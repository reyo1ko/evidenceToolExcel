"""
エビデンス自動貼り付けツール
- 起動時にExcelの貼り付け先ブックを選択
- Ctrl+Shift+E またはボタンで範囲選択起動
- 選択範囲をキャプチャしてExcelのアクティブシートに縦並び自動追記（原寸大）
- マルチモニター対応
"""

import os
import tempfile
import subprocess
import sys
import threading
import time
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox

import keyboard
import pystray
import pythoncom
import win32api
import win32com.client
import win32event
from PIL import Image, ImageDraw, ImageGrab
from screeninfo import get_monitors

MB_OK = 0x00000000
MB_ICONWARNING = 0x00000030
MB_ICONERROR = 0x00000010


# ===== 状態管理 =====

selected_target = {"workbook": None}  # シートはスクショ時にActiveSheetで動的取得
selection_state = {"active": False, "selector": None}
instance_lock = {"handle": None}


# ===== Excel接続 =====

def get_xl():
    """現在稼働中のExcelインスタンスを取得。見つからなければNoneを返す"""
    pythoncom.CoInitialize()
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return None


def get_excel_workbooks():
    """ブック一覧の読み取りだけを再試行し、COMの属性解決失敗にも対応する。"""
    last_error = None
    for attempt in range(3):
        try:
            xl = get_xl()
            if xl is None:
                raise RuntimeError("起動中のExcelに接続できません。")
            try:
                books = xl.Workbooks
            except AttributeError:
                # 動的ラッパーは名前解決中のCOMエラーをAttributeErrorに変える。
                # 直接取得して、元のCOMエラーを保持する。
                dispid = xl._oleobj_.GetIDsOfNames("Workbooks")
                books = win32com.client.Dispatch(
                    xl._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
                )
            names = [books.Item(i).Name for i in range(1, books.Count + 1)]
            return xl, books, names
        except (AttributeError, pythoncom.com_error, RuntimeError) as exc:
            last_error = exc
            if attempt < 2:
                time.sleep(0.3)
    raise RuntimeError(
        "Excelのブック一覧を取得できませんでした。\n"
        "Excelのセル編集中・ダイアログ表示中の場合は終了して、再試行してください。\n"
        "改善しない場合は、対象ブックを開き直して「貼り付け先を変更」で選択してください。\n"
        f"詳細: {last_error}"
    ) from last_error


def is_excel_running():
    """Excelが起動中であるかを判定する"""
    pythoncom.CoInitialize()
    try:
        win32com.client.GetActiveObject("Excel.Application")
        return True
    except Exception:
        pass

    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/NH"],
            capture_output=True,
            text=True,
            check=False,
        )
        output = result.stdout.strip()
        if result.returncode == 0 and output and "NO TASKS" not in output.upper():
            return True
    except Exception:
        pass

    return False


def create_single_instance_lock():
    """同じアプリケーションの多重起動を防止するためのロックを作成する"""
    name = "Global\\EvidenceToolExcelSingleInstanceMutex"
    handle = win32event.CreateMutex(None, False, name)
    if win32api.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        win32api.CloseHandle(handle)
        return None
    return handle


def show_native_message(title, message, flags):
    try:
        win32api.MessageBox(0, message, title, flags)
    except Exception:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, message, parent=root)
        root.destroy()


def release_single_instance_lock(handle):
    try:
        if handle is not None:
            win32event.ReleaseMutex(handle)
            win32api.CloseHandle(handle)
    except Exception:
        pass


# ===== 貼り付け先選択ダイアログ =====

def select_target_dialog(parent=None):
    """貼り付け先ブックを選択させる。選択完了でTrue、キャンセル・失敗でFalseを返す"""
    try:
        xl, workbooks, books = get_excel_workbooks()
    except RuntimeError as exc:
        messagebox.showerror("エラー", str(exc))
        return False

    if not books:
        messagebox.showerror("エラー", "Excelにブックが開かれていません。")
        return False

    # ブックが1つなら自動選択
    if len(books) == 1:
        selected_target["workbook"] = books[0]
        return True

    # 複数ブック：ファイル名幅に合わせてダイアログサイズを動的計算
    font = tkfont.Font(family="Arial", size=10)
    dialog_w = max(400, max(font.measure(b) for b in books) + 80)
    dialog_h = 120 + len(books) * 30

    dialog = tk.Toplevel(parent)
    dialog.title("貼り付け先のExcelを選択")
    dialog.geometry(f"{dialog_w}x{dialog_h}")
    dialog.resizable(True, False)
    dialog.attributes("-topmost", True)
    dialog.grab_set()

    tk.Label(dialog, text="貼り付け先のExcelファイルを選択してください",
             font=("Arial", 10), pady=12).pack()

    var = tk.StringVar(value=books[0])
    for book in books:
        tk.Radiobutton(dialog, text=book, variable=var, value=book,
                       font=("Arial", 10), wraplength=dialog_w - 60).pack(anchor="w", padx=30)

    confirmed = {"value": False}

    def on_ok():
        selected_target["workbook"] = var.get()
        confirmed["value"] = True
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=12)
    tk.Button(btn_frame, text="決定", command=on_ok, width=10).pack(side="left", padx=5)
    tk.Button(btn_frame, text="キャンセル", command=on_cancel, width=10).pack(side="left", padx=5)

    dialog.wait_window()
    return confirmed["value"]


# ===== 範囲選択オーバーレイ =====

class ScreenSelector:
    """全モニターを覆う半透明オーバーレイで範囲選択するUI"""

    def __init__(self, parent, callback):
        self.callback = callback
        self.start_x = 0
        self.start_y = 0
        self.rect = None
        self.parent = parent

        monitors = get_monitors()
        self.offset_x = min(m.x for m in monitors)
        self.offset_y = min(m.y for m in monitors)
        total_w = max(m.x + m.width for m in monitors) - self.offset_x
        total_h = max(m.y + m.height for m in monitors) - self.offset_y

        self.root = tk.Toplevel(parent)
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.3)
        self.root.geometry(f"{total_w}x{total_h}+{self.offset_x}+{self.offset_y}")
        self.root.configure(bg="black")
        self.root.config(cursor="crosshair")

        self.canvas = tk.Canvas(self.root, bg="black", highlightthickness=0,
                                width=total_w, height=total_h)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.root.bind("<Escape>", lambda e: self.cancel())

        # キャンセルボタン（右上）
        tk.Button(
            self.root, text="✕ キャンセル（Esc）",
            command=self.cancel,
            bg="#cc0000", fg="white",
            font=("Arial", 11, "bold"),
            relief="flat", padx=10, pady=5,
            cursor="arrow"
        ).place(relx=1.0, y=10, anchor="ne", x=-10)

        self.root.grab_set()
        self.root.focus_force()

    def run(self):
        self.root.wait_window()

    def _screen_to_canvas(self, sx, sy):
        """スクリーン座標をキャンバス座標に変換"""
        return sx - self.offset_x, sy - self.offset_y

    def on_press(self, event):
        self.start_x = event.x_root
        self.start_y = event.y_root
        if self.rect:
            self.canvas.delete(self.rect)

    def on_drag(self, event):
        if self.rect:
            self.canvas.delete(self.rect)
        cx1, cy1 = self._screen_to_canvas(self.start_x, self.start_y)
        cx2, cy2 = self._screen_to_canvas(event.x_root, event.y_root)
        self.rect = self.canvas.create_rectangle(
            cx1, cy1, cx2, cy2, outline="#ff0000", width=4, fill=""
        )

    def on_release(self, event):
        x1 = min(self.start_x, event.x_root)
        y1 = min(self.start_y, event.y_root)
        x2 = max(self.start_x, event.x_root)
        y2 = max(self.start_y, event.y_root)
        if x2 - x1 > 10 and y2 - y1 > 10:
            self.root.destroy()
            if self.parent is not None:
                self.parent.update_idletasks()
                self.parent.after(250, lambda: self.callback(x1, y1, x2, y2))
            else    :
                import time
                time.sleep(0.25)
                self.callback(x1, y1, x2, y2)
        else:
            self.root.destroy()

    def cancel(self):
        self.root.destroy()


# ===== Excelへの貼り付け =====

def capture_and_paste(x1, y1, x2, y2):
    """指定範囲をキャプチャしてExcelのアクティブシートに原寸大で追記する"""
    try:
        pythoncom.CoInitialize()

        if not selected_target["workbook"]:
            messagebox.showerror(
                "エラー", "貼り付け先が選択されていません。\n「貼り付け先を変更」ボタンで選択してください。"
            )
            return

        xl, workbooks, names = get_excel_workbooks()
        if selected_target["workbook"] not in names:
            raise RuntimeError(
                f"接続先Excelに対象ブック「{selected_target['workbook']}」が見つかりません。\n"
                "対象ブックを開き、「貼り付け先を変更」で選び直してください。"
            )
        wb = workbooks.Item(selected_target["workbook"])
        ws = wb.ActiveSheet

        # スクリーンショット取得
        img = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
        orig_w, orig_h = img.size

        # 貼り付け位置：既存画像とテキストセルの下端のうち大きい方
        MARGIN = 10
        next_top = MARGIN

        for i in range(1, ws.Shapes.Count + 1):
            shape = ws.Shapes.Item(i)
            next_top = max(next_top, shape.Top + shape.Height + MARGIN)

        try:
            used = ws.UsedRange
            next_top = max(next_top, used.Top + used.Height + MARGIN)
        except Exception:
            pass

        # PNGをブックに埋め込む。ExcelのShapesにはPasteメソッドがない。
        # 画面の前面切り替えやクリップボードの状態に依存しない。
        with tempfile.TemporaryDirectory(prefix="evidence_excel_") as temp_dir:
            image_path = os.path.join(temp_dir, "capture.png")
            img.save(image_path, format="PNG")
            ws.Shapes.AddPicture(
                image_path, 0, -1, MARGIN, next_top,
                orig_w * 0.75, orig_h * 0.75,
            )

    except Exception as e:
        messagebox.showerror("エラー", f"貼り付けに失敗しました:\n{e}")


# ===== ショートカット起動 =====

def launch_selector(parent=None, update_status=None):
    """スクリーンセレクタを起動する。既に選択中なら現在の選択を停止する"""
    def start():
        if not selected_target["workbook"] and parent is not None:
            if not select_target_dialog(parent):
                return
            if update_status is not None:
                parent.after(0, update_status)

        if selection_state["active"]:
            if selection_state["selector"] is not None:
                selection_state["selector"].cancel()
            return

        selection_state["active"] = True
        selection_state["selector"] = None

        try:
            selector = ScreenSelector(parent, capture_and_paste)
            selection_state["selector"] = selector
            selector.run()
        finally:
            selection_state["selector"] = None
            selection_state["active"] = False

    if parent is not None:
        parent.after(0, start)
    else:
        start()


def create_tray_icon_image():
    """タスクトレイアイコン用の小さな画像を生成する"""
    image = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rectangle((10, 20, 54, 44), fill="#0078d7", outline="white", width=2)
    draw.rectangle((22, 12, 42, 22), fill="#0078d7")
    draw.rectangle((28, 6, 36, 14), fill="white")
    return image


# ===== メイン =====

def main():
    pythoncom.CoInitialize()

    lock = create_single_instance_lock()
    if lock is None:
        show_native_message(
            "起動済み",
            "このアプリはすでに起動しています。既存のインスタンスを操作してください。",
            MB_ICONWARNING | MB_OK,
        )
        return

    if not is_excel_running():
        show_native_message(
            "エラー",
            "Excelが起動していません。Excelを起動してから再実行してください。",
            MB_ICONERROR | MB_OK,
        )
        release_single_instance_lock(lock)
        sys.exit(0)

    root = tk.Tk()
    root.title("エビデンスツール")
    root.geometry("320x185")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    root.deiconify()

    tk.Label(root, text="エビデンス自動貼り付けツール",
             font=("Arial", 11, "bold")).pack(pady=(15, 4))

    status_var = tk.StringVar(value="貼り付け先: 未選択")
    tk.Label(root, textvariable=status_var, font=("Arial", 9), fg="gray").pack()

    def update_status():
        wb = selected_target["workbook"]
        status_var.set(f"貼り付け先: {wb}（アクティブシート）" if wb else "貼り付け先: 未選択")

    def change_target():
        if select_target_dialog(root):
            update_status()

    tk.Button(root, text="📷 スクショ起動", command=lambda: launch_selector(root, update_status), width=22).pack(pady=(10, 0))
    tk.Label(root, text="Ctrl+Shift+E", font=("Arial", 9, "bold"), fg="#0066cc").pack(pady=(0, 4))
    tk.Button(root, text="貼り付け先を変更", command=change_target, width=22).pack(pady=2)
    tk.Button(root, text="終了", command=root.destroy, width=22).pack(pady=2)

    def show_window():
        root.deiconify()
        root.lift()
        root.attributes("-topmost", True)
        root.after(10, lambda: root.attributes("-topmost", False))

    def hide_window():
        root.withdraw()

    def on_exit():
        icon.stop()
        root.quit()

    def confirm_close():
        if messagebox.askyesno("終了確認", "アプリを終了しますか？"):
            on_exit()

    root.protocol("WM_DELETE_WINDOW", confirm_close)

    icon = pystray.Icon(
        "evidence_tool",
        create_tray_icon_image(),
        "エビデンスツール",
        menu=pystray.Menu(
            pystray.MenuItem("開く", lambda _: root.after(0, show_window)),
            pystray.MenuItem("スクショ起動", lambda _: root.after(0, lambda: launch_selector(root, update_status))),
            pystray.MenuItem("貼り付け先変更", lambda _: root.after(0, change_target)),
            pystray.MenuItem("終了", lambda _: root.after(0, on_exit)),
        ),
    )
    threading.Thread(target=icon.run, daemon=True).start()

    keyboard.add_hotkey("ctrl+shift+e", lambda: launch_selector(root, update_status))
    try:
        root.mainloop()
    finally:
        keyboard.unhook_all()
        release_single_instance_lock(lock)


if __name__ == "__main__":
    main()
