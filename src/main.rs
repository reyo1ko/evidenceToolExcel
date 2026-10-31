#![windows_subsystem = "windows"]
#![allow(unsafe_op_in_unsafe_fn, unused_must_use)]

use std::fs::{self, File};
use std::io::Write;
use std::mem::size_of;
use std::path::{Path, PathBuf};
use std::ptr::null_mut;
use std::sync::{Mutex, OnceLock};
use std::time::{SystemTime, UNIX_EPOCH};

use windows::Win32::Foundation::{
    BOOL, COLORREF, HINSTANCE, HWND, LPARAM, LRESULT, POINT, RECT, WPARAM,
};
use windows::Win32::Graphics::Gdi::{
    BI_RGB, BITMAPINFO, BITMAPINFOHEADER, BLACK_BRUSH, BeginPaint, BitBlt, CAPTUREBLT,
    CLEARTYPE_QUALITY, CLIP_DEFAULT_PRECIS, COLOR_WINDOW, CreateCompatibleBitmap,
    CreateCompatibleDC, CreateFontW, CreatePen, CreateSolidBrush, DEFAULT_CHARSET, DEFAULT_PITCH,
    DIB_RGB_COLORS, DT_CENTER, DT_SINGLELINE, DT_VCENTER, DeleteDC, DeleteObject, DrawTextW,
    EndPaint, FW_NORMAL, FillRect, GetDC, GetDIBits, GetStockObject, GetSysColorBrush, HBRUSH, HDC,
    HFONT, HGDIOBJ, InvalidateRect, NULL_BRUSH, OUT_DEFAULT_PRECIS, PAINTSTRUCT, PS_SOLID,
    Rectangle, ReleaseDC, RoundRect, SRCCOPY, SelectObject, SetBkMode, SetTextColor, TRANSPARENT,
    TextOutW, UpdateWindow,
};
use windows::Win32::System::Com::{
    CLSIDFromProgID, COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize, DISPATCH_FLAGS,
    DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPPARAMS, IDispatch,
};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Ole::GetActiveObject;
use windows::Win32::System::SystemServices::SS_CENTER;
use windows::Win32::UI::Controls::{DRAWITEMSTRUCT, ODS_SELECTED};
use windows::Win32::UI::HiDpi::{
    DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2, GetDpiForSystem, GetDpiForWindow,
    SetProcessDpiAwarenessContext,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    MOD_CONTROL, MOD_SHIFT, RegisterHotKey, ReleaseCapture, SetCapture, SetFocus, UnregisterHotKey,
    VK_ESCAPE,
};
use windows::Win32::UI::WindowsAndMessaging::*;
use windows::core::{BSTR, GUID, Interface, PCWSTR, VARIANT};

const ID_CAPTURE: usize = 1001;
const ID_EXIT: usize = 1002;
const ID_REFRESH: usize = 1003;
const ID_WORKBOOK: usize = 1004;
const ID_SHEET: usize = 1005;
const HOTKEY_ID: i32 = 1;
const MAIN_CLASS: &str = "EvidenceToolExcelMain";
const OVERLAY_CLASS: &str = "EvidenceToolExcelOverlay";

static mut MAIN_HWND: HWND = HWND(null_mut());
static mut STATUS_HWND: HWND = HWND(null_mut());
static mut WORKBOOK_HWND: HWND = HWND(null_mut());
static mut SHEET_HWND: HWND = HWND(null_mut());
static mut OVERLAY_HWND: HWND = HWND(null_mut());
static mut UI_FONT: HFONT = HFONT(null_mut());
static mut START: POINT = POINT { x: 0, y: 0 };
static mut CURRENT: POINT = POINT { x: 0, y: 0 };
static mut DRAGGING: bool = false;
static mut VIRTUAL_X: i32 = 0;
static mut VIRTUAL_Y: i32 = 0;

#[derive(Clone)]
struct WorkbookTarget {
    name: String,
    sheets: Vec<String>,
}

static TARGETS: OnceLock<Mutex<Vec<WorkbookTarget>>> = OnceLock::new();

type AppResult<T> = Result<T, String>;

fn wide(value: &str) -> Vec<u16> {
    value.encode_utf16().chain(Some(0)).collect()
}

unsafe extern "system" fn set_child_font(hwnd: HWND, font: LPARAM) -> BOOL {
    SendMessageW(hwnd, WM_SETFONT, WPARAM(font.0 as usize), LPARAM(1));
    BOOL(1)
}

unsafe fn set_status(text: &str) {
    let text = wide(text);
    let _ = SetWindowTextW(STATUS_HWND, PCWSTR(text.as_ptr()));
}

unsafe fn show_error(message: &str) {
    let title = wide("エラー");
    let body = wide(message);
    MessageBoxW(
        MAIN_HWND,
        PCWSTR(body.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_OK | MB_ICONERROR,
    );
}

unsafe fn register_class(name: &str, proc: WNDPROC, background: HBRUSH) -> AppResult<()> {
    let instance = GetModuleHandleW(None).map_err(|e| e.to_string())?;
    let class_name = wide(name);
    let cursor = LoadCursorW(None, IDC_ARROW).map_err(|e| e.to_string())?;
    let class = WNDCLASSW {
        hCursor: cursor,
        hInstance: HINSTANCE(instance.0),
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: proc,
        hbrBackground: background,
        ..Default::default()
    };
    if RegisterClassW(&class) == 0 {
        return Err("ウィンドウクラスを登録できませんでした。".into());
    }
    Ok(())
}

unsafe fn create_main_window() -> AppResult<HWND> {
    register_class(
        MAIN_CLASS,
        Some(main_proc),
        HBRUSH((COLOR_WINDOW.0 as usize + 1) as *mut _),
    )?;
    register_class(
        OVERLAY_CLASS,
        Some(overlay_proc),
        HBRUSH(GetStockObject(BLACK_BRUSH).0),
    )?;
    let instance = GetModuleHandleW(None).map_err(|e| e.to_string())?;
    let dpi = GetDpiForSystem().max(96) as i32;
    let s = |value: i32| (value * dpi + 48) / 96;
    let class_name = wide(MAIN_CLASS);
    let title = wide("エビデンス貼り付けツール");
    let hwnd = CreateWindowExW(
        WS_EX_TOPMOST,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        s(520),
        s(360),
        None,
        None,
        HINSTANCE(instance.0),
        None,
    )
    .map_err(|e| format!("メイン画面を作成できませんでした: {e}"))?;
    MAIN_HWND = hwnd;
    let static_class = wide("STATIC");
    let button_class = wide("BUTTON");
    let workbook_label = wide("貼り付け先のExcelブック");
    CreateWindowExW(
        Default::default(),
        PCWSTR(static_class.as_ptr()),
        PCWSTR(workbook_label.as_ptr()),
        WS_CHILD | WS_VISIBLE,
        s(28),
        s(24),
        s(300),
        s(22),
        hwnd,
        None,
        HINSTANCE(instance.0),
        None,
    );
    let combo_class = wide("COMBOBOX");
    WORKBOOK_HWND = CreateWindowExW(
        Default::default(),
        PCWSTR(combo_class.as_ptr()),
        PCWSTR::null(),
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
        s(28),
        s(49),
        s(352),
        s(200),
        hwnd,
        HMENU(ID_WORKBOOK as *mut _),
        HINSTANCE(instance.0),
        None,
    )
    .map_err(|e| e.to_string())?;
    let refresh = wide("再読込");
    CreateWindowExW(
        Default::default(),
        PCWSTR(button_class.as_ptr()),
        PCWSTR(refresh.as_ptr()),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_PUSHBUTTON as u32),
        s(394),
        s(49),
        s(98),
        s(32),
        hwnd,
        HMENU(ID_REFRESH as *mut _),
        HINSTANCE(instance.0),
        None,
    )
    .map_err(|e| e.to_string())?;
    let sheet_label = wide("貼り付け先のシート");
    CreateWindowExW(
        Default::default(),
        PCWSTR(static_class.as_ptr()),
        PCWSTR(sheet_label.as_ptr()),
        WS_CHILD | WS_VISIBLE,
        s(28),
        s(94),
        s(300),
        s(22),
        hwnd,
        None,
        HINSTANCE(instance.0),
        None,
    );
    SHEET_HWND = CreateWindowExW(
        Default::default(),
        PCWSTR(combo_class.as_ptr()),
        PCWSTR::null(),
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
        s(28),
        s(119),
        s(464),
        s(200),
        hwnd,
        HMENU(ID_SHEET as *mut _),
        HINSTANCE(instance.0),
        None,
    )
    .map_err(|e| e.to_string())?;
    let guide = wide("Excelが開かれていません。\r\nExcelを起動して［再読込］を押してください。");
    STATUS_HWND = CreateWindowExW(
        Default::default(),
        PCWSTR(static_class.as_ptr()),
        PCWSTR(guide.as_ptr()),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE(SS_CENTER.0),
        s(28),
        s(165),
        s(464),
        s(42),
        hwnd,
        None,
        HINSTANCE(instance.0),
        None,
    )
    .map_err(|e| e.to_string())?;
    let capture = wide("スクリーンショットを撮る  (Ctrl+Shift+E)");
    CreateWindowExW(
        Default::default(),
        PCWSTR(button_class.as_ptr()),
        PCWSTR(capture.as_ptr()),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_OWNERDRAW as u32),
        s(70),
        s(217),
        s(380),
        s(48),
        hwnd,
        HMENU(ID_CAPTURE as *mut _),
        HINSTANCE(instance.0),
        None,
    )
    .map_err(|e| e.to_string())?;
    let exit = wide("終了");
    CreateWindowExW(
        Default::default(),
        PCWSTR(button_class.as_ptr()),
        PCWSTR(exit.as_ptr()),
        WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_PUSHBUTTON as u32),
        s(200),
        s(285),
        s(120),
        s(34),
        hwnd,
        HMENU(ID_EXIT as *mut _),
        HINSTANCE(instance.0),
        None,
    )
    .map_err(|e| e.to_string())?;
    let face = wide("Yu Gothic UI");
    UI_FONT = CreateFontW(
        -s(15),
        0,
        0,
        0,
        FW_NORMAL.0 as i32,
        0,
        0,
        0,
        DEFAULT_CHARSET.0 as u32,
        OUT_DEFAULT_PRECIS.0 as u32,
        CLIP_DEFAULT_PRECIS.0 as u32,
        CLEARTYPE_QUALITY.0 as u32,
        DEFAULT_PITCH.0 as u32,
        PCWSTR(face.as_ptr()),
    );
    EnumChildWindows(hwnd, Some(set_child_font), LPARAM(UI_FONT.0 as isize));
    if RegisterHotKey(hwnd, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, 'E' as u32).is_err() {
        set_status("ホットキーを登録できません。ボタンから撮影できます。");
    }
    match refresh_targets() {
        Ok(()) => set_status("貼り付け先を選択して撮影してください"),
        Err(_) => {
            set_status("Excelが開かれていません。\r\nExcelを起動して［再読込］を押してください。")
        }
    }
    Ok(hwnd)
}

unsafe fn connect_excel() -> AppResult<IDispatch> {
    let clsid = CLSIDFromProgID(PCWSTR(wide("Excel.Application").as_ptr()))
        .map_err(|_| "Microsoft Excelがインストールされていません。".to_string())?;
    let mut unknown = None;
    GetActiveObject(&clsid, None, &mut unknown).map_err(|_| {
        "Excelが起動していません。Excelで対象ブックを開いてから「再読込」を押してください。"
            .to_string()
    })?;
    unknown
        .ok_or_else(|| "起動中のExcelに接続できません。".to_string())?
        .cast()
        .map_err(|e| format!("Excelに接続できません: {e}"))
}

unsafe fn variant_string(value: &VARIANT, label: &str) -> AppResult<String> {
    BSTR::try_from(value)
        .map(|s| s.to_string())
        .map_err(|_| format!("{label}の名前を取得できません。"))
}

unsafe fn refresh_targets() -> AppResult<()> {
    let previous = selected_target_names().ok();
    let excel = connect_excel()?;
    let workbooks = as_dispatch(&property(&excel, "Workbooks")?, "ブック一覧")?;
    let count = i32::try_from(&property(&workbooks, "Count")?)
        .map_err(|_| "Excelのブック数を取得できません。".to_string())?;
    if count == 0 {
        return Err("Excelでブックが開かれていません。対象ブックを開いてください。".into());
    }

    let mut targets = Vec::new();
    for index in 1..=count {
        let mut item_args = [VARIANT::from(index)];
        let book = as_dispatch(
            &invoke(
                &workbooks,
                "Item",
                DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                &mut item_args,
            )?,
            "ブック",
        )?;
        let name = variant_string(&property(&book, "Name")?, "ブック")?;
        let worksheets = as_dispatch(&property(&book, "Worksheets")?, "シート一覧")?;
        let sheet_count = i32::try_from(&property(&worksheets, "Count")?)
            .map_err(|_| "Excelのシート数を取得できません。".to_string())?;
        let mut sheets = Vec::new();
        for sheet_index in 1..=sheet_count {
            let mut sheet_args = [VARIANT::from(sheet_index)];
            let sheet = as_dispatch(
                &invoke(
                    &worksheets,
                    "Item",
                    DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                    &mut sheet_args,
                )?,
                "シート",
            )?;
            sheets.push(variant_string(&property(&sheet, "Name")?, "シート")?);
        }
        targets.push(WorkbookTarget { name, sheets });
    }

    let workbook_index = previous
        .as_ref()
        .and_then(|(book_name, _)| targets.iter().position(|book| book.name == *book_name))
        .unwrap_or(0);
    let preferred_sheet = previous.as_ref().and_then(|(book_name, sheet_name)| {
        targets
            .get(workbook_index)
            .filter(|book| book.name == *book_name)
            .and_then(|book| {
                book.sheets
                    .iter()
                    .any(|sheet| sheet == sheet_name)
                    .then(|| sheet_name.clone())
            })
    });

    let store = TARGETS.get_or_init(|| Mutex::new(Vec::new()));
    *store
        .lock()
        .map_err(|_| "選択先情報を更新できません。".to_string())? = targets;
    SendMessageW(WORKBOOK_HWND, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let targets = store
        .lock()
        .map_err(|_| "選択先情報を表示できません。".to_string())?;
    for target in targets.iter() {
        let name = wide(&target.name);
        SendMessageW(
            WORKBOOK_HWND,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(name.as_ptr() as isize),
        );
    }
    SendMessageW(
        WORKBOOK_HWND,
        CB_SETCURSEL,
        WPARAM(workbook_index),
        LPARAM(0),
    );
    drop(targets);
    populate_sheet_combo(workbook_index, preferred_sheet.as_deref())?;
    Ok(())
}

unsafe fn populate_sheet_combo(
    workbook_index: usize,
    preferred_sheet: Option<&str>,
) -> AppResult<()> {
    SendMessageW(SHEET_HWND, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let targets = TARGETS
        .get_or_init(|| Mutex::new(Vec::new()))
        .lock()
        .map_err(|_| "シート一覧を表示できません。".to_string())?;
    let target = targets
        .get(workbook_index)
        .ok_or_else(|| "Excelブックを選択してください。".to_string())?;
    for sheet in &target.sheets {
        let name = wide(sheet);
        SendMessageW(
            SHEET_HWND,
            CB_ADDSTRING,
            WPARAM(0),
            LPARAM(name.as_ptr() as isize),
        );
    }
    if !target.sheets.is_empty() {
        let sheet_index = preferred_sheet
            .and_then(|name| target.sheets.iter().position(|sheet| sheet == name))
            .unwrap_or(0);
        SendMessageW(SHEET_HWND, CB_SETCURSEL, WPARAM(sheet_index), LPARAM(0));
    }
    Ok(())
}

unsafe fn selected_target_names() -> AppResult<(String, String)> {
    let book_index = SendMessageW(WORKBOOK_HWND, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let sheet_index = SendMessageW(SHEET_HWND, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if book_index < 0 || sheet_index < 0 {
        return Err("貼り付け先のExcelブックとシートを選択してください。".into());
    }
    let targets = TARGETS
        .get_or_init(|| Mutex::new(Vec::new()))
        .lock()
        .map_err(|_| "選択先情報を読み取れません。".to_string())?;
    let book = targets
        .get(book_index as usize)
        .ok_or_else(|| "選択したExcelブックが見つかりません。再読込してください。".to_string())?;
    let sheet = book
        .sheets
        .get(sheet_index as usize)
        .ok_or_else(|| "選択したシートが見つかりません。再読込してください。".to_string())?;
    Ok((book.name.clone(), sheet.clone()))
}

unsafe fn resolve_selected_sheet() -> AppResult<IDispatch> {
    let (book_name, sheet_name) = selected_target_names()?;
    let excel = connect_excel()?;
    let workbooks = as_dispatch(&property(&excel, "Workbooks")?, "ブック一覧")?;
    let mut book_args = [VARIANT::from(book_name.as_str())];
    let workbook = as_dispatch(
        &invoke(
            &workbooks,
            "Item",
            DISPATCH_METHOD | DISPATCH_PROPERTYGET,
            &mut book_args,
        )
        .map_err(|_| format!("ブック「{book_name}」が見つかりません。再読込してください。"))?,
        "選択したブック",
    )?;
    let worksheets = as_dispatch(&property(&workbook, "Worksheets")?, "シート一覧")?;
    let mut sheet_args = [VARIANT::from(sheet_name.as_str())];
    as_dispatch(
        &invoke(
            &worksheets,
            "Item",
            DISPATCH_METHOD | DISPATCH_PROPERTYGET,
            &mut sheet_args,
        )
        .map_err(|_| format!("シート「{sheet_name}」が見つかりません。再読込してください。"))?,
        "選択したシート",
    )
}

unsafe fn resolve_active_sheet() -> AppResult<IDispatch> {
    let excel = connect_excel()?;
    let workbook = as_dispatch(&property(&excel, "ActiveWorkbook")?, "アクティブブック")
        .map_err(|_| "Excelでブックが開かれていません。".to_string())?;
    as_dispatch(&property(&workbook, "ActiveSheet")?, "アクティブシート")
}

unsafe fn prepare_selection() -> AppResult<()> {
    // 撮影に入る前に検証し、Excel未起動時に暗い選択画面へ遷移しないようにする。
    connect_excel()?;
    if selected_target_names().is_err() {
        refresh_targets()?;
    }
    resolve_selected_sheet()?;
    start_selection()
}

unsafe fn start_selection() -> AppResult<()> {
    if !OVERLAY_HWND.0.is_null() {
        return Ok(());
    }
    VIRTUAL_X = GetSystemMetrics(SM_XVIRTUALSCREEN);
    VIRTUAL_Y = GetSystemMetrics(SM_YVIRTUALSCREEN);
    let width = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    let height = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    let instance = GetModuleHandleW(None).map_err(|e| e.to_string())?;
    let class_name = wide(OVERLAY_CLASS);
    let hwnd = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_LAYERED | WS_EX_TOOLWINDOW,
        PCWSTR(class_name.as_ptr()),
        PCWSTR::null(),
        WS_POPUP,
        VIRTUAL_X,
        VIRTUAL_Y,
        width,
        height,
        None,
        None,
        HINSTANCE(instance.0),
        None,
    )
    .map_err(|e| format!("範囲選択画面を作成できませんでした: {e}"))?;
    OVERLAY_HWND = hwnd;
    SetLayeredWindowAttributes(hwnd, COLORREF(0), 105, LWA_ALPHA).map_err(|e| e.to_string())?;
    ShowWindow(hwnd, SW_SHOW);
    SetForegroundWindow(hwnd);
    SetFocus(hwnd);
    SetCapture(hwnd);
    set_status("範囲をドラッグしてください（Escでキャンセル）");
    Ok(())
}

unsafe fn close_overlay() {
    DRAGGING = false;
    ReleaseCapture();
    if !OVERLAY_HWND.0.is_null() {
        DestroyWindow(OVERLAY_HWND);
        OVERLAY_HWND = HWND(null_mut());
    }
}

fn signed_word(value: isize) -> i32 {
    value as i16 as i32
}

unsafe fn invalidate_selection_change(hwnd: HWND, old: POINT, new: POINT) {
    let left = START.x.min(old.x).min(new.x) - 48;
    let top = START.y.min(old.y).min(new.y) - 48;
    let right = START.x.max(old.x).max(new.x) + 180;
    let bottom = START.y.max(old.y).max(new.y) + 48;
    let dirty = RECT {
        left,
        top,
        right,
        bottom,
    };
    // WM_ERASEBKGNDを発生させず、WM_PAINT側のダブルバッファでまとめて更新する。
    InvalidateRect(hwnd, Some(&dirty), BOOL(0));
}

unsafe extern "system" fn main_proc(hwnd: HWND, msg: u32, w: WPARAM, l: LPARAM) -> LRESULT {
    match msg {
        WM_COMMAND => {
            let control_id = w.0 & 0xffff;
            let notification = (w.0 >> 16) & 0xffff;
            match control_id {
                ID_CAPTURE => {
                    if let Err(e) = prepare_selection() {
                        show_error(&e);
                    }
                }
                ID_REFRESH => match refresh_targets() {
                    Ok(()) => set_status("一覧を更新しました。貼り付け先を選択してください"),
                    Err(e) => {
                        set_status(
                            "Excelの一覧を取得できませんでした。\r\nExcelを確認して［再読込］を押してください。",
                        );
                        show_error(&e);
                    }
                },
                ID_WORKBOOK if notification == CBN_SELCHANGE as usize => {
                    let index = SendMessageW(WORKBOOK_HWND, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                    if index >= 0 {
                        if let Err(e) = populate_sheet_combo(index as usize, None) {
                            show_error(&e);
                        }
                    }
                }
                ID_EXIT => {
                    DestroyWindow(hwnd);
                }
                _ => {}
            }
            LRESULT(0)
        }
        WM_HOTKEY => {
            if w.0 == HOTKEY_ID as usize {
                if let Err(e) = prepare_selection() {
                    show_error(&e);
                }
            }
            LRESULT(0)
        }
        WM_CTLCOLORSTATIC => {
            let dc = HDC(w.0 as *mut _);
            SetBkMode(dc, TRANSPARENT);
            SetTextColor(dc, COLORREF(0x00202020));
            LRESULT(GetSysColorBrush(COLOR_WINDOW).0 as isize)
        }
        WM_DRAWITEM => {
            let item = &*(l.0 as *const DRAWITEMSTRUCT);
            if item.CtlID == ID_CAPTURE as u32 {
                let pressed = item.itemState.0 & ODS_SELECTED.0 != 0;
                // 先に親画面と同じ色で消去し、余白を取った角丸ボタンを描画する。
                FillRect(item.hDC, &item.rcItem, GetSysColorBrush(COLOR_WINDOW));
                let background = if pressed {
                    COLORREF(0x00a65406)
                } else {
                    COLORREF(0x00d96d00)
                };
                let brush = CreateSolidBrush(background);
                let pen = CreatePen(PS_SOLID, 1, background);
                let old_pen = SelectObject(item.hDC, HGDIOBJ(pen.0));
                let old_brush = SelectObject(item.hDC, HGDIOBJ(brush.0));
                RoundRect(
                    item.hDC,
                    item.rcItem.left + 2,
                    item.rcItem.top + 2,
                    item.rcItem.right - 2,
                    item.rcItem.bottom - 2,
                    12,
                    12,
                );
                SelectObject(item.hDC, old_pen);
                SelectObject(item.hDC, old_brush);
                DeleteObject(HGDIOBJ(pen.0));
                DeleteObject(HGDIOBJ(brush.0));

                SetBkMode(item.hDC, TRANSPARENT);
                SetTextColor(item.hDC, COLORREF(0x00ffffff));
                let old_font = if UI_FONT.0.is_null() {
                    HGDIOBJ(null_mut())
                } else {
                    SelectObject(item.hDC, HGDIOBJ(UI_FONT.0))
                };
                let mut label: Vec<u16> = "スクリーンショットを撮る  (Ctrl+Shift+E)"
                    .encode_utf16()
                    .collect();
                let mut text_rect = item.rcItem;
                DrawTextW(
                    item.hDC,
                    &mut label,
                    &mut text_rect,
                    DT_CENTER | DT_VCENTER | DT_SINGLELINE,
                );
                if !old_font.0.is_null() {
                    SelectObject(item.hDC, old_font);
                }
                return LRESULT(1);
            }
            DefWindowProcW(hwnd, msg, w, l)
        }
        WM_DESTROY => {
            UnregisterHotKey(hwnd, HOTKEY_ID);
            if !UI_FONT.0.is_null() {
                DeleteObject(HGDIOBJ(UI_FONT.0));
                UI_FONT = HFONT(null_mut());
            }
            PostQuitMessage(0);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, w, l),
    }
}

unsafe extern "system" fn overlay_proc(hwnd: HWND, msg: u32, w: WPARAM, l: LPARAM) -> LRESULT {
    match msg {
        WM_KEYDOWN if w.0 == VK_ESCAPE.0 as usize => {
            close_overlay();
            set_status("キャンセルしました");
            LRESULT(0)
        }
        WM_LBUTTONDOWN => {
            START = POINT {
                x: signed_word(l.0),
                y: signed_word(l.0 >> 16),
            };
            CURRENT = START;
            DRAGGING = true;
            invalidate_selection_change(hwnd, START, CURRENT);
            LRESULT(0)
        }
        WM_MOUSEMOVE if DRAGGING => {
            let old = CURRENT;
            CURRENT = POINT {
                x: signed_word(l.0),
                y: signed_word(l.0 >> 16),
            };
            invalidate_selection_change(hwnd, old, CURRENT);
            LRESULT(0)
        }
        WM_LBUTTONUP if DRAGGING => {
            CURRENT = POINT {
                x: signed_word(l.0),
                y: signed_word(l.0 >> 16),
            };
            let left = START.x.min(CURRENT.x) + VIRTUAL_X;
            let top = START.y.min(CURRENT.y) + VIRTUAL_Y;
            let right = START.x.max(CURRENT.x) + VIRTUAL_X;
            let bottom = START.y.max(CURRENT.y) + VIRTUAL_Y;
            close_overlay();
            if right - left < 10 || bottom - top < 10 {
                set_status("範囲が小さすぎるためキャンセルしました");
                return LRESULT(0);
            }
            std::thread::sleep(std::time::Duration::from_millis(180));
            set_status("Excelへ画像を追加しています...");
            match capture_and_insert(left, top, right, bottom) {
                Ok(()) => set_status("画像を追加しました。Excelで保存してください。"),
                Err(e) => {
                    set_status("画像の追加に失敗しました");
                    show_error(&e);
                }
            }
            LRESULT(0)
        }
        WM_PAINT => {
            let mut ps = PAINTSTRUCT::default();
            let dc = BeginPaint(hwnd, &mut ps);
            let paint_width = ps.rcPaint.right - ps.rcPaint.left;
            let paint_height = ps.rcPaint.bottom - ps.rcPaint.top;
            if paint_width <= 0 || paint_height <= 0 {
                EndPaint(hwnd, &ps);
                return LRESULT(0);
            }

            // 画面へ直接描かず、メモリ上で完成させてから一度に転送する。
            let buffer_dc = CreateCompatibleDC(dc);
            let buffer_bitmap = CreateCompatibleBitmap(dc, paint_width, paint_height);
            let old_bitmap = SelectObject(buffer_dc, HGDIOBJ(buffer_bitmap.0));
            let buffer_rect = RECT {
                left: 0,
                top: 0,
                right: paint_width,
                bottom: paint_height,
            };
            FillRect(
                buffer_dc,
                &buffer_rect,
                HBRUSH(GetStockObject(BLACK_BRUSH).0),
            );

            let offset_x = ps.rcPaint.left;
            let offset_y = ps.rcPaint.top;
            SetBkMode(buffer_dc, TRANSPARENT);
            SetTextColor(buffer_dc, COLORREF(0x00ffffff));
            let instruction: Vec<u16> = "ドラッグして撮影範囲を選択  ｜  Esc：キャンセル"
                .encode_utf16()
                .collect();
            TextOutW(buffer_dc, 24 - offset_x, 22 - offset_y, &instruction);
            if DRAGGING {
                let left = START.x.min(CURRENT.x);
                let top = START.y.min(CURRENT.y);
                let right = START.x.max(CURRENT.x);
                let bottom = START.y.max(CURRENT.y);
                let selected = RECT {
                    left: left - offset_x,
                    top: top - offset_y,
                    right: right - offset_x,
                    bottom: bottom - offset_y,
                };
                // 選択中の内側を明るくして、暗い未選択領域との差を明確にする。
                let fill = CreateSolidBrush(COLORREF(0x00ffffff));
                FillRect(buffer_dc, &selected, fill);
                DeleteObject(HGDIOBJ(fill.0));
                let pen = CreatePen(PS_SOLID, 4, COLORREF(0x0000ff));
                let old_pen = SelectObject(buffer_dc, HGDIOBJ(pen.0));
                let old_brush = SelectObject(buffer_dc, GetStockObject(NULL_BRUSH));
                Rectangle(
                    buffer_dc,
                    left - offset_x,
                    top - offset_y,
                    right - offset_x,
                    bottom - offset_y,
                );
                SelectObject(buffer_dc, old_pen);
                SelectObject(buffer_dc, old_brush);
                DeleteObject(HGDIOBJ(pen.0));

                SetTextColor(buffer_dc, COLORREF(0x0000ffff));
                let size_text: Vec<u16> = format!("{} × {} px", right - left, bottom - top)
                    .encode_utf16()
                    .collect();
                let label_x = (left + 8).max(8);
                let label_y = if top > 52 { top - 28 } else { bottom + 8 };
                TextOutW(
                    buffer_dc,
                    label_x - offset_x,
                    label_y - offset_y,
                    &size_text,
                );
            }

            BitBlt(
                dc,
                ps.rcPaint.left,
                ps.rcPaint.top,
                paint_width,
                paint_height,
                buffer_dc,
                0,
                0,
                SRCCOPY,
            );
            SelectObject(buffer_dc, old_bitmap);
            DeleteObject(HGDIOBJ(buffer_bitmap.0));
            DeleteDC(buffer_dc);
            EndPaint(hwnd, &ps);
            LRESULT(0)
        }
        WM_ERASEBKGND => LRESULT(1),
        WM_SETCURSOR => {
            SetCursor(LoadCursorW(None, IDC_CROSS).unwrap_or_default());
            LRESULT(1)
        }
        _ => DefWindowProcW(hwnd, msg, w, l),
    }
}

unsafe fn capture_and_insert(left: i32, top: i32, right: i32, bottom: i32) -> AppResult<()> {
    let width = right - left;
    let height = bottom - top;
    // Excelのサイズ単位はポイント。選択範囲の中心にあるウィンドウのDPIで
    // px→pt変換し、高DPIモニターの画像が96 DPI前提で引き伸ばされるのを防ぐ。
    let center = POINT {
        x: left + width / 2,
        y: top + height / 2,
    };
    let target_window = WindowFromPoint(center);
    let capture_dpi = if target_window.0.is_null() {
        GetDpiForSystem().max(96)
    } else {
        GetDpiForWindow(target_window).max(96)
    };
    let path = temp_bmp_path();
    capture_bmp(left, top, width, height, &path)?;
    let result = insert_into_excel(&path, width, height, capture_dpi);
    let _ = fs::remove_file(&path);
    result
}

fn temp_bmp_path() -> PathBuf {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_millis();
    std::env::temp_dir().join(format!(
        "evidence_tool_{}_{}.bmp",
        std::process::id(),
        stamp
    ))
}

unsafe fn capture_bmp(x: i32, y: i32, width: i32, height: i32, path: &Path) -> AppResult<()> {
    let screen = GetDC(None);
    if screen.0.is_null() {
        return Err("画面を取得できませんでした。".into());
    }
    let memory = CreateCompatibleDC(screen);
    let bitmap = CreateCompatibleBitmap(screen, width, height);
    if memory.0.is_null() || bitmap.0.is_null() {
        ReleaseDC(None, screen);
        return Err("画面キャプチャ用の領域を作成できませんでした。".into());
    }
    let old = SelectObject(memory, HGDIOBJ(bitmap.0));
    let copied = BitBlt(
        memory,
        0,
        0,
        width,
        height,
        screen,
        x,
        y,
        SRCCOPY | CAPTUREBLT,
    )
    .is_ok();
    let row_size = ((width * 32 + 31) / 32) * 4;
    let image_size = row_size * height;
    let mut pixels = vec![0u8; image_size as usize];
    let mut info = BITMAPINFO::default();
    info.bmiHeader = BITMAPINFOHEADER {
        biSize: size_of::<BITMAPINFOHEADER>() as u32,
        biWidth: width,
        biHeight: height,
        biPlanes: 1,
        biBitCount: 32,
        biCompression: BI_RGB.0,
        biSizeImage: image_size as u32,
        ..Default::default()
    };
    let lines = GetDIBits(
        memory,
        bitmap,
        0,
        height as u32,
        Some(pixels.as_mut_ptr().cast()),
        &mut info,
        DIB_RGB_COLORS,
    );
    SelectObject(memory, old);
    DeleteObject(HGDIOBJ(bitmap.0));
    DeleteDC(memory);
    ReleaseDC(None, screen);
    if !copied || lines == 0 {
        return Err("画面の読み取りに失敗しました。".into());
    }
    let mut file = File::create(path).map_err(|e| e.to_string())?;
    let pixel_offset = 14 + size_of::<BITMAPINFOHEADER>() as u32;
    let file_size = pixel_offset + image_size as u32;
    file.write_all(&[0x42, 0x4d]).map_err(|e| e.to_string())?;
    file.write_all(&file_size.to_le_bytes())
        .map_err(|e| e.to_string())?;
    file.write_all(&[0u8; 4]).map_err(|e| e.to_string())?;
    file.write_all(&pixel_offset.to_le_bytes())
        .map_err(|e| e.to_string())?;
    let header = std::slice::from_raw_parts(
        (&info.bmiHeader as *const BITMAPINFOHEADER).cast::<u8>(),
        size_of::<BITMAPINFOHEADER>(),
    );
    file.write_all(header).map_err(|e| e.to_string())?;
    file.write_all(&pixels).map_err(|e| e.to_string())?;
    Ok(())
}

unsafe fn dispid(object: &IDispatch, name: &str) -> AppResult<i32> {
    let name_w = wide(name);
    let names = [PCWSTR(name_w.as_ptr())];
    let mut id = 0i32;
    let iid_null = GUID::zeroed();
    object
        .GetIDsOfNames(&iid_null, names.as_ptr(), 1, 0x0411, &mut id)
        .or_else(|_| object.GetIDsOfNames(&iid_null, names.as_ptr(), 1, 0, &mut id))
        .map_err(|e| format!("Excelの {} を取得できません: {}", name, e))?;
    Ok(id)
}

unsafe fn invoke(
    object: &IDispatch,
    name: &str,
    flags: DISPATCH_FLAGS,
    args: &mut [VARIANT],
) -> AppResult<VARIANT> {
    let id = dispid(object, name)?;
    let mut params = DISPPARAMS {
        rgvarg: if args.is_empty() {
            null_mut()
        } else {
            args.as_mut_ptr()
        },
        rgdispidNamedArgs: null_mut(),
        cArgs: args.len() as u32,
        cNamedArgs: 0,
    };
    let mut result = VARIANT::new();
    let iid_null = GUID::zeroed();
    object
        .Invoke(
            id,
            &iid_null,
            0x0411,
            flags,
            &mut params,
            Some(&mut result),
            None,
            None,
        )
        .or_else(|_| {
            object.Invoke(
                id,
                &iid_null,
                0,
                flags,
                &mut params,
                Some(&mut result),
                None,
                None,
            )
        })
        .map_err(|e| format!("Excelの {} の実行に失敗しました: {}", name, e))?;
    Ok(result)
}

unsafe fn property(object: &IDispatch, name: &str) -> AppResult<VARIANT> {
    invoke(object, name, DISPATCH_PROPERTYGET, &mut [])
}

unsafe fn as_dispatch(value: &VARIANT, label: &str) -> AppResult<IDispatch> {
    IDispatch::try_from(value).map_err(|_| format!("Excelの {} を参照できません。", label))
}

unsafe fn insert_into_excel(
    path: &Path,
    width: i32,
    height: i32,
    capture_dpi: u32,
) -> AppResult<()> {
    let sheet = if WORKBOOK_HWND.0.is_null() {
        resolve_active_sheet()?
    } else {
        resolve_selected_sheet()?
    };
    let shapes = as_dispatch(&property(&sheet, "Shapes")?, "図形一覧")?;
    let mut next_top = 10.0f64;
    if let Ok(count_value) = property(&shapes, "Count") {
        if let Ok(count) = i32::try_from(&count_value) {
            for i in 1..=count {
                let mut args = [VARIANT::from(i)];
                if let Ok(item_value) = invoke(
                    &shapes,
                    "Item",
                    DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                    &mut args,
                ) {
                    if let Ok(item) = as_dispatch(&item_value, "図形") {
                        let top_value = property(&item, "Top")
                            .ok()
                            .and_then(|v| f64::try_from(&v).ok())
                            .unwrap_or(0.0);
                        let height_value = property(&item, "Height")
                            .ok()
                            .and_then(|v| f64::try_from(&v).ok())
                            .unwrap_or(0.0);
                        next_top = next_top.max(top_value + height_value + 10.0);
                    }
                }
            }
        }
    }
    if let Ok(used_value) = property(&sheet, "UsedRange") {
        if let Ok(used) = as_dispatch(&used_value, "使用範囲") {
            let top_value = property(&used, "Top")
                .ok()
                .and_then(|v| f64::try_from(&v).ok())
                .unwrap_or(0.0);
            let height_value = property(&used, "Height")
                .ok()
                .and_then(|v| f64::try_from(&v).ok())
                .unwrap_or(0.0);
            next_top = next_top.max(top_value + height_value + 10.0);
        }
    }
    let filename = path.to_string_lossy().to_string();
    // COMでは引数を逆順で渡す: Height, Width, Top, Left, SaveWithDocument, LinkToFile, Filename
    let points_per_pixel = 72.0 / capture_dpi.max(96) as f64;
    let mut args = [
        VARIANT::from(height as f64 * points_per_pixel),
        VARIANT::from(width as f64 * points_per_pixel),
        VARIANT::from(next_top),
        VARIANT::from(10.0f64),
        VARIANT::from(true),
        VARIANT::from(false),
        VARIANT::from(filename.as_str()),
    ];
    invoke(&shapes, "AddPicture", DISPATCH_METHOD, &mut args)?;
    Ok(())
}

fn main() {
    unsafe {
        // 複数ディスプレイで拡大率が異なっても、選択座標と取得ピクセルを一致させる。
        let _ = SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        if let Err(e) = CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() {
            let body = wide(&format!("初期化に失敗しました: {e}"));
            let title = wide("エラー");
            MessageBoxW(
                None,
                PCWSTR(body.as_ptr()),
                PCWSTR(title.as_ptr()),
                MB_OK | MB_ICONERROR,
            );
            return;
        }
        let args: Vec<String> = std::env::args().collect();
        if args.get(1).map(String::as_str) == Some("--self-test") {
            let result = capture_and_insert(0, 0, 64, 64);
            if let Some(result_path) = args.get(2) {
                let text = match &result {
                    Ok(()) => "OK".to_string(),
                    Err(e) => format!("ERROR: {e}"),
                };
                let _ = fs::write(result_path, text);
            }
            CoUninitialize();
            std::process::exit(if result.is_ok() { 0 } else { 1 });
        }
        match create_main_window() {
            Ok(hwnd) => {
                ShowWindow(hwnd, SW_SHOW);
                UpdateWindow(hwnd);
                let mut message = MSG::default();
                while GetMessageW(&mut message, None, 0, 0).as_bool() {
                    TranslateMessage(&message);
                    DispatchMessageW(&message);
                }
            }
            Err(e) => show_error(&e),
        }
        CoUninitialize();
    }
}
