fn main() {
    if cfg!(target_os = "windows") {
        let mut resource = winresource::WindowsResource::new();
        resource.set_icon("ico/evidence_tool.ico");
        resource.set("ProductName", "Evidence Tool Excel");
        resource.set("FileDescription", "Excel evidence capture tool");
        resource.set("LegalCopyright", "Copyright (c) 2026");
        resource
            .compile()
            .expect("failed to compile Windows resources");
    }
}
