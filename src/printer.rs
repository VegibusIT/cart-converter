/// Windows プリンター操作モジュール
/// - プリンター一覧の取得
/// - RAWデータ（SBPL）の直接送信

/// インストール済みプリンター名の一覧を取得
pub fn list_printers() -> Vec<String> {
    #[cfg(target_os = "windows")]
    {
        list_printers_windows()
    }
    #[cfg(not(target_os = "windows"))]
    {
        list_printers_unix()
    }
}

/// RAWデータをプリンターに直接送信
pub fn print_raw(printer_name: &str, data: &[u8]) -> Result<(), String> {
    #[cfg(target_os = "windows")]
    {
        print_raw_windows(printer_name, data)
    }
    #[cfg(not(target_os = "windows"))]
    {
        print_raw_unix(printer_name, data)
    }
}

// ============================================================
// Windows 実装
// ============================================================
#[cfg(target_os = "windows")]
#[repr(C)]
struct DocInfo1A {
    doc_name: *const u8,
    output_file: *const u8,
    datatype: *const u8,
}

#[cfg(target_os = "windows")]
extern "system" {
    fn OpenPrinterA(name: *const u8, handle: *mut isize, defaults: *const u8) -> i32;
    fn ClosePrinter(handle: isize) -> i32;
    fn StartDocPrinterA(handle: isize, level: u32, doc_info: *const DocInfo1A) -> u32;
    fn EndDocPrinter(handle: isize) -> i32;
    fn StartPagePrinter(handle: isize) -> i32;
    fn EndPagePrinter(handle: isize) -> i32;
    fn WritePrinter(handle: isize, buf: *const u8, count: u32, written: *mut u32) -> i32;
}

#[cfg(target_os = "windows")]
fn list_printers_windows() -> Vec<String> {
    // wmic は非推奨だが広く動作する。PowerShell より起動が速い。
    let output = std::process::Command::new("wmic")
        .args(["printer", "get", "name"])
        .output();

    match output {
        Ok(out) => {
            let text = String::from_utf8_lossy(&out.stdout);
            text.lines()
                .skip(1) // ヘッダー行 "Name" をスキップ
                .map(|l| l.trim().to_string())
                .filter(|l| !l.is_empty())
                .collect()
        }
        Err(_) => {
            // wmic が使えない場合は PowerShell にフォールバック
            let output = std::process::Command::new("powershell")
                .args(["-Command", "Get-Printer | Select-Object -ExpandProperty Name"])
                .output();
            match output {
                Ok(out) => {
                    let text = String::from_utf8_lossy(&out.stdout);
                    text.lines()
                        .map(|l| l.trim().to_string())
                        .filter(|l| !l.is_empty())
                        .collect()
                }
                Err(_) => Vec::new(),
            }
        }
    }
}

#[cfg(target_os = "windows")]
fn print_raw_windows(printer_name: &str, data: &[u8]) -> Result<(), String> {
    use std::ffi::CString;

    let c_printer = CString::new(printer_name).map_err(|e| format!("{e}"))?;
    let c_doc = CString::new("SCM Label").map_err(|e| format!("{e}"))?;
    let c_raw = CString::new("RAW").map_err(|e| format!("{e}"))?;

    unsafe {
        let mut handle: isize = 0;

        if OpenPrinterA(c_printer.as_ptr() as *const u8, &mut handle, std::ptr::null()) == 0 {
            return Err(format!("プリンター '{}' を開けません", printer_name));
        }

        let doc_info = DocInfo1A {
            doc_name: c_doc.as_ptr() as *const u8,
            output_file: std::ptr::null(),
            datatype: c_raw.as_ptr() as *const u8,
        };

        let job_id = StartDocPrinterA(handle, 1, &doc_info);
        if job_id == 0 {
            ClosePrinter(handle);
            return Err("印刷ジョブの開始に失敗しました".to_string());
        }

        if StartPagePrinter(handle) == 0 {
            EndDocPrinter(handle);
            ClosePrinter(handle);
            return Err("ページの開始に失敗しました".to_string());
        }

        let mut written: u32 = 0;
        let result = WritePrinter(handle, data.as_ptr(), data.len() as u32, &mut written);

        EndPagePrinter(handle);
        EndDocPrinter(handle);
        ClosePrinter(handle);

        if result == 0 {
            return Err("データの書き込みに失敗しました".to_string());
        }
    }

    Ok(())
}

// ============================================================
// Unix 実装 (Linux / macOS) — 開発・テスト用
// ============================================================
#[cfg(not(target_os = "windows"))]
fn list_printers_unix() -> Vec<String> {
    let output = std::process::Command::new("lpstat")
        .args(["-a"])
        .output();

    match output {
        Ok(out) => {
            let text = String::from_utf8_lossy(&out.stdout);
            text.lines()
                .filter_map(|l| l.split_whitespace().next())
                .map(|s| s.to_string())
                .collect()
        }
        Err(_) => vec!["(プリンター未検出)".to_string()],
    }
}

#[cfg(not(target_os = "windows"))]
fn print_raw_unix(printer_name: &str, data: &[u8]) -> Result<(), String> {
    use std::io::Write;
    use std::process::{Command, Stdio};

    let mut child = Command::new("lpr")
        .args(["-P", printer_name, "-o", "raw"])
        .stdin(Stdio::piped())
        .spawn()
        .map_err(|e| format!("lpr起動エラー: {e}"))?;

    child
        .stdin
        .as_mut()
        .unwrap()
        .write_all(data)
        .map_err(|e| format!("データ送信エラー: {e}"))?;

    let status = child.wait().map_err(|e| format!("{e}"))?;
    if !status.success() {
        return Err("lpr印刷に失敗しました".to_string());
    }
    Ok(())
}
