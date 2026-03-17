use eframe::egui;
use std::collections::HashSet;
use std::path::PathBuf;
use std::sync::{Arc, atomic::{AtomicBool, Ordering}};

use crate::google_auth::{self, AuthState, GoogleCredentials};
use crate::style::*;

// 設定用の定数
const SOURCE_FOLDER_ID: &str = "1N_a7WbcsYBsDCVr3LR1gNAXgx1dBzV9k";
const TARGET_SPREADSHEET_ID: &str = "1ga0RZkj9-LG75W4fqXWHaTLJhTOdzORSD03-7z5qwUc";

// --- Google API ヘルパー ---

/// Drive フォルダ内のファイル一覧を取得
fn list_drive_files(token: &str, folder_id: &str) -> Result<Vec<DriveFile>, String> {
    let query = format!("'{}' in parents and trashed = false", folder_id);
    let url = format!(
        "https://www.googleapis.com/drive/v3/files?q={}&fields=files(id,name,mimeType,modifiedTime)&orderBy=name desc&pageSize=200&supportsAllDrives=true&includeItemsFromAllDrives=true",
        urlencoding::encode(&query)
    );
    let resp = ureq::get(&url)
        .set("Authorization", &format!("Bearer {}", token))
        .call()
        .map_err(|e| format!("Drive API呼び出し失敗: {e}"))?;

    let raw = resp.into_string().map_err(|e| format!("レスポンス読み取り失敗: {e}"))?;
    let json: serde_json::Value = serde_json::from_str(&raw)
        .map_err(|e| format!("JSON解析失敗: {e}"))?;
    let files = json["files"]
        .as_array()
        .ok_or_else(|| format!("filesフィールドがありません: {}", raw))?;

    Ok(files
        .iter()
        .filter_map(|f| {
            Some(DriveFile {
                id: f["id"].as_str()?.to_string(),
                name: f["name"].as_str()?.to_string(),
                mime_type: f["mimeType"].as_str()?.to_string(),
                modified_time: f["modifiedTime"].as_str().unwrap_or("").to_string(),
            })
        })
        .collect())
}

/// Drive からファイルをダウンロード
fn download_drive_file(token: &str, file_id: &str) -> Result<Vec<u8>, String> {
    let url = format!(
        "https://www.googleapis.com/drive/v3/files/{}?alt=media",
        file_id
    );
    let resp = ureq::get(&url)
        .set("Authorization", &format!("Bearer {}", token))
        .call()
        .map_err(|e| format!("ダウンロード失敗: {e}"))?;

    let mut buf = Vec::new();
    resp.into_reader()
        .read_to_end(&mut buf)
        .map_err(|e| format!("読み取り失敗: {e}"))?;
    Ok(buf)
}

/// Sheets API: 既存データの行数を取得
fn get_sheet_row_count(token: &str, spreadsheet_id: &str, sheet_name: &str) -> Result<usize, String> {
    let range = format!("{}!A:A", sheet_name);
    let url = format!(
        "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}",
        spreadsheet_id,
        urlencoding::encode(&range)
    );
    let resp = ureq::get(&url)
        .set("Authorization", &format!("Bearer {}", token))
        .call()
        .map_err(|e| format!("API呼び出し失敗: {e}"))?;

    let json: serde_json::Value = resp.into_json().map_err(|e| format!("JSON解析失敗: {e}"))?;
    Ok(json["values"].as_array().map(|a| a.len()).unwrap_or(0))
}

/// Sheets API: データを追記
fn append_rows(
    token: &str,
    spreadsheet_id: &str,
    sheet_name: &str,
    start_row: usize,
    values: &[Vec<String>],
) -> Result<(), String> {
    let range = format!("{}!A{}:F{}", sheet_name, start_row, start_row + values.len());
    let url = format!(
        "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}?valueInputOption=USER_ENTERED",
        spreadsheet_id,
        urlencoding::encode(&range)
    );
    let body = serde_json::json!({
        "range": range,
        "majorDimension": "ROWS",
        "values": values,
    });
    ureq::put(&url)
        .set("Authorization", &format!("Bearer {}", token))
        .set("Content-Type", "application/json")
        .send_string(&body.to_string())
        .map_err(|e| format!("書き込み失敗: {e}"))?;
    Ok(())
}

// --- データ型 ---

#[derive(Clone, Debug)]
struct DriveFile {
    id: String,
    name: String,
    mime_type: String,
    modified_time: String,
}

/// 転記先に書き込む1行分のデータ
struct TransferRow {
    year_month: String,  // "202603"
    date: String,        // "2026年03月13日"
    jan_code: String,    // JANコード
    product_name: String, // 商品名
    quantity: String,     // 販売数量
    amount: String,       // 販売金額
}

/// 1ファイルから抽出した転記データ
struct FileTransferData {
    kikugawa_rows: Vec<TransferRow>,
    mori_rows: Vec<TransferRow>,
}

// --- 転記済み管理 ---

fn transferred_path() -> PathBuf {
    let base = if cfg!(target_os = "windows") {
        std::env::var_os("APPDATA")
            .map(PathBuf::from)
            .unwrap_or_else(|| PathBuf::from("."))
    } else if cfg!(target_os = "macos") {
        std::env::var_os("HOME")
            .map(|h| PathBuf::from(h).join("Library/Application Support"))
            .unwrap_or_else(|| PathBuf::from("."))
    } else {
        std::env::var_os("HOME")
            .map(|h| PathBuf::from(h).join(".config"))
            .unwrap_or_else(|| PathBuf::from("."))
    };
    base.join("cart-converter").join("entetu_transferred.json")
}

fn load_transferred() -> HashSet<String> {
    let Ok(s) = std::fs::read_to_string(transferred_path()) else {
        return HashSet::new();
    };
    serde_json::from_str(&s).unwrap_or_default()
}

fn save_transferred(ids: &HashSet<String>) {
    let dir = transferred_path().parent().unwrap().to_path_buf();
    let _ = std::fs::create_dir_all(&dir);
    if let Ok(json) = serde_json::to_string_pretty(ids) {
        let _ = std::fs::write(transferred_path(), json);
    }
}

// --- Excel解析ロジック ---

/// Excelファイルからデータを解析して転記データを生成
fn parse_excel_file(data: &[u8], file_name: &str) -> Result<FileTransferData, String> {
    use calamine::{Reader, Xlsx};
    use std::io::Cursor;

    let cursor = Cursor::new(data);
    let mut wb: Xlsx<_> = Xlsx::new(cursor)
        .map_err(|e| format!("Excel読み込み失敗: {e}"))?;

    let sheet_names = wb.sheet_names().to_vec();
    let is_yasaibus = file_name.contains("やさいバス");

    let mut result = FileTransferData {
        kikugawa_rows: Vec::new(),
        mori_rows: Vec::new(),
    };

    for sheet_name in &sheet_names {
        let range = wb.worksheet_range(sheet_name)
            .map_err(|e| format!("シート読み込み失敗: {e}"))?;

        // C5, E5 から日付を取得
        let date1 = get_cell_string(&range, 4, 2); // C5 (0-indexed: row=4, col=2)
        let date2 = get_cell_string(&range, 4, 4); // E5

        // 転記先シートの判別
        let target_is_mori = sheet_name.contains("森");

        // A8行〜 データ行を読み取り（A7は合計行なのでスキップ）
        for row_idx in 7..range.height() {
            let jan_code = get_cell_string(&range, row_idx as u32, 0);
            if jan_code.is_empty() || jan_code.starts_with("【") {
                continue;
            }
            let product_name = get_cell_string(&range, row_idx as u32, 1);

            // C列/D列: 1日目のデータ
            let qty1 = get_cell_string(&range, row_idx as u32, 2);
            let amt1 = get_cell_string(&range, row_idx as u32, 3);
            if !qty1.is_empty() || !amt1.is_empty() {
                let row = TransferRow {
                    year_month: extract_year_month(&date1),
                    date: clean_date(&date1),
                    jan_code: jan_code.clone(),
                    product_name: product_name.clone(),
                    quantity: qty1,
                    amount: amt1,
                };
                if target_is_mori {
                    result.mori_rows.push(row);
                } else {
                    result.kikugawa_rows.push(row);
                }
            }

            // E列/F列: 2日目のデータ
            let qty2 = get_cell_string(&range, row_idx as u32, 4);
            let amt2 = get_cell_string(&range, row_idx as u32, 5);
            if !qty2.is_empty() || !amt2.is_empty() {
                let row = TransferRow {
                    year_month: extract_year_month(&date2),
                    date: clean_date(&date2),
                    jan_code: jan_code.clone(),
                    product_name: product_name.clone(),
                    quantity: qty2,
                    amount: amt2,
                };
                if target_is_mori {
                    result.mori_rows.push(row);
                } else {
                    result.kikugawa_rows.push(row);
                }
            }
        }

        // やさいバスは菊川店のみ（森店シートなし）
        if is_yasaibus {
            // 全部菊川店に入れる（上のロジックでtarget_is_moriはfalse）
        }
    }

    Ok(result)
}

/// セル値を文字列として取得
fn get_cell_string(range: &calamine::Range<calamine::Data>, row: u32, col: u32) -> String {
    match range.get_value((row, col)) {
        Some(calamine::Data::String(s)) => s.clone(),
        Some(calamine::Data::Float(f)) => {
            if *f == (*f as i64) as f64 {
                (*f as i64).to_string()
            } else {
                f.to_string()
            }
        }
        Some(calamine::Data::Int(n)) => n.to_string(),
        Some(calamine::Data::DateTime(dt)) => {
            // Excel serial date
            format!("{}", dt.as_f64())
        }
        Some(calamine::Data::Bool(b)) => b.to_string(),
        _ => String::new(),
    }
}

/// ファイル名の日付が転記先に既に存在するかチェック
/// ファイル名例: "エムスクエアラボ日報03130314.xlsx" → 日付部分 "03130314" → "03月13日", "03月14日"
fn is_already_transferred(file_name: &str, existing_dates: &HashSet<String>) -> bool {
    let dates = extract_dates_from_filename(file_name);
    if dates.is_empty() {
        return false;
    }
    // 全日付が転記先に存在していれば転記済みと判定
    dates.iter().all(|d| {
        existing_dates.iter().any(|existing| existing.contains(d))
    })
}

/// ファイル名から日付部分を抽出
/// "エムスクエアラボ日報03130314.xlsx" → ["03月13日", "03月14日"]
/// "やさいバス日報0308.xlsx" → ["03月08日"]
fn extract_dates_from_filename(file_name: &str) -> Vec<String> {
    // "日報" の後の数字部分を抽出
    let name = file_name.trim_end_matches(".xlsx");
    let digits: String = name.chars()
        .rev()
        .take_while(|c| c.is_ascii_digit())
        .collect::<String>()
        .chars()
        .rev()
        .collect();

    let mut dates = Vec::new();
    match digits.len() {
        4 => {
            // "0308" → 1日分: "03月08日"
            let month = &digits[0..2];
            let day = &digits[2..4];
            dates.push(format!("{}月{}日", month, day));
        }
        8 => {
            // "03130314" → 2日分
            let m1 = &digits[0..2];
            let d1 = &digits[2..4];
            let m2 = &digits[4..6];
            let d2 = &digits[6..8];
            dates.push(format!("{}月{}日", m1, d1));
            dates.push(format!("{}月{}日", m2, d2));
        }
        _ => {}
    }
    dates
}

/// "2026年03月13日(金)" → "202603"
fn extract_year_month(date: &str) -> String {
    let date = date.trim();
    // "YYYY年MM月DD日" パターン
    if let Some(y_pos) = date.find('年') {
        if let Some(m_pos) = date.find('月') {
            let year = &date[..y_pos];
            let month = &date[y_pos + "年".len()..m_pos];
            return format!("{}{}", year, month.trim_start_matches('0'));
        }
    }
    String::new()
}

/// "2026年03月13日(金)" → "2026年03月13日"
fn clean_date(date: &str) -> String {
    let date = date.trim();
    if let Some(pos) = date.find('(') {
        date[..pos].trim().to_string()
    } else if let Some(pos) = date.find('（') {
        date[..pos].trim().to_string()
    } else {
        date.to_string()
    }
}

impl TransferRow {
    fn to_vec(&self) -> Vec<String> {
        vec![
            self.year_month.clone(),
            self.date.clone(),
            self.jan_code.clone(),
            self.product_name.clone(),
            self.quantity.clone(),
            self.amount.clone(),
        ]
    }
}

// --- UI ---

/// ファイル一覧の表示用
struct DisplayFile {
    file: DriveFile,
    transferred: bool,
    selected: bool,
}

pub struct EnteTuPage {
    auth_state: AuthState,
    credentials: Option<GoogleCredentials>,
    client_id_buf: String,
    client_secret_buf: String,
    show_credentials_form: bool,
    auth_cancel: Arc<AtomicBool>,

    // データ
    files: Vec<DisplayFile>,
    files_loaded: bool,
    transferred_ids: HashSet<String>,

    // 処理状態
    log: Vec<LogEntry>,
    progress: f32,
    is_processing: bool,
}

impl Default for EnteTuPage {
    fn default() -> Self {
        let credentials = google_auth::load_credentials()
            .unwrap_or_default();
        let auth_state = if !credentials.is_configured() {
            AuthState::NotAuthenticated
        } else {
            match google_auth::load_token() {
                Some(token) if google_auth::is_token_valid(&token) => AuthState::Authenticated(token),
                Some(token) => {
                    match google_auth::refresh_access_token(&credentials, &token) {
                        Ok(new_token) => AuthState::Authenticated(new_token),
                        Err(_) => AuthState::NotAuthenticated,
                    }
                }
                None => AuthState::NotAuthenticated,
            }
        };

        Self {
            auth_state,
            credentials: Some(credentials),
            client_id_buf: String::new(),
            auth_cancel: Arc::new(AtomicBool::new(false)),
            client_secret_buf: String::new(),
            show_credentials_form: false,
            files: Vec::new(),
            files_loaded: false,
            transferred_ids: load_transferred(),
            log: Vec::new(),
            progress: 0.0,
            is_processing: false,
        }
    }
}

impl EnteTuPage {
    pub fn show(&mut self, ctx: &egui::Context) -> bool {
        let mut go_back = false;

        egui::CentralPanel::default()
            .frame(
                egui::Frame::default()
                    .fill(BG)
                    .inner_margin(egui::Margin::symmetric(32, 28)),
            )
            .show(ctx, |ui| {
                if ui
                    .add(
                        egui::Button::new(
                            egui::RichText::new("< トップへ戻る").size(12.0).color(ACCENT),
                        )
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::NONE),
                    )
                    .clicked()
                {
                    go_back = true;
                }

                ui.add_space(8.0);
                ui.label(
                    egui::RichText::new("遠鉄ストア 消化仕入れデータ転記")
                        .size(16.0)
                        .strong()
                        .color(TEXT_PRIMARY),
                );
                ui.add_space(4.0);
                ui.label(
                    egui::RichText::new("Google Driveのメールデータをスプレッドシートに転記します")
                        .size(12.0)
                        .color(TEXT_SECONDARY),
                );
                ui.add_space(16.0);

                match &self.auth_state {
                    AuthState::NotAuthenticated | AuthState::Error(_) => {
                        self.show_auth_section(ui);
                    }
                    AuthState::WaitingForCallback => {
                        // トークンが保存されたか定期チェック
                        if let Some(token) = google_auth::load_token() {
                            if google_auth::is_token_valid(&token) {
                                self.auth_state = AuthState::Authenticated(token);
                                // 次のフレームで再描画
                                ui.ctx().request_repaint();
                                return;
                            }
                        }
                        ui.label(
                            egui::RichText::new("ブラウザで認証を完了してください...")
                                .size(13.0)
                                .color(TEXT_SECONDARY),
                        );
                        ui.add_space(8.0);
                        // 定期的に再描画してトークンチェック
                        ui.ctx().request_repaint_after(std::time::Duration::from_secs(1));
                        ui.horizontal(|ui| {
                            ui.spinner();
                            ui.add_space(12.0);
                            if ui
                                .add(
                                    egui::Button::new(
                                        egui::RichText::new("キャンセル")
                                            .size(13.0)
                                            .color(ERROR),
                                    )
                                    .fill(SURFACE)
                                    .stroke(egui::Stroke::new(1.0, ERROR))
                                    .corner_radius(egui::CornerRadius::same(8)),
                                )
                                .clicked()
                            {
                                self.auth_cancel.store(true, Ordering::Relaxed);
                                self.auth_state = AuthState::NotAuthenticated;
                            }
                        });
                    }
                    AuthState::Authenticated(_) => {
                        self.show_main_section(ui, ctx);
                    }
                }
            });

        go_back
    }

    fn show_auth_section(&mut self, ui: &mut egui::Ui) {
        if let AuthState::Error(msg) = &self.auth_state {
            ui.label(
                egui::RichText::new(format!("認証エラー: {msg}"))
                    .size(12.0)
                    .color(ERROR),
            );
            ui.add_space(8.0);
        }

        let needs_creds = self.credentials.as_ref().map(|c| !c.is_configured()).unwrap_or(true);
        if needs_creds || self.show_credentials_form {
            egui::Frame::default()
                .fill(SURFACE)
                .corner_radius(egui::CornerRadius::same(8))
                .stroke(egui::Stroke::new(1.0, BORDER))
                .inner_margin(egui::Margin::symmetric(16, 12))
                .show(ui, |ui| {
                    ui.label(
                        egui::RichText::new("Google OAuth クレデンシャル設定")
                            .size(13.0)
                            .strong()
                            .color(TEXT_PRIMARY),
                    );
                    ui.add_space(8.0);
                    ui.label(egui::RichText::new("Client ID").size(12.0).color(TEXT_SECONDARY));
                    ui.add(
                        egui::TextEdit::singleline(&mut self.client_id_buf)
                            .desired_width(ui.available_width() - 20.0)
                            .font(egui::TextStyle::Body),
                    );
                    ui.add_space(4.0);
                    ui.label(egui::RichText::new("Client Secret").size(12.0).color(TEXT_SECONDARY));
                    ui.add(
                        egui::TextEdit::singleline(&mut self.client_secret_buf)
                            .desired_width(ui.available_width() - 20.0)
                            .password(true)
                            .font(egui::TextStyle::Body),
                    );
                    ui.add_space(8.0);
                    if ui
                        .add(
                            egui::Button::new(
                                egui::RichText::new("保存").size(13.0).color(egui::Color32::WHITE),
                            )
                            .fill(ACCENT)
                            .corner_radius(egui::CornerRadius::same(8)),
                        )
                        .clicked()
                    {
                        if !self.client_id_buf.is_empty() && !self.client_secret_buf.is_empty() {
                            let creds = GoogleCredentials {
                                client_id: self.client_id_buf.clone(),
                                client_secret: self.client_secret_buf.clone(),
                            };
                            google_auth::save_credentials(&creds);
                            self.credentials = Some(creds);
                            self.show_credentials_form = false;
                        }
                    }
                });
            ui.add_space(8.0);
        }

        if let Some(creds) = &self.credentials {
            if !creds.is_configured() {
                return;
            }
            ui.horizontal(|ui| {
                if ui
                    .add(
                        egui::Button::new(
                            egui::RichText::new("Googleでログイン")
                                .size(14.0)
                                .color(egui::Color32::WHITE),
                        )
                        .fill(ACCENT)
                        .min_size(egui::vec2(200.0, 40.0))
                        .corner_radius(egui::CornerRadius::same(8)),
                    )
                    .clicked()
                {
                    let auth_url = google_auth::build_auth_url(creds);
                    let _ = open::that(&auth_url);
                    self.auth_cancel.store(false, Ordering::Relaxed);
                    self.auth_state = AuthState::WaitingForCallback;
                    let creds_clone = creds.clone();
                    let cancel = self.auth_cancel.clone();
                    let ctx = ui.ctx().clone();
                    std::thread::spawn(move || {
                        match google_auth::wait_for_callback_and_exchange(&creds_clone, &cancel) {
                            Ok(_) | Err(_) => { ctx.request_repaint(); }
                        }
                    });
                }
                ui.add_space(8.0);
                if ui
                    .add(
                        egui::Button::new(
                            egui::RichText::new("設定変更").size(12.0).color(TEXT_SECONDARY),
                        )
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::new(1.0, BORDER))
                        .corner_radius(egui::CornerRadius::same(8)),
                    )
                    .clicked()
                {
                    self.show_credentials_form = true;
                    if let Some(c) = &self.credentials {
                        self.client_id_buf = c.client_id.clone();
                        self.client_secret_buf = c.client_secret.clone();
                    }
                }
            });
        }
    }

    fn show_main_section(&mut self, ui: &mut egui::Ui, ctx: &egui::Context) {
        // トークンリフレッシュ
        if let AuthState::Authenticated(ref token) = self.auth_state {
            if !google_auth::is_token_valid(token) {
                if let Some(creds) = &self.credentials {
                    match google_auth::refresh_access_token(creds, token) {
                        Ok(new_token) => {
                            self.auth_state = AuthState::Authenticated(new_token);
                        }
                        Err(_) => {
                            self.auth_state = AuthState::NotAuthenticated;
                            return;
                        }
                    }
                }
            }
        }

        ui.horizontal(|ui| {
            ui.label(egui::RichText::new("Google認証済み").size(12.0).color(SUCCESS));
            if ui
                .add(
                    egui::Button::new(egui::RichText::new("ログアウト").size(11.0).color(ERROR))
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::new(1.0, ERROR))
                        .corner_radius(egui::CornerRadius::same(6)),
                )
                .clicked()
            {
                google_auth::clear_token();
                self.auth_state = AuthState::NotAuthenticated;
                self.files.clear();
                self.files_loaded = false;
                return;
            }
        });

        ui.add_space(12.0);

        // データ読み込みボタン（常に表示して再読み込み可能に）
        if ui
            .add(
                egui::Button::new(
                    egui::RichText::new(if self.files_loaded { "再読み込み" } else { "データを読み込む" })
                        .size(13.0)
                        .color(egui::Color32::WHITE),
                )
                .fill(ACCENT)
                .corner_radius(egui::CornerRadius::same(8)),
            )
            .clicked()
        {
            self.load_data();
        }

        ui.add_space(12.0);

        // ファイル一覧
        if self.files_loaded {
            let new_count = self.files.iter().filter(|f| !f.transferred).count();
            let total_count = self.files.len();

            if new_count > 0 {
                ui.label(
                    egui::RichText::new(format!("未転記: {}件 / 全{}件", new_count, total_count))
                        .size(14.0)
                        .strong()
                        .color(ACCENT),
                );
            } else {
                ui.label(
                    egui::RichText::new(format!("全{}件 転記済み", total_count))
                        .size(14.0)
                        .color(TEXT_SECONDARY),
                );
            }
            ui.add_space(4.0);

            // 未転記のみ選択 / 全選択 / 全解除
            ui.horizontal(|ui| {
                if ui.add(
                    egui::Button::new(egui::RichText::new("未転記を選択").size(11.0).color(ACCENT))
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::new(1.0, ACCENT))
                        .corner_radius(egui::CornerRadius::same(4)),
                ).clicked() {
                    for f in &mut self.files {
                        f.selected = !f.transferred;
                    }
                }
                if ui.add(
                    egui::Button::new(egui::RichText::new("全選択").size(11.0).color(TEXT_SECONDARY))
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::new(1.0, BORDER))
                        .corner_radius(egui::CornerRadius::same(4)),
                ).clicked() {
                    for f in &mut self.files { f.selected = true; }
                }
                if ui.add(
                    egui::Button::new(egui::RichText::new("全解除").size(11.0).color(TEXT_SECONDARY))
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::new(1.0, BORDER))
                        .corner_radius(egui::CornerRadius::same(4)),
                ).clicked() {
                    for f in &mut self.files { f.selected = false; }
                }
                let sel = self.files.iter().filter(|f| f.selected).count();
                ui.label(egui::RichText::new(format!("{}件選択中", sel)).size(11.0).color(TEXT_SECONDARY));
            });
            ui.add_space(4.0);

            egui::ScrollArea::vertical()
                .id_salt("source_files")
                .max_height(200.0)
                .show(ui, |ui| {
                    for file in &mut self.files {
                        ui.horizontal(|ui| {
                            ui.checkbox(&mut file.selected, "");
                            let name_color = if file.transferred { TEXT_SECONDARY } else { TEXT_PRIMARY };
                            ui.label(egui::RichText::new(&file.file.name).size(12.0).color(name_color));
                            if file.transferred {
                                ui.label(egui::RichText::new("転記済み").size(10.0).color(TEXT_SECONDARY));
                            } else {
                                ui.label(egui::RichText::new("未転記").size(10.0).color(ACCENT));
                            }
                        });
                    }
                });

            ui.add_space(12.0);

            // 転記実行ボタン
            if !self.is_processing {
                let btn = ui.add(
                    egui::Button::new(
                        egui::RichText::new("転記実行")
                            .size(15.0)
                            .color(egui::Color32::WHITE)
                            .strong(),
                    )
                    .fill(ACCENT)
                    .min_size(egui::vec2(ui.available_width(), 44.0))
                    .corner_radius(egui::CornerRadius::same(10)),
                );
                if btn.hovered() {
                    ctx.set_cursor_icon(egui::CursorIcon::PointingHand);
                }
                if btn.clicked() {
                    self.run_transfer();
                }
            }
        }

        ui.add_space(12.0);

        // Progress bar
        if self.progress > 0.0 {
            let (rect, _) = ui.allocate_exact_size(
                egui::vec2(ui.available_width(), 4.0),
                egui::Sense::hover(),
            );
            ui.painter().rect_filled(rect, egui::CornerRadius::same(2), PROGRESS_BG);
            let fill_rect = egui::Rect::from_min_size(
                rect.min,
                egui::vec2(rect.width() * self.progress, rect.height()),
            );
            ui.painter().rect_filled(fill_rect, egui::CornerRadius::same(2), ACCENT);
            ui.add_space(16.0);
        }

        show_log(ui, &self.log);
    }

    fn get_token(&self) -> Option<&str> {
        if let AuthState::Authenticated(ref token) = self.auth_state {
            Some(&token.access_token)
        } else {
            None
        }
    }

    fn load_data(&mut self) {
        let token = match self.get_token() {
            Some(t) => t.to_string(),
            None => {
                self.log.push(LogEntry { text: "認証が必要です".into(), kind: LogKind::Error });
                return;
            }
        };

        self.log.clear();
        self.log.push(LogEntry { text: "データを読み込んでいます...".into(), kind: LogKind::Info });

        // 転記先シートの既存日付を取得して重複チェック用セットを構築
        self.log.push(LogEntry { text: "転記先の既存データを確認中...".into(), kind: LogKind::Info });
        let existing_dates = self.load_existing_dates(&token);
        self.log.push(LogEntry {
            text: format!("転記先に{}件の日付データを確認", existing_dates.len()),
            kind: LogKind::Ok,
        });

        match list_drive_files(&token, SOURCE_FOLDER_ID) {
            Ok(drive_files) => {
                // xlsxファイルのみ
                let xlsx_files: Vec<DriveFile> = drive_files
                    .into_iter()
                    .filter(|f| f.name.ends_with(".xlsx"))
                    .collect();

                self.log.push(LogEntry {
                    text: format!("{}件のファイルを検出", xlsx_files.len()),
                    kind: LogKind::Ok,
                });

                // 転記済みチェック（ローカル記録 + 転記先データの日付照合）
                self.transferred_ids = load_transferred();

                self.files = xlsx_files
                    .into_iter()
                    .map(|f| {
                        let transferred = self.transferred_ids.contains(&f.id)
                            || is_already_transferred(&f.name, &existing_dates);
                        // ローカル記録にない既存データもIDを追加しておく
                        if transferred && !self.transferred_ids.contains(&f.id) {
                            self.transferred_ids.insert(f.id.clone());
                        }
                        DisplayFile { file: f, transferred, selected: false }
                    })
                    .collect();

                // ローカル記録を更新
                save_transferred(&self.transferred_ids);

                // 未転記ファイルを自動選択
                let new_count = self.files.iter().filter(|f| !f.transferred).count();
                for f in &mut self.files {
                    if !f.transferred {
                        f.selected = true;
                    }
                }

                self.files_loaded = true;

                if new_count > 0 {
                    self.log.push(LogEntry {
                        text: format!("未転記: {}件（自動選択済み）", new_count),
                        kind: LogKind::Ok,
                    });
                } else {
                    self.log.push(LogEntry {
                        text: "全ファイル転記済みです".into(),
                        kind: LogKind::Done,
                    });
                }
            }
            Err(e) => {
                self.log.push(LogEntry {
                    text: format!("フォルダ読み取りエラー: {e}"),
                    kind: LogKind::Error,
                });
            }
        }
    }

    /// 転記先シートの既存日付を読み取る（B列＝日付）
    fn load_existing_dates(&self, token: &str) -> HashSet<String> {
        let mut dates = HashSet::new();
        for sheet in &["菊川店", "森店"] {
            let range = format!("{}!B:B", sheet);
            let url = format!(
                "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}",
                TARGET_SPREADSHEET_ID,
                urlencoding::encode(&range)
            );
            if let Ok(resp) = ureq::get(&url)
                .set("Authorization", &format!("Bearer {}", token))
                .call()
            {
                if let Ok(json) = resp.into_json::<serde_json::Value>() {
                    if let Some(rows) = json["values"].as_array() {
                        for row in rows {
                            if let Some(cells) = row.as_array() {
                                if let Some(date) = cells.first().and_then(|c| c.as_str()) {
                                    dates.insert(date.to_string());
                                }
                            }
                        }
                    }
                }
            }
        }
        dates
    }

    fn run_transfer(&mut self) {
        let token = match self.get_token() {
            Some(t) => t.to_string(),
            None => {
                self.log.push(LogEntry { text: "認証が必要です".into(), kind: LogKind::Error });
                return;
            }
        };

        let selected: Vec<DriveFile> = self.files.iter()
            .filter(|f| f.selected)
            .map(|f| f.file.clone())
            .collect();

        if selected.is_empty() {
            self.log.push(LogEntry { text: "転記するファイルを選択してください".into(), kind: LogKind::Error });
            return;
        }

        self.log.clear();
        self.progress = 0.0;
        self.is_processing = true;

        self.log.push(LogEntry {
            text: format!("転記処理を開始（{}件）...", selected.len()),
            kind: LogKind::Info,
        });

        let total = selected.len();
        let mut total_kikugawa = 0usize;
        let mut total_mori = 0usize;
        let mut errors = 0usize;
        let mut newly_transferred: Vec<String> = Vec::new();

        // 全ファイルのデータを先に集める
        let mut all_kikugawa: Vec<Vec<String>> = Vec::new();
        let mut all_mori: Vec<Vec<String>> = Vec::new();

        for (idx, file) in selected.iter().enumerate() {
            self.log.push(LogEntry {
                text: format!("  ダウンロード中: {}", file.name),
                kind: LogKind::Info,
            });

            match download_drive_file(&token, &file.id) {
                Ok(data) => {
                    match parse_excel_file(&data, &file.name) {
                        Ok(transfer_data) => {
                            let k_count = transfer_data.kikugawa_rows.len();
                            let m_count = transfer_data.mori_rows.len();

                            for row in &transfer_data.kikugawa_rows {
                                all_kikugawa.push(row.to_vec());
                            }
                            for row in &transfer_data.mori_rows {
                                all_mori.push(row.to_vec());
                            }

                            total_kikugawa += k_count;
                            total_mori += m_count;
                            newly_transferred.push(file.id.clone());

                            self.log.push(LogEntry {
                                text: format!("    菊川店: {}行, 森店: {}行", k_count, m_count),
                                kind: LogKind::Ok,
                            });
                        }
                        Err(e) => {
                            errors += 1;
                            self.log.push(LogEntry {
                                text: format!("    解析エラー: {e}"),
                                kind: LogKind::Error,
                            });
                        }
                    }
                }
                Err(e) => {
                    errors += 1;
                    self.log.push(LogEntry {
                        text: format!("    ダウンロードエラー: {e}"),
                        kind: LogKind::Error,
                    });
                }
            }

            self.progress = (idx + 1) as f32 / (total as f32 * 2.0); // 前半50%
        }

        // 菊川店に書き込み
        if !all_kikugawa.is_empty() {
            self.log.push(LogEntry {
                text: format!("菊川店シートに{}行を書き込み中...", all_kikugawa.len()),
                kind: LogKind::Info,
            });
            match get_sheet_row_count(&token, TARGET_SPREADSHEET_ID, "菊川店") {
                Ok(existing) => {
                    let start_row = existing + 1;
                    match append_rows(&token, TARGET_SPREADSHEET_ID, "菊川店", start_row, &all_kikugawa) {
                        Ok(()) => {
                            self.log.push(LogEntry {
                                text: format!("  菊川店: {}行を書き込み完了", all_kikugawa.len()),
                                kind: LogKind::Ok,
                            });
                        }
                        Err(e) => {
                            self.log.push(LogEntry {
                                text: format!("  菊川店 書き込みエラー: {e}"),
                                kind: LogKind::Error,
                            });
                            errors += 1;
                        }
                    }
                }
                Err(e) => {
                    self.log.push(LogEntry {
                        text: format!("  菊川店 行数取得エラー: {e}"),
                        kind: LogKind::Error,
                    });
                    errors += 1;
                }
            }
        }
        self.progress = 0.75;

        // 森店に書き込み
        if !all_mori.is_empty() {
            self.log.push(LogEntry {
                text: format!("森店シートに{}行を書き込み中...", all_mori.len()),
                kind: LogKind::Info,
            });
            match get_sheet_row_count(&token, TARGET_SPREADSHEET_ID, "森店") {
                Ok(existing) => {
                    let start_row = existing + 1;
                    match append_rows(&token, TARGET_SPREADSHEET_ID, "森店", start_row, &all_mori) {
                        Ok(()) => {
                            self.log.push(LogEntry {
                                text: format!("  森店: {}行を書き込み完了", all_mori.len()),
                                kind: LogKind::Ok,
                            });
                        }
                        Err(e) => {
                            self.log.push(LogEntry {
                                text: format!("  森店 書き込みエラー: {e}"),
                                kind: LogKind::Error,
                            });
                            errors += 1;
                        }
                    }
                }
                Err(e) => {
                    self.log.push(LogEntry {
                        text: format!("  森店 行数取得エラー: {e}"),
                        kind: LogKind::Error,
                    });
                    errors += 1;
                }
            }
        }

        // 転記済みとして記録
        if errors == 0 {
            for id in &newly_transferred {
                self.transferred_ids.insert(id.clone());
            }
            save_transferred(&self.transferred_ids);

            // UI上の状態も更新
            for f in &mut self.files {
                if newly_transferred.contains(&f.file.id) {
                    f.transferred = true;
                    f.selected = false;
                }
            }
        }

        self.progress = 1.0;
        self.is_processing = false;

        self.log.push(LogEntry {
            text: format!(
                "完了 -- 菊川店: {}行, 森店: {}行, エラー: {}件",
                total_kikugawa, total_mori, errors
            ),
            kind: if errors == 0 { LogKind::Done } else { LogKind::Error },
        });
    }
}
