use eframe::egui;
use std::path::PathBuf;
use std::sync::{
    atomic::{AtomicBool, Ordering},
    Arc,
};

use crate::google_auth::{self, AuthState, GoogleCredentials};
use crate::style::*;

// --- Google Sheets API ヘルパー ---

/// スプレッドシートURLまたはIDからIDを抽出
fn extract_spreadsheet_id(input: &str) -> String {
    let input = input.trim();
    // URL: https://docs.google.com/spreadsheets/d/XXXXX/edit...
    if input.contains("/spreadsheets/d/") {
        if let Some(after) = input.split("/spreadsheets/d/").nth(1) {
            return after.split('/').next().unwrap_or(after).to_string();
        }
    }
    // そのままIDとして扱う
    input.to_string()
}

/// Sheets API でシートの全データを取得
fn read_sheet_values(
    token: &str,
    spreadsheet_id: &str,
    range: &str,
) -> Result<Vec<Vec<String>>, String> {
    let url = format!(
        "https://sheets.googleapis.com/v4/spreadsheets/{}/values/{}",
        spreadsheet_id,
        urlencoding::encode(range)
    );
    let resp = ureq::get(&url)
        .set("Authorization", &format!("Bearer {}", token))
        .call()
        .map_err(|e| format!("Sheets API呼び出し失敗: {e}"))?;

    let json: serde_json::Value = resp
        .into_json()
        .map_err(|e| format!("JSON解析失敗: {e}"))?;

    let rows = json["values"]
        .as_array()
        .ok_or("データが見つかりません")?;

    Ok(rows
        .iter()
        .map(|row| {
            row.as_array()
                .map(|cells| {
                    cells
                        .iter()
                        .map(|c| c.as_str().unwrap_or("").to_string())
                        .collect()
                })
                .unwrap_or_default()
        })
        .collect())
}

// --- データ構造 ---

struct PayeeInfo {
    name: String,
    code: String,
}

struct OriginalRow {
    user_id: String,
    producer_name: String,
    price_8: i64,
    total_payment: i64,
}

/// 支払先一覧スプレッドシートを読み込み
fn read_payee_from_sheets(
    token: &str,
    spreadsheet_id: &str,
) -> Result<std::collections::HashMap<String, PayeeInfo>, String> {
    let rows = read_sheet_values(token, spreadsheet_id, "A:Z")?;
    if rows.is_empty() {
        return Err("データが空です".into());
    }

    // ヘッダー行を探す
    let mut col_cid = None;
    let mut col_name = None;
    let mut col_code = None;
    let mut header_row = 0;

    for (row_idx, row) in rows.iter().enumerate().take(5) {
        for (col_idx, cell) in row.iter().enumerate() {
            match cell.trim() {
                "customer_id" => {
                    col_cid = Some(col_idx);
                    header_row = row_idx;
                }
                "支払先名" => col_name = Some(col_idx),
                "支払先コード" => col_code = Some(col_idx),
                _ => {}
            }
        }
    }

    let c_cid = col_cid.ok_or("customer_id列が見つかりません")?;
    let c_name = col_name.ok_or("支払先名列が見つかりません")?;
    let c_code = col_code.ok_or("支払先コード列が見つかりません")?;

    let mut map = std::collections::HashMap::new();
    for row in rows.iter().skip(header_row + 1) {
        let cid_raw = row.get(c_cid).map(|s| s.trim()).unwrap_or("");
        if cid_raw.is_empty() {
            continue;
        }
        // "1234.0" → "1234"
        let cid = if let Ok(f) = cid_raw.parse::<f64>() {
            format!("{}", f as i64)
        } else {
            cid_raw.to_string()
        };

        let name = row.get(c_name).map(|s| s.trim().to_string()).unwrap_or_default();
        let code = row.get(c_code).map(|s| s.trim().to_string()).unwrap_or_default();

        if !name.is_empty() && !code.is_empty() {
            map.insert(cid, PayeeInfo { name, code });
        }
    }

    Ok(map)
}

/// 生産者手取り額スプレッドシートを読み込み
fn read_original_from_sheets(
    token: &str,
    spreadsheet_id: &str,
) -> Result<Vec<OriginalRow>, String> {
    let rows = read_sheet_values(token, spreadsheet_id, "A:Z")?;
    if rows.is_empty() {
        return Err("データが空です".into());
    }

    // ヘッダー行を探す
    let mut col_uid = None;
    let mut col_prod = None;
    let mut col_8 = None;
    let mut col_total = None;
    let mut header_row = 0;

    for (row_idx, row) in rows.iter().enumerate().take(5) {
        for (col_idx, cell) in row.iter().enumerate() {
            let t = cell.trim();
            if t.contains("ユーザーID") {
                col_uid = Some(col_idx);
                header_row = row_idx;
            } else if t.contains("生産者名") {
                col_prod = Some(col_idx);
            } else if t.contains("生産者価格") && t.contains("8%") {
                col_8 = Some(col_idx);
            } else if t.contains("お支払合計額") {
                col_total = Some(col_idx);
            }
        }
    }

    let c_uid = col_uid.ok_or("ユーザーID列が見つかりません")?;
    let c_prod = col_prod.ok_or("生産者名列が見つかりません")?;
    let c_8 = col_8.ok_or("生産者価格(8%)列が見つかりません")?;
    let c_total = col_total.ok_or("お支払合計額列が見つかりません")?;

    let mut result = Vec::new();
    for row in rows.iter().skip(header_row + 1) {
        let uid = row.get(c_uid).map(|s| s.trim()).unwrap_or("");
        if uid.is_empty() {
            continue;
        }
        let user_id = if let Ok(f) = uid.parse::<f64>() {
            format!("{}", f as i64)
        } else {
            uid.to_string()
        };

        let producer_name = row.get(c_prod).map(|s| s.trim().to_string()).unwrap_or_default();
        let price_8 = parse_i64(row.get(c_8).map(|s| s.as_str()).unwrap_or(""));
        let total = parse_i64(row.get(c_total).map(|s| s.as_str()).unwrap_or(""));

        // 8%が0なら10%を確認（両方0なら除外）
        if price_8 == 0 && total == 0 {
            continue;
        }

        result.push(OriginalRow {
            user_id,
            producer_name,
            price_8,
            total_payment: total,
        });
    }

    Ok(result)
}

fn parse_i64(s: &str) -> i64 {
    let s = s.trim().replace(',', "");
    s.parse::<f64>().unwrap_or(0.0) as i64
}

/// 前月末日を "YYYY/MM/DD" で返す
fn last_day_of_prev_month() -> String {
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs() as i64;
    let (_y, m, _d) = crate::convert::unix_days_to_ymd(now / 86400);

    let (prev_y, prev_m) = if m == 1 { (_y - 1, 12) } else { (_y, m - 1) };
    let last_day = days_in_month(prev_y, prev_m);

    format!("{}/{:02}/{:02}", prev_y, prev_m, last_day)
}

fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if (year % 4 == 0 && year % 100 != 0) || year % 400 == 0 {
                29
            } else {
                28
            }
        }
        _ => 30,
    }
}

struct ConvertResult {
    file_count: usize,
    file_rows: Vec<(String, usize)>,
    unregistered: Vec<(String, String)>,
}

fn run_convert(
    token: &str,
    original_id: &str,
    payee_id: &str,
    payment_day: &str,
    output_dir: &std::path::Path,
) -> Result<ConvertResult, String> {
    let payee_map = read_payee_from_sheets(token, payee_id)?;
    let rows = read_original_from_sheets(token, original_id)?;
    let transaction_day = last_day_of_prev_month();

    let mut csv_rows: Vec<String> = Vec::new();
    let mut unregistered: Vec<(String, String)> = Vec::new();

    let header =
        "支払先名,支払先コード,取引日,購入分類コード,取引内容,請求金額,決済予定日,決済分類コード,備考";

    for row in &rows {
        if let Some(payee) = payee_map.get(&row.user_id) {
            let content = if row.price_8 != 0 {
                "商品(軽8%)"
            } else {
                "課仕 10%"
            };
            csv_rows.push(format!(
                "{},{},{},001,{},{},{},001,",
                payee.name, payee.code, transaction_day, content, row.total_payment, payment_day,
            ));
        } else {
            unregistered.push((row.user_id.clone(), row.producer_name.clone()));
        }
    }

    std::fs::create_dir_all(output_dir)
        .map_err(|e| format!("出力フォルダを作成できません: {e}"))?;

    // 100行ずつ分割して出力
    let chunk_size = 100;
    let mut file_count = 0;
    let mut file_rows_info = Vec::new();

    for (i, chunk) in csv_rows.chunks(chunk_size).enumerate() {
        let filename = format!("money_forward_import_list_{}.csv", i + 1);
        let filepath = output_dir.join(&filename);
        let content = format!("{}\n{}", header, chunk.join("\n"));
        std::fs::write(&filepath, content.as_bytes())
            .map_err(|e| format!("{} の書き込みに失敗: {e}", filename))?;
        file_rows_info.push((filename, chunk.len()));
        file_count += 1;
    }

    // 未登録ユーザーCSV
    if !unregistered.is_empty() {
        let unreg_path = output_dir.join("non_registered_user_id.csv");
        let mut content = String::from("ユーザーID,生産者名\n");
        for (uid, name) in &unregistered {
            content.push_str(&format!("{},{}\n", uid, name));
        }
        std::fs::write(&unreg_path, content.as_bytes())
            .map_err(|e| format!("未登録ユーザーファイルの書き込みに失敗: {e}"))?;
    }

    Ok(ConvertResult {
        file_count,
        file_rows: file_rows_info,
        unregistered,
    })
}

// --- 設定の永続化 ---

#[derive(serde::Serialize, serde::Deserialize, Default)]
struct MfSettings {
    original_url: Option<String>,
    payee_url: Option<String>,
    output_dir: Option<String>,
    payment_day: Option<String>,
}

fn settings_path() -> PathBuf {
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
    base.join("cart-converter").join("mf-settings.json")
}

fn load_settings() -> MfSettings {
    if let Ok(s) = std::fs::read_to_string(settings_path()) {
        serde_json::from_str(&s).unwrap_or_default()
    } else {
        MfSettings::default()
    }
}

fn save_settings(settings: &MfSettings) {
    if let Some(parent) = settings_path().parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(json) = serde_json::to_string_pretty(settings) {
        let _ = std::fs::write(settings_path(), json);
    }
}

// --- UI ---

pub struct MoneyForwardPage {
    // Google認証
    credentials: Option<GoogleCredentials>,
    auth_state: AuthState,
    auth_cancel: Arc<AtomicBool>,

    // 入力
    original_url: String,
    payee_url: String,
    payment_day: String,
    output_dir: Option<PathBuf>,

    // ログ
    log: Vec<LogEntry>,
    is_done: bool,
}

impl Default for MoneyForwardPage {
    fn default() -> Self {
        let settings = load_settings();
        let creds = google_auth::load_credentials()
            .unwrap_or_default();
        let auth_state = match google_auth::load_token() {
            Some(token) if google_auth::is_token_valid(&token) => {
                AuthState::Authenticated(token)
            }
            Some(token) => {
                // トークン期限切れ → リフレッシュ試行
                match google_auth::refresh_access_token(&creds, &token) {
                    Ok(new_token) => AuthState::Authenticated(new_token),
                    Err(_) => AuthState::NotAuthenticated,
                }
            }
            None => AuthState::NotAuthenticated,
        };

        Self {
            credentials: Some(creds),
            auth_state,
            auth_cancel: Arc::new(AtomicBool::new(false)),
            original_url: settings.original_url.unwrap_or_default(),
            payee_url: settings.payee_url.unwrap_or_else(|| {
                // デフォルト：支払先一覧スプレッドシート
                "https://docs.google.com/spreadsheets/d/1IdCd4QCvmJNeDHGt7qDRBt17Q8_TcJHQDRFBIezhKu8/edit#gid=0".to_string()
            }),
            payment_day: settings.payment_day.unwrap_or_default(),
            output_dir: settings.output_dir.map(PathBuf::from),
            log: Vec::new(),
            is_done: false,
        }
    }
}

impl MoneyForwardPage {
    pub fn show(&mut self, ctx: &egui::Context) -> bool {
        let go_back = false;

        egui::CentralPanel::default()
            .frame(
                egui::Frame::default()
                    .fill(BG)
                    .inner_margin(egui::Margin::symmetric(32, 28)),
            )
            .show(ctx, |ui| {
                ui.add_space(12.0);

                ui.label(
                    egui::RichText::new("MoneyForward 生産者入金データ変換")
                        .size(18.0)
                        .strong()
                        .color(TEXT_PRIMARY),
                );
                ui.add_space(4.0);
                ui.label(
                    egui::RichText::new(
                        "Googleスプレッドシートから直接読み込み → MoneyForwardインポート用CSVを生成",
                    )
                    .size(12.0)
                    .color(TEXT_SECONDARY),
                );
                ui.add_space(16.0);

                match &self.auth_state {
                    AuthState::NotAuthenticated => {
                        self.show_login_section(ui);
                    }
                    AuthState::WaitingForCallback => {
                        ui.label(
                            egui::RichText::new("ブラウザで認証中...")
                                .size(13.0)
                                .color(TEXT_SECONDARY),
                        );
                        ui.add_space(8.0);
                        if ui
                            .add(
                                egui::Button::new(
                                    egui::RichText::new("キャンセル")
                                        .size(12.0)
                                        .color(ERROR),
                                )
                                .fill(egui::Color32::TRANSPARENT)
                                .stroke(egui::Stroke::new(1.0, ERROR))
                                .corner_radius(egui::CornerRadius::same(6)),
                            )
                            .clicked()
                        {
                            self.auth_cancel.store(true, Ordering::Relaxed);
                            self.auth_state = AuthState::NotAuthenticated;
                        }

                        // コールバック完了チェック
                        if let Some(token) = google_auth::load_token() {
                            if google_auth::is_token_valid(&token) {
                                self.auth_state = AuthState::Authenticated(token);
                            }
                        }
                    }
                    AuthState::Authenticated(_) => {
                        self.show_main_section(ui, ctx);
                    }
                    AuthState::Error(e) => {
                        ui.label(
                            egui::RichText::new(format!("認証エラー: {e}"))
                                .size(13.0)
                                .color(ERROR),
                        );
                        if ui
                            .add(
                                egui::Button::new(
                                    egui::RichText::new("再ログイン")
                                        .size(13.0)
                                        .color(ACCENT),
                                )
                                .fill(SURFACE)
                                .stroke(egui::Stroke::new(1.0, BORDER))
                                .corner_radius(egui::CornerRadius::same(8)),
                            )
                            .clicked()
                        {
                            self.auth_state = AuthState::NotAuthenticated;
                        }
                    }
                }
            });

        go_back
    }

    fn show_login_section(&mut self, ui: &mut egui::Ui) {
        if let Some(creds) = &self.credentials {
            if !creds.is_configured() {
                ui.label(
                    egui::RichText::new("Google認証の設定が必要です")
                        .size(13.0)
                        .color(TEXT_SECONDARY),
                );
                return;
            }
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
                        Ok(_) | Err(_) => {
                            ctx.request_repaint();
                        }
                    }
                });
            }
        }
    }

    fn show_main_section(&mut self, ui: &mut egui::Ui, _ctx: &egui::Context) {
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
            ui.label(
                egui::RichText::new("Google認証済み")
                    .size(12.0)
                    .color(SUCCESS),
            );
            if ui
                .add(
                    egui::Button::new(
                        egui::RichText::new("ログアウト")
                            .size(11.0)
                            .color(ERROR),
                    )
                    .fill(egui::Color32::TRANSPARENT)
                    .stroke(egui::Stroke::new(1.0, ERROR))
                    .corner_radius(egui::CornerRadius::same(6)),
                )
                .clicked()
            {
                google_auth::clear_token();
                self.auth_state = AuthState::NotAuthenticated;
                return;
            }
        });

        ui.add_space(16.0);

        // 1. 生産者手取り額スプレッドシート
        ui.label(
            egui::RichText::new("生産者手取り額（スプレ��ドシートURL）")
                .size(13.0)
                .strong()
                .color(TEXT_PRIMARY),
        );
        ui.add_space(4.0);
        let resp = ui.add_sized(
            [ui.available_width(), 32.0],
            egui::TextEdit::singleline(&mut self.original_url)
                .hint_text("スプレッドシートのURLを貼り付け")
                .font(egui::TextStyle::Body),
        );
        if resp.changed() {
            self.save();
        }

        ui.add_space(12.0);

        // 2. 支払先一覧スプレッドシート
        ui.label(
            egui::RichText::new("支払先一覧（スプレッドシートURL）")
                .size(13.0)
                .strong()
                .color(TEXT_PRIMARY),
        );
        ui.add_space(4.0);
        let resp = ui.add_sized(
            [ui.available_width(), 32.0],
            egui::TextEdit::singleline(&mut self.payee_url)
                .hint_text("スプレッドシートのURLを貼り付け")
                .font(egui::TextStyle::Body),
        );
        if resp.changed() {
            self.save();
        }

        ui.add_space(12.0);

        // 3. 決済予定日
        ui.label(
            egui::RichText::new("決済予定日")
                .size(13.0)
                .strong()
                .color(TEXT_PRIMARY),
        );
        ui.add_space(4.0);
        let resp = ui.add_sized(
            [200.0, 32.0],
            egui::TextEdit::singleline(&mut self.payment_day)
                .hint_text("例: 2026/04/30")
                .font(egui::TextStyle::Body),
        );
        if resp.changed() {
            self.save();
        }

        ui.add_space(12.0);

        // 4. 出力先フォルダ
        ui.label(
            egui::RichText::new("出力先フォルダ")
                .size(13.0)
                .strong()
                .color(TEXT_PRIMARY),
        );
        ui.add_space(4.0);
        let output_display = self
            .output_dir
            .as_ref()
            .map(|p| p.to_string_lossy().to_string())
            .unwrap_or_default();
        if file_select_row(ui, &output_display) {
            if let Some(path) = rfd::FileDialog::new().pick_folder() {
                self.output_dir = Some(path);
                self.save();
            }
        }

        ui.add_space(20.0);

        // 変換ボタン
        let can_run = !self.original_url.is_empty()
            && !self.payee_url.is_empty()
            && !self.payment_day.is_empty()
            && self.output_dir.is_some();
        let btn = ui.add_sized(
            [ui.available_width(), 40.0],
            egui::Button::new(
                egui::RichText::new("変換実行")
                    .size(14.0)
                    .strong()
                    .color(if can_run {
                        egui::Color32::WHITE
                    } else {
                        TEXT_SECONDARY
                    }),
            )
            .fill(if can_run { ACCENT } else { BORDER })
            .corner_radius(egui::CornerRadius::same(10)),
        );

        if btn.clicked() && can_run {
            self.run_conversion();
        }

        ui.add_space(12.0);

        // ログ表示
        show_log(ui, &self.log);

        // 完了時：出力フォルダを開くボタン
        if self.is_done {
            ui.add_space(8.0);
            if ui
                .add(
                    egui::Button::new(
                        egui::RichText::new("出力フォルダを開く")
                            .size(13.0)
                            .color(ACCENT),
                    )
                    .fill(SURFACE)
                    .stroke(egui::Stroke::new(1.0, BORDER))
                    .corner_radius(egui::CornerRadius::same(8)),
                )
                .clicked()
            {
                if let Some(dir) = &self.output_dir {
                    let _ = open::that(dir);
                }
            }
        }
    }

    fn get_token(&self) -> Option<&str> {
        if let AuthState::Authenticated(ref token) = self.auth_state {
            Some(&token.access_token)
        } else {
            None
        }
    }

    fn run_conversion(&mut self) {
        self.log.clear();
        self.is_done = false;

        let token = match self.get_token() {
            Some(t) => t.to_string(),
            None => {
                self.log.push(LogEntry {
                    text: "認証が必要です".into(),
                    kind: LogKind::Error,
                });
                return;
            }
        };

        let original_id = extract_spreadsheet_id(&self.original_url);
        let payee_id = extract_spreadsheet_id(&self.payee_url);
        let output_dir = self.output_dir.as_ref().unwrap();

        self.log.push(LogEntry {
            text: "スプレッドシートを読み込み中...".into(),
            kind: LogKind::Info,
        });

        match run_convert(&token, &original_id, &payee_id, &self.payment_day, output_dir) {
            Ok(result) => {
                for (filename, count) in &result.file_rows {
                    self.log.push(LogEntry {
                        text: format!("{} （{}件）", filename, count),
                        kind: LogKind::Ok,
                    });
                }

                if !result.unregistered.is_empty() {
                    self.log.push(LogEntry {
                        text: format!(
                            "未登録ユーザー: {}件 → non_registered_user_id.csv",
                            result.unregistered.len()
                        ),
                        kind: LogKind::Info,
                    });
                    for (uid, name) in &result.unregistered {
                        self.log.push(LogEntry {
                            text: format!("  ID:{} {}", uid, name),
                            kind: LogKind::Info,
                        });
                    }
                }

                self.log.push(LogEntry {
                    text: format!(
                        "完了: {}ファイルを出力（取引日: {}）",
                        result.file_count,
                        last_day_of_prev_month()
                    ),
                    kind: LogKind::Done,
                });
                self.is_done = true;
            }
            Err(e) => {
                self.log.push(LogEntry {
                    text: format!("エラー: {e}"),
                    kind: LogKind::Error,
                });
            }
        }
    }

    fn save(&self) {
        let settings = MfSettings {
            original_url: Some(self.original_url.clone()),
            payee_url: Some(self.payee_url.clone()),
            output_dir: self
                .output_dir
                .as_ref()
                .map(|p| p.to_string_lossy().to_string()),
            payment_day: Some(self.payment_day.clone()),
        };
        save_settings(&settings);
    }
}
