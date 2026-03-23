use eframe::egui;
use std::collections::BTreeMap;
use std::path::PathBuf;

use crate::convert;
use crate::style::*;

/// イオン近畿生鮮MDアップロード用レコード
struct MdRecord {
    store_code: String,
    delivery_date: String, // YYYY/MM/DD
    eos_code: String,
    product_name: String,
    quantity: u32,
    lot: String,
    purchase_price: String,
    sell_price: String,
}

impl MdRecord {
    /// CSVフォーマットの1行（33フィールド）を出力
    fn to_csv_line(&self) -> String {
        // フィールド: 受注明細ID, 届け先コード, 最終納品先コード, 最終納品先納品日,
        //   便NO, 商品コード(発注用), 商品名, 商品名カナ, 処理種別,
        //   受注数量, 発注単位, 単価登録単位, 商品重量, 原単価, 売単価,
        //   都道府県コード, 原産エリアコード, 等級コード, 銘柄コード, バイオ区分,
        //   カラーコード, サイズコード, 栽培コード, 生産者コード, 団体コード,
        //   用途コード, 解凍区分, 商品状態区分, 養殖区分, 形状・部位コード,
        //   品種コード, 水域コード, 商品PR
        format!(
            ",,{},{},01,{},{},,,{},{},,,{},{},,,,,,,,,,,,,,,,,,",
            self.store_code,
            self.delivery_date,
            self.eos_code,
            self.product_name,
            self.quantity,
            self.lot,
            self.purchase_price,
            self.sell_price,
        )
    }
}

/// Excelから読み取ってレコードを生成
fn read_aeon_excel(path: &std::path::Path) -> Result<Vec<MdRecord>, String> {
    use calamine::{open_workbook, Reader, Xlsx, Data};

    let mut workbook: Xlsx<_> = open_workbook(path)
        .map_err(|e| format!("Excelファイルを開けません: {e}"))?;

    let sheet_names: Vec<String> = workbook.sheet_names().to_vec();
    let mut records = Vec::new();

    for sheet_name in &sheet_names {
        let range = workbook
            .worksheet_range(sheet_name)
            .map_err(|e| format!("シート '{}' の読み取りに失敗: {e}", sheet_name))?;

        // A1: 店舗コード
        let store_code = match range.get_value((0, 0)) {
            Some(Data::String(s)) => s.trim().to_string(),
            Some(Data::Float(f)) => format!("{:05}", *f as u64),
            Some(Data::Int(i)) => format!("{:05}", i),
            _ => continue, // 店舗コードがなければスキップ
        };

        if store_code.is_empty() || store_code.chars().any(|c| !c.is_ascii_digit()) {
            continue; // 数値コードでなければデータシートではない
        }

        // P6-V6 (row=5, col=15..=21): 納品日
        let mut dates: Vec<(usize, String)> = Vec::new(); // (col_index, formatted_date)
        for col in 15..=21 {
            if let Some(val) = range.get_value((5, col as u32)) {
                let date_str = match val {
                    Data::DateTime(dt) => {
                        // calamine DateTime → serial number → 日付変換
                        let serial = dt.as_f64();
                        serial_to_date_string(serial)
                    }
                    Data::Float(f) => serial_to_date_string(*f),
                    Data::String(s) => s.trim().to_string(),
                    _ => continue,
                };
                if !date_str.is_empty() {
                    dates.push((col, date_str));
                }
            }
        }

        if dates.is_empty() {
            continue;
        }

        // Row 8+ (row_idx=7+): データ行
        let height = range.height();
        for row_idx in 7..height {
            // C列 (col=2): EOS（商品コード）
            let eos_code = match range.get_value((row_idx as u32, 2)) {
                Some(Data::String(s)) => s.trim().to_string(),
                Some(Data::Float(f)) => format!("{}", *f as u64),
                Some(Data::Int(i)) => format!("{}", i),
                _ => continue,
            };
            if eos_code.is_empty() {
                continue;
            }

            // B列 (col=1): 商品名
            let product_name = match range.get_value((row_idx as u32, 1)) {
                Some(Data::String(s)) => s.trim().replace('\n', "").to_string(),
                Some(v) => format!("{}", v),
                None => continue,
            };

            // I列 (col=8): ロット（発注単位）
            let lot = match range.get_value((row_idx as u32, 8)) {
                Some(Data::Float(f)) => format!("{}", *f as u64),
                Some(Data::Int(i)) => format!("{}", i),
                Some(Data::String(s)) => s.trim().to_string(),
                _ => String::new(),
            };

            // L列 (col=11): 原単価（店頭着単価）
            let purchase_price = match range.get_value((row_idx as u32, 11)) {
                Some(Data::Float(f)) => format!("{}", *f as u64),
                Some(Data::Int(i)) => format!("{}", i),
                Some(Data::String(s)) => s.trim().to_string(),
                _ => String::new(),
            };

            // M列 (col=12): 売単価（店頭売価）
            let sell_price = match range.get_value((row_idx as u32, 12)) {
                Some(Data::Float(f)) => format!("{}", *f as u64),
                Some(Data::Int(i)) => format!("{}", i),
                Some(Data::String(s)) => s.trim().to_string(),
                _ => String::new(),
            };

            // P-V列 (col=15..=21): 各日の数量
            for &(col, ref date_str) in &dates {
                let qty = match range.get_value((row_idx as u32, col as u32)) {
                    Some(Data::Float(f)) => *f as u32,
                    Some(Data::Int(i)) => *i as u32,
                    Some(Data::String(s)) => s.trim().parse::<u32>().unwrap_or(0),
                    _ => 0,
                };
                if qty == 0 {
                    continue;
                }

                records.push(MdRecord {
                    store_code: store_code.clone(),
                    delivery_date: date_str.clone(),
                    eos_code: eos_code.clone(),
                    product_name: product_name.clone(),
                    quantity: qty,
                    lot: lot.clone(),
                    purchase_price: purchase_price.clone(),
                    sell_price: sell_price.clone(),
                });
            }
        }
    }

    Ok(records)
}

/// Excelシリアル値 → "YYYY/MM/DD" 文字列
fn serial_to_date_string(serial: f64) -> String {
    let days = serial as i64;
    if days < 1 {
        return String::new();
    }
    // Excel serial → Unix days: Excel epoch 1899-12-30, serial 25569 = 1970-01-01
    let unix_days = days - 25569;
    let (y, m, d) = convert::unix_days_to_ymd(unix_days);
    format!("{}/{:02}/{:02}", y, m, d)
}

/// 日付文字列 → ファイル名用 "YYYYMMDD"
fn date_to_filename(date_str: &str) -> String {
    // "YYYY/MM/DD" → "YYYYMMDD"
    let parts: Vec<&str> = date_str.split('/').collect();
    if parts.len() == 3 {
        format!(
            "{}{:02}{:02}",
            parts[0],
            parts[1].parse::<u32>().unwrap_or(0),
            parts[2].parse::<u32>().unwrap_or(0),
        )
    } else {
        date_str.replace('/', "")
    }
}

// --- 設定の永続化 ---

#[derive(serde::Serialize, serde::Deserialize)]
struct AeonSettings {
    input_file: Option<String>,
    output_dir: Option<String>,
}

impl Default for AeonSettings {
    fn default() -> Self {
        Self {
            input_file: None,
            output_dir: None,
        }
    }
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
    base.join("cart-converter").join("aeon-settings.json")
}

fn load_settings() -> AeonSettings {
    if let Ok(s) = std::fs::read_to_string(settings_path()) {
        serde_json::from_str(&s).unwrap_or_default()
    } else {
        AeonSettings::default()
    }
}

fn save_settings(settings: &AeonSettings) {
    if let Some(parent) = settings_path().parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(json) = serde_json::to_string_pretty(settings) {
        let _ = std::fs::write(settings_path(), json);
    }
}

// --- UI ---

pub struct AeonPage {
    input_file: Option<PathBuf>,
    output_dir: Option<PathBuf>,
    log: Vec<LogEntry>,
    is_done: bool,
}

impl Default for AeonPage {
    fn default() -> Self {
        let settings = load_settings();
        Self {
            input_file: settings.input_file.map(PathBuf::from),
            output_dir: settings.output_dir.map(PathBuf::from),
            log: Vec::new(),
            is_done: false,
        }
    }
}

impl AeonPage {
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
                    egui::RichText::new("イオン近畿 生鮮MDアップロード")
                        .size(18.0)
                        .strong()
                        .color(TEXT_PRIMARY),
                );
                ui.add_space(4.0);
                ui.label(
                    egui::RichText::new(
                        "商品リスト（Excel）から納品日別のアップロード用テキストファイルを生成します",
                    )
                    .size(12.0)
                    .color(TEXT_SECONDARY),
                );
                ui.add_space(20.0);

                // 入力ファイル選択
                ui.label(
                    egui::RichText::new("商品リスト（Excel）")
                        .size(13.0)
                        .strong()
                        .color(TEXT_PRIMARY),
                );
                ui.add_space(4.0);
                let input_display = self
                    .input_file
                    .as_ref()
                    .and_then(|p| p.file_name())
                    .map(|n| n.to_string_lossy().to_string())
                    .unwrap_or_default();
                if file_select_row(ui, &input_display) {
                    if let Some(path) = rfd::FileDialog::new()
                        .add_filter("Excel", &["xlsx", "xls"])
                        .pick_file()
                    {
                        self.input_file = Some(path);
                        self.save();
                    }
                }

                ui.add_space(16.0);

                // 出力フォルダ選択
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

                ui.add_space(24.0);

                // 変換ボタン
                let can_run = self.input_file.is_some() && self.output_dir.is_some();
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

                ui.add_space(16.0);

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
            });

        go_back
    }

    fn run_conversion(&mut self) {
        self.log.clear();
        self.is_done = false;

        let input_path = self.input_file.as_ref().unwrap();
        let output_dir = self.output_dir.as_ref().unwrap();

        self.log.push(LogEntry {
            text: format!("読み込み中: {}", input_path.display()),
            kind: LogKind::Info,
        });

        match read_aeon_excel(input_path) {
            Ok(records) => {
                if records.is_empty() {
                    self.log.push(LogEntry {
                        text: "数量が入った行が見つかりませんでした".to_string(),
                        kind: LogKind::Error,
                    });
                    return;
                }

                self.log.push(LogEntry {
                    text: format!("{}件のレコードを検出", records.len()),
                    kind: LogKind::Ok,
                });

                // 日付ごとにグループ化
                let mut by_date: BTreeMap<String, Vec<&MdRecord>> = BTreeMap::new();
                for rec in &records {
                    let key = date_to_filename(&rec.delivery_date);
                    by_date.entry(key).or_default().push(rec);
                }

                // ファイル出力
                let mut file_count = 0;
                for (date_key, recs) in &by_date {
                    let filename = format!("{}.txt", date_key);
                    let filepath = output_dir.join(&filename);

                    let content: String = recs
                        .iter()
                        .map(|r| r.to_csv_line())
                        .collect::<Vec<_>>()
                        .join("\r\n");

                    // Shift_JIS (CP932) で出力
                    match encode_to_cp932(&content) {
                        Ok(bytes) => {
                            match std::fs::write(&filepath, bytes) {
                                Ok(_) => {
                                    self.log.push(LogEntry {
                                        text: format!(
                                            "{} （{}件）",
                                            filename,
                                            recs.len()
                                        ),
                                        kind: LogKind::Ok,
                                    });
                                    file_count += 1;
                                }
                                Err(e) => {
                                    self.log.push(LogEntry {
                                        text: format!("{} の書き込みに失敗: {e}", filename),
                                        kind: LogKind::Error,
                                    });
                                }
                            }
                        }
                        Err(e) => {
                            // Shift_JISに変換できない文字がある場合はUTF-8で出力
                            self.log.push(LogEntry {
                                text: format!("{}: Shift_JIS変換不可、UTF-8で出力 ({})", filename, e),
                                kind: LogKind::Info,
                            });
                            match std::fs::write(&filepath, content.as_bytes()) {
                                Ok(_) => {
                                    self.log.push(LogEntry {
                                        text: format!(
                                            "{} （{}件）",
                                            filename,
                                            recs.len()
                                        ),
                                        kind: LogKind::Ok,
                                    });
                                    file_count += 1;
                                }
                                Err(e) => {
                                    self.log.push(LogEntry {
                                        text: format!("{} の書き込みに失敗: {e}", filename),
                                        kind: LogKind::Error,
                                    });
                                }
                            }
                        }
                    }
                }

                self.log.push(LogEntry {
                    text: format!("完了: {}ファイルを出力しました", file_count),
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
        let settings = AeonSettings {
            input_file: self.input_file.as_ref().map(|p| p.to_string_lossy().to_string()),
            output_dir: self.output_dir.as_ref().map(|p| p.to_string_lossy().to_string()),
        };
        save_settings(&settings);
    }
}

/// UTF-8文字列をCP932 (Shift_JIS) バイト列に変換
fn encode_to_cp932(text: &str) -> Result<Vec<u8>, String> {
    use encoding_rs::SHIFT_JIS;
    let (bytes, _, had_errors) = SHIFT_JIS.encode(text);
    if had_errors {
        Err("一部の文字がShift_JISに変換できません".to_string())
    } else {
        Ok(bytes.into_owned())
    }
}
