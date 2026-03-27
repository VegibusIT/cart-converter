use eframe::egui;
use std::path::PathBuf;

use crate::style::*;

// --- ITFバーコード画像生成 ---

/// ITF (Interleaved 2 of 5) バーコードをPNG画像としてメモリ上に生成する
fn generate_itf_png(data: &str) -> Result<Vec<u8>, String> {
    // 偶数桁チェック
    if data.len() % 2 != 0 {
        return Err("ITFバーコードは偶数桁が必要です".to_string());
    }

    let digits: Vec<usize> = data
        .chars()
        .map(|c| {
            c.to_digit(10)
                .ok_or_else(|| format!("不正な文字: {c}"))
                .map(|d| d as usize)
        })
        .collect::<Result<Vec<_>, _>>()?;

    // ITF 2-of-5 パターン (false=narrow, true=wide)
    const PATTERNS: [[bool; 5]; 10] = [
        [false, false, true, true, false],  // 0: NNWWN
        [true, false, false, false, true],   // 1: WNNNE
        [false, true, false, false, true],   // 2: NWNNW
        [true, true, false, false, false],   // 3: WWNNN
        [false, false, true, false, true],   // 4: NNWNW
        [true, false, true, false, false],   // 5: WNWNN
        [false, true, true, false, false],   // 6: NWWNN
        [false, false, false, true, true],   // 7: NNNWW
        [true, false, false, true, false],   // 8: WNNWN
        [false, true, false, true, false],   // 9: NWNWN
    ];

    let narrow = 2u32;
    let wide = 5u32;
    let bar_height = 80u32;
    let quiet_zone = 20u32; // 左右の余白

    // バー列を構築: (is_bar, width)
    let mut bars: Vec<(bool, u32)> = Vec::new();

    // スタートパターン: narrow bar, narrow space, narrow bar, narrow space
    bars.push((true, narrow));
    bars.push((false, narrow));
    bars.push((true, narrow));
    bars.push((false, narrow));

    // データペア
    for pair in digits.chunks(2) {
        let d1 = pair[0];
        let d2 = pair[1];
        for i in 0..5 {
            bars.push((true, if PATTERNS[d1][i] { wide } else { narrow }));
            bars.push((false, if PATTERNS[d2][i] { wide } else { narrow }));
        }
    }

    // エンドパターン: wide bar, narrow space, narrow bar
    bars.push((true, wide));
    bars.push((false, narrow));
    bars.push((true, narrow));

    let barcode_width: u32 = bars.iter().map(|(_, w)| *w).sum();
    let total_width = barcode_width + quiet_zone * 2;
    let total_height = bar_height;

    // 画像生成（白背景）
    let mut img = image::GrayImage::from_pixel(total_width, total_height, image::Luma([255u8]));

    // バーを描画
    let mut x = quiet_zone;
    for (is_bar, width) in &bars {
        if *is_bar {
            for dx in 0..*width {
                for y in 0..bar_height {
                    img.put_pixel(x + dx, y, image::Luma([0u8]));
                }
            }
        }
        x += width;
    }

    // PNGエンコード
    let mut buf = std::io::Cursor::new(Vec::new());
    image::DynamicImage::ImageLuma8(img)
        .write_to(&mut buf, image::ImageFormat::Png)
        .map_err(|e| format!("PNG生成エラー: {e}"))?;

    Ok(buf.into_inner())
}

// --- SCMバーコード固定値 ---
const SCM_PREFIX: &str = "5";          // 位置1: 物流識別コード
const SCM_VENDOR: &str = "8934";       // 位置2-5: 納品業者番号
const SCM_SUPPLIER: &str = "307270";   // 位置10-15: 取引先コード（やさいバス）
const SCM_RESERVE: &str = "88";        // 位置16-17: 予備
const SCM_SUFFIX: &str = "00810";      // 位置22-26: ルーティングコード

/// 26桁SCMバーコードを生成
fn build_scm_barcode(store_number: u32, seq: u32) -> String {
    format!(
        "{}{}{:04}{}{}{:04}{}",
        SCM_PREFIX,
        SCM_VENDOR,
        store_number,
        SCM_SUPPLIER,
        SCM_RESERVE,
        seq,
        SCM_SUFFIX,
    )
}

// --- 店舗マスター ---
fn store_name(store_number: u32) -> &'static str {
    match store_number {
        5 => "龍ヶ丘",
        27 => "湖北",
        336 => "東茂原",
        _ => "",
    }
}

// --- 設定の永続化 ---

#[derive(serde::Serialize, serde::Deserialize)]
struct KasumiSettings {
    output_dir: Option<String>,
    store_number: u32,
    next_seq: u32,
    label_count: u32,
    delivery_date: String,
}

impl Default for KasumiSettings {
    fn default() -> Self {
        Self {
            output_dir: None,
            store_number: 336,
            next_seq: 1,
            label_count: 1,
            delivery_date: String::new(),
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
    base.join("cart-converter").join("kasumi-settings.json")
}

fn load_settings() -> KasumiSettings {
    if let Ok(s) = std::fs::read_to_string(settings_path()) {
        serde_json::from_str(&s).unwrap_or_default()
    } else {
        KasumiSettings::default()
    }
}

fn save_settings(settings: &KasumiSettings) {
    if let Some(parent) = settings_path().parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(json) = serde_json::to_string_pretty(settings) {
        let _ = std::fs::write(settings_path(), json);
    }
}

// --- Excel出力 ---

fn generate_scm_excel(
    output_dir: &std::path::Path,
    store_number: u32,
    start_seq: u32,
    count: u32,
    delivery_date: &str,
) -> Result<(PathBuf, u32), String> {
    use rust_xlsxwriter::*;

    let name = store_name(store_number);
    let store_display = if name.is_empty() {
        format!("{:04}", store_number)
    } else {
        format!("{:04}_{}", store_number, name)
    };
    let filename = format!("SCMラベル_{}.xlsx", store_display);
    let filepath = output_dir.join(&filename);

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("SCMラベル").map_err(|e| format!("シート名設定エラー: {e}"))?;

    // --- ページ設定（30×50mmラベル用） ---
    // 余白を全て0に
    worksheet.set_margins(0.0, 0.0, 0.0, 0.0, 0.0, 0.0);
    worksheet.set_header("");
    worksheet.set_footer("");
    // 横向き（50mm幅 × 30mm高さ）
    worksheet.set_landscape();

    // 列幅設定: A列のみ使用（50mm ≈ 25文字幅）
    worksheet.set_column_width(0, 25).map_err(|e| format!("{e}"))?;

    // フォーマット定義（コンパクトサイズ）
    let fmt_title = Format::new()
        .set_font_name("游ゴシック")
        .set_font_size(9)
        .set_bold();

    let fmt_info = Format::new()
        .set_font_name("游ゴシック")
        .set_font_size(7);

    let fmt_barcode_text = Format::new()
        .set_font_name("游ゴシック")
        .set_font_size(7)
        .set_align(FormatAlign::Center);

    // 行の高さ（30mm ≈ 85pt を4行で配分）
    let row_height_title: f64 = 16.0;   // 行1: タイトル＋納品先
    let row_height_info: f64 = 12.0;    // 行2: 店番＋納品日
    let row_height_barcode: f64 = 45.0; // 行3: バーコード画像
    let row_height_text: f64 = 12.0;    // 行4: バーコード番号

    let end_seq = start_seq + count;
    let mut row: u32 = 0;
    let mut page_breaks: Vec<u32> = Vec::new();

    for seq in start_seq..end_seq {
        let barcode = build_scm_barcode(store_number, seq);
        let store_name_str = store_name(store_number);
        let store_label = if store_name_str.is_empty() {
            format!("店番:{:04}", store_number)
        } else {
            format!("店番:{:04} {}", store_number, store_name_str)
        };

        // 行1: やさいバス　カスミ佐倉流通センター 冷蔵 野菜
        worksheet.set_row_height(row, row_height_title).map_err(|e| format!("{e}"))?;
        worksheet.write_with_format(row, 0, "やさいバス カスミ佐倉流通センター 冷蔵 野菜", &fmt_title)
            .map_err(|e| format!("{e}"))?;
        row += 1;

        // 行2: 店番・店名（＋納品日）
        let line2 = if delivery_date.is_empty() {
            store_label
        } else {
            format!("{} 納品日:{}", store_label, delivery_date)
        };
        worksheet.set_row_height(row, row_height_info).map_err(|e| format!("{e}"))?;
        worksheet.write_with_format(row, 0, &line2, &fmt_info)
            .map_err(|e| format!("{e}"))?;
        row += 1;

        // 行3: ITFバーコード画像
        let png_data = generate_itf_png(&barcode)
            .map_err(|e| format!("バーコード生成エラー: {e}"))?;
        let barcode_image = Image::new_from_buffer(&png_data)
            .map_err(|e| format!("画像読込エラー: {e}"))?
            .set_scale_to_size(180.0, 40.0, false);
        worksheet.set_row_height(row, row_height_barcode).map_err(|e| format!("{e}"))?;
        worksheet.insert_image(row, 0, &barcode_image)
            .map_err(|e| format!("画像挿入エラー: {e}"))?;
        row += 1;

        // 行4: バーコード番号（人間可読テキスト）
        worksheet.set_row_height(row, row_height_text).map_err(|e| format!("{e}"))?;
        worksheet.write_with_format(row, 0, &barcode, &fmt_barcode_text)
            .map_err(|e| format!("{e}"))?;
        row += 1;

        // ラベル間にページ区切りを追加（最後のラベル以外）
        if seq < end_seq - 1 {
            page_breaks.push(row);
        }
    }

    // ページ区切り設定（1ラベル=1ページ）
    worksheet.set_page_breaks(&page_breaks)
        .map_err(|e| format!("ページ区切り設定エラー: {e}"))?;

    // 印刷範囲設定
    worksheet.set_print_area(0, 0, row - 1, 0)
        .map_err(|e| format!("印刷範囲設定エラー: {e}"))?;

    workbook.save(&filepath).map_err(|e| format!("Excel保存エラー: {e}"))?;

    Ok((filepath, end_seq))
}

// --- UI ---

pub struct KasumiPage {
    output_dir: Option<PathBuf>,
    store_number_input: String,
    store_number: u32,
    next_seq: u32,
    label_count_input: String,
    label_count: u32,
    delivery_date: String,
    log: Vec<LogEntry>,
    is_done: bool,
}

impl Default for KasumiPage {
    fn default() -> Self {
        let settings = load_settings();
        Self {
            output_dir: settings.output_dir.map(PathBuf::from),
            store_number_input: format!("{}", settings.store_number),
            store_number: settings.store_number,
            next_seq: settings.next_seq,
            label_count_input: format!("{}", settings.label_count),
            label_count: settings.label_count,
            delivery_date: settings.delivery_date,
            log: Vec::new(),
            is_done: false,
        }
    }
}

impl KasumiPage {
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
                    egui::RichText::new("カスミ SCMラベル発行")
                        .size(18.0)
                        .strong()
                        .color(TEXT_PRIMARY),
                );
                ui.add_space(4.0);
                ui.label(
                    egui::RichText::new(
                        "カスミ佐倉流通センター納品用の26桁SCMバーコードラベルを生成します",
                    )
                    .size(12.0)
                    .color(TEXT_SECONDARY),
                );
                ui.add_space(20.0);

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

                ui.add_space(16.0);

                // 店舗番号
                ui.horizontal(|ui| {
                    ui.label(
                        egui::RichText::new("店舗番号")
                            .size(13.0)
                            .strong()
                            .color(TEXT_PRIMARY),
                    );
                    ui.add_space(8.0);
                    let response = ui.add_sized(
                        [80.0, 28.0],
                        egui::TextEdit::singleline(&mut self.store_number_input)
                            .font(egui::TextStyle::Body),
                    );
                    if response.changed() {
                        if let Ok(n) = self.store_number_input.trim().parse::<u32>() {
                            self.store_number = n;
                            self.save();
                        }
                    }
                    // 店名表示
                    let name = store_name(self.store_number);
                    if !name.is_empty() {
                        ui.add_space(8.0);
                        ui.label(
                            egui::RichText::new(name)
                                .size(13.0)
                                .color(ACCENT),
                        );
                    }
                });

                ui.add_space(12.0);

                // 発行枚数
                ui.horizontal(|ui| {
                    ui.label(
                        egui::RichText::new("発行枚数")
                            .size(13.0)
                            .strong()
                            .color(TEXT_PRIMARY),
                    );
                    ui.add_space(8.0);
                    let response = ui.add_sized(
                        [80.0, 28.0],
                        egui::TextEdit::singleline(&mut self.label_count_input)
                            .font(egui::TextStyle::Body),
                    );
                    if response.changed() {
                        if let Ok(n) = self.label_count_input.trim().parse::<u32>() {
                            if n > 0 {
                                self.label_count = n;
                                self.save();
                            }
                        }
                    }
                    ui.label(
                        egui::RichText::new("枚")
                            .size(13.0)
                            .color(TEXT_PRIMARY),
                    );
                });

                ui.add_space(12.0);

                // 納品日（任意）
                ui.horizontal(|ui| {
                    ui.label(
                        egui::RichText::new("納品日")
                            .size(13.0)
                            .strong()
                            .color(TEXT_PRIMARY),
                    );
                    ui.add_space(20.0);
                    let response = ui.add_sized(
                        [140.0, 28.0],
                        egui::TextEdit::singleline(&mut self.delivery_date)
                            .hint_text("例: 3/28")
                            .font(egui::TextStyle::Body),
                    );
                    if response.changed() {
                        self.save();
                    }
                    ui.label(
                        egui::RichText::new("（任意）")
                            .size(11.0)
                            .color(TEXT_SECONDARY),
                    );
                });

                ui.add_space(12.0);

                // 連番情報
                ui.horizontal(|ui| {
                    ui.label(
                        egui::RichText::new(format!("次の連番: {}", self.next_seq))
                            .size(12.0)
                            .color(TEXT_SECONDARY),
                    );
                    ui.add_space(8.0);
                    if ui
                        .add(
                            egui::Button::new(
                                egui::RichText::new("リセット")
                                    .size(11.0)
                                    .color(ERROR),
                            )
                            .fill(SURFACE)
                            .stroke(egui::Stroke::new(1.0, BORDER))
                            .corner_radius(egui::CornerRadius::same(4)),
                        )
                        .clicked()
                    {
                        self.next_seq = 1;
                        self.save();
                    }
                });

                // バーコードプレビュー
                ui.add_space(8.0);
                let preview = build_scm_barcode(self.store_number, self.next_seq);
                ui.label(
                    egui::RichText::new(format!("プレビュー: {}", preview))
                        .size(11.0)
                        .color(TEXT_SECONDARY),
                );

                ui.add_space(20.0);

                // 発行ボタン
                let can_run = self.output_dir.is_some() && self.label_count > 0;
                let btn = ui.add_sized(
                    [ui.available_width(), 40.0],
                    egui::Button::new(
                        egui::RichText::new("ラベル発行")
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
                    self.run_generation();
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

    fn run_generation(&mut self) {
        self.log.clear();
        self.is_done = false;

        let output_dir = self.output_dir.as_ref().unwrap();

        self.log.push(LogEntry {
            text: format!(
                "店舗: {:04} {}  枚数: {}  連番: {}〜{}",
                self.store_number,
                store_name(self.store_number),
                self.label_count,
                self.next_seq,
                self.next_seq + self.label_count - 1,
            ),
            kind: LogKind::Info,
        });

        match generate_scm_excel(
            output_dir,
            self.store_number,
            self.next_seq,
            self.label_count,
            &self.delivery_date,
        ) {
            Ok((filepath, new_next_seq)) => {
                self.log.push(LogEntry {
                    text: format!(
                        "出力: {}",
                        filepath.file_name().unwrap_or_default().to_string_lossy()
                    ),
                    kind: LogKind::Ok,
                });

                let preview_start = build_scm_barcode(self.store_number, self.next_seq);
                let preview_end = build_scm_barcode(self.store_number, new_next_seq - 1);
                self.log.push(LogEntry {
                    text: format!("バーコード: {} 〜 {}", preview_start, preview_end),
                    kind: LogKind::Info,
                });

                // 連番を更新して保存
                self.next_seq = new_next_seq;
                self.save();

                self.log.push(LogEntry {
                    text: format!(
                        "完了: {}枚のラベルを生成しました（次の連番: {}）",
                        self.label_count, self.next_seq
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
        let settings = KasumiSettings {
            output_dir: self.output_dir.as_ref().map(|p| p.to_string_lossy().to_string()),
            store_number: self.store_number,
            next_seq: self.next_seq,
            label_count: self.label_count,
            delivery_date: self.delivery_date.clone(),
        };
        save_settings(&settings);
    }
}
