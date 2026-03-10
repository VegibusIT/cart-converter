// Prevents additional console window on Windows in release
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod convert;

use eframe::egui;
use serde::{Deserialize, Serialize};
use std::path::PathBuf;

use convert::ColumnMapping;

// やさいバス palette (vegibus.com準拠)
const HEADER_BG: egui::Color32 = egui::Color32::from_rgb(59, 100, 30);
const BG: egui::Color32 = egui::Color32::from_rgb(255, 255, 255);
const SURFACE: egui::Color32 = egui::Color32::from_rgb(250, 252, 248);
const TEXT_PRIMARY: egui::Color32 = egui::Color32::from_rgb(51, 51, 51);
const TEXT_SECONDARY: egui::Color32 = egui::Color32::from_rgb(130, 130, 130);
const ACCENT: egui::Color32 = egui::Color32::from_rgb(59, 100, 30);
const BORDER: egui::Color32 = egui::Color32::from_rgb(220, 220, 220);
const SUCCESS: egui::Color32 = egui::Color32::from_rgb(59, 100, 30);
const ERROR: egui::Color32 = egui::Color32::from_rgb(200, 50, 30);
const PROGRESS_BG: egui::Color32 = egui::Color32::from_rgb(230, 230, 230);

// --- Preset & Settings ---
#[derive(Serialize, Deserialize, Clone, Debug)]
struct Preset {
    name: String,
    mapping: ColumnMapping,
}

#[derive(Serialize, Deserialize)]
struct Settings {
    output_dir: Option<String>,
    #[serde(default)]
    presets: Vec<Preset>,
    #[serde(default)]
    selected_preset: usize,
    // 旧形式との互換用（読み込みのみ）
    #[serde(default, skip_serializing)]
    column_mapping: Option<ColumnMapping>,
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            output_dir: None,
            presets: vec![Preset {
                name: "デフォルト".to_string(),
                mapping: ColumnMapping::default(),
            }],
            selected_preset: 0,
            column_mapping: None,
        }
    }
}

/// ユーザー設定ディレクトリのパス（更新しても消えない）
fn config_dir_path() -> PathBuf {
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
    base.join("cart-converter")
}

fn settings_path() -> PathBuf {
    config_dir_path().join("settings.json")
}

/// 旧バージョンのexe隣接設定ファイルパス
fn legacy_settings_path() -> PathBuf {
    let mut path = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|p| p.to_path_buf()))
        .unwrap_or_else(|| PathBuf::from("."));
    path.push("cart-converter-settings.json");
    path
}

fn load_settings() -> Settings {
    // 1. 新しい場所から読む
    if let Ok(s) = std::fs::read_to_string(settings_path()) {
        if let Ok(settings) = serde_json::from_str::<Settings>(&s) {
            return migrate_old_format(settings);
        }
    }
    // 2. 旧場所から読んでマイグレーション
    if let Ok(s) = std::fs::read_to_string(legacy_settings_path()) {
        if let Ok(settings) = serde_json::from_str::<Settings>(&s) {
            let settings = migrate_old_format(settings);
            save_settings(&settings); // 新しい場所に保存
            return settings;
        }
    }
    Settings::default()
}

/// 旧形式（column_mapping単体）→新形式（presets配列）に変換
fn migrate_old_format(mut settings: Settings) -> Settings {
    if settings.presets.is_empty() {
        if let Some(mapping) = settings.column_mapping.take() {
            settings.presets.push(Preset {
                name: "デフォルト".to_string(),
                mapping,
            });
        } else {
            settings.presets.push(Preset {
                name: "デフォルト".to_string(),
                mapping: ColumnMapping::default(),
            });
        }
    }
    settings.column_mapping = None;
    if settings.selected_preset >= settings.presets.len() {
        settings.selected_preset = 0;
    }
    settings
}

fn save_settings(settings: &Settings) {
    let dir = config_dir_path();
    let _ = std::fs::create_dir_all(&dir);
    if let Ok(json) = serde_json::to_string_pretty(settings) {
        let _ = std::fs::write(settings_path(), json);
    }
}

/// デスクトップパスを取得（Windows / macOS / Linux 対応）
fn desktop_path() -> Option<PathBuf> {
    if let Some(home) = std::env::var_os("USERPROFILE")
        .or_else(|| std::env::var_os("HOME"))
    {
        let desktop = PathBuf::from(&home).join("Desktop");
        if desktop.exists() {
            return Some(desktop);
        }
        let desktop_jp = PathBuf::from(&home).join("デスクトップ");
        if desktop_jp.exists() {
            return Some(desktop_jp);
        }
        return Some(PathBuf::from(home));
    }
    None
}

fn today_folder_name() -> String {
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_secs() as i64;
    let days = now / 86400;
    let (y, m, d) = convert::unix_days_to_ymd(days);
    format!("{:04}{:02}{:02}", y, m, d)
}

// --- やさいバス style ---
fn setup_style(ctx: &egui::Context) {
    let mut style = (*ctx.style()).clone();

    // ライトモードをベースにする（ComboBoxポップアップ等のデフォルト色に影響）
    style.visuals = egui::Visuals::light();

    style.visuals.panel_fill = BG;
    style.visuals.window_fill = SURFACE;
    style.visuals.extreme_bg_color = SURFACE;

    style.visuals.widgets.inactive.bg_fill = SURFACE;
    style.visuals.widgets.inactive.bg_stroke = egui::Stroke::new(1.0, BORDER);
    style.visuals.widgets.inactive.fg_stroke = egui::Stroke::new(1.0, TEXT_PRIMARY);
    style.visuals.widgets.inactive.corner_radius = egui::CornerRadius::same(8);

    style.visuals.widgets.hovered.bg_fill = SURFACE;
    style.visuals.widgets.hovered.bg_stroke = egui::Stroke::new(1.5, ACCENT);
    style.visuals.widgets.hovered.fg_stroke = egui::Stroke::new(1.0, TEXT_PRIMARY);
    style.visuals.widgets.hovered.corner_radius = egui::CornerRadius::same(8);

    style.visuals.widgets.active.bg_fill = egui::Color32::from_rgb(245, 245, 247);
    style.visuals.widgets.active.bg_stroke = egui::Stroke::new(1.5, ACCENT);
    style.visuals.widgets.active.corner_radius = egui::CornerRadius::same(8);

    style.visuals.widgets.noninteractive.fg_stroke = egui::Stroke::new(1.0, TEXT_PRIMARY);
    style.visuals.widgets.noninteractive.bg_fill = BG;

    style.visuals.selection.bg_fill = ACCENT.linear_multiply(0.15);
    style.visuals.selection.stroke = egui::Stroke::new(1.0, TEXT_PRIMARY);
    style.visuals.window_stroke = egui::Stroke::new(1.0, BORDER);
    style.visuals.popup_shadow = egui::epaint::Shadow::NONE;

    style.spacing.item_spacing = egui::vec2(8.0, 8.0);
    style.spacing.button_padding = egui::vec2(16.0, 6.0);

    ctx.set_style(style);
}

/// ファイル/フォルダ選択行（パス表示 + 選択ボタン）→ クリック時trueを返す
fn file_select_row(ui: &mut egui::Ui, display: &str) -> bool {
    let total_width = ui.available_width();
    let btn_width = 56.0;
    let path_width = total_width - btn_width - 12.0;
    let mut clicked = false;

    ui.horizontal(|ui| {
        ui.set_max_width(total_width);
        egui::Frame::default()
            .fill(SURFACE)
            .corner_radius(egui::CornerRadius::same(8))
            .stroke(egui::Stroke::new(1.0, BORDER))
            .inner_margin(egui::Margin::symmetric(10, 8))
            .show(ui, |ui| {
                ui.set_max_width(path_width - 24.0);
                ui.set_min_width(path_width - 24.0);
                let text = if display.is_empty() { "未選択" } else { display };
                let color = if display.is_empty() { TEXT_SECONDARY } else { TEXT_PRIMARY };
                ui.add(
                    egui::Label::new(
                        egui::RichText::new(text).size(13.0).color(color),
                    )
                    .truncate(),
                );
            });
        if ui
            .add_sized(
                [btn_width, 32.0],
                egui::Button::new(
                    egui::RichText::new("選択").size(13.0).color(ACCENT),
                )
                .fill(SURFACE)
                .stroke(egui::Stroke::new(1.0, BORDER))
                .corner_radius(egui::CornerRadius::same(8)),
            )
            .clicked()
        {
            clicked = true;
        }
    });
    clicked
}

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_title("カート投入変換ツール")
            .with_inner_size([520.0, 680.0])
            .with_resizable(true),
        ..Default::default()
    };
    eframe::run_native(
        "カート投入変換ツール",
        options,
        Box::new(|cc| {
            let mut fonts = egui::FontDefinitions::default();
            let font_paths = [
                "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
                "C:\\Windows\\Fonts\\meiryo.ttc",
                "C:\\Windows\\Fonts\\YuGothR.ttc",
                "C:\\Windows\\Fonts\\msgothic.ttc",
            ];
            for path in &font_paths {
                if let Ok(data) = std::fs::read(path) {
                    fonts.font_data.insert(
                        "japanese".to_owned(),
                        std::sync::Arc::new(egui::FontData::from_owned(data)),
                    );
                    fonts
                        .families
                        .get_mut(&egui::FontFamily::Proportional)
                        .unwrap()
                        .insert(0, "japanese".to_owned());
                    fonts
                        .families
                        .get_mut(&egui::FontFamily::Monospace)
                        .unwrap()
                        .push("japanese".to_owned());
                    break;
                }
            }
            cc.egui_ctx.set_fonts(fonts);
            setup_style(&cc.egui_ctx);

            let settings = load_settings();
            let output_dir = settings
                .output_dir
                .as_ref()
                .map(PathBuf::from)
                .or_else(desktop_path);
            let selected = settings.selected_preset.min(settings.presets.len().saturating_sub(1));
            let mapping_ui = MappingUi::from_mapping(&settings.presets[selected].mapping);
            Ok(Box::new(App {
                output_dir,
                presets: settings.presets,
                selected_preset: selected,
                mapping_ui,
                preset_name_buf: String::new(),
                ..App::default()
            }))
        }),
    )
}

/// 列マッピングUI用の文字列バッファ
struct MappingUi {
    data_start_row: String,
    id_column: String,
    lot_column: String,
    order_start_column: String,
    order_column_count: String,
    date_header_row: String,
}

impl MappingUi {
    fn from_mapping(m: &ColumnMapping) -> Self {
        Self {
            data_start_row: m.data_start_row.to_string(),
            id_column: m.id_column.clone(),
            lot_column: m.lot_column.clone(),
            order_start_column: m.order_start_column.clone(),
            order_column_count: m.order_column_count.to_string(),
            date_header_row: m.date_header_row.to_string(),
        }
    }

    fn to_mapping(&self) -> Result<ColumnMapping, String> {
        let m = ColumnMapping {
            data_start_row: self.data_start_row.parse::<u32>()
                .map_err(|_| "データ開始行は数値で入力してください")?,
            id_column: self.id_column.trim().to_uppercase(),
            lot_column: self.lot_column.trim().to_uppercase(),
            order_start_column: self.order_start_column.trim().to_uppercase(),
            order_column_count: self.order_column_count.parse::<u32>()
                .map_err(|_| "発注数列数は数値で入力してください")?,
            date_header_row: self.date_header_row.parse::<u32>()
                .map_err(|_| "日付ヘッダー行は数値で入力してください")?,
        };
        m.validate()?;
        Ok(m)
    }
}

impl Default for MappingUi {
    fn default() -> Self {
        Self::from_mapping(&ColumnMapping::default())
    }
}

#[derive(Default)]
struct App {
    product_path: Option<PathBuf>,
    output_dir: Option<PathBuf>,
    presets: Vec<Preset>,
    selected_preset: usize,
    mapping_ui: MappingUi,
    preset_name_buf: String,
    show_mapping: bool,
    auto_detected: Option<String>, // 自動検出されたプリセット名
    confirm_delete_preset: bool,
    log: Vec<LogEntry>,
    progress: f32,
    confirm_overwrite: bool,
    pending_output_dir: Option<PathBuf>,
    completed_dir: Option<PathBuf>,
}

struct LogEntry {
    text: String,
    kind: LogKind,
}

#[derive(PartialEq)]
enum LogKind {
    Info,
    Ok,
    Error,
    Done,
}

impl eframe::App for App {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // サイト風ヘッダーバー
        egui::TopBottomPanel::top("header")
            .frame(
                egui::Frame::default()
                    .fill(HEADER_BG)
                    .inner_margin(egui::Margin::symmetric(20, 10)),
            )
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    ui.label(
                        egui::RichText::new("やさいバス")
                            .size(18.0)
                            .strong()
                            .color(egui::Color32::WHITE),
                    );
                    ui.add_space(12.0);
                    ui.label(
                        egui::RichText::new("カート投入変換ツール")
                            .size(13.0)
                            .color(egui::Color32::from_rgb(200, 220, 190)),
                    );
                });
            });

        // Overwrite confirmation dialog
        if self.confirm_overwrite {
            egui::Window::new("上書き確認")
                .collapsible(false)
                .resizable(false)
                .anchor(egui::Align2::CENTER_CENTER, [0.0, 0.0])
                .frame(
                    egui::Frame::default()
                        .fill(SURFACE)
                        .corner_radius(egui::CornerRadius::same(14))
                        .shadow(egui::epaint::Shadow {
                            offset: [0, 4],
                            blur: 20,
                            spread: 0,
                            color: egui::Color32::from_black_alpha(30),
                        })
                        .inner_margin(egui::Margin::same(24)),
                )
                .show(ctx, |ui| {
                    ui.label(
                        egui::RichText::new("出力先に既存ファイルがあります。\n上書きしますか？")
                            .size(14.0)
                            .color(TEXT_PRIMARY),
                    );
                    ui.add_space(16.0);
                    ui.horizontal(|ui| {
                        if ui
                            .add(
                                egui::Button::new(
                                    egui::RichText::new("上書き")
                                        .color(egui::Color32::WHITE)
                                        .size(13.0),
                                )
                                .fill(ACCENT)
                                .corner_radius(egui::CornerRadius::same(8)),
                            )
                            .clicked()
                        {
                            self.confirm_overwrite = false;
                            if let Some(dir) = self.pending_output_dir.take() {
                                self.do_convert(dir);
                            }
                        }
                        ui.add_space(8.0);
                        if ui
                            .add(
                                egui::Button::new(
                                    egui::RichText::new("キャンセル").size(13.0).color(TEXT_SECONDARY),
                                )
                                .fill(SURFACE)
                                .stroke(egui::Stroke::new(1.0, BORDER))
                                .corner_radius(egui::CornerRadius::same(8)),
                            )
                            .clicked()
                        {
                            self.confirm_overwrite = false;
                            self.pending_output_dir = None;
                        }
                    });
                });
        }

        egui::CentralPanel::default()
            .frame(
                egui::Frame::default()
                    .fill(BG)
                    .inner_margin(egui::Margin::symmetric(32, 28)),
            )
            .show(ctx, |ui| {
                ui.add_space(8.0);

                // File input: Product list
                ui.label(
                    egui::RichText::new("商品リスト")
                        .size(13.0)
                        .color(TEXT_SECONDARY),
                );
                ui.add_space(4.0);
                let product_display = self.product_path.as_ref().map(|p| p.display().to_string()).unwrap_or_default();
                if file_select_row(ui, &product_display) {
                    if let Some(path) = rfd::FileDialog::new()
                        .add_filter("Excel", &["xlsx"])
                        .pick_file()
                    {
                        self.product_path = Some(path.clone());
                        self.auto_detect_preset(&path);
                    }
                }

                // 自動検出表示
                if let Some(name) = &self.auto_detected {
                    ui.label(
                        egui::RichText::new(format!("  → プリセット「{name}」を自動選択"))
                            .size(11.0)
                            .color(ACCENT),
                    );
                }

                ui.add_space(16.0);

                // File input: Output folder
                ui.label(
                    egui::RichText::new("出力先フォルダ")
                        .size(13.0)
                        .color(TEXT_SECONDARY),
                );
                ui.add_space(4.0);
                let output_display = self.output_dir.as_ref().map(|p| p.display().to_string()).unwrap_or_default();
                if file_select_row(ui, &output_display) {
                    if let Some(path) = rfd::FileDialog::new().pick_folder() {
                        self.output_dir = Some(path.clone());
                        self.save_all();
                    }
                }

                ui.add_space(16.0);

                // プリセットタイル選択（常に表示）
                {
                    let mut new_selection: Option<usize> = None;
                    let names: Vec<String> = self.presets.iter().map(|p| p.name.clone()).collect();

                    ui.horizontal_wrapped(|ui| {
                        ui.spacing_mut().item_spacing = egui::vec2(6.0, 6.0);
                        for (i, name) in names.iter().enumerate() {
                            let is_selected = i == self.selected_preset;
                            let (bg, stroke_color, text_color) = if is_selected {
                                (ACCENT, ACCENT, egui::Color32::WHITE)
                            } else {
                                (BG, BORDER, TEXT_PRIMARY)
                            };
                            let btn = egui::Button::new(
                                egui::RichText::new(name).size(12.0).color(text_color)
                            )
                                .fill(bg)
                                .stroke(egui::Stroke::new(1.0, stroke_color))
                                .corner_radius(egui::CornerRadius::same(6))
                                .min_size(egui::vec2(0.0, 28.0));
                            if ui.add(btn).clicked() && !is_selected {
                                new_selection = Some(i);
                            }
                        }
                    });

                    if let Some(i) = new_selection {
                        self.save_current_to_preset();
                        self.selected_preset = i;
                        self.mapping_ui = MappingUi::from_mapping(&self.presets[i].mapping);
                        self.preset_name_buf = self.presets[i].name.clone();
                        self.save_all();
                    }
                }

                ui.add_space(8.0);

                // 列マッピング詳細設定（折りたたみ）
                let toggle_text = if self.show_mapping { "▼ 列マッピング詳細設定" } else { "▶ 列マッピング詳細設定" };
                if ui.add(
                    egui::Button::new(
                        egui::RichText::new(toggle_text).size(12.0).color(TEXT_SECONDARY),
                    )
                    .fill(egui::Color32::TRANSPARENT)
                    .stroke(egui::Stroke::NONE),
                ).clicked() {
                    self.show_mapping = !self.show_mapping;
                }

                if self.show_mapping {
                    egui::Frame::default()
                        .fill(SURFACE)
                        .corner_radius(egui::CornerRadius::same(8))
                        .stroke(egui::Stroke::new(1.0, BORDER))
                        .inner_margin(egui::Margin::symmetric(16, 12))
                        .show(ui, |ui| {
                            let preset_count = self.presets.len();

                            // プリセット名編集 + 追加・削除
                            ui.horizontal(|ui| {
                                ui.label(egui::RichText::new("名前").size(12.0).color(TEXT_SECONDARY));
                                if self.preset_name_buf.is_empty() {
                                    if let Some(p) = self.presets.get(self.selected_preset) {
                                        self.preset_name_buf = p.name.clone();
                                    }
                                }
                                let response = ui.add(
                                    egui::TextEdit::singleline(&mut self.preset_name_buf)
                                        .desired_width(160.0)
                                        .font(egui::TextStyle::Body),
                                );
                                if response.changed() {
                                    if let Some(p) = self.presets.get_mut(self.selected_preset) {
                                        p.name = self.preset_name_buf.clone();
                                    }
                                }
                                if response.lost_focus() {
                                    self.save_all();
                                }

                                // 追加ボタン
                                if ui.add(
                                    egui::Button::new(egui::RichText::new("+").size(13.0).color(ACCENT))
                                        .fill(egui::Color32::TRANSPARENT)
                                        .stroke(egui::Stroke::new(1.0, BORDER))
                                        .corner_radius(egui::CornerRadius::same(6))
                                        .min_size(egui::vec2(28.0, 24.0)),
                                ).clicked() {
                                    self.save_current_to_preset();
                                    let new_name = format!("プリセット{}", self.presets.len() + 1);
                                    self.presets.push(Preset {
                                        name: new_name,
                                        mapping: ColumnMapping::default(),
                                    });
                                    self.selected_preset = self.presets.len() - 1;
                                    self.mapping_ui = MappingUi::from_mapping(&ColumnMapping::default());
                                    self.preset_name_buf = self.presets[self.selected_preset].name.clone();
                                    self.save_all();
                                }

                                // 削除ボタン（2段階確認）
                                let can_delete = preset_count > 1;
                                if self.confirm_delete_preset {
                                    // 確認状態：「本当に削除」「キャンセル」
                                    if ui.add(
                                        egui::Button::new(egui::RichText::new("本当に削除").size(11.0).color(egui::Color32::WHITE))
                                            .fill(ERROR)
                                            .corner_radius(egui::CornerRadius::same(6))
                                            .min_size(egui::vec2(0.0, 24.0)),
                                    ).clicked() {
                                        self.presets.remove(self.selected_preset);
                                        if self.selected_preset >= self.presets.len() {
                                            self.selected_preset = self.presets.len() - 1;
                                        }
                                        self.mapping_ui = MappingUi::from_mapping(&self.presets[self.selected_preset].mapping);
                                        self.preset_name_buf = self.presets[self.selected_preset].name.clone();
                                        self.confirm_delete_preset = false;
                                        self.save_all();
                                    }
                                    if ui.add(
                                        egui::Button::new(egui::RichText::new("キャンセル").size(11.0).color(TEXT_PRIMARY))
                                            .fill(egui::Color32::TRANSPARENT)
                                            .stroke(egui::Stroke::new(1.0, BORDER))
                                            .corner_radius(egui::CornerRadius::same(6))
                                            .min_size(egui::vec2(0.0, 24.0)),
                                    ).clicked() {
                                        self.confirm_delete_preset = false;
                                    }
                                } else {
                                    if ui.add_enabled(can_delete,
                                        egui::Button::new(egui::RichText::new("削除").size(11.0).color(
                                            if can_delete { ERROR } else { TEXT_SECONDARY }
                                        ))
                                            .fill(egui::Color32::TRANSPARENT)
                                            .stroke(egui::Stroke::new(1.0, if can_delete { ERROR } else { BORDER }))
                                            .corner_radius(egui::CornerRadius::same(6))
                                            .min_size(egui::vec2(0.0, 24.0)),
                                    ).clicked() {
                                        self.confirm_delete_preset = true;
                                    }
                                }
                            });

                            ui.add_space(8.0);
                            ui.separator();
                            ui.add_space(4.0);

                            // マッピング入力フィールド
                            let field_width = 60.0;
                            egui::Grid::new("mapping_grid")
                                .num_columns(2)
                                .spacing([12.0, 6.0])
                                .show(ui, |ui| {
                                    ui.label(egui::RichText::new("商品ID列").size(12.0).color(TEXT_PRIMARY));
                                    ui.add(egui::TextEdit::singleline(&mut self.mapping_ui.id_column)
                                        .desired_width(field_width).font(egui::TextStyle::Body));
                                    ui.end_row();

                                    ui.label(egui::RichText::new("ロット列").size(12.0).color(TEXT_PRIMARY));
                                    ui.add(egui::TextEdit::singleline(&mut self.mapping_ui.lot_column)
                                        .desired_width(field_width).font(egui::TextStyle::Body));
                                    ui.end_row();

                                    ui.label(egui::RichText::new("発注数 開始列").size(12.0).color(TEXT_PRIMARY));
                                    ui.add(egui::TextEdit::singleline(&mut self.mapping_ui.order_start_column)
                                        .desired_width(field_width).font(egui::TextStyle::Body));
                                    ui.end_row();

                                    ui.label(egui::RichText::new("発注数 列数").size(12.0).color(TEXT_PRIMARY));
                                    ui.add(egui::TextEdit::singleline(&mut self.mapping_ui.order_column_count)
                                        .desired_width(field_width).font(egui::TextStyle::Body));
                                    ui.end_row();

                                    ui.label(egui::RichText::new("データ開始行").size(12.0).color(TEXT_PRIMARY));
                                    ui.add(egui::TextEdit::singleline(&mut self.mapping_ui.data_start_row)
                                        .desired_width(field_width).font(egui::TextStyle::Body));
                                    ui.end_row();

                                    ui.label(egui::RichText::new("日付ヘッダー行").size(12.0).color(TEXT_PRIMARY));
                                    ui.add(egui::TextEdit::singleline(&mut self.mapping_ui.date_header_row)
                                        .desired_width(field_width).font(egui::TextStyle::Body));
                                    ui.end_row();
                                });

                            // UIの値が変わったらプリセットに反映して保存
                            if let Ok(new_mapping) = self.mapping_ui.to_mapping() {
                                if let Some(p) = self.presets.get_mut(self.selected_preset) {
                                    if p.mapping != new_mapping {
                                        p.mapping = new_mapping;
                                        self.save_all();
                                    }
                                }
                            }
                        });
                }

                ui.add_space(16.0);

                // Convert button
                let btn = ui.add(
                    egui::Button::new(
                        egui::RichText::new("変換実行")
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
                    self.run_convert();
                }

                ui.add_space(20.0);

                // Progress bar
                if self.progress > 0.0 {
                    let (rect, _) = ui.allocate_exact_size(
                        egui::vec2(ui.available_width(), 4.0),
                        egui::Sense::hover(),
                    );
                    ui.painter()
                        .rect_filled(rect, egui::CornerRadius::same(2), PROGRESS_BG);
                    let fill_rect = egui::Rect::from_min_size(
                        rect.min,
                        egui::vec2(rect.width() * self.progress, rect.height()),
                    );
                    ui.painter()
                        .rect_filled(fill_rect, egui::CornerRadius::same(2), ACCENT);
                    ui.add_space(16.0);
                }

                // Log area
                if !self.log.is_empty() {
                    egui::ScrollArea::vertical()
                        .max_height(100.0)
                        .stick_to_bottom(true)
                        .show(ui, |ui| {
                            for entry in &self.log {
                                let color = match entry.kind {
                                    LogKind::Info => TEXT_SECONDARY,
                                    LogKind::Ok => SUCCESS,
                                    LogKind::Error => ERROR,
                                    LogKind::Done => ACCENT,
                                };
                                ui.label(
                                    egui::RichText::new(&entry.text).size(12.0).color(color),
                                );
                            }
                        });
                }

                // フォルダを開くボタン
                if let Some(dir) = &self.completed_dir {
                    let dir = dir.clone();
                    ui.add_space(8.0);
                    if ui
                        .add(
                            egui::Button::new(
                                egui::RichText::new("出力フォルダを開く").size(13.0).color(ACCENT),
                            )
                            .fill(SURFACE)
                            .stroke(egui::Stroke::new(1.0, BORDER))
                            .corner_radius(egui::CornerRadius::same(8)),
                        )
                        .clicked()
                    {
                        #[cfg(target_os = "windows")]
                        { let _ = std::process::Command::new("explorer").arg(&dir).spawn(); }
                        #[cfg(target_os = "macos")]
                        { let _ = std::process::Command::new("open").arg(&dir).spawn(); }
                        #[cfg(target_os = "linux")]
                        { let _ = std::process::Command::new("xdg-open").arg(&dir).spawn(); }
                    }
                }
            });
    }
}

impl App {
    /// 現在のUI値をプリセットに保存
    fn save_current_to_preset(&mut self) {
        if let Ok(mapping) = self.mapping_ui.to_mapping() {
            if let Some(p) = self.presets.get_mut(self.selected_preset) {
                p.mapping = mapping;
            }
        }
    }

    /// 設定をファイルに保存
    fn save_all(&self) {
        let settings = Settings {
            output_dir: self.output_dir.as_ref().map(|p| p.display().to_string()),
            presets: self.presets.clone(),
            selected_preset: self.selected_preset,
            column_mapping: None,
        };
        save_settings(&settings);
    }

    /// ファイルに対してプリセットを自動検出
    fn auto_detect_preset(&mut self, path: &std::path::Path) {
        self.auto_detected = None;
        if self.presets.len() <= 1 {
            return; // プリセットが1つだけなら検出不要
        }

        let mappings: Vec<ColumnMapping> = self.presets.iter().map(|p| p.mapping.clone()).collect();
        let scores = convert::score_presets(path, &mappings);

        // 最高スコアのプリセットを選択（0は除外）
        if let Some((best_idx, &best_score)) = scores.iter().enumerate().max_by_key(|(_, s)| *s) {
            if best_score > 0 && best_idx != self.selected_preset {
                self.save_current_to_preset();
                self.selected_preset = best_idx;
                self.mapping_ui = MappingUi::from_mapping(&self.presets[best_idx].mapping);
                self.preset_name_buf = self.presets[best_idx].name.clone();
                self.auto_detected = Some(self.presets[best_idx].name.clone());
                self.save_all();
            }
        }
    }

    fn run_convert(&mut self) {
        let product_path = match &self.product_path {
            Some(p) => p.clone(),
            None => {
                self.log.push(LogEntry {
                    text: "商品リストを選択してください".into(),
                    kind: LogKind::Error,
                });
                return;
            }
        };
        let base_dir = match &self.output_dir {
            Some(p) => p.clone(),
            None => {
                self.log.push(LogEntry {
                    text: "出力先フォルダを選択してください".into(),
                    kind: LogKind::Error,
                });
                return;
            }
        };

        if !product_path.exists() {
            self.log.push(LogEntry {
                text: format!("ファイルが見つかりません: {}", product_path.display()),
                kind: LogKind::Error,
            });
            return;
        }

        if product_path.extension().and_then(|e| e.to_str()) != Some("xlsx") {
            self.log.push(LogEntry {
                text: ".xlsx ファイルを選択してください".into(),
                kind: LogKind::Error,
            });
            return;
        }

        match convert::get_store_names(&product_path) {
            Ok(names) if names.is_empty() => {
                self.log.push(LogEntry {
                    text: "商品リストにシートがありません".into(),
                    kind: LogKind::Error,
                });
                return;
            }
            Err(e) => {
                self.log.push(LogEntry {
                    text: format!("ファイルを読み込めません: {e}"),
                    kind: LogKind::Error,
                });
                return;
            }
            _ => {}
        }

        let output_dir = base_dir.join(today_folder_name());

        if output_dir.exists() {
            self.confirm_overwrite = true;
            self.pending_output_dir = Some(output_dir);
            return;
        }

        self.do_convert(output_dir);
    }

    fn do_convert(&mut self, output_dir: PathBuf) {
        let product_path = self.product_path.as_ref().unwrap().clone();

        // 現在のUI値からマッピングを取得
        let mapping = match self.mapping_ui.to_mapping() {
            Ok(m) => m,
            Err(e) => {
                self.log.push(LogEntry {
                    text: format!("列マッピング設定エラー: {e}"),
                    kind: LogKind::Error,
                });
                return;
            }
        };
        self.save_current_to_preset();

        let preset_name = self.presets.get(self.selected_preset)
            .map(|p| p.name.as_str())
            .unwrap_or("?");

        self.log.clear();
        self.progress = 0.0;
        self.completed_dir = None;
        self.log.push(LogEntry {
            text: format!("変換開始（{preset_name}） → {}", output_dir.display()),
            kind: LogKind::Info,
        });

        match convert::convert_all(
            &product_path,
            &output_dir,
            &mapping,
            |store, count, current, total| {
                self.log.push(LogEntry {
                    text: format!("  {store}: {count}商品"),
                    kind: LogKind::Ok,
                });
                self.progress = current as f32 / total as f32;
            },
        ) {
            Ok(results) => {
                let total_products: usize = results.iter().map(|(_, c)| c).sum();
                self.log.push(LogEntry {
                    text: format!("完了 — {}店舗 / {}商品", results.len(), total_products),
                    kind: LogKind::Done,
                });
                self.progress = 1.0;
                self.completed_dir = Some(output_dir);
            }
            Err(e) => {
                self.log.push(LogEntry {
                    text: format!("エラー: {e}"),
                    kind: LogKind::Error,
                });
            }
        }
    }
}
