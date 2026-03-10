// Prevents additional console window on Windows in release
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod convert;

use eframe::egui;
use serde::{Deserialize, Serialize};
use std::path::PathBuf;

// やさいバス palette (vegibus.com準拠)
const HEADER_BG: egui::Color32 = egui::Color32::from_rgb(59, 100, 30);     // サイトナビバー濃緑
const BG: egui::Color32 = egui::Color32::from_rgb(255, 255, 255);          // 白背景
const SURFACE: egui::Color32 = egui::Color32::from_rgb(250, 252, 248);     // 薄い緑がかった白
const TEXT_PRIMARY: egui::Color32 = egui::Color32::from_rgb(51, 51, 51);   // 本文黒
const TEXT_SECONDARY: egui::Color32 = egui::Color32::from_rgb(130, 130, 130);
const ACCENT: egui::Color32 = egui::Color32::from_rgb(59, 100, 30);       // ナビバーと同じ濃緑
const BORDER: egui::Color32 = egui::Color32::from_rgb(220, 220, 220);     // グレーボーダー
const SUCCESS: egui::Color32 = egui::Color32::from_rgb(59, 100, 30);
const ERROR: egui::Color32 = egui::Color32::from_rgb(200, 50, 30);
const PROGRESS_BG: egui::Color32 = egui::Color32::from_rgb(230, 230, 230);

// --- Settings ---
#[derive(Serialize, Deserialize, Default)]
struct Settings {
    output_dir: Option<String>,
}

fn settings_path() -> PathBuf {
    let mut path = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|p| p.to_path_buf()))
        .unwrap_or_else(|| PathBuf::from("."));
    path.push("cart-converter-settings.json");
    path
}

fn load_settings() -> Settings {
    std::fs::read_to_string(settings_path())
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

fn save_settings(settings: &Settings) {
    if let Ok(json) = serde_json::to_string_pretty(settings) {
        let _ = std::fs::write(settings_path(), json);
    }
}

/// デスクトップパスを取得（Windows / macOS / Linux 対応）
fn desktop_path() -> Option<PathBuf> {
    // 1. Windows: USERPROFILE, macOS/Linux: HOME
    if let Some(home) = std::env::var_os("USERPROFILE")
        .or_else(|| std::env::var_os("HOME"))
    {
        let desktop = PathBuf::from(&home).join("Desktop");
        if desktop.exists() {
            return Some(desktop);
        }
        // 日本語Windowsの場合「デスクトップ」
        let desktop_jp = PathBuf::from(&home).join("デスクトップ");
        if desktop_jp.exists() {
            return Some(desktop_jp);
        }
        // macOS日本語
        let desktop_mac = PathBuf::from(&home).join("デスクトップ");
        if desktop_mac.exists() {
            return Some(desktop_mac);
        }
        // fallback: ホームディレクトリ
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

    style.visuals.panel_fill = BG;
    style.visuals.window_fill = SURFACE;
    style.visuals.extreme_bg_color = SURFACE;

    // Widgets
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

    style.spacing.item_spacing = egui::vec2(8.0, 8.0);
    style.spacing.button_padding = egui::vec2(16.0, 6.0);

    ctx.set_style(style);
}

/// ファイル/フォルダ選択行（パス表示 + 選択ボタン）→ クリック時trueを返す
fn file_select_row(
    ui: &mut egui::Ui,
    display: &str,
) -> bool {
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
            .with_inner_size([520.0, 560.0])
            .with_resizable(false),
        ..Default::default()
    };
    eframe::run_native(
        "カート投入変換ツール",
        options,
        Box::new(|cc| {
            let mut fonts = egui::FontDefinitions::default();
            let font_paths = [
                // Noto Sans CJK (サイトと同じフォント)
                "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
                // Windows: メイリオ（Noto Sansに近い）優先
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
                .map(PathBuf::from)
                .or_else(desktop_path);
            Ok(Box::new(App {
                output_dir,
                ..App::default()
            }))
        }),
    )
}

#[derive(Default)]
struct App {
    product_path: Option<PathBuf>,
    output_dir: Option<PathBuf>,
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
                        self.product_path = Some(path);
                    }
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
                        save_settings(&Settings {
                            output_dir: Some(path.display().to_string()),
                        });
                    }
                }

                ui.add_space(28.0);

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

        self.log.clear();
        self.progress = 0.0;
        self.completed_dir = None;
        self.log.push(LogEntry {
            text: format!("変換開始 → {}", output_dir.display()),
            kind: LogKind::Info,
        });

        match convert::convert_all(
            &product_path,
            &output_dir,
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
