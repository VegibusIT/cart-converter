// Prevents additional console window on Windows in release
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod convert;

use eframe::egui;
use std::path::PathBuf;

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_title("カート投入変換ツール")
            .with_inner_size([580.0, 440.0])
            .with_resizable(false),
        ..Default::default()
    };
    eframe::run_native(
        "カート投入変換ツール",
        options,
        Box::new(|_cc| Ok(Box::new(App::default()))),
    )
}

#[derive(Default)]
struct App {
    product_path: Option<PathBuf>,
    template_path: Option<PathBuf>,
    output_dir: Option<PathBuf>,
    log: Vec<LogEntry>,
    progress: f32,
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
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("カート投入変換ツール");
            ui.add_space(8.0);

            // 商品リスト
            ui.horizontal(|ui| {
                ui.label("商品リスト:");
                let text = self
                    .product_path
                    .as_ref()
                    .map(|p| p.display().to_string())
                    .unwrap_or_else(|| "未選択".into());
                ui.add(egui::TextEdit::singleline(&mut text.as_str()).desired_width(370.0));
                if ui.button("参照").clicked() {
                    if let Some(path) = rfd::FileDialog::new()
                        .add_filter("Excel", &["xlsx"])
                        .pick_file()
                    {
                        self.product_path = Some(path);
                    }
                }
            });

            ui.add_space(4.0);

            // カート投入用原本
            ui.horizontal(|ui| {
                ui.label("カート投入用原本:");
                let text = self
                    .template_path
                    .as_ref()
                    .map(|p| p.display().to_string())
                    .unwrap_or_else(|| "未選択".into());
                ui.add(egui::TextEdit::singleline(&mut text.as_str()).desired_width(370.0));
                if ui.button("参照").clicked() {
                    if let Some(path) = rfd::FileDialog::new()
                        .add_filter("Excel", &["xlsx"])
                        .pick_file()
                    {
                        self.template_path = Some(path);
                    }
                }
            });

            ui.add_space(4.0);

            // 出力先フォルダ
            ui.horizontal(|ui| {
                ui.label("出力先フォルダ:");
                let text = self
                    .output_dir
                    .as_ref()
                    .map(|p| p.display().to_string())
                    .unwrap_or_else(|| "未選択".into());
                ui.add(egui::TextEdit::singleline(&mut text.as_str()).desired_width(370.0));
                if ui.button("参照").clicked() {
                    if let Some(path) = rfd::FileDialog::new().pick_folder() {
                        self.output_dir = Some(path);
                    }
                }
            });

            ui.add_space(12.0);

            // 変換実行ボタン
            let btn = ui.add(
                egui::Button::new("変換実行").min_size(egui::vec2(ui.available_width(), 36.0)),
            );
            if btn.clicked() {
                self.run_convert();
            }

            ui.add_space(8.0);

            // プログレスバー
            ui.add(egui::ProgressBar::new(self.progress).show_percentage());

            ui.add_space(8.0);

            // ログ
            egui::ScrollArea::vertical()
                .max_height(160.0)
                .stick_to_bottom(true)
                .show(ui, |ui| {
                    for entry in &self.log {
                        let color = match entry.kind {
                            LogKind::Info => egui::Color32::GRAY,
                            LogKind::Ok => egui::Color32::from_rgb(39, 103, 73),
                            LogKind::Error => egui::Color32::from_rgb(197, 48, 48),
                            LogKind::Done => egui::Color32::from_rgb(44, 82, 130),
                        };
                        ui.colored_label(color, &entry.text);
                    }
                });
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
        let template_path = match &self.template_path {
            Some(p) => p.clone(),
            None => {
                self.log.push(LogEntry {
                    text: "カート投入用原本を選択してください".into(),
                    kind: LogKind::Error,
                });
                return;
            }
        };
        let output_dir = match &self.output_dir {
            Some(p) => p.clone(),
            None => {
                self.log.push(LogEntry {
                    text: "出力先フォルダを選択してください".into(),
                    kind: LogKind::Error,
                });
                return;
            }
        };

        self.log.clear();
        self.progress = 0.0;
        self.log.push(LogEntry {
            text: "変換開始...".into(),
            kind: LogKind::Info,
        });

        match convert::convert_all(
            &product_path,
            &template_path,
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
                    text: format!("完了！ {}店舗 / {}商品", results.len(), total_products),
                    kind: LogKind::Done,
                });
                self.progress = 1.0;
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
