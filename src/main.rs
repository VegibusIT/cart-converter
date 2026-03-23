// Prevents additional console window on Windows in release
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod aeon;
mod cart_converter;
mod convert;
mod entetu;
mod google_auth;
mod style;
mod updater;

use eframe::egui;
use style::*;

/// アプリの現在表示しているページ
enum Page {
    Top,
    CartConverter,
    EnteTu,
    Aeon,
}

struct App {
    page: Page,
    cart_converter: cart_converter::CartConverterPage,
    entetu: entetu::EnteTuPage,
    aeon: aeon::AeonPage,

    // バージョン管理
    releases: Vec<updater::ReleaseInfo>,
    releases_loaded: bool,
    releases_error: Option<String>,
    show_versions: bool,
    update_status: Option<String>,

    // 更新通知
    update_check_done: bool,
    latest_version: Option<updater::ReleaseInfo>,
    show_update_popup: bool,
}

impl Default for App {
    fn default() -> Self {
        Self {
            page: Page::Top,
            cart_converter: cart_converter::CartConverterPage::default(),
            entetu: entetu::EnteTuPage::default(),
            aeon: aeon::AeonPage::default(),
            releases: Vec::new(),
            releases_loaded: false,
            releases_error: None,
            show_versions: false,
            update_status: None,
            update_check_done: false,
            latest_version: None,
            show_update_popup: false,
        }
    }
}

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_title("やさいバス ツール")
            .with_inner_size([520.0, 680.0])
            .with_resizable(true),
        ..Default::default()
    };
    eframe::run_native(
        "やさいバス ツール",
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
            Ok(Box::new(App::default()))
        }),
    )
}

impl eframe::App for App {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // 起動時に最新バージョンチェック（1回のみ）
        if !self.update_check_done {
            self.update_check_done = true;
            if let Ok(releases) = updater::fetch_releases() {
                if let Some(latest) = releases.first() {
                    if !latest.is_current {
                        self.latest_version = Some(latest.clone());
                        self.show_update_popup = true;
                    }
                }
                self.releases = releases;
                self.releases_loaded = true;
            }
        }

        // 更新通知ポップアップ
        if self.show_update_popup {
            if let Some(latest) = self.latest_version.clone() {
                egui::Window::new("更新のお知らせ")
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
                            egui::RichText::new("新しいバージョンがあります")
                                .size(15.0)
                                .strong()
                                .color(TEXT_PRIMARY),
                        );
                        ui.add_space(8.0);
                        ui.label(
                            egui::RichText::new(format!(
                                "現在: v{}  →  最新: {}",
                                updater::current_version(),
                                latest.version
                            ))
                            .size(13.0)
                            .color(TEXT_PRIMARY),
                        );
                        ui.add_space(8.0);

                        // リリースノート
                        if !latest.release_notes.is_empty() {
                            egui::Frame::default()
                                .fill(BG)
                                .corner_radius(egui::CornerRadius::same(6))
                                .stroke(egui::Stroke::new(1.0, BORDER))
                                .inner_margin(egui::Margin::symmetric(12, 8))
                                .show(ui, |ui| {
                                    ui.label(
                                        egui::RichText::new(&latest.release_notes)
                                            .size(12.0)
                                            .color(TEXT_PRIMARY),
                                    );
                                });
                        }
                        ui.add_space(12.0);

                        if let Some(status) = &self.update_status {
                            ui.label(
                                egui::RichText::new(status)
                                    .size(12.0)
                                    .color(ACCENT),
                            );
                            ui.add_space(8.0);
                        }

                        ui.horizontal(|ui| {
                            if ui
                                .add(
                                    egui::Button::new(
                                        egui::RichText::new("今すぐ更新")
                                            .color(egui::Color32::WHITE)
                                            .size(13.0),
                                    )
                                    .fill(ACCENT)
                                    .corner_radius(egui::CornerRadius::same(8)),
                                )
                                .clicked()
                            {
                                self.update_status = Some("ダウンロード中...".into());
                                match updater::download_and_replace(&latest) {
                                    Ok(exe_path) => {
                                        updater::restart_app(&exe_path);
                                    }
                                    Err(e) => {
                                        self.update_status = Some(format!("更新失敗: {e}"));
                                    }
                                }
                            }
                            ui.add_space(8.0);
                            if ui
                                .add(
                                    egui::Button::new(
                                        egui::RichText::new("あとで")
                                            .size(13.0)
                                            .color(TEXT_SECONDARY),
                                    )
                                    .fill(SURFACE)
                                    .stroke(egui::Stroke::new(1.0, BORDER))
                                    .corner_radius(egui::CornerRadius::same(8)),
                                )
                                .clicked()
                            {
                                self.show_update_popup = false;
                            }
                        });
                    });
            }
        }

        // ヘッダーバー
        egui::TopBottomPanel::top("header")
            .frame(
                egui::Frame::default()
                    .fill(HEADER_BG)
                    .inner_margin(egui::Margin::symmetric(20, 10)),
            )
            .show(ctx, |ui| {
                ui.horizontal(|ui| {
                    // トップ以外では戻るボタンを表示
                    if !matches!(self.page, Page::Top) {
                        let back = ui.add(
                            egui::Button::new(
                                egui::RichText::new("<")
                                    .size(18.0)
                                    .strong()
                                    .color(egui::Color32::WHITE),
                            )
                            .fill(egui::Color32::TRANSPARENT)
                            .stroke(egui::Stroke::NONE),
                        );
                        if back.clicked() {
                            self.page = Page::Top;
                        }
                        if back.hovered() {
                            ctx.set_cursor_icon(egui::CursorIcon::PointingHand);
                        }
                    }

                    // ロゴ（クリックでトップに戻る）
                    let logo = ui.add(
                        egui::Button::new(
                            egui::RichText::new("やさいバス")
                                .size(18.0)
                                .strong()
                                .color(egui::Color32::WHITE),
                        )
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::NONE),
                    );
                    if logo.clicked() {
                        self.page = Page::Top;
                    }
                    if logo.hovered() {
                        ctx.set_cursor_icon(egui::CursorIcon::PointingHand);
                    }

                    ui.add_space(12.0);

                    let subtitle = match &self.page {
                        Page::Top => "ツール一覧",
                        Page::CartConverter => "カート投入変換ツール",
                        Page::EnteTu => "遠鉄ストア消化仕入れ",
                        Page::Aeon => "イオン近畿 生鮮MD",
                    };
                    ui.label(
                        egui::RichText::new(subtitle)
                            .size(13.0)
                            .color(egui::Color32::from_rgb(200, 220, 190)),
                    );
                });
            });

        // ページ描画
        match &self.page {
            Page::Top => self.show_top_page(ctx),
            Page::CartConverter => {
                if self.cart_converter.show(ctx) {
                    self.page = Page::Top;
                }
            }
            Page::EnteTu => {
                if self.entetu.show(ctx) {
                    self.page = Page::Top;
                }
            }
            Page::Aeon => {
                if self.aeon.show(ctx) {
                    self.page = Page::Top;
                }
            }
        }
    }
}

impl App {
    fn show_top_page(&mut self, ctx: &egui::Context) {
        egui::CentralPanel::default()
            .frame(
                egui::Frame::default()
                    .fill(BG)
                    .inner_margin(egui::Margin::symmetric(32, 28)),
            )
            .show(ctx, |ui| {
                ui.add_space(20.0);

                ui.label(
                    egui::RichText::new("ツール一覧")
                        .size(20.0)
                        .strong()
                        .color(TEXT_PRIMARY),
                );
                ui.add_space(4.0);
                ui.label(
                    egui::RichText::new("使用するツールを選択してください")
                        .size(13.0)
                        .color(TEXT_SECONDARY),
                );
                ui.add_space(24.0);

                let card_width = ui.available_width();

                // カート投入変換ツール
                if self.show_tool_card(
                    ui,
                    ctx,
                    card_width,
                    "カート投入変換ツール",
                    "商品リスト（Excel）からカート投入用ファイルを生成します",
                ) {
                    self.page = Page::CartConverter;
                }

                ui.add_space(12.0);

                // 遠鉄ストア消化仕入れ
                if self.show_tool_card(
                    ui,
                    ctx,
                    card_width,
                    "遠鉄ストア 消化仕入れ転記",
                    "Google Driveのメールデータをスプレッドシートに自動転記します",
                ) {
                    self.page = Page::EnteTu;
                }

                ui.add_space(12.0);

                // イオン近畿 生鮮MD
                if self.show_tool_card(
                    ui,
                    ctx,
                    card_width,
                    "イオン近畿 生鮮MDアップロード",
                    "商品リスト（Excel）から納品日別のアップロード用ファイルを生成します",
                ) {
                    self.page = Page::Aeon;
                }

                // バージョン管理セクション（画面下部）
                ui.add_space(24.0);
                ui.separator();
                ui.add_space(8.0);
                self.show_version_section(ui, ctx);
            });
    }

    fn show_version_section(&mut self, ui: &mut egui::Ui, ctx: &egui::Context) {
        ui.horizontal(|ui| {
            ui.label(
                egui::RichText::new(format!("v{}", updater::current_version()))
                    .size(12.0)
                    .color(TEXT_SECONDARY),
            );
            ui.add_space(8.0);

            let btn_text = if self.show_versions { "▼ バージョン管理" } else { "▶ バージョン管理" };
            if ui
                .add(
                    egui::Button::new(
                        egui::RichText::new(btn_text).size(12.0).color(TEXT_SECONDARY),
                    )
                    .fill(egui::Color32::TRANSPARENT)
                    .stroke(egui::Stroke::NONE),
                )
                .clicked()
            {
                self.show_versions = !self.show_versions;
                if self.show_versions && !self.releases_loaded {
                    match updater::fetch_releases() {
                        Ok(releases) => {
                            self.releases = releases;
                            self.releases_loaded = true;
                            self.releases_error = None;
                        }
                        Err(e) => {
                            self.releases_error = Some(e);
                        }
                    }
                }
            }
        });

        if !self.show_versions {
            return;
        }

        ui.add_space(8.0);

        // ステータスメッセージ
        if let Some(status) = &self.update_status {
            ui.label(
                egui::RichText::new(status)
                    .size(12.0)
                    .color(ACCENT),
            );
            ui.add_space(4.0);
        }

        if let Some(err) = &self.releases_error {
            ui.label(
                egui::RichText::new(format!("エラー: {err}"))
                    .size(12.0)
                    .color(ERROR),
            );
            return;
        }

        if self.releases.is_empty() {
            ui.label(
                egui::RichText::new("リリースが見つかりません")
                    .size(12.0)
                    .color(TEXT_SECONDARY),
            );
            return;
        }

        // バージョン一覧
        egui::Frame::default()
            .fill(SURFACE)
            .corner_radius(egui::CornerRadius::same(8))
            .stroke(egui::Stroke::new(1.0, BORDER))
            .inner_margin(egui::Margin::symmetric(16, 12))
            .show(ui, |ui| {
                for release in self.releases.clone() {
                    ui.horizontal(|ui| {
                        // バージョン名
                        let text_color = if release.is_current { ACCENT } else { TEXT_PRIMARY };
                        ui.label(
                            egui::RichText::new(&release.version)
                                .size(13.0)
                                .strong()
                                .color(text_color),
                        );

                        if release.is_current {
                            ui.label(
                                egui::RichText::new("（現在）")
                                    .size(11.0)
                                    .color(ACCENT),
                            );
                        } else {
                            // 切り替えボタン
                            let btn_label = if release.version > format!("v{}", updater::current_version()) {
                                "アップデート"
                            } else {
                                "ダウングレード"
                            };
                            if ui
                                .add(
                                    egui::Button::new(
                                        egui::RichText::new(btn_label)
                                            .size(11.0)
                                            .color(egui::Color32::WHITE),
                                    )
                                    .fill(ACCENT)
                                    .corner_radius(egui::CornerRadius::same(4)),
                                )
                                .clicked()
                            {
                                self.update_status = Some(format!("{} をダウンロード中...", release.version));
                                match updater::download_and_replace(&release) {
                                    Ok(exe_path) => {
                                        self.update_status = Some(format!(
                                            "{} のダウンロード完了。再起動します...",
                                            release.version
                                        ));
                                        ctx.request_repaint();
                                        updater::restart_app(&exe_path);
                                    }
                                    Err(e) => {
                                        self.update_status = Some(format!("更新失敗: {e}"));
                                    }
                                }
                            }
                        }
                    });
                    ui.add_space(2.0);
                }
            });
    }

    /// ツールカードを表示。クリックされたらtrueを返す
    fn show_tool_card(
        &self,
        ui: &mut egui::Ui,
        ctx: &egui::Context,
        width: f32,
        title: &str,
        description: &str,
    ) -> bool {
        let mut clicked = false;

        let response = egui::Frame::default()
            .fill(SURFACE)
            .corner_radius(egui::CornerRadius::same(12))
            .stroke(egui::Stroke::new(1.0, BORDER))
            .inner_margin(egui::Margin::symmetric(20, 16))
            .show(ui, |ui| {
                ui.set_min_width(width - 44.0);
                ui.horizontal(|ui| {
                    ui.vertical(|ui| {
                        ui.label(
                            egui::RichText::new(title)
                                .size(15.0)
                                .strong()
                                .color(TEXT_PRIMARY),
                        );
                        ui.add_space(4.0);
                        ui.label(
                            egui::RichText::new(description)
                                .size(12.0)
                                .color(TEXT_SECONDARY),
                        );
                    });
                    ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                        ui.label(
                            egui::RichText::new(">")
                                .size(18.0)
                                .color(ACCENT),
                        );
                    });
                });
            });

        if response.response.interact(egui::Sense::click()).clicked() {
            clicked = true;
        }
        if response.response.interact(egui::Sense::hover()).hovered() {
            ctx.set_cursor_icon(egui::CursorIcon::PointingHand);
        }

        clicked
    }
}
