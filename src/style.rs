use eframe::egui;

// やさいバス palette (vegibus.com準拠)
pub const HEADER_BG: egui::Color32 = egui::Color32::from_rgb(59, 100, 30);
pub const BG: egui::Color32 = egui::Color32::from_rgb(255, 255, 255);
pub const SURFACE: egui::Color32 = egui::Color32::from_rgb(250, 252, 248);
pub const TEXT_PRIMARY: egui::Color32 = egui::Color32::from_rgb(51, 51, 51);
pub const TEXT_SECONDARY: egui::Color32 = egui::Color32::from_rgb(130, 130, 130);
pub const ACCENT: egui::Color32 = egui::Color32::from_rgb(59, 100, 30);
pub const BORDER: egui::Color32 = egui::Color32::from_rgb(220, 220, 220);
pub const SUCCESS: egui::Color32 = egui::Color32::from_rgb(59, 100, 30);
pub const ERROR: egui::Color32 = egui::Color32::from_rgb(200, 50, 30);
pub const PROGRESS_BG: egui::Color32 = egui::Color32::from_rgb(230, 230, 230);

/// やさいバス style を適用
pub fn setup_style(ctx: &egui::Context) {
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
pub fn file_select_row(ui: &mut egui::Ui, display: &str) -> bool {
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

/// ログエントリ
pub struct LogEntry {
    pub text: String,
    pub kind: LogKind,
}

#[derive(PartialEq)]
pub enum LogKind {
    Info,
    Ok,
    Error,
    Done,
}

/// ログ表示エリア
pub fn show_log(ui: &mut egui::Ui, log: &[LogEntry]) {
    if !log.is_empty() {
        egui::ScrollArea::vertical()
            .id_salt("log_area")
            .max_height(100.0)
            .stick_to_bottom(true)
            .show(ui, |ui| {
                for entry in log {
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
}
