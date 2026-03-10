use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::Workbook;
use std::path::Path;

/// 固定ヘッダー（A〜U列、M/N列は空）
const HEADERS: &[&str] = &[
    "id", "name", "farmer", "area", "lot", "spec", "jan", "price",
    "shippingRatePrice", "shippingDay", "productionStart", "productionEnd",
    "", "", // M, N列は空
];

/// 1商品分のデータ
struct Product {
    id: String,
    lot: Option<f64>,
    /// 各曜日の発注数 (P-V列 = 7日分)
    daily_orders: [Option<f64>; 7],
}

/// 商品リストの6行目(index 5)から日付ヘッダーを読み取る (P-V列)
fn read_date_headers(range: &calamine::Range<calamine::Data>) -> Vec<String> {
    (15..22u32)
        .map(|col| {
            range
                .get_value((5, col))
                .map(|v| match v {
                    calamine::Data::DateTime(dt) => {
                        // Excel日付を "M/d" 形式に変換
                        // ExcelDateTimeはas_f64()で日数を返す
                        let days = dt.as_f64();
                        excel_days_to_md(days)
                    }
                    calamine::Data::Float(f) => excel_days_to_md(*f),
                    calamine::Data::String(s) => s.clone(),
                    other => format_cell_value(other),
                })
                .unwrap_or_default()
        })
        .collect()
}

/// Excelシリアル日付 → "M/d" 形式
fn excel_days_to_md(serial: f64) -> String {
    // Excel epoch: 1899-12-30
    let days = serial as i64;
    // 簡易変換: 1900-01-01 = serial 1
    let base = 25569i64; // 1970-01-01のExcelシリアル値
    let unix_days = days - base;
    let ts = unix_days * 86400;

    // 手動で年月日計算
    let (y, m, d) = unix_timestamp_to_ymd(ts);
    let _ = y;
    format!("{}/{}", m, d)
}

/// Unix日数 → (年, 月, 日) - main.rsからも使用
pub fn unix_days_to_ymd(days: i64) -> (i32, u32, u32) {
    unix_timestamp_to_ymd(days * 86400)
}

/// Unixタイムスタンプ → (年, 月, 日)
fn unix_timestamp_to_ymd(ts: i64) -> (i32, u32, u32) {
    let days = (ts / 86400) as i32;
    // 2000-03-01を基準とした計算
    let y400 = 146097; // 400年の日数
    let y100 = 36524;
    let y4 = 1461;

    let mut d = days + 719468; // 0000-03-01からの日数
    let era = if d >= 0 { d } else { d - y400 + 1 } / y400;
    let doe = (d - era * y400) as u32;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe as i32 + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = if mp < 10 { mp + 3 } else { mp - 9 };
    let year = if month <= 2 { y + 1 } else { y };
    (year, month, day)
}

/// 商品リストの1シート(店舗)を読み込む
fn read_store_sheet(
    workbook_path: &Path,
    sheet_name: &str,
) -> Result<Vec<Product>, String> {
    let mut wb: Xlsx<_> =
        open_workbook(workbook_path).map_err(|e| format!("ファイルを開けません: {e}"))?;

    let range = wb
        .worksheet_range(sheet_name)
        .map_err(|e| format!("シート '{sheet_name}' を開けません: {e}"))?;

    let mut products = Vec::new();

    // 8行目(index 7)からデータ行
    for row_idx in 7..range.height() {
        // AF列(index 31) = 商品ID
        let id = match range.get_value((row_idx as u32, 31)) {
            Some(val) => {
                let s = format_cell_value(val);
                if s.is_empty() {
                    continue;
                }
                s
            }
            None => continue,
        };

        // I列(index 8) = ロット
        let lot = get_numeric(&range, row_idx as u32, 8);

        // P-V列(index 15-21) = 各曜日の発注数
        let mut daily_orders = [None; 7];
        for i in 0..7 {
            let val = get_numeric(&range, row_idx as u32, 15 + i);
            if let Some(v) = val {
                if v != 0.0 {
                    daily_orders[i as usize] = Some(v);
                }
            }
        }

        products.push(Product {
            id,
            lot,
            daily_orders,
        });
    }

    Ok(products)
}

/// セル値を文字列に変換（数値はintっぽければ整数表示）
fn format_cell_value(val: &calamine::Data) -> String {
    match val {
        calamine::Data::Int(n) => n.to_string(),
        calamine::Data::Float(f) => {
            if *f == (*f as i64) as f64 {
                (*f as i64).to_string()
            } else {
                f.to_string()
            }
        }
        calamine::Data::String(s) => s.clone(),
        calamine::Data::Bool(b) => b.to_string(),
        calamine::Data::Empty => String::new(),
        _ => String::new(),
    }
}

/// 数値セルを取得
fn get_numeric(
    range: &calamine::Range<calamine::Data>,
    row: u32,
    col: u32,
) -> Option<f64> {
    match range.get_value((row, col))? {
        calamine::Data::Float(f) => Some(*f),
        calamine::Data::Int(n) => Some(*n as f64),
        _ => None,
    }
}

/// 商品リストの全シート名（店舗名）を取得
pub fn get_store_names(workbook_path: &Path) -> Result<Vec<String>, String> {
    let wb: Xlsx<_> =
        open_workbook(workbook_path).map_err(|e| format!("ファイルを開けません: {e}"))?;
    Ok(wb.sheet_names().to_vec())
}

/// 商品リストから日付ヘッダーを取得（最初のシートから）
fn get_date_headers(workbook_path: &Path) -> Result<Vec<String>, String> {
    let mut wb: Xlsx<_> =
        open_workbook(workbook_path).map_err(|e| format!("ファイルを開けません: {e}"))?;
    let first_sheet = wb.sheet_names().first().cloned()
        .ok_or("シートがありません")?;
    let range = wb
        .worksheet_range(&first_sheet)
        .map_err(|e| format!("シートを開けません: {e}"))?;
    Ok(read_date_headers(&range))
}

/// 1店舗分のカート投入用xlsxを生成（テンプレート不要版）
pub fn write_cart_file(
    workbook_path: &Path,
    sheet_name: &str,
    date_headers: &[String],
    output_path: &Path,
) -> Result<usize, String> {
    let products = read_store_sheet(workbook_path, sheet_name)?;

    let mut wb = Workbook::new();
    let ws = wb.add_worksheet().set_name("list").map_err(|e| e.to_string())?;

    // ヘッダー行 (4行目 = index 3)
    for (col, header) in HEADERS.iter().enumerate() {
        if !header.is_empty() {
            ws.write_string(3, col as u16, *header).map_err(|e| e.to_string())?;
        }
    }
    // O-U列 (index 14-20): 日付ヘッダー
    for (j, date) in date_headers.iter().enumerate() {
        if !date.is_empty() {
            ws.write_string(3, (14 + j) as u16, date).map_err(|e| e.to_string())?;
        }
    }

    // データ行 (5行目〜 = index 4〜)
    for (i, product) in products.iter().enumerate() {
        let row = (i + 4) as u32;

        // A列: id
        ws.write_string(row, 0, &product.id).map_err(|e| e.to_string())?;

        // E列: lot
        if let Some(lot) = product.lot {
            ws.write_number(row, 4, lot).map_err(|e| e.to_string())?;
        }

        // O-U列 (index 14-20): 各日の発注数
        for (j, order) in product.daily_orders.iter().enumerate() {
            if let Some(val) = order {
                ws.write_number(row, (14 + j) as u16, *val)
                    .map_err(|e| e.to_string())?;
            }
        }
    }

    wb.save(output_path).map_err(|e| format!("保存失敗: {e}"))?;
    Ok(products.len())
}

/// 全店舗を一括変換（テンプレート不要版）
pub fn convert_all(
    product_list_path: &Path,
    output_dir: &Path,
    mut on_progress: impl FnMut(&str, usize, usize, usize),
) -> Result<Vec<(String, usize)>, String> {
    let store_names = get_store_names(product_list_path)?;
    let date_headers = get_date_headers(product_list_path)?;
    let total = store_names.len();
    let mut results = Vec::new();

    std::fs::create_dir_all(output_dir)
        .map_err(|e| format!("出力フォルダを作成できません: {e}"))?;

    for (idx, store) in store_names.iter().enumerate() {
        let output_path = output_dir.join(format!("カート投入用_{store}.xlsx"));
        let count = write_cart_file(product_list_path, store, &date_headers, &output_path)?;
        results.push((store.clone(), count));
        on_progress(store, count, idx + 1, total);
    }

    Ok(results)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_convert() {
        let product_path = Path::new("../【集計】2w商品リスト (3月4日～3月10日) (4).xlsx");
        if !product_path.exists() {
            eprintln!("テストファイルが見つかりません。スキップ。");
            return;
        }

        let output_dir = Path::new("../test_output_rs2");
        let results = convert_all(product_path, output_dir, |store, count, current, total| {
            println!("{current}/{total} {store}: {count}商品");
        })
        .unwrap();

        assert!(!results.is_empty());
        for (store, count) in &results {
            println!("{store}: {count}商品");
            assert!(*count > 0);
        }
    }
}
