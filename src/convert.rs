use calamine::{open_workbook, Reader, Xlsx};
use rust_xlsxwriter::Workbook;
use serde::{Deserialize, Serialize};
use std::path::Path;

/// 固定ヘッダー（A〜U列、M/N列は空）
const HEADERS: &[&str] = &[
    "id", "name", "farmer", "area", "lot", "spec", "jan", "price",
    "shippingRatePrice", "shippingDay", "productionStart", "productionEnd",
    "", "", // M, N列は空
];

/// 元データの列マッピング設定
#[derive(Serialize, Deserialize, Clone, Debug, PartialEq)]
pub struct ColumnMapping {
    /// データ開始行（1始まり、デフォルト: 8）
    pub data_start_row: u32,
    /// 商品ID列（Excel列名、デフォルト: "AF"）
    pub id_column: String,
    /// ロット列（Excel列名、デフォルト: "I"）
    pub lot_column: String,
    /// 発注数の開始列（Excel列名、デフォルト: "P"）
    pub order_start_column: String,
    /// 発注数の列数（デフォルト: 7）
    pub order_column_count: u32,
    /// 日付ヘッダー行（1始まり、デフォルト: 6）
    pub date_header_row: u32,
}

impl Default for ColumnMapping {
    fn default() -> Self {
        Self {
            data_start_row: 8,
            id_column: "AF".to_string(),
            lot_column: "I".to_string(),
            order_start_column: "P".to_string(),
            order_column_count: 7,
            date_header_row: 6,
        }
    }
}

impl ColumnMapping {
    /// マッピング設定のバリデーション
    pub fn validate(&self) -> Result<(), String> {
        if self.data_start_row == 0 {
            return Err("データ開始行は1以上を指定してください".into());
        }
        if self.date_header_row == 0 {
            return Err("日付ヘッダー行は1以上を指定してください".into());
        }
        if self.order_column_count == 0 {
            return Err("発注数の列数は1以上を指定してください".into());
        }
        column_name_to_index(&self.id_column)
            .ok_or_else(|| format!("無効な列名（商品ID）: {}", self.id_column))?;
        column_name_to_index(&self.lot_column)
            .ok_or_else(|| format!("無効な列名（ロット）: {}", self.lot_column))?;
        column_name_to_index(&self.order_start_column)
            .ok_or_else(|| format!("無効な列名（発注数開始列）: {}", self.order_start_column))?;
        Ok(())
    }
}

/// Excel列名 → 0始まりインデックス ("A"→0, "B"→1, ..., "Z"→25, "AA"→26, "AF"→31)
pub fn column_name_to_index(name: &str) -> Option<u32> {
    let name = name.trim().to_uppercase();
    if name.is_empty() || !name.chars().all(|c| c.is_ascii_alphabetic()) {
        return None;
    }
    let mut idx: u32 = 0;
    for c in name.chars() {
        idx = idx * 26 + (c as u32 - 'A' as u32 + 1);
    }
    Some(idx - 1)
}

/// 0始まりインデックス → Excel列名 (0→"A", 25→"Z", 26→"AA", 31→"AF")
pub fn index_to_column_name(mut idx: u32) -> String {
    let mut result = String::new();
    loop {
        result.insert(0, (b'A' + (idx % 26) as u8) as char);
        if idx < 26 {
            break;
        }
        idx = idx / 26 - 1;
    }
    result
}

/// 1商品分のデータ
struct Product {
    id: String,
    lot: Option<f64>,
    /// 各曜日の発注数
    daily_orders: Vec<Option<f64>>,
}

/// 商品リストから日付ヘッダーを読み取る
fn read_date_headers(range: &calamine::Range<calamine::Data>, mapping: &ColumnMapping) -> Vec<String> {
    let order_start = column_name_to_index(&mapping.order_start_column).unwrap_or(15);
    let date_row = mapping.date_header_row.saturating_sub(1); // 1始まり→0始まり
    (order_start..order_start + mapping.order_column_count)
        .map(|col| {
            range
                .get_value((date_row, col))
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
    mapping: &ColumnMapping,
) -> Result<Vec<Product>, String> {
    let mut wb: Xlsx<_> =
        open_workbook(workbook_path).map_err(|e| format!("ファイルを開けません: {e}"))?;

    let range = wb
        .worksheet_range(sheet_name)
        .map_err(|e| format!("シート '{sheet_name}' を開けません: {e}"))?;

    let id_col = column_name_to_index(&mapping.id_column)
        .ok_or_else(|| format!("無効な列名: {}", mapping.id_column))?;
    let lot_col = column_name_to_index(&mapping.lot_column)
        .ok_or_else(|| format!("無効な列名: {}", mapping.lot_column))?;
    let order_start = column_name_to_index(&mapping.order_start_column)
        .ok_or_else(|| format!("無効な列名: {}", mapping.order_start_column))?;
    let data_start = mapping.data_start_row.saturating_sub(1) as usize; // 1始まり→0始まり

    let mut products = Vec::new();

    for row_idx in data_start..range.height() {
        let id = match range.get_value((row_idx as u32, id_col)) {
            Some(val) => {
                let s = format_cell_value(val);
                if s.is_empty() {
                    continue;
                }
                s
            }
            None => continue,
        };

        let lot = get_numeric(&range, row_idx as u32, lot_col);

        let mut daily_orders = vec![None; mapping.order_column_count as usize];
        for i in 0..mapping.order_column_count {
            let val = get_numeric(&range, row_idx as u32, order_start + i);
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
fn get_date_headers(workbook_path: &Path, mapping: &ColumnMapping) -> Result<Vec<String>, String> {
    let mut wb: Xlsx<_> =
        open_workbook(workbook_path).map_err(|e| format!("ファイルを開けません: {e}"))?;
    let first_sheet = wb.sheet_names().first().cloned()
        .ok_or("シートがありません")?;
    let range = wb
        .worksheet_range(&first_sheet)
        .map_err(|e| format!("シートを開けません: {e}"))?;
    Ok(read_date_headers(&range, mapping))
}

/// 1店舗分のカート投入用xlsxを生成（テンプレート不要版）
pub fn write_cart_file(
    workbook_path: &Path,
    sheet_name: &str,
    date_headers: &[String],
    output_path: &Path,
    mapping: &ColumnMapping,
) -> Result<usize, String> {
    let products = read_store_sheet(workbook_path, sheet_name, mapping)?;

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

        // O列〜: 各日の発注数
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
    mapping: &ColumnMapping,
    mut on_progress: impl FnMut(&str, usize, usize, usize),
) -> Result<Vec<(String, usize)>, String> {
    // マッピングのバリデーション
    mapping.validate()?;

    let store_names = get_store_names(product_list_path)?;
    let date_headers = get_date_headers(product_list_path, mapping)?;
    let total = store_names.len();
    let mut results = Vec::new();

    std::fs::create_dir_all(output_dir)
        .map_err(|e| format!("出力フォルダを作成できません: {e}"))?;

    for (idx, store) in store_names.iter().enumerate() {
        let output_path = output_dir.join(format!("カート投入用_{store}.xlsx"));
        let count = write_cart_file(product_list_path, store, &date_headers, &output_path, mapping)?;
        results.push((store.clone(), count));
        on_progress(store, count, idx + 1, total);
    }

    Ok(results)
}

/// ファイルに対して各プリセットのマッチ度をスコアリング（商品IDが見つかった行数）
/// 戻り値: 各プリセットのスコア（0=マッチしない）
pub fn score_presets(
    workbook_path: &Path,
    presets: &[ColumnMapping],
) -> Vec<usize> {
    let mut scores = vec![0usize; presets.len()];
    let Ok(mut wb) = open_workbook::<Xlsx<_>, _>(workbook_path) else {
        return scores;
    };
    let Some(first_sheet) = wb.sheet_names().first().cloned() else {
        return scores;
    };
    let Ok(range) = wb.worksheet_range(&first_sheet) else {
        return scores;
    };

    for (idx, mapping) in presets.iter().enumerate() {
        let Some(id_col) = column_name_to_index(&mapping.id_column) else {
            continue;
        };
        let data_start = mapping.data_start_row.saturating_sub(1) as usize;
        // 最大20行チェック
        let end = (data_start + 20).min(range.height());
        let mut count = 0usize;
        for row in data_start..end {
            if let Some(val) = range.get_value((row as u32, id_col)) {
                let s = format_cell_value(val);
                if !s.is_empty() {
                    count += 1;
                }
            }
        }
        scores[idx] = count;
    }
    scores
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_name_conversion() {
        assert_eq!(column_name_to_index("A"), Some(0));
        assert_eq!(column_name_to_index("B"), Some(1));
        assert_eq!(column_name_to_index("Z"), Some(25));
        assert_eq!(column_name_to_index("AA"), Some(26));
        assert_eq!(column_name_to_index("AF"), Some(31));
        assert_eq!(column_name_to_index("I"), Some(8));
        assert_eq!(column_name_to_index("P"), Some(15));
        assert_eq!(column_name_to_index(""), None);
        assert_eq!(column_name_to_index("1"), None);

        assert_eq!(index_to_column_name(0), "A");
        assert_eq!(index_to_column_name(25), "Z");
        assert_eq!(index_to_column_name(26), "AA");
        assert_eq!(index_to_column_name(31), "AF");
    }

    #[test]
    fn test_convert() {
        let product_path = Path::new("../【集計】2w商品リスト (3月4日～3月10日) (4).xlsx");
        if !product_path.exists() {
            eprintln!("テストファイルが見つかりません。スキップ。");
            return;
        }

        let output_dir = Path::new("../test_output_rs2");
        let mapping = ColumnMapping::default();
        let results = convert_all(product_path, output_dir, &mapping, |store, count, current, total| {
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
