use serde::Deserialize;
use std::path::PathBuf;

const GITHUB_REPO: &str = "VegibusIT/cart-converter";
const CURRENT_VERSION: &str = env!("CARGO_PKG_VERSION");

#[derive(Deserialize, Clone, Debug)]
struct GitHubRelease {
    tag_name: String,
    body: Option<String>,
    assets: Vec<GitHubAsset>,
}

#[derive(Deserialize, Clone, Debug)]
struct GitHubAsset {
    name: String,
    browser_download_url: String,
}

#[derive(Clone, Debug)]
pub struct ReleaseInfo {
    pub version: String,
    pub download_url: String,
    pub is_current: bool,
    pub release_notes: String,
}

/// 現在のバージョンを返す
pub fn current_version() -> &'static str {
    CURRENT_VERSION
}

/// GitHub Releasesから全バージョン一覧を取得
pub fn fetch_releases() -> Result<Vec<ReleaseInfo>, String> {
    let url = format!(
        "https://api.github.com/repos/{}/releases",
        GITHUB_REPO
    );

    let resp = ureq::get(&url)
        .set("User-Agent", "cart-converter-updater")
        .set("Accept", "application/vnd.github+json")
        .call()
        .map_err(|e| format!("リリース情報の取得に失敗: {e}"))?;

    let releases: Vec<GitHubRelease> = resp
        .into_json()
        .map_err(|e| format!("JSON解析失敗: {e}"))?;

    let current = format!("v{}", CURRENT_VERSION);
    let mut result = Vec::new();

    for release in &releases {
        // cart-converter.exe を含むリリースのみ
        if let Some(asset) = release.assets.iter().find(|a| a.name == "cart-converter.exe") {
            result.push(ReleaseInfo {
                version: release.tag_name.clone(),
                download_url: asset.browser_download_url.clone(),
                is_current: release.tag_name == current,
                release_notes: release.body.clone().unwrap_or_default(),
            });
        }
    }

    Ok(result)
}

/// 新しいexeをダウンロードして自己置き換え
/// Windowsでは実行中のexeをリネームできるので:
/// 1. 現在のexeを .old にリネーム
/// 2. 新しいexeをダウンロード
/// 3. アプリ再起動を促す
pub fn download_and_replace(release: &ReleaseInfo) -> Result<PathBuf, String> {
    let current_exe = std::env::current_exe()
        .map_err(|e| format!("現在のexeパスを取得できません: {e}"))?;

    let parent = current_exe.parent()
        .ok_or("exeの親ディレクトリが見つかりません")?;

    let old_exe = parent.join("cart-converter.old.exe");
    let new_exe = if current_exe.file_name().map(|n| n.to_str().unwrap_or("")) == Some("cart-converter.exe") {
        current_exe.clone()
    } else {
        // デバッグビルドなどの場合
        parent.join("cart-converter.exe")
    };

    // ダウンロード（一時ファイルに）
    let tmp_path = parent.join("cart-converter.new.exe");

    let resp = ureq::get(&release.download_url)
        .set("User-Agent", "cart-converter-updater")
        .call()
        .map_err(|e| format!("ダウンロード失敗: {e}"))?;

    let mut buf = Vec::new();
    resp.into_reader()
        .read_to_end(&mut buf)
        .map_err(|e| format!("読み取り失敗: {e}"))?;

    std::fs::write(&tmp_path, &buf)
        .map_err(|e| format!("一時ファイルの書き込み失敗: {e}"))?;

    // 現在のexeを .old にリネーム（存在する場合は先に削除）
    let _ = std::fs::remove_file(&old_exe);
    if current_exe.exists() && current_exe == new_exe {
        std::fs::rename(&current_exe, &old_exe)
            .map_err(|e| format!("現在のexeのリネーム失敗: {e}"))?;
    }

    // 新しいexeを配置
    std::fs::rename(&tmp_path, &new_exe)
        .map_err(|e| format!("新しいexeの配置失敗: {e}"))?;

    Ok(new_exe)
}

/// アプリを再起動（新しいexeで）
pub fn restart_app(exe_path: &PathBuf) {
    let _ = std::process::Command::new(exe_path).spawn();
    std::process::exit(0);
}
