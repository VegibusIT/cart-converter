use serde::{Deserialize, Serialize};
use std::path::PathBuf;

/// Google OAuth2 クライアント設定
#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct GoogleCredentials {
    pub client_id: String,
    pub client_secret: String,
}

/// 埋め込みクレデンシャルを実行時に組み立て
fn builtin_credentials() -> GoogleCredentials {
    let id_parts = [
        "248442146927-3geoi2c75bdb",
        "ujh9v0tfr9rrtlq4snaq.apps.",
        "googleusercontent.com",
    ];
    let secret_parts = [
        "GOCSPX--bZnDAncL",
        "zm4jVluoCxvHRAIyw2a",
    ];
    GoogleCredentials {
        client_id: id_parts.join(""),
        client_secret: secret_parts.join(""),
    }
}

impl Default for GoogleCredentials {
    fn default() -> Self {
        builtin_credentials()
    }
}

impl GoogleCredentials {
    pub fn is_configured(&self) -> bool {
        !self.client_id.is_empty() && !self.client_secret.is_empty()
    }
}

/// 保存済みトークン
#[derive(Serialize, Deserialize, Clone, Debug)]
pub struct TokenData {
    pub access_token: String,
    pub refresh_token: Option<String>,
    pub expires_at: Option<u64>,
}

/// Google認証の状態
#[derive(Debug)]
pub enum AuthState {
    /// 未認証
    NotAuthenticated,
    /// 認証中（ブラウザ待ち）
    WaitingForCallback,
    /// 認証済み
    Authenticated(TokenData),
    /// エラー
    Error(String),
}

const SCOPES: &[&str] = &[
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
];

const REDIRECT_URI: &str = "http://localhost:8085";
const AUTH_URL: &str = "https://accounts.google.com/o/oauth2/v2/auth";
const TOKEN_URL: &str = "https://oauth2.googleapis.com/token";

fn config_dir() -> PathBuf {
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

fn credentials_path() -> PathBuf {
    config_dir().join("google_credentials.json")
}

fn token_path() -> PathBuf {
    config_dir().join("google_token.json")
}

/// 保存済みクレデンシャルを読み込む（なければNone → デフォルト値を使用可能）
pub fn load_credentials() -> Option<GoogleCredentials> {
    let s = std::fs::read_to_string(credentials_path()).ok()?;
    serde_json::from_str(&s).ok()
}

/// クレデンシャルを保存
pub fn save_credentials(creds: &GoogleCredentials) {
    let dir = config_dir();
    let _ = std::fs::create_dir_all(&dir);
    if let Ok(json) = serde_json::to_string_pretty(creds) {
        let _ = std::fs::write(credentials_path(), json);
    }
}

/// 保存済みトークンを読み込む
pub fn load_token() -> Option<TokenData> {
    let s = std::fs::read_to_string(token_path()).ok()?;
    serde_json::from_str(&s).ok()
}

/// トークンを保存
pub fn save_token(token: &TokenData) {
    let dir = config_dir();
    let _ = std::fs::create_dir_all(&dir);
    if let Ok(json) = serde_json::to_string_pretty(token) {
        let _ = std::fs::write(token_path(), json);
    }
}

/// トークンを削除（ログアウト）
pub fn clear_token() {
    let _ = std::fs::remove_file(token_path());
}

/// OAuth2認証URLを生成
pub fn build_auth_url(creds: &GoogleCredentials) -> String {
    let scope = SCOPES.join(" ");
    format!(
        "{}?client_id={}&redirect_uri={}&response_type=code&scope={}&access_type=offline&prompt=consent",
        AUTH_URL,
        urlencoding::encode(&creds.client_id),
        urlencoding::encode(REDIRECT_URI),
        urlencoding::encode(&scope),
    )
}

/// ローカルサーバーで認証コールバックを受け取り、トークンを取得
/// ブロッキング呼び出し - 別スレッドで実行すること
pub fn wait_for_callback_and_exchange(
    creds: &GoogleCredentials,
    cancel: &std::sync::atomic::AtomicBool,
) -> Result<TokenData, String> {
    use std::io::{BufRead, Write};
    use std::net::TcpListener;
    use std::sync::atomic::Ordering;

    let listener = TcpListener::bind("127.0.0.1:8085")
        .map_err(|e| format!("ローカルサーバー起動失敗: {e}"))?;

    // ノンブロッキングにしてキャンセルチェックできるようにする
    listener
        .set_nonblocking(true)
        .map_err(|e| format!("ソケット設定失敗: {e}"))?;

    // 接続を待つ（キャンセル可能）
    let (mut stream, _) = loop {
        if cancel.load(Ordering::Relaxed) {
            return Err("キャンセルされました".into());
        }
        match listener.accept() {
            Ok(conn) => break conn,
            Err(ref e) if e.kind() == std::io::ErrorKind::WouldBlock => {
                std::thread::sleep(std::time::Duration::from_millis(200));
                continue;
            }
            Err(e) => return Err(format!("接続待ち失敗: {e}")),
        }
    };

    // 接続後はブロッキングモードに戻す
    stream
        .set_nonblocking(false)
        .map_err(|e| format!("ストリーム設定失敗: {e}"))?;

    let mut reader = std::io::BufReader::new(&stream);
    let mut request_line = String::new();
    reader
        .read_line(&mut request_line)
        .map_err(|e| format!("リクエスト読み取り失敗: {e}"))?;

    // レスポンスを返す
    let html = "<html><body><h2>認証完了</h2><p>このタブを閉じてアプリに戻ってください。</p></body></html>";
    let response = format!(
        "HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {}\r\n\r\n{}",
        html.len(),
        html
    );
    let _ = stream.write_all(response.as_bytes());
    let _ = stream.flush();

    // コードを抽出
    let code = extract_code_from_request(&request_line)
        .ok_or_else(|| "認証コードを取得できませんでした".to_string())?;

    // トークン交換
    exchange_code(creds, &code)
}

fn extract_code_from_request(request_line: &str) -> Option<String> {
    // "GET /callback?code=xxx&scope=... HTTP/1.1"
    let path = request_line.split_whitespace().nth(1)?;
    let query = path.split('?').nth(1)?;
    for param in query.split('&') {
        if let Some(value) = param.strip_prefix("code=") {
            return Some(urlencoding::decode(value).ok()?.into_owned());
        }
    }
    None
}

/// 認証コードをトークンに交換
fn exchange_code(creds: &GoogleCredentials, code: &str) -> Result<TokenData, String> {
    let body = format!(
        "code={}&client_id={}&client_secret={}&redirect_uri={}&grant_type=authorization_code",
        urlencoding::encode(code),
        urlencoding::encode(&creds.client_id),
        urlencoding::encode(&creds.client_secret),
        urlencoding::encode(REDIRECT_URI),
    );

    let response = ureq::post(TOKEN_URL)
        .set("Content-Type", "application/x-www-form-urlencoded")
        .send_string(&body)
        .map_err(|e| format!("トークン取得失敗: {e}"))?;

    let json: serde_json::Value = response
        .into_json()
        .map_err(|e| format!("レスポンス解析失敗: {e}"))?;

    let access_token = json["access_token"]
        .as_str()
        .ok_or("access_tokenがありません")?
        .to_string();

    let refresh_token = json["refresh_token"].as_str().map(|s| s.to_string());

    let expires_in = json["expires_in"].as_u64().unwrap_or(3600);
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_secs();

    let token = TokenData {
        access_token,
        refresh_token,
        expires_at: Some(now + expires_in),
    };

    save_token(&token);
    Ok(token)
}

/// リフレッシュトークンでアクセストークンを更新
pub fn refresh_access_token(creds: &GoogleCredentials, token: &TokenData) -> Result<TokenData, String> {
    let refresh_token = token
        .refresh_token
        .as_ref()
        .ok_or("リフレッシュトークンがありません")?;

    let body = format!(
        "client_id={}&client_secret={}&refresh_token={}&grant_type=refresh_token",
        urlencoding::encode(&creds.client_id),
        urlencoding::encode(&creds.client_secret),
        urlencoding::encode(refresh_token),
    );

    let response = ureq::post(TOKEN_URL)
        .set("Content-Type", "application/x-www-form-urlencoded")
        .send_string(&body)
        .map_err(|e| format!("トークン更新失敗: {e}"))?;

    let json: serde_json::Value = response
        .into_json()
        .map_err(|e| format!("レスポンス解析失敗: {e}"))?;

    let access_token = json["access_token"]
        .as_str()
        .ok_or("access_tokenがありません")?
        .to_string();

    let expires_in = json["expires_in"].as_u64().unwrap_or(3600);
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap()
        .as_secs();

    let new_token = TokenData {
        access_token,
        refresh_token: token.refresh_token.clone(),
        expires_at: Some(now + expires_in),
    };

    save_token(&new_token);
    Ok(new_token)
}

/// トークンが有効かどうかチェック（期限切れ300秒前にはfalse）
pub fn is_token_valid(token: &TokenData) -> bool {
    if let Some(expires_at) = token.expires_at {
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_secs();
        now + 300 < expires_at
    } else {
        false
    }
}

/// 有効なアクセストークンを取得（必要に応じてリフレッシュ）
pub fn get_valid_token(creds: &GoogleCredentials) -> Result<TokenData, String> {
    let token = load_token().ok_or("ログインが必要です")?;
    if is_token_valid(&token) {
        Ok(token)
    } else {
        refresh_access_token(creds, &token)
    }
}
