// ============================================================
//  口座情報登録フォーム — Google Apps Script (Code.gs)
// ============================================================
//
//  【セットアップ手順】
//
//  ① Google スプレッドシートを新規作成
//     - シート名を「口座情報」に変更
//     ※ 列の見出し（1行目）は、最初のデータ送信時に自動生成されます。
//
//  ② スプレッドシートの「拡張機能」→「Apps Script」を開く
//
//  ③ このファイルの内容をすべてコピーして Code.gs に貼り付け
//
//  ④ 下の定数を自分の環境に合わせて変更：
//     - SPREADSHEET_ID → スプレッドシートのID（URLの /d/ と /edit の間の文字列）
//     - DRIVE_FOLDER_ID → 通帳画像を保存するDriveフォルダのID
//
//  ⑤ デプロイ：
//     「デプロイ」→「新しいデプロイ」
//     → 種類：「ウェブアプリ」
//     → 実行ユーザー：「自分」
//     → アクセス：「全員」
//     → デプロイ → 表示されたURLをコピー
//
//  ⑥ HTMLファイルの GAS_URL にデプロイURLを貼り付け
//
// ============================================================

// ★★★ ここを自分の環境に合わせて変更 ★★★
const SPREADSHEET_ID  = '1sQ1nWtT0Gx2hEKNwffL5BDpeEhDmGouOS-Z9FS7y2e4';
const SHEET_NAME      = '口座情報';
const DRIVE_FOLDER_ID = '1QX44okhWslQe1uK9_jWiTvdkIP9RL_S5';

// ============================================================
//  doPost — フォームからのPOSTリクエストを処理
// ============================================================
function doPost(e) {
  try {
    // --- リクエストボディをパース ---
    const data = JSON.parse(e.postData.contents);

    // --- 通帳写真をDriveに保存 ---
    let photoUrl = '';
    if (data.photo_base64) {
      photoUrl = savePhotoToDrive(data.photo_base64, data.photo_filename, data.name_kanji);
    }

    // --- スプレッドシートに書き込み ---
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    // シートが空の場合、ヘッダーを自動生成
    if (sheet.getLastRow() === 0) {
      const headers = [
        'タイムスタンプ',
        '区分',
        '所属',
        '氏名（漢字）',
        '氏名（カタカナ）',
        '郵便番号',
        'ご住所',
        '番地・建物名',
        '電話番号',
        '記号',
        '番号',
        '口座名義（カタカナ）',
        '通帳写真URL'
      ];
      sheet.appendRow(headers);
    }

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    const row = [
      timestamp,              // A: タイムスタンプ
      data.category,          // B: 区分
      data.employee_branch,   // C: 所属
      data.name_kanji,        // D: 氏名（漢字）
      data.name_kana,         // E: 氏名（カタカナ）
      data.postal_code,       // F: 郵便番号
      data.address,           // G: ご住所
      data.address_detail,    // H: 番地・建物名
      data.phone || '',       // I: 電話番号
      data.symbol,            // J: 記号
      data.number,            // K: 番号
      data.account_holder,    // L: 口座名義（カタカナ）
      photoUrl                // M: 通帳写真URL
    ];

    sheet.appendRow(row);

    // --- 成功レスポンス ---
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: '登録完了' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    // --- エラーレスポンス ---
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
//  savePhotoToDrive — Base64画像をGoogle Driveに保存
// ============================================================
function savePhotoToDrive(base64DataUrl, originalFilename, userName) {
  // "data:image/jpeg;base64,/9j/4AAQ..." 形式から分離
  const parts     = base64DataUrl.split(',');
  const meta      = parts[0]; // "data:image/jpeg;base64"
  const raw       = parts[1]; // Base64データ本体

  // MIMEタイプを取得
  const mimeMatch = meta.match(/data:(.*?);/);
  const mimeType  = mimeMatch ? mimeMatch[1] : 'image/jpeg';

  // 拡張子を決定
  const extMap = {
    'image/jpeg': '.jpg',
    'image/png':  '.png',
    'image/gif':  '.gif',
    'image/webp': '.webp',
    'image/heic': '.heic',
    'image/heif': '.heif',
  };
  const ext = extMap[mimeType] || '.jpg';

  // ファイル名を生成（日時_名前_通帳）
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const safeName  = (userName || 'unknown').replace(/[\/\\:*?"<>|]/g, '_');
  const filename  = `${timestamp}_${safeName}_通帳${ext}`;

  // Base64をBlobに変換
  const decoded = Utilities.base64Decode(raw);
  const blob    = Utilities.newBlob(decoded, mimeType, filename);

  // 指定フォルダに保存
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const file   = folder.createFile(blob);

  // 社内で閲覧できるようにアクセス権を設定（必要に応じて変更）
  // file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

// ============================================================
//  doGet — GETアクセス時（テスト用）
// ============================================================
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      message: '口座情報登録APIは正常に稼働中です'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
