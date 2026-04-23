// 事例GP スライド自動生成 GAS（スプレッドシート紐付け）
// 設定エリア（初回はここを確認）
const CONFIG = {
  GEMINI_API_KEY   : 'AIzaSyCWQlTXU93MXYwSU3oXojiA4LjrY2Oo5PY',   // Gemini APIキー
  OUTPUT_FOLDER_ID : '13cmi42diyueRgDRYfE04LWT2IAZ8nr8e',   // 完成スライドの保存先フォルダID
  NOTIFY_EMAIL     : 'jnagai0423@gmail.com',       // 完成通知メール
  TEMPLATE_SHEET   : 'CONFIG',                       // テンプレートIDを保存するシート名
  ENABLE_UI_ALERT  : false,                          // ぐるぐる回避のため、通常は false 推奨
};

// スプレッドシート列番号（1始まり）
const COL = {
  CLIENT     : 2,  // B: 顧客名
  SUMMARY    : 3,  // C: 事例をざっくり
  INDUSTRY   : 4,  // D: 業種
  DETAIL     : 5,  // E: 事例の内容詳細
  SCORE      : 6,  // F: スコア
  SLIDE_URL  : 7,  // G: 生成スライドURL（自動書き込み）
  STATUS     : 8,  // H: ステータス（自動書き込み）
};


// 初回のみ実行: テンプレートスライドを作成
function createTemplate() {
  // テンプレートを新規作成
  const pres = SlidesApp.create('【テンプレート】事例GP_スライド');
  const slide = pres.getSlides()[0];

  // デフォルトのプレースホルダーを削除
  slide.getPlaceholders().forEach(placeholder => placeholder.remove());

  // レイアウト定数（単位: pt、16:9）
  const W = 720, H = 405;

  // 顧客名
  const clientBox = slide.insertTextBox('{{CLIENT_NAME}}', 20, 10, 400, 40);
  clientBox.getText().getTextStyle().setFontSize(24).setBold(true);

  // 業種
  const industryBox = slide.insertTextBox('業種：{{INDUSTRY}}', 20, 55, 300, 28);
  industryBox.getText().getTextStyle().setFontSize(13);

  // 事例概要
  const summaryBox = slide.insertTextBox('{{SUMMARY}}', 20, 90, W - 40, 50);
  summaryBox.getText().getTextStyle().setFontSize(18).setBold(true);

  // 事例内容ラベル
  const detailLabel = slide.insertTextBox('■ 事例内容', 20, 150, 200, 24);
  detailLabel.getText().getTextStyle().setFontSize(12).setBold(true);

  // 事例内容詳細
  const detailBox = slide.insertTextBox('{{DETAIL}}', 20, 178, 460, 110);
  detailBox.getText().getTextStyle().setFontSize(11);

  // スコアラベル
  const scoreLabelBox = slide.insertTextBox('スコア', 510, 150, 180, 24);
  scoreLabelBox.getText().getTextStyle().setFontSize(12).setBold(true);
  scoreLabelBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // スコア数値
  const scoreBox = slide.insertTextBox('{{SCORE}}点', 510, 178, 180, 70);
  scoreBox.getText().getTextStyle().setFontSize(40).setBold(true);
  scoreBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // AIコメントラベル
  const aiLabel = slide.insertTextBox('AI分析コメント', 20, 300, 200, 24);
  aiLabel.getText().getTextStyle().setFontSize(11).setBold(true);

  // AIコメント本文
  const aiBox = slide.insertTextBox('{{AI_COMMENT}}', 20, 324, W - 40, 40);
  aiBox.getText().getTextStyle().setFontSize(10);

  // ⑩ フッター（日付）
  const footerBox = slide.insertTextBox('{{DATE}}　Cloud Circus 事例GP', 20, H - 30, 400, 24);
  footerBox.getText().getTextStyle().setFontSize(9);

  // 作成直後はMyドライブ直下になるため、指定フォルダへ移動
  try {
    const outputFolder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    const templateFile = DriveApp.getFileById(pres.getId());
    outputFolder.addFile(templateFile);
    DriveApp.getRootFolder().removeFile(templateFile);
    Logger.log('テンプレートを保存先フォルダへ移動: ' + outputFolder.getName());
  } catch (e) {
    Logger.log('テンプレート移動に失敗（Myドライブに残ります）: ' + e.message);
  }

  // CONFIGシートにテンプレートIDを保存
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG.TEMPLATE_SHEET);
    configSheet.getRange('A1').setValue('TEMPLATE_SLIDE_ID');
  }
  configSheet.getRange('B1').setValue(pres.getId());

  const url = `https://docs.google.com/presentation/d/${pres.getId()}/edit`;
  Logger.log('テンプレート作成完了: ' + url);
  safeAlert(
    'テンプレート作成完了',
    `テンプレートスライドを作成しました。\n\n▼ 確認・編集はこちら\n${url}\n\nデザインはこのスライドを直接編集してカスタマイズできます。`
  );
}


// フォーム送信トリガー用（「フォーム送信時」に設定）
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();

  // ステータス更新（処理中）
  sheet.getRange(lastRow, COL.STATUS).setValue('処理中...');

  try {
    const row = sheet.getRange(lastRow, 1, 1, 6).getValues()[0];
    const data = {
      clientName : String(row[COL.CLIENT - 1] || '').trim(),
      summary    : String(row[COL.SUMMARY - 1] || '').trim(),
      industry   : String(row[COL.INDUSTRY - 1] || '').trim(),
      detail     : String(row[COL.DETAIL - 1] || '').trim(),
      score      : String(row[COL.SCORE - 1] || '').trim(),
    };

    if (!data.clientName) throw new Error('顧客名が空です');

    // Gemini API でAIコメント生成
    const aiComment = generateAIComment(data);

    // スライド生成
    const slideUrl = createSlide(data, aiComment);

    // スプレッドシートに結果書き込み
    sheet.getRange(lastRow, COL.SLIDE_URL).setValue(slideUrl);
    sheet.getRange(lastRow, COL.STATUS).setValue('完了');

    // メール通知
    sendNotification(data.clientName, slideUrl, aiComment);

    Logger.log(`[完了] ${data.clientName} → ${slideUrl}`);

  } catch (err) {
    sheet.getRange(lastRow, COL.STATUS).setValue('エラー: ' + err.message);
    Logger.log('[エラー] ' + err.toString());
    GmailApp.sendEmail(CONFIG.NOTIFY_EMAIL, '[エラー] 事例スライド生成失敗', err.toString());
  }
}


// Gemini APIでAIコメント生成
function generateAIComment(data) {
  if (!CONFIG.GEMINI_API_KEY || CONFIG.GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
    return '（APIキー未設定のためAIコメントなし）';
  }

  const prompt = `
あなたはBtoBデジタルマーケティングの専門家です。
以下の事例情報をもとに、イベント発表スライド用の印象的な一言コメントを80文字以内で生成してください。
数字・成果・ポジティブな変化を強調し、聴衆の共感を引く文章にしてください。
出力はコメント本文のみ（前置き・説明不要）。

顧客名: ${data.clientName}
業種: ${data.industry}
事例概要: ${data.summary}
詳細: ${data.detail}
スコア: ${data.score}
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { maxOutputTokens: 150, temperature: 0.75 }
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    if (json.candidates?.[0]?.content?.parts?.[0]?.text) {
      return json.candidates[0].content.parts[0].text.trim();
    }
    Logger.log('Gemini応答異常: ' + res.getContentText());
    return '（AIコメント生成に失敗しました）';
  } catch (err) {
    Logger.log('Gemini呼び出しエラー: ' + err);
    return '（AIコメント生成に失敗しました）';
  }
}


// スライド生成（テンプレートコピー + 置換）
function createSlide(data, aiComment) {
  // CONFIGシートからテンプレートID取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.TEMPLATE_SHEET);
  if (!configSheet) throw new Error('CONFIGシートが見つかりません。先に「①テンプレートを作成」を実行してください。');

  const templateId = String(configSheet.getRange('B1').getValue() || '').trim();
  if (!templateId) throw new Error('テンプレートIDが未設定です。先に「①テンプレートを作成」を実行してください。');

  // テンプレートをコピー
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  const fileName = `事例GP_${data.clientName}_${date}`;

  let folder;
  try {
    folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
    Logger.log('フォルダ取得成功: ' + folder.getName());
  } catch (e) {
    Logger.log('フォルダ取得失敗: ' + e.message);
    throw new Error(
      '保存先フォルダにアクセスできません。OUTPUT_FOLDER_ID と権限を確認してください。' +
      ' folderId=' + CONFIG.OUTPUT_FOLDER_ID
    );
  }

  const copy = DriveApp.getFileById(templateId).makeCopy(fileName, folder);
  Logger.log('生成ファイルID: ' + copy.getId());
  Logger.log('生成URL: https://docs.google.com/presentation/d/' + copy.getId() + '/edit');
  const pres = SlidesApp.openById(copy.getId());

  // プレースホルダー置換マップ
  const replacements = {
    '{{CLIENT_NAME}}' : data.clientName,
    '{{SUMMARY}}'     : data.summary,
    '{{INDUSTRY}}'    : data.industry,
    '{{DETAIL}}'      : data.detail,
    '{{SCORE}}'       : data.score,
    '{{AI_COMMENT}}'  : aiComment,
    '{{DATE}}'        : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日'),
  };

  // プレゼンテーション全体で一括置換（シェイプ単位より大幅に高速）
  Object.keys(replacements).forEach(key => {
    pres.replaceAllText(key, replacements[key]);
  });

  return `https://docs.google.com/presentation/d/${copy.getId()}/edit`;
}


// 完成通知メール
function sendNotification(clientName, url, aiComment) {
  if (!CONFIG.NOTIFY_EMAIL || CONFIG.NOTIFY_EMAIL === 'your-email@example.com') return;

  const subject = `【完成】${clientName} の事例スライドが生成されました`;
  const body = `
${clientName} の事例スライドが自動生成されました。

▼ スライドを開く
${url}

── AIコメント ──
${aiComment}

このメールは自動送信です。
  `.trim();

  GmailApp.sendEmail(CONFIG.NOTIFY_EMAIL, subject, body);
}


// スプレッドシートを開いた時のカスタムメニュー
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('事例GP')
    .addItem('① テンプレートを作成（初回のみ）', 'createTemplate')
    .addSeparator()
    .addItem('② 最新行でスライドを手動生成', 'runManually')
    .addToUi();
}

// UIアラートを安全に表示（表示不可環境ではログへ）
function safeAlert(title, message) {
  if (!CONFIG.ENABLE_UI_ALERT) {
    Logger.log(`safeAlert: 無効化中\n${title}\n${message}`);
    return;
  }

  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('safeAlert: UI表示をスキップしました: ' + e.message);
    Logger.log(`${title}\n${message}`);
  }
}


// 手動実行（最新行データ）
function runManually() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('データがありません。フォームから回答を入力してください。');
    return;
  }

  const row = sheet.getRange(lastRow, 1, 1, 6).getValues()[0];
  const clientName = row[COL.CLIENT - 1];

  const result = ui.alert(
    '手動生成',
    `最新行（行${lastRow}）のデータでスライドを生成します。\n\n顧客名: ${clientName}\n\n実行しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    onFormSubmit(null);
    const url = sheet.getRange(lastRow, COL.SLIDE_URL).getValue();
    safeAlert('完了', `スライドが生成されました。\n\n${url}`);
  }
}


