// 事例GP スライド自動生成 GAS（スプレッドシート紐付け）
// 設定エリア（初回はここを確認）
const CONFIG = {
  GEMINI_API_KEY   : 'AIzaSyCWQlTXU93MXYwSU3oXojiA4LjrY2Oo5PY',   // Gemini APIキー
  OUTPUT_FOLDER_ID : '13cmi42diyueRgDRYfE04LWT2IAZ8nr8e',   // 完成スライドの保存先フォルダID
  NOTIFY_EMAIL     : 'jnagai0423@gmail.com',       // 完成通知メール
  /** コピー元: ドライブの「jireiGp_SlideTemplate」プレゼン（URL 全体でも可） */
  TEMPLATE_SLIDE_ID: '1vukTwLSPjNdbFrr89SfP7kGVlsPh50gwoAwlwdxcyh8',
  ENABLE_UI_ALERT  : false,                          // ぐるぐる回避のため、通常は false 推奨
  /** true のとき、フォーム送信ごとにヘッダーと読み取った各列の値を Logger に出す（列マッピング確認用） */
  DEBUG_SHEET_HEADERS: false,
  /**
   * 回答を読むシート名（完全一致）。空なら名前が「フォームの回答」で始まる先頭タブを使う。
   * 複数のフォーム回答シートがあるときは、ここに例: 「フォームの回答 4」を指定。
   */
  FORM_RESPONSE_SHEET_NAME: '',
};

// スプレッドシート列番号（1始まり）
const COL = {
  SLIDE_URL_HEADER : '生成スライドURL',
  STATUS_HEADER    : 'ステータス',
};

/** 「事例グランプリ」見出しと K列成果本文で揃える赤 */
const ACCENT_BROWN = '#8B5A2B';


/**
 * コピー元スライドを「一から」作る（スプレッドシートのメニューから実行）。
 * 既にドライブに jireiGp_SlideTemplate などがあり CONFIG.TEMPLATE_SLIDE_ID で指している場合は不要（別ファイルが増えるだけ）。
 * SlidesApp.create の引数は、作成される Google スライドのファイル名（タイトル）になる。
 */
function createTemplate() {
  const fileTitle = '事例GP_コピー元スライド（メニューから新規作成）';
  const pres = SlidesApp.create(fileTitle);
  const slide = pres.getSlides()[0];

  // デフォルトのプレースホルダーを削除
  slide.getPlaceholders().forEach(placeholder => placeholder.remove());

  // レイアウト定数（単位: pt、16:9）
  const W = 720, H = 405;

  // タイトル補助
  const gpTitleBox = slide.insertTextBox('事例グランプリ', 20, 10, 220, 18);
  const gpTitleStyle = gpTitleBox.getText().getTextStyle();
  gpTitleStyle.setFontSize(10).setBold(true);
  safeSetTextColor(gpTitleStyle, ACCENT_BROWN);

  // タイトル（顧客名）
  const clientBox = slide.insertTextBox('{{CLIENT_NAME}}', 20, 24, 460, 36);
  clientBox.getText().getTextStyle().setFontSize(21).setBold(true);

  // 発表者
  const personBox = slide.insertTextBox('担当者：{{PERSON_NAME}}', 505, 42, 195, 20);
  personBox.getText().getTextStyle().setFontSize(11);
  personBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

  // 上段メタ情報
  const industryBox = slide.insertTextBox('・業種：{{INDUSTRY}}', 20, 62, W - 40, 16);
  industryBox.getText().getTextStyle().setFontSize(9);
  const planBox = slide.insertTextBox('・運用プラン：{{PLAN}}', 20, 77, W - 40, 16);
  planBox.getText().getTextStyle().setFontSize(9);
  const productBox = slide.insertTextBox('・導入製品：{{PRODUCTS}}', 20, 92, W - 40, 16);
  productBox.getText().getTextStyle().setFontSize(9);
  const siteUrlBox = slide.insertTextBox('・サイトURL：{{SITE_URL}}', 20, 107, W - 40, 16);
  siteUrlBox.getText().getTextStyle().setFontSize(9);
  const genreBox = slide.insertTextBox('・成果ジャンル：{{GENRE}}', 20, 122, W - 40, 16);
  genreBox.getText().getTextStyle().setFontSize(9);

  // 成果（本文は K列、KPIは見出し右横に表示）
  const featureLabel = slide.insertTextBox('成果を一言で', 20, 146, 110, 20);
  featureLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const kpiArrowBox = slide.insertTextBox('{{KPI_ARROW}}', 94, 146, 260, 20);
  kpiArrowBox.getText().getTextStyle().setFontSize(12).setBold(true);
  const featureBox = slide.insertTextBox('{{FEATURE}}', 20, 166, W - 40, 42);
  const featureStyle = featureBox.getText().getTextStyle();
  featureStyle.setFontSize(28).setBold(true);
  safeSetTextColor(featureStyle, ACCENT_BROWN);
  featureBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);

  // 成果内容（大きく表示）
  const detailLabel = slide.insertTextBox('事例の内容', 20, 212, 220, 20);
  detailLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const detailBox = slide.insertTextBox('{{DETAIL}}', 20, 232, W - 40, 58);
  detailBox.getText().getTextStyle().setFontSize(14);

  // AIコメント（右側に寄せる）
  const aiLabel = slide.insertTextBox('※', 360, 334, 20, 20);
  aiLabel.getText().getTextStyle().setFontSize(11).setBold(true);

  // AIコメント本文
  const aiBox = slide.insertTextBox('{{AI_COMMENT}}', 380, 334, 320, 34);
  aiBox.getText().getTextStyle().setFontSize(9);

  // ⑩ フッター（日付）
  const footerBox = slide.insertTextBox('{{FOOTER_DATE}}', 500, H - 24, 200, 18);
  footerBox.getText().getTextStyle().setFontSize(9);
  footerBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

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

  const url = `https://docs.google.com/presentation/d/${pres.getId()}/edit`;
  Logger.log('テンプレート作成完了: ' + url);
  safeAlert(
    'コピー元スライドの作成が完了',
    `コピー元となるスライドを新規作成しました。\n\n▼ 確認・編集はこちら\n${url}\n\nデザインはこのスライドを直接編集してカスタマイズできます。\n\nフォーム連携のコピー元に使うには、スクリプト先頭の CONFIG.TEMPLATE_SLIDE_ID を次の ID に更新してください。\n${pres.getId()}`
  );
}


/** フォーム回答シートを取得（先頭シートが CONFIG などでもずれないようにする） */
function getFormResponseSheet(ss) {
  const explicit = String(CONFIG.FORM_RESPONSE_SHEET_NAME || '').trim();
  if (explicit) {
    const named = ss.getSheetByName(explicit);
    if (named) return named;
    Logger.log('CONFIG.FORM_RESPONSE_SHEET_NAME が見つかりません: ' + explicit);
  }
  const formSheets = ss.getSheets().filter(s => /^フォームの回答/.test(s.getName()));
  if (formSheets.length) return formSheets[0];
  return ss.getSheets()[0];
}

// フォーム送信トリガー用（「フォーム送信時」に設定）
function onFormSubmit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getFormResponseSheet(ss);
  const lastRow = sheet.getLastRow();
  const outCol = ensureOutputColumns(sheet);

  // ステータス更新（処理中）
  sheet.getRange(lastRow, outCol.statusCol).setValue('処理中...');

  try {
    const data = getRowDataByHeader(sheet, lastRow);

    if (!data.clientName) throw new Error('顧客名が空です');

    // Gemini API でAIコメント生成
    const aiComment = generateAIComment(data);

    // スライド生成
    const slideUrl = createSlide(data, aiComment);

    // スプレッドシートに結果書き込み
    sheet.getRange(lastRow, outCol.slideUrlCol).setValue(slideUrl);
    sheet.getRange(lastRow, outCol.statusCol).setValue('完了');

    // メール通知
    sendNotification(data.clientName, slideUrl, aiComment);

    Logger.log(`[完了] ${data.clientName} → ${slideUrl}`);

  } catch (err) {
    sheet.getRange(lastRow, outCol.statusCol).setValue('エラー: ' + err.message);
    Logger.log('[エラー] ' + err.toString());
    GmailApp.sendEmail(CONFIG.NOTIFY_EMAIL, '[エラー] 事例スライド生成失敗', err.toString());
  }
}


// Gemini APIでAIコメント生成
function generateAIComment(data) {
  if (!CONFIG.GEMINI_API_KEY || CONFIG.GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
    return buildFallbackComment(data);
  }

  const prompt = `
あなたはBtoBデジタルマーケティングの専門家です。
以下の事例情報をもとに、イベント発表スライド用の印象的な一言コメントを80文字以内で生成してください。
数字・成果・ポジティブな変化を強調し、聴衆の共感を引く文章にしてください。
出力はコメント本文のみ（前置き・説明不要）。

顧客名: ${data.clientName}
業種: ${data.industry}
発表者: ${data.personName}
運用プラン: ${data.plan}
導入製品: ${data.products}
サイトURL: ${data.siteUrl}
成果ジャンル: ${data.genre}
成果（一言）: ${buildFeatureForSlide(data) || '（未入力）'}
${buildKpiArrowText(data) ? `成果KPI: ${buildKpiArrowText(data)}` : ''}
成果内容: ${data.detail}
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
    const status = res.getResponseCode();
    const body = res.getContentText();
    if (status < 200 || status >= 300) {
      Logger.log(`Gemini HTTPエラー: status=${status}, body=${body}`);
      return buildFallbackComment(data);
    }

    const json = JSON.parse(body);
    const text = json.candidates?.[0]?.content?.parts?.[0]?.text;
    if (text && String(text).trim()) {
      return String(text).trim();
    }
    Logger.log('Gemini応答異常: ' + body);
    return buildFallbackComment(data);
  } catch (err) {
    Logger.log('Gemini呼び出しエラー: ' + err);
    return buildFallbackComment(data);
  }
}

/** K列（成果の一言）本文のみを返す */
function buildFeatureForSlide(data) {
  return String(data.feature || '').trim();
}

/** L列・M列を見出し横に表示する文字列へ整形（例: （月間5件→月間8件）） */
function buildKpiArrowText(data) {
  const l = String(data.metric30Day || '').trim();
  const m = String(data.actualMetric || '').trim();
  if (!l && !m) return '';
  const arrowPart = l && m ? `${l}→${m}` : l ? `${l}→` : `→${m}`;
  return `（${arrowPart}）`;
}

// スライド生成（テンプレートコピー + 置換）
function createSlide(data, aiComment) {
  const templateId = getTemplateSlideId();

  // テンプレートをコピー
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  const fileName = `事例GP_${data.personName || '担当者未入力'}_${date}`;

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
  const aiCommentForSlide = trimForSlide(aiComment, 95);
  const featureForSlide = trimForSlide(buildFeatureForSlide(data), 200);
  const kpiArrowForSlide = trimForSlide(buildKpiArrowText(data), 80);
  const replacements = {
    '{{CLIENT_NAME}}' : data.clientName,
    '{{PERSON_NAME}}' : data.personName,
    '{{INDUSTRY}}'    : data.industry,
    '{{PLAN}}'        : data.plan,
    '{{PRODUCTS}}'    : data.products,
    '{{SITE_URL}}'    : data.siteUrl,
    '{{GENRE}}'       : data.genre,
    '{{FEATURE}}'     : featureForSlide,
    '{{KPI_ARROW}}'   : kpiArrowForSlide,
    '{{DETAIL}}'      : data.detail,
    '{{AI_COMMENT}}'  : aiCommentForSlide,
    '{{FOOTER_DATE}}' : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月') + '：事例グランプリ',
  };

  // プレゼンテーション全体で一括置換（シェイプ単位より大幅に高速）
  const corePlaceholderKeys = [
    '{{CLIENT_NAME}}',
    '{{PERSON_NAME}}',
    '{{PLAN}}',
    '{{GENRE}}',
    '{{FEATURE}}',
    '{{DETAIL}}',
  ];
  let replacedCoreCount = 0;
  Object.keys(replacements).forEach(key => {
    const n = pres.replaceAllText(key, replacements[key]);
    if (corePlaceholderKeys.includes(key)) {
      replacedCoreCount += Number(n || 0);
    }
  });

  applyFeatureHighlightTextColor(pres, featureForSlide);
  alignFeatureTextBox(pres, featureForSlide);
  insertKpiArrowNearFeatureLabel(pres, kpiArrowForSlide);

  // テンプレに {{...}} が無い（デザインのみ）場合は置換だけでは何も出ないため、テキストボックスで上書き描画する
  if (replacedCoreCount === 0) {
    Logger.log(
      'コア用プレースホルダーがテンプレ内に見つかりませんでした。テキストボックスで内容を描画します。' +
        ' デザインに埋め込む場合はスライドに {{CLIENT_NAME}} などを配置してください。'
    );
    renderContentFallback(pres, data, aiComment);
  }

  insertSitePreviewImage(pres, data.siteUrl);
  ensureThreeSameSlides(pres);

  return `https://docs.google.com/presentation/d/${copy.getId()}/edit`;
}

/** 1ページ目の見た目を基準に、同一構成の3ページへ揃える */
function ensureThreeSameSlides(presentation) {
  const slides = presentation.getSlides();
  if (!slides.length) return;

  const first = slides[0];

  // いったん1ページ目だけ残す（テンプレが複数ページでも同一構成3ページに統一）
  for (let i = slides.length - 1; i >= 1; i -= 1) {
    slides[i].remove();
  }

  // 1ページ目を複製して計3ページにする
  presentation.appendSlide(first);
  presentation.appendSlide(first);
}

function getTemplateSlideId() {
  const id = String(CONFIG.TEMPLATE_SLIDE_ID || '').trim();
  if (!id) {
    throw new Error('テンプレートID未設定です。CONFIG.TEMPLATE_SLIDE_ID を設定するか、メニューからテンプレート作成後に ID を反映してください。');
  }
  return extractSlideId(id);
}

function extractSlideId(value) {
  const text = String(value || '').trim();
  const m = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];
  return text;
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
    .addItem('① コピー元スライドを新規作成（手元に無いときのみ）', 'createTemplate')
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getFormResponseSheet(ss);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('データがありません。フォームから回答を入力してください。');
    return;
  }

  const data = getRowDataByHeader(sheet, lastRow);
  const clientName = data.clientName || '（未入力）';

  const result = ui.alert(
    '手動生成',
    `最新行（行${lastRow}）のデータでスライドを生成します。\n\n顧客名: ${clientName}\n\n実行しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    onFormSubmit();
    const outCol = ensureOutputColumns(sheet);
    const url = sheet.getRange(lastRow, outCol.slideUrlCol).getValue();
    safeAlert('完了', `スライドが生成されました。\n\n${url}`);
  }
}


// ヘッダー名ベースで行データを取得（フォーム項目変更に強くする）
function getRowDataByHeader(sheet, rowNumber) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const valueRow = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];

  const normalizedHeaders = headerRow.map(h => normalizeHeader(h));
  /** 複数列に部分一致する場合は、最も長いキーに一致した列を採用（短い「url」だけの誤マッチを減らす） */
  const getByKey = (keys, excludeKeys) => {
    const normalizedKeys = (keys || []).map(k => normalizeHeader(k)).filter(Boolean);
    const normalizedExcludes = (excludeKeys || []).map(k => normalizeHeader(k)).filter(Boolean);

    let bestIdx = -1;
    let bestScore = 0;

    normalizedHeaders.forEach((h, i) => {
      if (!h) return;
      if (normalizedExcludes.some(ex => ex && h.includes(ex))) return;

      let score = 0;
      for (const k of normalizedKeys) {
        if (h.includes(k) && k.length > score) score = k.length;
      }
      if (score > bestScore) {
        bestScore = score;
        bestIdx = i;
      }
    });

    if (bestIdx === -1 || bestScore === 0) return '';
    return String(valueRow[bestIdx] || '').trim();
  };

  const personName = getByKey([
    '自分の名前を入力してください',
    '自分の名前を入力',
    '自分の名前',
    '担当者名',
    '氏名',
  ]);
  const clientName = getByKey([
    '募集企業名および製品サービス名',
    '募集企業名',
    '製品サービス名',
    '顧客企業名を正式名称で入力',
    '顧客企業名',
    '顧客名',
  ]);
  const industry = getByKey([
    '顧客企業の業種を以下から選択してください',
    '顧客企業の業種',
    '業種',
  ]);
  const products = getByKey(
    [
      '導入済みのcloudcircus製品があれば選択してください',
      '導入済みのcloudcircus製品',
      '導入済のcloudcircus製品',
      'cloudcircus製品',
      'cloudcircus',
      '導入済',
    ],
    ['顧客企業名', '顧客名', '募集企業名', '生成スライドurl']
  );
  const plan = getByKey([
    '期待するコンサルティングの成果',
    '運用中のコンサルティングプラン',
    'コンサルティングの成果',
    'コンサルティングプラン',
    '運用中',
    'プラン',
    'コンサルティング',
  ]);
  const siteUrl = getByKey(
    [
      '顧客企業の対象url',
      '対象url',
      'サイトurlを入力',
      'サイトurl',
      'webサイトurl',
      'ホームページurl',
      'url',
    ],
    ['生成スライドurl', 'メールアドレス']
  );
  const genre = getByKey([
    '成果事例のジャンルを以下から選択してください',
    '成果事例のジャンル',
    'ジャンル',
  ]);
  const feature = getByKey([
    '成果事例の成果を一言でアピール',
    '成果事例の成果を一言で',
    '成果事例の成果',
    '成果事例の特徴を一言で表すと',
    '成果事例の特徴',
    '15文字',
  ]);
  const detail = getByKey([
    '成果事例の内容をできるだけ詳細に記述してください',
    '成果事例の内容',
    '300文字',
    '詳細',
  ]);
  const metric30DayByHeader = getByKey([
    'kpi改善数',
    'KPI改善数',
    '30日内数値',
    '30日内',
  ]);
  const actualMetricByHeader = getByKey([
    'kpi実績数値',
    'KPI実績数値',
    '実績数値',
    '実績数値あれば',
  ]);
  // L/M 列を明示フォールバック（列追加や文言変更で見出しマッチしない時の保険）
  const metric30Day =
    metric30DayByHeader || (sheet.getLastColumn() >= 12 ? String(valueRow[11] || '').trim() : '');
  const actualMetric =
    actualMetricByHeader || (sheet.getLastColumn() >= 13 ? String(valueRow[12] || '').trim() : '');

  const data = {
    personName,
    clientName,
    industry,
    products,
    siteUrl,
    genre,
    feature,
    detail,
    plan,
    metric30Day,
    actualMetric,
  };

  if (CONFIG.DEBUG_SHEET_HEADERS) {
    Logger.log('[DEBUG] シート: ' + sheet.getName() + ' 行: ' + rowNumber);
    Logger.log('[DEBUG] ヘッダー: ' + JSON.stringify(headerRow));
    Logger.log('[DEBUG] 読み取り: ' + JSON.stringify(data));
  }

  return data;
}

/** テンプレに {{CLIENT_NAME}} 等が無いとき、1枚目にフォーム内容をテキストで描画する */
function renderContentFallback(pres, data, aiComment) {
  const slide = pres.getSlides()[0];
  const W = 720;
  const H = 405;

  const gpTitleBox = slide.insertTextBox('事例グランプリ', 20, 10, 220, 18);
  const gpTitleStyle = gpTitleBox.getText().getTextStyle();
  gpTitleStyle.setFontSize(10).setBold(true);
  safeSetTextColor(gpTitleStyle, ACCENT_BROWN);

  const clientBox = slide.insertTextBox(data.clientName || '（顧客名未入力）', 20, 24, 460, 36);
  clientBox.getText().getTextStyle().setFontSize(21).setBold(true);

  const personBox = slide.insertTextBox(`担当者：${data.personName || '未入力'}`, 505, 42, 195, 20);
  personBox.getText().getTextStyle().setFontSize(11);
  personBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

  const industryBox = slide.insertTextBox(`・業種：${data.industry || '未入力'}`, 20, 62, W - 40, 16);
  industryBox.getText().getTextStyle().setFontSize(9);

  const planBox = slide.insertTextBox(`・期待する成果：${data.plan || '未入力'}`, 20, 77, W - 40, 16);
  planBox.getText().getTextStyle().setFontSize(9);

  const productBox = slide.insertTextBox(`・導入製品：${data.products || '未入力'}`, 20, 92, W - 40, 16);
  productBox.getText().getTextStyle().setFontSize(9);

  const siteUrlBox = slide.insertTextBox(`・サイトURL：${data.siteUrl || '未入力'}`, 20, 107, W - 40, 16);
  siteUrlBox.getText().getTextStyle().setFontSize(9);

  const genreBox = slide.insertTextBox(`・成果ジャンル：${data.genre || '未入力'}`, 20, 122, W - 40, 16);
  genreBox.getText().getTextStyle().setFontSize(9);

  const featureLabel = slide.insertTextBox('成果を一言で', 20, 146, 110, 20);
  featureLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const kpiArrowText = trimForSlide(buildKpiArrowText(data), 80);
  const kpiArrowBox = slide.insertTextBox(kpiArrowText, 94, 146, 260, 20);
  kpiArrowBox.getText().getTextStyle().setFontSize(12).setBold(true);
  const featureText =
    trimForSlide(buildFeatureForSlide(data), 200) || '（特徴未入力）';
  const featureBox = slide.insertTextBox(featureText, 20, 166, W - 40, 42);
  const featureBodyStyle = featureBox.getText().getTextStyle();
  featureBodyStyle.setFontSize(24).setBold(true);
  safeSetTextColor(featureBodyStyle, ACCENT_BROWN);
  featureBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);

  const detailLabel = slide.insertTextBox('事例の内容', 20, 212, 220, 20);
  detailLabel.getText().getTextStyle().setFontSize(12).setBold(true);
  const detailBox = slide.insertTextBox(data.detail || '（詳細未入力）', 20, 232, W - 40, 58);
  detailBox.getText().getTextStyle().setFontSize(14);

  const aiLabel = slide.insertTextBox('※', 360, 334, 20, 20);
  aiLabel.getText().getTextStyle().setFontSize(11).setBold(true);
  const aiBox = slide.insertTextBox(trimForSlide(aiComment || '（AIコメントなし）', 95), 380, 334, 320, 34);
  aiBox.getText().getTextStyle().setFontSize(9);

  const footerText = `${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月')}：事例グランプリ`;
  const footerBox = slide.insertTextBox(footerText, 500, H - 24, 200, 18);
  footerBox.getText().getTextStyle().setFontSize(9);
  footerBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
}

function buildFallbackComment(data) {
  const feature = buildFeatureForSlide(data) || '成果を創出';
  const kpi = buildKpiArrowText(data);
  const genre = data.genre || '成果領域';
  return `「${feature}${kpi}」を実現。${genre}で再現性のある運用成果が確認できる好事例です。`;
}

function trimForSlide(text, maxLength) {
  const str = String(text || '').trim();
  if (!str) return '';
  if (str.length <= maxLength) return str;
  return `${str.slice(0, maxLength - 1)}…`;
}

function safeSetTextColor(textStyle, colorHex) {
  try {
    textStyle.setForegroundColor(colorHex);
  } catch (e) {
    Logger.log(`文字色の設定をスキップ: ${e}`);
  }
}

/** 1枚目で本文が K列相当（featureForSlide）と一致するシェイプをこげ茶にする（{{FEATURE}} 置換後のテンプレ用） */
function applyFeatureHighlightTextColor(presentation, displayText) {
  const needle = String(displayText || '').trim();
  if (!needle) return;
  presentation.getSlides()[0].getPageElements().forEach(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
    const shape = pe.asShape();
    let full = '';
    try {
      full = String(shape.getText().asString() || '').trim();
    } catch (e) {
      return;
    }
    if (full !== needle) return;
    safeSetTextColor(shape.getText().getTextStyle(), ACCENT_BROWN);
    shape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  });
}

/** 1枚目で本文が K列相当（displayText）のシェイプを左寄せ・左位置に補正する */
function alignFeatureTextBox(presentation, displayText) {
  const needle = String(displayText || '').trim();
  if (!needle) return;
  presentation.getSlides()[0].getPageElements().forEach(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
    const shape = pe.asShape();
    let full = '';
    try {
      full = String(shape.getText().asString() || '').trim();
    } catch (e) {
      return;
    }
    if (full !== needle) return;
    // 見出し「成果を一言で」と同じ左マージンに合わせる
    pe.setLeft(20);
    pe.setWidth(680);
    shape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  });
}

/** 1枚目で「成果」見出しの右に KPI 文字列を黒で描画（本文には連結しない） */
function insertKpiArrowNearFeatureLabel(presentation, kpiText) {
  const text = String(kpiText || '').trim();
  if (!text) return;
  const slide = presentation.getSlides()[0];

  const hasSameText = slide.getPageElements().some(pe => {
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return false;
    try {
      return String(pe.asShape().getText().asString() || '').trim() === text;
    } catch (e) {
      return false;
    }
  });
  if (hasSameText) return;

  let inserted = false;
  slide.getPageElements().forEach(pe => {
    if (inserted) return;
    if (pe.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
    const shape = pe.asShape();
    let label = '';
    try {
      label = String(shape.getText().asString() || '').trim();
    } catch (e) {
      return;
    }
    if (label !== '成果を一言で' && label !== '成果') return;

    const left = pe.getLeft();
    const top = pe.getTop();
    const width = pe.getWidth();
    const h = pe.getHeight();
    const kpiBox = slide.insertTextBox(text, left + width + 6, top, 300, h + 2);
    kpiBox.getText().getTextStyle().setFontSize(12);
    kpiBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    inserted = true;
  });
}

function insertSitePreviewImage(pres, siteUrl) {
  const imageBlobs = fetchSitePreviewImageBlobs(siteUrl);
  if (!imageBlobs.length) return;

  try {
    const slide = pres.getSlides()[0];
    // 右側に最大3枚を重ねて配置（少しずつオフセット）
    const placements = [
      { x: 392, y: 64, w: 288, h: 211 },
      { x: 407, y: 78, w: 288, h: 211 },
      { x: 422, y: 92, w: 288, h: 211 },
    ];
    let insertedCount = 0;
    imageBlobs.slice(0, 12).forEach((blob, i) => {
      if (insertedCount >= 3) return;
      const p = placements[i] || placements[placements.length - 1];
      try {
        slide.insertImage(blob, p.x, p.y, p.w, p.h);
        insertedCount += 1;
      } catch (insertErr) {
        Logger.log('画像挿入をスキップ: ' + insertErr);
      }
    });
    // 挿入成功が足りない場合は、最初の画像を再利用して3枚に揃える
    while (insertedCount > 0 && insertedCount < 3) {
      const p = placements[insertedCount];
      slide.insertImage(imageBlobs[0], p.x, p.y, p.w, p.h);
      insertedCount += 1;
    }
    if (insertedCount < 3) {
      Logger.log(`画像挿入は${insertedCount}件でした（有効画像不足）`);
    }
  } catch (e) {
    Logger.log('サイト画像の挿入をスキップ: ' + e);
  }
}

function fetchSitePreviewImageBlobs(siteUrl) {
  const url = String(siteUrl || '').trim();
  if (!url) return [];
  if (!/^https?:\/\//i.test(url)) return [];

  try {
    const htmlRes = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GAS-bot/1.0)' }
    });

    if (htmlRes.getResponseCode() < 200 || htmlRes.getResponseCode() >= 300) {
      Logger.log(`サイトHTML取得失敗: status=${htmlRes.getResponseCode()} url=${url}`);
      return [];
    }

    const html = htmlRes.getContentText();
    const candidates = [
      ...pickMetaContents(html, 'property', 'og:image'),
      ...pickMetaContents(html, 'name', 'twitter:image'),
      ...pickMetaContents(html, 'property', 'og:image:url'),
      ...pickImgSrcs(html),
      ...pickImgSrcsetCandidates(html),
    ];

    if (!candidates.length) return [];

    const uniqueUrls = [];
    candidates.forEach(v => {
      const abs = toAbsoluteUrl(url, v);
      if (abs && !uniqueUrls.includes(abs)) uniqueUrls.push(abs);
    });

    const blobs = [];
    uniqueUrls.slice(0, 60).forEach(imageUrl => {
      try {
        const imageRes = UrlFetchApp.fetch(imageUrl, {
          method: 'get',
          muteHttpExceptions: true,
          followRedirects: true,
          headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GAS-bot/1.0)' }
        });
        if (imageRes.getResponseCode() < 200 || imageRes.getResponseCode() >= 300) return;
        const blob = imageRes.getBlob();
        const contentType = String(blob.getContentType() || '').toLowerCase();
        if (!contentType.startsWith('image/')) return;
        // Slides に挿入しやすい形式を優先（svg や ico は除外）
        if (!/image\/(png|jpeg|jpg|gif|webp|bmp)/.test(contentType)) return;
        if (blob.getBytes().length < 4000) return; // 小さすぎるアイコン画像を除外
        const dim = getImageDimensions(blob, contentType);
        if (!dim) return;
        // 縦長画像は除外（正方形・横長のみ採用）
        if (dim.width < dim.height) return;
        blobs.push(blob);
      } catch (fetchErr) {
        Logger.log(`画像取得をスキップ: ${imageUrl} err=${fetchErr}`);
      }
    });
    return blobs;
  } catch (e) {
    Logger.log('サイト画像取得エラー: ' + e);
    return [];
  }
}

/** ページ内の <img src="..."> を抽出（data URI は除外） */
function pickImgSrcs(html) {
  const pattern = /<img[^>]*\ssrc\s*=\s*["']([^"']+)["'][^>]*>/ig;
  const out = [];
  const str = String(html || '');
  let m;
  while ((m = pattern.exec(str)) !== null) {
    const v = String(m[1] || '').trim();
    if (!v) continue;
    if (/^data:/i.test(v)) continue;
    out.push(v);
  }
  return out;
}

function pickImgSrcsetCandidates(html) {
  const pattern = /<img[^>]*\ssrcset\s*=\s*["']([^"']+)["'][^>]*>/ig;
  const out = [];
  const str = String(html || '');
  let m;
  while ((m = pattern.exec(str)) !== null) {
    const srcset = String(m[1] || '');
    srcset.split(',').forEach(part => {
      const first = String(part || '').trim().split(/\s+/)[0];
      if (!first) return;
      if (/^data:/i.test(first)) return;
      out.push(first);
    });
  }
  return out;
}

function getImageDimensions(blob, contentType) {
  try {
    const bytes = blob.getBytes();
    if (!bytes || bytes.length < 24) return null;

    if (/image\/png/.test(contentType)) {
      return {
        width: readUInt32BE(bytes, 16),
        height: readUInt32BE(bytes, 20),
      };
    }
    if (/image\/gif/.test(contentType)) {
      return {
        width: readUInt16LE(bytes, 6),
        height: readUInt16LE(bytes, 8),
      };
    }
    if (/image\/bmp/.test(contentType)) {
      return {
        width: Math.abs(readInt32LE(bytes, 18)),
        height: Math.abs(readInt32LE(bytes, 22)),
      };
    }
    if (/image\/jpe?g/.test(contentType)) {
      return readJpegDimensions(bytes);
    }
    if (/image\/webp/.test(contentType)) {
      return readWebpDimensions(bytes);
    }
    return null;
  } catch (e) {
    Logger.log('画像サイズ取得失敗: ' + e);
    return null;
  }
}

function readUInt16LE(bytes, i) {
  return (bytes[i] & 0xff) | ((bytes[i + 1] & 0xff) << 8);
}

function readInt32LE(bytes, i) {
  const b0 = bytes[i] & 0xff;
  const b1 = (bytes[i + 1] & 0xff) << 8;
  const b2 = (bytes[i + 2] & 0xff) << 16;
  const b3 = (bytes[i + 3] & 0xff) << 24;
  return (b0 | b1 | b2 | b3);
}

function readUInt32BE(bytes, i) {
  return ((bytes[i] & 0xff) << 24) | ((bytes[i + 1] & 0xff) << 16) | ((bytes[i + 2] & 0xff) << 8) | (bytes[i + 3] & 0xff);
}

function readJpegDimensions(bytes) {
  let i = 2;
  while (i + 9 < bytes.length) {
    if ((bytes[i] & 0xff) !== 0xff) {
      i += 1;
      continue;
    }
    const marker = bytes[i + 1] & 0xff;
    const length = ((bytes[i + 2] & 0xff) << 8) | (bytes[i + 3] & 0xff);
    if (length < 2) return null;
    if (marker >= 0xc0 && marker <= 0xc3 && i + 8 < bytes.length) {
      const height = ((bytes[i + 5] & 0xff) << 8) | (bytes[i + 6] & 0xff);
      const width = ((bytes[i + 7] & 0xff) << 8) | (bytes[i + 8] & 0xff);
      return { width, height };
    }
    i += 2 + length;
  }
  return null;
}

function readWebpDimensions(bytes) {
  if (bytes.length < 30) return null;
  const riff = String.fromCharCode(bytes[0] & 0xff, bytes[1] & 0xff, bytes[2] & 0xff, bytes[3] & 0xff);
  const webp = String.fromCharCode(bytes[8] & 0xff, bytes[9] & 0xff, bytes[10] & 0xff, bytes[11] & 0xff);
  if (riff !== 'RIFF' || webp !== 'WEBP') return null;
  const chunk = String.fromCharCode(bytes[12] & 0xff, bytes[13] & 0xff, bytes[14] & 0xff, bytes[15] & 0xff);

  if (chunk === 'VP8X' && bytes.length >= 30) {
    const w = 1 + ((bytes[24] & 0xff) | ((bytes[25] & 0xff) << 8) | ((bytes[26] & 0xff) << 16));
    const h = 1 + ((bytes[27] & 0xff) | ((bytes[28] & 0xff) << 8) | ((bytes[29] & 0xff) << 16));
    return { width: w, height: h };
  }
  if (chunk === 'VP8L' && bytes.length >= 25) {
    const b0 = bytes[21] & 0xff;
    const b1 = bytes[22] & 0xff;
    const b2 = bytes[23] & 0xff;
    const b3 = bytes[24] & 0xff;
    const w = 1 + (b0 | ((b1 & 0x3f) << 8));
    const h = 1 + ((b1 >> 6) | (b2 << 2) | ((b3 & 0x0f) << 10));
    return { width: w, height: h };
  }
  if (chunk === 'VP8 ' && bytes.length >= 30) {
    const w = readUInt16LE(bytes, 26) & 0x3fff;
    const h = readUInt16LE(bytes, 28) & 0x3fff;
    return { width: w, height: h };
  }
  return null;
}

function pickMetaContents(html, attrName, attrValue) {
  const escaped = escapeRegex(attrValue);
  const pattern = new RegExp(
    `<meta[^>]*${attrName}\\s*=\\s*["']${escaped}["'][^>]*content\\s*=\\s*["']([^"']+)["'][^>]*>|` +
    `<meta[^>]*content\\s*=\\s*["']([^"']+)["'][^>]*${attrName}\\s*=\\s*["']${escaped}["'][^>]*>`,
    'ig'
  );
  const out = [];
  const str = String(html || '');
  let m;
  while ((m = pattern.exec(str)) !== null) {
    const v = (m[1] || m[2] || '').trim();
    if (v) out.push(v);
  }
  return out;
}

function toAbsoluteUrl(baseUrl, maybeRelative) {
  const raw = String(maybeRelative || '').trim();
  if (!raw) return '';
  if (/^https?:\/\//i.test(raw)) return raw;
  if (/^\/\//.test(raw)) {
    const scheme = /^https:\/\//i.test(baseUrl) ? 'https:' : 'http:';
    return `${scheme}${raw}`;
  }

  const hostMatch = String(baseUrl).match(/^(https?:\/\/[^\/?#]+)/i);
  if (!hostMatch) return raw;
  const origin = hostMatch[1];

  if (raw.startsWith('/')) return `${origin}${raw}`;

  const pathBase = String(baseUrl).replace(/[#?].*$/, '').replace(/\/[^/]*$/, '/');
  return `${pathBase}${raw}`;
}

function escapeRegex(str) {
  return String(str || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeHeader(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .replace(/[（）()！!：:・、,./]/g, '')
    .toLowerCase();
}

function ensureOutputColumns(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const headerRow = headerRange.getValues()[0];
  const normalizedHeaders = headerRow.map(h => normalizeHeader(h));

  const slideHeader = normalizeHeader(COL.SLIDE_URL_HEADER);
  const statusHeader = normalizeHeader(COL.STATUS_HEADER);

  let slideUrlCol = normalizedHeaders.findIndex(h => h === slideHeader) + 1;
  let statusCol = normalizedHeaders.findIndex(h => h === statusHeader) + 1;

  if (!slideUrlCol) {
    slideUrlCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, slideUrlCol).setValue(COL.SLIDE_URL_HEADER);
  }
  if (!statusCol) {
    statusCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, statusCol).setValue(COL.STATUS_HEADER);
  }

  return { slideUrlCol, statusCol };
}
