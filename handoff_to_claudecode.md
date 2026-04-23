# 事例GP スライド自動生成 GAS — Claude Code 引き継ぎメモ

## 概要
Google スプレッドシート + Google Apps Script (GAS) で、フォーム入力から自動的に Google Slides の事例紹介スライドを生成するシステム。Gemini API でAIコメントも自動生成する。

---

## 現在の状況

### 完了していること
- GASスクリプト本体は完成（下記ファイル参照）
- `createTemplate()` 実行済み → CONFIGシートのB1にテンプレートスライドIDが保存されている
- `testRun()` or `runManually()` を実行 → スライド生成自体は完了している

### 未解決の問題
**「スライドが見つかりません」エラー**
- スライド生成は完了しているが、生成されたスライドにアクセスできない
- 考えられる原因：
  1. `OUTPUT_FOLDER_ID` のフォルダへのアクセス権限がない → ルートに保存されているかも
  2. スプレッドシートG列（SLIDE_URL）に書き込まれたURLが壊れている
  3. スライドは生成されているがGoogleドライブの別の場所に保存されている

### 確認すべきこと
1. スプレッドシートのH列（ステータス）の表示内容
2. CONFIGシートのB1セルにIDが入っているか
3. GASの実行ログ（Apps Script エディタ → 実行数 → 最新の実行ログ）
4. Googleドライブのルートフォルダに `事例GP_` で始まるファイルがないか

---

## ファイル構成

### GASスクリプト（最新版）
`jirei_gp_gas.js` に完全版あり。主な関数：

| 関数名 | 役割 |
|--------|------|
| `createTemplate()` | テンプレートスライドを新規作成。初回のみ実行。 |
| `onFormSubmit(e)` | フォーム送信トリガー。メイン処理。 |
| `generateAIComment(data)` | Gemini API でAIコメント生成 |
| `createSlide(data, aiComment)` | テンプレートをコピーしてテキスト置換 |
| `sendNotification(...)` | 完成通知メール送信 |
| `runManually()` | 最新行で手動実行 |
| `testRun()` | ダミーデータでテスト実行 |

### 設定値（CONFIG）
```javascript
const CONFIG = {
  GEMINI_API_KEY   : 'AIzaSyCWQlTXU93MXYwSU3oXojiA4LjrY2Oo5PY',
  OUTPUT_FOLDER_ID : '13cmi42diyueRgDRYfE04LWT2IAZ8nr8e',
  NOTIFY_EMAIL     : 'jnagai0423@gmail.com',
  TEMPLATE_SHEET   : 'CONFIG',
};
```

### スプレッドシート列構成
| 列 | 内容 |
|----|------|
| A | タイムスタンプ |
| B | 顧客名 |
| C | 事例概要 |
| D | 業種 |
| E | 事例内容詳細 |
| F | スコア |
| G | 生成スライドURL（自動書き込み） |
| H | ステータス（自動書き込み） |

---

## これまでの修正履歴

1. `createTemplate()` に `slide.getPlaceholders().forEach(p => p.remove())` を追加
   → Googleスライドのデフォルト「クリックしてタイトルを追加」枠を削除するため

2. `createSlide()` の置換処理を高速化
   - 変更前：シェイプ単位でループ → 重くてタイムアウト気味
   - 変更後：`pres.replaceAllText(key, value)` で一括置換

3. `pres.saveAndClose()` を削除
   - 追加したことでぐるぐる回って完了しなくなっていた

---

## 次にやること（未解決問題の対処）

```javascript
// createSlide() のフォルダ取得部分にデバッグログを追加して原因特定
let folder;
try {
  folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  Logger.log('フォルダ取得成功: ' + folder.getName());
} catch (e) {
  Logger.log('フォルダ取得失敗: ' + e.message);
  folder = DriveApp.getRootFolder();
  Logger.log('ルートフォルダに保存します');
}

// コピー後にURLをログ出力して確認
const copy = DriveApp.getFileById(templateId).makeCopy(fileName, folder);
Logger.log('生成ファイルID: ' + copy.getId());
Logger.log('生成URL: https://docs.google.com/presentation/d/' + copy.getId() + '/edit');
```

上記を追加して `testRun()` を実行 → ログでURLを直接確認する。
