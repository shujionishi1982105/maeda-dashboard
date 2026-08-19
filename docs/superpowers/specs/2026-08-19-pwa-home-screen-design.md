# PWA化（ホーム画面への追加）設計

- 日付: 2026-08-19
- 対象: `app.py`（まえだ耳鼻咽喉科 経営分析ダッシュボード, Streamlit, GitHub: shujionishi1982105/maeda-dashboard, Streamlit Community Cloud上で公開）
- 対象コミット: 6caff11（スマホ対応レイアウト完了時点）
- 前提スペック: [2026-08-19-mobile-responsive-layout-design.md](2026-08-19-mobile-responsive-layout-design.md)

## 背景・目的

スマホ対応レイアウトの完了に続き、スマホのホーム画面にアイコンを追加して「アプリ風」に開けるようにする。ユーザーからのヒアリングにより、目的は「ホーム画面に追加してアプリ風に開きたい」ことであり、オフラインでの閲覧は不要と確認済み。

## スコープ

**含む**
- ホーム画面への追加時に使うアイコン（192×192px・512×512px・Apple用180×180px）
- Web App Manifest（`manifest.json`）による、アプリ名・アイコン・スタンドアロン表示設定
- Android/Chromeの「ホーム画面に追加」導線が正しく機能するための最小限のService Worker（キャッシュ処理なし、登録のみ）
- iOS Safari向けの`apple-touch-icon`・`apple-mobile-web-app-capable`等のメタタグ

**含まない（対象外）**
- オフラインでのデータ閲覧・キャッシュ（Service Workerはキャッシュしない）
- JavaScriptによる画面幅・デバイスの自動判定（[スマホ対応レイアウト設計](2026-08-19-mobile-responsive-layout-design.md)の方針を踏襲し、今回も導入しない）
- 既存のログイン機能・データ読み込みロジックの変更

## アイコン素材

ユーザーから提供された画像（192×170px、RGBA、透過背景のイラスト）を使用する。正方形にパディングした上で、以下のサイズを生成する：
- `icon-192.png`（192×192px、Web App Manifest用）
- `icon-512.png`（512×512px、Web App Manifest用）
- `apple-touch-icon.png`（180×180px、iOS Safari用）

元画像が192×170pxと小さいため、512px版は拡大時に多少のぼやけが生じる可能性がある。実害が大きい場合は、後日より高解像度の画像に差し替え可能な構成にする。

## 技術アーキテクチャ

### 1. 静的ファイル配信（Streamlit公式機能）

`.streamlit/config.toml` を新規作成し、以下を設定する：

```toml
[server]
enableStaticServing = true
```

これにより、プロジェクトルートの `static/` フォルダ内のファイルが `app/static/{ファイル名}` というURLパスで配信される（例: `https://<app-url>/app/static/manifest.json`）。これはStreamlit公式のドキュメント化された機能であり、Streamlitパッケージ自体を書き換える必要がない。

配置するファイル：
- `static/manifest.json`
- `static/icon-192.png`
- `static/icon-512.png`
- `static/apple-touch-icon.png`
- `static/sw.js`（Service Worker）

### 2. `<head>`へのタグ注入（JavaScriptインジェクション）

Streamlitは`st.markdown`で`<head>`タグを直接操作できないため、`st.components.v1.html`でJavaScriptを埋め込み、`window.parent.document.head`に対して以下のタグを注入する：

```html
<link rel="manifest" href="app/static/manifest.json">
<link rel="apple-touch-icon" href="app/static/apple-touch-icon.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="経営分析ダッシュボード">
<meta name="theme-color" content="#2C3E50">
```

加えて同じスクリプト内でService Workerを登録する：

```javascript
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('app/static/sw.js');
}
```

この注入処理は `app.py` 内に `inject_pwa_head()` のような小さな関数としてまとめ、スクリプトの先頭付近（`st.set_page_config` の直後、ログイン画面より前）で1回呼び出す。ログイン前後どちらの画面でも同じタグが適用される。

`st.components.v1.html`はiframe内でHTMLを描画する公式APIだが、`window.parent.document`への操作自体はStreamlit非公式のコミュニティ手法である。同一オリジンのiframeとして描画されるため通常は問題なく動作するが、将来のブラウザ仕様変更やStreamlitのセキュリティ強化により動作しなくなる可能性はゼロではない。

### 3. 最小限のService Worker

`static/sw.js` は以下のような、キャッシュを一切行わない最小構成にする：

```javascript
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));
```

`fetch`イベントのリスナーは登録しない（＝すべての通信は素通りでネットワークに任せる）。目的はAndroid/Chromeの「ホーム画面に追加」導線を有効にすることのみで、オフラインキャッシュは行わない。

## Web App Manifest の内容

```json
{
  "name": "経営分析ダッシュボード",
  "short_name": "経営分析ダッシュボード",
  "start_url": "/",
  "scope": "/",
  "display": "standalone",
  "background_color": "#FFFFFF",
  "theme_color": "#2C3E50",
  "icons": [
    { "src": "app/static/icon-192.png", "sizes": "192x192", "type": "image/png" },
    { "src": "app/static/icon-512.png", "sizes": "512x512", "type": "image/png" }
  ]
}
```

`theme_color`は既存アプリのヘッダー文字色（`#2C3E50`）に合わせる。

## 影響を受けないもの

- ログイン処理、CSV読み込み・集計処理、既存の分析ページすべて
- スマホ対応レイアウト（表示モード切替・ナビ・マトリクス表）の既存ロジック

## テスト・検証方法

このサンドボックス環境ではヘッドレスブラウザ（Playwright）を用意しようとしたが、依存ライブラリ不足とsudo権限がないため実機能検証はできなかった。そのため、以下の段階で検証する：

1. **静的チェック**：`ast.parse`によるPython構文チェック、`json.loads`による`manifest.json`の妥当性チェック
2. **配信確認**：`streamlit run app.py`をローカル起動し、`curl`で`app/static/manifest.json`・各アイコン・`sw.js`が200 OKで取得できることを確認
3. **注入内容の確認**：`st.components.v1.html`に渡すHTML文字列に、想定した`<link>`・`<meta>`タグと登録スクリプトが正しく含まれていることを文字列レベルで確認
4. **実機確認（人手・必須）**：GitHubへのpush・Streamlit Community Cloudへのデプロイ後、実際のスマホ（iOS Safari・Android Chromeの両方が望ましい）で「ホーム画面に追加」を行い、アイコン・アプリ名が正しく表示され、アイコンをタップした際にブラウザのURLバーなしで開く（スタンドアロン表示になる）ことを確認する

自動テストで検証しきれない最終確認（実機での見え方・動作）が必須である点は、前回のスマホ対応レイアウト作業（AppTestによる自動E2E検証がほぼ完結した）と異なり、本作業固有の制約として明記しておく。
