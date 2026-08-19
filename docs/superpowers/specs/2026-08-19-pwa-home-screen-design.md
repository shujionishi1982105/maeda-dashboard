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
- iOS Safari向けの`apple-touch-icon`・`apple-mobile-web-app-capable`等のメタタグ

**含まない（対象外）**
- Service Worker（下記「検討したが見送った項目」を参照）
- オフラインでのデータ閲覧・キャッシュ
- JavaScriptによる画面幅・デバイスの自動判定（[スマホ対応レイアウト設計](2026-08-19-mobile-responsive-layout-design.md)の方針を踏襲し、今回も導入しない）
- 既存のログイン機能・データ読み込みロジックの変更

## 検討したが見送った項目：Service Worker

当初はAndroid/Chromeの「ホーム画面に追加」導線を強化する目的で、キャッシュ処理を行わない最小限のService Workerを追加する予定だった。しかし技術調査の結果、以下の理由で見送った。

- Streamlitの静的ファイル配信（`enableStaticServing`）は`app/static/`配下にしかファイルを置けない
- ブラウザの仕様上、Service Workerは「配信されたディレクトリ配下」しか制御できない（既定の最大スコープ）。これを広げるには`Service-Worker-Allowed`レスポンスヘッダーが必要だが、Streamlitの静的ファイル配信はレスポンスヘッダーを制御する手段を提供していない
- そのため`app/static/sw.js`として配信しても、実際のダッシュボード本体（`/`配下）を制御できず、登録しても実質的に機能しない（無駄な登録エラーが発生する可能性もある）

iOS Safariの「ホーム画面に追加」はService Workerを必要としないため影響なし。Androidも`manifest.json`の`display: standalone`設定により、基本的な「ホーム画面に追加」自体は動作する見込みだが、一部端末・Chromeバージョンでは完全な単独アプリ的挙動（URLバー非表示等）が弱まる可能性がある。この点はユーザーに説明済みで、了承を得ている。将来的にAndroidでの挙動を強化したい場合は、`sw.js`をルート直下（`/`）で配信できる別のホスティング構成（リバースプロキシ等）が必要になる。

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

`unsafe_allow_javascript=True`の利用に`streamlit>=1.52.0`が必要なため、`requirements.txt`の`streamlit>=1.30.0`も`streamlit>=1.52.0`に引き上げる。

### 2. `<head>`へのタグ注入（`st.html`の公式JavaScript実行機能を使用）

Streamlit 1.52.0で追加された `st.html(body, unsafe_allow_javascript=True)` を使う。この関数は`st.components.v1.html`と異なりiframeを使わず、内容をページの実際のDOMに直接描画するため、埋め込んだ`<script>`は`window.parent`を経由せず`document`に直接アクセスできる（Streamlit公式ドキュメント・インストール済みパッケージのdocstringで確認済み：「``st.html`` content is **not** iframed... To execute JavaScript contained in your HTML, set ``unsafe_allow_javascript=True``」）。これにより以下のタグを`document.head`に注入する：

```html
<link rel="manifest" href="app/static/manifest.json">
<link rel="apple-touch-icon" href="app/static/apple-touch-icon.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="経営分析ダッシュボード">
<meta name="theme-color" content="#2C3E50">
```

スクリプトの先頭で「既に注入済みなら何もしない」ガード（`document.querySelector('link[rel="manifest"]')`の存在チェック）を入れる。Streamlitはユーザー操作のたびにスクリプト全体を再実行するため、ガードがないと再実行のたびに同じタグが重複して`<head>`に追加されてしまう。

この注入処理は `app.py` 内に `inject_pwa_head()` のような小さな関数としてまとめ、スクリプトの先頭付近（`st.set_page_config` の直後、ログイン画面より前）で1回呼び出す。ログイン前後どちらの画面でも同じタグが適用される。

`unsafe_allow_javascript=True`はStreamlit公式のドキュメント化された機能（1.52.0〜）であり、`st.components.v1.html`のiframe越しに`window.parent`を操作する非公式のコミュニティ手法より安定性が高い。ただし本番で使うには `requirements.txt` の `streamlit>=1.30.0` を `streamlit>=1.52.0` 以上に引き上げる必要がある。

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
2. **配信確認**：`streamlit run app.py`をローカル起動し、`curl`で`app/static/manifest.json`・各アイコンが200 OKで取得できることを確認
3. **注入内容の確認**：`st.html`に渡すHTML文字列に、想定した`<link>`・`<meta>`タグが正しく含まれていることを文字列レベルで確認。また`streamlit.testing.v1.AppTest`でログイン〜各ページ遷移を通しで実行し、`inject_pwa_head()`呼び出しを含めて例外が発生しないことを確認する
4. **実機確認（人手・必須）**：GitHubへのpush・Streamlit Community Cloudへのデプロイ後、実際のスマホ（iOS Safari・Android Chromeの両方が望ましい）で「ホーム画面に追加」を行い、アイコン・アプリ名が正しく表示され、アイコンをタップした際にブラウザのURLバーなしで開く（スタンドアロン表示になる）ことを確認する

自動テストで検証しきれない最終確認（実機での見え方・動作）が必須である点は、前回のスマホ対応レイアウト作業（AppTestによる自動E2E検証がほぼ完結した）と異なり、本作業固有の制約として明記しておく。
