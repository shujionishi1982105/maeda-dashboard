# PWA化（ホーム画面への追加） Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `app.py`（Streamlitダッシュボード）を、スマホのホーム画面に追加したときにクリニックのアイコン・アプリ名でスタンドアロン表示されるようにする。

**Architecture:** Streamlit公式の静的ファイル配信機能（`enableStaticServing`）で`static/`配下にWeb App Manifestとアイコン画像を配置し、`st.html(body, unsafe_allow_javascript=True)`（Streamlit 1.52.0〜の公式機能）で`<link rel="manifest">`等のタグを`document.head`に直接注入する。Service Workerは、Streamlitの静的配信のスコープ制約により実質機能しないため今回は導入しない（詳細はスペック参照）。

**Tech Stack:** Streamlit >=1.52.0（`st.html(unsafe_allow_javascript=True)`, `enableStaticServing`）, Pillow（アイコン画像生成）。

**参照スペック:** `docs/superpowers/specs/2026-08-19-pwa-home-screen-design.md`

**このプロジェクトに自動テストは存在しない。** 各タスクの検証は (a) `python3 -c "import ast; ast.parse(open('app.py').read())"` によるシンタックスチェック、(b) 必要に応じた使い捨ての検証スクリプト、(c) `streamlit.testing.v1.AppTest`によるE2E確認、(d) `streamlit run app.py`実行中の`curl`での静的配信確認、の組み合わせで行う。**実機（スマホ）での最終確認は自動化できないため、最後のタスクで手順を明記し、人手での確認を依頼する。**

---

### Task 1: アイコン画像の生成

**Files:**
- Create: `static/icon-192.png`
- Create: `static/icon-512.png`
- Create: `static/apple-touch-icon.png`

- [ ] **Step 1: 生成スクリプトを書く**

`/tmp/generate_pwa_icons.py` を作成:

```python
from PIL import Image
import os

SRC = "/home/shujionishi/.claude/image-cache/40b1f4fe-847a-4128-81d8-d14fb87f115d/1.png"
OUT_DIR = "static"

os.makedirs(OUT_DIR, exist_ok=True)

im = Image.open(SRC).convert("RGBA")
w, h = im.size
side = max(w, h)

# 正方形に透過パディング
canvas = Image.new("RGBA", (side, side), (255, 255, 255, 0))
canvas.paste(im, ((side - w) // 2, (side - h) // 2), im)

canvas.resize((192, 192), Image.LANCZOS).save(os.path.join(OUT_DIR, "icon-192.png"))
canvas.resize((512, 512), Image.LANCZOS).save(os.path.join(OUT_DIR, "icon-512.png"))

# apple-touch-iconは透過非推奨（iOSが透過部を黒く塗りつぶすため）、白背景に合成
apple_bg = Image.new("RGBA", (side, side), (255, 255, 255, 255))
apple_bg.paste(canvas, (0, 0), canvas)
apple_bg.convert("RGB").resize((180, 180), Image.LANCZOS).save(os.path.join(OUT_DIR, "apple-touch-icon.png"))

print("icons generated")
```

- [ ] **Step 2: 実行する**

Run: `python3 /tmp/generate_pwa_icons.py`（リポジトリのルートディレクトリで実行すること。`static/`が作成され、3つのPNGが出力される）
Expected: `icons generated` が出力される

- [ ] **Step 3: 生成結果を検証する**

`/tmp/verify_pwa_icons.py` を作成して実行:

```python
from PIL import Image

icon192 = Image.open("static/icon-192.png")
assert icon192.size == (192, 192), icon192.size

icon512 = Image.open("static/icon-512.png")
assert icon512.size == (512, 512), icon512.size

apple_icon = Image.open("static/apple-touch-icon.png")
assert apple_icon.size == (180, 180), apple_icon.size
assert apple_icon.mode == "RGB", apple_icon.mode  # 透過なし（iOS向け）

print("OK")
```

Run: `python3 /tmp/verify_pwa_icons.py`
Expected: `OK` が出力される

- [ ] **Step 4: 使い捨てスクリプトを削除しコミット**

```bash
rm /tmp/generate_pwa_icons.py /tmp/verify_pwa_icons.py
git add static/icon-192.png static/icon-512.png static/apple-touch-icon.png
git commit -m "PWA用アイコン画像（192px/512px/apple-touch-icon）を追加"
```

---

### Task 2: 静的アセット・設定ファイルの作成

**Files:**
- Create: `static/manifest.json`
- Create: `.streamlit/config.toml`
- Modify: `requirements.txt`

- [ ] **Step 1: `static/manifest.json` を作成する**

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

- [ ] **Step 2: `.streamlit/config.toml` を作成する**

```toml
[server]
enableStaticServing = true
```

- [ ] **Step 3: `requirements.txt` のstreamlitバージョンを引き上げる**

現在の内容（`requirements.txt`）:
```
streamlit>=1.30.0
pandas>=2.0.0
plotly>=5.18.0
```

`streamlit>=1.30.0` の行だけを次のように変更する:
```
streamlit>=1.52.0
```

（`unsafe_allow_javascript`パラメータがStreamlit 1.52.0で追加されたため。他の2行は変更しない）

- [ ] **Step 4: 検証する**

```bash
python3 -c "import json; d = json.load(open('static/manifest.json')); assert d['display'] == 'standalone'; assert len(d['icons']) == 2; print('manifest OK')"
python3 -c "import tomllib; d = tomllib.load(open('.streamlit/config.toml', 'rb')); assert d['server']['enableStaticServing'] is True; print('config OK')"
grep -q "streamlit>=1.52.0" requirements.txt && echo "requirements OK"
```

Expected: `manifest OK` / `config OK` / `requirements OK` がそれぞれ出力される（`tomllib`はPython 3.11以降の標準ライブラリ。このマシンのPythonバージョンは3.12なので利用可能）

- [ ] **Step 5: コミット**

```bash
git add static/manifest.json .streamlit/config.toml requirements.txt
git commit -m "PWA用manifest.json・静的配信設定・streamlitバージョン引き上げを追加"
```

---

### Task 3: `app.py` にPWAヘッダー注入処理を追加

**Files:**
- Modify: `app.py`（`st.set_page_config(...)` の直後、ログイン機能セクションより前）

- [ ] **Step 1: 現在のコードを確認する**

`app.py` の冒頭は次の通り:

```python
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import glob
import re
import os
import unicodedata

# キャッシュをリセット
st.cache_data.clear()

st.set_page_config(page_title="まえだ耳鼻咽喉科 経営分析", layout="wide")

# ==========================================
# 🔒 ログイン機能の設定
# ==========================================
```

- [ ] **Step 2: `st.set_page_config(...)` と `# 🔒 ログイン機能の設定` セクションの間にPWA注入処理を追加する**

```python
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import glob
import re
import os
import unicodedata

# キャッシュをリセット
st.cache_data.clear()

st.set_page_config(page_title="まえだ耳鼻咽喉科 経営分析", layout="wide")

# ==========================================
# 📱 PWA対応（ホーム画面への追加）
# ==========================================
def inject_pwa_head():
    st.html(
        """
        <script>
        (function() {
            if (document.querySelector('link[rel="manifest"]')) {
                return;
            }

            function addLink(rel, href) {
                var link = document.createElement('link');
                link.rel = rel;
                link.href = href;
                document.head.appendChild(link);
            }

            function addMeta(name, content) {
                var meta = document.createElement('meta');
                meta.name = name;
                meta.content = content;
                document.head.appendChild(meta);
            }

            addLink('manifest', 'app/static/manifest.json');
            addLink('apple-touch-icon', 'app/static/apple-touch-icon.png');
            addMeta('apple-mobile-web-app-capable', 'yes');
            addMeta('apple-mobile-web-app-status-bar-style', 'black-translucent');
            addMeta('apple-mobile-web-app-title', '経営分析ダッシュボード');
            addMeta('theme-color', '#2C3E50');
        })();
        </script>
        """,
        unsafe_allow_javascript=True,
    )

inject_pwa_head()

# ==========================================
# 🔒 ログイン機能の設定
# ==========================================
```

- [ ] **Step 3: シンタックスチェック**

Run: `python3 -c "import ast; ast.parse(open('app.py').read())"`
Expected: エラーなく終了

- [ ] **Step 4: 注入内容をgrepで確認する**

```bash
grep -c "addLink('manifest', 'app/static/manifest.json')" app.py
grep -c "addLink('apple-touch-icon', 'app/static/apple-touch-icon.png')" app.py
grep -c "unsafe_allow_javascript=True" app.py
grep -c 'querySelector(.link\[rel="manifest"\]' app.py
```

Expected: いずれも1以上（重複防止ガードと各タグ生成コードが存在すること）

- [ ] **Step 5: コミット**

```bash
git add app.py
git commit -m "PWAヘッダー注入処理（inject_pwa_head）をapp.pyに追加"
```

---

### Task 4: ローカル動作確認（静的配信・AppTest）

**Files:**
- Modify: なし（確認のみ）

- [ ] **Step 1: 依存関係を確認する**

このマシンには既に `streamlit`, `pandas`, `plotly`, `jinja2`, `Pillow` がインストール済み（前回のスマホ対応レイアウト作業時にセットアップ済み）。もし無ければ:

```bash
pip install --break-system-packages --quiet streamlit plotly jinja2 pillow
python3 -c "import streamlit; print(streamlit.__version__)"
```

Expected: `streamlit.__version__` が `1.52.0` 以上であること

- [ ] **Step 2: Streamlitをバックグラウンドで起動する**

リポジトリのルートディレクトリで:

```bash
nohup python3 -m streamlit run app.py --server.headless true --server.port 8766 --server.address 127.0.0.1 > /tmp/streamlit_pwa_smoke.log 2>&1 &
sleep 8
```

- [ ] **Step 3: 静的ファイルが配信されることを確認する**

```bash
curl -s -o /dev/null -w "manifest.json: %{http_code}\n" http://127.0.0.1:8766/app/static/manifest.json
curl -s -o /dev/null -w "icon-192.png: %{http_code}\n" http://127.0.0.1:8766/app/static/icon-192.png
curl -s -o /dev/null -w "icon-512.png: %{http_code}\n" http://127.0.0.1:8766/app/static/icon-512.png
curl -s -o /dev/null -w "apple-touch-icon.png: %{http_code}\n" http://127.0.0.1:8766/app/static/apple-touch-icon.png
```

Expected: すべて `200`

- [ ] **Step 4: サーバーログにエラーが出ていないことを確認する**

```bash
grep -i -E "error|traceback|exception" /tmp/streamlit_pwa_smoke.log || echo "NO ERRORS"
```

Expected: `NO ERRORS`

- [ ] **Step 5: AppTestでログイン〜全ページ遷移を通しで確認する**

`AppTest.from_file()`はスクリプト自身が置かれたディレクトリからの相対パスを解決するため、`/tmp`に検証スクリプトを置くと`app.py`への相対参照が解決できない。そのため`app.py`は絶対パスで指定する（リポジトリのルートディレクトリの絶対パスに置き換えること）。

`/tmp/verify_pwa_apptest.py` を作成:

```python
from streamlit.testing.v1 import AppTest

at = AppTest.from_file("/tmp/claude-1000/-home-shujionishi/39910c5c-937e-40e2-a576-dc003c792c67/scratchpad/maeda-dashboard/app.py", default_timeout=60)
at.run()
assert not at.exception, f"起動時に例外: {list(at.exception)}"

at.text_input[0].input("admin").run()
at.text_input[1].input("maeda2026").run()
at.button(key="FormSubmitter:login_form-ログイン").click().run()
assert not at.exception, f"ログイン後に例外: {list(at.exception)}"

at.radio(key="view_mode_radio").set_value("📱 スマホ表示").run()
assert not at.exception, f"スマホ表示切替後に例外: {list(at.exception)}"

pages = ["レセプト分析", "外来収入金額推移分析", "受付患者数（初再診別）推移分析", "年齢別構成比分析", "診療行為一覧分析", "検査一覧分析", "医師別診療実績（月別）", "AI総合経営アドバイス"]
for p in pages:
    sb = next(s for s in at.selectbox if s.label == "分析モードを選択")
    sb.set_value(p).run()
    assert not at.exception, f"{p} で例外: {list(at.exception)}"

print("ALL OK")
```

Run（リポジトリのルートディレクトリで）: `python3 /tmp/verify_pwa_apptest.py`
Expected: `ALL OK` が出力される（例外が発生した場合はAssertionErrorで即座に分かる）

- [ ] **Step 6: Streamlitサーバーを停止し、使い捨てファイルを削除する**

```bash
pkill -f "streamlit run app.py --server.headless true --server.port 8766" || true
sleep 1
rm -f /tmp/streamlit_pwa_smoke.log /tmp/verify_pwa_apptest.py
```

- [ ] **Step 7: 変更がなければコミット不要**（このタスクはコード変更を伴わない確認のみのため、コミットは発生しない）

---

### Task 5: 総合確認・ブランチ統合・実機確認手順の案内

**Files:**
- Modify: なし
- Sync: `/mnt/c/Users/大西修司/Downloads/app.py`（ローカル作業コピー）

- [ ] **Step 1: git logで一連のコミットを確認する**

Run: `git log --oneline main..HEAD`
Expected: Task 1・2・3の3件のコミットが表示される

- [ ] **Step 2: ローカル作業コピー（Downloads）に最新の `app.py` を同期する**

```bash
cp app.py "/mnt/c/Users/大西修司/Downloads/app.py"
diff -q app.py "/mnt/c/Users/大西修司/Downloads/app.py" && echo "sync OK"
```

- [ ] **Step 3: `superpowers:finishing-a-development-branch` スキルの手順に従い、ブランチの完了処理（マージ／PR作成／保持／破棄）をユーザーに確認する**

- [ ] **Step 4: 実機確認の手順をユーザーに案内する**

以下の内容をユーザーに伝える（このステップはコマンド実行ではなく案内文の提示）:

1. GitHubへのpush後、Streamlit Community Cloudでの再デプロイ完了を待つ
2. スマホのブラウザで本番URLを開く
3. **iPhoneの場合**：Safariで開き、共有ボタン→「ホーム画面に追加」。追加されたアイコンが正しいイラストになっているか、ホーム画面のアプリ名が「経営分析ダッシュボード」になっているかを確認し、アイコンをタップしてSafariのURLバーが表示されない（スタンドアロン表示になる）ことを確認する
4. **Androidの場合**：Chromeで開き、メニュー（︙）→「ホーム画面に追加」または「アプリをインストール」。同様にアイコン・アプリ名・起動時の表示を確認する
5. 期待通りに動かない場合（アイコンが違う、URLバーが表示されたままなど）は、その症状を教えてもらい追加調査する
