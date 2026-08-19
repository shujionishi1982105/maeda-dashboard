# スマホ対応レイアウト Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `app.py`（Streamlitダッシュボード）を、同一URL・同一コードのままPC表示とスマホ表示の両方で見やすくする。

**Architecture:** `st.query_params["view"]`（`"pc"`／`"mobile"`）で表示モードを保持し、ヘッダー直下のトグルで切り替える。`view == "mobile"` のときだけ、(1) 分析メニューをプルダウンに、(2) 列数の多い月別マトリクス表を直近6ヶ月＋全期間表示ボタンに切り替える。加えて画面幅ベースのCSSメディアクエリでフォント・余白を微調整する。

**Tech Stack:** Streamlit >=1.30 (`st.query_params` 新API), pandas, 既存の `app.py` 単一ファイル構成を踏襲。

**参照スペック:** `docs/superpowers/specs/2026-08-19-mobile-responsive-layout-design.md`

**このプロジェクトに自動テストは存在しない。** 各タスクの検証は (a) `python3 -c "import ast; ast.parse(open('app.py').read())"` によるシンタックスチェック、(b) 必要に応じた使い捨ての手動検証スクリプト、(c) `streamlit run app.py` での目視確認、の組み合わせで行う。

---

### Task 1: 直近Nヶ月に絞り込むヘルパー関数を追加

**Files:**
- Modify: `app.py`（「共通関数 (全ページで利用)」セクション、`get_act_summary_for_ai` 関数の直後・246行目付近）

- [ ] **Step 1: 動作を確認する使い捨て検証スクリプトを書く**

`/tmp/verify_limit_recent_months.py` を作成:

```python
import pandas as pd

def limit_recent_months(df, month_cols, n=6):
    recent_months = month_cols[-n:] if len(month_cols) > n else list(month_cols)
    other_cols = [c for c in df.columns if c not in month_cols]
    return df[recent_months + other_cols]

# 8ヶ月分のデータ + 年間合計列を持つダミーのマトリクス表
month_cols = [f"{i}月" for i in range(1, 9)]
df = pd.DataFrame(
    [[i * 10 for i in range(1, 9)] + [360]],
    columns=month_cols + ["年間合計"],
    index=["テスト行"],
)

result = limit_recent_months(df, month_cols, n=6)
assert list(result.columns) == ["3月", "4月", "5月", "6月", "7月", "8月", "年間合計"], result.columns
assert result.loc["テスト行", "年間合計"] == 360

# 月数がnより少ない場合はそのまま全部残る
short_cols = [f"{i}月" for i in range(1, 4)]
df_short = pd.DataFrame([[1, 2, 3, 6]], columns=short_cols + ["年間合計"], index=["テスト行"])
result_short = limit_recent_months(df_short, short_cols, n=6)
assert list(result_short.columns) == ["1月", "2月", "3月", "年間合計"], result_short.columns

print("OK")
```

- [ ] **Step 2: 実行して期待通り動くことを確認する**

Run: `python3 /tmp/verify_limit_recent_months.py`
Expected: `OK` が出力される（AssertionErrorが出ないこと）

- [ ] **Step 3: `app.py` に関数を追加する**

`app.py` の `get_act_summary_for_ai` 関数（214〜253行目）の直後、`# ==========================================` の次のセクション区切りの前に以下を追加する:

```python

def limit_recent_months(df, month_cols, n=6):
    """month_cols（時系列順の月名リスト）のうち直近n件だけ残した列構成のdfを返す。年間合計など月以外の列はそのまま残す。"""
    recent_months = month_cols[-n:] if len(month_cols) > n else list(month_cols)
    other_cols = [c for c in df.columns if c not in month_cols]
    return df[recent_months + other_cols]
```

- [ ] **Step 4: シンタックスチェック**

Run: `python3 -c "import ast; ast.parse(open('app.py').read())"`
Expected: エラーなく終了（何も出力されない）

- [ ] **Step 5: 使い捨て検証スクリプトを削除しコミット**

```bash
rm /tmp/verify_limit_recent_months.py
git add app.py
git commit -m "直近nヶ月に絞り込むlimit_recent_months関数を追加"
```

---

### Task 2: 表示モード（PC/スマホ）の状態管理とトグルUIを追加

**Files:**
- Modify: `app.py:282-286`（ヘッダーのログアウトボタンとナビゲーションメニューの間）

- [ ] **Step 1: 現在のコードを確認する**

`app.py` の該当箇所（ヘッダーブロックの直後）は次の通り:

```python
with header_cols[2]:
    st.write("<br>", unsafe_allow_html=True)
    if st.button("🚪 ログアウト", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

# ==========================================
# ナビゲーションメニュー
# ==========================================
```

- [ ] **Step 2: 表示モードのトグルを挿入する**

上記の `with header_cols[2]:` ブロックと `# ナビゲーションメニュー` セクションの間に、以下を追加する:

```python
with header_cols[2]:
    st.write("<br>", unsafe_allow_html=True)
    if st.button("🚪 ログアウト", use_container_width=True):
        st.session_state.logged_in = False
        st.rerun()

# ==========================================
# 表示モード切替（PC／スマホ）
# ==========================================
view = st.query_params.get("view", "pc")
view_choice = st.radio(
    "表示モード",
    ["🖥 PC表示", "📱 スマホ表示"],
    index=0 if view == "pc" else 1,
    horizontal=True,
    key="view_mode_radio",
    label_visibility="collapsed",
)
new_view = "pc" if view_choice == "🖥 PC表示" else "mobile"
if new_view != view:
    st.query_params["view"] = new_view
    st.rerun()

# ==========================================
# ナビゲーションメニュー
# ==========================================
```

- [ ] **Step 3: シンタックスチェック**

Run: `python3 -c "import ast; ast.parse(open('app.py').read())"`
Expected: エラーなく終了

- [ ] **Step 4: 手動確認**

Run: `streamlit run app.py`（WSL側の実行手順は前回セッション同様 `--break-system-packages` でpandas等インストール済みの環境を使う）

1. ログイン後、ヘッダー直下に「🖥 PC表示 / 📱 スマホ表示」のラジオボタンが表示されることを確認
2. 「📱 スマホ表示」をクリックし、ブラウザのURLに `?view=mobile` が付与されることを確認
3. ページをリロードしても `📱 スマホ表示` が選択されたままであることを確認（URLに状態が保持されている）
4. 「🖥 PC表示」に戻すと `?view=pc` になり、見た目が今まで通りであることを確認

- [ ] **Step 5: コミット**

```bash
git add app.py
git commit -m "PC/スマホ表示モードの切替UIとquery_params管理を追加"
```

---

### Task 3: 分析メニューのナビゲーションをスマホ表示時はプルダウンに変更

**Files:**
- Modify: `app.py:288-314`付近（Task 2適用後は行番号がずれるため、`# ナビゲーションメニュー` セクションを検索して特定する）

- [ ] **Step 1: 現在のコードを確認する**

```python
# ==========================================
# ナビゲーションメニュー
# ==========================================
pages = [
    "レセプト分析", 
    "外来収入金額推移分析", 
    "受付患者数（初再診別）推移分析",
    "年齢別構成比分析",
    "診療行為一覧分析",
    "検査一覧分析",
    "医師別診療実績（月別）",
    "AI総合経営アドバイス"
]

if 'current_page' not in st.session_state:
    st.session_state.current_page = pages[0]

st.write("### 🔍 分析メニュー")

for i in range(0, len(pages), 4):
    cols = st.columns(4)
    for j in range(4):
        if i + j < len(pages):
            page_name = pages[i + j]
            with cols[j]:
                btn_type = "primary" if st.session_state.current_page == page_name else "secondary"
                if st.button(page_name, use_container_width=True, key=f"nav_btn_{i+j}", type=btn_type):
                    st.session_state.current_page = page_name
                    st.rerun()

st.write("---")
```

- [ ] **Step 2: `st.write("### 🔍 分析メニュー")` 以降を、表示モードで分岐させる**

`st.write("### 🔍 分析メニュー")` から `st.write("---")` の直前までを、以下に置き換える:

```python
st.write("### 🔍 分析メニュー")

if view == "mobile":
    selected_page = st.selectbox(
        "分析モードを選択",
        pages,
        index=pages.index(st.session_state.current_page),
        label_visibility="collapsed",
    )
    if selected_page != st.session_state.current_page:
        st.session_state.current_page = selected_page
else:
    for i in range(0, len(pages), 4):
        cols = st.columns(4)
        for j in range(4):
            if i + j < len(pages):
                page_name = pages[i + j]
                with cols[j]:
                    btn_type = "primary" if st.session_state.current_page == page_name else "secondary"
                    if st.button(page_name, use_container_width=True, key=f"nav_btn_{i+j}", type=btn_type):
                        st.session_state.current_page = page_name
                        st.rerun()
```

（コードレビューで指摘・修正：`st.selectbox`はユーザー操作時にStreamlitが自動的に1回再実行するため、`st.session_state.current_page`の更新は次の`analysis_mode = st.session_state.current_page`行より前に反映される。ボタン版と違い見た目の状態（primary/secondary色）を明示的に更新する必要がないため、`st.rerun()`は不要かつ二重再実行による遅延の原因になる。）

- [ ] **Step 3: シンタックスチェック**

Run: `python3 -c "import ast; ast.parse(open('app.py').read())"`
Expected: エラーなく終了

- [ ] **Step 4: 手動確認**

`streamlit run app.py` で以下を確認:
1. 「🖥 PC表示」のとき：今まで通り4列ボタングリッドが表示される
2. 「📱 スマホ表示」のとき：ボタングリッドの代わりにプルダウン（セレクトボックス）が表示される
3. プルダウンで別の分析モードを選ぶと、正しくそのページに切り替わる
4. スマホ表示のままブラウザ幅を戻してもプルダウンのまま（`view`はウィンドウ幅ではなくトグルで決まる）であることを確認

- [ ] **Step 5: コミット**

```bash
git add app.py
git commit -m "スマホ表示時の分析メニューをプルダウンに変更"
```

---

### Task 4: 診療行為一覧分析のマトリクス表に直近6ヶ月表示を適用

**Files:**
- Modify: `app.py`（`診療行為一覧分析` セクション内、`st.write("#### 📋 月別詳細テーブル（すべての診療行為・総点数）")` を含むブロック）

- [ ] **Step 1: 現在のコードを確認する**

```python
    st.write("#### 📋 月別詳細テーブル（すべての診療行為・総点数）")
    
    matrix_full_df = df_curr_act.pivot_table(index='診療行為名称', columns='月', values='総点数', aggfunc='sum').fillna(0)
    matrix_cols = [m for m in valid_months_act if m in matrix_full_df.columns]
    matrix_full_df = matrix_full_df.reindex(columns=matrix_cols)
    matrix_full_df['年間合計'] = matrix_full_df.sum(axis=1)
    matrix_full_df = matrix_full_df.sort_values('年間合計', ascending=False)
    
    sum_row_act = matrix_full_df.sum(numeric_only=True)
    sum_row_act.name = '★点数合計'
    matrix_full_df = pd.concat([matrix_full_df, pd.DataFrame([sum_row_act])])
    
    def style_full_matrix(df):
        styler = df.style
        fmt = {col: "{:,.0f}" for col in df.columns}
        styler = styler.format(fmt)
        
        def apply_bold_total(row):
            if row.name == '★点数合計':
                return ['font-weight: bold; background-color: #f0f2f6'] * len(row)
            return [''] * len(row)
            
        return styler.apply(apply_bold_total, axis=1)
        
    st.dataframe(style_full_matrix(matrix_full_df), use_container_width=True)

# ==========================================
```

（このファイルには同一パターンの表が「検査一覧分析」セクションにも存在する。区別するため、直後の `# ==========================================` の前に「診療行為一覧分析」セクション内であることを確認してから編集すること。Task 5と混同しないこと。）

- [ ] **Step 2: `st.dataframe(style_full_matrix(matrix_full_df), use_container_width=True)` の行だけを、直近6ヶ月切替ロジックに置き換える**

置き換え前（この1行）:
```python
    st.dataframe(style_full_matrix(matrix_full_df), use_container_width=True)
```

置き換え後:
```python
    if view == "mobile":
        show_full_act = st.session_state.get("show_full_matrix_act", False)
        btn_label_act = "📅 直近6ヶ月だけ表示" if show_full_act else "📅 全期間を表示"
        if st.button(btn_label_act, key="show_full_matrix_act_btn"):
            st.session_state["show_full_matrix_act"] = not show_full_act
            st.rerun()
        display_matrix_act = matrix_full_df if show_full_act else limit_recent_months(matrix_full_df, matrix_cols, n=6)
    else:
        display_matrix_act = matrix_full_df

    st.dataframe(style_full_matrix(display_matrix_act), use_container_width=True)
```

- [ ] **Step 3: シンタックスチェック**

Run: `python3 -c "import ast; ast.parse(open('app.py').read())"`
Expected: エラーなく終了

- [ ] **Step 4: 手動確認**

`streamlit run app.py` で「診療行為一覧分析」を開き:
1. PC表示：今まで通り全期間の表がそのまま表示される（ボタンは出ない）
2. スマホ表示：直近6ヶ月＋年間合計列だけの表と、「📅 全期間を表示」ボタンが出る
3. ボタンを押すと全期間表示に変わり、ボタンのラベルが「📅 直近6ヶ月だけ表示」に変わる
4. もう一度押すと直近6ヶ月表示に戻る

- [ ] **Step 5: コミット**

```bash
git add app.py
git commit -m "診療行為一覧分析のマトリクス表にスマホ用の直近6ヶ月表示を追加"
```

---

### Task 5: 検査一覧分析のマトリクス表に直近6ヶ月表示を適用

**Files:**
- Modify: `app.py`（`検査一覧分析` セクション内、`st.write("#### 📋 月別詳細テーブル（すべての検査・総点数）")` を含むブロック）

- [ ] **Step 1: 現在のコードを確認する**

Task 4と同一パターンだが、こちらは「検査一覧分析」セクション内で、直前の見出しが `st.write("#### 📋 月別詳細テーブル（すべての検査・総点数）")` になっている方を編集する。

```python
    st.write("#### 📋 月別詳細テーブル（すべての検査・総点数）")
    
    matrix_full_df = df_curr_act.pivot_table(index='診療行為名称', columns='月', values='総点数', aggfunc='sum').fillna(0)
    matrix_cols = [m for m in valid_months_act if m in matrix_full_df.columns]
    matrix_full_df = matrix_full_df.reindex(columns=matrix_cols)
    matrix_full_df['年間合計'] = matrix_full_df.sum(axis=1)
    matrix_full_df = matrix_full_df.sort_values('年間合計', ascending=False)
    
    sum_row_act = matrix_full_df.sum(numeric_only=True)
    sum_row_act.name = '★点数合計'
    matrix_full_df = pd.concat([matrix_full_df, pd.DataFrame([sum_row_act])])
    
    def style_full_matrix(df):
        styler = df.style
        fmt = {col: "{:,.0f}" for col in df.columns}
        styler = styler.format(fmt)
        
        def apply_bold_total(row):
            if row.name == '★点数合計':
                return ['font-weight: bold; background-color: #f0f2f6'] * len(row)
            return [''] * len(row)
            
        return styler.apply(apply_bold_total, axis=1)
        
    st.dataframe(style_full_matrix(matrix_full_df), use_container_width=True)
```

- [ ] **Step 2: 検査一覧分析セクション側の `st.dataframe(...)` の行だけを置き換える**

Task 4と同じ内容だが、セッションステートのキーは重複しないよう `_exam` サフィックスにする:

```python
    if view == "mobile":
        show_full_exam = st.session_state.get("show_full_matrix_exam", False)
        btn_label_exam = "📅 直近6ヶ月だけ表示" if show_full_exam else "📅 全期間を表示"
        if st.button(btn_label_exam, key="show_full_matrix_exam_btn"):
            st.session_state["show_full_matrix_exam"] = not show_full_exam
            st.rerun()
        display_matrix_exam = matrix_full_df if show_full_exam else limit_recent_months(matrix_full_df, matrix_cols, n=6)
    else:
        display_matrix_exam = matrix_full_df

    st.dataframe(style_full_matrix(display_matrix_exam), use_container_width=True)
```

- [ ] **Step 3: シンタックスチェック**

Run: `python3 -c "import ast; ast.parse(open('app.py').read())"`
Expected: エラーなく終了

- [ ] **Step 4: 手動確認**

「検査一覧分析」ページで、Task 4のStep 4と同じ4項目を確認する。加えて、診療行為一覧分析のボタン（Task 4）と検査一覧分析のボタン（本タスク）の開閉状態が互いに独立していることを確認する（片方を「全期間表示」にしても、もう片方は直近6ヶ月のまま）。

- [ ] **Step 5: コミット**

```bash
git add app.py
git commit -m "検査一覧分析のマトリクス表にスマホ用の直近6ヶ月表示を追加"
```

---

### Task 6: 医師別診療実績のマトリクス表に直近6ヶ月表示を適用

**Files:**
- Modify: `app.py`（`医師別診療実績（月別）` セクション内、`st.write("### 📋 月別詳細マトリクス（医師×月・保険点数）")` を含むブロック）

- [ ] **Step 1: 現在のコードを確認する**

```python
    st.write("---")
    st.write("### 📋 月別詳細マトリクス（医師×月・保険点数）")

    matrix_doc_df = df_curr_doc.pivot_table(index='医師', columns='月', values='保険点数', aggfunc='sum').fillna(0)
    matrix_doc_cols = [m for m in valid_months_doc if m in matrix_doc_df.columns]
    matrix_doc_df = matrix_doc_df.reindex(columns=matrix_doc_cols)
    matrix_doc_df['年間合計'] = matrix_doc_df.sum(axis=1)
    matrix_doc_df = matrix_doc_df.sort_values('年間合計', ascending=False)

    sum_row_matrix_doc = matrix_doc_df.sum(numeric_only=True)
    sum_row_matrix_doc.name = '★点数合計'
    matrix_doc_df = pd.concat([matrix_doc_df, pd.DataFrame([sum_row_matrix_doc])])

    def style_full_matrix_doc(df):
        styler = df.style
        fmt = {col: "{:,.0f}" for col in df.columns}
        styler = styler.format(fmt)

        def apply_bold_total(row):
            if row.name == '★点数合計':
                return ['font-weight: bold; background-color: #f0f2f6'] * len(row)
            return [''] * len(row)

        return styler.apply(apply_bold_total, axis=1)

    st.dataframe(style_full_matrix_doc(matrix_doc_df), use_container_width=True)
```

- [ ] **Step 2: `st.dataframe(style_full_matrix_doc(matrix_doc_df), use_container_width=True)` の行だけを置き換える**

```python
    if view == "mobile":
        show_full_doc = st.session_state.get("show_full_matrix_doc", False)
        btn_label_doc = "📅 直近6ヶ月だけ表示" if show_full_doc else "📅 全期間を表示"
        if st.button(btn_label_doc, key="show_full_matrix_doc_btn"):
            st.session_state["show_full_matrix_doc"] = not show_full_doc
            st.rerun()
        display_matrix_doc = matrix_doc_df if show_full_doc else limit_recent_months(matrix_doc_df, matrix_doc_cols, n=6)
    else:
        display_matrix_doc = matrix_doc_df

    st.dataframe(style_full_matrix_doc(display_matrix_doc), use_container_width=True)
```

- [ ] **Step 3: シンタックスチェック**

Run: `python3 -c "import ast; ast.parse(open('app.py').read())"`
Expected: エラーなく終了

- [ ] **Step 4: 手動確認**

「医師別診療実績（月別）」ページで、Task 4のStep 4と同じ4項目を確認する。

- [ ] **Step 5: コミット**

```bash
git add app.py
git commit -m "医師別診療実績のマトリクス表にスマホ用の直近6ヶ月表示を追加"
```

---

### Task 7: スマホ幅向けのCSSメディアクエリを追加

**Files:**
- Modify: `app.py:150-155`付近（共通CSS設定の `</style>` 直前）

- [ ] **Step 1: 現在のコードを確認する**

```python
    .ai-box {
        border: 2px solid #D5D8DC;
        background-color: #F8F9F9;
        border-radius: 10px;
        padding: 15px;
        min-height: 220px;
    }
    </style>
    """, unsafe_allow_html=True)
```

- [ ] **Step 2: `</style>` の直前にメディアクエリを追加する**

```python
    .ai-box {
        border: 2px solid #D5D8DC;
        background-color: #F8F9F9;
        border-radius: 10px;
        padding: 15px;
        min-height: 220px;
    }

    /* === スマホ幅（768px以下）での見やすさ調整 === */
    @media (max-width: 768px) {
        .header-title {
            font-size: 1.3rem !important;
        }
        div.stButton > button {
            font-size: 13px !important;
            height: 55px !important;
        }
        .stDataFrame, .stTable {
            font-size: 12px !important;
        }
    }
    </style>
    """, unsafe_allow_html=True)
```

- [ ] **Step 3: シンタックスチェック**

Run: `python3 -c "import ast; ast.parse(open('app.py').read())"`
Expected: エラーなく終了

- [ ] **Step 4: 手動確認**

`streamlit run app.py` をブラウザで開き、開発者ツールでビューポート幅を375px程度に変更し、見出し・ボタン・表の文字サイズが小さくなって収まりやすくなることを確認する（PC幅に戻すと通常サイズに戻ることも確認）。

- [ ] **Step 5: コミット**

```bash
git add app.py
git commit -m "スマホ幅向けのCSSメディアクエリを追加"
```

---

### Task 8: 総合動作確認とローカル作業コピーへの同期

**Files:**
- Modify: なし（確認のみ）
- Sync: `/mnt/c/Users/大西修司/Downloads/app.py`（ローカル作業コピー）

- [ ] **Step 1: 全8分析モードをPC表示・スマホ表示の両方で一通り確認する**

`streamlit run app.py` で以下を確認する:
1. `?view` パラメータなしでアクセスした場合、変更前と見た目・挙動が変わらないこと（PC表示が既定のまま）
2. 「📱 スマホ表示」に切り替えた状態で、8つの分析モードすべてがプルダウンから選択できること
3. 診療行為一覧分析・検査一覧分析・医師別診療実績（月別）の3箇所で、直近6ヶ月表示⇄全期間表示の切り替えボタンがそれぞれ独立して正しく動作すること
4. レセプト分析・外来収入金額推移分析・受付患者数推移分析・年齢別構成比分析・AI総合経営アドバイスの各ページが、スマホ表示でも今まで通り正しくデータを表示すること（これらのページは今回変更していないため、表示崩れがないことの確認が目的）

- [ ] **Step 2: ローカル作業コピー（Downloads）に最新の `app.py` を同期する**

```bash
cp app.py "/mnt/c/Users/大西修司/Downloads/app.py"
```

- [ ] **Step 3: git logで一連のコミットを確認する**

Run: `git log --oneline -10`
Expected: Task 1〜7の7件のコミット（+ 既存のスペックドキュメントコミット）が新しい順に並んでいる

GitHubへのpush（`git push origin main`）は、ユーザーに確認のうえ別途実行する。
