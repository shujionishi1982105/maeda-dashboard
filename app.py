import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import glob
import re
import os
import unicodedata

# ページ設定は一番最初に行う必要があります
st.set_page_config(page_title="まえだ耳鼻咽喉科 経営分析", layout="wide")

# ==========================================
# 🔒 セキュリティ：ログイン認証機能
# ==========================================
def check_password():
    """IDとパスワードが正しいかチェックする関数"""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        # ログイン画面のデザイン
        st.markdown("<h2 style='text-align: center; color: #2C3E50; margin-top: 50px;'>🏥 まえだ耳鼻咽喉科<br>経営分析システム ログイン</h2>", unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("<div style='background-color: #F8F9F9; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>", unsafe_allow_html=True)
            username = st.text_input("👤 ID（ユーザー名）")
            password = st.text_input("🔑 パスワード", type="password")
            
            if st.button("ログイン", use_container_width=True, type="primary"):
                # ★★★ ここで好きなIDとパスワードを設定します ★★★
                # 現在は ID: admin / パスワード: maeda2026 になっています。後で自由に変更してください。
                if username == "admin" and password == "maeda2026":
                    st.session_state["password_correct"] = True
                    st.rerun() # 画面をリロードしてダッシュボードを表示
                else:
                    st.error("⚠️ IDまたはパスワードが間違っています。")
            st.markdown("</div>", unsafe_allow_html=True)
        return False
    return True

# 認証が通っていない場合はここで処理をストップ（これより下のコードは実行・表示されません）
if not check_password():
    st.stop()


# ==========================================
# 以下、これまでのダッシュボードの全コード
# ==========================================

# --- 共通CSS設定 ---
st.markdown("""
    <style>
    [data-testid="stTable"] th:first-child, [data-grid-column-id="月"], .stDataFrame div[data-testid="stTable"] div:first-child { min-width: 100px !important; max-width: 100px !important; width: 100px !important; }
    .record-text { font-size: 1.1rem; font-weight: bold; color: #E74C3C; }
    div[data-baseweb="select"], div[data-baseweb="select"] * { cursor: pointer !important; }
    div.stButton > button { height: 65px !important; border-radius: 10px !important; font-weight: bold !important; font-size: 15px !important; box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important; transition: all 0.2s ease-in-out !important; white-space: normal !important; line-height: 1.3 !important; }
    div.stButton > button[kind="secondary"] { background-color: #EBF5FB !important; border: 1px solid #AED6F1 !important; color: #154360 !important; }
    div.stButton > button[kind="primary"] { background-color: #E74C3C !important; border: 1px solid #E74C3C !important; color: #FFFFFF !important; }
    div.stButton > button:hover { transform: translateY(-2px) !important; filter: brightness(0.95) !important; }
    .header-container { display: flex; align-items: center; gap: 20px; margin-bottom: 20px; }
    [data-testid="stImage"] { margin-top: 10px !important; }
    .header-title { margin: 0 !important; padding-top: 15px !important; }
    .kpi-container { height: 250px; overflow-y: auto; border-radius: 0.5rem; padding: 1rem; }
    .kpi-container.stInfo { border: 1px solid rgba(25, 230, 255, 0.2); background-color: rgba(25, 230, 255, 0.1); }
    .kpi-container.stSuccess { border: 1px solid rgba(0, 255, 128, 0.2); background-color: rgba(0, 255, 128, 0.1); }
    .kpi-container.stWarning { border: 1px solid rgba(255, 204, 0, 0.2); background-color: rgba(255, 204, 0, 0.1); }
    .target-box { border: 2px solid #AED6F1; background-color: #EBF5FB; border-radius: 10px; padding: 15px; min-height: 220px; }
    .ai-box { border: 2px solid #D5D8DC; background-color: #F8F9F9; border-radius: 10px; padding: 15px; min-height: 220px; }
    .revision-box { border: 2px solid #F5B041; background-color: #FEF9E7; border-radius: 10px; padding: 15px; margin-top: 15px; height: 100%; }
    .revision-box h4 { margin-top: 0; margin-bottom: 10px; color: #2C3E50; font-size: 1.1rem; border-bottom: 1px dashed #ccc; padding-bottom: 5px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 共通関数 
# ==========================================
@st.cache_data(ttl=600)
def get_latest_news(url):
    try:
        df = pd.read_csv(url)
        return df if not df.empty else None
    except Exception: return None

@st.cache_data(ttl=3600)
def get_clean_df(year_str):
    files = glob.glob(f"*{year_str}*レセプト*.csv")
    if not files: return None
    for enc in ['utf-8-sig', 'cp932', 'shift-jis']:
        try:
            df = pd.read_csv(files[0], encoding=enc)
            df.columns = [re.sub(r'\s*[\(（].*?[\)）]', '', c).strip() for c in df.columns]
            return df
        except: continue
    return None

def get_latest_complete_month(df_rece, year_str):
    valid_months = [f"{i}月" for i in range(1, 13)]
    df_tmp = df_rece[df_rece['月'].isin(valid_months)].copy()
    df_tmp['レセ単価_num'] = pd.to_numeric(df_tmp['レセ単価'].astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce').fillna(0)
    active_rows = df_tmp[df_tmp['レセ単価_num'] > 0]
    if active_rows.empty: return None, None, False, None
    original_latest_month = active_rows.iloc[-1]['月']
    try: target_year_num = int(re.search(r'\d+', year_str).group())
    except: return original_latest_month, active_rows.iloc[-1], False, original_latest_month
    for idx in range(len(active_rows) - 1, -1, -1):
        tmp_month = active_rows.iloc[idx]['月']
        target_month_num = int(re.search(r'\d+', tmp_month).group())
        act_exists = False
        for f in glob.glob("*診療行為一覧*.csv"):
            f_norm = unicodedata.normalize('NFKC', f)
            match = re.search(r'(?:R|令和)?0*(\d+)年0*(\d+)月', f_norm)
            if match and int(match.group(1)) == target_year_num and int(match.group(2)) == target_month_num:
                act_exists = True; break
        if act_exists: return tmp_month, active_rows.iloc[idx], (tmp_month != original_latest_month), original_latest_month
    return original_latest_month, active_rows.iloc[-1], False, original_latest_month

@st.cache_data(ttl=3600)
def get_act_summary_for_ai(y_str, m_str):
    try:
        target_y = int(re.search(r'\d+', y_str).group())
        target_m = int(re.search(r'\d+', m_str).group())
    except: return pd.DataFrame()
    for f in glob.glob("*診療行為一覧*.csv"):
        f_norm = unicodedata.normalize('NFKC', f)
        match = re.search(r'(?:R|令和)?0*(\d+)年0*(\d+)月', f_norm)
        if match and int(match.group(1)) == target_y and int(match.group(2)) == target_m:
            for enc in ['utf-8-sig', 'cp932', 'shift-jis']:
                try:
                    df_act = pd.read_csv(f, encoding=enc)
                    col_name = next((c for c in df_act.columns if '診療行為名称' in c.replace(' ', '')), None)
                    if not col_name: continue
                    df_act = df_act.dropna(subset=[col_name])
                    df_act = df_act[~df_act[col_name].astype(str).str.contains('合計')]
                    col_cat = '診療区分' if '診療区分' in df_act.columns else None
                    for col in ['回数', '総点数 (点)']:
                        if col in df_act.columns: df_act[col] = pd.to_numeric(df_act[col].astype(str).str.replace(r'[，,点円％%]', '', regex=True), errors='coerce').fillna(0)
                    if col_cat:
                        summary = df_act.groupby([col_cat, col_name])[['回数', '総点数 (点)']].sum().reset_index()
                        return summary.rename(columns={col_name: '名称', col_cat: '区分', '総点数 (点)': '総点数'})
                    else:
                        summary = df_act.groupby(col_name)[['回数', '総点数 (点)']].sum().reset_index()
                        summary['区分'] = '不明'
                        return summary.rename(columns={col_name: '名称', '総点数 (点)': '総点数'})
                except Exception: continue
    return pd.DataFrame()

# ==========================================
# タイトル＆メニュー 
# ==========================================
try: base_dir = os.path.dirname(os.path.abspath(__file__))
except NameError: base_dir = os.getcwd()

header_cols = st.columns([1, 15])
with header_cols[0]:
    if os.path.exists("logo.png"): st.image("logo.png", use_container_width=True)
    else: st.markdown("<h1 style='text-align: center; margin: 0; padding: 0;'>🏥</h1>", unsafe_allow_html=True)
with header_cols[1]: st.markdown('<h1 class="header-title">まえだ耳鼻咽喉科 経営分析ダッシュボード</h1>', unsafe_allow_html=True)

pages = ["レセプト分析", "外来収入金額推移分析", "受付患者数（初再診別）推移分析", "年齢別構成比分析", "診療行為一覧分析", "検査一覧分析", "AI総合経営アドバイス"]
if 'current_page' not in st.session_state: st.session_state.current_page = pages[0]

st.write("### 🔍 分析メニュー")
for i in range(0, len(pages), 4):
    cols = st.columns(4)
    for j in range(4):
        if i + j < len(pages):
            page_name = pages[i + j]
            with cols[j]:
                if st.button(page_name, use_container_width=True, key=f"nav_btn_{i+j}", type="primary" if st.session_state.current_page == page_name else "secondary"):
                    st.session_state.current_page = page_name
                    st.rerun()
st.write("---")
analysis_mode = st.session_state.current_page

# ==========================================
# A. レセプト分析
# ==========================================
if analysis_mode == "レセプト分析":
    files_all = glob.glob("*レセプト*.csv")
    years_found = sorted(list(set(re.search(r'(R\d+年)', f).group(1) for f in files_all if re.search(r'(R\d+年)', f))))
    if not years_found: st.error("レセプトのCSVファイルが見つかりません。"); st.stop()

    kpi_placeholder = st.container()
    col_year, col_item = st.columns(2)
    with col_year: selected_year = st.selectbox("📅 表示する年度を選択", years_found, index=len(years_found)-1)

    prev_year = f"R{int(re.search(r'\d+', selected_year).group()) - 1}年"
    df_now_full = get_clean_df(selected_year)
    df_prev_full = get_clean_df(prev_year)
    valid_months = [f"{i}月" for i in range(1, 13)]

    if df_now_full is not None:
        choices = [c for c in df_now_full.columns.tolist() if df_now_full[c].dtype in ['float64', 'int64'] and not any(k in c for k in ["月", "総", "前年", "比", "枚数"])]
        with col_item: selected_item = st.selectbox("🔍 グラフ項目", options=choices, index=choices.index("レセ単価") if "レセ単価" in choices else 0)

    with kpi_placeholder:
        if df_now_full is not None and 'レセ単価' in df_now_full.columns:
            TARGET_PRICE = 750
            target_month, latest_row, is_fallback, original_latest_month = get_latest_complete_month(df_now_full, selected_year)
            
            if target_month:
                price_curr = latest_row['レセ単価_num']
                latest_m_num = int(target_month.replace('月', ''))
                df_tmp = df_now_full[df_now_full['月'].isin(valid_months)].copy()
                
                prev_month = "12月" if latest_m_num == 1 else f"{latest_m_num - 1}月"
                price_prev = 0
                if latest_m_num == 1 and df_prev_full is not None and 'レセ単価' in df_prev_full.columns:
                    p_val = df_prev_full.loc[df_prev_full['月'] == prev_month, 'レセ単価'].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                    price_prev = float(p_val.values[0]) if len(p_val) > 0 and p_val.values[0] != '' else 0
                elif latest_m_num != 1:
                    df_tmp['レセ単価_num'] = pd.to_numeric(df_tmp['レセ単価'].astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce').fillna(0)
                    p_val = df_tmp.loc[df_tmp['月'] == prev_month, 'レセ単価_num']
                    price_prev = p_val.values[0] if len(p_val) > 0 else 0
                    
                gap = TARGET_PRICE - price_curr
                gap_html = f"<span style='color:#E74C3C; font-size:1.2rem;'><b>{gap:,.0f} 点 不足</b></span>" if gap > 0 else f"<span style='color:#27AE60; font-size:1.2rem;'><b>目標達成！ (+{abs(gap):,.0f} 点)</b></span>"
                fallback_notice = f"<div style='background-color:#FFF3CD; color:#856404; padding:5px 10px; border-radius:5px; font-size:0.85em; margin-bottom:10px;'>⚠️ {original_latest_month}の診療データが未確定のため、データが揃っている<b>【{target_month}】</b>を基準に表示しています。</div>" if is_fallback else ""

                if price_prev == 0: ai_comment_html = "<p>前月のデータが存在しないため、比較分析はスキップしました。</p>"
                else:
                    diff_price = price_curr - price_prev
                    trend = "低下" if diff_price < 0 else "上昇"
                    trend_color = "#E74C3C" if diff_price < 0 else "#2E86C1"
                    
                    df_a_curr = get_act_summary_for_ai(selected_year, target_month)
                    prev_y_str = f"R{int(re.search(r'\d+', selected_year).group()) - 1}年" if latest_m_num == 1 else selected_year
                    df_a_prev = get_act_summary_for_ai(prev_y_str, prev_month)
                    
                    detailed_comment = ""
                    if not df_a_curr.empty and not df_a_prev.empty:
                        df_a_diff = pd.merge(df_a_curr, df_a_prev, on=['区分', '名称'], how='outer', suffixes=('_curr', '_prev')).fillna(0)
                        df_a_diff['点数差'] = df_a_diff['総点数_curr'] - df_a_diff['総点数_prev']
                        top_a_diffs = df_a_diff[df_a_diff['点数差'] < 0].sort_values('点数差', ascending=True).head(3) if diff_price < 0 else df_a_diff[df_a_diff['点数差'] > 0].sort_values('点数差', ascending=False).head(3)
                        
                        detailed_comment = "<ul style='line-height:1.5; padding-left:15px; font-size:0.9em;'>"
                        for _, r in top_a_diffs.iterrows():
                            sign_a = "+" if r['点数差'] > 0 else ""
                            cat_clean = re.sub(r'^\d+\s*', '', str(r['区分']))
                            if '外来' in str(r['名称']) or '指導' in cat_clean or '管理' in cat_clean: cat_clean = '管理'
                            elif '手術' in cat_clean: cat_clean = '手術'
                            elif '検査' in cat_clean: cat_clean = '検査'
                            elif '処置' in cat_clean: cat_clean = '処置'
                            detailed_comment += f"<li style='margin-bottom:3px;'><b>【{cat_clean}】</b>{r['名称']} <span style='color:#555;'>({sign_a}{r['点数差']:,.0f}点)</span></li>"
                        detailed_comment += "</ul>"
                    
                    ai_comment_html = f"<p>【{target_month}】のレセ単価は前月({prev_month})と比較して <b style='color:{trend_color};'>{abs(diff_price):,.0f} 点 {trend}</b> しています。</p><p style='font-size:0.95em;'>主に以下の要因が影響しています。</p>{detailed_comment}"

                col_kpi, col_ai = st.columns(2)
                with col_kpi: st.markdown("<div class='target-box'>" + f"{fallback_notice}" + "<h4 style='margin-top:0; margin-bottom:10px; color:#2C3E50; font-size:1.1rem; border-bottom:1px dashed #ccc; padding-bottom:5px;'>🎯 今年の目標レセプト単価</h4><div style='font-size: 2.8rem; font-weight: bold; text-align: center; margin: 15px 0; color: #2E86C1;'>750 <span style='font-size: 1.2rem; color:#333;'>点</span></div><hr style='margin: 10px 0;'><p style='margin-bottom:5px;'><b>【最新実績】</b> " + f"{selected_year}{target_month}： <b style='font-size:1.2rem;'>{price_curr:,.0f} 点</b></p><p style='margin-bottom:0;'><b>【進捗】</b> {gap_html}</p></div>", unsafe_allow_html=True)
                with col_ai: st.markdown("<div class='ai-box'><h4 style='margin-top:0; margin-bottom:10px; color:#2C3E50; font-size:1.1rem; border-bottom:1px dashed #ccc; padding-bottom:5px;'>🤖 AI要因分析レポート</h4>" + f"{ai_comment_html}" + "</div>", unsafe_allow_html=True)
            st.write("---")

    if df_now_full is not None:
        df_now = df_now_full[df_now_full['月'].isin(valid_months)].copy()
        df_prev = df_prev_full[df_prev_full['月'].isin(valid_months)].copy() if df_prev_full is not None else None
        plot_df = df_now[['月', selected_item]].copy()
        plot_df.columns = ['月', '当年']
        if df_prev is not None and selected_item in df_prev.columns:
            prev_data = df_prev[['月', selected_item]].copy()
            prev_data.columns = ['月', '前年']
            plot_df = pd.merge(plot_df, prev_data, on='月', how='left').fillna(0)
            plot_df['比率'] = plot_df.apply(lambda x: round((x['当年'] / x['前年'] * 100), 1) if x['前年'] > 0 else 0, axis=1)
        else: plot_df['前年'], plot_df['比率'] = 0, 0

        colors = ['#E74C3C' if 0 < val < 100 else ('#2E86C1' if val >= 100 else '#000000') for val in plot_df['比率']]
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=plot_df['月'], y=plot_df['当年'], mode='lines+markers+text', name='当年', line=dict(color='#2E86C1', width=4), text=plot_df['比率'].apply(lambda x: f"{x}%" if x > 0 else ""), textposition="top center", textfont=dict(color=colors, size=13, family="Arial Black"), hovertemplate="<b>%{x}</b><br>当年: %{y:,.0f}<extra></extra>"))
        if df_prev is not None: fig.add_trace(go.Scatter(x=plot_df['月'], y=plot_df['前年'], mode='lines+markers', name='前年', line=dict(color='#ABB2B9', width=2, dash='dot'), hovertemplate="前年: %{y:,.0f}<extra></extra>"))
        fig.update_layout(xaxis=dict(title="診療月", type='category', categoryorder='array', categoryarray=valid_months), yaxis_title="点数 / 件数", hovermode="x", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

        st.write("---")
        st.write("### 📋 詳細数値データ（年間一覧）")
        display_df = df_now_full.copy()
        display_df['月'] = display_df['月'].replace('総数', '総数・平均')
        prev_total_row = df_prev_full[df_prev_full['月'] == '総数'].iloc[0] if df_prev_full is not None and not df_prev_full[df_prev_full['月'] == '総数'].empty else None

        def make_styled_df(df):
            styled = df.set_index('月').style.format({c: "{:.1f}%" if "比" in c else "{:,.0f}" for c in df.columns if c != '月'})
            def apply_colors(row):
                styles = ['color: black'] * len(row)
                for i, col in enumerate(row.index):
                    val = row[col]
                    if "比" in col:
                        try:
                            v = float(val)
                            if v >= 100: styles[i] = 'color: #2E86C1; font-weight: bold'
                            elif 0 < v < 100: styles[i] = 'color: #E74C3C; font-weight: bold'
                        except: pass
                    elif row.name == '総数・平均' and prev_total_row is not None:
                        p_val = prev_total_row.get(col, 0)
                        if val >= p_val and p_val != 0: styles[i] = 'color: #2E86C1; font-weight: bold'
                        elif val < p_val and p_val != 0: styles[i] = 'color: #E74C3C; font-weight: bold'
                return styles
            return styled.apply(apply_colors, axis=1)

        st.dataframe(make_styled_df(display_df), use_container_width=True)

# ==========================================
# B. 外来収入金額推移分析
# ==========================================
elif analysis_mode == "外来収入金額推移分析":
    income_files = glob.glob("*外来収入金額*.csv")
    if not income_files: st.error("ファイルが見つかりません。"); st.stop()

    data_list = []
    for f in income_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str, month_str = match.group(1) + "年", match.group(2) + "月"
            for enc in ['utf-8-sig', 'cp932', 'shift-jis']:
                try:
                    df_m = pd.read_csv(f, encoding=enc)
                    df_m = df_m[df_m['日'].str.contains('日', na=False)]
                    for col in df_m.columns:
                        if col != '日': df_m[col] = pd.to_numeric(df_m[col].astype(str).str.replace(',', ''), errors='coerce')
                    summary = df_m.drop(columns=['日']).sum()
                    summary['年度'], summary['月'] = year_str, month_str
                    data_list.append(summary)
                    break
                except: continue

    income_df = pd.DataFrame(data_list).drop_duplicates(subset=['年度', '月'])
    available_years = sorted(list(income_df['年度'].unique()), key=lambda x: int(re.search(r'\d+', x).group()))
    col_year_inc, col_item_inc = st.columns(2)
    with col_year_inc: selected_year_inc = st.selectbox("📅 表示する年度", available_years, index=len(available_years)-1, key="inc_year")
    current_num_inc = int(re.search(r'\d+', selected_year_inc).group())
    prev_year_inc = f"R{current_num_inc - 1}年"
    
    items_inc = [c for c in income_df.columns if c not in ['年度', '月']]
    with col_item_inc: selected_income = st.selectbox("🔍 分析項目", items_inc, index=items_inc.index("外来収入金額 (円)") if "外来収入金額 (円)" in items_inc else 0, key="inc_item")

    st.subheader(f"📊 {selected_year_inc} 【{selected_income}】の推移分析")
    valid_months_inc = [f"{i}月" for i in range(1, 13)]
    df_curr_inc = income_df[income_df['年度'] == selected_year_inc].copy()
    df_prev_inc = income_df[income_df['年度'] == prev_year_inc].copy()
    
    plot_df_inc = pd.merge(pd.DataFrame({'月': valid_months_inc}), df_curr_inc[['月', selected_income]], on='月', how='left').rename(columns={selected_income: '当年'})
    plot_df_inc = pd.merge(plot_df_inc, df_prev_inc[['月', selected_income]], on='月', how='left').rename(columns={selected_income: '前年'}).fillna(0)
    plot_df_inc['前年比'] = plot_df_inc.apply(lambda x: round(x['当年']/x['前年']*100, 1) if x['前年'] > 0 else 0, axis=1)

    fig_i = go.Figure()
    if not df_prev_inc.empty: fig_i.add_trace(go.Bar(x=plot_df_inc['月'], y=plot_df_inc['前年'], name=f'前年 ({prev_year_inc})', marker_color='#ABB2B9', text=plot_df_inc['前年'].apply(lambda x: f"{x/10000:.0f}万" if x >= 10000 else (f"{x:,.0f}" if x > 0 else "")), textposition='outside', textfont=dict(size=11), hovertemplate="前年: %{y:,.0f}<extra></extra>"))
    fig_i.add_trace(go.Bar(x=plot_df_inc['月'], y=plot_df_inc['当年'], name=f'当年 ({selected_year_inc})', marker_color='#2E86C1', text=plot_df_inc['当年'].apply(lambda x: f"{x/10000:.0f}万" if x >= 10000 else (f"{x:,.0f}" if x > 0 else "")), textposition='outside', textfont=dict(size=11), hovertemplate="当年: %{y:,.0f}<extra></extra>"))
    
    max_val = max(plot_df_inc['当年'].max(), plot_df_inc['前年'].max())
    fig_i.update_layout(barmode='group', xaxis_title="診療月", yaxis_title=selected_income, yaxis=dict(range=[0, max_val * 1.15]), hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), xaxis=dict(type='category', categoryorder='array', categoryarray=valid_months_inc))
    st.plotly_chart(fig_i, use_container_width=True, config={'displayModeBar': False})

    st.write("---")
    st.write("### 📋 詳細数値データ（年間一覧）")
    sum_curr, sum_prev = plot_df_inc['当年'].sum(), plot_df_inc['前年'].sum()
    sum_ratio = round(sum_curr / sum_prev * 100, 1) if sum_prev > 0 else 0
    st_df_inc = pd.concat([plot_df_inc, pd.DataFrame([{'月': '総数・平均', '当年': sum_curr, '前年': sum_prev, '前年比': sum_ratio}])], ignore_index=True)
    
    def style_income_table(df):
        styler = df.set_index('月').style.format({'当年': "{:,.0f}", '前年': "{:,.0f}", '前年比': "{:.1f}%"})
        def apply_colors(row):
            styles = ['color: black'] * len(row)
            for i, col in enumerate(row.index):
                val = row[col]
                if col == '前年比':
                    try:
                        v = float(val)
                        if v >= 100: styles[i] = 'color: #2E86C1; font-weight: bold'
                        elif 0 < v < 100: styles[i] = 'color: #E74C3C; font-weight: bold'
                    except: pass
                elif row.name == '総数・平均' and col == '当年':
                    p_val = row['前年']
                    if val >= p_val and p_val != 0: styles[i] = 'color: #2E86C1; font-weight: bold'
                    elif val < p_val and p_val != 0: styles[i] = 'color: #E74C3C; font-weight: bold'
            return styles
        return styler.apply(apply_colors, axis=1)
    st.dataframe(style_income_table(st_df_inc), use_container_width=True)

# ==========================================
# C. 受付患者数推移分析
# ==========================================
elif analysis_mode == "受付患者数（初再診別）推移分析":
    patient_files = [f for f in glob.glob("*日別受付患者数*.csv") if "年齢別" not in f]
    if not patient_files: st.error("ファイルが見つかりません。"); st.stop()

    monthly_data_list, daily_data_list = [], []
    for f in patient_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str, month_str = match.group(1), match.group(2)
            for enc in ['utf-8-sig', 'cp932', 'shift-jis']:
                try:
                    df_p = pd.read_csv(f, encoding=enc)
                    df_p = df_p[df_p['日'].str.contains('日', na=False)].copy()
                    df_p['フル日付'] = f"{year_str}年{month_str}月" + df_p['日']
                    for col in df_p.columns:
                        if col not in ['日', 'フル日付']: df_p[col] = pd.to_numeric(df_p[col].astype(str).str.replace(',', ''), errors='coerce')
                    active_days = (df_p['延べ患者数 (人)'] > 0).sum()
                    if active_days == 0: continue
                    monthly_data_list.append({
                        '年月': f"{year_str}年{int(month_str):02d}月", 'ソート用': int(re.search(r'\d+', year_str).group()) * 100 + int(month_str),
                        '延べ患者数': df_p['延べ患者数 (人)'].sum(), '新患初診': df_p['新患初診 (人)'].sum(),
                        '再来初診': df_p.get('再来初診 (人)', pd.Series(dtype=float)).sum(), '再診': df_p.get('再診 (人)', pd.Series(dtype=float)).sum(),
                        '稼働日数': active_days
                    })
                    daily_data_list.append(df_p[['フル日付', '延べ患者数 (人)']])
                    break
                except: continue

    df_monthly = pd.DataFrame(monthly_data_list).sort_values('ソート用').drop_duplicates(subset=['年月']).reset_index(drop=True)
    df_daily = pd.concat(daily_data_list, ignore_index=True)
    df_monthly['新患率'] = df_monthly.apply(lambda row: round((row['新患初診'] / row['延べ患者数']) * 100, 1) if row['延べ患者数'] > 0 else 0, axis=1)
    df_monthly['再初診率'] = df_monthly.apply(lambda row: round((row['再来初診'] / (row['再来初診'] + row['再診'])) * 100, 1) if (row['再来初診'] + row['再診']) > 0 else 0, axis=1)
    df_monthly['一日平均'] = round(df_monthly['延べ患者数'] / df_monthly['稼働日数'], 1)
    df_monthly['累計_延べ患者数'] = df_monthly['延べ患者数'].cumsum()
    df_monthly['累計_新患初診'] = df_monthly['新患初診'].cumsum()

    max_daily_idx = df_daily['延べ患者数 (人)'].idxmax()
    max_avg_idx = df_monthly['一日平均'].idxmax()

    st.subheader("🏆 まえだ耳鼻咽喉科 過去最高記録")
    col_rec1, col_rec2 = st.columns(2)
    with col_rec1: st.info(f"**最高来院数（1日）:** \n\n ### {df_daily.loc[max_daily_idx, '延べ患者数 (人)']:,.0f} 人 \n\n 📅 記録日: {df_daily.loc[max_daily_idx, 'フル日付']}")
    with col_rec2: st.success(f"**最高平均来院数（月間）:** \n\n ### {df_monthly.loc[max_avg_idx, '一日平均']:,.1f} 人/日 \n\n 📅 記録月: {df_monthly.loc[max_avg_idx, '年月']}")

    st.write("---")
    st.subheader("📊 患者数トレンド分析（全期間）")
    fig1 = make_subplots(specs=[[{"secondary_y": True}]])
    fig1.add_trace(go.Bar(x=df_monthly['年月'], y=df_monthly['延べ患者数'], name="延べ患者数", marker_color="#3498DB"), secondary_y=False)
    fig1.add_trace(go.Scatter(x=df_monthly['年月'], y=df_monthly['新患率'], name="新患率(%)", mode='lines+markers', line=dict(color="#E74C3C", width=3)), secondary_y=True)
    fig1.update_layout(hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    fig1.update_yaxes(title_text="人数 (人)", secondary_y=False)
    fig1.update_yaxes(title_text="率 (%)", secondary_y=True, range=[0, max(df_monthly['新患率']) * 1.2], showgrid=False)
    st.plotly_chart(fig1, use_container_width=True, config={'displayModeBar': False})

    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=df_monthly['年月'], y=df_monthly['累計_延べ患者数'], name="累計 延べ患者数", fill='tozeroy', line=dict(color="#2980B9")))
    fig2.add_trace(go.Scatter(x=df_monthly['年月'], y=df_monthly['累計_新患初診'], name="累計 新患初診数", fill='tozeroy', line=dict(color="#1ABC9C")))
    fig2.update_layout(hovermode="x unified", yaxis_title="累計人数 (人)", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})

    fig3 = go.Figure()
    fig3.add_trace(go.Bar(x=df_monthly['年月'], y=df_monthly['一日平均'], name="一日平均", marker_color="#27AE60", text=df_monthly['一日平均'], texttemplate="%{text:.1f}", textposition='outside', textfont=dict(size=11)))
    fig3.update_layout(hovermode="x unified", yaxis_title="平均人数 (人/日)", yaxis=dict(range=[0, max(df_monthly['一日平均']) * 1.15]))
    st.plotly_chart(fig3, use_container_width=True, config={'displayModeBar': False})

    fig4 = make_subplots(specs=[[{"secondary_y": True}]])
    fig4.add_trace(go.Bar(x=df_monthly['年月'], y=df_monthly['再診'], name="再診", marker_color="#AED6F1"), secondary_y=False)
    fig4.add_trace(go.Bar(x=df_monthly['年月'], y=df_monthly['再来初診'], name="再来初診", marker_color="#5DADE2"), secondary_y=False)
    fig4.add_trace(go.Bar(x=df_monthly['年月'], y=df_monthly['新患初診'], name="新患初診", marker_color="#2E86C1"), secondary_y=False)
    fig4.add_trace(go.Scatter(x=df_monthly['年月'], y=df_monthly['再初診率'], name="再初診率(%)", mode='lines+markers', line=dict(color="#F39C12", width=3)), secondary_y=True)
    fig4.update_layout(barmode='stack', hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    fig4.update_yaxes(title_text="人数 (人)", secondary_y=False)
    fig4.update_yaxes(title_text="率 (%)", secondary_y=True, range=[0, max(df_monthly['再初診率']) * 1.2], showgrid=False)
    st.plotly_chart(fig4, use_container_width=True, config={'displayModeBar': False})

# ==========================================
# D. 年齢別構成比分析
# ==========================================
elif analysis_mode == "年齢別構成比分析":
    age_files = glob.glob("*受付患者数（年齢別）*.csv")
    if not age_files: st.error("ファイルが見つかりません。"); st.stop()

    data_list = []
    for f in age_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str, month_str = match.group(1) + "年", match.group(2) + "月"
            for enc in ['utf-8-sig', 'cp932', 'shift-jis']:
                try:
                    df_a = pd.read_csv(f, encoding=enc)
                    df_a = df_a[df_a['日'].str.contains('日', na=False)]
                    for col in df_a.columns:
                        if col != '日': df_a[col] = pd.to_numeric(df_a[col].astype(str).str.replace(',', ''), errors='coerce')
                    summary = df_a.drop(columns=['日']).sum()
                    summary['年度'], summary['月'] = year_str, month_str
                    data_list.append(summary)
                    break
                except: continue

    age_df = pd.DataFrame(data_list).drop_duplicates(subset=['年度', '月'])
    available_years = sorted(list(age_df['年度'].unique()), key=lambda x: int(re.search(r'\d+', x).group()))
    selected_year_age = st.selectbox("📅 表示する年度を選択", available_years, index=len(available_years)-1)
    
    df_curr_age_raw = age_df[age_df['年度'] == selected_year_age].copy()
    valid_months_age = [f"{i}月" for i in range(1, 13)]
    df_curr_age = pd.merge(pd.DataFrame({'月': valid_months_age}), df_curr_age_raw, on='月', how='left').fillna(0)

    st.subheader(f"📊 {selected_year_age} 年齢別構成比分析")
    t1_cols = ['～3歳', '3～6歳']
    t2_cols = ['～3歳', '3～6歳', '6～12歳', '12～20歳', '20～30歳', '30～50歳']
    curr_total_sum = df_curr_age['延べ患者数 (人)'].sum()

    col_kpi1, col_kpi2 = st.columns(2)
    with col_kpi1:
        if [c for c in t1_cols if c in df_curr_age.columns]:
            t1_sum = df_curr_age[[c for c in t1_cols if c in df_curr_age.columns]].sum().sum()
            t1_ratio = round((t1_sum / curr_total_sum) * 100, 1) if curr_total_sum > 0 else 0
            st.info(f"**👦 小児層（～6歳）** \n\n **累計人数:** {t1_sum:,.0f} 人　｜　**構成比:** {t1_ratio}%")
    with col_kpi2:
        if [c for c in t2_cols if c in df_curr_age.columns]:
            t2_sum = df_curr_age[[c for c in t2_cols if c in df_curr_age.columns]].sum().sum()
            t2_ratio = round((t2_sum / curr_total_sum) * 100, 1) if curr_total_sum > 0 else 0
            st.success(f"**👪 親子層（～50歳）** \n\n **累計人数:** {t2_sum:,.0f} 人　｜　**構成比:** {t2_ratio}%")

    st.write("---")
    ordered_age_categories = ['～3歳', '3～6歳', '6～12歳', '12～20歳', '20～30歳', '30～50歳', '50～60歳', '60～65歳', '65～75歳', '75歳～']
    age_categories = [c for c in ordered_age_categories if c in age_df.columns]

    col_chart1, col_chart2 = st.columns([2, 1])
    with col_chart1:
        fig_stack = go.Figure()
        for age_cat in age_categories:
            fig_stack.add_trace(go.Bar(x=df_curr_age['月'], y=df_curr_age[age_cat], name=age_cat))
        fig_stack.update_layout(barmode='stack', hovermode="x unified", xaxis=dict(type='category', categoryorder='array', categoryarray=valid_months_age))
        st.plotly_chart(fig_stack, use_container_width=True, config={'displayModeBar': False})

    with col_chart2:
        total_by_age = df_curr_age[age_categories].sum()
        total_by_age = total_by_age[total_by_age > 0]
        fig_pie = go.Figure(data=[go.Pie(labels=total_by_age.index, values=total_by_age.values, hole=.4, sort=False, direction='clockwise')])
        fig_pie.update_layout(showlegend=False, margin=dict(t=20, b=20, l=0, r=0))
        st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})

    st.write("---")
    st_df_age = df_curr_age[['月', '延べ患者数 (人)'] + age_categories].copy()
    sum_row_df = pd.DataFrame([st_df_age.sum(numeric_only=True)])
    sum_row_df['月'] = '総計'
    st_df_age = pd.concat([st_df_age, sum_row_df], ignore_index=True)
    st.dataframe(st_df_age.set_index('月').style.format("{:,.0f}").apply(lambda r: ['font-weight: bold; background-color: #f0f2f6'] * len(r) if r.name == '総計' else [''] * len(r), axis=1), use_container_width=True)

# ==========================================
# E. 診療行為一覧分析
# ==========================================
elif analysis_mode == "診療行為一覧分析":
    act_files = glob.glob("*診療行為一覧*.csv")
    if not act_files: st.error("ファイルが見つかりません。"); st.stop()

    data_list = []
    for f in act_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str, month_str = match.group(1) + "年", match.group(2) + "月"
            for enc in ['utf-8-sig', 'cp932', 'shift-jis']:
                try:
                    df_act = pd.read_csv(f, encoding=enc)
                    col_name = next((c for c in df_act.columns if '診療行為名称' in c.replace(' ', '')), None)
                    if not col_name: continue
                    df_act = df_act.dropna(subset=[col_name])
                    df_act = df_act[~df_act[col_name].astype(str).str.contains('合計')]
                    if '診療区分' in df_act.columns: df_act = df_act[~df_act['診療区分'].astype(str).isin(['検査', '保険外費用'])]
                    for col in ['回数', '総点数 (点)']:
                        if col in df_act.columns: df_act[col] = pd.to_numeric(df_act[col].astype(str).str.replace(r'[，,点円％%]', '', regex=True), errors='coerce').fillna(0)
                    summary = df_act.groupby(col_name)[['回数', '総点数 (点)']].sum().reset_index()
                    summary.rename(columns={col_name: '診療行為名称', '総点数 (点)': '総点数'}, inplace=True)
                    summary['年度'], summary['月'] = year_str, month_str
                    data_list.append(summary)
                    break
                except: continue

    act_df = pd.concat(data_list, ignore_index=True)
    available_years = sorted(list(act_df['年度'].unique()), key=lambda x: int(re.search(r'\d+', x).group()))
    selected_year_act = st.selectbox("📅 表示する年度を選択", available_years, index=len(available_years)-1, key="act_year")
    
    st.subheader(f"📊 {selected_year_act} 診療行為 分析 (処置・手術・医学管理等)")
    df_curr_act = act_df[act_df['年度'] == selected_year_act].copy()
    rank_df = df_curr_act.groupby('診療行為名称')[['回数', '総点数']].sum().reset_index()
    rank_df['単価'] = rank_df.apply(lambda x: round(x['総点数'] / x['回数'], 1) if x['回数'] > 0 else 0, axis=1)

    top10_df = rank_df.sort_values('総点数', ascending=False).head(10).reset_index(drop=True)
    top10_df.index = top10_df.index + 1 
    st.write("#### 🏆 年間 総点数 TOP10 ランキング")
    st.dataframe(top10_df[['診療行為名称', '単価', '回数', '総点数']].style.format({'単価': "{:,.1f} 点", '回数': "{:,.0f} 回", '総点数': "{:,.0f} 点"}), use_container_width=True)

    st.write("---")
    st.write("### 📈 診療行為別 月別推移")
    valid_months_act = [f"{i}月" for i in range(1, 13)]
    all_act_items = sorted(list(set(df_curr_act['診療行為名称'].unique().tolist())))
    if all_act_items:
        selected_act_item = st.selectbox("🔍 グラフ表示する診療行為を選択", all_act_items, index=all_act_items.index(top10_df.iloc[0]['診療行為名称']) if not top10_df.empty else 0)
        curr_item_data = df_curr_act[df_curr_act['診療行為名称'] == selected_act_item].groupby('月')['総点数'].sum().reset_index().rename(columns={'総点数': '当年'})
        plot_df_act = pd.merge(pd.DataFrame({'月': valid_months_act}), curr_item_data, on='月', how='left').fillna(0)

        fig_act = go.Figure()
        fig_act.add_trace(go.Scatter(x=plot_df_act['月'], y=plot_df_act['当年'], mode='lines+markers', name=f'当年', line=dict(color='#2E86C1', width=4)))
        fig_act.update_layout(xaxis_title="診療月", yaxis_title="総点数 (点)", hovermode="x unified")
        st.plotly_chart(fig_act, use_container_width=True, config={'displayModeBar': False})
        
    st.write("#### 📋 月別詳細テーブル（すべての診療行為・総点数）")
    matrix_full_df = df_curr_act.pivot_table(index='診療行為名称', columns='月', values='総点数', aggfunc='sum').fillna(0).reindex(columns=valid_months_act).fillna(0)
    matrix_full_df['年間合計'] = matrix_full_df.sum(axis=1)
    matrix_full_df = matrix_full_df.sort_values('年間合計', ascending=False)
    sum_row_act = matrix_full_df.sum(numeric_only=True)
    sum_row_act.name = '★点数合計'
    st.dataframe(pd.concat([matrix_full_df, pd.DataFrame([sum_row_act])]).style.format("{:,.0f}").apply(lambda r: ['font-weight: bold; background-color: #f0f2f6'] * len(r) if r.name == '★点数合計' else [''] * len(r), axis=1), use_container_width=True)

# ==========================================
# F. 検査一覧分析
# ==========================================
elif analysis_mode == "検査一覧分析":
    act_files = glob.glob("*診療行為一覧*.csv")
    if not act_files: st.error("ファイルが見つかりません。"); st.stop()

    data_list = []
    for f in act_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str, month_str = match.group(1) + "年", match.group(2) + "月"
            for enc in ['utf-8-sig', 'cp932', 'shift-jis']:
                try:
                    df_act = pd.read_csv(f, encoding=enc)
                    col_name = next((c for c in df_act.columns if '診療行為名称' in c.replace(' ', '')), None)
                    if not col_name: continue
                    df_act = df_act.dropna(subset=[col_name])
                    df_act = df_act[~df_act[col_name].astype(str).str.contains('合計')]
                    if '診療区分' in df_act.columns: df_act = df_act[df_act['診療区分'].astype(str).str.contains('検査')]
                    else: continue
                    for col in ['回数', '総点数 (点)']:
                        if col in df_act.columns: df_act[col] = pd.to_numeric(df_act[col].astype(str).str.replace(r'[，,点円％%]', '', regex=True), errors='coerce').fillna(0)
                    summary = df_act.groupby(col_name)[['回数', '総点数 (点)']].sum().reset_index()
                    summary.rename(columns={col_name: '診療行為名称', '総点数 (点)': '総点数'}, inplace=True)
                    summary['年度'], summary['月'] = year_str, month_str
                    data_list.append(summary)
                    break
                except: continue

    act_df = pd.concat(data_list, ignore_index=True)
    available_years = sorted(list(act_df['年度'].unique()), key=lambda x: int(re.search(r'\d+', x).group()))
    selected_year_act = st.selectbox("📅 表示する年度を選択", available_years, index=len(available_years)-1, key="ins_year")
    
    st.subheader(f"📊 {selected_year_act} 検査一覧 分析")
    df_curr_act = act_df[act_df['年度'] == selected_year_act].copy()

    st.write("#### 🎯 注力検査項目 実績")
    col_t1, col_t2, col_t3 = st.columns(3)
    def get_kpi_data(item_names):
        d = df_curr_act[df_curr_act['診療行為名称'].isin(item_names)]
        return d['回数'].sum(), d['総点数'].sum()

    c1_count, c1_pts = get_kpi_data(['チンパノメトリー'])
    with col_t1: st.markdown(f"<div class='kpi-container stInfo'><p><b>👂 チンパノメトリー</b></p><br><p><b>回数:</b> {c1_count:,.0f} 回 ｜ <b>総点数:</b> {c1_pts:,.0f} 点</p></div>", unsafe_allow_html=True)
        
    c2_count, c2_pts = get_kpi_data(['標準純音聴力検査', '簡易聴力検査（気導純音聴力）'])
    h1_c, h1_p = get_kpi_data(['標準純音聴力検査'])
    h2_c, h2_p = get_kpi_data(['簡易聴力検査（気導純音聴力）'])
    with col_t2: st.markdown(f"<div class='kpi-container stSuccess'><p><b>🎧 聴力検査（合算）</b></p><br><p><b>回数:</b> {c2_count:,.0f} 回 ｜ <b>総点数:</b> {c2_pts:,.0f} 点</p><hr><ul><li><b>標準純音</b>: {h1_c:,.0f}回 ｜ {h1_p:,.0f}点</li><li><b>簡易聴力</b>: {h2_c:,.0f}回 ｜ {h2_p:,.0f}点</li></ul></div>", unsafe_allow_html=True)
        
    c3_count, c3_pts = get_kpi_data(['ＥＦ－中耳', 'ＥＦ－喉頭', 'ＥＦ－嗅裂・鼻咽腔・副鼻腔'])
    f1_c, f1_p = get_kpi_data(['ＥＦ－中耳'])
    f2_c, f2_p = get_kpi_data(['ＥＦ－喉頭'])
    f3_c, f3_p = get_kpi_data(['ＥＦ－嗅裂・鼻咽腔・副鼻腔'])
    with col_t3: st.markdown(f"<div class='kpi-container stWarning'><p><b>🔬 ファイバー（合算）</b></p><br><p><b>回数:</b> {c3_count:,.0f} 回 ｜ <b>総点数:</b> {c3_pts:,.0f} 点</p><hr><ul><li><b>中耳</b>: {f1_c:,.0f}回 ｜ {f1_p:,.0f}点</li><li><b>喉頭</b>: {f2_c:,.0f}回 ｜ {f2_p:,.0f}点</li><li><b>鼻咽腔等</b>: {f3_c:,.0f}回 ｜ {f3_p:,.0f}点</li></ul></div>", unsafe_allow_html=True)

    st.write("---")
    rank_df = df_curr_act.groupby('診療行為名称')[['回数', '総点数']].sum().reset_index()
    rank_df['単価'] = rank_df.apply(lambda x: round(x['総点数'] / x['回数'], 1) if x['回数'] > 0 else 0, axis=1)
    top10_df = rank_df.sort_values('総点数', ascending=False).head(10).reset_index(drop=True)
    top10_df.index = top10_df.index + 1 
    
    st.write("#### 🏆 年間 総点数 TOP10 ランキング (検査)")
    st.dataframe(top10_df[['診療行為名称', '単価', '回数', '総点数']].style.format({'単価': "{:,.1f} 点", '回数': "{:,.0f} 回", '総点数': "{:,.0f} 点"}), use_container_width=True)

# ==========================================
# G. AI総合経営アドバイス
# ==========================================
elif analysis_mode == "AI総合経営アドバイス":
    
    NEWS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTS8MOjzU72nTwsnrZK54XoKkkk5DeI91kcgCkcx0D98YhQxFv8MrK4btU8ffp8KOVrhK-LsRextVKI/pub?output=csv"
    df_news = get_latest_news(NEWS_CSV_URL)

    if df_news is not None and not df_news.empty:
        if '日付' in df_news.columns:
            df_news['日付_date'] = pd.to_datetime(df_news['日付'], errors='coerce')
            df_news = df_news.sort_values('日付_date', ascending=False)
        
        st.markdown("<h3 style='margin-top:0px; margin-bottom:15px; color:#2C3E50;'>📰 最新の医療・改定ニュース（トップ3）</h3>", unsafe_allow_html=True)
        
        top3_news = df_news.head(3)
        for _, row in top3_news.iterrows():
            date = row.get('日付', '')
            source = row.get('カテゴリ', row.get('情報ソース', '情報局'))
            raw_content = str(row.get('ニュース内容', ''))
            
            url_match = re.search(r'(https?://[^\s]+)', raw_content)
            url = url_match.group(1) if url_match else ""
            title = raw_content.replace(url, '').strip() if url else raw_content
            
            link_html = f"<a href='{url}' target='_blank' style='text-decoration:none; color:#2E86C1;'>🔗 {title}</a>" if url else title

            news_html = f"""
            <div style='background-color:#fff; border-left: 5px solid #3498DB; padding: 12px 15px; margin-bottom: 12px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);'>
                <div style='font-size:0.85rem; color:#7F8C8D; margin-bottom:5px;'>📅 {date} ｜ 🏢 {source}</div>
                <div style='font-size:1rem; font-weight:bold; line-height:1.4;'>{link_html}</div>
            </div>
            """
            st.markdown(news_html, unsafe_allow_html=True)

        with st.expander("🔍 過去のニュースを検索・閲覧する（全件表示）"):
            search_query = st.text_input("キーワードで検索（例：ベースアップ、ガイドライン）")
            filtered_news = df_news
            if search_query:
                filtered_news = df_news[df_news['ニュース内容'].str.contains(search_query, na=False, case=False)]
            
            if not filtered_news.empty:
                st.write(f"該当件数: {len(filtered_news)}件")
                for _, row in filtered_news.iterrows():
                    date = row.get('日付', '')
                    source = row.get('カテゴリ', '')
                    raw_content = str(row.get('ニュース内容', ''))
                    url_match = re.search(r'(https?://[^\s]+)', raw_content)
                    url = url_match.group(1) if url_match else ""
                    title = raw_content.replace(url, '').strip() if url else raw_content
                    
                    st.markdown(f"- **{date}** [{source}] {'[' + title + '](' + url + ')' if url else title}")
            else:
                st.info("該当するニュースが見つかりませんでした。")
        
        st.write("---")

    st.subheader("🤖 AI総合経営アドバイス（診療報酬改定2026対応・上位コンサル版）")
    st.caption("※現在表示されているアドバイスは、当院の過去実績データに基づく専用コンサルティングレポートです。最新ニュースを自動解析して動的にアドバイスを生成するには、別途AI API（Gemini等）の連携が必要です。")
    
    with st.spinner("レセプトデータ・年齢構成比と、2026年改定答申データをクロス分析しています..."):
        rece_files = glob.glob("*レセプト*.csv")
        if not rece_files: st.error("データが見つかりません。"); st.stop()
        latest_year_str = sorted(list(set(re.search(r'(R\d+年)', f).group(1) for f in rece_files if re.search(r'(R\d+年)', f))))[-1]
        df_rece_ai = get_clean_df(latest_year_str)
        
        rece_tanka, rece_patients, latest_m_ai, kids_ratio, kids_cnt, total_age_cnt = 0, 0, "不明", 0, 0, 0
        
        if df_rece_ai is not None:
            target_month, latest_row, _, _ = get_latest_complete_month(df_rece_ai, latest_year_str)
            if target_month:
                rece_tanka = latest_row['レセ単価_num']
                latest_m_ai = target_month
                cnt_col = next((c for c in df_rece_ai.columns if '件数' in c or '枚数' in c), None)
                if cnt_col:
                    cnt_str = str(latest_row[cnt_col]).replace(',', '').replace('件', '').replace('枚', '').strip()
                    rece_patients = float(cnt_str) if cnt_str.replace('.', '', 1).isdigit() else 0

        try:
            target_y_num = int(re.search(r'\d+', latest_year_str).group())
            target_m_num = int(re.search(r'\d+', latest_m_ai).group())
            for f in glob.glob("*受付患者数（年齢別）*.csv"):
                f_norm = unicodedata.normalize('NFKC', f)
                match = re.search(r'(?:R|令和)?0*(\d+)年0*(\d+)月', f_norm)
                if match and int(match.group(1)) == target_y_num and int(match.group(2)) == target_m_num:
                    for enc in ['utf-8-sig', 'cp932', 'shift-jis']:
                        try:
                            df_age = pd.read_csv(f, encoding=enc)
                            df_age = df_age[df_age['日'].str.contains('日', na=False)] 
                            for col in df_age.columns:
                                if col != '日': df_age[col] = pd.to_numeric(df_age[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                            sum_age = df_age.drop(columns=['日', '延べ患者数 (人)', 'フル日付'], errors='ignore').sum()
                            t1_cols = ['～3歳', '3～6歳', '6～12歳']
                            valid_t1 = [c for c in t1_cols if c in sum_age.index]
                            if valid_t1:
                                kids_cnt = sum_age[valid_t1].sum()
                                total_age_cnt = sum_age.sum()
                                kids_ratio = (kids_cnt / total_age_cnt * 100) if total_age_cnt > 0 else 0
                            break
                        except: continue
                    break
        except: pass

        gap_ai = 750 - rece_tanka

    st.markdown(f"<div class='target-box' style='background-color:#F8F9F9; border-color:#D5D8DC;'><h3>📊 現在地確認 ({latest_year_str}{latest_m_ai})</h3><p style='font-size:1.1rem;'>最新のレセプト単価： <b>{rece_tanka:,.0f} 点</b> <span style='color:{'#E74C3C' if gap_ai>0 else '#27AE60'};'>（目標まで {'あと '+str(int(gap_ai))+' 点' if gap_ai>0 else 'クリア！'}）</span><br>小児患者（12歳以下）割合： <b style='color:#E74C3C; font-size:1.3rem;'>約 {kids_ratio:.1f} %</b> （月間約 {kids_cnt:,.0f} 人）</p></div>", unsafe_allow_html=True)

    st.markdown("<h3 style='margin-top:30px;'>📑 【超具体化】当院の実績ベース改定対応アクション</h3>", unsafe_allow_html=True)
    
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        st.markdown(f"""
        <div class='revision-box'>
            <h4>💊 小児抗菌薬適正使用支援加算の徹底</h4>
            <p style='font-size:0.95rem; line-height:1.6;'>
                当院は小児割合が<b>{kids_ratio:.1f}%</b>と極めて高く、急性中耳炎・副鼻腔炎の来院が多いという強みがあります。2026年改定では小児への抗菌薬適正使用がさらに評価されます。<br>
                <b>【具体策】</b>現在、月間約{kids_cnt:,.0f}人の小児が来院していますが、対象疾患（急性中耳炎等）の患者の半数（約{kids_cnt*0.5:,.0f}人）にガイドラインに沿った文書説明を行い算定（80点）した場合、それだけで<b>月額約{(kids_cnt*0.5*80*10):,.0f}円の増収</b>となります。受付での「説明文書配布フロー」をマニュアル化してください。
            </p>
        </div>
        """, unsafe_allow_html=True)

    with col_r2:
        st.markdown(f"""
        <div class='revision-box' style='border-color:#5DADE2; background-color:#EBF5FB;'>
            <h4>📋 「小児特定疾患カウンセリング料」等の算定漏れ防止</h4>
            <p style='font-size:0.95rem; line-height:1.6;'>
                ベースアップ評価料等の基礎対応が完了している当院における次のレバーは<b>「特定疾患の管理料」</b>です。小児のアレルギー性鼻炎（舌下免疫療法含む）に対する継続管理が重要です。<br>
                <b>【具体策】</b>毎月の処方と同時に、対象となる小児患者への療養指導が算定漏れしていないかレセコンのチェックルールを見直してください。例えば対象患者100人に算定漏れが発覚した場合、<b>年間で数十万円単位の機会損失</b>を防げます。
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.write("---")
    st.markdown("#### 🤖 コンサルタント総評")
    st.success(f"**マイナ保険証利用率7割達成**、および**ベースアップ評価料の早期導入・維持**、非常に素晴らしい経営手腕です。この水準のクリニックにおいては、もはや『DX化』や『賃上げ要件』といったマクロな対策にリソースを割く必要はありません。今後は当院の最大の武器である**「月間 {kids_cnt:,.0f} 人の小児・ファミリー層」に対する医学管理料や指導料のニッチな算定要件**（文書提供、情報通信機器の活用など）にフォーカスし、750点達成に向けた『チリツモ』の収益を確実に取りこぼさない体制を築くことが最優先事項です。")