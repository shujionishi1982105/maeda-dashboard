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

# --- 共通CSS設定 ---
st.markdown("""
    <style>
    /* テーブルの「月」列の幅固定 */
    [data-testid="stTable"] th:first-child,
    [data-grid-column-id="月"],
    .stDataFrame div[data-testid="stTable"] div:first-child {
        min-width: 100px !important;
        max-width: 100px !important;
        width: 100px !important;
    }
    .record-text {
        font-size: 1.1rem;
        font-weight: bold;
        color: #E74C3C;
    }

    /* === セレクトボックスのカーソルを指マークにする === */
    div[data-baseweb="select"],
    div[data-baseweb="select"] * {
        cursor: pointer !important;
    }
    
    /* === ナビゲーションボタンの確実なカスタムデザイン === */
    div.stButton > button {
        height: 65px !important;
        border-radius: 10px !important;
        font-weight: bold !important;
        font-size: 15px !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1) !important;
        transition: all 0.2s ease-in-out !important;
        white-space: normal !important;
        line-height: 1.3 !important;
    }
    
    /* 【未選択のボタン (secondary)】薄い青背景 */
    div.stButton > button[kind="secondary"],
    div.stButton > button[data-testid="baseButton-secondary"] {
        background-color: #EBF5FB !important;
        border: 1px solid #AED6F1 !important;
    }
    div.stButton > button[kind="secondary"] p,
    div.stButton > button[data-testid="baseButton-secondary"] p,
    div.stButton > button[kind="secondary"] div,
    div.stButton > button[data-testid="baseButton-secondary"] div {
        color: #154360 !important;
    }

    /* 【選択中のボタン (primary)】赤背景 */
    div.stButton > button[kind="primary"],
    div.stButton > button[data-testid="baseButton-primary"] {
        background-color: #E74C3C !important;
        border: 1px solid #E74C3C !important;
    }
    div.stButton > button[kind="primary"] p,
    div.stButton > button[data-testid="baseButton-primary"] p,
    div.stButton > button[kind="primary"] div,
    div.stButton > button[data-testid="baseButton-primary"] div {
        color: #FFFFFF !important;
    }

    /* マウスホバー時 */
    div.stButton > button:hover {
        transform: translateY(-2px) !important;
        filter: brightness(0.95) !important;
    }

    /* === ヘッダーのロゴとタイトルの高さ合わせ === */
    .header-container {
        display: flex;
        align-items: center; 
        gap: 20px; 
        margin-bottom: 20px;
    }
    [data-testid="stImage"] {
        margin-top: 10px !important; 
    }
    .header-title {
        margin: 0 !important; 
        padding-top: 15px !important; 
    }

    /* === KPI枠・AI分析枠のCSS === */
    .kpi-container {
        height: 250px; 
        overflow-y: auto; 
        margin-bottom: 0px !important; 
        padding-bottom: 0px !important;
        border-radius: 0.5rem; 
        padding: 1rem; 
        background-color: transparent; 
    }
    .kpi-container.stInfo { border: 1px solid rgba(25, 230, 255, 0.2); background-color: rgba(25, 230, 255, 0.1); }
    .kpi-container.stSuccess { border: 1px solid rgba(0, 255, 128, 0.2); background-color: rgba(0, 255, 128, 0.1); }
    .kpi-container.stWarning { border: 1px solid rgba(255, 204, 0, 0.2); background-color: rgba(255, 204, 0, 0.1); }
    .kpi-container p, .kpi-container ul { margin: 0; padding: 0; list-style-type: none; }
    .kpi-container hr { margin: 10px 0; border: 0; border-top: 1px solid rgba(49, 51, 63, 0.1); }
    
    .target-box {
        border: 2px solid #AED6F1;
        background-color: #EBF5FB;
        border-radius: 10px;
        padding: 15px;
        min-height: 220px;
    }
    .ai-box {
        border: 2px solid #D5D8DC;
        background-color: #F8F9F9;
        border-radius: 10px;
        padding: 15px;
        min-height: 220px;
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 共通関数 (全ページで利用)
# ==========================================
def get_clean_df(year_str):
    files = glob.glob(f"*{year_str}*レセプト*.csv")
    if not files: return None
    # 修正：すべての文字コードを試し、正しく「月」列が読めたものだけを採用する絶対安全な仕組み
    for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
        try:
            df = pd.read_csv(files[0], encoding=enc)
            df.columns = [re.sub(r'\s*[\(（].*?[\)）]', '', str(c)).strip() for c in df.columns]
            if '月' in df.columns:
                return df
        except: continue
    return None

def get_latest_complete_month(df_rece, year_str):
    valid_months = [f"{i}月" for i in range(1, 13)]
    df_tmp = df_rece[df_rece['月'].isin(valid_months)].copy()
    df_tmp['レセ単価_num'] = pd.to_numeric(df_tmp['レセ単価'].astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce').fillna(0)
    active_rows = df_tmp[df_tmp['レセ単価_num'] > 0]
    
    if active_rows.empty:
        return None, None, False, None
        
    original_latest_month = active_rows.iloc[-1]['月']
    
    try:
        target_year_num = int(re.search(r'\d+', year_str).group())
    except:
        return original_latest_month, active_rows.iloc[-1], False, original_latest_month
    
    for idx in range(len(active_rows) - 1, -1, -1):
        tmp_month = active_rows.iloc[idx]['月']
        target_month_num = int(re.search(r'\d+', tmp_month).group())
        
        act_exists = False
        for f in glob.glob("*診療行為一覧*.csv"):
            f_norm = unicodedata.normalize('NFKC', f)
            match = re.search(r'(?:R|令和)?0*(\d+)年0*(\d+)月', f_norm)
            if match:
                if int(match.group(1)) == target_year_num and int(match.group(2)) == target_month_num:
                    act_exists = True
                    break
        
        if act_exists:
            is_fallback = (tmp_month != original_latest_month)
            return tmp_month, active_rows.iloc[idx], is_fallback, original_latest_month
            
    return original_latest_month, active_rows.iloc[-1], False, original_latest_month

def get_act_summary_for_ai(y_str, m_str):
    try:
        target_y = int(re.search(r'\d+', y_str).group())
        target_m = int(re.search(r'\d+', m_str).group())
    except:
        return pd.DataFrame()
        
    act_files = glob.glob("*診療行為一覧*.csv")
    for f in act_files:
        f_norm = unicodedata.normalize('NFKC', f)
        match = re.search(r'(?:R|令和)?0*(\d+)年0*(\d+)月', f_norm)
        if match:
            if int(match.group(1)) == target_y and int(match.group(2)) == target_m:
                for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
                    try:
                        df_act = pd.read_csv(f, encoding=enc)
                        col_name = None
                        if '診 療 行 為 名 称' in df_act.columns: col_name = '診 療 行 為 名 称'
                        elif '診療行為名称' in df_act.columns: col_name = '診療行為名称'
                        if not col_name: continue # 正しく列名が読めなければ次の文字コードへ
                            
                        df_act = df_act.dropna(subset=[col_name])
                        df_act = df_act[~df_act[col_name].astype(str).str.contains('合計')]
                        
                        col_cat = '診療区分' if '診療区分' in df_act.columns else None
                        
                        for col in ['回数', '総点数 (点)']:
                            if col in df_act.columns:
                                df_act[col] = pd.to_numeric(df_act[col].astype(str).str.replace(r'[，,点円％%]', '', regex=True), errors='coerce').fillna(0)
                        
                        if col_cat:
                            summary = df_act.groupby([col_cat, col_name])[['回数', '総点数 (点)']].sum().reset_index()
                            return summary.rename(columns={col_name: '名称', col_cat: '区分', '総点数 (点)': '総点数'})
                        else:
                            summary = df_act.groupby(col_name)[['回数', '総点数 (点)']].sum().reset_index()
                            summary['区分'] = '不明'
                            return summary.rename(columns={col_name: '名称', '総点数 (点)': '総点数'})
                    except Exception:
                        continue
    return pd.DataFrame()

# ==========================================
# タイトル＆ロゴ表示 
# ==========================================
try:
    base_dir = os.path.dirname(os.path.abspath(__file__))
except NameError:
    base_dir = os.getcwd()

logo_path_png = os.path.join(base_dir, "logo.png")
logo_path_jpg = os.path.join(base_dir, "logo.jpg")

header_cols = st.columns([1, 15])
with header_cols[0]:
    if os.path.exists(logo_path_png):
        st.image(logo_path_png, use_container_width=True)
    elif os.path.exists(logo_path_jpg):
        st.image(logo_path_jpg, use_container_width=True)
    elif os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    else:
        st.markdown("<h1 style='text-align: center; margin: 0; padding: 0;'>🏥</h1>", unsafe_allow_html=True)

with header_cols[1]:
    st.markdown(f'<h1 class="header-title">まえだ耳鼻咽喉科 経営分析ダッシュボード</h1>', unsafe_allow_html=True)

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

analysis_mode = st.session_state.current_page


# ==========================================
# A. レセプト分析
# ==========================================
if analysis_mode == "レセプト分析":
    files_all = glob.glob("*レセプト*.csv")
    years_found = sorted(list(set(re.search(r'(R\d+年)', f).group(1) for f in files_all if re.search(r'(R\d+年)', f))))

    if not years_found:
        st.error("レセプトのCSVファイルが見つかりません。")
        st.stop()

    kpi_placeholder = st.container()
    st.write("<br>", unsafe_allow_html=True) 

    col_year, col_item = st.columns(2)
    with col_year:
        selected_year = st.selectbox("📅 表示する年度を選択してください", years_found, index=len(years_found)-1)

    current_num = int(re.search(r'\d+', selected_year).group())
    prev_year = f"R{current_num - 1}年"

    df_now_full = get_clean_df(selected_year)
    df_prev_full = get_clean_df(prev_year)

    valid_months = [f"{i}月" for i in range(1, 13)]

    if df_now_full is not None:
        all_cols = df_now_full.columns.tolist()
        forbidden = ["月", "総", "前年", "比", "枚数"]
        choices = [c for c in all_cols if df_now_full[c].dtype in ['float64', 'int64'] 
                   and not any(k in c for k in forbidden)]

        with col_item:
            selected_item = st.selectbox("🔍 グラフで推移を確認したい項目を選んでください", options=choices, 
                                        index=choices.index("レセ単価") if "レセ単価" in choices else 0)

    # ----------------------------------------------------
    # 【プレースホルダー内に描画】レセプト単価 目標＆AI分析 KPIブロック
    # ----------------------------------------------------
    with kpi_placeholder:
        if selected_year == "R8年" and df_now_full is not None and 'レセ単価' in df_now_full.columns:
            TARGET_PRICE = 750
            
            target_month, latest_row, is_fallback, original_latest_month = get_latest_complete_month(df_now_full, selected_year)
            
            if target_month:
                latest_month = target_month
                price_curr = latest_row['レセ単価_num']
                
                latest_m_num = int(latest_month.replace('月', ''))
                
                df_tmp = df_now_full[df_now_full['月'].isin(valid_months)].copy()
                
                if latest_m_num == 1:
                    prev_month = "12月"
                    if df_prev_full is not None and 'レセ単価' in df_prev_full.columns:
                        prev_val = df_prev_full.loc[df_prev_full['月'] == prev_month, 'レセ単価'].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                        price_prev = float(prev_val.values[0]) if len(prev_val) > 0 and prev_val.values[0] != '' else 0
                    else:
                        price_prev = 0
                else:
                    prev_month = f"{latest_m_num - 1}月"
                    df_tmp['レセ単価_num'] = pd.to_numeric(df_tmp['レセ単価'].astype(str).str.replace(r'[^\d.-]', '', regex=True), errors='coerce').fillna(0)
                    prev_val = df_tmp.loc[df_tmp['月'] == prev_month, 'レセ単価_num']
                    price_prev = prev_val.values[0] if len(prev_val) > 0 else 0
                    
                gap = TARGET_PRICE - price_curr
                if gap > 0:
                    gap_html = f"<span style='color:#E74C3C; font-size:1.2rem;'><b>{gap:,.0f} 点 不足</b></span>"
                else:
                    gap_html = f"<span style='color:#27AE60; font-size:1.2rem;'><b>目標達成！ (+{abs(gap):,.0f} 点)</b></span>"

                fallback_notice = ""
                if is_fallback:
                    fallback_notice = f"<div style='background-color:#FFF3CD; color:#856404; padding:5px 10px; border-radius:5px; font-size:0.85em; margin-bottom:10px;'>⚠️ {original_latest_month}の診療データが未確定のため、データが揃っている<b>【{latest_month}】</b>を基準に表示しています。</div>"

                if price_prev == 0:
                    ai_comment_html = "<p style='margin-top:10px;'>前月のデータが存在しないため、比較分析はスキップしました。</p>"
                elif price_curr == price_prev:
                    ai_comment_html = f"<p style='margin-top:10px;'>【{latest_month}】のレセ単価は前月({prev_month})と同水準（<b>{price_curr:,.0f} 点</b>）を維持しています。</p>"
                else:
                    diff_price = price_curr - price_prev
                    trend = "低下" if diff_price < 0 else "上昇"
                    trend_color = "#E74C3C" if diff_price < 0 else "#2E86C1"
                    reason = "減少" if diff_price < 0 else "増加"
                    
                    df_a_curr = get_act_summary_for_ai(selected_year, latest_month)
                    prev_y_str_act = f"R{int(re.search(r'\d+', selected_year).group()) - 1}年" if latest_m_num == 1 else selected_year
                    df_a_prev = get_act_summary_for_ai(prev_y_str_act, prev_month)
                    
                    if df_a_curr.empty or df_a_prev.empty:
                        exclude_cols = ["月", "総", "前年", "比", "枚数", "件数", "日数", "レセ単価", "日当点", "平均", "合計", "レセ単価_num"]
                        comp_cols = [c for c in df_now_full.columns if not any(k in c for k in exclude_cols)]
                        diffs = []
                        try:
                            row_c = df_tmp[df_tmp['月'] == latest_month].iloc[0]
                            row_p = None
                            if latest_m_num == 1 and df_prev_full is not None:
                                prev_rows = df_prev_full[df_prev_full['月'] == prev_month]
                                if not prev_rows.empty: row_p = prev_rows.iloc[0]
                            else:
                                prev_rows = df_tmp[df_tmp['月'] == prev_month]
                                if not prev_rows.empty: row_p = prev_rows.iloc[0]
                                    
                            if row_p is not None:
                                for col in comp_cols:
                                    if col in row_c and col in row_p:
                                        v_c_str = str(row_c[col]).replace(',', '').replace('点', '').replace('円', '').strip()
                                        v_p_str = str(row_p[col]).replace(',', '').replace('点', '').replace('円', '').strip()
                                        v_c = float(v_c_str) if v_c_str.replace('.', '', 1).isdigit() else 0.0
                                        v_p = float(v_p_str) if v_p_str.replace('.', '', 1).isdigit() else 0.0
                                        diffs.append({'名称': col, '点数差': v_c - v_p})
                                df_diff = pd.DataFrame(diffs)
                            else: df_diff = pd.DataFrame()
                        except: df_diff = pd.DataFrame()
                        
                        if df_diff.empty:
                            ai_comment_html = "<p style='margin-bottom:8px;'>" + f"【{latest_month}】のレセ単価は前月({prev_month})と比較して <b style='color:{trend_color};'>{abs(diff_price):,.0f} 点 {trend}</b> しています。" + "</p><p style='color:#E74C3C; font-size:0.9em; margin-top:10px;'>※各項目の詳細データが抽出できなかったため、要因分析はスキップしました。</p>"
                        else:
                            top_diffs = df_diff[df_diff['点数差'] < 0].sort_values('点数差', ascending=True).head(3) if diff_price < 0 else df_diff[df_diff['点数差'] > 0].sort_values('点数差', ascending=False).head(3)
                            details = ""
                            for _, row in top_diffs.iterrows():
                                if row['点数差'] != 0:
                                    sign_p = "+" if row['点数差'] > 0 else ""
                                    details += f"<li style='margin-bottom:3px;'>・<b>{row['名称']}</b> <span style='font-size:0.9em; color:#555;'>(全体で {sign_p}{row['点数差']:,.0f}点)</span></li>"
                            ai_comment_html = "<p style='margin-bottom:8px;'>" + f"【{latest_month}】のレセ単価は前月({prev_month})と比較して <b style='color:{trend_color};'>{abs(diff_price):,.0f} 点 {trend}</b> しています。" + "</p>" + f"<p style='font-size:0.95em; margin-bottom:5px;'>レセプトデータを分析した結果、主に以下の項目の<b>{reason}</b>が影響していると考えられます。</p>" + "<ul style='line-height:1.4; padding-left:10px; list-style-type:none;'>" + f"{details}" + "</ul>"
                    else:
                        df_diff = pd.merge(df_a_curr, df_a_prev, on=['区分', '名称'], how='outer', suffixes=('_curr', '_prev')).fillna(0)
                        df_diff['点数差'] = df_diff['総点数_curr'] - df_diff['総点数_prev']
                        
                        top_diffs = df_diff[df_diff['点数差'] < 0].sort_values('点数差', ascending=True).head(3) if diff_price < 0 else df_diff[df_diff['点数差'] > 0].sort_values('点数差', ascending=False).head(3)
                        
                        details = ""
                        for _, row in top_diffs.iterrows():
                            if row['点数差'] != 0:
                                sign_p = "+" if row['点数差'] > 0 else ""
                                
                                cat_raw = str(row['区分'])
                                cat_clean = re.sub(r'^\d+\s*', '', cat_raw)
                                item_name = str(row['名称'])
                                
                                if '小児科外来' in item_name or '指導' in cat_clean or '管理' in cat_clean: cat_clean = '管理'
                                elif '手術' in cat_clean: cat_clean = '手術'
                                elif '検査' in cat_clean: cat_clean = '検査'
                                elif '処置' in cat_clean: cat_clean = '処置'
                                elif '投薬' in cat_clean or '処方' in cat_clean: cat_clean = '投薬'
                                elif '初診' in cat_clean or '再診' in cat_clean or '診察' in cat_clean: cat_clean = '診察'
                                elif '画像' in cat_clean or 'レントゲン' in cat_clean: cat_clean = '画像'
                                
                                details += f"<li style='margin-bottom:3px;'><b>【{cat_clean}】</b>{item_name} <span style='font-size:0.9em; color:#555;'>({sign_p}{row['点数差']:,.0f}点)</span></li>"
                        
                        ai_comment_html = "<p style='margin-bottom:8px;'><span style='background-color:#EBF5FB; padding:2px 6px; border-radius:4px; font-size:0.85em; color:#2E86C1;'>ℹ️ 診療行為詳細データと同期</span><br>" + f"【{latest_month}】のレセ単価は前月({prev_month})と比較して <b style='color:{trend_color};'>{abs(diff_price):,.0f} 点 {trend}</b> しています。</p>" + f"<p style='font-size:0.95em; margin-bottom:5px;'>詳細データを分析した結果、主に以下の具体的な診療項目の<b>{reason}</b>が影響していると考えられます。</p>" + "<ul style='line-height:1.4; padding-left:10px; list-style-type:none;'>" + f"{details}" + "</ul>"

                col_kpi, col_ai = st.columns(2)
                with col_kpi:
                    kpi_block = "<div class='target-box'>" + f"{fallback_notice}" + "<h4 style='margin-top:0; margin-bottom:10px; color:#2C3E50; font-size:1.1rem; border-bottom:1px dashed #ccc; padding-bottom:5px;'>🎯 今年の目標レセプト単価</h4>" + "<div style='font-size: 2.8rem; font-weight: bold; text-align: center; margin: 15px 0; color: #2E86C1;'>750 <span style='font-size: 1.2rem; color:#333;'>点</span></div>" + "<hr style='margin: 10px 0;'>" + f"<p style='margin-bottom:5px;'><b>【最新実績】</b> {selected_year}{latest_month}： <b style='font-size:1.2rem;'>{price_curr:,.0f} 点</b></p>" + f"<p style='margin-bottom:0;'><b>【進捗】</b> {gap_html}</p>" + "</div>"
                    st.markdown(kpi_block, unsafe_allow_html=True)
                    
                with col_ai:
                    ai_block = "<div class='ai-box'>" + "<h4 style='margin-top:0; margin-bottom:10px; color:#2C3E50; font-size:1.1rem; border-bottom:1px dashed #ccc; padding-bottom:5px;'>🤖 AI分析レポート</h4>" + f"{ai_comment_html}" + "</div>"
                    st.markdown(ai_block, unsafe_allow_html=True)
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
            plot_df['比率'] = plot_df.apply(
                lambda x: round((x['当年'] / x['前年'] * 100), 1) if x['前年'] > 0 else 0, axis=1
            )
        else:
            plot_df['前年'] = 0
            plot_df['比率'] = 0

        colors = ['#E74C3C' if 0 < val < 100 else ('#2E86C1' if val >= 100 else '#000000') for val in plot_df['比率']]

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=plot_df['月'], y=plot_df['当年'],
            mode='lines+markers+text', name='当年',
            line=dict(color='#2E86C1', width=4),
            text=plot_df['比率'].apply(lambda x: f"{x}%" if x > 0 else ""),
            textposition="top center", textfont=dict(color=colors, size=13, family="Arial Black"),
            hovertemplate="<b>%{x}</b><br>当年: %{y:,.0f}<extra></extra>"
        ))
        if df_prev is not None:
            fig.add_trace(go.Scatter(
                x=plot_df['月'], y=plot_df['前年'],
                mode='lines+markers', name='前年',
                line=dict(color='#ABB2B9', width=2, dash='dot'),
                hovertemplate="前年: %{y:,.0f}<extra></extra>"
            ))
        fig.update_layout(
            xaxis=dict(title="診療月", type='category', categoryorder='array', categoryarray=valid_months),
            yaxis_title="点数 / 件数", hovermode="x",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

        st.write("---")
        st.write("### 📋 詳細数値データ（年間一覧）")

        display_df = df_now_full.copy()
        display_df['月'] = display_df['月'].replace('総数', '総数・平均')
        
        prev_total_row = None
        if df_prev_full is not None:
            p_row = df_prev_full[df_prev_full['月'] == '総数']
            if not p_row.empty:
                prev_total_row = p_row.iloc[0]

        def make_styled_df(df):
            styled = df.set_index('月').style
            fmt = {}
            for col in df.columns:
                if col == '月': continue
                if "比" in col: fmt[col] = "{:.1f}%"
                else: fmt[col] = "{:,.0f}"
            styled = styled.format(fmt)

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
    if not income_files:
        st.error("外来収入金額のCSVファイルが見つかりません。")
        st.stop()

    data_list = []
    for f in income_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str = match.group(1) + "年"
            month_str = match.group(2) + "月"
            for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
                try:
                    df_m = pd.read_csv(f, encoding=enc)
                    if '日' not in df_m.columns:
                        continue # 列名が正しく読めなければ次へ
                    df_m = df_m[df_m['日'].str.contains('日', na=False)]
                    for col in df_m.columns:
                        if col != '日':
                            df_m[col] = pd.to_numeric(df_m[col].astype(str).str.replace(',', ''), errors='coerce')
                    
                    summary = df_m.drop(columns=['日']).sum()
                    summary['年度'] = year_str
                    summary['月'] = month_str
                    data_list.append(summary)
                    break
                except: continue

    if not data_list:
        st.error("有効なデータを読み込めませんでした。")
        st.stop()

    income_df = pd.DataFrame(data_list).drop_duplicates(subset=['年度', '月'])
    available_years = sorted(list(income_df['年度'].unique()), key=lambda x: int(re.search(r'\d+', x).group()))
    
    col_year_inc, col_item_inc = st.columns(2)
    with col_year_inc:
        selected_year_inc = st.selectbox("📅 表示する年度を選択してください", available_years, index=len(available_years)-1, key="inc_year")
    
    current_num_inc = int(re.search(r'\d+', selected_year_inc).group())
    prev_year_inc = f"R{current_num_inc - 1}年"
    
    items_inc = [c for c in income_df.columns if c not in ['年度', '月']]
    with col_item_inc:
        selected_income = st.selectbox("🔍 分析項目を選択してください", items_inc, 
                                       index=items_inc.index("外来収入金額 (円)") if "外来収入金額 (円)" in items_inc else 0, key="inc_item")

    st.subheader(f"📊 {selected_year_inc} 【{selected_income}】の推移分析")

    valid_months_inc = [f"{i}月" for i in range(1, 13)]
    df_curr_inc = income_df[income_df['年度'] == selected_year_inc].copy()
    df_prev_inc = income_df[income_df['年度'] == prev_year_inc].copy()
    
    plot_df_inc = pd.DataFrame({'月': valid_months_inc})
    plot_df_inc = pd.merge(plot_df_inc, df_curr_inc[['月', selected_income]], on='月', how='left').rename(columns={selected_income: '当年'})
    plot_df_inc = pd.merge(plot_df_inc, df_prev_inc[['月', selected_income]], on='月', how='left').rename(columns={selected_income: '前年'})
    plot_df_inc = plot_df_inc.fillna(0)
    plot_df_inc['前年比'] = plot_df_inc.apply(lambda x: round(x['当年']/x['前年']*100, 1) if x['前年'] > 0 else 0, axis=1)

    fig_i = go.Figure()
    if not df_prev_inc.empty:
        fig_i.add_trace(go.Bar(
            x=plot_df_inc['月'], y=plot_df_inc['前年'], name=f'前年 ({prev_year_inc})', marker_color='#ABB2B9',
            text=plot_df_inc['前年'].apply(lambda x: f"{x/10000:.0f}万" if x >= 10000 else (f"{x:,.0f}" if x > 0 else "")),
            textposition='outside', textangle=0, textfont=dict(size=11), hovertemplate="前年: %{y:,.0f}<extra></extra>"
        ))
        
    fig_i.add_trace(go.Bar(
        x=plot_df_inc['月'], y=plot_df_inc['当年'], name=f'当年 ({selected_year_inc})', marker_color='#2E86C1',
        text=plot_df_inc['当年'].apply(lambda x: f"{x/10000:.0f}万" if x >= 10000 else (f"{x:,.0f}" if x > 0 else "")),
        textposition='outside', textangle=0, textfont=dict(size=11), hovertemplate="当年: %{y:,.0f}<extra></extra>"
    ))
    
    max_val = max(plot_df_inc['当年'].max(), plot_df_inc['前年'].max())
    fig_i.update_layout(
        barmode='group', xaxis_title="診療月", yaxis_title=selected_income, yaxis=dict(range=[0, max_val * 1.15]),
        hovermode="x unified", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(type='category', categoryorder='array', categoryarray=valid_months_inc)
    )
    st.plotly_chart(fig_i, use_container_width=True, config={'displayModeBar': False})

    st.write("---")
    st.write("### 📋 詳細数値データ（年間一覧）")
    
    st_df_inc = plot_df_inc.copy()
    sum_curr = st_df_inc['当年'].sum()
    sum_prev = st_df_inc['前年'].sum()
    sum_ratio = round(sum_curr / sum_prev * 100, 1) if sum_prev > 0 else 0
    
    sum_row = pd.DataFrame([{'月': '総数・平均', '当年': sum_curr, '前年': sum_prev, '前年比': sum_ratio}])
    st_df_inc = pd.concat([st_df_inc, sum_row], ignore_index=True)
    
    def style_income_table(df):
        styler = df.set_index('月').style
        fmt = {'当年': "{:,.0f}", '前年': "{:,.0f}", '前年比': "{:.1f}%"}
        styler = styler.format(fmt)
        
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
# C. 受付患者数推移分析 (初再診別)
# ==========================================
elif analysis_mode == "受付患者数（初再診別）推移分析":
    patient_files = glob.glob("*日別受付患者数*.csv")
    patient_files = [f for f in patient_files if "年齢別" not in f]
    
    if not patient_files:
        st.error("受付患者数のCSVファイルが見つかりません。")
        st.stop()

    monthly_data_list = []
    daily_data_list = []

    for f in patient_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str = match.group(1)
            month_str = match.group(2)
            
            for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
                try:
                    df_p = pd.read_csv(f, encoding=enc)
                    if '日' not in df_p.columns:
                        continue # 列名が正しく読めなければ次へ
                    df_p = df_p[df_p['日'].str.contains('日', na=False)].copy()
                    
                    df_p['フル日付'] = f"{year_str}年{month_str}月" + df_p['日']
                    
                    for col in df_p.columns:
                        if col not in ['日', 'フル日付']:
                            df_p[col] = pd.to_numeric(df_p[col].astype(str).str.replace(',', ''), errors='coerce')
                    
                    active_days = (df_p['延べ患者数 (人)'] > 0).sum()
                    if active_days == 0: continue
                    
                    sum_nobe = df_p['延べ患者数 (人)'].sum()
                    sum_shinkan = df_p['新患初診 (人)'].sum()
                    sum_sairai = df_p.get('再来初診 (人)', pd.Series(dtype=float)).sum()
                    sum_saishin = df_p.get('再診 (人)', pd.Series(dtype=float)).sum()
                    
                    monthly_data_list.append({
                        '年月': f"{year_str}年{int(month_str):02d}月",
                        'ソート用': int(re.search(r'\d+', year_str).group()) * 100 + int(month_str),
                        '延べ患者数': sum_nobe,
                        '新患初診': sum_shinkan,
                        '再来初診': sum_sairai,
                        '再診': sum_saishin,
                        '稼働日数': active_days
                    })
                    
                    daily_data_list.append(df_p[['フル日付', '延べ患者数 (人)']])
                    break
                except: continue

    if not monthly_data_list:
        st.error("有効なデータを読み込めませんでした。")
        st.stop()

    df_monthly = pd.DataFrame(monthly_data_list).sort_values('ソート用').drop_duplicates(subset=['年月']).reset_index(drop=True)
    df_daily = pd.concat(daily_data_list, ignore_index=True)

    df_monthly['新患率'] = df_monthly.apply(
        lambda row: round((row['新患初診'] / row['延べ患者数']) * 100, 1) if row['延べ患者数'] > 0 else 0, axis=1
    )

    df_monthly['再初診率'] = df_monthly.apply(
        lambda row: round((row['再来初診'] / (row['再来初診'] + row['再診'])) * 100, 1) if (row['再来初診'] + row['再診']) > 0 else 0, axis=1
    )

    df_monthly['一日平均'] = round(df_monthly['延べ患者数'] / df_monthly['稼働日数'], 1)
    df_monthly['累計_延べ患者数'] = df_monthly['延べ患者数'].cumsum()
    df_monthly['累計_新患初診'] = df_monthly['新患初診'].cumsum()

    max_daily_idx = df_daily['延べ患者数 (人)'].idxmax()
    max_daily_date = df_daily.loc[max_daily_idx, 'フル日付']
    max_daily_val = df_daily.loc[max_daily_idx, '延べ患者数 (人)']

    max_avg_idx = df_monthly['一日平均'].idxmax()
    max_avg_month = df_monthly.loc[max_avg_idx, '年月']
    max_avg_val = df_monthly.loc[max_avg_idx, '一日平均']

    st.subheader("🏆 まえだ耳鼻咽喉科 過去最高記録")
    col_rec1, col_rec2 = st.columns(2)
    with col_rec1:
        st.info(f"**最高来院数（1日）:** \n\n ### {max_daily_val:,.0f} 人 \n\n 📅 記録日: {max_daily_date}")
    with col_rec2:
        st.success(f"**最高平均来院数（月間）:** \n\n ### {max_avg_val:,.1f} 人/日 \n\n 📅 記録月: {max_avg_month}")

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

    max_nobe = df_monthly['累計_延べ患者数'].max()
    if pd.notna(max_nobe) and max_nobe > 0:
        for m in range(50000, int(max_nobe) + 1, 50000):
            achieved_rows = df_monthly[df_monthly['累計_延べ患者数'] >= m]
            if not achieved_rows.empty:
                achieved_row = achieved_rows.iloc[0]
                month = achieved_row['年月']
                fig2.add_hline(y=m, line_dash="dash", line_color="#3498DB", opacity=0.7, layer="below")
                fig2.add_trace(go.Scatter(
                    x=[month], y=[achieved_row['累計_延べ患者数']], mode='markers+text',
                    marker=dict(color='#E74C3C', size=12, symbol='star'),
                    text=[f"<b>延べ{int(m/10000)}万人達成</b><br>({month})"], textposition="top left",
                    textfont=dict(color="#2980B9", size=11), showlegend=False, hoverinfo="skip"
                ))

    max_shinkan = df_monthly['累計_新患初診'].max()
    if pd.notna(max_shinkan) and max_shinkan > 0:
        for m in range(10000, int(max_shinkan) + 1, 10000):
            achieved_rows = df_monthly[df_monthly['累計_新患初診'] >= m]
            if not achieved_rows.empty:
                achieved_row = achieved_rows.iloc[0]
                month = achieved_row['年月']
                fig2.add_hline(y=m, line_dash="dot", line_color="#1ABC9C", opacity=0.7, layer="below")
                fig2.add_trace(go.Scatter(
                    x=[month], y=[achieved_row['累計_新患初診']], mode='markers+text',
                    marker=dict(color='#E74C3C', size=12, symbol='star'),
                    text=[f"<b>新患{int(m/10000)}万人達成</b><br>({month})"], textposition="top right", 
                    textfont=dict(color="#FFFFFF", size=11), showlegend=False, hoverinfo="skip"
                ))

    fig2.update_layout(hovermode="x unified", yaxis_title="累計人数 (人)", yaxis=dict(tickformat=",.0f", dtick=10000), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})

    fig3 = go.Figure()
    fig3.add_trace(go.Bar(
        x=df_monthly['年月'], y=df_monthly['一日平均'], name="一日平均", marker_color="#27AE60",
        text=df_monthly['一日平均'], texttemplate="%{text:.1f}", textposition='outside', textangle=0, textfont=dict(size=11),
        hovertemplate="<b>%{x}</b><br>一日平均: %{y:.1f} 人<extra></extra>"
    ))
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
    if not age_files:
        st.error("受付患者数（年齢別）のCSVファイルが見つかりません。")
        st.stop()

    data_list = []
    for f in age_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str = match.group(1) + "年"
            month_str = match.group(2) + "月"
            for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
                try:
                    df_a = pd.read_csv(f, encoding=enc)
                    if '日' not in df_a.columns:
                        continue # 列名が正しく読めなければ次へ
                    df_a = df_a[df_a['日'].str.contains('日', na=False)]
                    for col in df_a.columns:
                        if col != '日':
                            df_a[col] = pd.to_numeric(df_a[col].astype(str).str.replace(',', ''), errors='coerce')
                    
                    summary = df_a.drop(columns=['日']).sum()
                    summary['年度'] = year_str
                    summary['月'] = month_str
                    data_list.append(summary)
                    break
                except: continue

    if not data_list:
        st.error("有効なデータを読み込めませんでした。")
        st.stop()

    age_df = pd.DataFrame(data_list).drop_duplicates(subset=['年度', '月'])
    available_years = sorted(list(age_df['年度'].unique()), key=lambda x: int(re.search(r'\d+', x).group()))
    
    selected_year_age = st.selectbox("📅 表示する年度を選択してください", available_years, index=len(available_years)-1, key="age_year")
    
    current_num_age = int(re.search(r'\d+', selected_year_age).group())
    prev_year_age = f"R{current_num_age - 1}年"

    st.subheader(f"📊 {selected_year_age} 年齢別構成比分析")

    df_curr_age_raw = age_df[age_df['年度'] == selected_year_age].copy()
    active_months = df_curr_age_raw['月'].unique().tolist()
    
    valid_months_age = [f"{i}月" for i in range(1, 13)]
    base_df = pd.DataFrame({'月': valid_months_age})
    df_curr_age = pd.merge(base_df, df_curr_age_raw, on='月', how='left').fillna(0)

    df_prev_age = age_df[age_df['年度'] == prev_year_age].copy()
    df_prev_age_ytd = df_prev_age[df_prev_age['月'].isin(active_months)]

    t1_cols = ['～3歳', '3～6歳']
    t2_cols = ['～3歳', '3～6歳', '6～12歳', '12～20歳', '20～30歳', '30～50歳']
    
    valid_t1 = [c for c in t1_cols if c in df_curr_age.columns]
    valid_t2 = [c for c in t2_cols if c in df_curr_age.columns]

    curr_total_sum = df_curr_age['延べ患者数 (人)'].sum()

    st.write("#### 🎯 当院メインターゲット層 実績")
    
    if active_months:
        active_months_sorted = sorted(active_months, key=lambda x: int(re.search(r'\d+', x).group()))
        if len(active_months_sorted) == 12:
             st.caption(f"※前年比は、前年の年間トータルと比較しています。")
        else:
             st.caption(f"※前年比は、当年のデータが存在する月（{active_months_sorted[0]}〜{active_months_sorted[-1]}）の同期間で比較しています。")
             
    col_kpi1, col_kpi2 = st.columns(2)

    with col_kpi1:
        if valid_t1:
            t1_curr_sum = df_curr_age[valid_t1].sum().sum()
            t1_curr_ratio = round((t1_curr_sum / curr_total_sum) * 100, 1) if curr_total_sum > 0 else 0
            
            if not df_prev_age_ytd.empty:
                t1_prev_sum = df_prev_age_ytd[valid_t1].sum().sum()
                t1_yoy = round((t1_curr_sum / t1_prev_sum) * 100, 1) if t1_prev_sum > 0 else 0
                t1_yoy_text = f"{t1_yoy}%"
                t1_color = "🟢" if t1_yoy >= 100 else "🔴"
            else:
                t1_yoy_text = "-"
                t1_color = "⚪"
                
            st.info(f"**👦 小児層（～6歳）** \n\n **累計人数:** {t1_curr_sum:,.0f} 人　｜　**構成比:** {t1_curr_ratio}%　｜　**前年比:** {t1_color} {t1_yoy_text}")

    with col_kpi2:
        if valid_t2:
            t2_curr_sum = df_curr_age[valid_t2].sum().sum()
            t2_curr_ratio = round((t2_curr_sum / curr_total_sum) * 100, 1) if curr_total_sum > 0 else 0
            
            if not df_prev_age_ytd.empty:
                t2_prev_sum = df_prev_age_ytd[valid_t2].sum().sum()
                t2_yoy = round((t2_curr_sum / t2_prev_sum) * 100, 1) if t2_prev_sum > 0 else 0
                t2_yoy_text = f"{t2_yoy}%"
                t2_color = "🟢" if t2_yoy >= 100 else "🔴"
            else:
                t2_yoy_text = "-"
                t2_color = "⚪"
                
            st.success(f"**👪 親子層（～50歳）** \n\n **累計人数:** {t2_curr_sum:,.0f} 人　｜　**構成比:** {t2_curr_ratio}%　｜　**前年比:** {t2_color} {t2_yoy_text}")

    st.write("---")

    ordered_age_categories = ['～3歳', '3～6歳', '6～12歳', '12～20歳', '20～30歳', '30～50歳', '50～60歳', '60～65歳', '65～75歳', '75歳～']
    age_categories = [c for c in ordered_age_categories if c in age_df.columns]

    col_chart1, col_chart2 = st.columns([2, 1])

    with col_chart1:
        st.write("#### 📈 月別 年齢層積み上げ推移")
        fig_stack = go.Figure()
        for age_cat in age_categories:
            fig_stack.add_trace(go.Bar(
                x=df_curr_age['月'], 
                y=df_curr_age[age_cat], 
                name=age_cat,
                hovertemplate=f"<b>%{{x}}</b><br>{age_cat}: %{{y:,.0f}}人<extra></extra>"
            ))
        fig_stack.update_layout(
            barmode='stack', 
            xaxis_title="診療月", 
            yaxis_title="人数 (人)", 
            hovermode="x unified",
            xaxis=dict(type='category', categoryorder='array', categoryarray=valid_months_age),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_stack, use_container_width=True, config={'displayModeBar': False})

    with col_chart2:
        st.write("#### 🍩 年間 年齢層構成比")
        total_by_age = df_curr_age[age_categories].sum()
        total_by_age = total_by_age[total_by_age > 0]
        
        fig_pie = go.Figure(data=[go.Pie(
            labels=total_by_age.index, 
            values=total_by_age.values, 
            hole=.4,
            textinfo='label+percent',
            textposition='inside',
            insidetextorientation='horizontal',
            sort=False,
            direction='clockwise'
        )])
        fig_pie.update_layout(
            showlegend=False, 
            margin=dict(t=20, b=20, l=0, r=0)
        )
        st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})

    st.write("---")
    st.write("### 📋 詳細数値データ（月別・年齢別一覧）")
    
    st_df_age = df_curr_age[['月', '延べ患者数 (人)'] + age_categories].copy()
    sum_row = st_df_age.sum(numeric_only=True)
    sum_row_df = pd.DataFrame([sum_row])
    sum_row_df['月'] = '総計'
    st_df_age = pd.concat([st_df_age, sum_row_df], ignore_index=True)
    
    def style_age_matrix(df):
        styler = df.set_index('月').style
        fmt = {col: "{:,.0f}" for col in df.columns}
        styler = styler.format(fmt)
        
        def apply_bold_total(row):
            if row.name == '総計':
                return ['font-weight: bold; background-color: #f0f2f6'] * len(row)
            return [''] * len(row)
            
        return styler.apply(apply_bold_total, axis=1)
        
    st.dataframe(style_age_matrix(st_df_age), use_container_width=True)

# ==========================================
# E. 診療行為一覧分析
# ==========================================
elif analysis_mode == "診療行為一覧分析":
    act_files = glob.glob("*診療行為一覧*.csv")
    if not act_files:
        st.error("診療行為一覧のCSVファイルが見つかりません。")
        st.stop()

    data_list = []
    for f in act_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str = match.group(1) + "年"
            month_str = match.group(2) + "月"
            for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
                try:
                    df_act = pd.read_csv(f, encoding=enc)
                    
                    col_name = None
                    if '診 療 行 為 名 称' in df_act.columns:
                        col_name = '診 療 行 為 名 称'
                    elif '診療行為名称' in df_act.columns:
                        col_name = '診療行為名称'
                    
                    if not col_name:
                        continue # 列名が正しく読めなければ次へ
                        
                    df_act = df_act.dropna(subset=[col_name])
                    df_act = df_act[~df_act[col_name].astype(str).str.contains('合計')]
                    
                    if '診療区分' in df_act.columns:
                        df_act = df_act[~df_act['診療区分'].astype(str).isin(['検査', '保険外費用'])]
                    
                    for col in ['回数', '総点数 (点)']:
                        if col in df_act.columns:
                            df_act[col] = pd.to_numeric(
                                df_act[col].astype(str).str.replace(r'[，,点円％%]', '', regex=True), 
                                errors='coerce'
                            ).fillna(0)
                    
                    summary = df_act.groupby(col_name)[['回数', '総点数 (点)']].sum().reset_index()
                    summary.rename(columns={col_name: '診療行為名称', '総点数 (点)': '総点数'}, inplace=True)
                    summary['年度'] = year_str
                    summary['月'] = month_str
                    data_list.append(summary)
                    break
                except Exception as e:
                    continue

    if not data_list:
        st.error("有効なデータを読み込めませんでした。")
        st.stop()

    act_df = pd.concat(data_list, ignore_index=True)
    
    available_years = sorted(list(act_df['年度'].unique()), key=lambda x: int(re.search(r'\d+', x).group()))
    
    selected_year_act = st.selectbox("📅 表示する年度を選択してください", available_years, index=len(available_years)-1, key="act_year")
    
    current_num_act = int(re.search(r'\d+', selected_year_act).group())
    prev_year_act = f"R{current_num_act - 1}年"

    st.subheader(f"📊 {selected_year_act} 診療行為 分析 (処置・手術・医学管理等)")

    df_curr_act = act_df[act_df['年度'] == selected_year_act].copy()
    active_months_act = df_curr_act['月'].unique().tolist()
    
    df_prev_act_full = act_df[act_df['年度'] == prev_year_act].copy()
    df_prev_act_ytd = df_prev_act_full[df_prev_act_full['月'].isin(active_months_act)]

    curr_sum_act = df_curr_act.groupby('診療行為名称')[['回数', '総点数']].sum().reset_index()
    
    if not df_prev_act_ytd.empty:
        prev_sum_act = df_prev_act_ytd.groupby('診療行為名称')[['総点数']].sum().reset_index().rename(columns={'総点数': '前年総点数'})
        rank_df = pd.merge(curr_sum_act, prev_sum_act, on='診療行為名称', how='left').fillna(0)
    else:
        rank_df = curr_sum_act.copy()
        rank_df['前年総点数'] = 0

    rank_df['単価'] = rank_df.apply(lambda x: round(x['総点数'] / x['回数'], 1) if x['回数'] > 0 else 0, axis=1)
    rank_df['前年比'] = rank_df.apply(lambda x: round((x['総点数'] / x['前年総点数']) * 100, 1) if x['前年総点数'] > 0 else 0, axis=1)

    top10_df = rank_df.sort_values('総点数', ascending=False).head(10).reset_index(drop=True)
    top10_df.index = top10_df.index + 1 
    
    st.write("#### 🏆 年間 総点数 TOP10 ランキング")
    
    if active_months_act:
        active_months_sorted = sorted(active_months_act, key=lambda x: int(re.search(r'\d+', x).group()))
        if len(active_months_sorted) == 12:
             st.caption(f"※前年比は、前年の年間トータルと比較しています。")
        else:
             st.caption(f"※前年比は、当年のデータが存在する月（{active_months_sorted[0]}〜{active_months_sorted[-1]}）の同期間で比較しています。")
             
    disp_top10 = top10_df[['診療行為名称', '単価', '回数', '総点数', '前年比']].copy()
    
    def style_top10(df):
        styler = df.style
        fmt = {
            '単価': "{:,.1f} 点",
            '回数': "{:,.0f} 回",
            '総点数': "{:,.0f} 点",
            '前年比': "{:.1f}%"
        }
        styler = styler.format(fmt)
        
        def color_yoy(val):
            try:
                v = float(val)
                if v >= 100: return 'color: #2E86C1; font-weight: bold'
                elif 0 < v < 100: return 'color: #E74C3C; font-weight: bold'
                return ''
            except:
                return ''
                
        styler = styler.map(color_yoy, subset=['前年比'])
        return styler

    st.dataframe(style_top10(disp_top10), use_container_width=True)

    st.write("---")
    st.write("### 📈 診療行為別 月別推移")
    
    valid_months_act = [f"{i}月" for i in range(1, 13)]
    
    all_act_items = sorted(list(set(df_curr_act['診療行為名称'].unique().tolist() + df_prev_act_full['診療行為名称'].unique().tolist())))
    default_act = top10_df.iloc[0]['診療行為名称'] if not top10_df.empty else (all_act_items[0] if all_act_items else "")
    
    if all_act_items:
        selected_act_item = st.selectbox("🔍 グラフ表示する診療行為を選択してください", all_act_items, index=all_act_items.index(default_act) if default_act in all_act_items else 0)
        
        plot_df_act = pd.DataFrame({'月': valid_months_act})
        curr_item_data = df_curr_act[df_curr_act['診療行為名称'] == selected_act_item].groupby('月')['総点数'].sum().reset_index().rename(columns={'総点数': '当年'})
        prev_item_data = df_prev_act_full[df_prev_act_full['診療行為名称'] == selected_act_item].groupby('月')['総点数'].sum().reset_index().rename(columns={'総点数': '前年'})
        
        plot_df_act = pd.merge(plot_df_act, curr_item_data, on='月', how='left')
        plot_df_act = pd.merge(plot_df_act, prev_item_data, on='月', how='left').fillna(0)
        
        plot_df_act['前年比'] = plot_df_act.apply(lambda x: round((x['当年'] / x['前年']) * 100, 1) if x['前年'] > 0 else 0, axis=1)
        colors_act = ['#E74C3C' if 0 < val < 100 else ('#2E86C1' if val >= 100 else '#000000') for val in plot_df_act['前年比']]

        fig_act = go.Figure()
        fig_act.add_trace(go.Scatter(
            x=plot_df_act['月'], y=plot_df_act['当年'],
            mode='lines+markers+text', name=f'当年 ({selected_year_act})',
            line=dict(color='#2E86C1', width=4),
            text=plot_df_act['前年比'].apply(lambda x: f"{x}%" if x > 0 else ""),
            textposition="top center", textfont=dict(color=colors_act, size=13, family="Arial Black"),
            hovertemplate="<b>%{x}</b><br>当年: %{y:,.0f} 点<extra></extra>"
        ))
        
        if not df_prev_act_full.empty:
            fig_act.add_trace(go.Scatter(
                x=plot_df_act['月'], y=plot_df_act['前年'],
                mode='lines+markers', name=f'前年 ({prev_year_act})',
                line=dict(color='#ABB2B9', width=2, dash='dot'),
                hovertemplate="前年: %{y:,.0f} 点<extra></extra>"
            ))

        fig_act.update_layout(
            xaxis_title="診療月",
            yaxis_title="総点数 (点)",
            hovermode="x unified",
            xaxis=dict(type='category', categoryorder='array', categoryarray=valid_months_act),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_act, use_container_width=True, config={'displayModeBar': False})
        
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
# F. 検査一覧分析
# ==========================================
elif analysis_mode == "検査一覧分析":
    act_files = glob.glob("*診療行為一覧*.csv")
    if not act_files:
        st.error("診療行為一覧のCSVファイルが見つかりません。")
        st.stop()

    data_list = []
    for f in act_files:
        match = re.search(r'(R\d+)年(\d+)月', f)
        if match:
            year_str = match.group(1) + "年"
            month_str = match.group(2) + "月"
            for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
                try:
                    df_act = pd.read_csv(f, encoding=enc)
                    
                    col_name = None
                    if '診 療 行 為 名 称' in df_act.columns:
                        col_name = '診 療 行 為 名 称'
                    elif '診療行為名称' in df_act.columns:
                        col_name = '診療行為名称'
                    
                    if not col_name:
                        continue # 列名が正しく読めなければ次へ
                        
                    df_act = df_act.dropna(subset=[col_name])
                    df_act = df_act[~df_act[col_name].astype(str).str.contains('合計')]
                    
                    # 【検査のみ抽出】
                    if '診療区分' in df_act.columns:
                        df_act = df_act[df_act['診療区分'].astype(str).str.contains('検査')]
                    else:
                        continue
                    
                    for col in ['回数', '総点数 (点)']:
                        if col in df_act.columns:
                            df_act[col] = pd.to_numeric(
                                df_act[col].astype(str).str.replace(r'[，,点円％%]', '', regex=True), 
                                errors='coerce'
                            ).fillna(0)
                    
                    summary = df_act.groupby(col_name)[['回数', '総点数 (点)']].sum().reset_index()
                    summary.rename(columns={col_name: '診療行為名称', '総点数 (点)': '総点数'}, inplace=True)
                    summary['年度'] = year_str
                    summary['月'] = month_str
                    data_list.append(summary)
                    break
                except Exception as e:
                    continue

    if not data_list:
        st.error("有効な検査データを読み込めませんでした。")
        st.stop()

    act_df = pd.concat(data_list, ignore_index=True)
    
    available_years = sorted(list(act_df['年度'].unique()), key=lambda x: int(re.search(r'\d+', x).group()))
    
    selected_year_act = st.selectbox("📅 表示する年度を選択してください", available_years, index=len(available_years)-1, key="ins_year")
    
    current_num_act = int(re.search(r'\d+', selected_year_act).group())
    prev_year_act = f"R{current_num_act - 1}年"

    st.subheader(f"📊 {selected_year_act} 検査一覧 分析")

    df_curr_act = act_df[act_df['年度'] == selected_year_act].copy()
    active_months_act = df_curr_act['月'].unique().tolist()
    valid_months_act = [f"{i}月" for i in range(1, 13)]
    
    df_prev_act_full = act_df[act_df['年度'] == prev_year_act].copy()
    df_prev_act_ytd = df_prev_act_full[df_prev_act_full['月'].isin(active_months_act)]

    st.write("#### 🎯 注力検査項目 実績")
    
    if active_months_act:
        active_months_sorted = sorted(active_months_act, key=lambda x: int(re.search(r'\d+', x).group()))
        if len(active_months_sorted) == 12:
             st.caption(f"※前年比は、前年の年間トータルと比較しています。")
        else:
             st.caption(f"※前年比は、当年のデータが存在する月（{active_months_sorted[0]}〜{active_months_sorted[-1]}）の同期間で比較しています。")

    col_t1, col_t2, col_t3 = st.columns(3)
    
    def get_kpi_data(item_names):
        if isinstance(item_names, str):
            item_names = [item_names]
        curr_data = df_curr_act[df_curr_act['診療行為名称'].isin(item_names)]
        prev_data = df_prev_act_ytd[df_prev_act_ytd['診療行為名称'].isin(item_names)]
        
        curr_count = curr_data['回数'].sum()
        curr_pts = curr_data['総点数'].sum()
        prev_pts = prev_data['総点数'].sum()
        
        yoy = round((curr_pts / prev_pts) * 100, 1) if prev_pts > 0 else 0
        return curr_count, curr_pts, yoy

    c1_count, c1_pts, c1_yoy = get_kpi_data(['チンパノメトリー'])
    c1_color = "🟢" if c1_yoy >= 100 else "🔴"
    c1_yoy_str = f"{c1_yoy}%" if c1_pts > 0 or c1_count > 0 else "-"
    with col_t1:
        info_block = "<div class='kpi-container stInfo'><p><b>👂 チンパノメトリー</b> </p><br>" + f"<p><b>回数:</b> {c1_count:,.0f} 回 ｜ <b>総点数:</b> {c1_pts:,.0f} 点 </p><br>" + f"<p><b>前年比:</b> {c1_color} {c1_yoy_str}</p></div>"
        st.markdown(info_block, unsafe_allow_html=True)
        
    c2_count, c2_pts, c2_yoy = get_kpi_data(['標準純音聴力検査', '簡易聴力検査（気導純音聴力）'])
    c2_color = "🟢" if c2_yoy >= 100 else "🔴"
    c2_yoy_str = f"{c2_yoy}%" if c2_pts > 0 or c2_count > 0 else "-"
    
    h1_c, h1_p, h1_y = get_kpi_data('標準純音聴力検査')
    h1_y_str = f"{h1_y}%" if h1_p > 0 or h1_c > 0 else "-"
    h2_c, h2_p, h2_y = get_kpi_data('簡易聴力検査（気導純音聴力）')
    h2_y_str = f"{h2_y}%" if h2_p > 0 or h2_c > 0 else "-"

    with col_t2:
        succ_block = "<div class='kpi-container stSuccess'><p><b>🎧 聴力検査（合算）</b> </p><br>" + f"<p><b>回数:</b> {c2_count:,.0f} 回 ｜ <b>総点数:</b> {c2_pts:,.0f} 点 </p><br>" + f"<p><b>前年比:</b> {c2_color} {c2_yoy_str}</p><hr><ul>" + f"<li><b>標準純音</b>: {h1_c:,.0f}回 ｜ {h1_p:,.0f}点 ｜ 前年: {h1_y_str}</li>" + f"<li><b>簡易聴力</b>: {h2_c:,.0f}回 ｜ {h2_p:,.0f}点 ｜ 前年: {h2_y_str}</li></ul></div>"
        st.markdown(succ_block, unsafe_allow_html=True)
        
    c3_count, c3_pts, c3_yoy = get_kpi_data(['ＥＦ－中耳', 'ＥＦ－喉頭', 'ＥＦ－嗅裂・鼻咽腔・副鼻腔'])
    c3_color = "🟢" if c3_yoy >= 100 else "🔴"
    c3_yoy_str = f"{c3_yoy}%" if c3_pts > 0 or c3_count > 0 else "-"
    
    f1_c, f1_p, f1_y = get_kpi_data('ＥＦ－中耳')
    f1_y_str = f"{f1_y}%" if f1_p > 0 or f1_c > 0 else "-"
    f2_c, f2_p, f2_y = get_kpi_data('ＥＦ－喉頭')
    f2_y_str = f"{f2_y}%" if f2_p > 0 or f2_c > 0 else "-"
    f3_c, f3_p, f3_y = get_kpi_data('ＥＦ－嗅裂・鼻咽腔・副鼻腔')
    f3_y_str = f"{f3_y}%" if f3_p > 0 or f3_c > 0 else "-"

    with col_t3:
        warn_block = "<div class='kpi-container stWarning'><p><b>🔬 ファイバー（合算）</b> </p><br>" + f"<p><b>回数:</b> {c3_count:,.0f} 回 ｜ <b>総点数:</b> {c3_pts:,.0f} 点 </p><br>" + f"<p><b>前年比:</b> {c3_color} {c3_yoy_str}</p><hr><ul>" + f"<li><b>中耳</b>: {f1_c:,.0f}回 ｜ {f1_p:,.0f}点 ｜ 前年: {f1_y_str}</li>" + f"<li><b>喉頭</b>: {f2_c:,.0f}回 ｜ {f2_p:,.0f}点 ｜ 前年: {f2_y_str}</li>" + f"<li><b>鼻咽腔等</b>: {f3_c:,.0f}回 ｜ {f3_p:,.0f}点 ｜ 前年: {f3_y_str}</li></ul></div>"
        st.markdown(warn_block, unsafe_allow_html=True)

    col_g1, col_g2, col_g3 = st.columns(3)

    def create_mini_trend_fig(item_names):
        if isinstance(item_names, str):
            item_names = [item_names]
            
        curr_data = df_curr_act[df_curr_act['診療行為名称'].isin(item_names)]
        curr_monthly = curr_data.groupby('月')['回数'].sum().reindex(valid_months_act).fillna(0)
        
        prev_data = df_prev_act_full[df_prev_act_full['診療行為名称'].isin(item_names)]
        prev_monthly = prev_data.groupby('月')['回数'].sum().reindex(valid_months_act).fillna(0)
        
        fig = go.Figure()
        
        if not prev_data.empty:
            fig.add_trace(go.Scatter(
                x=valid_months_act, y=prev_monthly, 
                mode='lines+markers', name='前年', 
                line=dict(color='#ABB2B9', width=2, dash='dot'),
                hovertemplate="前年: %{y:,.0f} 回<extra></extra>"
            ))
            
        fig.add_trace(go.Scatter(
            x=valid_months_act, y=curr_monthly, 
            mode='lines+markers', name='当年', 
            line=dict(color='#2E86C1', width=3),
            hovertemplate="当年: %{y:,.0f} 回<extra></extra>"
        ))
        
        fig.update_layout(
            height=180,
            margin=dict(l=10, r=10, t=10, b=10),
            xaxis=dict(showticklabels=True, tickfont=dict(size=10)),
            yaxis=dict(showticklabels=True, tickfont=dict(size=10)),
            showlegend=False,
            hovermode="x unified",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)"
        )
        return fig

    with col_g1:
        st.plotly_chart(create_mini_trend_fig(['チンパノメトリー']), use_container_width=True, config={'displayModeBar': False})
    with col_g2:
        st.plotly_chart(create_mini_trend_fig(['標準純音聴力検査', '簡易聴力検査（気導純音聴力）']), use_container_width=True, config={'displayModeBar': False})
    with col_g3:
        st.plotly_chart(create_mini_trend_fig(['ＥＦ－中耳', 'ＥＦ－喉頭', 'ＥＦ－嗅裂・鼻咽腔・副鼻腔']), use_container_width=True, config={'displayModeBar': False})

    st.write("---")

    curr_sum_act = df_curr_act.groupby('診療行為名称')[['回数', '総点数']].sum().reset_index()
    
    if not df_prev_act_ytd.empty:
        prev_sum_act = df_prev_act_ytd.groupby('診療行為名称')[['総点数']].sum().reset_index().rename(columns={'総点数': '前年総点数'})
        rank_df = pd.merge(curr_sum_act, prev_sum_act, on='診療行為名称', how='left').fillna(0)
    else:
        rank_df = curr_sum_act.copy()
        rank_df['前年総点数'] = 0

    rank_df['単価'] = rank_df.apply(lambda x: round(x['総点数'] / x['回数'], 1) if x['回数'] > 0 else 0, axis=1)
    rank_df['前年比'] = rank_df.apply(lambda x: round((x['総点数'] / x['前年総点数']) * 100, 1) if x['前年総点数'] > 0 else 0, axis=1)

    top10_df = rank_df.sort_values('総点数', ascending=False).head(10).reset_index(drop=True)
    top10_df.index = top10_df.index + 1 
    
    st.write("#### 🏆 年間 総点数 TOP10 ランキング (検査)")
             
    disp_top10 = top10_df[['診療行為名称', '単価', '回数', '総点数', '前年比']].copy()
    
    def style_top10(df):
        styler = df.style
        fmt = {
            '単価': "{:,.1f} 点",
            '回数': "{:,.0f} 回",
            '総点数': "{:,.0f} 点",
            '前年比': "{:.1f}%"
        }
        styler = styler.format(fmt)
        
        def color_yoy(val):
            try:
                v = float(val)
                if v >= 100: return 'color: #2E86C1; font-weight: bold'
                elif 0 < v < 100: return 'color: #E74C3C; font-weight: bold'
                return ''
            except:
                return ''
                
        styler = styler.map(color_yoy, subset=['前年比'])
        return styler

    st.dataframe(style_top10(disp_top10), use_container_width=True)

    st.write("---")
    st.write("### 📈 検査別 月別推移")
    
    all_act_items = sorted(list(set(df_curr_act['診療行為名称'].unique().tolist() + df_prev_act_full['診療行為名称'].unique().tolist())))
    default_act = top10_df.iloc[0]['診療行為名称'] if not top10_df.empty else (all_act_items[0] if all_act_items else "")
    
    if all_act_items:
        selected_act_item = st.selectbox("🔍 グラフ表示する検査を選択してください", all_act_items, index=all_act_items.index(default_act) if default_act in all_act_items else 0)
        
        plot_df_act = pd.DataFrame({'月': valid_months_act})
        curr_item_data = df_curr_act[df_curr_act['診療行為名称'] == selected_act_item].groupby('月')['総点数'].sum().reset_index().rename(columns={'総点数': '当年'})
        prev_item_data = df_prev_act_full[df_prev_act_full['診療行為名称'] == selected_act_item].groupby('月')['総点数'].sum().reset_index().rename(columns={'総点数': '前年'})
        
        plot_df_act = pd.merge(plot_df_act, curr_item_data, on='月', how='left')
        plot_df_act = pd.merge(plot_df_act, prev_item_data, on='月', how='left').fillna(0)
        
        plot_df_act['前年比'] = plot_df_act.apply(lambda x: round((x['当年'] / x['前年']) * 100, 1) if x['前年'] > 0 else 0, axis=1)
        colors_act = ['#E74C3C' if 0 < val < 100 else ('#2E86C1' if val >= 100 else '#000000') for val in plot_df_act['前年比']]

        fig_act = go.Figure()
        fig_act.add_trace(go.Scatter(
            x=plot_df_act['月'], y=plot_df_act['当年'],
            mode='lines+markers+text', name=f'当年 ({selected_year_act})',
            line=dict(color='#2E86C1', width=4),
            text=plot_df_act['前年比'].apply(lambda x: f"{x}%" if x > 0 else ""),
            textposition="top center", textfont=dict(color=colors_act, size=13, family="Arial Black"),
            hovertemplate="<b>%{x}</b><br>当年: %{y:,.0f} 点<extra></extra>"
        ))
        
        if not df_prev_act_full.empty:
            fig_act.add_trace(go.Scatter(
                x=plot_df_act['月'], y=plot_df_act['前年'],
                mode='lines+markers', name=f'前年 ({prev_year_act})',
                line=dict(color='#ABB2B9', width=2, dash='dot'),
                hovertemplate="前年: %{y:,.0f} 点<extra></extra>"
            ))

        fig_act.update_layout(
            xaxis_title="診療月",
            yaxis_title="総点数 (点)",
            hovermode="x unified",
            xaxis=dict(type='category', categoryorder='array', categoryarray=valid_months_act),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_act, use_container_width=True, config={'displayModeBar': False})
        
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

# ==========================================
# G. AI総合経営アドバイス (完全動的データ生成版)
# ==========================================
elif analysis_mode == "AI総合経営アドバイス":
    st.subheader("🤖 AI総合経営アドバイス（目標：レセプト単価750点達成に向けて）")
    
    with st.spinner("すべてのデータを統合・多角的に分析しています..."):
        
        # --- 1. レセプトデータの取得 ---
        rece_files = glob.glob("*レセプト*.csv")
        years_found = sorted(list(set(re.search(r'(R\d+年)', f).group(1) for f in rece_files if re.search(r'(R\d+年)', f))))
        
        if not years_found:
            st.error("データが見つかりません。")
            st.stop()
            
        latest_year_str = years_found[-1]
        df_rece_ai = get_clean_df(latest_year_str)
        
        rece_tanka = 0
        rece_patients = 0
        latest_m_ai = "不明"
        is_fallback_ai = False
        gap_ai = 0
        
        # 課題抽出用
        worst_item = "不明"
        worst_diff = 0

        if df_rece_ai is not None and 'レセ単価' in df_rece_ai.columns:
            target_month, latest_row, is_fallback_ai, original_latest_month = get_latest_complete_month(df_rece_ai, latest_year_str)
            
            if target_month:
                rece_tanka = latest_row['レセ単価_num']
                latest_m_ai = target_month
                
                cnt_col = next((c for c in df_rece_ai.columns if '件数' in c or '枚数' in c), None)
                if cnt_col:
                    cnt_str = str(latest_row[cnt_col]).replace(',', '').replace('件', '').replace('枚', '').strip()
                    rece_patients = float(cnt_str) if cnt_str.replace('.', '', 1).isdigit() else 0
                
                gap_ai = 750 - rece_tanka

                latest_m_num = int(latest_m_ai.replace('月', ''))
                df_tmp_ai = df_rece_ai[df_rece_ai['月'].isin([f"{i}月" for i in range(1, 13)])].copy()
                
                if latest_m_num == 1:
                    prev_year_str = f"R{int(re.search(r'\d+', latest_year_str).group()) - 1}年"
                    df_prev_ai = get_clean_df(prev_year_str)
                    prev_month = "12月"
                    row_p = df_prev_ai[df_prev_ai['月'] == prev_month].iloc[0] if df_prev_ai is not None and not df_prev_ai[df_prev_ai['月'] == prev_month].empty else None
                else:
                    prev_month = f"{latest_m_num - 1}月"
                    row_p = df_tmp_ai[df_tmp_ai['月'] == prev_month].iloc[0] if not df_tmp_ai[df_tmp_ai['月'] == prev_month].empty else None
                
                if row_p is not None:
                    row_c = latest_row
                    exclude_cols = ["月", "総", "前年", "比", "枚数", "件数", "日数", "レセ単価", "日当点", "平均", "合計", "レセ単価_num"]
                    comp_cols = [c for c in df_rece_ai.columns if not any(k in c for k in exclude_cols)]
                    
                    diffs = []
                    for col in comp_cols:
                        if col in row_c and col in row_p:
                            v_c_str = str(row_c[col]).replace(',', '').replace('点', '').replace('円', '').strip()
                            v_p_str = str(row_p[col]).replace(',', '').replace('点', '').replace('円', '').strip()
                            v_c = float(v_c_str) if v_c_str.replace('.', '', 1).isdigit() else 0.0
                            v_p = float(v_p_str) if v_p_str.replace('.', '', 1).isdigit() else 0.0
                            diffs.append({'名称': col, '点数差': v_c - v_p})
                    
                    if diffs:
                        df_diff_ai = pd.DataFrame(diffs)
                        if not df_diff_ai.empty:
                            worst_row = df_diff_ai.sort_values('点数差', ascending=True).iloc[0]
                            if worst_row['点数差'] < 0:
                                worst_item = worst_row['名称']
                                worst_diff = worst_row['点数差']

        # --- 3. 診療行為データから実施率を取得 ---
        df_a_curr = get_act_summary_for_ai(latest_year_str, latest_m_ai)
        fiber_cnt = 0
        audio_cnt = 0
        tympa_cnt = 0
        if not df_a_curr.empty:
            fiber_cnt = df_a_curr[df_a_curr['名称'].str.contains('ファイバー|ＥＦ', na=False)]['回数'].sum()
            audio_cnt = df_a_curr[df_a_curr['名称'].str.contains('聴力', na=False)]['回数'].sum()
            tympa_cnt = df_a_curr[df_a_curr['名称'].str.contains('チンパノ', na=False)]['回数'].sum()

        fiber_rate = (fiber_cnt / rece_patients * 100) if rece_patients > 0 else 0
        tympa_rate = (tympa_cnt / rece_patients * 100) if rece_patients > 0 else 0

        # --- 4. 年齢構成データの取得 ---
        kids_cnt = 0
        total_age_cnt = 0
        try:
            target_y_num = int(re.search(r'\d+', latest_year_str).group())
            target_m_num = int(re.search(r'\d+', latest_m_ai).group())
            
            age_files = glob.glob("*受付患者数（年齢別）*.csv")
            for f in age_files:
                f_norm = unicodedata.normalize('NFKC', f)
                match = re.search(r'(?:R|令和)?0*(\d+)年0*(\d+)月', f_norm)
                if match and int(match.group(1)) == target_y_num and int(match.group(2)) == target_m_num:
                    for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
                        try:
                            df_age = pd.read_csv(f, encoding=enc)
                            if '日' not in df_age.columns: continue # 列確認
                            df_age = df_age[df_age['日'].str.contains('日', na=False)]
                            for col in df_age.columns:
                                if col != '日':
                                    df_age[col] = pd.to_numeric(df_age[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                            sum_age = df_age.drop(columns=['日']).sum()
                            t1_cols = [c for c in ['～3歳', '3～6歳', '6～12歳'] if c in sum_age.index]
                            kids_cnt = sum_age[t1_cols].sum()
                            total_age_cnt = sum_age['延べ患者数 (人)'] if '延べ患者数 (人)' in sum_age.index else sum_age.sum()
                            break
                        except: continue
        except:
            pass
            
        kids_ratio = (kids_cnt / total_age_cnt * 100) if total_age_cnt > 0 else 0

        # --- 5. シミュレーション計算 ---
        if gap_ai > 0 and rece_patients > 0:
            required_points_total = gap_ai * rece_patients
            add_fiber_cnt = required_points_total / 600
            new_fiber_rate = ((fiber_cnt + add_fiber_cnt) / rece_patients * 100)
            
            add_tympa_cnt = required_points_total / 340
            new_tympa_rate = ((tympa_cnt + add_tympa_cnt) / rece_patients * 100)
            
            req_audio = required_points_total / 350
        else:
            required_points_total = 0
            add_fiber_cnt = 0
            new_fiber_rate = 0
            add_tympa_cnt = 0
            new_tympa_rate = 0
            req_audio = 0

    st.write("<br>", unsafe_allow_html=True)
    
    if gap_ai > 0:
        status_color = "#E74C3C"
        status_text = f"あと {gap_ai:,.0f} 点 不足しています"
    else:
        status_color = "#27AE60"
        status_text = f"目標を {abs(gap_ai):,.0f} 点 上回って達成中です！"

    fallback_html = ""
    if is_fallback_ai:
        fallback_html = f"<div style='background-color:#FFF3CD; color:#856404; padding:5px 10px; border-radius:5px; font-size:0.85em; margin-bottom:10px;'>⚠️ 最新月の診療データが未確定のため、データが揃っている<b>【{latest_m_ai}】</b>を基準に表示しています。</div>"

    top_block = "<div style='background-color: #F8F9F9; padding: 20px; border-radius: 10px; border: 1px solid #D5D8DC; margin-bottom: 20px;'>" + f"{fallback_html}" + "<h3 style='color: #2C3E50; margin-top: 0; margin-bottom: 10px;'>📊 現在地と目標のギャップ確認</h3>" + "<p style='font-size: 1.1rem; line-height: 1.6; margin-bottom:0;'>" + "<span style='background-color:#EBF5FB; padding:2px 6px; border-radius:4px; font-size:0.85em; color:#2E86C1;'>ℹ️ レセプトデータと同期した確定月を表示</span><br>" + f"最新実績（{latest_year_str} {latest_m_ai}）におけるレセプト単価は <b>{rece_tanka:,.0f} 点</b> です。<br>" + f"今年の目標である <b>750点</b> に対して、現在 <b style='color: {status_color}; font-size: 1.3rem;'>{status_text}</b>。" + "</p></div>"
    st.markdown(top_block, unsafe_allow_html=True)
    
    col_h1, col_h2 = st.columns(2)
    with col_h1:
        st.markdown("<div style='height: 60px; display: flex; align-items: flex-end; margin-bottom: 15px;'><h4 style='margin: 0; padding: 0; color: #31333F; font-size: 1.25rem;'>💡 達成に向けたキードライバー<br><span style='font-size: 1rem; font-weight: normal; color: #555;'>（AIが抽出した現状の課題と対策）</span></h4></div>", unsafe_allow_html=True)
    with col_h2:
        st.markdown("<div style='height: 60px; display: flex; align-items: flex-end; margin-bottom: 15px;'><h4 style='margin: 0; padding: 0; color: #31333F; font-size: 1.25rem;'>📈 シミュレーション<br><span style='font-size: 1rem; font-weight: normal; color: #555;'>（あとどれくらい必要か？）</span></h4></div>", unsafe_allow_html=True)
        
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        issue_text = ""
        if kids_ratio > 0:
            issue_text += f"<li><b>小児層の集患基盤：</b> 現在、患者全体の<b>約{kids_ratio:.1f}%</b>が12歳以下の小児層です。ファミリー層からの強い支持を活かしたアプローチが有効です。</li>"
            
        if worst_diff < 0:
            issue_text += f"<li style='margin-top:10px;'><b>「{worst_item}」項目の下落：</b> レセプトデータ分析の結果、前月比で「{worst_item}」が全体で<b>{abs(worst_diff):,.0f}点低下</b>し、単価の足を引っ張っています。算定漏れがないかレセコンの入力確認が必要です。</li>"
            
        if tympa_rate > 0 and tympa_rate < 20 and kids_ratio > 20:
            issue_text += f"<li style='margin-top:10px;'><b>小児割合に対する検査の少なさ：</b> 小児が多いにも関わらず、チンパノメトリーの実施率が <b>{tympa_rate:.1f}%</b> に留まっています。中耳炎の客観的評価の機会を損失している可能性があります。</li>"
            
        if not issue_text:
            issue_text = "<li>目立ったマイナス要因は見当たらず、全体的に安定した推移を見せています。</li>"
            
        adv_block = "<div class='target-box' style='height: 100%;'>" + "<ul style='font-size: 1.05rem; line-height: 1.6; padding-left: 20px;'>" + f"{issue_text}" + "</ul>" + "<hr style='border-top: 1px dashed #AED6F1;'>" + "<b>【AIからの推奨アクション】</b><br>" + "小児の鼻症状・耳症状受付時に、スタッフ主導でチンパノメトリーやファイバー検査へ誘導する<b>「セット実施フロー」</b>を構築し、医師の負担を減らしつつ算定漏れを防ぎましょう。" + "</div>"
        st.markdown(adv_block, unsafe_allow_html=True)

    with col_c2:
        if gap_ai > 0 and rece_patients > 0:
            sim_block = "<div style='border: 2px solid #AED6F1; border-radius: 10px; padding: 15px; background-color: #EBF5FB; height: 100%;'>" + f"<p>現在の月間患者数（約 <b>{rece_patients:,.0f} 人</b>）をベースに計算すると、単価をあと {gap_ai:,.0f} 点引き上げるには、月に <b>約 {required_points_total:,.0f} 点</b> の積み上げが必要です。</p>" + "<hr style='border-top: 1px dashed #AED6F1;'>" + "<p style='margin-bottom: 5px;'>これを当院の<b>注力検査</b>の実施件数で換算すると...</p>" + "<div style='background-color:#fff; padding:10px; border-radius:5px; margin-bottom:10px;'>" + f"<b style='color:#2E86C1;'>① チンパノメトリーの徹底</b><br>現在の実施率 {tympa_rate:.1f}% を <b>{new_tympa_rate:.1f}%</b> に引き上げる。<br><span style='font-size:0.9em; color:#555;'>（月にあと <b>約 {add_tympa_cnt:,.0f} 件</b> 追加 / 1日約 {add_tympa_cnt/20:.1f}件）</span>" + "</div>" + "<div style='background-color:#fff; padding:10px; border-radius:5px; margin-bottom:10px;'>" + f"<b style='color:#2E86C1;'>② ファイバーの適応拡大</b><br>現在の実施率 {fiber_rate:.1f}% を <b>{new_fiber_rate:.1f}%</b> に引き上げる。<br><span style='font-size:0.9em; color:#555;'>（月にあと <b>約 {add_fiber_cnt:,.0f} 件</b> 追加 / 1日約 {add_fiber_cnt/20:.1f}件）</span>" + "</div>" + "<p style='font-size: 0.85rem; color: #777; margin-top: 10px;'>※1ヶ月の診療日数を20日として計算。すべての不足分を1つの検査で補った場合の目安です。</p>" + "</div>"
            st.markdown(sim_block, unsafe_allow_html=True)
        else:
            if gap_ai <= 0:
                succ_block2 = "<div style='border: 2px solid #27AE60; border-radius: 10px; padding: 15px; background-color: #EAFAF1; height: 100%;'>" + "<p style='font-size:1.1rem; color:#1E8449; font-weight:bold; text-align:center;'>🎉 すでに目標単価をクリアしています！</p>" + "<p style='margin-top:10px;'>現在の素晴らしい取り組みを継続しつつ、さらに「患者満足度の向上」や「スタッフの業務効率化」など、次のステージに向けた施策にシフトしていくことをお勧めします。</p>" + "</div>"
                st.markdown(succ_block2, unsafe_allow_html=True)
            else:
                err_block = "<div style='border: 2px solid #E74C3C; border-radius: 10px; padding: 15px; background-color: #FDEDEC; height: 100%;'>" + "<p style='font-size:1.1rem; color:#C0392B; font-weight:bold; text-align:center;'>患者数のデータが取得できません</p>" + "<p style='margin-top:10px;'>レセプトデータの中に「件数」や「枚数」を示す項目が見つからなかったため、シミュレーションの計算ができませんでした。</p>" + "</div>"
                st.markdown(err_block, unsafe_allow_html=True)
        
    st.write("---")
    st.markdown("#### 🤖 AIからの総評")
    st.markdown("まえだ耳鼻咽喉科の最大の強みは、**明確なターゲット基盤（小児・ファミリー層）**がすでに構築されている点にあります。無理に新しい患者層を開拓するよりも、**「今来院されている患者様に対して、必要な検査や治療の選択肢を漏れなく提案し、医療の質を深めること」**が、目標単価750点達成の最短ルートです。明日から、**「今日はファイバーをあと○件実施しよう」**という具体的な1日の目標件数をスタッフ全員で共有し、朝礼などで意識づけを行うことをお勧めします。")
