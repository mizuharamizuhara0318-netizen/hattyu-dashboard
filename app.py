import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import datetime
import os
import json
import random

st.set_page_config(page_title="発注検討表（Web管理システム）", layout="wide")

st.markdown("""
<style>
    header[data-testid="stHeader"] {
        height: 2rem !important;
        background: transparent !important;
    }
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
    }
    .app-header {
        font-family: "Meiryo", sans-serif;
        font-size: 22px;
        font-weight: bold;
        color: #1f2937;
        margin-top: 4px;
        margin-bottom: 10px;
        padding-bottom: 6px;
        border-bottom: 3px solid #2563eb;
        line-height: 1.4;
    }
    .yayoi-form-box {
        background-color: #f0f7f4;
        border: 2px solid #2e7d32;
        border-radius: 6px;
        padding: 20px;
        margin-bottom: 20px;
    }
    .yayoi-form-header {
        background-color: #2e7d32;
        color: white;
        padding: 8px 12px;
        font-weight: bold;
        font-size: 15px;
        border-radius: 4px 4px 0 0;
        margin: -20px -20px 20px -20px;
    }
</style>
""", unsafe_allow_html=True)

if os.path.exists("/content/drive/MyDrive/発注検討"):
    BASE_DIR = "/content/drive/MyDrive/発注検討"
else:
    BASE_DIR = "."

MAIN_EXCEL = os.path.join(BASE_DIR, "発注検討表.xlsx")
if not os.path.exists(MAIN_EXCEL):
    MAIN_EXCEL = os.path.join(BASE_DIR, "発注検討表-修正.xlsx")

YAYOI_EXCEL = os.path.join(BASE_DIR, "弥生確保リスト.xlsx")
YAYOI_ITEM_EXCEL = os.path.join(BASE_DIR, "弥生商品番号.xlsx")
YAYOI_COMP_EXCEL = os.path.join(BASE_DIR, "弥生会社名.xlsx")
SAVE_CSV = os.path.join(BASE_DIR, "meeting_decisions_v8.csv")
MANUAL_YAYOI_CSV = os.path.join(BASE_DIR, "yayoi_manual_additions.csv")

def safe_get_rows(ws):
    rows = []
    it = ws.iter_rows(values_only=True)
    while True:
        try:
            rows.append(next(it))
        except StopIteration:
            break
        except ValueError:
            break
    return rows

def is_na_val(val):
    if val is None:
        return True
    s = str(val).strip().lower()
    return s in ("#n/a", "#ref!", "#div/0!", "#value!", "none", "nan", "<na>")

def clean_str(val):
    if is_na_val(val):
        return ""
    return str(val).strip()

def to_num(val):
    try:
        if is_na_val(val) or str(val).strip() == "":
            return 0.0
        return float(val)
    except:
        return 0.0

@st.cache_data
def load_yayoi_data(file_source):
    df = pd.DataFrame()
    if os.path.exists(file_source):
        try:
            df = pd.read_excel(file_source, sheet_name=0, header=4)
        except Exception:
            pass
    
    if not df.empty and "商品コード" in df.columns:
        df = df[df["商品コード"].notna()].copy()
        df["商品コード"] = df["商品コード"].astype(str).str.strip()
        df["得意先名"] = df["得意先名"].apply(lambda x: clean_str(x) if clean_str(x) else "不明")
        df["受注残数"] = pd.to_numeric(df["受注残数"], errors="coerce").fillna(0).astype(int)
        df["数量"] = pd.to_numeric(df["数量"], errors="coerce").fillna(0).astype(int)
        df["受注日_dt"] = pd.to_datetime(df["受注日"], errors="coerce")
        df["納入期日_dt"] = pd.to_datetime(df["納入期日"], errors="coerce")
        df["受注日"] = df["受注日_dt"].dt.strftime("%Y-%m-%d").fillna("")
        df["納入期日"] = df["納入期日_dt"].dt.strftime("%Y-%m-%d").fillna("")
        df["備考"] = df["備考"].apply(clean_str)
        df["倉庫名"] = df["倉庫名"].apply(clean_str)
        df["商品名/摘要"] = df["商品名/摘要"].apply(clean_str)
        df["担当者名"] = df["担当者名"].apply(clean_str) if "担当者名" in df.columns else ""
    else:
        df = pd.DataFrame(columns=["得意先名", "商品コード", "受注残数", "数量", "受注日", "納入期日", "備考", "倉庫名", "担当者名", "受注日_dt", "納入期日_dt"])

    if os.path.exists(MANUAL_YAYOI_CSV):
        try:
            df_man = pd.read_csv(MANUAL_YAYOI_CSV).fillna("")
            if not df_man.empty:
                df_man["受注日_dt"] = pd.to_datetime(df_man["受注日"], errors="coerce")
                df_man["納入期日_dt"] = pd.to_datetime(df_man["納入期日"], errors="coerce")
                df = pd.concat([df, df_man], ignore_index=True)
        except Exception:
            pass

    if not df.empty:
        df = df.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
    return df

@st.cache_data
def load_sales_csvs(uploaded_files):
    if not uploaded_files:
        return pd.Series(dtype=int)
    
    all_dfs = []
    for f in uploaded_files:
        try:
            df_s = pd.read_csv(f, encoding="utf-8-sig")
        except Exception:
            try:
                df_s = pd.read_csv(f, encoding="shift_jis")
            except Exception:
                continue
        
        code_col = next((c for c in df_s.columns if any(k in str(c) for k in ["商品コード", "品番", "コード", "商品"])), None)
        qty_col = next((c for c in df_s.columns if any(k in str(c) for k in ["数量", "数", "残", "受注残"])), None)
        
        if code_col and qty_col:
            temp = pd.DataFrame()
            temp["商品コード"] = df_s[code_col].astype(str).str.strip()
            temp["数量"] = pd.to_numeric(df_s[qty_col], errors="coerce").fillna(0)
            all_dfs.append(temp)
            
    if not all_dfs:
        return pd.Series(dtype=int)
        
    combined = pd.concat(all_dfs, ignore_index=True)
    sales_summary = combined.groupby("商品コード")["数量"].sum().astype(int)
    return sales_summary

@st.cache_data
def load_excel_master(file_source):
    try:
        wb = openpyxl.load_workbook(file_source, read_only=True, data_only=True)
    except Exception:
        return pd.DataFrame(), pd.DataFrame(), "前回決定事項"
        
    target_sheet = "マスター (新) (3)" if "マスター (新) (3)" in wb.sheetnames else wb.sheetnames[0]
    if "マスター (新) (3)" not in wb.sheetnames and "80828" in wb.sheetnames:
        target_sheet = "80828"
    ws = wb[target_sheet]
    rows = safe_get_rows(ws)
    
    prev_col_title = "前回決定事項"
    if rows and len(rows[0]) > 28 and rows[0][28]:
        prev_col_title = clean_str(rows[0][28])

    dec_80828 = {}
    prev_80828 = {}
    if "80828" in wb.sheetnames:
        for r in safe_get_rows(wb["80828"])[1:]:
            if len(r) > 26 and r[2]:
                code_k = str(r[2]).strip()
                if r[25]:
                    prev_80828[code_k] = clean_str(r[25])
                if r[26]:
                    dec_80828[code_k] = clean_str(r[26])

    items = []
    for r in rows[1:]:
        if len(r) < 27 or r[2] is None:
            continue
        hinban = clean_str(r[2])
        if not hinban:
            continue
            
        na_flag = is_na_val(r[14]) or is_na_val(r[25]) or clean_str(r[12]) != ""
        
        prev_dec = clean_str(r[28]) if len(r) > 28 else ""
        if not prev_dec and hinban in prev_80828:
            prev_dec = prev_80828[hinban]
            
        curr_dec = clean_str(r[29]) if len(r) > 29 else ""
        if not curr_dec and hinban in dec_80828 and target_sheet == "80828":
            curr_dec = dec_80828[hinban]
            
        ndate = clean_str(r[18]) if len(r) > 18 else ""
        v_2ka = clean_str(r[15]) if len(r) > 15 and not is_na_val(r[15]) else ""
        v_rin = clean_str(r[27]) if len(r) > 27 and not is_na_val(r[27]) else ""

        items.append({
            "No": int(to_num(r[0])) if not is_na_val(r[0]) else "",
            "商品センター": clean_str(r[1]),
            "品番": hinban,
            "商品区分": clean_str(r[3]),
            "1課発注点": int(round(to_num(r[14]), 0)),
            "2課発注点": v_2ka,
            "在庫数_ロジザード": int(to_num(r[16])),
            "入荷予定数_ロジザード": int(to_num(r[17])),
            "入荷予定日_ロジザード": ndate,
            "確保数_ロジザード": int(to_num(r[19])),
            "在庫確保_弥生": int(to_num(r[20])),
            "在庫確保_丸藤": int(to_num(r[21])),
            "確保数のズレ": int(to_num(r[22])),
            "入荷後のフリー在庫": int(to_num(r[24])),
            "発注検討フラグ": int(round(to_num(r[25]), 0)),
            "在庫回転率": int(round(to_num(r[26]) * 100, 0)),
            "リンベル確保数": v_rin,
            "前回決定事項": prev_dec,
            "決定事項": curr_dec,
            "発注日": "",
            "NA行": na_flag
        })
    df_main = pd.DataFrame(items)

    date_sheets = [s for s in wb.sheetnames if s.isdigit() and len(s) == 5]
    hist_list = []
    for s in date_sheets:
        ws_d = wb[s]
        r_d = safe_get_rows(ws_d)
        if not r_d:
            continue
        header = r_d[0]
        h_col = next((i for i, c in enumerate(header) if c == "品番"), 2)
        k_cols = [i for i, c in enumerate(header) if c and "決定事項" in str(c) and "前回" not in str(c)]
        s_clean = s.lstrip("0")
        rest = s_clean[1:] if s_clean.startswith("8") else s_clean
        dt_str = f"2026-{rest[:2]}-{rest[2:]}" if len(rest) == 4 else (f"2026-0{rest[0]}-{rest[1:]}" if len(rest) == 3 else s)
        for r in r_d[1:]:
            if len(r) <= h_col or not r[h_col]:
                continue
            c_val = clean_str(r[h_col])
            for kc in k_cols:
                if kc < len(r) and clean_str(r[kc]) not in ("", "0"):
                    hist_list.append({"会議日": dt_str, "シート": s, "品番": c_val, "決定事項": clean_str(r[kc])})
    df_hist = pd.DataFrame(hist_list).drop_duplicates(subset=["会議日", "品番", "決定事項"]) if hist_list else pd.DataFrame(columns=["会議日", "シート", "品番", "決定事項"])
    return df_main, df_hist, prev_col_title

head_col1, head_col2 = st.columns([3, 2])
with head_col1:
    st.markdown('<div class="app-header">📊 発注検討表（Web管理システム）</div>', unsafe_allow_html=True)
with head_col2:
    with st.expander("📂 最新データ更新 ＆ 売れ数CSV読込"):
        if st.button("🔄 最新データを再読み込み"):
            st.cache_data.clear()
            if "master_df" in st.session_state:
                del st.session_state.master_df
            st.rerun()
        up_main = st.file_uploader("① 発注検討表.xlsx", type=["xlsx"], key="up_m")
        up_yayoi = st.file_uploader("② 弥生確保リスト.xlsx", type=["xlsx"], key="up_y")
        up_sales = st.file_uploader("③ 売上CSV群（ yayoi_xxxx.csv ）", type=["csv"], accept_multiple_files=True, key="up_s")

main_src = up_main if up_main is not None else (MAIN_EXCEL if os.path.exists(MAIN_EXCEL) else None)
yayoi_src = up_yayoi if up_yayoi is not None else (YAYOI_EXCEL if os.path.exists(YAYOI_EXCEL) else None)

if main_src is None:
    st.error("⚠️ メインの『発注検討表.xlsx』が見つかりません。")
    st.stop()

df_yayoi = load_yayoi_data(yayoi_src) if yayoi_src is not None else pd.DataFrame()
df_main, df_hist, prev_col_title = load_excel_master(main_src)

sales_series = load_sales_csvs(up_sales)
if not sales_series.empty:
    df_main["今月売れ数"] = df_main["品番"].map(sales_series).fillna(0).astype(int)
else:
    df_main["今月売れ数"] = 0

if df_main.empty:
    st.error("⚠️ 発注検討表のデータ読み込みに失敗しました。")
    st.stop()

if not df_yayoi.empty:
    yayoi_summary = df_yayoi[df_yayoi["受注残数"] > 0].groupby("商品コード").agg(
        弥生実データ合計=("受注残数", "sum"),
        確保会社数=("得意先名", "nunique"),
        確保会社一覧=("得意先名", lambda x: " / ".join(x.unique()))
    ).reset_index().rename(columns={"商品コード": "品番"})
    df_main = pd.merge(df_main, yayoi_summary, on="品番", how="left")
    df_main["弥生実データ合計"] = df_main["弥生実データ合計"].fillna(0).astype(int)
    df_main["確保会社数"] = df_main["確保会社数"].fillna(0).astype(int)
    df_main["確保会社一覧"] = df_main["確保会社一覧"].apply(clean_str)
else:
    df_main["弥生実データ合計"] = 0
    df_main["確保会社数"] = 0
    df_main["確保会社一覧"] = ""

if os.path.exists(SAVE_CSV):
    df_saved = pd.read_csv(SAVE_CSV).fillna("")
    for _, s_row in df_saved.iterrows():
        c = clean_str(s_row["品番"])
        if c in df_main["品番"].values:
            df_main.loc[df_main["品番"] == c, "入荷予定日_ロジザード"] = clean_str(s_row.get("入荷予定日_ロジザード", ""))
            df_main.loc[df_main["品番"] == c, "決定事項"] = clean_str(s_row.get("決定事項", ""))
            df_main.loc[df_main["品番"] == c, "発注日"] = clean_str(s_row.get("発注日", ""))

if "master_df" not in st.session_state or up_main is not None:
    st.session_state.master_df = df_main.copy()

m_df = st.session_state.master_df

yayoi_dict = {}
if not df_yayoi.empty:
    for code, grp in df_yayoi.groupby("商品コード", sort=False):
        grp_sorted = grp.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
        yayoi_dict[code] = grp_sorted[["得意先名", "受注残数", "数量", "受注日", "納入期日", "倉庫名", "備考", "担当者名"]].to_dict(orient="records")

hist_dict = {}
if not df_hist.empty:
    for code, grp in df_hist.sort_values("会議日", ascending=False).groupby("品番"):
        hist_dict[code] = grp[["会議日", "決定事項", "シート"]].to_dict(orient="records")

tab1, tab2, tab3, tab4 = st.tabs([
    "📋 1. 発注検討表（直接編集・クリックポップアップ）",
    "🏢 2. 会社（得意先）別 確保データ検索",
    "➕ 3. 新規確保データの登録（弥生風フォーム）",
    "🕒 4. 入荷予定・過去の会議決定履歴"
])

with tab1:
    f_col1, f_col2, f_col3 = st.columns([2.5, 1.5, 3])
    filter_mode = f_col1.selectbox(
        "表示絞り込み",
        ["すべて表示（Excelと同じ並び）", "🚨 要発注（発注検討フラグがマイナスのみ）", "🔥 【ヒット商品】在庫があるのに売れている", "確保数のズレがある商品", "入荷予定がある商品", "弥生確保がある商品"],
        index=0
    )
    center_opt = ["すべて"] + sorted([x for x in m_df["商品センター"].unique() if x])
    sel_center = f_col2.selectbox("商品センター", center_opt)
    search_kw = f_col3.text_input("🔍 品番・商品名・会社名・決定事項で検索", placeholder="例: AHL-2203 / ライト / 丸藤")

    v_df = m_df.copy()
    if filter_mode == "🚨 要発注（発注検討フラグがマイナスのみ）":
        v_df = v_df[(v_df["発注検討フラグ"] < 0) & (~v_df["NA行"])]
    elif filter_mode == "🔥 【ヒット商品】在庫があるのに売れている":
        v_df = v_df[(v_df["入荷後のフリー在庫"] >= 1) & (v_df["今月売れ数"] >= 1) & (~v_df["NA行"])]
    elif filter_mode == "確保数のズレがある商品":
        v_df = v_df[v_df["確保数のズレ"] != 0]
    elif filter_mode == "入荷予定がある商品":
        v_df = v_df[v_df["入荷予定数_ロジザード"] > 0]
    elif filter_mode == "弥生確保がある商品":
        v_df = v_df[(v_df["在庫確保_弥生"] > 0) | (v_df["在庫確保_丸藤"] > 0) | (v_df["弥生実データ合計"] > 0)]
    
    if sel_center != "すべて":
        v_df = v_df[v_df["商品センター"] == sel_center]
    if search_kw:
        v_df = v_df[
            v_df["品番"].str.contains(search_kw, case=False, na=False) |
            v_df["商品区分"].str.contains(search_kw, case=False, na=False) |
            v_df["確保会社一覧"].str.contains(search_kw, case=False, na=False) |
            v_df["前回決定事項"].str.contains(search_kw, case=False, na=False) |
            v_df["決定事項"].str.contains(search_kw, case=False, na=False)
        ]

    records_json = json.dumps(v_df.to_dict(orient="records"), ensure_ascii=False)
    yayoi_json = json.dumps(yayoi_dict, ensure_ascii=False)
    hist_json = json.dumps(hist_dict, ensure_ascii=False)
    prev_title_html = prev_col_title.replace("(", "<br>(")

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        * {{ box-sizing: border-box; font-family: "Meiryo", "Hiragino Kaku Gothic ProN", sans-serif; }}
        body {{ margin: 0; padding: 0; background: #fff; }}
        
        .toolbar {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            background: #f8fafc;
            padding: 8px 12px;
            border: 1px solid #cbd5e1;
            border-bottom: none;
            border-radius: 6px 6px 0 0;
        }}
        .toolbar-info {{ font-size: 13px; color: #334155; font-weight: bold; }}
        .toolbar-hint {{ font-size: 12px; color: #0369a1; background: #e0f2fe; padding: 4px 10px; border-radius: 4px; font-weight: bold; }}
        .btn-group {{ display: flex; gap: 8px; }}
        .btn {{
            padding: 6px 14px;
            font-size: 13px;
            font-weight: bold;
            border-radius: 4px;
            cursor: pointer;
            border: 1px solid transparent;
            transition: 0.15s;
        }}
        .btn-download {{ background: #2563eb; color: #fff; }}
        .btn-download:hover {{ background: #1d4ed8; }}
        
        .table-wrap {{
            width: 100%;
            height: 610px;
            overflow: auto;
            border: 1px solid #94a3b8;
            position: relative;
        }}
        table {{
            border-collapse: separate;
            border-spacing: 0;
            width: max-content;
            min-width: 100%;
        }}
        th, td {{
            border-right: 1px solid #94a3b8;
            border-bottom: 1px solid #94a3b8;
            padding: 4px 6px;
            font-size: 13px;
            height: 32px;
        }}
        
        thead th {{
            position: sticky;
            top: 0;
            background-color: #ffffff;
            color: #000;
            font-weight: bold;
            text-align: center;
            vertical-align: middle;
            line-height: 1.25;
            height: 72px;
            z-index: 10;
            border-bottom: 2px solid #475569;
            user-select: none;
        }}
        .red-txt {{ color: #e11d48; font-weight: bold; font-size: 11px; }}
        
        .sticky-col-1 {{ position: sticky; left: 0px; z-index: 5; background: #fff; }}
        .sticky-col-2 {{ position: sticky; left: 45px; z-index: 5; background: #fff; }}
        .sticky-col-3 {{ position: sticky; left: 115px; z-index: 5; background: #fff; }}
        .sticky-col-4 {{ position: sticky; left: 255px; z-index: 5; background: #fff; border-right: 2px solid #64748b; }}
        thead th.sticky-col-1, thead th.sticky-col-2, thead th.sticky-col-3, thead th.sticky-col-4 {{ z-index: 15; }}

        .w-no {{ width: 45px; min-width: 45px; max-width: 45px; text-align: center; white-space: nowrap; }}
        .w-center {{ width: 70px; min-width: 70px; max-width: 70px; text-align: center; white-space: nowrap; }}
        .w-code {{ width: 140px; min-width: 140px; max-width: 140px; text-align: left; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .w-name {{ width: 210px; min-width: 210px; max-width: 210px; text-align: left; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .w-num {{ width: 78px; min-width: 78px; max-width: 78px; text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }}
        .w-date {{ width: 105px; min-width: 105px; max-width: 105px; text-align: center; white-space: nowrap; }}
        .w-prev-dec {{ width: 170px; min-width: 170px; max-width: 170px; text-align: left; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .w-dec {{ width: 230px; min-width: 230px; max-width: 230px; text-align: left; }}

        .bg-cyan {{ background-color: #67e8f9 !important; }}
        .bg-flag-neg {{ background-color: #fecdd3 !important; color: #be123c; font-weight: bold; }}
        .bg-green {{ background-color: #a3e635 !important; }}
        .bg-pink-rate {{ background-color: #fbcfe8 !important; color: #be123c; font-weight: bold; }}
        .bg-yellow-ro {{ background-color: #fef08a !important; color: #1e293b; }}
        .bg-yellow-edit {{ background-color: #ffff00 !important; }}
        .bg-sales-hit {{ background-color: #fed7aa !important; font-weight: bold; color: #c2410c; }}
        
        tr.na-row td {{ background-color: #4b5563 !important; color: #f3f4f6 !important; }}
        tr.na-row td .click-link {{ color: #93c5fd !important; }}
        
        tbody tr:hover td {{ filter: brightness(0.95); }}

        .click-link {{
            color: #0284c7;
            font-weight: bold;
            text-decoration: underline;
            cursor: pointer;
            display: inline-block;
            width: 100%;
            white-space: nowrap;
        }}
        .click-link:hover {{ color: #0369a1; background-color: rgba(255,255,255,0.6); border-radius: 3px; }}
        
        .cell-input {{
            width: 100%;
            height: 24px;
            border: 1px solid #94a3b8;
            border-radius: 2px;
            padding: 2px 5px;
            font-size: 13px;
            background: rgba(255, 255, 255, 0.85);
        }}
        .cell-input:focus {{
            outline: 2px solid #2563eb;
            background: #fff;
        }}

        .modal-overlay {{
            display: none;
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(15, 23, 42, 0.55);
            z-index: 9999;
            justify-content: center;
            align-items: center;
        }}
        .modal-window {{
            background: #fff;
            width: 94%;
            max-width: 1060px;
            max-height: 88vh;
            border-radius: 8px;
            box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.25);
            display: flex;
            flex-direction: column;
            overflow: hidden;
            border: 2px solid #1e40af;
        }}
        .modal-header {{
            background: #1e3a8a;
            color: #fff;
            padding: 10px 18px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .modal-title {{ font-size: 16px; font-weight: bold; }}
        .modal-close {{
            background: #ef4444;
            color: #fff;
            border: none;
            border-radius: 4px;
            padding: 4px 12px;
            font-size: 13px;
            font-weight: bold;
            cursor: pointer;
        }}
        .modal-body {{
            padding: 14px 16px;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 12px;
        }}
        .summary-bar {{
            display: flex;
            gap: 14px;
            flex-wrap: wrap;
            background: #f1f5f9;
            padding: 8px 14px;
            border-radius: 6px;
            border-left: 5px solid #2563eb;
            font-size: 13px;
        }}
        .summary-item {{ font-weight: bold; color: #1e293b; }}
        .summary-item span {{ color: #0284c7; font-size: 15px; margin-left: 3px; }}
        
        .quick-edit-box {{
            background: #fffbeb;
            border: 1px solid #fde68a;
            padding: 10px 14px;
            border-radius: 6px;
            display: flex;
            flex-direction: column;
            gap: 8px;
        }}
        .stamp-btn {{
            background: #e2e8f0;
            border: 1px solid #94a3b8;
            border-radius: 4px;
            padding: 2px 8px;
            font-size: 12px;
            cursor: pointer;
            font-weight: bold;
        }}
        .stamp-btn:hover {{ background: #cbd5e1; }}

        .modal-grid {{
            display: grid;
            grid-template-columns: 3.2fr 1.8fr;
            gap: 16px;
        }}
        .sub-table {{ width: 100%; border-collapse: collapse; }}
        .sub-table th {{ background: #e2e8f0; color: #1e293b; padding: 6px; font-size: 12px; height: auto; position: static; border: 1px solid #cbd5e1; white-space: nowrap; }}
        .sub-table td {{ padding: 6px; font-size: 12px; height: auto; border: 1px solid #cbd5e1; }}
        .nowrap-cell {{ white-space: nowrap; text-align: center; }}
        .sec-title {{ font-size: 14px; font-weight: bold; color: #1e3a8a; margin-bottom: 6px; border-bottom: 2px solid #93c5fd; padding-bottom: 3px; }}
    </style>
    </head>
    <body>

    <div class="toolbar">
        <div class="toolbar-info">表示件数: <span id="row-count">0</span> 件</div>
        <div class="toolbar-hint">💡 品番・在庫確保数クリックでポップアップ表示 ｜ 売上CSV読込で「今月売れ数」を自動反映！</div>
        <div class="btn-group">
            <button class="btn btn-download" onclick="exportEditedCSV()">📥 編集内容を反映したCSVをダウンロード</button>
        </div>
    </div>

    <div class="table-wrap">
        <table id="excel-table">
            <thead>
                <tr>
                    <th class="w-no sticky-col-1">No</th>
                    <th class="w-center sticky-col-2">商品<br>センター</th>
                    <th class="w-code sticky-col-3">品番</th>
                    <th class="w-name sticky-col-4">商品区分</th>
                    <th class="w-num">1課発注<br>点</th>
                    <th class="w-num">2課発注<br>点</th>
                    <th class="w-num">在庫数<br><span class="red-txt">ロジザード</span></th>
                    <th class="w-num">入荷予定<br>数<br><span class="red-txt">ロジザード</span></th>
                    <th class="w-date">入荷予定<br>日<br><span class="red-txt">ロジザード</span></th>
                    <th class="w-num">確保数<br><span class="red-txt">ロジザード<br>(在庫確保)</span></th>
                    <th class="w-num">在庫確保<br><span class="red-txt">(弥生)</span></th>
                    <th class="w-num">在庫確保<br>丸藤</th>
                    <th class="w-num">確保数の<br>ズレ</th>
                    <th class="w-num">入荷後の<br>フリー在庫</th>
                    <th class="w-num">発注検討<br>フラグ<br><span class="red-txt">マイナスの<br>場合発注検討</span></th>
                    <th class="w-num">今月売れ数<br><span style="color:#c2410c; font-weight:bold;">🔥ヒット検出</span></th>
                    <th class="w-num">在庫回転<br>率</th>
                    <th class="w-num">リンベル<br>確保数</th>
                    <th class="w-prev-dec">{prev_title_html}</th>
                    <th class="w-dec">決定事項</th>
                    <th class="w-date">発注日</th>
                </tr>
            </thead>
            <tbody id="table-body"></tbody>
        </table>
    </div>

    <div class="modal-overlay" id="detail-modal" onclick="closeModalOnBg(event)">
        <div class="modal-window">
            <div class="modal-header">
                <div class="modal-title" id="modal-title-text">📦 品番詳細・弥生確保データ</div>
                <button class="modal-close" onclick="closeModal()">✕ 閉じる</button>
            </div>
            <div class="modal-body">
                <div class="summary-bar" id="modal-summary"></div>
                
                <div class="quick-edit-box">
                    <div style="font-size:12px; font-weight:bold; color:#92400e; display:flex; justify-content:space-between; align-items:center;">
                        <span>✏️ この場で決定事項・入荷予定日を入力（表にもリアルタイム連動します）</span>
                        <span>
                            ワンクリック入力:
                            <button class="stamp-btn" onclick="addStamp('様子見')">＋様子見</button>
                            <button class="stamp-btn" onclick="addStamp('発注済')">＋発注済</button>
                            <button class="stamp-btn" onclick="addStamp('確認中')">＋確認中</button>
                            <button class="stamp-btn" onclick="addStamp('廃盤予定')">＋廃盤予定</button>
                        </span>
                    </div>
                    <div style="display:flex; gap:10px; align-items:center;">
                        <label style="font-size:12px; font-weight:bold;">前回決定事項:</label>
                        <input type="text" id="modal-view-prev" class="cell-input" style="width:180px; background:#fef08a; color:#475569;" readonly>
                        <label style="font-size:12px; font-weight:bold;">入荷予定日:</label>
                        <input type="text" id="modal-input-ndate" class="cell-input" style="width:120px; background:#fff;" oninput="syncFromModal()">
                        <label style="font-size:12px; font-weight:bold;">決定事項:</label>
                        <input type="text" id="modal-input-dec" class="cell-input" style="flex:1; background:#fff;" oninput="syncFromModal()">
                        <label style="font-size:12px; font-weight:bold;">発注日:</label>
                        <input type="text" id="modal-input-odate" class="cell-input" style="width:110px; background:#fff;" oninput="syncFromModal()">
                    </div>
                </div>

                <div class="modal-grid">
                    <div>
                        <div class="sec-title">🏢 得意先別・弥生確保リスト（受注日の最新順）</div>
                        <div id="modal-yayoi-area"></div>
                    </div>
                    <div>
                        <div class="sec-title">📜 過去の会議決定履歴（最新順）</div>
                        <div id="modal-hist-area"></div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        const tableData = {records_json};
        const yayoiMap = {yayoi_json};
        const histMap = {hist_json};
        let currentModalIdx = null;

        const savedEdits = JSON.parse(localStorage.getItem("hattyu_edits_v1"] || "{{}}");
        tableData.forEach(r => {{
            if (savedEdits[r["品番"]]) {{
                r["入荷予定日_ロジザード"] = savedEdits[r["品番"]].ndate ?? r["入荷予定日_ロジザード"];
                r["決定事項"] = savedEdits[r["品番"]].dec ?? r["決定事項"];
                r["発注日"] = savedEdits[r["品番"]].odate ?? r["発注日"];
            }}
        }});

        function fmtNum(val) {{
            if (val === "" || val === null || val === undefined) return "";
            const n = Number(val);
            if (isNaN(n)) return val;
            return n.toLocaleString("ja-JP");
        }}

        function renderTable() {{
            const tbody = document.getElementById("table-body");
            document.getElementById("row-count").textContent = tableData.length.toLocaleString();
            let html = "";

            tableData.forEach((r, idx) => {{
                const rowCls = r["NA行"] ? "na-row" : "";
                const flagCls = (Number(r["発注検討フラグ"]) < 0 && !r["NA行"]) ? "bg-flag-neg" : "bg-cyan";
                const rateVal = Number(r["在庫回転率"]);
                const rateCls = (!r["NA行"] && rateVal <= 5) ? "bg-pink-rate" : "bg-green";
                const salesVal = Number(r["今月売れ数"] || 0);
                const salesCls = (salesVal >= 1 && Number(r["入荷後のフリー在庫"]) >= 1) ? "bg-sales-hit" : "";
                
                html += `<tr class="${{rowCls}}">
                    <td class="w-no sticky-col-1">${{r["No"]}}</td>
                    <td class="w-center sticky-col-2">${{r["商品センター"]}}</td>
                    <td class="w-code sticky-col-3" title="${{r["品番"]}}">
                        <span class="click-link" onclick="openModal(${{idx}})">${{r["品番"]}}</span>
                    </td>
                    <td class="w-name sticky-col-4" title="${{r["商品区分"]}}">${{r["商品区分"]}}</td>
                    <td class="w-num">${{r["NA行"] ? "#N/A" : fmtNum(r["1課発注点"])}}</td>
                    <td class="w-num">${{fmtNum(r["2課発注点"])}}</td>
                    <td class="w-num bg-cyan">${{fmtNum(r["在庫数_ロジザード"])}}</td>
                    <td class="w-num bg-cyan">${{fmtNum(r["入荷予定数_ロジザード"])}}</td>
                    <td class="w-date bg-cyan">
                        <input type="text" class="cell-input" id="ndate-${{idx}}" value="${{r["入荷予定日_ロジザード"] || ""}}" oninput="updateCell(${{idx}}, '入荷予定日_ロジザード', this.value)">
                    </td>
                    <td class="w-num bg-cyan">${{fmtNum(r["確保数_ロジザード"])}}</td>
                    <td class="w-num bg-cyan">
                        <span class="click-link" onclick="openModal(${{idx}})" title="クリックで弥生確保の会社内訳を表示">${{fmtNum(r["在庫確保_弥生"])}}</span>
                    </td>
                    <td class="w-num bg-cyan">
                        <span class="click-link" onclick="openModal(${{idx}})">${{fmtNum(r["在庫確保_丸藤"])}}</span>
                    </td>
                    <td class="w-num bg-cyan">${{fmtNum(r["確保数のズレ"])}}</td>
                    <td class="w-num bg-cyan">${{fmtNum(r["入荷後のフリー在庫"])}}</td>
                    <td class="w-num ${{flagCls}}">${{r["NA行"] ? "#N/A" : fmtNum(r["発注検討フラグ"])}}</td>
                    <td class="w-num ${{salesCls}}" title="今月売れ数">${{salesVal > 0 ? "🔥 " + fmtNum(salesVal) : "0"}}</td>
                    <td class="w-num ${{rateCls}}">${{r["NA行"] ? "#N/A" : fmtNum(r["在庫回転率"]) + "%"}}</td>
                    <td class="w-num">${{fmtNum(r["リンベル確保数"])}}</td>
                    <td class="w-prev-dec bg-yellow-ro" title="${{r["前回決定事項"] || ""}}">${{r["前回決定事項"] || ""}}</td>
                    <td class="w-dec bg-yellow-edit">
                        <input type="text" class="cell-input" id="dec-${{idx}}" value="${{r["決定事項"] || ""}}" oninput="updateCell(${{idx}}, '決定事項', this.value)">
                    </td>
                    <td class="w-date bg-yellow-edit">
                        <input type="text" class="cell-input" id="odate-${{idx}}" value="${{r["発注日"] || ""}}" oninput="updateCell(${{idx}}, '発注日', this.value)">
                    </td>
                </tr>`;
            }});
            tbody.innerHTML = html;
        }}

        function saveToStorage(hinban, rowObj) {{
            const store = JSON.parse(localStorage.getItem("hattyu_edits_v1") || "{{}}");
            store[hinban] = {{
                ndate: rowObj["入荷予定日_ロジザード"],
                dec: rowObj["決定事項"],
                odate: rowObj["発注日"]
            }};
            localStorage.setItem("hattyu_edits_v1", JSON.stringify(store));
        }}

        function updateCell(idx, field, val) {{
            tableData[idx][field] = val;
            saveToStorage(tableData[idx]["品番"], tableData[idx]);
        }}

        function openModal(idx) {{
            currentModalIdx = idx;
            const r = tableData[idx];
            const code = r["品番"];
            
            document.getElementById("modal-title-text").innerHTML = `📦 品番: 【${{code}}】 ${{r["商品区分"]}} の受注・確保詳細`;
            document.getElementById("modal-summary").innerHTML = `
                <div class="summary-item">商品センター: <span>${{r["商品センター"]}}</span></div>
                <div class="summary-item">1課発注点: <span>${{fmtNum(r["1課発注点"])}}</span></div>
                <div class="summary-item">ロジ在庫数: <span>${{fmtNum(r["在庫数_ロジザード"])}}</span></div>
                <div class="summary-item">ロジ確保数: <span>${{fmtNum(r["確保数_ロジザード"])}}</span></div>
                <div class="summary-item">今月売れ数: <span style="color:#c2410c;">${{fmtNum(r["今月売れ数"])}}</span></div>
                <div class="summary-item">フリー在庫: <span>${{fmtNum(r["入荷後のフリー在庫"])}}</span></div>
                <div class="summary-item">発注検討フラグ: <span style="color:#e11d48;">${{fmtNum(r["発注検討フラグ"])}}</span></div>
            `;

            document.getElementById("modal-view-prev").value = r["前回決定事項"] || "";
            document.getElementById("modal-input-ndate").value = r["入荷予定日_ロジザード"] || "";
            document.getElementById("modal-input-dec").value = r["決定事項"] || "";
            document.getElementById("modal-input-odate").value = r["発注日"] || "";

            const rawYList = yayoiMap[code] || [];
            const yList = [...rawYList].sort((a, b) => (b["受注日"] || "").localeCompare(a["受注日"] || ""));
            if (yList.length === 0) {{
                document.getElementById("modal-yayoi-area").innerHTML = `<div style="padding:12px; color:#64748b; background:#f8fafc; border-radius:4px;">現在、弥生確保リストに『${{code}}』の受注残データはありません。</div>`;
            }} else {{
                let yHtml = `<table class="sub-table"><thead><tr>
                    <th>得意先名</th><th>受注残数</th><th>受注数</th><th>受注日 (最新順)</th><th>納入期日</th><th>倉庫名</th><th>備考</th>
                </tr></thead><tbody>`;
                yList.forEach(y => {{
                    yHtml += `<tr>
                        <td style="font-weight:bold;">${{y["得意先名"]}}</td>
                        <td style="text-align:right; font-weight:bold; color:#0284c7;">${{fmtNum(y["受注残数"])}}</td>
                        <td style="text-align:right;">${{fmtNum(y["数量"])}}</td>
                        <td class="nowrap-cell" style="font-weight:bold;">${{y["受注日"]}}</td>
                        <td class="nowrap-cell">${{y["納入期日"]}}</td>
                        <td class="nowrap-cell">${{y["倉庫名"]}}</td>
                        <td>${{y["備考"]}}</td>
                    </tr>`;
                }});
                yHtml += `</tbody></table>`;
                document.getElementById("modal-yayoi-area").innerHTML = yHtml;
            }}

            const hList = histMap[code] || [];
            if (hList.length === 0) {{
                document.getElementById("modal-hist-area").innerHTML = `<div style="padding:12px; color:#64748b; background:#f8fafc; border-radius:4px;">過去の会議シートに決定事項の履歴はありません。</div>`;
            }} else {{
                let hHtml = `<table class="sub-table"><thead><tr>
                    <th style="width:95px;">会議日</th><th>決定事項</th>
                </tr></thead><tbody>`;
                hList.forEach(h => {{
                    hHtml += `<tr>
                        <td class="nowrap-cell">${{h["会議日"]}}</td>
                        <td style="font-weight:bold;">${{h["決定事項"]}}</td>
                    </tr>`;
                }});
                hHtml += `</tbody></table>`;
                document.getElementById("modal-hist-area").innerHTML = hHtml;
            }}

            document.getElementById("detail-modal").style.display = "flex";
        }}

        function syncFromModal() {{
            if (currentModalIdx === null) return;
            const ndate = document.getElementById("modal-input-ndate").value;
            const dec = document.getElementById("modal-input-dec").value;
            const odate = document.getElementById("modal-input-odate").value;

            tableData[currentModalIdx]["入荷予定日_ロジザード"] = ndate;
            tableData[currentModalIdx]["決定事項"] = dec;
            tableData[currentModalIdx]["発注日"] = odate;

            document.getElementById(`ndate-${{currentModalIdx}}`).value = ndate;
            document.getElementById(`dec-${{currentModalIdx}}`).value = dec;
            document.getElementById(`odate-${{currentModalIdx}}`).value = odate;

            saveToStorage(tableData[currentModalIdx]["品番"], tableData[currentModalIdx]);
        }}

        function addStamp(text) {{
            const decInput = document.getElementById("modal-input-dec");
            const curr = decInput.value.trim();
            decInput.value = curr ? `${{curr}} / ${{text}}` : text;
            syncFromModal();
        }}

        function closeModal() {{
            document.getElementById("detail-modal").style.display = "none";
            currentModalIdx = null;
        }}

        function closeModalOnBg(e) {{
            if (e.target.id === "detail-modal") closeModal();
        }}

        function exportEditedCSV() {{
            const cols = [
                "No", "商品センター", "品番", "商品区分", "1課発注点", "2課発注点",
                "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード",
                "確保数_ロジザード", "在庫確保_弥生", "在庫確保_丸藤", "確保数のズレ",
                "入荷後のフリー在庫", "発注検討フラグ", "今月売れ数", "在庫回転率", "リンベル確保数",
                "前回決定事項", "決定事項", "発注日"
            ];
            let csv = "\\uFEFF" + cols.join(",") + "\\n";
            tableData.forEach(r => {{
                const rowStr = cols.map(c => {{
                    let val = r[c] === undefined || r[c] === null ? "" : String(r[c]);
                    val = val.replace(/"/g, '""');
                    return `"${{val}}"`;
                }}).join(",");
                csv += rowStr + "\\n";
            }});
            const blob = new Blob([csv], {{ type: "text/csv;charset=utf-8;" }});
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `発注検討表_出力_${{new Date().toISOString().slice(0,10)}}.csv`;
            a.click();
        }}

        renderTable();
    </script>
    </body>
    </html>
    """

    components.html(html_code, height=675, scrolling=False)

with tab2:
    st.subheader("🏢 会社（得意先）別 確保データ検索 ＆ 月別スケジュール")
    if df_yayoi.empty:
        st.warning("『弥生確保リスト.xlsx』、または追加データがありません。")
    else:
        comp_totals = df_yayoi.groupby("得意先名")["受注残数"].sum().sort_values(ascending=False)
        comp_opts = ["すべての会社を表示"] + [f"{c} （確保合計: {v:,} 個）" for c, v in comp_totals.items()]
        c1, c2 = st.columns(2)
        sel_comp = c1.selectbox("🏢 会社名（得意先）をプルダウンで選択", comp_opts)
        kw_comp = c2.text_input("🔍 品番・備考でさらに絞り込み", placeholder="例: AHL-2203 / 11月")

        yv = df_yayoi.copy()
        if sel_comp != "すべての会社を表示":
            cname = sel_comp.split(" （確保合計:")[0]
            yv = yv[yv["得意先名"] == cname]
        if kw_comp:
            yv = yv[yv["得意先名"].str.contains(kw_comp, case=False, na=False) | yv["商品コード"].str.contains(kw_comp, case=False, na=False) | yv["備考"].str.contains(kw_comp, case=False, na=False)]

        yv_merged = pd.merge(yv, m_df[["品番", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "入荷後のフリー在庫", "発注検討フラグ", "前回決定事項", "決定事項"]], left_on="商品コード", right_on="品番", how="left")
        yv_merged = yv_merged.fillna("")
        for col in yv_merged.select_dtypes(include=['object']).columns:
            yv_merged[col] = yv_merged[col].apply(lambda x: "" if is_na_val(x) else str(x))

        if not yv_merged.empty and "納入期日" in yv_merged.columns:
            valid_dates = yv_merged[yv_merged["納入期日"] != ""]["納入期日"]
            if not valid_dates.empty:
                yv_merged["納入月"] = pd.to_datetime(valid_dates, errors="coerce").dt.strftime("%Y年%m月")
                monthly_sch = yv_merged[yv_merged["納入月"].notna()].groupby("納入月")["受注残数"].sum()
                if not monthly_sch.empty:
                    st.markdown("##### 📈 選択した会社の月別 納入スケジュール（受注残数合計）")
                    st.bar_chart(monthly_sch)

        st.dataframe(yv_merged[["得意先名", "商品コード", "商品名/摘要", "受注残数", "受注日", "納入期日", "備考", "倉庫名", "入荷予定日_ロジザード", "前回決定事項", "決定事項"]], hide_index=True, use_container_width=True, height=450)

with tab3:
    st.markdown("""
    <div class="yayoi-form-box">
        <div class="yayoi-form-header">💻 受注伝票入力（弥生販売風・新規確保データの登録）</div>
    """, unsafe_allow_html=True)
    
    # 1. 弥生会社名.xlsx から会社リストを読み込み
    company_options = []
    if os.path.exists(YAYOI_COMP_EXCEL):
        try:
            df_c_m = pd.read_excel(YAYOI_COMP_EXCEL, sheet_name=0)
            c_col = next((c for c in df_c_m.columns if any(k in str(c) for k in ["会社", "得意先", "名称", "顧客"])), df_c_m.columns[0])
            company_options = df_c_m[c_col].dropna().astype(str).str.strip().tolist()
        except Exception:
            pass
    if not company_options and not df_yayoi.empty:
        company_options = sorted(df_yayoi["得意先名"].unique().tolist())
    if not company_options:
        company_options = ["アマゾンベンダー（ホーム）", "株式会社オバタ", "丸藤", "トップ産業"]

    # 2. 弥生商品番号.xlsx から商品リストを読み込み
    item_options = []
    item_dict = {}
    if os.path.exists(YAYOI_ITEM_EXCEL):
        try:
            df_i_m = pd.read_excel(YAYOI_ITEM_EXCEL, sheet_name=0)
            code_col_m = next((c for c in df_i_m.columns if any(k in str(c) for k in ["コード", "品番", "商品コード"])), df_i_m.columns[0])
            name_col_m = next((c for c in df_i_m.columns if any(k in str(c) for k in ["名称", "商品名", "品名", "摘要"])), df_i_m.columns[1] if len(df_i_m.columns) > 1 else df_i_m.columns[0])
            for _, r in df_i_m.iterrows():
                c_val = str(r[code_col_m]).strip()
                n_val = str(r[name_col_m]).strip()
                if c_val and c_val != "nan":
                    item_options.append(f"{c_val} ｜ {n_val}")
                    item_dict[c_val] = n_val
        except Exception:
            pass
    if not item_options:
        item_options = [f"{row['品番']} ｜ {row['商品区分']}" for _, row in m_df.iterrows()]
        for _, row in m_df.iterrows():
            item_dict[row['品番']] = row['商品区分']

    with st.form("add_yayoi_form_yayoi_style"):
        # 自動重複なし受注番号の生成
        rand_suffix = random.randint(100, 999)
        next_order_no = f"WEB-{datetime.date.today().strftime('%Y%m%d')}-{rand_suffix}"
        
        f_h1, f_h2, f_h3 = st.columns(3)
        f_h1.text_input("受注番号 (自動発行)", value=next_order_no, disabled=True)
        add_date = f_h2.date_input("受注日", value=datetime.date.today())
        add_limit = f_h3.date_input("納入期日", value=datetime.date.today())

        f_d1, f_d2, f_d3 = st.columns(3)
        add_comp = f_d1.selectbox("得意先 (会社名)", options=company_options)
        
        selected_item_str = f_d2.selectbox("商品コード (品番 ｜ 商品名)", options=item_options)
        add_code = selected_item_str.split(" ｜ ")[0].strip() if " ｜ " in selected_item_str else selected_item_str.strip()
        
        add_qty = f_d3.number_input("受注残数 (数量)", min_value=1, value=1)

        f_e1, f_e2 = st.columns(2)
        staff_options = ["松平", "杉中", "田中", "鈴木", "佐藤"]
        add_staff = f_e1.selectbox("担当者", options=staff_options)
        add_memo = f_e2.text_input("摘要 / 備考", placeholder="例: コーポさっぽろ 12/3w")

        submitted = st.form_submit_button("💾 伝票を登録する（自動反映）", type="primary")
        if submitted:
            if not add_comp or not add_code:
                st.error("⚠️ 「得意先」と「商品コード」は必須です。")
            else:
                new_row = {
                    "受注番号": next_order_no,
                    "得意先名": str(add_comp).strip(),
                    "商品コード": str(add_code).strip(),
                    "受注残数": int(add_qty),
                    "数量": int(add_qty),
                    "受注日": add_date.strftime("%Y-%m-%d"),
                    "納入期日": add_limit.strftime("%Y-%m-%d"),
                    "受注日_dt": pd.to_datetime(add_date),
                    "納入期日_dt": pd.to_datetime(add_limit),
                    "備考": str(add_memo).strip(),
                    "倉庫名": "WEB入力",
                    "担当者名": str(add_staff).strip()
                }
                df_new = pd.DataFrame([new_row])
                if os.path.exists(MANUAL_YAYOI_CSV):
                    df_existing = pd.read_csv(MANUAL_YAYOI_CSV)
                    df_combined_man = pd.concat([df_existing, df_new], ignore_index=True)
                else:
                    df_combined_man = df_new
                df_combined_man.to_csv(MANUAL_YAYOI_CSV, index=False, encoding="utf-8-sig")
                
                st.cache_data.clear()
                if "master_df" in st.session_state:
                    del st.session_state.master_df
                st.success(f"✅ 受注番号『{next_order_no}』で登録しました！表を更新しています...")
                st.rerun()
                
    st.markdown("</div>", unsafe_allow_html=True)

with tab4:
    ch1, ch2 = st.columns(2)
    with ch1:
        st.subheader("🚢 入荷予定あり・発注済みの商品一覧")
        inc = m_df[(m_df["入荷予定数_ロジザード"] > 0) | (m_df["発注日"] != "") | (m_df["入荷予定日_ロジザード"] != "")].copy()
        st.dataframe(inc[["品番", "商品区分", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "発注日", "前回決定事項", "決定事項"]], hide_index=True, use_container_width=True, height=500)
    with ch2:
        st.subheader("📜 過去28回の会議シート 決定事項 全履歴")
        h_kw = st.text_input("品番・決定内容で検索", placeholder="例: 発注 / 様子見")
        hv = df_hist.sort_values(["会議日", "品番"], ascending=[False, True]) if not df_hist.empty else pd.DataFrame()
        if not hv.empty and h_kw:
            hv = hv[hv["品番"].str.contains(h_kw, case=False, na=False) | hv["決定事項"].str.contains(h_kw, case=False, na=False)]
        st.dataframe(hv, hide_index=True, use_container_width=True, height=450)
