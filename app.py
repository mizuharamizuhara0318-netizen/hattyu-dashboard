import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import datetime
import os
import glob
import json
import random

st.set_page_config(page_title="発注検討表 ＆ 弥生受注管理システム", layout="wide")

st.markdown("""
<style>
    header[data-testid="stHeader"] {
        height: 2rem !important;
        background: transparent !important;
    }
    .block-container {
        padding-top: 2.5rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.2rem !important;
        padding-right: 1.2rem !important;
        max-width: 100% !important;
    }
    .app-header {
        font-family: "Meiryo", sans-serif;
        font-size: 20px;
        font-weight: bold;
        color: #1f2937;
        margin-bottom: 8px;
        padding-bottom: 4px;
        border-bottom: 3px solid #2563eb;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 1. ファイルパスの自動探索（GitHub上・ローカルドライブ両対応）
# ---------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else "."

SEARCH_DIRS = [
    SCRIPT_DIR,
    ".",
    r"G:\マイドライブ\発注検討",
    "/content/drive/MyDrive/発注検討"
]

def find_target_file(keywords, ext=".xlsx"):
    for d in SEARCH_DIRS:
        if os.path.exists(d):
            try:
                for fn in os.listdir(d):
                    if fn.endswith(ext) and not fn.startswith("~$"):
                        if any(k in fn for k in keywords):
                            return os.path.join(d, fn)
            except Exception:
                pass
    return None

MAIN_EXCEL = find_target_file(["発注検討表"])
YAYOI_EXCEL = find_target_file(["弥生確保リスト", "確保リスト"])
YAYOI_ITEM_EXCEL = find_target_file(["弥生商品番号", "商品番号"])
YAYOI_COMP_EXCEL = find_target_file(["弥生会社名", "会社名"])

SAVE_CSV = os.path.join(SCRIPT_DIR, "meeting_decisions_v8.csv")
MANUAL_YAYOI_CSV = os.path.join(SCRIPT_DIR, "yayoi_manual_additions.csv")

# ---------------------------------------------------------
# 2. 補助関数（数値・文字列の安全な変換）
# ---------------------------------------------------------
def is_na_val(val):
    if val is None or pd.isna(val):
        return True
    s = str(val).strip().lower()
    return s in ("#n/a", "#ref!", "#div/0!", "#value!", "#name?", "none", "nan", "<na>", "")

def clean_str(val):
    if is_na_val(val):
        return ""
    s = str(val).strip()
    if s.endswith(".0") and s[:-2].replace("-", "").isdigit():
        return s[:-2]
    return s

def to_num(val):
    if is_na_val(val):
        return 0.0
    try:
        s = str(val).replace(",", "").replace("%", "").strip()
        return float(s)
    except Exception:
        return 0.0

# ---------------------------------------------------------
# 3. 確実なExcel・CSVデータ読み込み関数（0件バグ完全解消版）
# ---------------------------------------------------------
@st.cache_data
def load_excel_master(file_source):
    if file_source is None:
        return pd.DataFrame(), pd.DataFrame(), "前回決定事項"
    try:
        all_sheets = pd.read_excel(file_source, sheet_name=None, header=None)
    except Exception as e:
        st.error(f"Excel読み込みエラー: {e}")
        return pd.DataFrame(), pd.DataFrame(), "前回決定事項"

    target_df = None
    # シート名に「マスター」が含まれるものを最優先で探す
    for s_name, df_s in all_sheets.items():
        if "マスター" in str(s_name) and len(df_s) > 5:
            target_df = df_s
            break
    # なければ一番行数が多いシート（かつ品番データがあるもの）を自動選択
    if target_df is None:
        max_len = 0
        for s_name, df_s in all_sheets.items():
            if len(df_s) > max_len and df_s.shape[1] >= 20:
                max_len = len(df_s)
                target_df = df_s

    if target_df is None or target_df.empty:
        return pd.DataFrame(), pd.DataFrame(), "前回決定事項"

    # ヘッダー行（「品番」という文字がある行）を探す
    header_idx = 0
    for idx in range(min(15, len(target_df))):
        row_vals = [clean_str(x) for x in target_df.iloc[idx].values]
        if "品番" in row_vals:
            header_idx = idx
            break

    header_row = [clean_str(x) for x in target_df.iloc[header_idx].values]
    prev_col_title = "前回決定事項"
    if len(header_row) > 28 and header_row[28]:
        prev_col_title = header_row[28]

    items = []
    for idx in range(header_idx + 1, len(target_df)):
        r = target_df.iloc[idx].values
        if len(r) < 25:
            continue
        hinban = clean_str(r[2])
        if not hinban or hinban == "品番":
            continue

        raw_14 = str(r[14]).strip().lower() if len(r) > 14 and pd.notna(r[14]) else ""
        raw_25 = str(r[25]).strip().lower() if len(r) > 25 and pd.notna(r[25]) else ""
        na_flag = ("#n/a" in raw_14) or ("#n/a" in raw_25) or (clean_str(r[12]) != "" if len(r) > 12 else False)

        prev_dec = clean_str(r[28]) if len(r) > 28 else ""
        curr_dec = clean_str(r[29]) if len(r) > 29 else ""
        ndate = clean_str(r[18]) if len(r) > 18 else ""
        if "00:00:00" in ndate:
            ndate = ndate.split(" ")[0]

        v_2ka = clean_str(r[15]) if len(r) > 15 else ""
        v_rin = clean_str(r[27]) if len(r) > 27 else ""

        rate_raw = to_num(r[26]) if len(r) > 26 else 0.0
        rate_val = int(round(rate_raw * 100, 0)) if abs(rate_raw) <= 5.0 else int(round(rate_raw, 0))

        items.append({
            "No": int(to_num(r[0])) if not is_na_val(r[0]) else idx,
            "商品センター": clean_str(r[1]),
            "品番": hinban,
            "商品区分": clean_str(r[3]),
            "1課発注点": int(round(to_num(r[14]), 0)) if len(r) > 14 else 0,
            "2課発注点": v_2ka,
            "在庫数_ロジザード": int(to_num(r[16])) if len(r) > 16 else 0,
            "入荷予定数_ロジザード": int(to_num(r[17])) if len(r) > 17 else 0,
            "入荷予定日_ロジザード": ndate,
            "確保数_ロジザード": int(to_num(r[19])) if len(r) > 19 else 0,
            "在庫確保_弥生": int(to_num(r[20])) if len(r) > 20 else 0,
            "在庫確保_丸藤": int(to_num(r[21])) if len(r) > 21 else 0,
            "確保数のズレ": int(to_num(r[22])) if len(r) > 22 else 0,
            "入荷後のフリー在庫": int(to_num(r[24])) if len(r) > 24 else 0,
            "発注検討フラグ": int(round(to_num(r[25]), 0)) if len(r) > 25 else 0,
            "在庫回転率": rate_val,
            "リンベル確保数": v_rin,
            "前回決定事項": prev_dec,
            "決定事項": curr_dec,
            "発注日": "",
            "NA行": na_flag
        })

    df_main = pd.DataFrame(items)

    # 過去の会議履歴シートの抽出
    hist_list = []
    for s_name, df_s in all_sheets.items():
        s_str = str(s_name).strip()
        if s_str.isdigit() and len(s_str) == 5 and not df_s.empty:
            h_row = [clean_str(x) for x in df_s.iloc[0].values]
            h_col = next((i for i, c in enumerate(h_row) if c == "品番"), 2)
            k_cols = [i for i, c in enumerate(h_row) if c and "決定事項" in c and "前回" not in c]
            s_clean = s_str.lstrip("0")
            rest = s_clean[1:] if s_clean.startswith("8") else s_clean
            dt_str = f"2026-{rest[:2]}-{rest[2:]}" if len(rest) == 4 else (f"2026-0{rest[0]}-{rest[1:]}" if len(rest) == 3 else s_str)
            for _, r_d in df_s.iloc[1:].iterrows():
                r_vals = r_d.values
                if len(r_vals) <= h_col:
                    continue
                c_val = clean_str(r_vals[h_col])
                if not c_val:
                    continue
                for kc in k_cols:
                    if kc < len(r_vals) and clean_str(r_vals[kc]) not in ("", "0"):
                        hist_list.append({"会議日": dt_str, "シート": s_str, "品番": c_val, "決定事項": clean_str(r_vals[kc])})

    df_hist = pd.DataFrame(hist_list).drop_duplicates(subset=["会議日", "品番", "決定事項"]) if hist_list else pd.DataFrame(columns=["会議日", "シート", "品番", "決定事項"])
    return df_main, df_hist, prev_col_title

@st.cache_data
def load_yayoi_data(file_source):
    df = pd.DataFrame()
    if file_source is not None:
        try:
            df_raw = pd.read_excel(file_source, sheet_name=0, header=None)
            h_idx = 4
            for i in range(min(15, len(df_raw))):
                row_v = [clean_str(x) for x in df_raw.iloc[i].values]
                if "商品コード" in row_v or "得意先名" in row_v:
                    h_idx = i
                    break
            df = pd.read_excel(file_source, sheet_name=0, header=h_idx)
        except Exception:
            pass

    if not df.empty and "商品コード" in df.columns:
        df = df[df["商品コード"].notna()].copy()
        df["商品コード"] = df["商品コード"].apply(clean_str)
        df = df[df["商品コード"] != ""]
        df["受注番号"] = df["受注番号"].apply(clean_str) if "受注番号" in df.columns else "YAYOI-EX"
        df["得意先名"] = df["得意先名"].apply(lambda x: clean_str(x) if clean_str(x) else "不明") if "得意先名" in df.columns else "不明"
        df["受注残数"] = pd.to_numeric(df["受注残数"], errors="coerce").fillna(0).astype(int) if "受注残数" in df.columns else 0
        df["数量"] = pd.to_numeric(df["数量"], errors="coerce").fillna(0).astype(int) if "数量" in df.columns else df["受注残数"]
        df["受注日_dt"] = pd.to_datetime(df["受注日"], errors="coerce") if "受注日" in df.columns else pd.NaT
        df["納入期日_dt"] = pd.to_datetime(df["納入期日"], errors="coerce") if "納入期日" in df.columns else pd.NaT
        df["受注日"] = df["受注日_dt"].dt.strftime("%Y-%m-%d").fillna("")
        df["納入期日"] = df["納入期日_dt"].dt.strftime("%Y-%m-%d").fillna("")
        df["備考"] = df["備考"].apply(clean_str) if "備考" in df.columns else ""
        df["倉庫名"] = df["倉庫名"].apply(clean_str) if "倉庫名" in df.columns else ""
        df["商品名/摘要"] = df["商品名/摘要"].apply(clean_str) if "商品名/摘要" in df.columns else ""
        df["担当者名"] = df["担当者名"].apply(clean_str) if "担当者名" in df.columns else ""
        df["単価"] = pd.to_numeric(df["単価"], errors="coerce").fillna(0).astype(int) if "単価" in df.columns else 0
    else:
        df = pd.DataFrame(columns=["受注番号", "得意先名", "商品コード", "商品名/摘要", "受注残数", "数量", "単価", "受注日", "納入期日", "備考", "倉庫名", "担当者名", "受注日_dt", "納入期日_dt"])

    if os.path.exists(MANUAL_YAYOI_CSV):
        try:
            df_man = pd.read_csv(MANUAL_YAYOI_CSV).fillna("")
            if not df_man.empty:
                df_man["受注残数"] = pd.to_numeric(df_man["受注残数"], errors="coerce").fillna(0).astype(int)
                df_man["数量"] = pd.to_numeric(df_man["数量"], errors="coerce").fillna(0).astype(int)
                df_man["単価"] = pd.to_numeric(df_man.get("単価", 0), errors="coerce").fillna(0).astype(int)
                df_man["受注日_dt"] = pd.to_datetime(df_man["受注日"], errors="coerce")
                df_man["納入期日_dt"] = pd.to_datetime(df_man["納入期日"], errors="coerce")
                for col in ["受注番号", "得意先名", "商品コード", "商品名/摘要", "備考", "倉庫名", "担当者名"]:
                    if col not in df_man.columns:
                        df_man[col] = ""
                df = pd.concat([df, df_man], ignore_index=True)
        except Exception:
            pass

    if not df.empty:
        df = df.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
    return df

@st.cache_data
def load_sales_csvs(uploaded_files):
    all_dfs = []
    
    # 1. ローカルドライブ（G:\マイドライブ\売上ダッシュボード\済）が存在する場合は自動スキャン
    local_sales_dirs = [
        r"G:\マイドライブ\売上ダッシュボード\済",
        r"G:\マイドライブ\売上ダッシュボード",
        "/content/drive/MyDrive/売上ダッシュボード/済"
    ]
    local_csv_paths = []
    for d in local_sales_dirs:
        if os.path.exists(d):
            local_csv_paths.extend(glob.glob(os.path.join(d, "yayoi_*.csv")))
            break

    sources = uploaded_files if uploaded_files else local_csv_paths
    if not sources:
        return pd.Series(dtype=int)

    for f in sources:
        df_s = None
        for enc in ["cp932", "shift_jis", "utf-8-sig", "utf-8"]:
            try:
                if isinstance(f, str):
                    df_s = pd.read_csv(f, encoding=enc)
                else:
                    f.seek(0)
                    df_s = pd.read_csv(f, encoding=enc)
                if not df_s.empty:
                    break
            except Exception:
                continue

        if df_s is None or df_s.empty:
            continue

        code_col = next((c for c in df_s.columns if any(k in str(c) for k in ["商品コード", "品番", "コード", "商品"])), None)
        qty_col = next((c for c in df_s.columns if any(k in str(c) for k in ["数量", "売上数量", "数", "受注数", "残"])), None)

        if code_col and qty_col:
            temp = pd.DataFrame()
            temp["商品コード"] = df_s[code_col].apply(clean_str)
            temp["数量"] = pd.to_numeric(df_s[qty_col], errors="coerce").fillna(0)
            all_dfs.append(temp)
        else:
            # ヘッダー無しの弥生インポート形式CSV等の場合、全列から英数字ハイフンの品番列と数値列を自動推定
            for col_i in range(df_s.shape[1]):
                sample_str = df_s.iloc[:, col_i].dropna().astype(str).str.strip()
                if sample_str.str.contains(r"^[A-Z]{2,4}-[0-9A-Z]+", regex=True).any():
                    temp = pd.DataFrame()
                    temp["商品コード"] = df_s.iloc[:, col_i].apply(clean_str)
                    # その右側の列から数量らしき列を探す
                    for q_i in range(col_i + 1, min(col_i + 8, df_s.shape[1])):
                        q_num = pd.to_numeric(df_s.iloc[:, q_i], errors="coerce").fillna(0)
                        if (q_num > 0).any() and q_num.max() < 100000:
                            temp["数量"] = q_num
                            all_dfs.append(temp)
                            break
                    break

    if not all_dfs:
        return pd.Series(dtype=int)

    combined = pd.concat(all_dfs, ignore_index=True)
    combined = combined[combined["商品コード"] != ""]
    return combined.groupby("商品コード")["数量"].sum().astype(int)

# ---------------------------------------------------------
# 4. 上部ヘッダー ＆ データ更新・売上CSV読込バー
# ---------------------------------------------------------
head_col1, head_col2 = st.columns([3.2, 2.0])
with head_col1:
    st.markdown('<div class="app-header">📊 発注検討表 ＆ 弥生受注伝票システム（完全連動版）</div>', unsafe_allow_html=True)
with head_col2:
    with st.expander("📂 最新データの差し替え ＆ 売上CSV（売れ数）読込"):
        if st.button("🔄 データを最新状態に再読み込み", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
        up_sales = st.file_uploader("① 今月の売上CSV（ yayoi_xxxx.csv を複数まとめて選択可）", type=["csv"], accept_multiple_files=True, key="up_s")
        up_main = st.file_uploader("② 発注検討表.xlsx（差し替える場合のみ）", type=["xlsx"], key="up_m")
        up_yayoi = st.file_uploader("③ 弥生確保リスト.xlsx（差し替える場合のみ）", type=["xlsx"], key="up_y")

main_src = up_main if up_main is not None else MAIN_EXCEL
yayoi_src = up_yayoi if up_yayoi is not None else YAYOI_EXCEL

df_main, df_hist, prev_col_title = load_excel_master(main_src)
df_yayoi = load_yayoi_data(yayoi_src)

if df_main.empty:
    st.error("⚠️ 『発注検討表.xlsx』のデータが見つからないか読み込めませんでした。右上の「最新データの差し替え」から『発注検討表.xlsx』をアップロードしてください。")
    st.stop()

# 売上CSVの集計マッピング
sales_series = load_sales_csvs(up_sales)
if not sales_series.empty:
    df_main["今月売れ数"] = df_main["品番"].map(sales_series).fillna(0).astype(int)
else:
    df_main["今月売れ数"] = 0

# 弥生確保データの集計＆発注検討表へのリアルタイム連動（手動登録分も自動加算）
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
    # WEB伝票で追加された場合、弥生実データ合計が元の数値を上回れば自動反映
    df_main["在庫確保_弥生"] = df_main[["在庫確保_弥生", "弥生実データ合計"]].max(axis=1)
else:
    df_main["弥生実データ合計"] = 0
    df_main["確保会社数"] = 0
    df_main["確保会社一覧"] = ""

if os.path.exists(SAVE_CSV):
    try:
        df_saved = pd.read_csv(SAVE_CSV).fillna("")
        for _, s_row in df_saved.iterrows():
            c = clean_str(s_row["品番"])
            if c in df_main["品番"].values:
                df_main.loc[df_main["品番"] == c, "入荷予定日_ロジザード"] = clean_str(s_row.get("入荷予定日_ロジザード", ""))
                df_main.loc[df_main["品番"] == c, "決定事項"] = clean_str(s_row.get("決定事項", ""))
                df_main.loc[df_main["品番"] == c, "発注日"] = clean_str(s_row.get("発注日", ""))
    except Exception:
        pass

m_df = df_main.copy()

yayoi_dict = {}
if not df_yayoi.empty:
    for code, grp in df_yayoi.groupby("商品コード", sort=False):
        grp_sorted = grp.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
        yayoi_dict[code] = grp_sorted[["受注番号", "得意先名", "受注残数", "数量", "受注日", "納入期日", "倉庫名", "備考", "担当者名"]].to_dict(orient="records")

hist_dict = {}
if not df_hist.empty:
    for code, grp in df_hist.sort_values("会議日", ascending=False).groupby("品番"):
        hist_dict[code] = grp[["会議日", "決定事項", "シート"]].to_dict(orient="records")

# ---------------------------------------------------------
# 5. メインタブ構成
# ---------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    f"📋 1. 発注検討表（全 {len(m_df):,} 件・クリック即ポップアップ）",
    "🏢 2. 会社（得意先）別 確保データ検索",
    "💻 3. 弥生受注伝票（新規登録 ＆ 検索ダブルクリック修正）",
    "🕒 4. 入荷予定・過去の会議決定履歴"
])

# =========================================================
# タブ1：発注検討表（Excel完全再現・直接編集・ポップアップ）
# =========================================================
with tab1:
    f_col1, f_col2, f_col3 = st.columns([2.5, 1.3, 3.0])
    filter_mode = f_col1.selectbox(
        "表示絞り込み",
        [
            "すべて表示（Excelと同じ並び）",
            "🚨 要発注（発注検討フラグがマイナスのみ）",
            "🔥 【ヒット商品】在庫があるのに今月売れている",
            "確保数のズレがある商品",
            "入荷予定がある商品",
            "弥生確保がある商品"
        ],
        index=0
    )
    center_opt = ["すべて"] + sorted([x for x in m_df["商品センター"].unique() if x])
    sel_center = f_col2.selectbox("商品センター", center_opt)
    search_kw = f_col3.text_input("🔍 品番・商品名・会社名・決定事項で検索", placeholder="例: AHL-2203 / ライト / 丸藤")

    v_df = m_df.copy()
    if filter_mode == "🚨 要発注（発注検討フラグがマイナスのみ）":
        v_df = v_df[(v_df["発注検討フラグ"] < 0) & (~v_df["NA行"])]
    elif filter_mode == "🔥 【ヒット商品】在庫があるのに今月売れている":
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

    html_tab1 = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        * {{ box-sizing: border-box; font-family: "Meiryo", "Hiragino Kaku Gothic ProN", sans-serif; }}
        body {{ margin: 0; padding: 0; background: #fff; }}
        .toolbar {{
            display: flex; align-items: center; justify-content: space-between;
            background: #f8fafc; padding: 6px 12px; border: 1px solid #cbd5e1;
            border-bottom: none; border-radius: 6px 6px 0 0;
        }}
        .toolbar-info {{ font-size: 13px; color: #1e293b; font-weight: bold; }}
        .toolbar-hint {{ font-size: 12px; color: #0369a1; background: #e0f2fe; padding: 4px 10px; border-radius: 4px; font-weight: bold; }}
        .btn-download {{
            padding: 6px 14px; font-size: 12px; font-weight: bold; border-radius: 4px;
            cursor: pointer; border: none; background: #2563eb; color: #fff;
        }}
        .btn-download:hover {{ background: #1d4ed8; }}
        .table-wrap {{ width: 100%; height: 600px; overflow: auto; border: 1px solid #94a3b8; position: relative; }}
        table {{ border-collapse: separate; border-spacing: 0; width: max-content; min-width: 100%; }}
        th, td {{ border-right: 1px solid #94a3b8; border-bottom: 1px solid #94a3b8; padding: 3px 6px; font-size: 12.5px; height: 30px; }}
        thead th {{
            position: sticky; top: 0; background-color: #ffffff; color: #000;
            font-weight: bold; text-align: center; vertical-align: middle;
            line-height: 1.25; height: 68px; z-index: 10; border-bottom: 2px solid #475569;
        }}
        .red-txt {{ color: #e11d48; font-weight: bold; font-size: 10.5px; }}
        .sticky-col-1 {{ position: sticky; left: 0px; z-index: 5; background: #fff; }}
        .sticky-col-2 {{ position: sticky; left: 45px; z-index: 5; background: #fff; }}
        .sticky-col-3 {{ position: sticky; left: 110px; z-index: 5; background: #fff; }}
        .sticky-col-4 {{ position: sticky; left: 245px; z-index: 5; background: #fff; border-right: 2px solid #64748b; }}
        thead th.sticky-col-1, thead th.sticky-col-2, thead th.sticky-col-3, thead th.sticky-col-4 {{ z-index: 15; }}

        .w-no {{ width: 45px; min-width: 45px; text-align: center; }}
        .w-center {{ width: 65px; min-width: 65px; text-align: center; }}
        .w-code {{ width: 135px; min-width: 135px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .w-name {{ width: 210px; min-width: 210px; max-width: 210px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .w-num {{ width: 74px; min-width: 74px; text-align: right; font-variant-numeric: tabular-nums; }}
        .w-date {{ width: 100px; min-width: 100px; text-align: center; }}
        .w-prev-dec {{ width: 165px; min-width: 165px; max-width: 165px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .w-dec {{ width: 220px; min-width: 220px; }}

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

        .click-link {{ color: #0284c7; font-weight: bold; text-decoration: underline; cursor: pointer; display: block; width: 100%; }}
        .cell-input {{ width: 100%; height: 23px; border: 1px solid #94a3b8; border-radius: 2px; padding: 1px 4px; font-size: 12px; background: rgba(255,255,255,0.9); }}

        .modal-overlay {{
            display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(15, 23, 42, 0.55); z-index: 9999; justify-content: center; align-items: center;
        }}
        .modal-window {{
            background: #fff; width: 95%; max-width: 1080px; max-height: 88vh; border-radius: 8px;
            box-shadow: 0 20px 25px -5px rgba(0,0,0,0.3); display: flex; flex-direction: column; overflow: hidden; border: 2px solid #1e40af;
        }}
        .modal-header {{ background: #1e3a8a; color: #fff; padding: 10px 16px; display: flex; justify-content: space-between; align-items: center; }}
        .modal-title {{ font-size: 15px; font-weight: bold; }}
        .modal-close {{ background: #ef4444; color: #fff; border: none; border-radius: 4px; padding: 4px 12px; font-size: 12px; font-weight: bold; cursor: pointer; }}
        .modal-body {{ padding: 14px; overflow-y: auto; display: flex; flex-direction: column; gap: 12px; }}
        .summary-bar {{ display: flex; gap: 14px; flex-wrap: wrap; background: #f1f5f9; padding: 8px 12px; border-radius: 6px; border-left: 5px solid #2563eb; font-size: 12.5px; }}
        .summary-item {{ font-weight: bold; color: #1e293b; }}
        .summary-item span {{ color: #0284c7; font-size: 14px; margin-left: 3px; }}
        .quick-edit-box {{ background: #fffbeb; border: 1px solid #fde68a; padding: 10px 12px; border-radius: 6px; display: flex; flex-direction: column; gap: 8px; }}
        .stamp-btn {{ background: #e2e8f0; border: 1px solid #94a3b8; border-radius: 4px; padding: 2px 8px; font-size: 11.5px; cursor: pointer; font-weight: bold; }}
        .modal-grid {{ display: grid; grid-template-columns: 3.3fr 1.7fr; gap: 14px; }}
        .sub-table {{ width: 100%; border-collapse: collapse; }}
        .sub-table th {{ background: #e2e8f0; color: #1e293b; padding: 5px; font-size: 11.5px; height: auto; position: static; border: 1px solid #cbd5e1; }}
        .sub-table td {{ padding: 5px; font-size: 11.5px; height: auto; border: 1px solid #cbd5e1; }}
        .sec-title {{ font-size: 13.5px; font-weight: bold; color: #1e3a8a; margin-bottom: 6px; border-bottom: 2px solid #93c5fd; padding-bottom: 3px; }}
    </style>
    </head>
    <body>
    <div class="toolbar">
        <div class="toolbar-info">表示件数: <span id="row-count">0</span> 件</div>
        <div class="toolbar-hint">💡 品番・在庫確保(弥生)クリックで内訳ポップアップ ｜ 売上CSV読込で「今月売れ数」を自動計算！</div>
        <button class="btn-download" onclick="exportEditedCSV()">📥 編集内容を反映したCSV出力</button>
    </div>
    <div class="table-wrap">
        <table>
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
                    <div style="font-size:12px; font-weight:bold; color:#92400e; display:flex; justify-content:space-between;">
                        <span>✏️ この場で決定事項・入荷予定日を入力（表にもリアルタイム連動）</span>
                        <span>
                            スタンプ:
                            <button class="stamp-btn" onclick="addStamp('様子見')">＋様子見</button>
                            <button class="stamp-btn" onclick="addStamp('発注済')">＋発注済</button>
                            <button class="stamp-btn" onclick="addStamp('確認中')">＋確認中</button>
                            <button class="stamp-btn" onclick="addStamp('廃盤予定')">＋廃盤予定</button>
                        </span>
                    </div>
                    <div style="display:flex; gap:8px; align-items:center;">
                        <label style="font-size:12px; font-weight:bold;">前回:</label>
                        <input type="text" id="modal-view-prev" class="cell-input" style="width:160px; background:#fef08a;" readonly>
                        <label style="font-size:12px; font-weight:bold;">入荷予定日:</label>
                        <input type="text" id="modal-input-ndate" class="cell-input" style="width:110px;" oninput="syncFromModal()">
                        <label style="font-size:12px; font-weight:bold;">決定事項:</label>
                        <input type="text" id="modal-input-dec" class="cell-input" style="flex:1;" oninput="syncFromModal()">
                        <label style="font-size:12px; font-weight:bold;">発注日:</label>
                        <input type="text" id="modal-input-odate" class="cell-input" style="width:100px;" oninput="syncFromModal()">
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

        const savedEdits = JSON.parse(localStorage.getItem("hattyu_edits_v1") || "{{}}");
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
            return isNaN(n) ? val : n.toLocaleString("ja-JP");
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
                    <td class="w-code sticky-col-3" title="${{r["品番"]}}"><span class="click-link" onclick="openModal(${{idx}})">${{r["品番"]}}</span></td>
                    <td class="w-name sticky-col-4" title="${{r["商品区分"]}}">${{r["商品区分"]}}</td>
                    <td class="w-num">${{r["NA行"] ? "#N/A" : fmtNum(r["1課発注点"])}}</td>
                    <td class="w-num">${{fmtNum(r["2課発注点"])}}</td>
                    <td class="w-num bg-cyan">${{fmtNum(r["在庫数_ロジザード"])}}</td>
                    <td class="w-num bg-cyan">${{fmtNum(r["入荷予定数_ロジザード"])}}</td>
                    <td class="w-date bg-cyan"><input type="text" class="cell-input" id="ndate-${{idx}}" value="${{r["入荷予定日_ロジザード"] || ""}}" oninput="updateCell(${{idx}}, '入荷予定日_ロジザード', this.value)"></td>
                    <td class="w-num bg-cyan">${{fmtNum(r["確保数_ロジザード"])}}</td>
                    <td class="w-num bg-cyan"><span class="click-link" onclick="openModal(${{idx}})">${{fmtNum(r["在庫確保_弥生"])}}</span></td>
                    <td class="w-num bg-cyan"><span class="click-link" onclick="openModal(${{idx}})">${{fmtNum(r["在庫確保_丸藤"])}}</span></td>
                    <td class="w-num bg-cyan">${{fmtNum(r["確保数のズレ"])}}</td>
                    <td class="w-num bg-cyan">${{fmtNum(r["入荷後のフリー在庫"])}}</td>
                    <td class="w-num ${{flagCls}}">${{r["NA行"] ? "#N/A" : fmtNum(r["発注検討フラグ"])}}</td>
                    <td class="w-num ${{salesCls}}">${{salesVal > 0 ? "🔥 " + fmtNum(salesVal) : "0"}}</td>
                    <td class="w-num ${{rateCls}}">${{r["NA行"] ? "#N/A" : fmtNum(r["在庫回転率"]) + "%"}}</td>
                    <td class="w-num">${{fmtNum(r["リンベル確保数"])}}</td>
                    <td class="w-prev-dec bg-yellow-ro" title="${{r["前回決定事項"] || ""}}">${{r["前回決定事項"] || ""}}</td>
                    <td class="w-dec bg-yellow-edit"><input type="text" class="cell-input" id="dec-${{idx}}" value="${{r["決定事項"] || ""}}" oninput="updateCell(${{idx}}, '決定事項', this.value)"></td>
                    <td class="w-date bg-yellow-edit"><input type="text" class="cell-input" id="odate-${{idx}}" value="${{r["発注日"] || ""}}" oninput="updateCell(${{idx}}, '発注日', this.value)"></td>
                </tr>`;
            }});
            tbody.innerHTML = html;
        }}

        function saveToStorage(hinban, rowObj) {{
            const store = JSON.parse(localStorage.getItem("hattyu_edits_v1") || "{{}}");
            store[hinban] = {{ ndate: rowObj["入荷予定日_ロジザード"], dec: rowObj["決定事項"], odate: rowObj["発注日"] }};
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
                <div class="summary-item">センター: <span>${{r["商品センター"]}}</span></div>
                <div class="summary-item">1課発注点: <span>${{fmtNum(r["1課発注点"])}}</span></div>
                <div class="summary-item">ロジ在庫: <span>${{fmtNum(r["在庫数_ロジザード"])}}</span></div>
                <div class="summary-item">弥生確保: <span>${{fmtNum(r["在庫確保_弥生"])}}</span></div>
                <div class="summary-item">今月売れ数: <span style="color:#c2410c;">${{fmtNum(r["今月売れ数"])}}</span></div>
                <div class="summary-item">フリー在庫: <span>${{fmtNum(r["入荷後のフリー在庫"])}}</span></div>
                <div class="summary-item">検討フラグ: <span style="color:#e11d48;">${{fmtNum(r["発注検討フラグ"])}}</span></div>
            `;
            document.getElementById("modal-view-prev").value = r["前回決定事項"] || "";
            document.getElementById("modal-input-ndate").value = r["入荷予定日_ロジザード"] || "";
            document.getElementById("modal-input-dec").value = r["決定事項"] || "";
            document.getElementById("modal-input-odate").value = r["発注日"] || "";

            const rawYList = yayoiMap[code] || [];
            const yList = [...rawYList].sort((a, b) => (b["受注日"] || "").localeCompare(a["受注日"] || ""));
            if (yList.length === 0) {{
                document.getElementById("modal-yayoi-area").innerHTML = `<div style="padding:10px; color:#64748b; background:#f8fafc;">現在、弥生確保リストに『${{code}}』の受注残データはありません。</div>`;
            }} else {{
                let yHtml = `<table class="sub-table"><thead><tr><th>受注番号</th><th>得意先名</th><th>受注残数</th><th>受注日</th><th>納入期日</th><th>担当者</th><th>備考</th></tr></thead><tbody>`;
                yList.forEach(y => {{
                    yHtml += `<tr>
                        <td style="text-align:center;">${{y["受注番号"] || ""}}</td>
                        <td style="font-weight:bold;">${{y["得意先名"]}}</td>
                        <td style="text-align:right; font-weight:bold; color:#0284c7;">${{fmtNum(y["受注残数"])}}</td>
                        <td style="text-align:center;">${{y["受注日"]}}</td>
                        <td style="text-align:center;">${{y["納入期日"]}}</td>
                        <td style="text-align:center;">${{y["担当者名"] || ""}}</td>
                        <td>${{y["備考"]}}</td>
                    </tr>`;
                }});
                yHtml += `</tbody></table>`;
                document.getElementById("modal-yayoi-area").innerHTML = yHtml;
            }}

            const hList = histMap[code] || [];
            if (hList.length === 0) {{
                document.getElementById("modal-hist-area").innerHTML = `<div style="padding:10px; color:#64748b; background:#f8fafc;">過去の会議決定履歴はありません。</div>`;
            }} else {{
                let hHtml = `<table class="sub-table"><thead><tr><th style="width:90px;">会議日</th><th>決定事項</th></tr></thead><tbody>`;
                hList.forEach(h => {{
                    hHtml += `<tr><td style="text-align:center;">${{h["会議日"]}}</td><td style="font-weight:bold;">${{h["決定事項"]}}</td></tr>`;
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

        function closeModal() {{ document.getElementById("detail-modal").style.display = "none"; currentModalIdx = null; }}
        function closeModalOnBg(e) {{ if (e.target.id === "detail-modal") closeModal(); }}

        function exportEditedCSV() {{
            const cols = ["No", "商品センター", "品番", "商品区分", "1課発注点", "2課発注点", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "確保数_ロジザード", "在庫確保_弥生", "在庫確保_丸藤", "確保数のズレ", "入荷後のフリー在庫", "発注検討フラグ", "今月売れ数", "在庫回転率", "リンベル確保数", "前回決定事項", "決定事項", "発注日"];
            let csv = "\\uFEFF" + cols.join(",") + "\\n";
            tableData.forEach(r => {{
                csv += cols.map(c => `"${{String(r[c] ?? "").replace(/"/g, '""')}}"`).join(",") + "\\n";
            }});
            const blob = new Blob([csv], {{ type: "text/csv;charset=utf-8;" }});
            const a = document.createElement("a");
            a.href = URL.createObjectURL(blob);
            a.download = `発注検討表_出力_${{new Date().toISOString().slice(0,10)}}.csv`;
            a.click();
        }}
        renderTable();
    </script>
    </body>
    </html>
    """
    components.html(html_tab1, height=650, scrolling=False)

# =========================================================
# タブ2：会社（得意先）別 確保データ検索
# =========================================================
with tab2:
    st.subheader("🏢 会社（得意先）別 確保データ検索 ＆ 月別スケジュール")
    if df_yayoi.empty:
        st.warning("確保リストのデータがありません。")
    else:
        comp_totals = df_yayoi.groupby("得意先名")["受注残数"].sum().sort_values(ascending=False)
        comp_opts = ["すべての会社を表示"] + [f"{c} （確保合計: {v:,} 個）" for c, v in comp_totals.items()]
        c1, c2 = st.columns(2)
        sel_comp = c1.selectbox("🏢 会社名（得意先）を選択", comp_opts)
        kw_comp = c2.text_input("🔍 品番・備考・受注番号で絞り込み", placeholder="例: AHL-2203 / コープ")

        yv = df_yayoi.copy()
        if sel_comp != "すべての会社を表示":
            cname = sel_comp.split(" （確保合計:")[0]
            yv = yv[yv["得意先名"] == cname]
        if kw_comp:
            yv = yv[
                yv["得意先名"].str.contains(kw_comp, case=False, na=False) |
                yv["商品コード"].str.contains(kw_comp, case=False, na=False) |
                yv["備考"].str.contains(kw_comp, case=False, na=False)
            ]

        yv_merged = pd.merge(yv, m_df[["品番", "商品区分", "在庫数_ロジザード", "入荷予定日_ロジザード", "入荷後のフリー在庫", "発注検討フラグ", "決定事項"]], left_on="商品コード", right_on="品番", how="left").fillna("")
        disp_cols = [c for c in ["受注番号", "得意先名", "商品コード", "商品区分", "受注残数", "受注日", "納入期日", "担当者名", "備考", "倉庫名", "入荷予定日_ロジザード", "決定事項"] if c in yv_merged.columns]
        st.dataframe(yv_merged[disp_cols], hide_index=True, use_container_width=True, height=480)

# =========================================================
# タブ3：動画の弥生販売を完全再現した「受注伝票 ＆ 検索ダブルクリック」システム
# =========================================================
with tab3:
    # 1. 会社マスタの準備（弥生会社名.xlsx ＋ 既存確保リスト）
    comp_records = []
    seen_comps = set()
    if YAYOI_COMP_EXCEL and os.path.exists(YAYOI_COMP_EXCEL):
        try:
            df_cm = pd.read_excel(YAYOI_COMP_EXCEL, sheet_name=0)
            for _, r in df_cm.iterrows():
                vals = [clean_str(x) for x in r.values if clean_str(x)]
                if len(vals) >= 2:
                    c_code, c_name = vals[0], vals[1]
                elif len(vals) == 1:
                    c_code, c_name = "", vals[0]
                else:
                    continue
                if c_name not in seen_comps and c_name != "得意先名":
                    seen_comps.add(c_name)
                    comp_records.append({"code": c_code, "name": c_name})
        except Exception:
            pass

    if not df_yayoi.empty:
        for cname in sorted(df_yayoi["得意先名"].unique()):
            if cname and cname not in seen_comps:
                seen_comps.add(cname)
                comp_records.append({"code": "", "name": cname})

    for default_c in [
        {"code": "104400", "name": "清水産業 株式会社"},
        {"code": "127400", "name": "株式会社 丸藤"},
        {"code": "127410", "name": "株式会社 丸藤 東日本"},
        {"code": "133000", "name": "株式会社オバタ"},
        {"code": "132600", "name": "アーネスト株式会社"},
        {"code": "116300", "name": "トップ産業株式会社"},
        {"code": "102400", "name": "株式会社カワタキコーポレーション"},
        {"code": "900100", "name": "ら・でぃっしゅぼーや株式会社"},
        {"code": "900200", "name": "アマゾンベンダー（ホーム）"}
    ]:
        if default_c["name"] not in seen_comps:
            seen_comps.add(default_c["name"])
            comp_records.append(default_c)

    # 2. 商品マスタの準備（弥生商品番号.xlsx ＋ 発注検討表）
    item_records = []
    seen_items = set()
    if YAYOI_ITEM_EXCEL and os.path.exists(YAYOI_ITEM_EXCEL):
        try:
            df_im = pd.read_excel(YAYOI_ITEM_EXCEL, sheet_name=0)
            for _, r in df_im.iterrows():
                vals = [clean_str(x) for x in r.values if clean_str(x)]
                if len(vals) >= 2:
                    i_code, i_name = vals[0], vals[1]
                    if i_code not in seen_items and i_code != "商品コード":
                        seen_items.add(i_code)
                        item_records.append({"code": i_code, "name": i_name, "center": "名古屋"})
        except Exception:
            pass

    for _, r in m_df.iterrows():
        ic = clean_str(r["品番"])
        if ic and ic not in seen_items:
            seen_items.add(ic)
            item_records.append({"code": ic, "name": clean_str(r["商品区分"]), "center": clean_str(r["商品センター"]) or "名古屋"})

    # 3. 既存の登録済み伝票一覧（受注伝票検索ポップアップ用）
    existing_slips = []
    if not df_yayoi.empty:
        # 受注番号がない過去データは「得意先＋受注日」で擬似グループ化
        df_y_slip = df_yayoi.copy()
        df_y_slip["slip_id"] = df_y_slip.apply(
            lambda x: x["受注番号"] if clean_str(x.get("受注番号", "")) not in ("", "YAYOI-EX") else f"EX-{x['受注日'].replace('-','')}-{abs(hash(x['得意先名']))%1000:03d}",
            axis=1
        )
        for sid, grp in df_y_slip.groupby("slip_id", sort=False):
            first_r = grp.iloc[0]
            lines = []
            for _, lr in grp.iterrows():
                lines.append({
                    "code": clean_str(lr["商品コード"]),
                    "name": clean_str(lr.get("商品名/摘要", "")),
                    "wh": clean_str(lr.get("倉庫名", "名古屋")),
                    "qty": int(lr["受注残数"]),
                    "price": int(lr.get("単価", 0)),
                    "note": clean_str(lr.get("備考", ""))
                })
            existing_slips.append({
                "order_no": str(sid),
                "order_date": clean_str(first_r["受注日"]),
                "delivery_date": clean_str(first_r["納入期日"]),
                "comp_name": clean_str(first_r["得意先名"]),
                "staff": clean_str(first_r.get("担当者名", "松平")),
                "summary": clean_str(first_r.get("備考", "")),
                "is_manual": str(sid).startswith("0") or str(sid).startswith("WEB"),
                "lines": lines
            })

    # 次の自動採番（00002717 のような弥生形式連番）
    max_seq = 2716
    if os.path.exists(MANUAL_YAYOI_CSV):
        try:
            df_m_check = pd.read_csv(MANUAL_YAYOI_CSV)
            for ono in df_m_check.get("受注番号", []):
                s_ono = clean_str(ono).replace("WEB-", "")
                if s_ono.isdigit():
                    max_seq = max(max_seq, int(s_ono))
        except Exception:
            pass
    next_yayoi_no = f"{max_seq + 1:08d}"

    # 編集対象の伝票がセッションにセットされているかチェック
    active_slip = st.session_state.get("active_yayoi_slip", {
        "order_no": next_yayoi_no,
        "order_date": datetime.date.today().strftime("%Y-%m-%d"),
        "delivery_date": (datetime.date.today() + datetime.timedelta(days=5)).strftime("%Y-%m-%d"),
        "comp_name": comp_records[0]["name"] if comp_records else "株式会社オバタ",
        "staff": "松平",
        "summary": "",
        "lines": [
            {"code": "", "name": "", "wh": "名古屋", "qty": 0, "price": 0, "note": ""} for _ in range(5)
        ]
    })

    # --- 上部ツールバー：伝票検索ポップアップ（弥生のF8検索を再現） ---
    @st.dialog("🔍 受注伝票検索（行をクリックすると伝票に呼び出します）", width="large")
    def open_slip_search_dialog():
        st.caption("💡 弥生販売の「受注伝票検索」と同じ画面です。確認・修正したい注文の右端にある「📂 伝票を開く」を押すと、メイン伝票画面に全明細が展開されます。")
        kw_s = st.text_input("🔍 得意先名・受注番号・品番で絞り込み", placeholder="例: 丸藤 / オバタ / 000027")
        
        filtered_slips = existing_slips
        if kw_s:
            filtered_slips = [
                s for s in existing_slips
                if kw_s.lower() in s["comp_name"].lower() or kw_s.lower() in s["order_no"].lower() or any(kw_s.lower() in l["code"].lower() for l in s["lines"])
            ]

        # 検索結果テーブルヘッダー
        h_c1, h_c2, h_c3, h_c4, h_c5, h_c6 = st.columns([1.3, 1.3, 3.0, 1.0, 1.0, 1.4])
        h_c1.markdown("**受注日**")
        h_c2.markdown("**受注番号**")
        h_c3.markdown("**得意先名称**")
        h_c4.markdown("**担当者**")
        h_c5.markdown("**明細数**")
        h_c6.markdown("**操作**")
        st.divider()

        for idx_s, s in enumerate(filtered_slips[:40]):
            r_c1, r_c2, r_c3, r_c4, r_c5, r_c6 = st.columns([1.3, 1.3, 3.0, 1.0, 1.0, 1.4])
            r_c1.write(s["order_date"])
            r_c2.write(f"`{s['order_no']}`")
            r_c3.write(f"**{s['comp_name']}**")
            r_c4.write(s["staff"] or "—")
            r_c5.write(f"{len(s['lines'])} 行")
            if r_c6.button("📂 伝票を開く", key=f"open_slip_{idx_s}_{s['order_no']}", type="primary"):
                padded_lines = list(s["lines"])
                while len(padded_lines) < 5:
                    padded_lines.append({"code": "", "name": "", "wh": "名古屋", "qty": 0, "price": 0, "note": ""})
                st.session_state["active_yayoi_slip"] = {
                    "order_no": s["order_no"],
                    "order_date": s["order_date"] or datetime.date.today().strftime("%Y-%m-%d"),
                    "delivery_date": s["delivery_date"] or datetime.date.today().strftime("%Y-%m-%d"),
                    "comp_name": s["comp_name"],
                    "staff": s["staff"] or "松平",
                    "summary": s["summary"],
                    "lines": padded_lines
                }
                st.rerun()

    # 弥生ツールバー風ボタン群
    tb_c1, tb_c2, tb_c3, tb_c4 = st.columns([1.5, 2.2, 2.0, 3.3])
    if tb_c1.button("📄 新規伝票作成 (F2)", use_container_width=True):
        st.session_state["active_yayoi_slip"] = {
            "order_no": next_yayoi_no,
            "order_date": datetime.date.today().strftime("%Y-%m-%d"),
            "delivery_date": (datetime.date.today() + datetime.timedelta(days=5)).strftime("%Y-%m-%d"),
            "comp_name": comp_records[0]["name"] if comp_records else "株式会社オバタ",
            "staff": "松平",
            "summary": "",
            "lines": [{"code": "", "name": "", "wh": "名古屋", "qty": 0, "price": 0, "note": ""} for _ in range(5)]
        }
        st.rerun()

    if tb_c2.button("🔍 登録済み伝票を検索して開く (F8)", type="primary", use_container_width=True):
        open_slip_search_dialog()

    # 受注伝票本体（動画の弥生販売デザインを忠実に再現）
    st.markdown("""
    <div style="background:#2e6b5e; color:#fff; padding:6px 14px; font-weight:bold; font-size:15px; border-radius:4px 4px 0 0;">
        受注伝票 （弥生販売・商品確保システム）
    </div>
    """, unsafe_allow_html=True)

    with st.form("yayoi_slip_main_form"):
        # ヘッダー部（グリーン背景）
        h_col1, h_col2, h_col3, h_col4 = st.columns([1.5, 1.5, 1.5, 2.5])
        in_order_no = h_col1.text_input("受注番号 (自動採番)", value=active_slip["order_no"])
        
        try:
            def_odate = pd.to_datetime(active_slip["order_date"]).date()
        except Exception:
            def_odate = datetime.date.today()
        try:
            def_ddate = pd.to_datetime(active_slip["delivery_date"]).date()
        except Exception:
            def_ddate = datetime.date.today()

        in_order_date = h_col2.date_input("受注日", value=def_odate)
        in_deliv_date = h_col3.date_input("納入期日", value=def_ddate)

        staff_list = ["12 松平", "01 小泉", "02 杉中", "03 田中", "04 鈴木", "05 佐藤"]
        default_staff_idx = 0
        for idx_st, st_v in enumerate(staff_list):
            if active_slip["staff"] in st_v:
                default_staff_idx = idx_st
                break
        in_staff = h_col4.selectbox("担当者", options=staff_list, index=default_staff_idx)

        comp_names_only = [c["name"] for c in comp_records]
        if active_slip["comp_name"] not in comp_names_only:
            comp_names_only.insert(0, active_slip["comp_name"])
        in_comp = st.selectbox("得意先（会社名 ※文字入力で検索できます）", options=comp_names_only, index=comp_names_only.index(active_slip["comp_name"]))
        in_summary = st.text_input("摘要 / 全体備考", value=active_slip["summary"], placeholder="例: コープさっぽろ 12/3w (12/9〜12/15)")

        st.markdown("##### 📋 内訳明細（品番をプルダウンで検索・選択して数量を入力してください）")
        item_select_opts = ["(未選択)"] + [f"{it['code']} ｜ {it['name']}" for it in item_records]
        code_to_opt = {it["code"]: f"{it['code']} ｜ {it['name']}" for it in item_records}
        code_to_name = {it["code"]: it["name"] for it in item_records}

        form_lines = []
        for line_i in range(len(active_slip["lines"])):
            cur_l = active_slip["lines"][line_i]
            l_c1, l_c2, l_c3, l_c4, l_c5 = st.columns([0.5, 3.2, 1.2, 1.2, 2.5])
            l_c1.markdown(f"<div style='padding-top:28px; font-weight:bold; text-align:center;'>{line_i+1}</div>", unsafe_allow_html=True)

            cur_opt = code_to_opt.get(cur_l["code"], "(未選択)")
            if cur_l["code"] and cur_opt == "(未選択)":
                custom_opt = f"{cur_l['code']} ｜ {cur_l['name']}"
                item_select_opts.append(custom_opt)
                cur_opt = custom_opt

            sel_it = l_c2.selectbox(
                f"商品コード / 商品名 #{line_i+1}",
                options=item_select_opts,
                index=item_select_opts.index(cur_opt) if cur_opt in item_select_opts else 0,
                key=f"slip_item_{active_slip['order_no']}_{line_i}"
            )
            wh_opts = ["名古屋", "青山", "伊賀", "共通倉庫"]
            sel_wh = l_c3.selectbox(
                f"倉庫 #{line_i+1}",
                options=wh_opts,
                index=wh_opts.index(cur_l["wh"]) if cur_l["wh"] in wh_opts else 0,
                key=f"slip_wh_{active_slip['order_no']}_{line_i}"
            )
            sel_qty = l_c4.number_input(
                f"受注残数(数量) #{line_i+1}",
                min_value=0,
                value=int(cur_l["qty"]),
                key=f"slip_qty_{active_slip['order_no']}_{line_i}"
            )
            sel_note = l_c5.text_input(
                f"備考 #{line_i+1}",
                value=cur_l["note"],
                placeholder="個別備考",
                key=f"slip_note_{active_slip['order_no']}_{line_i}"
            )
            form_lines.append({"sel_it": sel_it, "wh": sel_wh, "qty": sel_qty, "note": sel_note})

        act_c1, act_c2 = st.columns([3, 1])
        save_btn = act_c1.form_submit_button("💾 この伝票を登録・上書き保存する (F12)", type="primary", use_container_width=True)
        del_btn = act_c2.form_submit_button("🗑️ この伝票を削除する", use_container_width=True)

        if save_btn:
            valid_rows = []
            staff_clean = in_staff.split(" ")[-1] if " " in in_staff else in_staff
            for fl in form_lines:
                if fl["sel_it"] != "(未選択)" and fl["qty"] > 0:
                    c_code = fl["sel_it"].split(" ｜ ")[0].strip()
                    c_name = code_to_name.get(c_code, fl["sel_it"].split(" ｜ ")[-1].strip())
                    valid_rows.append({
                        "受注番号": clean_str(in_order_no),
                        "得意先名": clean_str(in_comp),
                        "商品コード": c_code,
                        "商品名/摘要": c_name,
                        "受注残数": int(fl["qty"]),
                        "数量": int(fl["qty"]),
                        "単価": 0,
                        "受注日": in_order_date.strftime("%Y-%m-%d"),
                        "納入期日": in_deliv_date.strftime("%Y-%m-%d"),
                        "備考": clean_str(fl["note"]) or clean_str(in_summary),
                        "倉庫名": fl["wh"],
                        "担当者名": staff_clean
                    })

            if not valid_rows:
                st.error("⚠️ 商品が選択され、かつ数量が1以上の行を少なくとも1行入力してください。")
            else:
                df_new_rows = pd.DataFrame(valid_rows)
                if os.path.exists(MANUAL_YAYOI_CSV):
                    df_exist = pd.read_csv(MANUAL_YAYOI_CSV).fillna("")
                    df_exist = df_exist[df_exist["受注番号"].astype(str) != clean_str(in_order_no)]
                    df_out = pd.concat([df_exist, df_new_rows], ignore_index=True)
                else:
                    df_out = df_new_rows

                df_out.to_csv(MANUAL_YAYOI_CSV, index=False, encoding="utf-8-sig")
                st.cache_data.clear()
                st.success(f"✅ 受注番号『{in_order_no}』（{in_comp} / {len(valid_rows)}明細）を保存し、発注検討表に即時反映しました！")
                # 保存後は新規伝票番号にリセット
                if "active_yayoi_slip" in st.session_state:
                    del st.session_state["active_yayoi_slip"]
                st.rerun()

        if del_btn:
            if os.path.exists(MANUAL_YAYOI_CSV):
                df_exist = pd.read_csv(MANUAL_YAYOI_CSV).fillna("")
                df_exist = df_exist[df_exist["受注番号"].astype(str) != clean_str(in_order_no)]
                df_exist.to_csv(MANUAL_YAYOI_CSV, index=False, encoding="utf-8-sig")
                st.cache_data.clear()
                if "active_yayoi_slip" in st.session_state:
                    del st.session_state["active_yayoi_slip"]
                st.success(f"🗑️ 受注番号『{in_order_no}』を削除しました。")
                st.rerun()

# =========================================================
# タブ4：入荷予定・過去の会議決定履歴
# =========================================================
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
