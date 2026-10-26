import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import datetime
import os
import glob
import json

st.set_page_config(page_title="発注検討表 ＆ 弥生受注管理システム", layout="wide")

st.markdown("""
<style>
    header[data-testid="stHeader"] {
        height: 2rem !important;
        background: transparent !important;
    }
    .block-container {
        padding-top: 2.2rem !important;
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
    .yayoi-slip-header {
        background-color: #2e6b5e;
        color: white;
        padding: 8px 14px;
        font-weight: bold;
        font-size: 15px;
        border-radius: 6px 6px 0 0;
        margin-top: 6px;
    }
    .yayoi-slip-body {
        background-color: #f2f7f5;
        border: 2px solid #2e6b5e;
        border-top: none;
        border-radius: 0 0 6px 6px;
        padding: 14px;
        margin-bottom: 16px;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 1. ファイルパスの自動探索（GitHub・ローカルGドライブ両対応）
# ---------------------------------------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else "."

SEARCH_DIRS = [
    SCRIPT_DIR,
    ".",
    r"G:\マイドライブ\発注検討",
    r"G:\マイドライブ\売上ダッシュボード\済",
    r"G:\マイドライブ\売上ダッシュボード",
    "/content/drive/MyDrive/発注検討",
    "/content/drive/MyDrive/売上ダッシュボード/済"
]

def find_target_file(keywords, ext=".xlsx"):
    for d in SEARCH_DIRS:
        if os.path.exists(d):
            try:
                for fn in os.listdir(d):
                    if fn.lower().endswith(ext) and not fn.startswith("~$"):
                        if any(k in fn for k in keywords):
                            return os.path.join(d, fn)
            except Exception:
                pass
    return None

MAIN_EXCEL = find_target_file(["発注検討表"])
YAYOI_EXCEL = find_target_file(["弥生確保リスト", "確保リスト"])
YAYOI_ITEM_EXCEL = find_target_file(["弥生商品番号", "商品番号", "商品マスタ", "商品一覧"])
YAYOI_COMP_EXCEL = find_target_file(["弥生会社名", "会社名", "得意先マスタ", "得意先一覧", "tokuisaki"])

SAVE_CSV = os.path.join(SCRIPT_DIR, "meeting_decisions_v8.csv")
MANUAL_YAYOI_CSV = os.path.join(SCRIPT_DIR, "yayoi_manual_additions.csv")

# ---------------------------------------------------------
# 2. 補助関数（正規化・数値変換）
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

def normalize_code(val):
    """品番照合用の正規化（空白除去・大文字化・全角ハイフン統一）"""
    s = clean_str(val).upper()
    s = s.replace(" ", "").replace(" ", "").replace("－", "-").replace("ー", "-")
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
# 3. 各種データ読み込み関数
# ---------------------------------------------------------
@st.cache_data
def load_excel_master(file_source):
    if file_source is None:
        return pd.DataFrame(), pd.DataFrame(), "前回決定事項"
    try:
        all_sheets = pd.read_excel(file_source, sheet_name=None, header=None)
    except Exception as e:
        st.error(f"発注検討表の読み込みエラー: {e}")
        return pd.DataFrame(), pd.DataFrame(), "前回決定事項"

    target_df = None
    for s_name, df_s in all_sheets.items():
        if "マスター" in str(s_name) and len(df_s) > 5:
            target_df = df_s
            break
    if target_df is None:
        max_len = 0
        for s_name, df_s in all_sheets.items():
            if len(df_s) > max_len and df_s.shape[1] >= 20:
                max_len = len(df_s)
                target_df = df_s

    if target_df is None or target_df.empty:
        return pd.DataFrame(), pd.DataFrame(), "前回決定事項"

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
            "品番_正規化": normalize_code(hinban),
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
        df["受注番号"] = df["受注番号"].apply(clean_str) if "受注番号" in df.columns else ""
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
        
        # 受注番号が空欄のExcelデータには、得意先＋受注日ごとに一貫した番号を付与
        empty_mask = df["受注番号"] == ""
        if empty_mask.any():
            df.loc[empty_mask, "受注番号"] = df[empty_mask].apply(
                lambda r: f"Y-{r['受注日'].replace('-','')}-{abs(hash(r['得意先名']))%10000:04d}", axis=1
            )
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
                # もし同じ受注番号をWEB上で上書き編集していた場合は、手動側を優先する
                edited_onos = set(df_man["受注番号"].astype(str).unique())
                df = df[~df["受注番号"].astype(str).isin(edited_onos)]
                df = pd.concat([df, df_man], ignore_index=True)
        except Exception:
            pass

    if not df.empty:
        df = df.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
    return df

@st.cache_data
def load_company_master(file_source):
    """弥生会社名.xlsx から全シート・全行の得意先を漏れなく抽出"""
    comps = []
    seen = set()
    if file_source is not None:
        try:
            sheets = pd.read_excel(file_source, sheet_name=None, header=None)
            for _, df_c in sheets.items():
                if df_c.empty:
                    continue
                for _, r in df_c.iterrows():
                    vals = [clean_str(x) for x in r.values if clean_str(x)]
                    if not vals:
                        continue
                    # 1列目がコード（数字）で2列目が会社名の場合、または1列目が会社名の場合の両対応
                    if len(vals) >= 2 and vals[0].replace("-", "").isdigit():
                        c_code, c_name = vals[0], vals[1]
                    else:
                        c_code, c_name = "", vals[0]
                    if c_name in ("得意先名", "会社名", "名称", "得意先名称", "コード", "得意先コード", "tokuisaki"):
                        continue
                    if c_name not in seen:
                        seen.add(c_name)
                        comps.append({"code": c_code, "name": c_name})
        except Exception:
            pass
    return comps

@st.cache_data
def load_item_master(file_source):
    """弥生商品番号.xlsx から全シート・全行の商品を漏れなく抽出"""
    items = []
    seen = set()
    if file_source is not None:
        try:
            sheets = pd.read_excel(file_source, sheet_name=None, header=None)
            for _, df_i in sheets.items():
                if df_i.empty:
                    continue
                for _, r in df_i.iterrows():
                    vals = [clean_str(x) for x in r.values if clean_str(x)]
                    if not vals:
                        continue
                    i_code = vals[0]
                    i_name = vals[1] if len(vals) >= 2 else vals[0]
                    if i_code in ("商品コード", "品番", "コード", "商品番号", "名称", "商品名"):
                        continue
                    if i_code not in seen:
                        seen.add(i_code)
                        items.append({"code": i_code, "name": i_name})
        except Exception:
            pass
    return items

@st.cache_data
def load_sales_csvs(uploaded_files, valid_norm_codes_tuple):
    """売上CSV群から今月売れ数を確実に集計（品番の正規化マッチング対応）"""
    valid_codes = set(valid_norm_codes_tuple)
    all_dfs = []

    local_csv_paths = []
    for d in SEARCH_DIRS:
        if os.path.exists(d):
            local_csv_paths.extend(glob.glob(os.path.join(d, "*yayoi*.csv")))
            local_csv_paths.extend(glob.glob(os.path.join(d, "*売上*.csv")))
    local_csv_paths = sorted(list(set(local_csv_paths)))

    sources = uploaded_files if uploaded_files else local_csv_paths
    if not sources:
        return pd.Series(dtype=int), 0

    loaded_file_count = 0
    for f in sources:
        df_s = None
        for enc in ["cp932", "shift_jis", "utf-8-sig", "utf-8"]:
            try:
                if isinstance(f, str):
                    df_s = pd.read_csv(f, encoding=enc, header=None)
                else:
                    f.seek(0)
                    df_s = pd.read_csv(f, encoding=enc, header=None)
                if df_s is not None and not df_s.empty:
                    break
            except Exception:
                continue

        if df_s is None or df_s.empty:
            continue

        loaded_file_count += 1

        # 1. まず1行目にヘッダーがあるか確認
        first_row = [clean_str(x) for x in df_s.iloc[0].values]
        code_col_idx = next((i for i, c in enumerate(first_row) if any(k in c for k in ["商品コード", "品番", "商品番号", "コード"])), None)
        qty_col_idx = next((i for i, c in enumerate(first_row) if any(k in c for k in ["数量", "売上数", "受注数", "受注残", "数"]) and "単価" not in c and "金額" not in c), None)

        if code_col_idx is not None and qty_col_idx is not None:
            temp = pd.DataFrame()
            temp["code_norm"] = df_s.iloc[1:, code_col_idx].apply(normalize_code)
            temp["qty"] = pd.to_numeric(df_s.iloc[1:, qty_col_idx].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
            all_dfs.append(temp)
            continue

        # 2. ヘッダーが無い（または特殊な）弥生CSVの場合、マスターの品番と一致する列を自動特定する
        best_code_col = None
        best_match_cnt = 0
        for col_i in range(df_s.shape[1]):
            col_norms = df_s.iloc[:, col_i].apply(normalize_code)
            match_cnt = col_norms.isin(valid_codes).sum()
            if match_cnt > best_match_cnt:
                best_match_cnt = match_cnt
                best_code_col = col_i

        if best_code_col is not None and best_match_cnt > 0:
            # 品番列の右側にある数値列のうち、数量らしき列（整数中心・単価や金額より小さい列）を特定
            chosen_qty_col = None
            for q_i in range(best_code_col + 1, min(best_code_col + 10, df_s.shape[1])):
                q_vals = pd.to_numeric(df_s.iloc[:, q_i].astype(str).str.replace(",", ""), errors="coerce").dropna()
                if len(q_vals) > 0 and (q_vals > 0).any() and q_vals.median() <= 5000:
                    chosen_qty_col = q_i
                    break
            if chosen_qty_col is not None:
                temp = pd.DataFrame()
                temp["code_norm"] = df_s.iloc[:, best_code_col].apply(normalize_code)
                temp["qty"] = pd.to_numeric(df_s.iloc[:, chosen_qty_col].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
                all_dfs.append(temp)

    if not all_dfs:
        return pd.Series(dtype=int), loaded_file_count

    combined = pd.concat(all_dfs, ignore_index=True)
    combined = combined[combined["code_norm"] != ""]
    return combined.groupby("code_norm")["qty"].sum().astype(int), loaded_file_count

# ---------------------------------------------------------
# 4. 上部ヘッダー ＆ マスタ・売上CSV読込バー
# ---------------------------------------------------------
head_col1, head_col2 = st.columns([3.0, 2.2])
with head_col1:
    st.markdown('<div class="app-header">📊 発注検討表 ＆ 弥生受注伝票システム（完全連動版）</div>', unsafe_allow_html=True)
with head_col2:
    with st.expander("📂 売上CSV（今月売れ数）＆ マスタExcelの読込・更新", expanded=False):
        up_sales = st.file_uploader("🔥 今月の売上CSV（ yayoi_xxxx.csv をすべて選択してドロップ）", type=["csv"], accept_multiple_files=True, key="up_s")
        u_c1, u_c2 = st.columns(2)
        up_comp = u_c1.file_uploader("🏢 弥生会社名.xlsx", type=["xlsx"], key="up_comp")
        up_item = u_c2.file_uploader("📦 弥生商品番号.xlsx", type=["xlsx"], key="up_item")
        u_c3, u_c4 = st.columns(2)
        up_main = u_c3.file_uploader("📋 発注検討表.xlsx", type=["xlsx"], key="up_m")
        up_yayoi = u_c4.file_uploader("📑 弥生確保リスト.xlsx", type=["xlsx"], key="up_y")
        if st.button("🔄 キャッシュをクリアして最新データに更新", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

main_src = up_main if up_main is not None else MAIN_EXCEL
yayoi_src = up_yayoi if up_yayoi is not None else YAYOI_EXCEL
comp_src = up_comp if up_comp is not None else YAYOI_COMP_EXCEL
item_src = up_item if up_item is not None else YAYOI_ITEM_EXCEL

df_main, df_hist, prev_col_title = load_excel_master(main_src)
df_yayoi = load_yayoi_data(yayoi_src)

if df_main.empty:
    st.error("⚠️ 『発注検討表.xlsx』が見つかりません。右上の「📂 売上CSV ＆ マスタExcelの読込・更新」から『発注検討表.xlsx』をアップロードしてください。")
    st.stop()

# 売上CSVの集計・マッピング
valid_norm_tuple = tuple(df_main["品番_正規化"].unique())
sales_series, sales_file_cnt = load_sales_csvs(up_sales, valid_norm_tuple)
if not sales_series.empty:
    df_main["今月売れ数"] = df_main["品番_正規化"].map(sales_series).fillna(0).astype(int)
else:
    df_main["今月売れ数"] = 0

# 弥生確保データの集計＆連動
if not df_yayoi.empty:
    df_yayoi["品番_正規化"] = df_yayoi["商品コード"].apply(normalize_code)
    yayoi_summary = df_yayoi[df_yayoi["受注残数"] > 0].groupby("品番_正規化").agg(
        弥生実データ合計=("受注残数", "sum"),
        確保会社数=("得意先名", "nunique"),
        確保会社一覧=("得意先名", lambda x: " / ".join(x.unique()))
    ).reset_index()
    df_main = pd.merge(df_main, yayoi_summary, on="品番_正規化", how="left")
    df_main["弥生実データ合計"] = df_main["弥生実データ合計"].fillna(0).astype(int)
    df_main["確保会社数"] = df_main["確保会社数"].fillna(0).astype(int)
    df_main["確保会社一覧"] = df_main["確保会社一覧"].apply(clean_str)
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

# 会社マスタ・商品マスタの統合構築
master_comps = load_company_master(comp_src)
master_items = load_item_master(item_src)

seen_c = {c["name"] for c in master_comps}
if not df_yayoi.empty:
    for cn in sorted(df_yayoi["得意先名"].unique()):
        if cn and cn not in seen_c:
            seen_c.add(cn)
            master_comps.append({"code": "", "name": cn})

seen_i = {it["code"] for it in master_items}
for _, r in m_df.iterrows():
    ic = clean_str(r["品番"])
    if ic and ic not in seen_i:
        seen_i.add(ic)
        master_items.append({"code": ic, "name": clean_str(r["商品区分"])})

yayoi_dict = {}
if not df_yayoi.empty:
    for code_n, grp in df_yayoi.groupby("品番_正規化", sort=False):
        grp_sorted = grp.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
        yayoi_dict[code_n] = grp_sorted[["受注番号", "得意先名", "受注残数", "数量", "受注日", "納入期日", "倉庫名", "備考", "担当者名"]].to_dict(orient="records")

hist_dict = {}
if not df_hist.empty:
    for code, grp in df_hist.sort_values("会議日", ascending=False).groupby("品番"):
        hist_dict[code] = grp[["会議日", "決定事項", "シート"]].to_dict(orient="records")

# ステータス情報バー
total_sales_qty = int(m_df["今月売れ数"].sum())
st.caption(
    f"📊 **マスタ読込状況** ｜ 発注検討表: **{len(m_df):,}** 商品 ｜ 得意先選択肢: **{len(master_comps):,}** 社 "
    f"（{'弥生会社名.xlsx読込済' if comp_src else '確保リストから抽出※右上で弥生会社名.xlsxを読込可'}） ｜ "
    f"商品選択肢: **{len(master_items):,}** 点 ｜ "
    f"🔥 今月売れ数合計: **{total_sales_qty:,}** 個 （読込CSV: {sales_file_cnt} ファイル {'※右上でyayoi_xxxx.csvを読込可' if sales_file_cnt == 0 else ''}）"
)

# ---------------------------------------------------------
# 5. タブ構成
# ---------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    f"📋 1. 発注検討表（全 {len(m_df):,} 件・クリック即ポップアップ）",
    "🏢 2. 会社（得意先）別・受注番号別 確保データ検索",
    "💻 3. 弥生受注伝票（最大100行登録 ＆ 検索ダブルクリック修正）",
    "🕒 4. 入荷予定・過去の会議決定履歴"
])

# =========================================================
# タブ1：発注検討表
# =========================================================
with tab1:
    f_col1, f_col2, f_col3 = st.columns([2.6, 1.3, 3.0])
    filter_mode = f_col1.selectbox(
        "表示絞り込み",
        [
            "すべて表示（Excelと同じ並び）",
            "🚨 要発注（発注検討フラグがマイナスのみ）",
            "🔥 【ヒット商品】在庫があるのに今月売れている",
            "🔥 今月売れている商品すべて（売れ数1以上）",
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
    elif filter_mode == "🔥 今月売れている商品すべて（売れ数1以上）":
        v_df = v_df[v_df["今月売れ数"] >= 1].sort_values(by="今月売れ数", ascending=False)
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
        .table-wrap {{ width: 100%; height: 590px; overflow: auto; border: 1px solid #94a3b8; position: relative; }}
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
        <div class="toolbar-hint">💡 品番・在庫確保(弥生)クリックで内訳ポップアップ ｜ 右上で売上CSVを読込むと「今月売れ数」が自動反映されます</div>
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
                const salesCls = (salesVal >= 1) ? "bg-sales-hit" : "";

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
            const codeNorm = r["品番_正規化"] || code;
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

            const rawYList = yayoiMap[codeNorm] || [];
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
    components.html(html_tab1, height=640, scrolling=False)

# =========================================================
# タブ2：会社（得意先）別 確保データ検索（受注番号ごとに集約＆クリックで明細表示）
# =========================================================
with tab2:
    st.subheader("🏢 会社（得意先）別・受注番号別 確保データ検索")
    if df_yayoi.empty:
        st.warning("確保リストのデータがありません。")
    else:
        comp_totals = df_yayoi.groupby("得意先名")["受注残数"].sum().sort_values(ascending=False)
        comp_opts = ["すべての会社を表示"] + [f"{c} （確保合計: {v:,} 個）" for c, v in comp_totals.items()]
        c1, c2 = st.columns(2)
        sel_comp = c1.selectbox("🏢 会社名（得意先）をプルダウンで選択", comp_opts)
        kw_comp = c2.text_input("🔍 受注番号・品番・備考で絞り込み", placeholder="例: 000027 / AHL-2203 / コープ")

        yv = df_yayoi.copy()
        if sel_comp != "すべての会社を表示":
            cname = sel_comp.split(" （確保合計:")[0]
            yv = yv[yv["得意先名"] == cname]
        if kw_comp:
            yv = yv[
                yv["受注番号"].astype(str).str.contains(kw_comp, case=False, na=False) |
                yv["得意先名"].str.contains(kw_comp, case=False, na=False) |
                yv["商品コード"].str.contains(kw_comp, case=False, na=False) |
                yv["備考"].str.contains(kw_comp, case=False, na=False)
            ]

        # 受注番号ごとにまとめて一覧を作成
        slip_grouped = yv.groupby("受注番号", as_index=False).agg(
            受注日=("受注日", "first"),
            納入期日=("納入期日", "first"),
            得意先名=("得意先名", "first"),
            担当者名=("担当者名", "first"),
            登録商品数=("商品コード", "count"),
            受注残数_合計=("受注残数", "sum"),
            代表商品=("商品コード", lambda x: " / ".join(list(x)[:3]) + (" ..." if len(x) > 3 else "")),
            備考=("備考", "first")
        ).sort_values(by=["受注日", "受注番号"], ascending=[False, False]).reset_index(drop=True)

        st.markdown("##### 📋 受注伝票一覧（行をクリックすると、下にその伝票の登録明細が表示されます）")
        sel_event = st.dataframe(
            slip_grouped[["受注日", "受注番号", "得意先名", "担当者名", "登録商品数", "受注残数_合計", "納入期日", "代表商品", "備考"]],
            hide_index=True,
            use_container_width=True,
            height=300,
            selection_mode="single-row",
            on_select="rerun"
        )

        sel_rows = sel_event.get("selection", {}).get("rows", [])
        if sel_rows:
            chosen_ono = slip_grouped.iloc[sel_rows[0]]["受注番号"]
        elif not slip_grouped.empty:
            chosen_ono = slip_grouped.iloc[0]["受注番号"]
        else:
            chosen_ono = None

        if chosen_ono:
            detail_rows = df_yayoi[df_yayoi["受注番号"].astype(str) == str(chosen_ono)].copy()
            detail_merged = pd.merge(
                detail_rows,
                m_df[["品番", "商品区分", "在庫数_ロジザード", "入荷予定日_ロジザード", "入荷後のフリー在庫", "発注検討フラグ", "決定事項"]],
                left_on="商品コード", right_on="品番", how="left"
            ).fillna("")
            first_info = detail_rows.iloc[0]
            st.markdown(
                f"""
                <div style="background:#eff6ff; border:2px solid #2563eb; border-radius:6px; padding:12px; margin-top:8px;">
                    <div style="font-size:15px; font-weight:bold; color:#1e3a8a; margin-bottom:6px;">
                        📂 選択中の伝票内容 ｜ 受注番号: 【{chosen_ono}】 ／ 得意先: 【{first_info['得意先名']}】
                        （受注日: {first_info['受注日']} ／ 納入期日: {first_info['納入期日']} ／ 担当者: {first_info.get('担当者名','')}）
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )
            d_cols = [c for c in ["商品コード", "商品区分", "商品名/摘要", "受注残数", "数量", "倉庫名", "備考", "在庫数_ロジザード", "入荷後のフリー在庫", "入荷予定日_ロジザード", "決定事項"] if c in detail_merged.columns]
            st.dataframe(detail_merged[d_cols], hide_index=True, use_container_width=True)

# =========================================================
# タブ3：弥生受注伝票（勝手に登録されない・最大100行追加・検索呼出）
# =========================================================
with tab3:
    # 次の自動採番（1からの連番または既存最大値+1）
    max_seq = 0
    if os.path.exists(MANUAL_YAYOI_CSV):
        try:
            df_m_check = pd.read_csv(MANUAL_YAYOI_CSV)
            for ono in df_m_check.get("受注番号", []):
                s_ono = clean_str(ono).replace("WEB-", "")
                if s_ono.isdigit():
                    max_seq = max(max_seq, int(s_ono))
        except Exception:
            pass
    next_order_no_str = f"{max_seq + 1:08d}"

    # 選択肢リストの作成
    comp_select_list = [f"{c['name']}" + (f" （{c['code']}）" if c["code"] else "") for c in master_comps]
    if not comp_select_list:
        comp_select_list = ["株式会社オバタ", "株式会社 丸藤", "清水産業 株式会社"]

    item_select_list = [f"{it['code']} ｜ {it['name']}" for it in master_items]
    code_to_item_label = {it["code"]: f"{it['code']} ｜ {it['name']}" for it in master_items}
    code_to_item_name = {it["code"]: it["name"] for it in master_items}

    # セッション初期化（勝手に登録されないよう、入力状態と保存処理を完全分離）
    if "slip_editor_state" not in st.session_state:
        st.session_state["slip_editor_state"] = {
            "order_no": next_order_no_str,
            "order_date": datetime.date.today(),
            "delivery_date": datetime.date.today() + datetime.timedelta(days=5),
            "comp_label": comp_select_list[0],
            "staff": "12 松平",
            "summary": "",
            "grid_df": pd.DataFrame([
                {"商品（品番 ｜ 商品名）": "", "倉庫": "名古屋", "受注残数(数量)": 0, "備考": ""}
                for _ in range(10)
            ])
        }

    slip_state = st.session_state["slip_editor_state"]

    # --- 登録済み伝票の検索・呼び出しダイアログ（受注番号ごとに集約） ---
    @st.dialog("🔍 受注伝票検索（受注番号・得意先をクリックして伝票を呼び出し）", width="large")
    def open_yayoi_search_modal():
        st.caption("💡 登録済みの伝票一覧です。確認・修正したい行をクリックすると、すぐ下に明細が表示され、「📂 この伝票を編集画面に開く」で呼び出せます。")
        kw_modal = st.text_input("🔍 受注番号・得意先名・品番で検索", placeholder="例: 00000001 / オバタ / 丸藤")

        df_search_src = df_yayoi.copy()
        if kw_modal:
            df_search_src = df_search_src[
                df_search_src["受注番号"].astype(str).str.contains(kw_modal, case=False, na=False) |
                df_search_src["得意先名"].str.contains(kw_modal, case=False, na=False) |
                df_search_src["商品コード"].str.contains(kw_modal, case=False, na=False)
            ]

        if df_search_src.empty:
            st.info("該当する伝票がありません。")
            return

        modal_summary = df_search_src.groupby("受注番号", as_index=False).agg(
            受注日=("受注日", "first"),
            納入期日=("納入期日", "first"),
            得意先名=("得意先名", "first"),
            担当者名=("担当者名", "first"),
            明細数=("商品コード", "count"),
            合計数量=("受注残数", "sum"),
            備考=("備考", "first")
        ).sort_values(by=["受注日", "受注番号"], ascending=[False, False]).reset_index(drop=True)

        ev = st.dataframe(
            modal_summary[["受注日", "受注番号", "得意先名", "担当者名", "明細数", "合計数量", "納入期日", "備考"]],
            hide_index=True,
            use_container_width=True,
            height=260,
            selection_mode="single-row",
            on_select="rerun"
        )

        s_rows = ev.get("selection", {}).get("rows", [])
        target_idx = s_rows[0] if s_rows else 0
        target_ono = modal_summary.iloc[target_idx]["受注番号"]
        target_lines = df_yayoi[df_yayoi["受注番号"].astype(str) == str(target_ono)]

        st.markdown(f"**📦 選択中の伝票明細 【受注番号: `{target_ono}` ／ 得意先: `{modal_summary.iloc[target_idx]['得意先名']}`】**")
        st.dataframe(
            target_lines[["商品コード", "商品名/摘要", "受注残数", "倉庫名", "納入期日", "備考"]],
            hide_index=True,
            use_container_width=True,
            height=180
        )

        if st.button("📂 この伝票をメイン画面に呼び出して修正・追加する", type="primary", use_container_width=True):
            first_r = target_lines.iloc[0]
            grid_rows = []
            for _, lr in target_lines.iterrows():
                c_code = clean_str(lr["商品コード"])
                lbl = code_to_item_label.get(c_code, f"{c_code} ｜ {clean_str(lr.get('商品名/摘要', ''))}")
                if lbl not in item_select_list:
                    item_select_list.append(lbl)
                grid_rows.append({
                    "商品（品番 ｜ 商品名）": lbl,
                    "倉庫": clean_str(lr.get("倉庫名", "名古屋")) or "名古屋",
                    "受注残数(数量)": int(lr["受注残数"]),
                    "備考": clean_str(lr.get("備考", ""))
                })
            while len(grid_rows) < 10:
                grid_rows.append({"商品（品番 ｜ 商品名）": "", "倉庫": "名古屋", "受注残数(数量)": 0, "備考": ""})

            matched_comp = next((c for c in comp_select_list if c.startswith(clean_str(first_r["得意先名"]))), clean_str(first_r["得意先名"]))
            try:
                od_dt = pd.to_datetime(first_r["受注日"]).date()
            except Exception:
                od_dt = datetime.date.today()
            try:
                dd_dt = pd.to_datetime(first_r["納入期日"]).date()
            except Exception:
                dd_dt = datetime.date.today()

            st.session_state["slip_editor_state"] = {
                "order_no": str(target_ono),
                "order_date": od_dt,
                "delivery_date": dd_dt,
                "comp_label": matched_comp,
                "staff": clean_str(first_r.get("担当者名", "12 松平")) or "12 松平",
                "summary": clean_str(first_r.get("備考", "")),
                "grid_df": pd.DataFrame(grid_rows[:100])
            }
            st.rerun()

    # 上部操作ボタンバー
    top_b1, top_b2, top_b3, top_b4 = st.columns([1.6, 2.2, 1.8, 3.0])
    if top_b1.button("📄 新規伝票を作成 (F2)", use_container_width=True):
        st.session_state["slip_editor_state"] = {
            "order_no": next_order_no_str,
            "order_date": datetime.date.today(),
            "delivery_date": datetime.date.today() + datetime.timedelta(days=5),
            "comp_label": comp_select_list[0],
            "staff": "12 松平",
            "summary": "",
            "grid_df": pd.DataFrame([
                {"商品（品番 ｜ 商品名）": "", "倉庫": "名古屋", "受注残数(数量)": 0, "備考": ""}
                for _ in range(10)
            ])
        }
        st.rerun()

    if top_b2.button("🔍 登録済み伝票を検索・呼出 (F8)", type="primary", use_container_width=True):
        open_yayoi_search_modal()

    if top_b3.button("➕ 入力行を10行追加（最大100行）", use_container_width=True):
        cur_df = slip_state["grid_df"]
        if len(cur_df) < 100:
            add_cnt = min(10, 100 - len(cur_df))
            extra_df = pd.DataFrame([
                {"商品（品番 ｜ 商品名）": "", "倉庫": "名古屋", "受注残数(数量)": 0, "備考": ""}
                for _ in range(add_cnt)
            ])
            slip_state["grid_df"] = pd.concat([cur_df, extra_df], ignore_index=True)
            st.rerun()

    # 弥生風伝票ヘッダー入力部（フォーム外なのでEnterやクリックで勝手に登録されません）
    st.markdown('<div class="yayoi-slip-header">💻 受注伝票入力（登録ボタンを押すまで勝手には登録されません ／ 最大100行対応）</div>', unsafe_allow_html=True)
    st.markdown('<div class="yayoi-slip-body">', unsafe_allow_html=True)

    h1, h2, h3, h4 = st.columns([1.4, 1.4, 1.4, 2.0])
    in_ono = h1.text_input("受注番号（自動発行・重複不可）", value=slip_state["order_no"])
    in_odate = h2.date_input("受注日", value=slip_state["order_date"])
    in_ddate = h3.date_input("納入期日", value=slip_state["delivery_date"])

    staff_choices = ["12 松平", "01 小泉", "02 杉中", "03 田中", "04 鈴木", "05 佐藤"]
    if slip_state["staff"] not in staff_choices:
        staff_choices.insert(0, slip_state["staff"])
    in_staff = h4.selectbox("担当者（直接選択）", options=staff_choices, index=staff_choices.index(slip_state["staff"]))

    hc1, hc2 = st.columns([2.5, 2.5])
    if slip_state["comp_label"] not in comp_select_list:
        comp_select_list.insert(0, slip_state["comp_label"])
    in_comp_label = hc1.selectbox(
        f"得意先（会社名 ※全 {len(comp_select_list):,} 社から文字入力で検索可能）",
        options=comp_select_list,
        index=comp_select_list.index(slip_state["comp_label"])
    )
    in_summary = hc2.text_input("摘要 / 伝票全体メモ", value=slip_state["summary"], placeholder="例: コープさっぽろ 12/3w (12/9〜12/15)")

    st.markdown(f"##### 📝 内訳明細（現在 {len(slip_state['grid_df'])} 行表示中 ／ 表の下端または上の「＋入力行を10行追加」で最大100行まで追加可能）")

    edited_grid = st.data_editor(
        slip_state["grid_df"],
        num_rows="dynamic",
        use_container_width=True,
        height=340,
        key=f"yayoi_grid_{slip_state['order_no']}",
        column_config={
            "商品（品番 ｜ 商品名）": st.column_config.SelectboxColumn(
                f"商品コード / 商品名（全 {len(item_select_list):,} 点）",
                options=item_select_list,
                width="large",
                help="クリックして品番または商品名の一部を入力すると絞り込めます"
            ),
            "倉庫": st.column_config.SelectboxColumn(
                "倉庫名",
                options=["名古屋", "青山", "伊賀", "共通倉庫"],
                default="名古屋",
                width="small"
            ),
            "受注残数(数量)": st.column_config.NumberColumn(
                "受注残数 (数量)",
                min_value=0,
                max_value=999999,
                step=1,
                default=0,
                width="small"
            ),
            "備考": st.column_config.TextColumn(
                "備考 / 個別摘要",
                width="medium"
            )
        }
    )

    st.markdown("</div>", unsafe_allow_html=True)

    # 明示的な登録ボタン・削除ボタン（ここをクリックした時だけ登録・削除が走る）
    b_col1, b_col2 = st.columns([3.2, 1.0])
    if b_col1.button("💾 この伝票を登録・上書き保存する（クリックで確定）", type="primary", use_container_width=True):
        valid_items = []
        comp_clean = in_comp_label.split(" （")[0].strip()
        staff_clean = in_staff.split(" ")[-1].strip() if " " in in_staff else in_staff.strip()

        for idx_r, row_g in edited_grid.iterrows():
            if idx_r >= 100:
                break
            sel_prod = clean_str(row_g.get("商品（品番 ｜ 商品名）", ""))
            qty_v = int(to_num(row_g.get("受注残数(数量)", 0)))
            if sel_prod and qty_v > 0:
                p_code = sel_prod.split(" ｜ ")[0].strip()
                p_name = code_to_item_name.get(p_code, sel_prod.split(" ｜ ")[-1].strip())
                valid_items.append({
                    "受注番号": clean_str(in_ono),
                    "得意先名": comp_clean,
                    "商品コード": p_code,
                    "商品名/摘要": p_name,
                    "受注残数": qty_v,
                    "数量": qty_v,
                    "単価": 0,
                    "受注日": in_odate.strftime("%Y-%m-%d"),
                    "納入期日": in_ddate.strftime("%Y-%m-%d"),
                    "備考": clean_str(row_g.get("備考", "")) or clean_str(in_summary),
                    "倉庫名": clean_str(row_g.get("倉庫", "名古屋")) or "名古屋",
                    "担当者名": staff_clean
                })

        if not valid_items:
            st.error("⚠️ 商品が選択され、かつ「受注残数(数量)」が1以上の行がありません。商品と数量を入力してから登録ボタンを押してください。")
        else:
            df_new_slip = pd.DataFrame(valid_items)
            if os.path.exists(MANUAL_YAYOI_CSV):
                df_exist = pd.read_csv(MANUAL_YAYOI_CSV).fillna("")
                df_exist = df_exist[df_exist["受注番号"].astype(str) != clean_str(in_ono)]
                df_out = pd.concat([df_exist, df_new_slip], ignore_index=True)
            else:
                df_out = df_new_slip

            df_out.to_csv(MANUAL_YAYOI_CSV, index=False, encoding="utf-8-sig")
            st.cache_data.clear()
            del st.session_state["slip_editor_state"]
            st.success(f"✅ 受注番号『{in_ono}』（得意先: {comp_clean} ／ 明細 {len(valid_items)} 行）を保存しました！")
            st.rerun()

    if b_col2.button("🗑️ この伝票を削除する", use_container_width=True):
        if os.path.exists(MANUAL_YAYOI_CSV):
            df_exist = pd.read_csv(MANUAL_YAYOI_CSV).fillna("")
            df_exist = df_exist[df_exist["受注番号"].astype(str) != clean_str(in_ono)]
            df_exist.to_csv(MANUAL_YAYOI_CSV, index=False, encoding="utf-8-sig")
            st.cache_data.clear()
            del st.session_state["slip_editor_state"]
            st.success(f"🗑️ 受注番号『{in_ono}』を削除しました。")
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
