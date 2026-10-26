import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import datetime
import os
import glob
import json
import zlib
import base64
import random
import time

st.set_page_config(page_title="発注検討表 ＆ 弥生受注管理システム", layout="wide")

st.markdown("""
<style>
    header[data-testid="stHeader"] {
        height: 1.8rem !important;
        background: transparent !important;
    }
    .block-container {
        padding-top: 2.0rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1.0rem !important;
        padding-right: 1.0rem !important;
        max-width: 100% !important;
    }
    .app-header {
        font-family: "Meiryo", sans-serif;
        font-size: 21px;
        font-weight: bold;
        color: #1f2937;
        margin-bottom: 6px;
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
        margin-bottom: 14px;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 1. 圧縮内蔵データ（バックアップ用）
# ---------------------------------------------------------
EMBEDDED_COMP_B64 = (
    "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"
)

EMBEDDED_ITEM_B64 = (
    "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"
    "udT+UeL//f8BWn3yfA=="
)

def get_embedded_masters():
    try:
        raw_c = json.loads(zlib.decompress(base64.b64decode(EMBEDDED_COMP_B64)).decode("utf-8"))
        comps = [{"code": c[0], "name": c[1]} for c in raw_c]
    except Exception:
        comps = []
    try:
        raw_i = json.loads(zlib.decompress(base64.b64decode(EMBEDDED_ITEM_B64)).decode("utf-8"))
        items = [{"code": i[0], "name": i[1], "wh": i[2] if len(i) > 2 and i[2] else ""} for i in raw_i]
    except Exception:
        items = []
    return comps, items

embed_comps, embed_items = get_embedded_masters()

# ---------------------------------------------------------
# 2. ファイルパスの自動探索
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
YAYOI_ITEM_EXCEL = find_target_file(["弥生商品番号", "商品番号"])
YAYOI_COMP_EXCEL = find_target_file(["弥生会社名", "会社名"])

SAVE_CSV = os.path.join(SCRIPT_DIR, "meeting_decisions_v8.csv")
MANUAL_YAYOI_CSV = os.path.join(SCRIPT_DIR, "yayoi_manual_additions.csv")

STAFF_OPTIONS_CLEAN = ["", "小泉", "鈴木", "松平", "原", "西山", "田中", "水原", "川村", "穴山"]

# ---------------------------------------------------------
# 3. 補助関数（正規化・数値変換）
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
    s = clean_str(val).upper()
    return s.replace(" ", "").replace(" ", "").replace("－", "-").replace("ー", "-").replace("―", "-")

def to_num(val):
    if is_na_val(val):
        return 0.0
    try:
        s = str(val).replace(",", "").replace("，", "").replace("%", "").replace("円", "").strip()
        return float(s)
    except Exception:
        return 0.0

@st.cache_data
def load_company_master(file_source):
    loaded = []
    seen = set()
    if file_source and os.path.exists(file_source):
        try:
            df_c = pd.read_excel(file_source, sheet_name=0, header=None)
            for idx in range(len(df_c)):
                r = df_c.iloc[idx].values
                if len(r) < 3:
                    continue
                c_code = clean_str(r[1])
                c_name = clean_str(r[2])
                if not c_name or c_name in ("名称", "得意先名", "会社名", "得意先リスト", "コード順"):
                    continue
                if c_name not in seen:
                    seen.add(c_name)
                    loaded.append({"code": c_code, "name": c_name})
        except Exception:
            pass
    if not loaded:
        loaded = list(embed_comps)
    return loaded

@st.cache_data
def load_item_master(file_source):
    loaded = []
    seen = set()
    if file_source and os.path.exists(file_source):
        try:
            df_i = pd.read_excel(file_source, sheet_name=0, header=None)
            for idx in range(len(df_i)):
                r = df_i.iloc[idx].values
                if len(r) < 3:
                    continue
                i_code = clean_str(r[1])
                i_name = clean_str(r[2])
                i_wh = clean_str(r[20]) if len(r) > 20 else ""
                if not i_code or i_code in ("コード", "商品コード", "品番", "商品リスト", "コード順"):
                    continue
                if i_code not in seen:
                    seen.add(i_code)
                    loaded.append({"code": i_code, "name": i_name, "wh": i_wh if i_wh in ("青山", "名古屋", "共通") else ""})
        except Exception:
            pass
    if not loaded:
        loaded = list(embed_items)
    return loaded

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
                edited_onos = set(df_man["受注番号"].astype(str).unique())
                df = df[~df["受注番号"].astype(str).isin(edited_onos)]
                df = pd.concat([df, df_man], ignore_index=True)
        except Exception:
            pass

    if not df.empty:
        df = df.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
    return df

@st.cache_data
def load_sales_csvs(uploaded_files, valid_norm_codes_tuple):
    valid_codes = set(valid_norm_codes_tuple)
    all_records = []

    local_csv_paths = []
    for d in SEARCH_DIRS:
        if os.path.exists(d):
            local_csv_paths.extend(glob.glob(os.path.join(d, "*yayoi*.csv")))
            local_csv_paths.extend(glob.glob(os.path.join(d, "*売上*.csv")))
    local_csv_paths = sorted(list(set(local_csv_paths)))

    sources = uploaded_files if uploaded_files else local_csv_paths
    if not sources:
        return pd.Series(dtype=int), 0, pd.DataFrame()

    loaded_file_count = 0
    for f in sources:
        df_s = None
        fname = f if isinstance(f, str) else getattr(f, "name", "uploaded.csv")
        for enc in ["cp932", "shift_jis", "utf-8-sig", "utf-8"]:
            try:
                if isinstance(f, str):
                    df_s = pd.read_csv(f, encoding=enc, header=None, dtype=str)
                else:
                    f.seek(0)
                    df_s = pd.read_csv(f, encoding=enc, header=None, dtype=str)
                if df_s is not None and not df_s.empty:
                    break
            except Exception:
                continue

        if df_s is None or df_s.empty:
            continue

        loaded_file_count += 1

        header_row_idx = None
        code_col_idx = None
        qty_col_idx = None
        for r_i in range(min(10, len(df_s))):
            row_strs = [clean_str(x) for x in df_s.iloc[r_i].values]
            c_idx = next((i for i, c in enumerate(row_strs) if c in ("商品コード", "品番", "商品番号", "コード") or "商品コード" in c), None)
            q_idx = next((
                i for i, c in enumerate(row_strs)
                if c in ("数量", "売上数量", "受注数量", "出荷数量", "受注残数", "売上数", "バラ数量", "総数量")
            ), None)
            if q_idx is None:
                q_idx = next((
                    i for i, c in enumerate(row_strs)
                    if ("数量" in c or "売上数" in c) and not any(ng in c for ng in ["単価", "金額", "入数", "ケース", "小数", "税"])
                ), None)
            if c_idx is not None and q_idx is not None:
                header_row_idx = r_i
                code_col_idx = c_idx
                qty_col_idx = q_idx
                break

        if header_row_idx is not None:
            for r_i in range(header_row_idx + 1, len(df_s)):
                r_vals = df_s.iloc[r_i].values
                if len(r_vals) <= max(code_col_idx, qty_col_idx):
                    continue
                c_norm = normalize_code(r_vals[code_col_idx])
                if c_norm in valid_codes:
                    q_v = to_num(r_vals[qty_col_idx])
                    if q_v != 0:
                        all_records.append({"file": os.path.basename(fname), "code_norm": c_norm, "qty": q_v})
            continue

        col_code_hits = {}
        for col_i in range(df_s.shape[1]):
            col_norms = df_s.iloc[:, col_i].apply(normalize_code)
            hits = int(col_norms.isin(valid_codes).sum())
            if hits > 0:
                col_code_hits[col_i] = hits

        if not col_code_hits:
            continue

        best_code_col = max(col_code_hits, key=col_code_hits.get)
        matched_rows_mask = df_s.iloc[:, best_code_col].apply(normalize_code).isin(valid_codes)
        df_matched = df_s[matched_rows_mask].copy()

        num_cols = {}
        for q_i in range(best_code_col + 1, min(best_code_col + 14, df_matched.shape[1])):
            s_num = df_matched.iloc[:, q_i].apply(to_num)
            if (s_num != 0).sum() > 0:
                num_cols[q_i] = s_num

        col_votes = {q_i: 0 for q_i in num_cols}
        c_keys = sorted(num_cols.keys())
        for idx_a in range(len(c_keys)):
            for idx_b in range(idx_a + 1, len(c_keys)):
                for idx_c in range(idx_b + 1, len(c_keys)):
                    ca, cb, cc = c_keys[idx_a], c_keys[idx_b], c_keys[idx_c]
                    va, vb, vc = num_cols[ca], num_cols[cb], num_cols[cc]
                    valid_m = (va != 0) & (vb >= 20) & (vc != 0) & ((va * vb - vc).abs() <= 2)
                    col_votes[ca] += int(valid_m.sum())

        best_qty_col = None
        if col_votes and max(col_votes.values()) > 0:
            best_qty_col = max(col_votes, key=col_votes.get)
        else:
            for cand_col in c_keys:
                s_cand = num_cols[cand_col]
                non_zero = s_cand[s_cand != 0]
                if len(non_zero) == 0:
                    continue
                if non_zero.median() > 1500:
                    continue
                best_qty_col = cand_col
                break

        if best_qty_col is not None:
            for _, r_m in df_matched.iterrows():
                c_norm = normalize_code(r_m.iloc[best_code_col])
                q_v = to_num(r_m.iloc[best_qty_col])
                if q_v != 0:
                    all_records.append({"file": os.path.basename(fname), "code_norm": c_norm, "qty": q_v})

    if not all_records:
        return pd.Series(dtype=int), loaded_file_count, pd.DataFrame()

    df_rec = pd.DataFrame(all_records)
    summary_series = df_rec.groupby("code_norm")["qty"].sum().round().astype(int)
    return summary_series, loaded_file_count, df_rec

# ---------------------------------------------------------
# ヘッダー ＆ アップローダー
# ---------------------------------------------------------
head_col1, head_col2 = st.columns([3.0, 2.2])
with head_col1:
    st.markdown('<div class="app-header">📊 発注検討表 ＆ 弥生受注伝票システム（完全連動版）</div>', unsafe_allow_html=True)
with head_col2:
    with st.expander("📂 売上CSV（今月売れ数）＆ 最新Excelの読込・更新", expanded=False):
        up_sales = st.file_uploader("🔥 今月の売上CSV（ yayoi_xxxx.csv をすべて選択してドロップ）", type=["csv"], accept_multiple_files=True, key="up_s")
        u_c1, u_c2 = st.columns(2)
        up_main = u_c1.file_uploader("📋 発注検討表.xlsx", type=["xlsx"], key="up_m")
        up_yayoi = u_c2.file_uploader("📑 弥生確保リスト.xlsx", type=["xlsx"], key="up_y")
        u_c3, u_c4 = st.columns(2)
        up_comp = u_c3.file_uploader("🏢 弥生会社名.xlsx", type=["xlsx"], key="up_comp")
        up_item = u_c4.file_uploader("📦 弥生商品番号.xlsx", type=["xlsx"], key="up_item")
        if st.button("🔄 最新状態に再読み込み", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

main_src = up_main if up_main is not None else MAIN_EXCEL
yayoi_src = up_yayoi if up_yayoi is not None else YAYOI_EXCEL
comp_src = up_comp if up_comp is not None else YAYOI_COMP_EXCEL
item_src = up_item if up_item is not None else YAYOI_ITEM_EXCEL

df_main, df_hist, prev_col_title = load_excel_master(main_src)
df_yayoi = load_yayoi_data(yayoi_src)
master_comps = load_company_master(comp_src)
master_items = load_item_master(item_src)

if df_main.empty:
    st.error("⚠️ 『発注検討表.xlsx』が見つかりません。右上の「📂 売上CSV ＆ 最新Excelの読込・更新」から『発注検討表.xlsx』をアップロードしてください。")
    st.stop()

# 商品コード → U列倉庫・商品名のマッピング辞書
item_wh_map = {}
item_name_map = {}
for it in master_items:
    c_norm = normalize_code(it["code"])
    raw_wh = clean_str(it.get("wh", ""))
    item_wh_map[c_norm] = raw_wh
    item_wh_map[it["code"]] = raw_wh
    item_name_map[c_norm] = it["name"]
    item_name_map[it["code"]] = it["name"]

seen_i_codes = {normalize_code(it["code"]) for it in master_items}
for _, r in df_main.iterrows():
    ic = clean_str(r["品番"])
    ic_n = normalize_code(ic)
    if ic and ic_n not in seen_i_codes:
        seen_i_codes.add(ic_n)
        wh_def = clean_str(r["商品センター"]) if clean_str(r["商品センター"]) in ("青山", "名古屋") else ""
        master_items.append({"code": ic, "name": clean_str(r["商品区分"]), "wh": wh_def})
        item_wh_map[ic] = wh_def
        item_wh_map[ic_n] = wh_def
        item_name_map[ic] = clean_str(r["商品区分"])

all_valid_norm_codes = tuple(set(df_main["品番_正規化"].unique()) | set(item_name_map.keys()))
sales_series, sales_file_cnt, df_sales_detail = load_sales_csvs(up_sales, all_valid_norm_codes)
if not sales_series.empty:
    df_main["今月売れ数"] = df_main["品番_正規化"].map(sales_series).fillna(0).astype(int)
else:
    df_main["今月売れ数"] = 0

if not df_yayoi.empty:
    df_yayoi["品番_正規化"] = df_yayoi["商品コード"].apply(normalize_code)
    df_yayoi["倉庫名"] = df_yayoi.apply(
        lambda r: item_wh_map.get(normalize_code(r["商品コード"]), "")
        if clean_str(r.get("倉庫名", "")) in ("", "WEB入力", "弥生風入力") else clean_str(r["倉庫名"]),
        axis=1
    )
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

yayoi_dict = {}
if not df_yayoi.empty:
    for code_n, grp in df_yayoi.groupby("品番_正規化", sort=False):
        grp_sorted = grp.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
        yayoi_dict[code_n] = grp_sorted[["受注番号", "得意先名", "受注残数", "数量", "受注日", "納入期日", "倉庫名", "備考", "担当者名"]].to_dict(orient="records")

hist_dict = {}
if not df_hist.empty:
    for code, grp in df_hist.sort_values("会議日", ascending=False).groupby("品番"):
        hist_dict[code] = grp[["会議日", "決定事項", "シート"]].to_dict(orient="records")

total_sales_qty = int(m_df["今月売れ数"].sum())
asc001_qty = int(sales_series.get("ASC-001", 0)) if not sales_series.empty else 0
st.caption(
    f"✅ **マスタ完全連動中** ｜ 発注検討表: **{len(m_df):,}** 商品 ｜ "
    f"得意先マスタ: **{len(master_comps):,}** 社 ｜ "
    f"商品マスタ: **{len(master_items):,}** 点（U列倉庫連動） ｜ "
    f"🔥 今月売れ数合計: **{total_sales_qty:,}** 個 （読込CSV: {sales_file_cnt} 件 ／ ASC-001集計値: **{asc001_qty:,}** 個）"
)

# ---------------------------------------------------------
# タブ構成
# ---------------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    f"📋 1. 発注検討表（全 {len(m_df):,} 件・特大ポップアップ）",
    "🏢 2. 会社（得意先）別・受注番号別 確保リスト（特大ポップアップ）",
    "💻 3. 弥生受注伝票（コード・商品名・倉庫名 分離版 ／ 最大100行）",
    "🕒 4. 入荷予定・会議履歴 ＆ 売上CSV検証"
])

# =========================================================
# タブ1：発注検討表（特大ポップアップウィンドウ 約2倍）
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
    search_kw = f_col3.text_input("🔍 品番・商品名・会社名・決定事項で検索", placeholder="例: ASC-001 / AHL-2203 / ライト / 丸藤")

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
        .table-wrap {{ width: 100%; height: 610px; overflow: auto; border: 1px solid #94a3b8; position: relative; }}
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

        /* 特大ポップアップウィンドウ（2倍サイズ・大文字仕様） */
        .modal-overlay {{
            display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(15, 23, 42, 0.7); z-index: 9999; justify-content: center; align-items: center;
        }}
        .modal-window {{
            background: #fff; width: 99%; max-width: 1750px; height: 96vh; border-radius: 10px;
            box-shadow: 0 25px 40px -5px rgba(0,0,0,0.45); display: flex; flex-direction: column; overflow: hidden; border: 3px solid #1e40af;
        }}
        .modal-header {{ background: #1e3a8a; color: #fff; padding: 16px 24px; display: flex; justify-content: space-between; align-items: center; }}
        .modal-title {{ font-size: 20px; font-weight: bold; }}
        .modal-close {{ background: #ef4444; color: #fff; border: none; border-radius: 6px; padding: 8px 22px; font-size: 15px; font-weight: bold; cursor: pointer; }}
        .modal-body {{ padding: 22px; overflow-y: auto; display: flex; flex-direction: column; gap: 18px; flex: 1; }}
        .summary-bar {{ display: flex; gap: 24px; flex-wrap: wrap; background: #f1f5f9; padding: 14px 20px; border-radius: 8px; border-left: 6px solid #2563eb; font-size: 15.5px; }}
        .summary-item {{ font-weight: bold; color: #1e293b; }}
        .summary-item span {{ color: #0284c7; font-size: 18px; margin-left: 5px; }}
        .quick-edit-box {{ background: #fffbeb; border: 2px solid #fde68a; padding: 14px 18px; border-radius: 8px; display: flex; flex-direction: column; gap: 12px; }}
        .stamp-btn {{ background: #e2e8f0; border: 1px solid #94a3b8; border-radius: 5px; padding: 5px 12px; font-size: 14px; cursor: pointer; font-weight: bold; }}
        .modal-grid {{ display: grid; grid-template-columns: 3.5fr 1.5fr; gap: 20px; }}
        .sub-table {{ width: 100%; border-collapse: collapse; }}
        .sub-table th {{ background: #e2e8f0; color: #1e293b; padding: 10px 12px; font-size: 14.5px; height: auto; position: static; border: 1px solid #cbd5e1; }}
        .sub-table td {{ padding: 10px 12px; font-size: 14.5px; height: auto; border: 1px solid #cbd5e1; }}
        .sec-title {{ font-size: 16.5px; font-weight: bold; color: #1e3a8a; margin-bottom: 8px; border-bottom: 2px solid #93c5fd; padding-bottom: 4px; }}
    </style>
    </head>
    <body>
    <div class="toolbar">
        <div class="toolbar-info">表示件数: <span id="row-count">0</span> 件</div>
        <div class="toolbar-hint">💡 品番・在庫確保(弥生)クリックで特大ポップアップが開きます！</div>
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
                    <div style="font-size:14.5px; font-weight:bold; color:#92400e; display:flex; justify-content:space-between;">
                        <span>✏️ この場で決定事項・入荷予定日を入力（表にもリアルタイム連動）</span>
                        <span>
                            スタンプ:
                            <button class="stamp-btn" onclick="addStamp('様子見')">＋様子見</button>
                            <button class="stamp-btn" onclick="addStamp('発注済')">＋発注済</button>
                            <button class="stamp-btn" onclick="addStamp('確認中')">＋確認中</button>
                            <button class="stamp-btn" onclick="addStamp('廃盤予定')">＋廃盤予定</button>
                        </span>
                    </div>
                    <div style="display:flex; gap:12px; align-items:center;">
                        <label style="font-size:14px; font-weight:bold;">前回:</label>
                        <input type="text" id="modal-view-prev" class="cell-input" style="width:200px; height:32px; background:#fef08a; font-size:14px;" readonly>
                        <label style="font-size:14px; font-weight:bold;">入荷予定日:</label>
                        <input type="text" id="modal-input-ndate" class="cell-input" style="width:140px; height:32px; font-size:14px;" oninput="syncFromModal()">
                        <label style="font-size:14px; font-weight:bold;">決定事項:</label>
                        <input type="text" id="modal-input-dec" class="cell-input" style="flex:1; height:32px; font-size:14px;" oninput="syncFromModal()">
                        <label style="font-size:14px; font-weight:bold;">発注日:</label>
                        <input type="text" id="modal-input-odate" class="cell-input" style="width:130px; height:32px; font-size:14px;" oninput="syncFromModal()">
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
                document.getElementById("modal-yayoi-area").innerHTML = `<div style="padding:16px; font-size:14.5px; color:#64748b; background:#f8fafc;">現在、弥生確保リストに『${{code}}』の受注残データはありません。</div>`;
            }} else {{
                let yHtml = `<table class="sub-table"><thead><tr><th>受注番号</th><th>得意先名</th><th>受注残数</th><th>受注日</th><th>納入期日</th><th>倉庫名</th><th>担当者</th><th>備考</th></tr></thead><tbody>`;
                yList.forEach(y => {{
                    yHtml += `<tr>
                        <td style="text-align:center; font-weight:bold;">${{y["受注番号"] || ""}}</td>
                        <td style="font-weight:bold;">${{y["得意先名"]}}</td>
                        <td style="text-align:right; font-weight:bold; color:#0284c7; font-size:15.5px;">${{fmtNum(y["受注残数"])}}</td>
                        <td style="text-align:center;">${{y["受注日"]}}</td>
                        <td style="text-align:center;">${{y["納入期日"]}}</td>
                        <td style="text-align:center; font-weight:bold;">${{y["倉庫名"] || ""}}</td>
                        <td style="text-align:center;">${{y["担当者名"] || ""}}</td>
                        <td>${{y["備考"]}}</td>
                    </tr>`;
                }});
                yHtml += `</tbody></table>`;
                document.getElementById("modal-yayoi-area").innerHTML = yHtml;
            }}

            const hList = histMap[code] || [];
            if (hList.length === 0) {{
                document.getElementById("modal-hist-area").innerHTML = `<div style="padding:16px; font-size:14.5px; color:#64748b; background:#f8fafc;">過去の会議決定履歴はありません。</div>`;
            }} else {{
                let hHtml = `<table class="sub-table"><thead><tr><th style="width:110px;">会議日</th><th>決定事項</th></tr></thead><tbody>`;
                hList.forEach(h => {{
                    hHtml += `<tr><td style="text-align:center; font-weight:bold;">${{h["会議日"]}}</td><td style="font-weight:bold;">${{h["決定事項"]}}</td></tr>`;
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
    components.html(html_tab1, height=670, scrolling=False)

# =========================================================
# タブ2：会社（得意先）別 確保リスト（特大ポップアップウィンドウ 約2倍）
# =========================================================
with tab2:
    if df_yayoi.empty:
        st.warning("確保リストのデータがありません。")
    else:
        comp_totals = df_yayoi.groupby("得意先名")["受注残数"].sum().sort_values(ascending=False)
        comp_opts = ["すべての会社を表示"] + [f"{c} （確保合計: {v:,} 個）" for c, v in comp_totals.items()]
        c1, c2 = st.columns(2)
        sel_comp = c1.selectbox("🏢 会社名（得意先）をプルダウンで選択", comp_opts)
        kw_comp = c2.text_input("🔍 受注番号・得意先名・品番・備考で検索", placeholder="例: 00000001 / オバタ / 丸藤 / AHL-2203")

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

        yv_full = pd.merge(
            yv,
            m_df[["品番_正規化", "商品区分", "在庫数_ロジザード", "入荷予定日_ロジザード", "入荷後のフリー在庫", "発注検討フラグ", "決定事項"]],
            on="品番_正規化", how="left"
        ).fillna("")

        slips_summary_list = []
        slip_details_map = {}

        for ono, grp in yv_full.groupby("受注番号", sort=False):
            ono_str = str(ono)
            first_r = grp.iloc[0]
            codes_preview = " / ".join(list(grp["商品コード"].unique())[:4])
            if len(grp["商品コード"].unique()) > 4:
                codes_preview += " 他"

            slips_summary_list.append({
                "受注番号": ono_str,
                "受注日": clean_str(first_r["受注日"]),
                "納入期日": clean_str(first_r["納入期日"]),
                "得意先名": clean_str(first_r["得意先名"]),
                "担当者名": clean_str(first_r.get("担当者名", "")),
                "登録商品数": len(grp),
                "受注残数合計": int(grp["受注残数"].sum()),
                "商品内訳プレビュー": codes_preview,
                "備考": clean_str(first_r.get("備考", ""))
            })

            d_list = []
            for _, dr in grp.iterrows():
                p_name = clean_str(dr.get("商品区分", "")) or clean_str(dr.get("商品名/摘要", "")) or item_name_map.get(clean_str(dr["商品コード"]), "")
                d_list.append({
                    "商品コード": clean_str(dr["商品コード"]),
                    "商品名": p_name,
                    "受注残数": int(dr["受注残数"]),
                    "数量": int(dr["数量"]),
                    "倉庫名": clean_str(dr.get("倉庫名", "")),
                    "納入期日": clean_str(dr.get("納入期日", "")),
                    "備考": clean_str(dr.get("備考", "")),
                    "ロジ在庫": int(to_num(dr.get("在庫数_ロジザード", 0))),
                    "フリー在庫": int(to_num(dr.get("入荷後のフリー在庫", 0))),
                    "入荷予定日": clean_str(dr.get("入荷予定日_ロジザード", "")),
                    "決定事項": clean_str(dr.get("決定事項", ""))
                })
            slip_details_map[ono_str] = d_list

        slips_summary_list.sort(key=lambda x: (x["受注日"], x["受注番号"]), reverse=True)
        slips_json = json.dumps(slips_summary_list, ensure_ascii=False)
        slip_details_json = json.dumps(slip_details_map, ensure_ascii=False)

        html_tab2 = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="utf-8">
        <style>
            * {{ box-sizing: border-box; font-family: "Meiryo", "Hiragino Kaku Gothic ProN", sans-serif; }}
            body {{ margin: 0; padding: 0; background: #fff; }}
            .toolbar {{
                display: flex; align-items: center; justify-content: space-between;
                background: #f8fafc; padding: 8px 14px; border: 1px solid #cbd5e1;
                border-bottom: none; border-radius: 6px 6px 0 0;
            }}
            .toolbar-info {{ font-size: 13.5px; color: #1e293b; font-weight: bold; }}
            .toolbar-hint {{ font-size: 12.5px; color: #0369a1; background: #e0f2fe; padding: 4px 12px; border-radius: 4px; font-weight: bold; }}
            .table-wrap {{ width: 100%; height: 580px; overflow: auto; border: 1px solid #94a3b8; }}
            table {{ border-collapse: separate; border-spacing: 0; width: 100%; }}
            th, td {{ border-right: 1px solid #cbd5e1; border-bottom: 1px solid #cbd5e1; padding: 7px 10px; font-size: 13.5px; }}
            thead th {{
                position: sticky; top: 0; background: #1e3a8a; color: #fff;
                font-weight: bold; text-align: center; height: 40px; z-index: 10;
            }}
            tbody tr:nth-child(even) {{ background: #f8fafc; }}
            tbody tr:hover {{ background: #e0f2fe; cursor: pointer; }}
            .click-link {{ color: #0284c7; font-weight: bold; text-decoration: underline; cursor: pointer; }}
            .click-link:hover {{ color: #1d4ed8; }}

            .modal-overlay {{
                display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
                background: rgba(15, 23, 42, 0.7); z-index: 9999; justify-content: center; align-items: center;
            }}
            .modal-window {{
                background: #fff; width: 99%; max-width: 1750px; height: 96vh; border-radius: 10px;
                box-shadow: 0 25px 40px -5px rgba(0,0,0,0.45); display: flex; flex-direction: column; overflow: hidden; border: 3px solid #1e40af;
            }}
            .modal-header {{ background: #1e3a8a; color: #fff; padding: 16px 24px; display: flex; justify-content: space-between; align-items: center; }}
            .modal-title {{ font-size: 20px; font-weight: bold; }}
            .modal-close {{ background: #ef4444; color: #fff; border: none; border-radius: 6px; padding: 8px 22px; font-size: 15px; font-weight: bold; cursor: pointer; }}
            .modal-body {{ padding: 22px; overflow-y: auto; display: flex; flex-direction: column; gap: 18px; flex: 1; }}
            .summary-bar {{ display: flex; gap: 24px; flex-wrap: wrap; background: #f1f5f9; padding: 14px 20px; border-radius: 8px; border-left: 6px solid #2563eb; font-size: 15.5px; }}
            .summary-item {{ font-weight: bold; color: #1e293b; }}
            .summary-item span {{ color: #0284c7; font-size: 18px; margin-left: 5px; }}
            .sub-table {{ width: 100%; border-collapse: collapse; }}
            .sub-table th {{ background: #e2e8f0; color: #1e293b; padding: 10px 12px; font-size: 14.5px; border: 1px solid #cbd5e1; position: static; height: auto; }}
            .sub-table td {{ padding: 10px 12px; font-size: 14.5px; border: 1px solid #cbd5e1; }}
            .sec-title {{ font-size: 16.5px; font-weight: bold; color: #1e3a8a; border-bottom: 2px solid #93c5fd; padding-bottom: 4px; }}
        </style>
        </head>
        <body>
        <div class="toolbar">
            <div class="toolbar-info">受注伝票件数: <span id="slip-count">0</span> 件（受注番号ごとに集約）</div>
            <div class="toolbar-hint">💡 行をクリックすると、特大ポップアップで登録内容が開きます！</div>
        </div>
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th style="width:130px;">受注番号</th>
                        <th style="width:105px;">受注日</th>
                        <th style="width:105px;">納入期日</th>
                        <th style="width:260px;">得意先名（会社名）</th>
                        <th style="width:85px;">担当者</th>
                        <th style="width:85px;">商品行数</th>
                        <th style="width:110px;">受注残数 合計</th>
                        <th>登録商品コード内訳</th>
                        <th style="width:180px;">備考 / 摘要</th>
                    </tr>
                </thead>
                <tbody id="slip-tbody"></tbody>
            </table>
        </div>

        <div class="modal-overlay" id="slip-modal" onclick="closeSlipModalOnBg(event)">
            <div class="modal-window">
                <div class="modal-header">
                    <div class="modal-title" id="slip-modal-title">📦 受注伝票 登録内容詳細</div>
                    <button class="modal-close" onclick="closeSlipModal()">✕ 閉じる</button>
                </div>
                <div class="modal-body">
                    <div class="summary-bar" id="slip-modal-summary"></div>
                    <div class="sec-title">📋 登録されている商品明細一覧</div>
                    <div id="slip-modal-table-area"></div>
                </div>
            </div>
        </div>

        <script>
            const slipsData = {slips_json};
            const slipDetailsMap = {slip_details_json};

            function fmtNum(val) {{
                const n = Number(val);
                return isNaN(n) ? val : n.toLocaleString("ja-JP");
            }}

            function renderSlips() {{
                document.getElementById("slip-count").textContent = slipsData.length.toLocaleString();
                const tbody = document.getElementById("slip-tbody");
                let html = "";
                slipsData.forEach((s, idx) => {{
                    html += `<tr onclick="openSlipModal(${{idx}})">
                        <td style="text-align:center;"><span class="click-link">${{s["受注番号"]}}</span></td>
                        <td style="text-align:center;">${{s["受注日"]}}</td>
                        <td style="text-align:center;">${{s["納入期日"]}}</td>
                        <td><span class="click-link">${{s["得意先名"]}}</span></td>
                        <td style="text-align:center;">${{s["担当者名"]}}</td>
                        <td style="text-align:right; font-weight:bold;">${{fmtNum(s["登録商品数"])}} 点</td>
                        <td style="text-align:right; font-weight:bold; color:#0284c7;">${{fmtNum(s["受注残数合計"])}} 個</td>
                        <td style="color:#334155;">${{s["商品内訳プレビュー"]}}</td>
                        <td>${{s["備考"]}}</td>
                    </tr>`;
                }});
                tbody.innerHTML = html;
            }}

            function openSlipModal(idx) {{
                const s = slipsData[idx];
                const ono = s["受注番号"];
                const details = slipDetailsMap[ono] || [];

                document.getElementById("slip-modal-title").innerHTML = `🏢 受注番号: 【${{ono}}】 ${{s["得意先名"]}} の登録内容詳細`;
                document.getElementById("slip-modal-summary").innerHTML = `
                    <div class="summary-item">得意先名: <span>${{s["得意先名"]}}</span></div>
                    <div class="summary-item">受注番号: <span>${{ono}}</span></div>
                    <div class="summary-item">受注日: <span>${{s["受注日"]}}</span></div>
                    <div class="summary-item">納入期日: <span>${{s["納入期日"]}}</span></div>
                    <div class="summary-item">担当者: <span>${{s["担当者名"] || "—"}}</span></div>
                    <div class="summary-item">明細行数: <span>${{fmtNum(s["登録商品数"])}} 行</span></div>
                    <div class="summary-item">受注残数 合計: <span style="color:#e11d48;">${{fmtNum(s["受注残数合計"])}} 個</span></div>
                `;

                let tHtml = `<table class="sub-table"><thead><tr>
                    <th style="width:55px;">No</th>
                    <th style="width:150px;">商品コード</th>
                    <th>商品名</th>
                    <th style="width:110px;">受注残数</th>
                    <th style="width:100px;">倉庫名</th>
                    <th style="width:115px;">納入期日</th>
                    <th style="width:105px;">ロジ在庫</th>
                    <th style="width:105px;">フリー在庫</th>
                    <th style="width:115px;">入荷予定日</th>
                    <th>備考 / 決定事項</th>
                </tr></thead><tbody>`;

                details.forEach((d, i) => {{
                    tHtml += `<tr>
                        <td style="text-align:center;">${{i + 1}}</td>
                        <td style="font-weight:bold; color:#0284c7;">${{d["商品コード"]}}</td>
                        <td style="font-weight:bold;">${{d["商品名"]}}</td>
                        <td style="text-align:right; font-weight:bold; color:#be123c; background:#fff1f2; font-size:15.5px;">${{fmtNum(d["受注残数"])}}</td>
                        <td style="text-align:center; font-weight:bold;">${{d["倉庫名"]}}</td>
                        <td style="text-align:center;">${{d["納入期日"]}}</td>
                        <td style="text-align:right; background:#ecfeff;">${{fmtNum(d["ロジ在庫"])}}</td>
                        <td style="text-align:right; background:#ecfeff;">${{fmtNum(d["フリー在庫"])}}</td>
                        <td style="text-align:center;">${{d["入荷予定日"]}}</td>
                        <td>${{d["備考"]}} ${{d["決定事項"] ? "（決定: " + d["決定事項"] + "）" : ""}}</td>
                    </tr>`;
                }});
                tHtml += `</tbody></table>`;
                document.getElementById("slip-modal-table-area").innerHTML = tHtml;
                document.getElementById("slip-modal").style.display = "flex";
            }}

            function closeSlipModal() {{ document.getElementById("slip-modal").style.display = "none"; }}
            function closeSlipModalOnBg(e) {{ if (e.target.id === "slip-modal") closeSlipModal(); }}

            renderSlips();
        </script>
        </body>
        </html>
        """
        components.html(html_tab2, height=650, scrolling=False)

# =========================================================
# タブ3：弥生受注伝票（コード・商品名・倉庫名 分離版 ／ 最大100行）
# =========================================================
with tab3:
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

    comp_select_list = [""] + [f"{c['name']}" + (f" （{c['code']}）" if c["code"] else "") for c in master_comps]
    item_code_list = [""] + [it["code"] for it in master_items]

    def create_blank_slip(order_no_val):
        return {
            "order_no": order_no_val,
            "order_date": datetime.date.today(),
            "delivery_date": datetime.date.today() + datetime.timedelta(days=5),
            "comp_label": "",
            "staff": "",
            "summary": "",
            "version": int(time.time() * 1000) % 1000000,
            "grid_df": pd.DataFrame([
                {"商品コード": "", "商品名": "", "倉庫名": "", "受注残数(数量)": 0, "備考": ""}
                for _ in range(10)
            ])
        }

    if "slip_editor_state" not in st.session_state:
        st.session_state["slip_editor_state"] = create_blank_slip(next_order_no_str)

    slip_state = st.session_state["slip_editor_state"]
    v_key = slip_state.get("version", 1)

    @st.dialog("🔍 受注伝票検索（クリックで詳細確認 ＆ そのまま編集画面へ呼び出し）", width="large")
    def open_yayoi_search_modal():
        st.caption("💡 登録されている伝票を受注番号ごとに一覧表示しています。行をクリックすると詳細が開き、編集画面に呼び出せます。")
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
            height=250,
            selection_mode="single-row",
            on_select="rerun"
        )

        s_rows = ev.get("selection", {}).get("rows", [])
        target_idx = s_rows[0] if s_rows else 0
        target_ono = modal_summary.iloc[target_idx]["受注番号"]
        target_lines = df_yayoi[df_yayoi["受注番号"].astype(str) == str(target_ono)]

        st.markdown(
            f"""
            <div style="background:#eff6ff; border-left:5px solid #2563eb; padding:8px 12px; border-radius:4px; margin:6px 0;">
                <b>📦 選択中の伝票明細</b> ｜ 受注番号: <b>{target_ono}</b> ／ 得意先: <b>{modal_summary.iloc[target_idx]['得意先名']}</b> ／ 担当者: <b>{modal_summary.iloc[target_idx]['担当者名']}</b>
            </div>
            """,
            unsafe_allow_html=True
        )
        st.dataframe(
            target_lines[["商品コード", "商品名/摘要", "倉庫名", "受注残数", "納入期日", "備考"]],
            hide_index=True,
            use_container_width=True,
            height=170
        )

        if st.button("📂 この伝票を編集画面に開く（明細の追加・数量修正・削除）", type="primary", use_container_width=True):
            first_r = target_lines.iloc[0]
            grid_rows = []
            for _, lr in target_lines.iterrows():
                c_code = clean_str(lr["商品コード"])
                if c_code and c_code not in item_code_list:
                    item_code_list.append(c_code)
                p_name = item_name_map.get(normalize_code(c_code), clean_str(lr.get("商品名/摘要", "")))
                u_wh = item_wh_map.get(normalize_code(c_code), clean_str(lr.get("倉庫名", "")))
                grid_rows.append({
                    "商品コード": c_code,
                    "商品名": p_name,
                    "倉庫名": u_wh,
                    "受注残数(数量)": int(lr["受注残数"]),
                    "備考": clean_str(lr.get("備考", ""))
                })
            while len(grid_rows) < 10:
                grid_rows.append({"商品コード": "", "商品名": "", "倉庫名": "", "受注残数(数量)": 0, "備考": ""})

            matched_comp = next((c for c in comp_select_list if c and c.startswith(clean_str(first_r["得意先名"]))), clean_str(first_r["得意先名"]))
            try:
                od_dt = pd.to_datetime(first_r["受注日"]).date()
            except Exception:
                od_dt = datetime.date.today()
            try:
                dd_dt = pd.to_datetime(first_r["納入期日"]).date()
            except Exception:
                dd_dt = datetime.date.today()

            st_name = clean_str(first_r.get("担当者名", ""))
            for clean_st in STAFF_OPTIONS_CLEAN:
                if clean_st and clean_st == st_name:
                    st_name = clean_st
                    break

            st.session_state["slip_editor_state"] = {
                "order_no": str(target_ono),
                "order_date": od_dt,
                "delivery_date": dd_dt,
                "comp_label": matched_comp,
                "staff": st_name,
                "summary": clean_str(first_r.get("備考", "")),
                "version": int(time.time() * 1000) % 1000000,
                "grid_df": pd.DataFrame(grid_rows[:100])
            }
            st.rerun()

    top_b1, top_b2, top_b3, top_b4 = st.columns([1.8, 2.3, 1.9, 2.5])
    if top_b1.button("📄 新規作成（すべて白紙にリセット）", use_container_width=True):
        st.session_state["slip_editor_state"] = create_blank_slip(next_order_no_str)
        st.rerun()

    if top_b2.button("🔍 登録済み伝票を検索・修正 (F8)", type="primary", use_container_width=True):
        open_yayoi_search_modal()

    if top_b3.button("➕ 入力行を10行追加（最大100行）", use_container_width=True):
        cur_df = slip_state["grid_df"]
        if len(cur_df) < 100:
            add_cnt = min(10, 100 - len(cur_df))
            extra_df = pd.DataFrame([
                {"商品コード": "", "商品名": "", "倉庫名": "", "受注残数(数量)": 0, "備考": ""}
                for _ in range(add_cnt)
            ])
            slip_state["grid_df"] = pd.concat([cur_df, extra_df], ignore_index=True)
            st.rerun()

    st.markdown('<div class="yayoi-slip-header">💻 受注伝票入力（商品コードを選ぶと「商品名(C列)」と「倉庫名(U列)」が自動で入ります）</div>', unsafe_allow_html=True)
    st.markdown('<div class="yayoi-slip-body">', unsafe_allow_html=True)

    h1, h2, h3, h4 = st.columns([1.4, 1.4, 1.4, 1.8])
    in_ono = h1.text_input("受注番号（1から自動採番）", value=slip_state["order_no"], key=f"ono_{v_key}")
    in_odate = h2.date_input("受注日", value=slip_state["order_date"], key=f"odate_{v_key}")
    in_ddate = h3.date_input("納入期日", value=slip_state["delivery_date"], key=f"ddate_{v_key}")

    staff_list = list(STAFF_OPTIONS_CLEAN)
    if slip_state["staff"] not in staff_list:
        staff_list.append(slip_state["staff"])
    in_staff = h4.selectbox(
        "担当者（番号なし・新規は白紙）",
        options=staff_list,
        index=staff_list.index(slip_state["staff"]),
        key=f"staff_{v_key}"
    )

    hc1, hc2 = st.columns([2.6, 2.4])
    if slip_state["comp_label"] not in comp_select_list:
        comp_select_list.insert(1, slip_state["comp_label"])
    in_comp_label = hc1.selectbox(
        f"得意先（弥生会社名.xlsx 全 {len(master_comps):,} 社対応 ／ 新規は白紙）",
        options=comp_select_list,
        index=comp_select_list.index(slip_state["comp_label"]),
        key=f"comp_{v_key}"
    )
    in_summary = hc2.text_input(
        "摘要 / 伝票全体メモ（新規は白紙）",
        value=slip_state["summary"],
        placeholder="例: コープさっぽろ 12/3w (12/9〜12/15)",
        key=f"sum_{v_key}"
    )

    st.markdown(f"##### 📝 内訳明細（全 {len(master_items):,} 商品対応 ／ 「商品コード」列を選択すると隣の「商品名」「倉庫名」が自動表示されます）")

    raw_edited_grid = st.data_editor(
        slip_state["grid_df"],
        num_rows="dynamic",
        use_container_width=True,
        height=350,
        key=f"yayoi_grid_{v_key}",
        column_config={
            "商品コード": st.column_config.SelectboxColumn(
                f"商品コード (B列: 全 {len(master_items):,} 点)",
                options=item_code_list,
                width="medium",
                help="クリックして商品コードを入力すると絞り込めます"
            ),
            "商品名": st.column_config.TextColumn(
                "商品名 (C列自動)",
                width="large",
                disabled=True,
                help="商品コードを選ぶと自動で商品名が入ります"
            ),
            "倉庫名": st.column_config.TextColumn(
                "倉庫名 (U列自動)",
                width="small",
                disabled=True,
                help="弥生商品番号のU列（青山・名古屋・共通・空白）が自動で入ります"
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

    need_sync = False
    synced_rows = []
    for _, r_g in raw_edited_grid.iterrows():
        c_code = clean_str(r_g.get("商品コード", ""))
        cur_name = clean_str(r_g.get("商品名", ""))
        cur_wh = clean_str(r_g.get("倉庫名", ""))
        qty_v = int(to_num(r_g.get("受注残数(数量)", 0)))
        note_v = clean_str(r_g.get("備考", ""))

        if c_code:
            expected_name = item_name_map.get(normalize_code(c_code), item_name_map.get(c_code, ""))
            expected_wh = item_wh_map.get(normalize_code(c_code), item_wh_map.get(c_code, ""))
        else:
            expected_name = ""
            expected_wh = ""

        if cur_name != expected_name or cur_wh != expected_wh:
            need_sync = True

        synced_rows.append({
            "商品コード": c_code,
            "商品名": expected_name,
            "倉庫名": expected_wh,
            "受注残数(数量)": qty_v,
            "備考": note_v
        })

    if need_sync:
        slip_state["grid_df"] = pd.DataFrame(synced_rows)
        slip_state["version"] = int(time.time() * 1000) % 1000000
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    b_col1, b_col2 = st.columns([3.2, 1.0])
    if b_col1.button("💾 この伝票を登録・上書き保存する（クリックで確定）", type="primary", use_container_width=True):
        comp_clean = in_comp_label.split(" （")[0].strip()
        staff_clean = in_staff.strip()

        if not comp_clean:
            st.error("⚠️ 「得意先（会社名）」が白紙です。プルダウンから得意先を選択してください。")
        elif not staff_clean:
            st.error("⚠️ 「担当者」が白紙です。プルダウンから担当者を選択してください。")
        else:
            valid_items = []
            for idx_r, row_g in raw_edited_grid.iterrows():
                if idx_r >= 100:
                    break
                p_code = clean_str(row_g.get("商品コード", ""))
                qty_v = int(to_num(row_g.get("受注残数(数量)", 0)))
                if p_code and qty_v > 0:
                    p_name = item_name_map.get(normalize_code(p_code), item_name_map.get(p_code, ""))
                    u_wh = item_wh_map.get(normalize_code(p_code), item_wh_map.get(p_code, ""))
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
                        "倉庫名": u_wh,
                        "担当者名": staff_clean
                    })

            if not valid_items:
                st.error("⚠️ 商品コードが選択され、かつ「受注残数(数量)」が1以上の行がありません。入力してから登録ボタンを押してください。")
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
                new_next_no = f"{max(max_seq + 1, int(clean_str(in_ono)) if clean_str(in_ono).isdigit() else max_seq + 1) + 1:08d}"
                st.session_state["slip_editor_state"] = create_blank_slip(new_next_no)
                st.success(f"✅ 受注番号『{in_ono}』（得意先: {comp_clean} ／ 担当: {staff_clean} ／ 明細 {len(valid_items)} 行）を保存しました！")
                st.rerun()

    if b_col2.button("🗑️ この伝票を削除する", use_container_width=True):
        if os.path.exists(MANUAL_YAYOI_CSV):
            df_exist = pd.read_csv(MANUAL_YAYOI_CSV).fillna("")
            df_exist = df_exist[df_exist["受注番号"].astype(str) != clean_str(in_ono)]
            df_exist.to_csv(MANUAL_YAYOI_CSV, index=False, encoding="utf-8-sig")
            st.cache_data.clear()
            st.session_state["slip_editor_state"] = create_blank_slip(next_order_no_str)
            st.success(f"🗑️ 受注番号『{in_ono}』を削除しました。")
            st.rerun()

# =========================================================
# タブ4：入荷予定・会議履歴 ＆ 売上CSV検証
# =========================================================
with tab4:
    st.subheader("🔍 今月の売上CSV 品番別・ファイル別 集計明細チェック")
    if df_sales_detail.empty:
        st.info("現在、売上CSVが読み込まれていません。画面右上の「📂 売上CSV（今月売れ数）＆ 最新Excelの読込・更新」から yayoi_xxxx.csv をすべて選択して読み込んでください。")
    else:
        chk_c1, chk_c2 = st.columns([1, 2])
        kw_sales_chk = chk_c1.text_input("品番で売上集計を検索（例: ASC-001）", value="ASC-001")
        df_s_filt = df_sales_detail[df_sales_detail["code_norm"].str.contains(kw_sales_chk.strip().upper(), na=False)] if kw_sales_chk else df_sales_detail
        chk_c2.metric(f"『{kw_sales_chk or '全商品'}』の集計売れ数 合計", f"{int(df_s_filt['qty'].sum()):,} 個", f"該当明細行数: {len(df_s_filt)} 行")
        st.dataframe(df_s_filt, hide_index=True, use_container_width=True, height=220)

    st.divider()
    ch1, ch2 = st.columns(2)
    with ch1:
        st.subheader("🚢 入荷予定あり・発注済みの商品一覧")
        inc = m_df[(m_df["入荷予定数_ロジザード"] > 0) | (m_df["発注日"] != "") | (m_df["入荷予定日_ロジザード"] != "")].copy()
        st.dataframe(inc[["品番", "商品区分", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "発注日", "前回決定事項", "決定事項"]], hide_index=True, use_container_width=True, height=450)
    with ch2:
        st.subheader("📜 過去28回の会議シート 決定事項 全履歴")
        h_kw = st.text_input("品番・決定内容で検索", placeholder="例: 発注 / 様子見")
        hv = df_hist.sort_values(["会議日", "品番"], ascending=[False, True]) if not df_hist.empty else pd.DataFrame()
        if not hv.empty and h_kw:
            hv = hv[hv["品番"].str.contains(h_kw, case=False, na=False) | hv["決定事項"].str.contains(h_kw, case=False, na=False)]
        st.dataframe(hv, hide_index=True, use_container_width=True, height=450)
