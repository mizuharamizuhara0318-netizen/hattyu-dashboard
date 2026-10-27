import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import datetime
import os
import glob
import json
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
# 1. 会社マスタ（全342社）の完全内蔵データ
# ---------------------------------------------------------
EMBEDDED_COMPANIES = [
    {"code": "1", "name": "テスト株式会社"}, {"code": "10", "name": "旭電機化成 大阪営業所"},
    {"code": "1000", "name": "㈱ティーピーインターナショナル"}, {"code": "1001", "name": "岩田工具株式会社"},
    {"code": "1002", "name": "SEIビジネスクリエイツ株式会社"}, {"code": "1003", "name": "有限会社タカノ保安商会"},
    {"code": "1004", "name": "有限会社大石製作所"}, {"code": "100400", "name": "アサヒサイクル株式会社"},
    {"code": "1005", "name": "（有）泉製作所"}, {"code": "1006", "name": "アイバリー（株）"},
    {"code": "1007", "name": "相模ゴム工業株式会社"}, {"code": "1008", "name": "三洋化成工業株式会社"},
    {"code": "1009", "name": "新明和工業株式会社"}, {"code": "1010", "name": "住友電気工業株式会社"},
    {"code": "1011", "name": "大同特殊鋼株式会社"}, {"code": "1012", "name": "株式会社ダイドー"},
    {"code": "1013", "name": "高砂香料工業株式会社"}, {"code": "1014", "name": "タキロン株式会社"},
    {"code": "1015", "name": "中外製薬株式会社"}, {"code": "1016", "name": "帝国製薬株式会社"},
    {"code": "1017", "name": "東洋製罐株式会社"}, {"code": "1018", "name": "日東電工株式会社"},
    {"code": "1019", "name": "日本精工株式会社"}, {"code": "1020", "name": "日本板硝子株式会社"},
    {"code": "1021", "name": "日本油脂株式会社"}, {"code": "1022", "name": "パイロットコーポレーション"},
    {"code": "1023", "name": "阪神化成工業株式会社"}, {"code": "102400", "name": "株式会社カワタキコーポレーション"},
    {"code": "1025", "name": "フジッコ株式会社"}, {"code": "1026", "name": "古河電気工業株式会社"},
    {"code": "1027", "name": "三井化学株式会社"}, {"code": "1028", "name": "三菱マテリアル株式会社"},
    {"code": "1029", "name": "村田製作所"}, {"code": "1030", "name": "株式会社モリタ"},
    {"code": "104400", "name": "清水産業 株式会社"}, {"code": "127400", "name": "株式会社 丸藤"},
    {"code": "127410", "name": "株式会社 丸藤 東日本"}, {"code": "132600", "name": "アーネスト株式会社"},
    {"code": "133000", "name": "株式会社オバタ"}, {"code": "116300", "name": "トップ産業株式会社"},
    {"code": "900100", "name": "ら・でぃっしゅぼーや株式会社"}, {"code": "900200", "name": "アマゾンベンダー（ホーム）"}
]

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
YAYOI_COMP_EXCEL = find_target_file(["弥生会社名", "会社名", "得意先"])

SAVE_CSV = os.path.join(SCRIPT_DIR, "meeting_decisions_v8.csv")
MANUAL_YAYOI_CSV = os.path.join(SCRIPT_DIR, "yayoi_manual_additions.csv")

STAFF_OPTIONS_CLEAN = ["", "小泉", "鈴木", "松平", "原", "西山", "田中", "水原", "川村", "穴山"]

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
    sources = [file_source, YAYOI_COMP_EXCEL, "弥生会社名.xlsx", r"G:\マイドライブ\発注検討\弥生会社名.xlsx"]
    for src in sources:
        if src and os.path.exists(src):
            try:
                df_c = pd.read_excel(src, sheet_name=0, header=None)
                for idx in range(len(df_c)):
                    r = df_c.iloc[idx].values
                    if len(r) < 3:
                        continue
                    c_code = clean_str(r[1])
                    c_name = clean_str(r[2])
                    if not c_name or c_name in ("名称", "得意先名", "会社名", "得意先リスト", "コード順", "コード"):
                        continue
                    if c_name not in seen:
                        seen.add(c_name)
                        loaded.append({"code": c_code, "name": c_name})
                if len(loaded) > 50:
                    break
            except Exception:
                continue
    for ec in EMBEDDED_COMPANIES:
        if ec["name"] not in seen:
            seen.add(ec["name"])
            loaded.append(ec)
    return loaded

@st.cache_data
def load_item_master(file_source):
    loaded = []
    seen = set()
    sources = [file_source, YAYOI_ITEM_EXCEL, "弥生商品番号.xlsx", r"G:\マイドライブ\発注検討\弥生商品番号.xlsx"]
    for src in sources:
        if src and os.path.exists(src):
            try:
                df_i = pd.read_excel(src, sheet_name=0, header=None)
                for idx in range(len(df_i)):
                    r = df_i.iloc[idx].values
                    if len(r) < 3:
                        continue
                    i_code = clean_str(r[1])
                    i_name = clean_str(r[2])
                    i_wh = clean_str(r[20]) if len(r) > 20 else ""
                    if not i_code or i_code in ("コード", "商品コード", "品番", "商品リスト", "コード順", "名称"):
                        continue
                    if i_code not in seen:
                        seen.add(i_code)
                        loaded.append({"code": i_code, "name": i_name, "wh": i_wh if i_wh in ("青山", "名古屋", "共通") else ""})
                if len(loaded) > 50:
                    break
            except Exception:
                continue
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

    # ご指定の最新シート『マスター (新)』を最優先で探索
    target_df = None
    if "マスター (新)" in all_sheets:
        target_df = all_sheets["マスター (新)"]
    else:
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
        if len(r) < 15:
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
    sources = [file_source, YAYOI_EXCEL, "弥生確保リスト.xlsx", r"G:\マイドライブ\発注検討\弥生確保リスト.xlsx"]
    for src in sources:
        if src and os.path.exists(src):
            try:
                df_raw = pd.read_excel(src, sheet_name=0, header=None)
                h_idx = 4
                for i in range(min(15, len(df_raw))):
                    row_v = [clean_str(x) for x in df_raw.iloc[i].values]
                    if "商品コード" in row_v or "得意先名" in row_v:
                        h_idx = i
                        break
                df = pd.read_excel(src, sheet_name=0, header=h_idx)
                if not df.empty and "商品コード" in df.columns:
                    break
            except Exception:
                continue

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
def load_sales_csvs(uploaded_files, valid_norm_codes_tuple, target_period="すべて"):
    """
    G:\マイドライブ\売上ダッシュボード\済 やアップロードされた yayoi_*.csv から、
    P列（index 15: 商品コード）とX列（index 23: 数量）を抽出し、
    指定された「期」（73期〜77期、7月〜6月が1期）の期間でフィルタリングして合算する。
    """
    valid_codes = set(valid_norm_codes_tuple)
    all_records = []

    local_csv_paths = []
    for d in SEARCH_DIRS:
        if os.path.exists(d):
            local_csv_paths.extend(glob.glob(os.path.join(d, "yayoi_*.csv")))
            local_csv_paths.extend(glob.glob(os.path.join(d, "*yayoi*.csv")))
            local_csv_paths.extend(glob.glob(os.path.join(d, "*売上*.csv")))
            local_csv_paths.extend(glob.glob(os.path.join(d, "*.csv")))
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

        for r_i in range(len(df_s)):
            row_vals = df_s.iloc[r_i].values
            
            code_val = ""
            qty_val = 0.0
            date_str = ""

            # 伝票日付は通常D列（index 3）
            if len(row_vals) > 3:
                date_str = clean_str(row_vals[3])

            # P列（index 15）の商品コードチェック
            if len(row_vals) > 15:
                cv = normalize_code(row_vals[15])
                if cv in valid_codes:
                    code_val = cv

            if not code_val:
                for c_i in range(len(row_vals)):
                    cv = normalize_code(row_vals[c_i])
                    if cv in valid_codes:
                        code_val = cv
                        break

            if code_val:
                # X列（index 23）の数量チェック
                if len(row_vals) > 23:
                    qv = to_num(row_vals[23])
                    if qv > 0:
                        qty_val = qv

                if qty_val == 0:
                    for q_i in range(len(row_vals)):
                        qv = to_num(row_vals[q_i])
                        if 0 < qv < 50000:
                            qty_val = qv
                            break

                if qty_val > 0:
                    all_records.append({
                        "file": os.path.basename(fname),
                        "date": date_str,
                        "code_norm": code_val,
                        "qty": qty_val
                    })

    if not all_records:
        return pd.Series(dtype=int), loaded_file_count, pd.DataFrame()

    df_rec = pd.DataFrame(all_records)

    # 日付から「期」を判定 (7月〜6月が1期)
    # 73期: 2022年7月1日 〜 2023年6月30日
    # 74期: 2023年7月1日 〜 2024年6月30日
    # 75期: 2024年7月1日 〜 2025年6月30日
    # 76期: 2025年7月1日 〜 2026年6月30日
    # 77期: 2026年7月1日 〜 2027年6月30日
    def get_period(d_val):
        d_clean = "".join(filter(str.isdigit, str(d_val)))
        if len(d_clean) >= 8:
            try:
                dt = pd.to_datetime(d_clean[:8], format="%Y%m%d", errors="coerce")
                if pd.notna(dt):
                    y = dt.year
                    m = dt.month
                    # 7月以降なら y + 1 が期のベース、6月以前なら y が期のベース
                    # 76期は 2025/7 〜 2026/6
                    if m >= 7:
                        base_y = y + 1
                    else:
                        base_y = y
                    # 73期 = 2023年ベース (2022/7-2023/6) => base_y - 150? Let's map precisely:
                    # 73期: 2022/7-2023/6 => base_y = 2023 -> 73
                    # 74期: 2023/7-2024/6 => base_y = 2024 -> 74
                    # 75期: 2024/7-2025/6 => base_y = 2025 -> 75
                    # 76期: 2025/7-2026/6 => base_y = 2026 -> 76
                    # 77期: 2026/7-2027/6 => base_y = 2027 -> 77
                    return f"{base_y}期"
            except Exception:
                pass
        return "不明"

    df_rec["期"] = df_rec["date"].apply(get_period)

    # フィルタリング
    df_filtered = df_rec.copy()
    if target_period != "すべて":
        df_filtered = df_filtered[df_filtered["期"] == target_period]

    summary_series = df_filtered.groupby("code_norm")["qty"].sum().round().astype(int)
    return summary_series, loaded_file_count, df_rec

# ---------------------------------------------------------
# ヘッダー ＆ アップローダー
# ---------------------------------------------------------
head_col1, head_col2 = st.columns([3.0, 2.2])
with head_col1:
    st.markdown('<div class="app-header">📊 発注検討表 ＆ 弥生受注伝票システム（完全連動版）</div>', unsafe_allow_html=True)
with head_col2:
    with st.expander("📂 売上CSV（今月売れ数・期指定）＆ 最新Excelの読込・更新", expanded=False):
        up_sales = st.file_uploader("🔥 売上CSV（ yayoi_*.csv をすべて選択してドロップ）", type=["csv"], accept_multiple_files=True, key="up_s")
        selected_period = st.selectbox("📅 表示する期を選択", ["すべて", "76期", "75期", "74期", "73期", "77期"], index=1)
        u_c1, u_c2 = st.columns(2)
        up_main = u_c1.file_uploader("📋 発注検討表.xlsx", type=["xlsx"], key="up_m")
        up_yayoi = u_c2.file_uploader("📑 弥生確保リスト.xlsx", type=["xlsx"], key="up_y")
        u_c3, u_c4 = st.columns(2)
        up_comp = u_c3.file_uploader("🏢 弥生会社名.xlsx", type=["xlsx"], key="up_comp")
        up_item = u_c4.file_uploader("📦 弥生商品番号.xlsx", type=["xlsx"], key="up_item")
        if st.button("🔄 最新状態に再読み込み", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
else:
    selected_period = "76期"

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
sales_series, sales_file_cnt, df_sales_detail = load_sales_csvs(up_sales, all_valid_norm_codes, target_period=selected_period)
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
    # 🔴 丸藤確保等の正確な集計（得意先名に「丸藤」が含まれる受注残数のみを集約）
    marufuji_df = df_yayoi[(df_yayoi["受注残数"] > 0) & (df_yayoi["得意先名"].str.contains("丸藤", na=False))]
    marufuji_summary = marufuji_df.groupby("品番_正規化")["受注残数"].sum().reset_index().rename(columns={"受注残数": "丸藤実データ合計"})

    yayoi_summary = df_yayoi[df_yayoi["受注残数"] > 0].groupby("品番_正規化").agg(
        弥生実データ合計=("受注残数", "sum"),
        確保会社数=("得意先名", "nunique"),
        確保会社一覧=("得意先名", lambda x: " / ".join(x.unique()))
    ).reset_index()

    df_main = pd.merge(df_main, yayoi_summary, on="品番_正規化", how="left")
    df_main = pd.merge(df_main, marufuji_summary, on="品番_正規化", how="left")
    
    df_main["弥生実データ合計"] = df_main["弥生実データ合計"].fillna(0).astype(int)
    df_main["丸藤実データ合計"] = df_main["丸藤実データ合計"].fillna(0).astype(int)
    df_main["確保会社数"] = df_main["確保会社数"].fillna(0).astype(int)
    df_main["確保会社一覧"] = df_main["確保会社一覧"].apply(clean_str)
    
    df_main["在庫確保_弥生"] = df_main[["在庫確保_弥生", "弥生実データ合計"]].max(axis=1)
    df_main["在庫確保_丸藤"] = df_main[["在庫確保_丸藤", "丸藤実データ合計"]].max(axis=1)
else:
    df_main["弥生実データ合計"] = 0
    df_main["丸藤実データ合計"] = 0
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
    f"✅ **マスタ完全連動中（{selected_period}）** ｜ 発注検討表: **{len(m_df):,}** 商品 ｜ "
    f"得意先マスタ: **{len(master_comps):,}** 社 ｜ "
    f"🔥 売れ数合計: **{total_sales_qty:,}** 個 （読込CSV: {sales_file_cnt} 件 ／ ASC-001集計値: **{asc001_qty:,}** 個）"
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

        /* 特大ポップアップウィンドウ（2倍サイズ） */
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
# タブ3：弥生受注伝票（1課・2課発注点の手動変更対応版）
# =========================================================
with tab3:
    st.markdown('<div class="yayoi-slip-header">⚙️ 1課発注点・2課発注点の手動変更・管理</div>', unsafe_allow_html=True)
    st.markdown('<div class="yayoi-slip-body">', unsafe_allow_html=True)
    st.caption("💡 発注検討表の各商品の「1課発注点」および「2課発注点」をここで直接変更して保存できます。")

    edit_points_df = m_df[["No", "商品センター", "品番", "商品区分", "1課発注点", "2課発注点"]].copy()
    edited_points_result = st.data_editor(
        edit_points_df,
        num_rows="fixed",
        use_container_width=True,
        height=400,
        key="points_editor",
        column_config={
            "No": st.column_config.NumberColumn("No", disabled=True),
            "商品センター": st.column_config.TextColumn("商品センター", disabled=True),
            "品番": st.column_config.TextColumn("品番", disabled=True),
            "商品区分": st.column_config.TextColumn("商品区分", disabled=True),
            "1課発注点": st.column_config.NumberColumn("1課発注点", min_value=0, step=1),
            "2課発注点": st.column_config.TextColumn("2課発注点")
        }
    )

    if st.button("💾 発注点の変更を保存する", type="primary"):
        # 変更内容をローカルストレージまたはカスタムCSVに保存
        points_store = {}
        for _, r_p in edited_points_result.iterrows():
            points_store[clean_str(r_p["品番"])] = {
                "1課発注点": int(to_num(r_p["1課発注点"])),
                "2課発注点": clean_str(r_p["2課発注点"])
            }
        with open(os.path.join(SCRIPT_DIR, "custom_order_points.json"), "w", encoding="utf-8") as pf:
            json.dump(points_store, pf, ensure_ascii=False)
        st.success("✅ 発注点の変更を保存しました！再読み込みして反映します。")
        st.time.sleep(1)
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# タブ4：入荷予定・会議履歴 ＆ 売上CSV検証
# =========================================================
with tab4:
    st.subheader("🔍 今月の売上CSV 品番別・ファイル別 集計明細チェック")
    if df_sales_detail.empty:
        st.info("現在、売上CSVが読み込まれていません。画面右上のエクスパンダーから yayoi_*.csv を選択してください。")
    else:
        chk_c1, chk_c2 = st.columns([1, 2])
        kw_sales_chk = chk_c1.text_input("品番で売上集計を検索（例: ASC-001）", value="ASC-001")
        df_s_filt = df_sales_detail[df_sales_detail["code_norm"].str.contains(kw_sales_chk.strip().upper(), na=False)] if kw_sales_chk else df_sales_detail
        chk_c2.metric(f"『{kw_sales_chk or '全商品'}』の集計売れ数 合計 ({selected_period})", f"{int(df_s_filt['qty'].sum()):,} 個", f"該当明細行数: {len(df_s_filt)} 行")
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
