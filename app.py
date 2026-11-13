import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
import datetime
import os
import json
import zipfile
import io
import time
import re
import math
import urllib.request

st.set_page_config(page_title="発注検討表 ＆ 弥生受注管理システム（Googleドライブ完全連動版）", layout="wide")

st.markdown("""
<style>
    header[data-testid="stHeader"] {
        height: 1.8rem !important;
        background: transparent !important;
    }
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 1rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        max-width: 100% !important;
    }
    .app-header {
        font-family: "Meiryo", sans-serif;
        font-size: 20px;
        font-weight: bold;
        color: #1f2937;
        margin-bottom: 6px;
        padding-bottom: 4px;
        border-bottom: 3px solid #2563eb;
    }
</style>
""", unsafe_allow_html=True)

# ① 発注検討表のGoogleドライブURL
GDRIVE_ORDER_URL = "https://docs.google.com/spreadsheets/d/1JequkGmSRMwoo9rK5HAodfnY6foXl-F2/edit?usp=drive_link"
# ② 在庫確保.xlsx のGoogleドライブURL（空欄でも同じGoogleドライブ内から自動検索します）
GDRIVE_KAKUHO_URL = ""

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else "."
MANUAL_YAYOI_CSV = os.path.join(SCRIPT_DIR, "yayoi_manual_additions.csv")

STAFF_OPTIONS_CLEAN = ["", "小泉", "鈴木", "松平", "原", "西山", "田中", "水原", "川村", "穴山"]

MARUFUJI_TARGET_NAMES = {
    "株式会社 丸藤",
    "株式会社 丸藤 九州営業グループ",
    "株式会社 丸藤 東日本",
    "株式会社丸藤",
    "株式会社丸藤東日本"
}

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

def parse_jp_date(val):
    """
    『2027年01月12日』『2026/12/23』『2026-12-23』などあらゆる日付形式を
    確実に認識して (Timestamp, 'YYYY-MM-DD') に変換する関数
    """
    if is_na_val(val):
        return pd.NaT, ""
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        ts = pd.Timestamp(val)
        return ts, ts.strftime("%Y-%m-%d")
    s = clean_str(val)
    if not s:
        return pd.NaT, ""
    # 「2027年01月12日」→「2027-01-12」に変換
    s_norm = re.sub(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?', r'\1-\2-\3', s)
    s_norm = s_norm.replace("/", "-").split(" ")[0]
    dt = pd.to_datetime(s_norm, errors="coerce")
    if pd.notna(dt):
        return dt, dt.strftime("%Y-%m-%d")
    return pd.NaT, s

def clean_arrival_date(val):
    if is_na_val(val):
        return ""
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        return val.strftime("%Y-%m-%d")
    s = clean_str(val)
    if "00:00:00" in s:
        s = s.split(" ")[0].strip()
    if s.replace(",", "").replace(".", "").isdigit() or s == "0":
        return ""
    return s

def normalize_code(val):
    s = clean_str(val).upper()
    return s.replace(" ", "").replace(" ", "").replace("－", "-").replace("ー", "-").replace("―", "-")

def to_num(val):
    if is_na_val(val):
        return 0.0
    try:
        s = str(val).replace(",", "").replace("，", "").replace("%", "").replace("円", "").replace("(", "-").replace(")", "").strip()
        return float(s)
    except Exception:
        return 0.0

def fix_corrupted_xlsx_filters(xlsx_bytes: bytes) -> io.BytesIO:
    """
    Excelファイル（発注検討表や在庫確保.xlsx）にオートフィルターがかかっていても
    エラーにならず全行（非表示行含む）を読み込めるようにメモリ上で修復する関数
    """
    in_mem = io.BytesIO(xlsx_bytes)
    out_mem = io.BytesIO()
    try:
        with zipfile.ZipFile(in_mem, 'r') as zin, zipfile.ZipFile(out_mem, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
                    text = content.decode('utf-8', errors='ignore')
                    text = re.sub(r'<autoFilter[^>]*>.*?</autoFilter>', '', text, flags=re.DOTALL)
                    text = re.sub(r'<autoFilter[^>]*/>', '', text)
                    content = text.encode('utf-8')
                zout.writestr(item, content)
        out_mem.seek(0)
        return out_mem
    except Exception:
        in_mem.seek(0)
        return in_mem

def read_excel_safely(src, sheet_name=0, header=None):
    """どんなExcelファイル（フィルター付き・破損XML付き）でも確実にDataFrameとして読み込む"""
    raw_bytes = None
    if isinstance(src, bytes):
        raw_bytes = src
    elif hasattr(src, "getvalue"):
        raw_bytes = src.getvalue()
    elif isinstance(src, str) and os.path.exists(src):
        with open(src, "rb") as f:
            raw_bytes = f.read()
    if raw_bytes is None:
        return pd.DataFrame()

    clean_io = fix_corrupted_xlsx_filters(raw_bytes)
    try:
        return pd.read_excel(clean_io, sheet_name=sheet_name, header=header)
    except Exception:
        clean_io.seek(0)
        wb = openpyxl.load_workbook(clean_io, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[sheet_name]] if isinstance(sheet_name, int) else wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return pd.DataFrame()
        if header is None:
            return pd.DataFrame(rows)
        elif header < len(rows):
            cols = [str(c) if c is not None else f"Unnamed:{i}" for i, c in enumerate(rows[header])]
            return pd.DataFrame(rows[header + 1:], columns=cols)
        return pd.DataFrame(rows)

# =========================================================
# Googleドライブ接続・ファイル自動取得処理
# =========================================================
def get_gdrive_service():
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        creds_info = None
        if "gcp_service_account" in st.secrets:
            creds_info = dict(st.secrets["gcp_service_account"])
        elif "google_service_account" in st.secrets:
            creds_info = dict(st.secrets["google_service_account"])
        elif "private_key" in st.secrets and "client_email" in st.secrets:
            creds_info = dict(st.secrets)
        else:
            for k in st.secrets:
                val = st.secrets[k]
                if hasattr(val, "keys") and "private_key" in val:
                    creds_info = dict(val)
                    break
                elif isinstance(val, str) and "private_key" in val:
                    try:
                        creds_info = json.loads(val)
                        break
                    except Exception:
                        pass
        if creds_info:
            creds = service_account.Credentials.from_service_account_info(
                creds_info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
            )
            return build("drive", "v3", credentials=creds)
    except Exception:
        pass
    return None

def download_file_by_id(service, file_id: str):
    if service is not None:
        from googleapiclient.http import MediaIoBaseDownload
        try:
            req = service.files().get_media(fileId=file_id, supportsAllDrives=True)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            if len(fh.getvalue()) > 500:
                return fh.getvalue(), "Google Drive API"
        except Exception:
            try:
                req = service.files().export_media(
                    fileId=file_id,
                    mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, req)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                if len(fh.getvalue()) > 500:
                    return fh.getvalue(), "Google Drive API (スプレッドシート)"
            except Exception:
                pass

    dl_urls = [
        f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx",
        f"https://drive.google.com/uc?export=download&id={file_id}",
        f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t"
    ]
    for u in dl_urls:
        try:
            req = urllib.request.Request(u, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=20) as resp:
                data = resp.read()
                if len(data) > 1000 and data[:2] == b'PK':
                    return data, "Googleドライブ共有リンク"
        except Exception:
            continue
    return None, None

@st.cache_data(ttl=3600, show_spinner="☁️ Googleドライブから最新データを同期中...")
def sync_all_files_from_gdrive(order_url: str, kakuho_url: str = ""):
    m_ord = re.search(r'/d/([a-zA-Z0-9_-]+)', order_url or "")
    target_ord_id = m_ord.group(1) if m_ord else ""

    m_kak = re.search(r'/d/([a-zA-Z0-9_-]+)', kakuho_url or "")
    target_kak_id = m_kak.group(1) if m_kak else ""

    service = get_gdrive_service()

    result = {
        "main_bytes": None,
        "yayoi_bytes": None,
        "yayoi_filename": "",
        "comp_bytes": None,
        "item_bytes": None,
        "source_label": "",
        "error_detail": ""
    }

    if target_ord_id:
        b_main, src_lbl = download_file_by_id(service, target_ord_id)
        if b_main:
            result["main_bytes"] = b_main
            result["source_label"] = src_lbl
        else:
            has_secret = "設定あり" if service is not None else "未設定"
            result["error_detail"] = f"発注検討表(ID: {target_ord_id})にアクセスできませんでした。（Secrets認証: {has_secret}）"

    if target_kak_id:
        b_kak, _ = download_file_by_id(service, target_kak_id)
        if b_kak:
            result["yayoi_bytes"] = b_kak
            result["yayoi_filename"] = "Googleドライブ(URL指定: 在庫確保)"

    if service is not None:
        try:
            q = "trashed = false and (name contains '在庫確保' or name contains '弥生確保' or name contains '確保リスト' or name contains '発注検討表' or name contains '弥生会社名' or name contains '弥生商品番号')"
            res = service.files().list(
                q=q,
                pageSize=50,
                fields="files(id, name, modifiedTime)",
                orderBy="modifiedTime desc",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            files_list = res.get("files", [])

            # 『在庫確保』という名前のファイルを最優先で探す
            for f in files_list:
                fname = f["name"]
                fid = f["id"]
                if fname.startswith("~$"):
                    continue
                if result["yayoi_bytes"] is None and "在庫確保" in fname:
                    b, _ = download_file_by_id(service, fid)
                    if b:
                        result["yayoi_bytes"] = b
                        result["yayoi_filename"] = f"GDrive: {fname}"
                        break

            for f in files_list:
                fname = f["name"]
                fid = f["id"]
                if fname.startswith("~$"):
                    continue
                if result["main_bytes"] is None and "発注検討表" in fname:
                    b, lbl = download_file_by_id(service, fid)
                    if b:
                        result["main_bytes"] = b
                        result["source_label"] = f"{lbl}: {fname}"
                elif result["yayoi_bytes"] is None and any(k in fname for k in ["弥生確保", "確保リスト"]):
                    b, _ = download_file_by_id(service, fid)
                    if b:
                        result["yayoi_bytes"] = b
                        result["yayoi_filename"] = f"GDrive: {fname}"
                elif result["comp_bytes"] is None and ("弥生会社名" in fname or "得意先" in fname):
                    b, _ = download_file_by_id(service, fid)
                    if b:
                        result["comp_bytes"] = b
                elif result["item_bytes"] is None and ("弥生商品番号" in fname or "商品番号" in fname):
                    b, _ = download_file_by_id(service, fid)
                    if b:
                        result["item_bytes"] = b
        except Exception:
            pass

    return result

def find_local_file(keywords, ext=".xlsx"):
    for fn in os.listdir(SCRIPT_DIR):
        if fn.lower().endswith(ext) and not fn.startswith("~$"):
            if any(k in fn for k in keywords):
                return os.path.join(SCRIPT_DIR, fn)
    return None

@st.cache_data
def load_company_master(file_bytes_or_path):
    loaded = []
    seen = set()
    sources = [file_bytes_or_path, find_local_file(["弥生会社名", "会社名", "得意先"])]
    for src in sources:
        if src is not None:
            try:
                df_c = read_excel_safely(src, sheet_name=0, header=None)
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
    return loaded

@st.cache_data
def load_item_master(file_bytes_or_path):
    loaded = []
    seen = set()
    sources = [file_bytes_or_path, find_local_file(["弥生商品番号", "商品番号"])]
    for src in sources:
        if src is not None:
            try:
                df_i = read_excel_safely(src, sheet_name=0, header=None)
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
    raw_bytes = None
    if file_source is not None:
        if isinstance(file_source, bytes):
            raw_bytes = file_source
        elif hasattr(file_source, "getvalue"):
            raw_bytes = file_source.getvalue()
        elif isinstance(file_source, str) and os.path.exists(file_source):
            with open(file_source, "rb") as f:
                raw_bytes = f.read()

    if raw_bytes is None:
        return pd.DataFrame(), pd.DataFrame(), "前回決定事項", {}, "", {}

    clean_io = fix_corrupted_xlsx_filters(raw_bytes)
    all_sheets = {}
    try:
        all_sheets = pd.read_excel(clean_io, sheet_name=None, header=None)
    except Exception:
        try:
            clean_io.seek(0)
            wb = openpyxl.load_workbook(clean_io, read_only=True, data_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                all_sheets[sname] = pd.DataFrame(list(ws.iter_rows(values_only=True)))
        except Exception as ex:
            st.error(f"発注検討表の読み込みエラー: {ex}")
            return pd.DataFrame(), pd.DataFrame(), "前回決定事項", {}, "", {}

    target_sheet_name = list(all_sheets.keys())[0]
    if "マスター (新) (3)" in all_sheets:
        target_sheet_name = "マスター (新) (3)"
    else:
        for cand in reversed(list(all_sheets.keys())):
            if "マスター" in str(cand) or (str(cand).strip().isdigit() and len(str(cand).strip()) == 5):
                target_sheet_name = cand
                break
    target_df = all_sheets[target_sheet_name]

    z_free_sum = {}
    z_sec_sum = {}
    if "在庫貼り付け" in all_sheets:
        df_z = all_sheets["在庫貼り付け"].iloc[4:]
        for _, rz in df_z.iterrows():
            vals = rz.values
            if len(vals) > 2 and clean_str(vals[1]):
                c_n = normalize_code(vals[1])
                z_free_sum[c_n] = z_free_sum.get(c_n, 0) + int(to_num(vals[2]))
            if len(vals) > 6 and clean_str(vals[5]):
                c_n2 = normalize_code(vals[5])
                z_sec_sum[c_n2] = z_sec_sum.get(c_n2, 0) + int(to_num(vals[6]))

    nyuka_sum = {}
    if "入荷予定貼り付け" in all_sheets:
        df_n = all_sheets["入荷予定貼り付け"].iloc[4:]
        for _, rn in df_n.iterrows():
            vals = rn.values
            if len(vals) > 2 and clean_str(vals[1]):
                c_n = normalize_code(vals[1])
                nyuka_sum[c_n] = nyuka_sum.get(c_n, 0) + int(to_num(vals[2]))

    lot_map = {}
    prod_name_sheet_map = {}
    if "商品一覧表 (3)" in all_sheets:
        df_p = all_sheets["商品一覧表 (3)"].iloc[2:]
        for _, rp in df_p.iterrows():
            vals = rp.values
            if len(vals) > 6 and clean_str(vals[3]):
                c_n = normalize_code(vals[3])
                lot_v = int(to_num(vals[6]))
                if lot_v > 0:
                    lot_map[c_n] = lot_v
                if clean_str(vals[4]):
                    prod_name_sheet_map[c_n] = clean_str(vals[4])

    sheet_yayoi_sum = {}
    sheet_maru_sum = {}
    if "確保貼り付け" in all_sheets:
        df_k = all_sheets["確保貼り付け"].iloc[4:]
        for _, rk in df_k.iterrows():
            vals = rk.values
            if len(vals) > 14 and clean_str(vals[13]) and clean_str(vals[13]) not in ("行ラベル", "総計", "(空白)"):
                c_n = normalize_code(vals[13])
                sheet_yayoi_sum[c_n] = sheet_yayoi_sum.get(c_n, 0) + int(to_num(vals[14]))
            if len(vals) > 11 and clean_str(vals[9]) and clean_str(vals[9]) != "商品コード":
                c_n2 = normalize_code(vals[9])
                sheet_maru_sum[c_n2] = sheet_maru_sum.get(c_n2, 0) + int(to_num(vals[11]))

    latest_ship_map = {}
    for sname in all_sheets:
        if "total" in str(sname).lower():
            df_tot = all_sheets[sname]
            for _, rt in df_tot.iterrows():
                vals = rt.values
                if len(vals) > 5 and clean_str(vals[1]):
                    c_n = normalize_code(vals[1])
                    latest_ship_map[c_n] = int(to_num(vals[5]))

    header_idx = 0
    for idx in range(min(15, len(target_df))):
        row_vals = [str(x) for x in target_df.iloc[idx].values]
        if any("品番" in x for x in row_vals):
            header_idx = idx
            break

    h_row = [str(x).replace("\n", "").replace(" ", "").replace(" ", "") for x in target_df.iloc[header_idx].values]
    col_map = {}
    for i, c in enumerate(h_row):
        if "商品センター" in c and "center" not in col_map: col_map["center"] = i
        elif "品番" in c and "code" not in col_map: col_map["code"] = i
        elif "商品区分" in c and "name" not in col_map: col_map["name"] = i
        elif "73期" in c and "k73" not in col_map: col_map["k73"] = i
        elif "74期" in c and "k74" not in col_map: col_map["k74"] = i
        elif "75期" in c and "k75" not in col_map: col_map["k75"] = i
        elif "76期" in c and "k76" not in col_map: col_map["k76"] = i
        elif "77期" in c and "k77" not in col_map: col_map["k77"] = i
        elif ("2課発注点" in c or "２課発注点" in c) and "point2" not in col_map: col_map["point2"] = i
        elif "発注点" in c and "point1" not in col_map: col_map["point1"] = i
        elif "在庫数" in c and "roji_stock" not in col_map: col_map["roji_stock"] = i
        elif "入荷予定日" in c and "ndate" not in col_map: col_map["ndate"] = i
        elif "入荷予定" in c and "roji_in" not in col_map: col_map["roji_in"] = i
        elif "ロジザード" in c and ("確保" in c or "在庫確保" in c) and "roji_sec" not in col_map: col_map["roji_sec"] = i
        elif "丸藤" in c and "maru_sec" not in col_map: col_map["maru_sec"] = i
        elif ("在庫確保" in c or "弥生" in c) and "ロジザード" not in c and "丸藤" not in c and "yayoi_sec" not in col_map: col_map["yayoi_sec"] = i
        elif ("フリー在庫" in c or "有効数量" in c) and "free" not in col_map: col_map["free"] = i
        elif "回転率" in c and "rate" not in col_map: col_map["rate"] = i
        elif ("リンベル" in c or "ﾘﾝﾍﾞﾙ" in c) and "rinbell" not in col_map: col_map["rinbell"] = i
        elif "前回決定事項" in c and "prev_dec" not in col_map: col_map["prev_dec"] = i
        elif "決定事項" in c and "curr_dec" not in col_map: col_map["curr_dec"] = i

    c_center = col_map.get("center", 1)
    c_code = col_map.get("code", 2)
    c_name = col_map.get("name", 3)
    c_k73 = col_map.get("k73", 6)
    c_k74 = col_map.get("k74", 7)
    c_k75 = col_map.get("k75", 9)
    c_k76 = col_map.get("k76", 10)
    c_k77 = col_map.get("k77", 11)
    c_point1 = col_map.get("point1", 14)
    c_point2 = col_map.get("point2", -1)
    c_roji_stock = col_map.get("roji_stock", 16)
    c_roji_in = col_map.get("roji_in", 17)
    c_ndate = col_map.get("ndate", -1)
    c_roji_sec = col_map.get("roji_sec", 19)
    c_yayoi_sec = col_map.get("yayoi_sec", 20)
    c_maru_sec = col_map.get("maru_sec", 21)
    c_rate = col_map.get("rate", 26)
    c_rin = col_map.get("rinbell", 27)
    c_prev = col_map.get("prev_dec", 28)
    c_curr = col_map.get("curr_dec", 29)

    raw_h = target_df.iloc[header_idx].values
    prev_col_title = clean_str(raw_h[c_prev]) if c_prev < len(raw_h) and clean_str(raw_h[c_prev]) else "前回決定事項"

    items = []
    for idx in range(header_idx + 1, len(target_df)):
        r = target_df.iloc[idx].values
        if len(r) <= c_code:
            continue
        hinban = clean_str(r[c_code])
        if not hinban or hinban == "品番":
            continue
        c_norm = normalize_code(hinban)

        k77_val = int(to_num(r[c_k77])) if c_k77 < len(r) and not is_na_val(r[c_k77]) else latest_ship_map.get(c_norm, 0)
        p1_val = int(round(to_num(r[c_point1]), 0)) if c_point1 < len(r) and not is_na_val(r[c_point1]) else int(round(k77_val / 12 * 5, 0))
        p2_val = int(round(to_num(r[c_point2]), 0)) if 0 <= c_point2 < len(r) and not is_na_val(r[c_point2]) else 0

        stock_roji = z_free_sum.get(c_norm, int(to_num(r[c_roji_stock])) if c_roji_stock < len(r) else 0)
        in_roji = nyuka_sum.get(c_norm, int(to_num(r[c_roji_in])) if c_roji_in < len(r) else 0)
        sec_roji_cell = int(to_num(r[c_roji_sec])) if c_roji_sec < len(r) else 0
        sec_roji = z_sec_sum.get(c_norm, sec_roji_cell) if z_sec_sum.get(c_norm, 0) > 0 else sec_roji_cell

        cell_yayoi = int(to_num(r[c_yayoi_sec])) if (0 <= c_yayoi_sec < len(r)) else 0
        cell_maru = int(to_num(r[c_maru_sec])) if (0 <= c_maru_sec < len(r)) else 0
        y_sec_init = sheet_yayoi_sum.get(c_norm, cell_yayoi)
        m_sec_init = sheet_maru_sum.get(c_norm, cell_maru)

        free_val = (stock_roji + sec_roji + in_roji) - (y_sec_init + m_sec_init)
        flag_val = free_val - (p1_val + p2_val)

        rate_raw = to_num(r[c_rate]) if c_rate < len(r) else 0.0
        rate_val = int(round(rate_raw * 100, 0)) if abs(rate_raw) <= 5.0 else int(round(rate_raw, 0))
        ndate_val = clean_arrival_date(r[c_ndate]) if (0 <= c_ndate < len(r)) else ""

        items.append({
            "No": int(to_num(r[0])) if not is_na_val(r[0]) else idx,
            "商品センター": clean_str(r[c_center]) if c_center < len(r) else "",
            "品番": hinban,
            "品番_正規化": c_norm,
            "商品区分": clean_str(r[c_name]) if c_name < len(r) else prod_name_sheet_map.get(c_norm, ""),
            "入数": lot_map.get(c_norm, 1),
            "73期": int(to_num(r[c_k73])) if c_k73 < len(r) else 0,
            "74期": int(to_num(r[c_k74])) if c_k74 < len(r) else 0,
            "75期": int(to_num(r[c_k75])) if c_k75 < len(r) else 0,
            "76期": int(to_num(r[c_k76])) if c_k76 < len(r) else 0,
            "77期": k77_val,
            "1課発注点": p1_val,
            "2課発注点": p2_val,
            "在庫数_ロジザード": stock_roji,
            "入荷予定数_ロジザード": in_roji,
            "入荷予定日_ロジザード": ndate_val,
            "確保数_ロジザード": sec_roji,
            "在庫確保_弥生": y_sec_init,
            "在庫確保_丸藤": m_sec_init,
            "確保数のズレ": sec_roji - (y_sec_init + m_sec_init),
            "入荷後のフリー在庫": free_val,
            "発注検討フラグ": flag_val,
            "在庫回転率": rate_val,
            "リンベル確保数": clean_str(r[c_rin]) if c_rin < len(r) else "",
            "前回決定事項": clean_str(r[c_prev]) if c_prev < len(r) else "",
            "決定事項": clean_str(r[c_curr]) if c_curr < len(r) else "",
            "発注日": "",
            "関連会社検索用": ""
        })

    df_main = pd.DataFrame(items)

    hist_list = []
    for s_name, df_s in all_sheets.items():
        s_str = str(s_name).strip()
        if s_str.isdigit() and len(s_str) == 5 and not df_s.empty:
            h_row_vals = [clean_str(x) for x in df_s.iloc[0].values]
            h_col_idx = next((i for i, c in enumerate(h_row_vals) if c == "品番"), 2)
            k_cols = [i for i, c in enumerate(h_row_vals) if c and "決定事項" in c and "前回" not in c]
            s_clean = s_str.lstrip("0")
            rest = s_clean[1:] if s_clean.startswith("8") else s_clean
            dt_str = f"2026-{rest[:2]}-{rest[2:]}" if len(rest) == 4 else (f"2026-0{rest[0]}-{rest[1:]}" if len(rest) == 3 else s_str)
            for _, r_d in df_s.iloc[1:].iterrows():
                r_vals = r_d.values
                if len(r_vals) <= h_col_idx:
                    continue
                c_val = clean_str(r_vals[h_col_idx])
                if not c_val:
                    continue
                for kc in k_cols:
                    if kc < len(r_vals) and clean_str(r_vals[kc]) not in ("", "0"):
                        hist_list.append({"会議日": dt_str, "シート": s_str, "品番": c_val, "決定事項": clean_str(r_vals[kc])})

    df_hist = pd.DataFrame(hist_list).drop_duplicates(subset=["会議日", "品番", "決定事項"]) if hist_list else pd.DataFrame(columns=["会議日", "シート", "品番", "決定事項"])
    extra_meta = {
        "z_free_sum": z_free_sum,
        "nyuka_sum": nyuka_sum,
        "lot_map": lot_map,
        "prod_name_sheet_map": prod_name_sheet_map,
        "sheet_yayoi_sum": sheet_yayoi_sum,
        "sheet_maru_sum": sheet_maru_sum
    }
    return df_main, df_hist, prev_col_title, extra_meta, target_sheet_name, all_sheets

@st.cache_data
def load_yayoi_data(file_bytes_or_path, all_sheets_from_main=None):
    """
    オートフィルター付きの『在庫確保.xlsx』や『2027年01月12日』といった日付も
    1行も漏らさず読み込む強化版関数
    """
    df = pd.DataFrame()
    used_src_name = ""
    sources = [
        (file_bytes_or_path, "Googleドライブ/アップロード (在庫確保)"),
        (find_local_file(["在庫確保"]), "ローカル: 在庫確保.xlsx"),
        (find_local_file(["弥生確保リスト", "確保リスト"]), "ローカル: 弥生確保リスト.xlsx")
    ]
    for src, s_label in sources:
        if src is not None:
            try:
                df_raw = read_excel_safely(src, sheet_name=0, header=None)
                if df_raw.empty:
                    continue
                h_idx = 0
                for i in range(min(20, len(df_raw))):
                    row_v = [clean_str(x) for x in df_raw.iloc[i].values]
                    if any(k in row_v for k in ["商品コード", "品番", "型番", "得意先名", "得意先名称", "受注残数"]):
                        h_idx = i
                        break
                df_cand = read_excel_safely(src, sheet_name=0, header=h_idx)
                if not df_cand.empty:
                    df = df_cand
                    used_src_name = s_label
                    break
            except Exception:
                continue

    if not df.empty:
        cols_str = [clean_str(c) for c in df.columns]
        def find_col(cands, fallback_idx=None):
            for cand in cands:
                for orig_c, clean_c in zip(df.columns, cols_str):
                    if cand == clean_c or cand in clean_c:
                        return orig_c
            if fallback_idx is not None and len(df.columns) > fallback_idx:
                return df.columns[fallback_idx]
            return None

        col_c_code = find_col(["商品コード", "品番", "型番", "機種名"], 13)
        col_c_qty = find_col(["受注残数", "数量", "確保数", "残数"], 21)
        col_c_cust = find_col(["得意先名", "得意先名称", "会社名", "得意先"], 5)
        col_c_ord = find_col(["受注番号", "伝票番号"], 2)
        col_c_odate = find_col(["受注日", "伝票日付"], 1)
        col_c_ddate = find_col(["納入期日", "納期"], 32)
        col_c_wh = find_col(["倉庫名", "倉庫"], 19)
        col_c_staff = find_col(["担当者", "担当者名", "入力者"], 10)
        col_c_note = find_col(["備考", "摘要", "メモ"], 30)

        if col_c_code is not None and col_c_qty is not None:
            df = df[df[col_c_code].notna()].copy()
            df["商品コード"] = df[col_c_code].apply(clean_str)
            df = df[~df["商品コード"].isin(["", "商品コード", "品番", "型番", "合計", "総計"])]
            df["得意先名"] = df[col_c_cust].apply(lambda x: clean_str(x) if clean_str(x) else "一般確保") if col_c_cust in df.columns else "一般確保"
            df["数量"] = df[col_c_qty].apply(lambda x: int(to_num(x)))
            df = df[df["数量"] != 0].copy()
            df["受注残数"] = df["数量"]
            df["受注番号"] = df[col_c_ord].apply(clean_str) if col_c_ord in df.columns else ""

            # 2027年01月12日 などの日本語日付を確実にパース（未来日付も一切除外しない）
            if col_c_odate in df.columns:
                parsed_o = df[col_c_odate].apply(parse_jp_date)
                df["受注日_dt"] = [p[0] for p in parsed_o]
                df["受注日"] = [p[1] for p in parsed_o]
            else:
                df["受注日_dt"] = pd.NaT
                df["受注日"] = ""

            if col_c_ddate in df.columns:
                parsed_d = df[col_c_ddate].apply(parse_jp_date)
                df["納入期日_dt"] = [p[0] for p in parsed_d]
                df["納入期日"] = [p[1] for p in parsed_d]
            else:
                df["納入期日_dt"] = pd.NaT
                df["納入期日"] = ""

            df["倉庫名"] = df[col_c_wh].apply(clean_str) if col_c_wh in df.columns else ""
            df["担当者名"] = df[col_c_staff].apply(clean_str) if col_c_staff in df.columns else ""
            df["備考"] = df[col_c_note].apply(clean_str) if col_c_note in df.columns else ""

            empty_mask = df["受注番号"] == ""
            if empty_mask.any():
                df.loc[empty_mask, "受注番号"] = [f"K-{idx:05d}" for idx in range(1, empty_mask.sum() + 1)]
        else:
            df = pd.DataFrame()

    # 発注検討表の『確保貼り付け』シートにある明細で、万が一『在庫確保.xlsx』に載っていない品番があれば補完する
    if all_sheets_from_main and "確保貼り付け" in all_sheets_from_main:
        existing_codes = set(df["商品コード"].apply(normalize_code)) if not df.empty else set()
        supp_rows = []
        df_k = all_sheets_from_main["確保貼り付け"].iloc[4:]
        for idx_k, rk in df_k.iterrows():
            vals = rk.values
            if len(vals) > 11 and clean_str(vals[9]) and clean_str(vals[9]) != "商品コード":
                c_raw = clean_str(vals[9])
                if normalize_code(c_raw) not in existing_codes:
                    supp_rows.append({
                        "受注番号": f"MARU-{idx_k:04d}",
                        "得意先名": "株式会社 丸藤（確保貼付シート分）",
                        "商品コード": c_raw,
                        "商品名/摘要": clean_str(vals[10]),
                        "数量": int(to_num(vals[11])),
                        "受注残数": int(to_num(vals[11])),
                        "受注日": "",
                        "納入期日": "",
                        "備考": clean_str(vals[12]) if len(vals) > 12 else "",
                        "倉庫名": "",
                        "担当者名": "",
                        "受注日_dt": pd.NaT,
                        "納入期日_dt": pd.NaT
                    })
            if len(vals) > 14 and clean_str(vals[13]) and clean_str(vals[13]) not in ("行ラベル", "総計", "(空白)"):
                c_raw2 = clean_str(vals[13])
                if normalize_code(c_raw2) not in existing_codes:
                    supp_rows.append({
                        "受注番号": f"YAYOI-{idx_k:04d}",
                        "得意先名": "弥生確保（確保貼付シート分）",
                        "商品コード": c_raw2,
                        "商品名/摘要": "",
                        "数量": int(to_num(vals[14])),
                        "受注残数": int(to_num(vals[14])),
                        "受注日": "",
                        "納入期日": "",
                        "備考": "",
                        "倉庫名": "",
                        "担当者名": "",
                        "受注日_dt": pd.NaT,
                        "納入期日_dt": pd.NaT
                    })
        if supp_rows:
            df_supp = pd.DataFrame(supp_rows)
            df = pd.concat([df, df_supp], ignore_index=True) if not df.empty else df_supp
            if not used_src_name:
                used_src_name = "発注検討表内の『確保貼り付け』シート"

    if os.path.exists(MANUAL_YAYOI_CSV):
        try:
            df_man = pd.read_csv(MANUAL_YAYOI_CSV).fillna("")
            if not df_man.empty:
                df_man["数量"] = pd.to_numeric(df_man["数量"], errors="coerce").fillna(0).astype(int)
                df_man["受注残数"] = df_man["数量"]
                df_man["受注日_dt"] = pd.to_datetime(df_man["受注日"], errors="coerce")
                df_man["納入期日_dt"] = pd.to_datetime(df_man["納入期日"], errors="coerce")
                edited_onos = set(df_man["受注番号"].astype(str).unique())
                df = df[~df["受注番号"].astype(str).isin(edited_onos)]
                df = pd.concat([df, df_man], ignore_index=True)
        except Exception:
            pass

    if not df.empty:
        # 2027年などの未来の受注日も一番上に来るように降順ソート
        df = df.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
    return df, used_src_name

# ---------------------------------------------------------
# ヘッダー ＆ Googleドライブ同期ボタン ＆ アップローダー
# ---------------------------------------------------------
head_col1, head_col2, head_col3 = st.columns([2.7, 1.1, 1.8])
with head_col1:
    st.markdown('<div class="app-header">📊 発注検討表 ＆ 弥生受注管理システム（Googleドライブ完全連動版）</div>', unsafe_allow_html=True)
with head_col2:
    if st.button("🔄 Googleドライブから最新データを同期", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
with head_col3:
    with st.expander("📂 手動アップロード / GDrive設定", expanded=False):
        gdrive_url_input = st.text_input("☁️ ①発注検討表 GoogleドライブURL", value=GDRIVE_ORDER_URL)
        gdrive_kakuho_input = st.text_input("☁️ ②在庫確保.xlsx GoogleドライブURL (任意)", value=GDRIVE_KAKUHO_URL, placeholder="空欄でも自動検索します。直接指定する場合はURLを貼付")
        u_c1, u_c2 = st.columns(2)
        up_main = u_c1.file_uploader("📋 発注検討表.xlsx", type=["xlsx"], key="up_m")
        up_yayoi = u_c2.file_uploader("📑 在庫確保.xlsx", type=["xlsx"], key="up_y")
        u_c3, u_c4 = st.columns(2)
        up_comp = u_c3.file_uploader("🏢 弥生会社名.xlsx", type=["xlsx"], key="up_comp")
        up_item = u_c4.file_uploader("📦 弥生商品番号.xlsx", type=["xlsx"], key="up_item")

gdrive_sync = sync_all_files_from_gdrive(gdrive_url_input, gdrive_kakuho_input)

if gdrive_sync["error_detail"] and up_main is None:
    st.warning(f"⚠️ **Googleドライブ未接続**\n{gdrive_sync['error_detail']}")

main_src = up_main if up_main is not None else (gdrive_sync["main_bytes"] if gdrive_sync["main_bytes"] is not None else find_local_file(["発注検討表"]))
yayoi_src = up_yayoi if up_yayoi is not None else gdrive_sync["yayoi_bytes"]
comp_src = up_comp if up_comp is not None else gdrive_sync["comp_bytes"]
item_src = up_item if up_item is not None else gdrive_sync["item_bytes"]

df_main, df_hist, prev_col_title, extra_meta, loaded_sheet_name, all_sheets_raw = load_excel_master(main_src)
df_yayoi, yayoi_src_label = load_yayoi_data(yayoi_src, all_sheets_raw)
if gdrive_sync["yayoi_filename"] and up_yayoi is None and "ローカル" not in yayoi_src_label:
    yayoi_src_label = gdrive_sync["yayoi_filename"]

master_comps = load_company_master(comp_src)
master_items = load_item_master(item_src)

if df_main.empty:
    st.error("⚠️ 『発注検討表.xlsx』が読み込めませんでした。Googleドライブの共有権限を確認するか、右上のエクスパンダーからアップロードしてください。")
    st.stop()

marufuji_dict = {}
yayoi_dict = {}
yayoi_comp_search_map = {}

if not df_yayoi.empty:
    df_yayoi["品番_正規化"] = df_yayoi["商品コード"].apply(normalize_code)

    def is_marufuji_name(name):
        n_clean = str(name).strip().replace(" ", " ")
        return any(target.replace(" ", " ") in n_clean for target in MARUFUJI_TARGET_NAMES) or ("丸藤" in n_clean)

    df_yayoi["is_maru"] = df_yayoi["得意先名"].apply(is_marufuji_name)
    df_maru_only = df_yayoi[df_yayoi["is_maru"]]
    df_non_maru = df_yayoi[~df_yayoi["is_maru"]]

    # 2027年など未来の受注分も含めて全期間・全行をSUMIFS合算
    yayoi_non_maru_totals = df_non_maru.groupby("品番_正規化")["数量"].sum().to_dict()
    marufuji_real_totals = df_maru_only.groupby("品番_正規化")["数量"].sum().to_dict()

    mapped_yayoi = df_main["品番_正規化"].map(yayoi_non_maru_totals)
    mapped_maru = df_main["品番_正規化"].map(marufuji_real_totals)

    df_main["在庫確保_弥生"] = mapped_yayoi.fillna(df_main["在庫確保_弥生"]).astype(int)
    df_main["在庫確保_丸藤"] = mapped_maru.fillna(df_main["在庫確保_丸藤"]).astype(int)

    total_sec_series = df_main["在庫確保_弥生"] + df_main["在庫確保_丸藤"]
    df_main["確保数のズレ"] = df_main["確保数_ロジザード"] - total_sec_series
    df_main["入荷後のフリー在庫"] = (df_main["在庫数_ロジザード"] + df_main["確保数_ロジザード"] + df_main["入荷予定数_ロジザード"]) - total_sec_series
    df_main["発注検討フラグ"] = df_main["入荷後のフリー在庫"] - (df_main["1課発注点"] + df_main["2課発注点"])

    for code_n, grp in df_yayoi.groupby("品番_正規化", sort=False):
        grp_sorted = grp.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
        yayoi_dict[code_n] = grp_sorted[["受注番号", "得意先名", "数量", "受注日", "納入期日", "倉庫名", "備考", "担当者名"]].to_dict(orient="records")
        yayoi_comp_search_map[code_n] = " ".join(grp_sorted["得意先名"].unique())
        maru_grp = grp_sorted[grp_sorted["is_maru"]]
        if not maru_grp.empty:
            marufuji_dict[code_n] = maru_grp[["受注番号", "得意先名", "数量", "受注日", "納入期日", "倉庫名", "備考", "担当者名"]].to_dict(orient="records")

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

target_norm_codes = set(df_main["品番_正規化"].unique())
df_main["関連会社検索用"] = df_main["品番_正規化"].map(yayoi_comp_search_map).fillna("")
m_df = df_main.copy()

unregistered_list = []
all_kakuho_check = {}
if not df_yayoi.empty:
    for c_n, grp in df_yayoi.groupby("品番_正規化"):
        all_kakuho_check[c_n] = int(grp["数量"].sum())
else:
    for c_n, q in extra_meta.get("sheet_yayoi_sum", {}).items():
        all_kakuho_check[c_n] = all_kakuho_check.get(c_n, 0) + q
    for c_n, q in extra_meta.get("sheet_maru_sum", {}).items():
        all_kakuho_check[c_n] = all_kakuho_check.get(c_n, 0) + q

for c_n, k_qty in all_kakuho_check.items():
    if k_qty > 0 and c_n not in target_norm_codes:
        z_f = extra_meta.get("z_free_sum", {}).get(c_n, 0)
        n_q = extra_meta.get("nyuka_sum", {}).get(c_n, 0)
        p_nm = item_name_map.get(c_n) or extra_meta.get("prod_name_sheet_map", {}).get(c_n, "マスター未登録（特注/枝番）")
        unregistered_list.append({
            "品番": c_n,
            "商品名": p_nm,
            "弥生確保数": k_qty,
            "ロジザード在庫": z_f,
            "入荷予定数": n_q,
            "差引フリー在庫": z_f + n_q - k_qty,
            "入数": extra_meta.get("lot_map", {}).get(c_n, 1)
        })
df_unreg = pd.DataFrame(unregistered_list)

hist_dict = {}
if df_hist is not None and not df_hist.empty:
    for p_code, grp in df_hist.groupby("品番", sort=False):
        hist_dict[p_code] = grp.sort_values("会議日", ascending=False)[["会議日", "決定事項"]].to_dict(orient="records")

src_status = f"✅ Googleドライブ同期完了 ({gdrive_sync['source_label']})" if gdrive_sync["main_bytes"] is not None else "📂 ローカル参照中"
st.caption(
    f"**{src_status}** ｜ 参照シート: **{loaded_sheet_name}** ({len(m_df):,} 品番) ｜ "
    f"📑 確保データ元: **{yayoi_src_label or '未検出'}** (全 **{len(df_yayoi):,}** 行・2027年以降の先行確保も全件反映済) ｜ "
    f"🚨 要発注（検討フラグがマイナス）: **{int((m_df['発注検討フラグ'] < 0).sum())}** 件"
)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    f"📋 1. 発注検討表（SUMIFS完全合算 ＆ 高速検索）",
    "🏢 2. 会社（得意先）別・受注番号別 確保リスト",
    "💻 3. 弥生受注伝票（リアルタイム発注判定つき）",
    "🕒 4. 入荷予定 ＆ 過去の会議決定履歴",
    f"🚀 5. 営業用・大口注文即判定シミュレーター ＆ 抜け防止 ({len(df_unreg)}件)"
])

with tab1:
    center_list = sorted([x for x in m_df["商品センター"].unique() if x])
    centers_json = json.dumps(center_list, ensure_ascii=False)
    records_json = json.dumps(m_df.to_dict(orient="records"), ensure_ascii=False)
    yayoi_json = json.dumps(yayoi_dict, ensure_ascii=False)
    marufuji_json = json.dumps(marufuji_dict, ensure_ascii=False)
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
        .instant-filter-bar {{
            display: grid; grid-template-columns: 2.3fr 1.2fr 3.0fr; gap: 14px;
            background: #f8fafc; padding: 10px 14px; border: 1px solid #cbd5e1;
            border-radius: 6px; margin-bottom: 8px;
        }}
        .filter-group {{ display: flex; flex-direction: column; gap: 4px; }}
        .filter-label {{ font-size: 12px; font-weight: bold; color: #334155; }}
        .filter-select, .filter-input {{
            height: 34px; border: 1px solid #94a3b8; border-radius: 5px;
            padding: 4px 10px; font-size: 13px; background: #fff; color: #0f172a;
        }}
        .filter-input:focus, .filter-select:focus {{ outline: 2px solid #2563eb; border-color: transparent; }}

        .toolbar {{
            display: flex; align-items: center; justify-content: space-between;
            background: #f1f5f9; padding: 6px 12px; border: 1px solid #cbd5e1;
            border-bottom: none; border-radius: 6px 6px 0 0;
        }}
        .toolbar-info {{ font-size: 13px; color: #1e293b; font-weight: bold; }}
        .btn-download {{
            padding: 6px 14px; font-size: 12px; font-weight: bold; border-radius: 4px;
            cursor: pointer; border: none; background: #2563eb; color: #fff;
        }}
        .btn-download:hover {{ background: #1d4ed8; }}
        .table-wrap {{ width: 100%; height: 620px; overflow: auto; border: 1px solid #94a3b8; position: relative; }}
        table {{ border-collapse: separate; border-spacing: 0; width: max-content; min-width: 100%; }}
        th, td {{ border-right: 1px solid #94a3b8; border-bottom: 1px solid #94a3b8; padding: 3px 6px; font-size: 12px; height: 30px; }}
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
        .w-name {{ width: 200px; min-width: 200px; max-width: 200px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .w-num {{ width: 72px; min-width: 72px; text-align: right; font-variant-numeric: tabular-nums; }}
        .w-date {{ width: 95px; min-width: 95px; text-align: center; }}
        .w-prev-dec {{ width: 160px; min-width: 160px; max-width: 160px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .w-dec {{ width: 220px; min-width: 220px; }}

        .bg-cyan {{ background-color: #67e8f9 !important; }}
        .bg-flag-neg {{ background-color: #fecdd3 !important; color: #be123c; font-weight: bold; }}
        .bg-green {{ background-color: #a3e635 !important; }}
        .bg-pink-rate {{ background-color: #fbcfe8 !important; color: #be123c; font-weight: bold; }}
        .bg-yellow-ro {{ background-color: #fef08a !important; color: #1e293b; }}
        .bg-yellow-edit {{ background-color: #ffff00 !important; }}
        .bg-hist-ki {{ background-color: #f1f5f9; color: #475569; }}
        .bg-zure-diff {{ background-color: #ffe4e6 !important; color: #be123c; font-weight: bold; }}

        tbody tr:hover td {{ filter: brightness(0.95); }}
        .click-link {{ color: #0284c7; font-weight: bold; text-decoration: underline; cursor: pointer; display: block; width: 100%; }}
        .cell-input {{ width: 100%; height: 23px; border: 1px solid #94a3b8; border-radius: 2px; padding: 1px 4px; font-size: 12px; background: rgba(255,255,255,0.9); }}
        .cell-input-point {{ width: 100%; height: 23px; border: 1px solid #3b82f6; border-radius: 2px; padding: 1px 4px; font-size: 12px; font-weight: bold; text-align: right; background: #eff6ff; }}

        .modal-overlay {{
            display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(15, 23, 42, 0.7); z-index: 9999; justify-content: center; align-items: center;
        }}
        .modal-window {{
            background: #fff; width: 99%; max-width: 1750px; height: 94vh; border-radius: 10px;
            box-shadow: 0 25px 40px -5px rgba(0,0,0,0.45); display: flex; flex-direction: column; overflow: hidden; border: 3px solid #1e40af;
        }}
        .modal-header {{ background: #1e3a8a; color: #fff; padding: 14px 24px; display: flex; justify-content: space-between; align-items: center; }}
        .modal-title {{ font-size: 19px; font-weight: bold; }}
        .modal-close {{ background: #ef4444; color: #fff; border: none; border-radius: 6px; padding: 8px 22px; font-size: 15px; font-weight: bold; cursor: pointer; }}
        .modal-body {{ padding: 20px; overflow-y: auto; display: flex; flex-direction: column; gap: 16px; flex: 1; }}
        .summary-bar {{ display: flex; gap: 22px; flex-wrap: wrap; background: #f1f5f9; padding: 12px 18px; border-radius: 8px; border-left: 6px solid #2563eb; font-size: 15px; align-items: center; }}
        .summary-item {{ font-weight: bold; color: #1e293b; }}
        .summary-item span {{ color: #0284c7; font-size: 17px; margin-left: 4px; }}
        .mode-btn {{ padding: 6px 14px; border-radius: 6px; border: 1px solid #cbd5e1; font-weight: bold; cursor: pointer; font-size: 13.5px; background: #fff; color: #1e293b; }}
        .mode-btn.active {{ background: #ea580c; color: #fff; border-color: #c2410c; }}
        .quick-edit-box {{ background: #fffbeb; border: 2px solid #fde68a; padding: 12px 16px; border-radius: 8px; display: flex; flex-direction: column; gap: 10px; }}
        .stamp-btn {{ background: #e2e8f0; border: 1px solid #94a3b8; border-radius: 5px; padding: 4px 10px; font-size: 13px; cursor: pointer; font-weight: bold; }}
        .modal-grid {{ display: grid; grid-template-columns: 3.5fr 1.5fr; gap: 20px; }}
        .sub-table {{ width: 100%; border-collapse: collapse; }}
        .sub-table th {{ background: #e2e8f0; color: #1e293b; padding: 8px 10px; font-size: 14px; height: auto; position: static; border: 1px solid #cbd5e1; }}
        .sub-table td {{ padding: 8px 10px; font-size: 14px; height: auto; border: 1px solid #cbd5e1; }}
        .sec-title {{ font-size: 16px; font-weight: bold; color: #1e3a8a; margin-bottom: 8px; border-bottom: 2px solid #93c5fd; padding-bottom: 4px; }}
    </style>
    </head>
    <body>
    <div class="instant-filter-bar">
        <div class="filter-group">
            <label class="filter-label">表示絞り込み（瞬時切替）</label>
            <select id="js-filter-mode" class="filter-select" onchange="applyInstantFilter()">
                <option value="all">すべて表示（Excelと同じ並び）</option>
                <option value="need_order">🚨 要発注（発注検討フラグがマイナスのみ）</option>
                <option value="zure">確保数のズレがある商品</option>
                <option value="incoming">入荷予定がある商品</option>
                <option value="secured">弥生確保または丸藤確保がある商品</option>
            </select>
        </div>
        <div class="filter-group">
            <label class="filter-label">商品センター</label>
            <select id="js-filter-center" class="filter-select" onchange="applyInstantFilter()">
                <option value="すべて">すべて</option>
            </select>
        </div>
        <div class="filter-group">
            <label class="filter-label">🔍 品番・商品名・会社名（確保先）・決定事項でリアルタイム検索</label>
            <input type="text" id="js-search-kw" class="filter-input" placeholder="文字を打つと0.01秒で絞り込みます（例: ABO-08PK / ライト / 丸藤 / トップ産業）" oninput="applyInstantFilter()">
        </div>
    </div>

    <div class="toolbar">
        <div class="toolbar-info">表示件数: <span id="row-count">0</span> 件 ｜ 💡 「品番」や「在庫確保」の数字をクリックすると、2027年以降の先行予約も含めた確保内訳POPが開きます！</div>
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
                    <th class="w-num bg-hist-ki">74期<br>実績</th>
                    <th class="w-num bg-hist-ki">75期<br>実績</th>
                    <th class="w-num bg-hist-ki">76期<br>実績</th>
                    <th class="w-num bg-hist-ki">77期<br>今期出荷</th>
                    <th class="w-num" style="background:#e0f2fe;">1課発注点<br><span style="color:#0369a1; font-size:10px;">✏️手動修正可</span></th>
                    <th class="w-num" style="background:#e0f2fe;">2課発注点<br><span style="color:#0369a1; font-size:10px;">✏️手動修正可</span></th>
                    <th class="w-num bg-cyan">在庫数<br><span class="red-txt">ロジザード</span></th>
                    <th class="w-num bg-cyan">入荷予定<br>数<br><span class="red-txt">ロジザード</span></th>
                    <th class="w-date bg-cyan">入荷予定<br>日<br><span class="red-txt">ロジザード</span></th>
                    <th class="w-num bg-cyan">確保数<br><span class="red-txt">ロジザード<br>(在庫確保)</span></th>
                    <th class="w-num bg-cyan">在庫確保<br><span class="red-txt">(弥生)</span></th>
                    <th class="w-num bg-cyan">在庫確保<br><span style="color:#b91c1c; font-weight:bold;">丸藤</span></th>
                    <th class="w-num bg-cyan">確保数の<br>ズレ</th>
                    <th class="w-num bg-cyan">入荷後の<br>フリー在庫</th>
                    <th class="w-num">発注検討<br>フラグ<br><span class="red-txt">マイナスの<br>場合発注検討</span></th>
                    <th class="w-num">在庫回転<br>率</th>
                    <th class="w-num">リンベル<br>確保数</th>
                    <th class="w-prev-dec bg-yellow-ro">{prev_title_html}</th>
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
                <div class="modal-title" id="modal-title-text">📦 品番詳細・確保内訳データ</div>
                <div style="display:flex; gap:10px; align-items:center;">
                    <button class="mode-btn" id="btn-mode-yayoi" onclick="switchModalMode('yayoi')">🏢 弥生確保リスト（すべて）</button>
                    <button class="mode-btn" id="btn-mode-maru" onclick="switchModalMode('marufuji')">🏢 丸藤確保のみ表示</button>
                    <button class="modal-close" onclick="closeModal()">✕ 閉じる</button>
                </div>
            </div>
            <div class="modal-body">
                <div class="summary-bar" id="modal-summary"></div>
                <div class="quick-edit-box">
                    <div style="font-size:14px; font-weight:bold; color:#92400e; display:flex; justify-content:space-between;">
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
                        <input type="text" id="modal-input-ndate" class="cell-input" style="width:140px; height:32px; font-size:14px;" placeholder="例: 11/20" oninput="syncFromModal()">
                        <label style="font-size:14px; font-weight:bold;">決定事項:</label>
                        <input type="text" id="modal-input-dec" class="cell-input" style="flex:1; height:32px; font-size:14px;" oninput="syncFromModal()">
                        <label style="font-size:14px; font-weight:bold;">発注日:</label>
                        <input type="text" id="modal-input-odate" class="cell-input" style="width:130px; height:32px; font-size:14px;" oninput="syncFromModal()">
                    </div>
                </div>
                <div class="modal-grid">
                    <div>
                        <div class="sec-title" id="modal-sub-title">🏢 確保明細リスト</div>
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
        const centerOptions = {centers_json};
        const yayoiMap = {yayoi_json};
        const marufujiMap = {marufuji_json};
        const histMap = {hist_json};
        let currentModalIdx = null;
        let currentModalMode = "yayoi";
        let visibleIndices = [];

        const centerSel = document.getElementById("js-filter-center");
        centerOptions.forEach(c => {{
            const opt = document.createElement("option");
            opt.value = c;
            opt.textContent = c;
            centerSel.appendChild(opt);
        }});

        function cleanInvalidNdate(val) {{
            if (!val) return "";
            const s = String(val).trim();
            if (/^[0-9,.]+$/.test(s)) return "";
            return s;
        }}

        const savedEdits = JSON.parse(localStorage.getItem("hattyu_edits_v6") || "{{}}");
        tableData.forEach((r, idx) => {{
            r["_origIdx"] = idx;
            if (savedEdits[r["品番"]]) {{
                r["1課発注点"] = savedEdits[r["品番"]].p1 ?? r["1課発注点"];
                r["2課発注点"] = savedEdits[r["品番"]].p2 ?? r["2課発注点"];
                r["入荷予定日_ロジザード"] = cleanInvalidNdate(savedEdits[r["品番"]].ndate ?? r["入荷予定日_ロジザード"]);
                r["決定事項"] = savedEdits[r["品番"]].dec ?? r["決定事項"];
                r["発注日"] = savedEdits[r["品番"]].odate ?? r["発注日"];
                const pSum = Number(r["1課発注点"] || 0) + Number(r["2課発注点"] || 0);
                r["発注検討フラグ"] = Number(r["入荷後のフリー在庫"] || 0) - pSum;
            }} else {{
                r["入荷予定日_ロジザード"] = cleanInvalidNdate(r["入荷予定日_ロジザード"]);
            }}
        }});

        function fmtNum(val) {{
            if (val === "" || val === null || val === undefined) return "";
            const n = Number(val);
            return isNaN(n) ? val : n.toLocaleString("ja-JP");
        }}

        function handleKeyDown(event, colName, origIdx) {{
            if (event.key === "Enter") {{
                event.preventDefault();
                const pos = visibleIndices.indexOf(origIdx);
                if (pos !== -1 && pos + 1 < visibleIndices.length) {{
                    const nextOrigIdx = visibleIndices[pos + 1];
                    const nextInput = document.getElementById(`${{colName}}-${{nextOrigIdx}}`);
                    if (nextInput) {{
                        nextInput.focus();
                        if (nextInput.select) nextInput.select();
                    }}
                }}
            }}
        }}

        function recalculateRow(origIdx) {{
            const r = tableData[origIdx];
            const pSum = Number(r["1課発注点"] || 0) + Number(r["2課発注点"] || 0);
            r["発注検討フラグ"] = Number(r["入荷後のフリー在庫"] || 0) - pSum;

            const flagCell = document.getElementById(`flag-${{origIdx}}`);
            if (flagCell) {{
                flagCell.textContent = fmtNum(r["発注検討フラグ"]);
                flagCell.className = (r["発注検討フラグ"] < 0) ? "w-num bg-flag-neg" : "w-num bg-cyan";
            }}
            saveToStorage(r["品番"], r);
        }}

        function applyInstantFilter() {{
            const fMode = document.getElementById("js-filter-mode").value;
            const fCenter = document.getElementById("js-filter-center").value;
            const kw = document.getElementById("js-search-kw").value.trim().toLowerCase();

            let filtered = tableData.filter(r => {{
                if (fMode === "need_order" && !(Number(r["発注検討フラグ"]) < 0)) return false;
                if (fMode === "zure" && !(Number(r["確保数のズレ"]) !== 0)) return false;
                if (fMode === "incoming" && !(Number(r["入荷予定数_ロジザード"]) > 0 || (r["入荷予定日_ロジザード"] && r["入荷予定日_ロジザード"] !== ""))) return false;
                if (fMode === "secured" && !(Number(r["在庫確保_弥生"]) > 0 || Number(r["在庫確保_丸藤"]) > 0)) return false;

                if (fCenter !== "すべて" && r["商品センター"] !== fCenter) return false;

                if (kw) {{
                    const targetStr = [
                        r["品番"], r["商品区分"], r["前回決定事項"], r["決定事項"], r["関連会社検索用"]
                    ].join(" ").toLowerCase();
                    if (!targetStr.includes(kw)) return false;
                }}
                return true;
            }});

            visibleIndices = filtered.map(r => r["_origIdx"]);
            renderTable(filtered);
        }}

        function renderTable(rowsToRender) {{
            const tbody = document.getElementById("table-body");
            document.getElementById("row-count").textContent = rowsToRender.length.toLocaleString();
            let html = "";
            rowsToRender.forEach(r => {{
                const idx = r["_origIdx"];
                const flagVal = Number(r["発注検討フラグ"]);
                const flagCls = (flagVal < 0) ? "bg-flag-neg" : "bg-cyan";
                const rateVal = Number(r["在庫回転率"]);
                const rateCls = (rateVal <= 5) ? "bg-pink-rate" : "bg-green";
                const zureVal = Number(r["確保数のズレ"] || 0);
                const zureCls = (zureVal !== 0) ? "bg-zure-diff" : "bg-cyan";

                html += `<tr>
                    <td class="w-no sticky-col-1">${{r["No"]}}</td>
                    <td class="w-center sticky-col-2">${{r["商品センター"]}}</td>
                    <td class="w-code sticky-col-3" title="${{r["品番"]}}"><span class="click-link" onclick="openModal(${{idx}}, 'yayoi')">${{r["品番"]}}</span></td>
                    <td class="w-name sticky-col-4" title="${{r["商品区分"]}}">${{r["商品区分"]}}</td>
                    <td class="w-num bg-hist-ki">${{fmtNum(r["74期"])}}</td>
                    <td class="w-num bg-hist-ki">${{fmtNum(r["75期"])}}</td>
                    <td class="w-num bg-hist-ki">${{fmtNum(r["76期"])}}</td>
                    <td class="w-num bg-hist-ki" style="font-weight:bold; color:#0f172a;">${{fmtNum(r["77期"])}}</td>
                    <td class="w-num"><input type="number" class="cell-input-point" id="point1-${{idx}}" value="${{r["1課発注点"]}}" oninput="tableData[${{idx}}]['1課発注点']=Number(this.value); recalculateRow(${{idx}})" onkeydown="handleKeyDown(event, 'point1', ${{idx}})"></td>
                    <td class="w-num"><input type="number" class="cell-input-point" id="point2-${{idx}}" value="${{r["2課発注点"]}}" oninput="tableData[${{idx}}]['2課発注点']=Number(this.value); recalculateRow(${{idx}})" onkeydown="handleKeyDown(event, 'point2', ${{idx}})"></td>
                    <td class="w-num bg-cyan">${{fmtNum(r["在庫数_ロジザード"])}}</td>
                    <td class="w-num bg-cyan">${{fmtNum(r["入荷予定数_ロジザード"])}}</td>
                    <td class="w-date bg-cyan"><input type="text" class="cell-input" id="ndate-${{idx}}" value="${{r["入荷予定日_ロジザード"] || ""}}" placeholder="日付入力" oninput="updateCell(${{idx}}, '入荷予定日_ロジザード', this.value)" onkeydown="handleKeyDown(event, 'ndate', ${{idx}})"></td>
                    <td class="w-num bg-cyan">${{fmtNum(r["確保数_ロジザード"])}}</td>
                    <td class="w-num bg-cyan"><span class="click-link" onclick="openModal(${{idx}}, 'yayoi')">${{fmtNum(r["在庫確保_弥生"])}}</span></td>
                    <td class="w-num bg-cyan"><span class="click-link" style="color:#b91c1c;" onclick="openModal(${{idx}}, 'marufuji')">${{fmtNum(r["在庫確保_丸藤"])}}</span></td>
                    <td class="w-num ${{zureCls}}" id="zure-${{idx}}">${{fmtNum(r["確保数のズレ"])}}</td>
                    <td class="w-num bg-cyan">${{fmtNum(r["入荷後のフリー在庫"])}}</td>
                    <td class="w-num ${{flagCls}}" id="flag-${{idx}}">${{fmtNum(r["発注検討フラグ"])}}</td>
                    <td class="w-num ${{rateCls}}">${{fmtNum(r["在庫回転率"]) + "%"}}</td>
                    <td class="w-num">${{fmtNum(r["リンベル確保数"])}}</td>
                    <td class="w-prev-dec bg-yellow-ro" title="${{r["前回決定事項"] || ""}}">${{r["前回決定事項"] || ""}}</td>
                    <td class="w-dec bg-yellow-edit"><input type="text" class="cell-input" id="dec-${{idx}}" value="${{r["決定事項"] || ""}}" oninput="updateCell(${{idx}}, '決定事項', this.value)" onkeydown="handleKeyDown(event, 'dec', ${{idx}})"></td>
                    <td class="w-date bg-yellow-edit"><input type="text" class="cell-input" id="odate-${{idx}}" value="${{r["発注日"] || ""}}" oninput="updateCell(${{idx}}, '発注日', this.value)" onkeydown="handleKeyDown(event, 'odate', ${{idx}})"></td>
                </tr>`;
            }});
            tbody.innerHTML = html;
        }}

        function saveToStorage(hinban, rowObj) {{
            const store = JSON.parse(localStorage.getItem("hattyu_edits_v6") || "{{}}");
            store[hinban] = {{
                p1: rowObj["1課発注点"],
                p2: rowObj["2課発注点"],
                ndate: rowObj["入荷予定日_ロジザード"],
                dec: rowObj["決定事項"],
                odate: rowObj["発注日"]
            }};
            localStorage.setItem("hattyu_edits_v6", JSON.stringify(store));
        }}

        function updateCell(idx, field, val) {{
            tableData[idx][field] = val;
            saveToStorage(tableData[idx]["品番"], tableData[idx]);
        }}

        function switchModalMode(mode) {{
            if (currentModalIdx === null) return;
            openModal(currentModalIdx, mode);
        }}

        function openModal(idx, mode) {{
            currentModalIdx = idx;
            currentModalMode = mode;
            const r = tableData[idx];
            const code = r["品番"];
            const codeNorm = r["品番_正規化"] || code;

            document.getElementById("btn-mode-yayoi").className = (mode === 'yayoi') ? "mode-btn active" : "mode-btn";
            document.getElementById("btn-mode-maru").className = (mode === 'marufuji') ? "mode-btn active" : "mode-btn";

            document.getElementById("modal-title-text").innerHTML = `📦 品番: 【${{code}}】 ${{r["商品区分"]}}`;
            document.getElementById("modal-summary").innerHTML = `
                <div class="summary-item">74〜77期出荷: <span>${{fmtNum(r["74期"])}} / ${{fmtNum(r["75期"])}} / ${{fmtNum(r["76期"])}} / ${{fmtNum(r["77期"])}}</span></div>
                <div class="summary-item">1課/2課発注点: <span>${{fmtNum(r["1課発注点"])}} / ${{fmtNum(r["2課発注点"])}}</span></div>
                <div class="summary-item">ロジ在庫: <span>${{fmtNum(r["在庫数_ロジザード"])}}</span></div>
                <div class="summary-item">弥生確保(一般): <span>${{fmtNum(r["在庫確保_弥生"])}}</span></div>
                <div class="summary-item">丸藤確保: <span style="color:#b91c1c;">${{fmtNum(r["在庫確保_丸藤"])}}</span></div>
                <div class="summary-item">入荷後のフリー在庫: <span>${{fmtNum(r["入荷後のフリー在庫"])}}</span></div>
                <div class="summary-item">発注検討フラグ: <span style="color:#e11d48;">${{fmtNum(r["発注検討フラグ"])}}</span></div>
            `;
            document.getElementById("modal-view-prev").value = r["前回決定事項"] || "";
            document.getElementById("modal-input-ndate").value = r["入荷予定日_ロジザード"] || "";
            document.getElementById("modal-input-dec").value = r["決定事項"] || "";
            document.getElementById("modal-input-odate").value = r["発注日"] || "";

            let targetList = [];
            if (mode === 'marufuji') {{
                document.getElementById("modal-sub-title").innerHTML = "🏢 【丸藤専用】確保リスト（2027年以降を含む全受注を表示）";
                targetList = marufujiMap[codeNorm] || [];
            }} else {{
                document.getElementById("modal-sub-title").innerHTML = "🏢 得意先別・弥生確保リスト（2027年以降を含む全受注を表示）";
                targetList = yayoiMap[codeNorm] || [];
            }}

            if (targetList.length === 0) {{
                document.getElementById("modal-yayoi-area").innerHTML = `<div style="padding:16px; font-size:14.5px; color:#64748b; background:#f8fafc;">該当する受注確保データはありません。</div>`;
            }} else {{
                let yHtml = `<table class="sub-table"><thead><tr><th>受注番号</th><th>得意先名</th><th>数量</th><th>受注日</th><th>納入期日</th><th>倉庫名</th><th>担当者</th><th>備考</th></tr></thead><tbody>`;
                targetList.forEach(y => {{
                    yHtml += `<tr>
                        <td style="text-align:center; font-weight:bold;">${{y["受注番号"] || ""}}</td>
                        <td style="font-weight:bold;">${{y["得意先名"]}}</td>
                        <td style="text-align:right; font-weight:bold; color:#0284c7; font-size:15.5px;">${{fmtNum(y["数量"])}}</td>
                        <td style="text-align:center; font-weight:bold;">${{y["受注日"]}}</td>
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
            const ndateEl = document.getElementById(`ndate-${{currentModalIdx}}`);
            const decEl = document.getElementById(`dec-${{currentModalIdx}}`);
            const odateEl = document.getElementById(`odate-${{currentModalIdx}}`);
            if (ndateEl) ndateEl.value = ndate;
            if (decEl) decEl.value = dec;
            if (odateEl) odateEl.value = odate;
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
            const cols = ["No", "商品センター", "品番", "商品区分", "74期", "75期", "76期", "77期", "1課発注点", "2課発注点", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "確保数_ロジザード", "在庫確保_弥生", "在庫確保_丸藤", "確保数のズレ", "入荷後のフリー在庫", "発注検討フラグ", "在庫回転率", "リンベル確保数", "前回決定事項", "決定事項", "発注日"];
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
        applyInstantFilter();
    </script>
    </body>
    </html>
    """
    components.html(html_tab1, height=740, scrolling=False)

with tab2:
    if df_yayoi.empty:
        st.warning("確保リストのデータがありません。")
    else:
        comp_totals = df_yayoi.groupby("得意先名")["数量"].sum().sort_values(ascending=False)
        comp_opts = ["すべての会社を表示"] + [f"{c} （確保合計: {v:,} 個）" for c, v in comp_totals.items()]
        c1, c2 = st.columns(2)
        sel_comp = c1.selectbox("🏢 会社名（得意先）をプルダウンで選択", comp_opts)
        kw_comp = c2.text_input("🔍 受注番号・得意先名・品番・受注日(2027など)・備考で検索", placeholder="例: 2027 / 00002657 / 丸藤 / ABO-08PK")

        yv = df_yayoi.copy()
        if sel_comp != "すべての会社を表示":
            cname = sel_comp.split(" （確保合計:")[0]
            yv = yv[yv["得意先名"] == cname]
        if kw_comp:
            yv = yv[
                yv["受注番号"].astype(str).str.contains(kw_comp, case=False, na=False) |
                yv["得意先名"].str.contains(kw_comp, case=False, na=False) |
                yv["商品コード"].str.contains(kw_comp, case=False, na=False) |
                yv["受注日"].str.contains(kw_comp, case=False, na=False) |
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
                "受注残数合計": int(grp["数量"].sum()),
                "商品内訳プレビュー": codes_preview,
                "備考": clean_str(first_r.get("備考", ""))
            })

            d_list = []
            for _, dr in grp.iterrows():
                p_name = clean_str(dr.get("商品区分", "")) or clean_str(dr.get("商品名/摘要", "")) or item_name_map.get(clean_str(dr["商品コード"]), "")
                d_list.append({
                    "商品コード": clean_str(dr["商品コード"]),
                    "商品名": p_name,
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
            .modal-overlay {{
                display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
                background: rgba(15, 23, 42, 0.7); z-index: 9999; justify-content: center; align-items: center;
            }}
            .modal-window {{
                background: #fff; width: 99%; max-width: 1750px; height: 94vh; border-radius: 10px;
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
            <div class="toolbar-info">受注伝票件数: <span id="slip-count">0</span> 件（2027年以降の先行受注も含め全件表示）</div>
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
                        <td style="text-align:center; font-weight:bold;">${{s["受注日"]}}</td>
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
                        <td style="text-align:right; font-weight:bold; color:#be123c; background:#fff1f2; font-size:15.5px;">${{fmtNum(d["数量"])}}</td>
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

master_sim_map = {}
for _, r_m in m_df.iterrows():
    master_sim_map[r_m["品番_正規化"]] = {
        "name": r_m["商品区分"],
        "free": int(r_m["入荷後のフリー在庫"]),
        "point": int(r_m["1課発注点"] + r_m["2課発注点"]),
        "lot": max(int(r_m.get("入数", 1)), 1)
    }

def evaluate_order_line(c_code, qty_v):
    if not c_code or qty_v <= 0:
        return ""
    c_n = normalize_code(c_code)
    if c_n not in master_sim_map:
        return "🚨 マスター未登録（要発注確認！）"
    info = master_sim_map[c_n]
    avail = info["free"]
    after_q = avail - qty_v
    margin = after_q - info["point"]
    lot = info["lot"]
    if after_q < 0:
        rec = math.ceil(abs(margin) / lot) * lot
        return f"🚨 欠品・即発注！(不足:{abs(after_q):,} / 推奨発注:{rec:,})"
    elif margin < 0:
        rec = math.ceil(abs(margin) / lot) * lot
        return f"⚠️ 発注点割れ(残:{after_q:,} / 推奨発注:{rec:,})"
    return f"✅ 在庫OK(残:{after_q:,})"

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
                {"商品コード": "", "商品名": "", "倉庫名": "", "受注残数(数量)": 0, "備考": "", "発注判定": ""}
                for _ in range(10)
            ])
        }

    if "slip_editor_state" not in st.session_state:
        st.session_state["slip_editor_state"] = create_blank_slip(next_order_no_str)

    slip_state = st.session_state["slip_editor_state"]
    v_key = slip_state.get("version", 1)

    @st.dialog("🔍 登録済み伝票を検索・呼び出し", width="large")
    def open_yayoi_search_modal():
        st.caption("💡 登録されている伝票を受注番号ごとに一覧表示しています。行を選択して「編集画面に開く」を押すと入力フォームに呼び出せます。")
        kw_modal = st.text_input("🔍 受注番号・得意先名・商品コードで検索", placeholder="例: 00002657 / 丸藤 / ABO-08PK")

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
            合計数量=("数量", "sum"),
            備考=("備考", "first")
        ).sort_values(by=["受注日", "受注番号"], ascending=[False, False]).reset_index(drop=True)

        ev = st.dataframe(
            modal_summary[["受注日", "受注番号", "得意先名", "担当者名", "明細数", "合計数量", "納入期日", "備考"]],
            hide_index=True,
            use_container_width=True,
            height=240,
            selection_mode="single-row",
            on_select="rerun"
        )

        s_rows = ev.get("selection", {}).get("rows", [])
        target_idx = s_rows[0] if s_rows else 0
        target_ono = modal_summary.iloc[target_idx]["受注番号"]
        target_lines = df_yayoi[df_yayoi["受注番号"].astype(str) == str(target_ono)]

        st.dataframe(
            target_lines[["商品コード", "倉庫名", "数量", "納入期日", "備考"]],
            hide_index=True,
            use_container_width=True,
            height=160
        )

        if st.button("📂 この伝票を入力フォームに開く（内容確認・修正・削除）", type="primary", use_container_width=True):
            first_r = target_lines.iloc[0]
            grid_rows = []
            for _, lr in target_lines.iterrows():
                c_code = clean_str(lr["商品コード"])
                if c_code and c_code not in item_code_list:
                    item_code_list.append(c_code)
                p_name = item_name_map.get(normalize_code(c_code), clean_str(lr.get("商品名/摘要", "")))
                u_wh = item_wh_map.get(normalize_code(c_code), clean_str(lr.get("倉庫名", "")))
                q_v = int(lr["数量"])
                nt_v = clean_str(lr.get("備考", ""))
                grid_rows.append({
                    "商品コード": c_code,
                    "商品名": p_name,
                    "倉庫名": u_wh,
                    "受注残数(数量)": q_v,
                    "備考": nt_v,
                    "発注判定": evaluate_order_line(c_code, q_v)
                })
            while len(grid_rows) < 10:
                grid_rows.append({"商品コード": "", "商品名": "", "倉庫名": "", "受注残数(数量)": 0, "備考": "", "発注判定": ""})

            matched_comp = next((c for c in comp_select_list if c and c.startswith(clean_str(first_r["得意先名"]))), clean_str(first_r["得意先名"]))
            od_ts, _ = parse_jp_date(first_r["受注日"])
            od_dt = od_ts.date() if pd.notna(od_ts) else datetime.date.today()
            dd_ts, _ = parse_jp_date(first_r["納入期日"])
            dd_dt = dd_ts.date() if pd.notna(dd_ts) else datetime.date.today()

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

    top_b1, top_b2, top_b3 = st.columns([1.8, 2.4, 2.0])
    if top_b1.button("📄 新規作成（白紙リセット）", use_container_width=True):
        st.session_state["slip_editor_state"] = create_blank_slip(next_order_no_str)
        st.rerun()

    if top_b2.button("🔍 登録済み伝票を検索・呼び出し", type="primary", use_container_width=True):
        open_yayoi_search_modal()

    if top_b3.button("➕ 入力行を10行追加", use_container_width=True):
        cur_df = slip_state["grid_df"]
        if len(cur_df) < 100:
            add_cnt = min(10, 100 - len(cur_df))
            extra_df = pd.DataFrame([
                {"商品コード": "", "商品名": "", "倉庫名": "", "受注残数(数量)": 0, "備考": "", "発注判定": ""}
                for _ in range(add_cnt)
            ])
            slip_state["grid_df"] = pd.concat([cur_df, extra_df], ignore_index=True)
            st.rerun()

    h1, h2, h3, h4 = st.columns([1.4, 1.4, 1.4, 1.8])
    in_ono = h1.text_input("受注番号（自動採番）", value=slip_state["order_no"], key=f"ono_{v_key}")
    in_odate = h2.date_input("受注日", value=slip_state["order_date"], key=f"odate_{v_key}")
    in_ddate = h3.date_input("納入期日", value=slip_state["delivery_date"], key=f"ddate_{v_key}")

    staff_list = list(STAFF_OPTIONS_CLEAN)
    if slip_state["staff"] not in staff_list:
        staff_list.append(slip_state["staff"])
    in_staff = h4.selectbox(
        "担当者（新規は白紙）",
        options=staff_list,
        index=staff_list.index(slip_state["staff"]),
        key=f"staff_{v_key}"
    )

    hc1, hc2 = st.columns([2.6, 2.4])
    if slip_state["comp_label"] not in comp_select_list:
        comp_select_list.insert(1, slip_state["comp_label"])
    in_comp_label = hc1.selectbox(
        f"得意先（全 {len(master_comps):,} 社対応 ／ 新規は白紙）",
        options=comp_select_list,
        index=comp_select_list.index(slip_state["comp_label"]),
        key=f"comp_{v_key}"
    )
    in_summary = hc2.text_input(
        "摘要 / 伝票全体メモ（新規は白紙）",
        value=slip_state["summary"],
        placeholder="例: コープさっぽろ 12/3w",
        key=f"sum_{v_key}"
    )

    if "発注判定" not in slip_state["grid_df"].columns:
        slip_state["grid_df"]["発注判定"] = ""

    raw_edited_grid = st.data_editor(
        slip_state["grid_df"],
        num_rows="dynamic",
        use_container_width=True,
        height=320,
        key=f"yayoi_grid_{v_key}",
        column_config={
            "商品コード": st.column_config.SelectboxColumn(
                f"商品コード (全 {len(master_items):,} 点)",
                options=item_code_list,
                width="medium"
            ),
            "商品名": st.column_config.TextColumn("商品名 (自動)", width="medium", disabled=True),
            "倉庫名": st.column_config.TextColumn("倉庫名 (自動)", width="small", disabled=True),
            "受注残数(数量)": st.column_config.NumberColumn("受注残数 (数量)", min_value=0, max_value=999999, step=1, default=0, width="small"),
            "備考": st.column_config.TextColumn("備考 / 納期メモ", width="medium"),
            "発注判定": st.column_config.TextColumn("🚀 発注要否判定 (自動)", width="large", disabled=True)
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
        cur_judge = clean_str(r_g.get("発注判定", ""))

        if c_code:
            expected_name = item_name_map.get(normalize_code(c_code), item_name_map.get(c_code, ""))
            expected_wh = item_wh_map.get(normalize_code(c_code), item_wh_map.get(c_code, ""))
            expected_judge = evaluate_order_line(c_code, qty_v)
        else:
            expected_name = ""
            expected_wh = ""
            expected_judge = ""

        if cur_name != expected_name or cur_wh != expected_wh or cur_judge != expected_judge:
            need_sync = True

        synced_rows.append({
            "商品コード": c_code,
            "商品名": expected_name,
            "倉庫名": expected_wh,
            "受注残数(数量)": qty_v,
            "備考": note_v,
            "発注判定": expected_judge
        })

    if need_sync:
        slip_state["grid_df"] = pd.DataFrame(synced_rows)
        slip_state["version"] = int(time.time() * 1000) % 1000000
        st.rerun()

    b_col1, b_col2 = st.columns([3.2, 1.0])
    if b_col1.button("💾 この伝票を登録・保存する（クリックで確定）", type="primary", use_container_width=True):
        comp_clean = in_comp_label.split(" （")[0].strip()
        staff_clean = in_staff.strip()

        if not comp_clean:
            st.error("⚠️ 「得意先」を選択してください。")
        elif not staff_clean:
            st.error("⚠️ 「担当者」を選択してください。")
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
                        "受注日": in_odate.strftime("%Y-%m-%d"),
                        "納入期日": in_ddate.strftime("%Y-%m-%d"),
                        "備考": clean_str(row_g.get("備考", "")) or clean_str(in_summary),
                        "倉庫名": u_wh,
                        "担当者名": staff_clean
                    })

            if not valid_items:
                st.error("⚠️ 商品コードが選択され、かつ「受注残数(数量)」が1以上の行がありません。")
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
                st.success(f"✅ 受注番号『{in_ono}』（得意先: {comp_clean} ／ 明細 {len(valid_items)} 行）を保存しました！")
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

with tab4:
    ch1, ch2 = st.columns(2)
    with ch1:
        st.subheader("🚢 入荷予定あり・発注済みの商品一覧")
        inc = m_df[(m_df["入荷予定数_ロジザード"] > 0) | (m_df["発注日"] != "") | (m_df["入荷予定日_ロジザード"] != "")].copy()
        st.dataframe(inc[["品番", "商品区分", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "発注日", "前回決定事項", "決定事項"]], hide_index=True, use_container_width=True, height=520)
    with ch2:
        st.subheader("📜 過去の会議シート 決定事項 全履歴")
        h_kw = st.text_input("品番・決定内容で検索", placeholder="例: 発注 / 様子見")
        hv = df_hist.sort_values(["会議日", "品番"], ascending=[False, True]) if not df_hist.empty else pd.DataFrame()
        if not hv.empty and h_kw:
            hv = hv[hv["品番"].str.contains(h_kw, case=False, na=False) | hv["決定事項"].str.contains(h_kw, case=False, na=False)]
        st.dataframe(hv, hide_index=True, use_container_width=True, height=450)

with tab5:
    st.subheader("🚀 営業用：大量注文・確保依頼の即判定シミュレーター")
    st.caption("営業担当者が大口注文や確保依頼を受けた際、下の表に『品番』と『今回の注文数』を入れる（またはExcelからコピペする）だけで、発注が必要かどうか0.1秒で一括判定します。")

    sim_init = pd.DataFrame([
        {"品番": "ABO-08PK", "今回の注文数": 5000},
        {"品番": "ASL-3302", "今回の注文数": 1000},
        {"品番": "AKN-05TOP-10", "今回の注文数": 19842},
    ])
    sim_edited = st.data_editor(
        sim_init,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "品番": st.column_config.TextColumn("【入力】品番 (Excelから複数行コピペ可)", width="medium"),
            "今回の注文数": st.column_config.NumberColumn("【入力】今回の注文数", min_value=0, step=1, width="medium")
        }
    )

    if not sim_edited.empty:
        sim_rows = []
        for _, sr in sim_edited.iterrows():
            c_raw = clean_str(sr.get("品番", ""))
            if not c_raw:
                continue
            c_n = normalize_code(c_raw)
            q_order = int(to_num(sr.get("今回の注文数", 0)))

            if c_n not in master_sim_map:
                sim_rows.append({
                    "品番": c_raw,
                    "商品名": item_name_map.get(c_n, "マスター未登録（特注/OEM枝番）"),
                    "今回の注文数": q_order,
                    "【判定結果】": "🚨 マスター未登録（発注担当へ即確認！）",
                    "現在のフリー在庫": 0,
                    "注文後の残り在庫": -q_order,
                    "発注点(安全在庫)": 0,
                    "入数": extra_meta.get("lot_map", {}).get(c_n, 1),
                    "推奨発注数(入数切上)": q_order
                })
            else:
                info = master_sim_map[c_n]
                avail = info["free"]
                after_q = avail - q_order
                margin = after_q - info["point"]
                lot = info["lot"]
                if after_q < 0:
                    rec_q = math.ceil(abs(margin) / lot) * lot
                    verdict = f"🚨 欠品・即発注！（{abs(after_q):,}個 不足）"
                elif margin < 0:
                    rec_q = math.ceil(abs(margin) / lot) * lot
                    verdict = "⚠️ 出荷可だが発注点割れ（要追加発注）"
                else:
                    rec_q = 0
                    verdict = "✅ 在庫OK（通常受注可）"

                sim_rows.append({
                    "品番": c_raw,
                    "商品名": info["name"],
                    "今回の注文数": q_order,
                    "【判定結果】": verdict,
                    "現在のフリー在庫": avail,
                    "注文後の残り在庫": after_q,
                    "発注点(安全在庫)": info["point"],
                    "入数": lot,
                    "推奨発注数(入数切上)": rec_q
                })
        if sim_rows:
            st.markdown("#### 📋 判定結果一覧")
            st.dataframe(pd.DataFrame(sim_rows), hide_index=True, use_container_width=True)

    st.divider()
    st.subheader(f"🚨 発注抜け防止アラート：弥生確保があるのに発注検討マスターに未登録の商品（{len(df_unreg)}件）")
    st.caption("『在庫確保.xlsx』や『弥生確保リスト』に受注残があるにもかかわらず、発注検討マスターに品番が登録されていないため定期発注チェックから漏れている一覧です。")
    if df_unreg.empty:
        st.success("✅ 現在、マスター未登録の確保商品は0件です。抜け漏れはありません。")
    else:
        st.dataframe(df_unreg.sort_values("差引フリー在庫"), hide_index=True, use_container_width=True)
