# ==============================================================================
# 【パート 1 / 3】初期設定・Googleドライブ & ローカル自動同期・新原価・週次売上エンジン
# ==============================================================================
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
from concurrent.futures import ThreadPoolExecutor, as_completed

st.set_page_config(page_title="発注検討表 ＆ 売上・粗利・在庫統合システム", layout="wide")

st.markdown("""
<style>
    header[data-testid="stHeader"] {
        height: 1.8rem !important;
        background: transparent !important;
    }
    .block-container {
        padding-top: 1.0rem !important;
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
# ② 在庫確保.xlsx / order_r（マスター）.xlsx のURL（空欄でもドライブ内から自動検索します）
GDRIVE_KAKUHO_URL = ""
GDRIVE_MASTER_R_URL = ""

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else "."
MANUAL_YAYOI_CSV = os.path.join(SCRIPT_DIR, "yayoi_manual_additions.csv")

# ローカルPC実行時（G:\マイドライブ\売上ダッシュボード 等）の自動探索ディレクトリ
LOCAL_SEARCH_DIRS = [
    SCRIPT_DIR,
    ".",
    r"G:\マイドライブ\売上ダッシュボード",
    r"G:\マイドライブ\売上ダッシュボード\済",
    r"G:\マイドライブ\発注検討",
    "/content/drive/MyDrive/売上ダッシュボード",
    "/content/drive/MyDrive/売上ダッシュボード/済",
    "/content/drive/MyDrive/発注検討"
]

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
        s = str(val).replace(",", "").replace("，", "").replace("%", "").replace("円", "").replace("¥", "").replace("(", "-").replace(")", "").strip()
        return float(s)
    except Exception:
        return 0.0

def parse_jp_date(val, fallback_year=None):
    if is_na_val(val):
        return pd.NaT, ""
    if isinstance(val, (datetime.datetime, datetime.date, pd.Timestamp)):
        ts = pd.Timestamp(val)
        return ts, ts.strftime("%Y-%m-%d")
    s = clean_str(val)
    if not s:
        return pd.NaT, ""
    s_digits = s.replace("-", "").replace("/", "")
    if len(s_digits) == 8 and s_digits.isdigit():
        dt = pd.to_datetime(s_digits, format="%Y%m%d", errors="coerce")
        if pd.notna(dt):
            return dt, dt.strftime("%Y-%m-%d")
    m_reiwa = re.match(r'^(?:令和|R|r)\s*(\d{1,2})[年./\-](\d{1,2})[月./\-](\d{1,2})', s)
    if m_reiwa:
        y = 2018 + int(m_reiwa.group(1))
        dt = pd.to_datetime(f"{y}-{int(m_reiwa.group(2)):02d}-{int(m_reiwa.group(3)):02d}", errors="coerce")
        if pd.notna(dt):
            return dt, dt.strftime("%Y-%m-%d")
    s_norm = re.sub(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?', r'\1-\2-\3', s)
    s_norm = s_norm.replace("/", "-").replace(".", "-").split(" ")[0]
    m_short_y = re.match(r'^(\d{2})-(\d{1,2})-(\d{1,2})$', s_norm)
    if m_short_y:
        s_norm = f"20{m_short_y.group(1)}-{int(m_short_y.group(2)):02d}-{int(m_short_y.group(3)):02d}"
    m_md = re.match(r'^(\d{1,2})[月\-/](\d{1,2})日?$', s)
    if m_md and fallback_year:
        s_norm = f"{fallback_year}-{int(m_md.group(1)):02d}-{int(m_md.group(2)):02d}"

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

def find_local_file(keywords, exts=(".xlsx", ".csv")):
    if isinstance(exts, str):
        exts = (exts,)
    for d in LOCAL_SEARCH_DIRS:
        if os.path.exists(d):
            try:
                for fn in os.listdir(d):
                    if fn.lower().endswith(exts) and not fn.startswith("~$"):
                        if any(k.lower() in fn.lower() for k in keywords):
                            return os.path.join(d, fn)
            except Exception:
                pass
    return None

def fix_corrupted_xlsx_filters(xlsx_bytes: bytes) -> io.BytesIO:
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
# Googleドライブ接続 ＆ 全ファイル自動取得処理
# =========================================================
def get_gdrive_creds_info():
    try:
        if "gcp_service_account" in st.secrets:
            return dict(st.secrets["gcp_service_account"])
        elif "google_service_account" in st.secrets:
            return dict(st.secrets["google_service_account"])
        elif "private_key" in st.secrets and "client_email" in st.secrets:
            return dict(st.secrets)
        else:
            for k in st.secrets:
                val = st.secrets[k]
                if hasattr(val, "keys") and "private_key" in val:
                    return dict(val)
                elif isinstance(val, str) and "private_key" in val:
                    try:
                        return json.loads(val)
                    except Exception:
                        pass
    except Exception:
        pass
    return None

def get_gdrive_service():
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        creds_info = get_gdrive_creds_info()
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
            if len(fh.getvalue()) > 0:
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
                if len(fh.getvalue()) > 0:
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
                if len(data) > 500 and data[:2] == b'PK':
                    return data, "Googleドライブ共有リンク"
        except Exception:
            continue
    return None, None

@st.cache_data(ttl=3600, show_spinner="☁️ Googleドライブから発注検討表・在庫確保・order_r（新原価）・ロジザード在庫を同期中...")
def sync_all_files_from_gdrive(order_url: str, kakuho_url: str = "", master_r_url: str = ""):
    m_ord = re.search(r'/d/([a-zA-Z0-9_-]+)', order_url or "")
    target_ord_id = m_ord.group(1) if m_ord else ""
    m_kak = re.search(r'/d/([a-zA-Z0-9_-]+)', kakuho_url or "")
    target_kak_id = m_kak.group(1) if m_kak else ""
    m_mr = re.search(r'/d/([a-zA-Z0-9_-]+)', master_r_url or "")
    target_mr_id = m_mr.group(1) if m_mr else ""

    service = get_gdrive_service()

    result = {
        "main_bytes": None,
        "yayoi_bytes": None,
        "yayoi_filename": "",
        "order_r_bytes": None,
        "order_r_filename": "",
        "roji_csv_list": [],
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
            result["yayoi_filename"] = "GDrive(URL指定: 在庫確保)"

    if target_mr_id:
        b_mr, _ = download_file_by_id(service, target_mr_id)
        if b_mr:
            result["order_r_bytes"] = b_mr
            result["order_r_filename"] = "GDrive(URL指定: order_r)"

    if service is not None:
        try:
            q = (
                "trashed = false and ("
                "name contains '在庫確保' or name contains '弥生確保' or name contains '確保リスト' or "
                "name contains 'order_r' or name contains '新原価' or name contains 'ロジザード' or "
                "name contains '発注検討表' or name contains '弥生会社名' or name contains '弥生商品番号'"
                ")"
            )
            res = service.files().list(
                q=q,
                pageSize=80,
                fields="files(id, name, modifiedTime)",
                orderBy="modifiedTime desc",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            files_list = res.get("files", [])

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
                elif result["order_r_bytes"] is None and ("order_r" in fname.lower() or "新原価" in fname):
                    b, _ = download_file_by_id(service, fid)
                    if b:
                        result["order_r_bytes"] = b
                        result["order_r_filename"] = f"GDrive: {fname}"
                elif "ロジザード" in fname and fname.lower().endswith(".csv") and len(result["roji_csv_list"]) < 4:
                    b, _ = download_file_by_id(service, fid)
                    if b:
                        result["roji_csv_list"].append((fname, b))

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

# =========================================================
# 『order_r（マスター）.xlsx』の「新原価」「item」シート読み込み処理
# =========================================================
@st.cache_data
def load_order_r_cost_master(file_source):
    """
    order_r（マスター）.xlsx の『新原価』シートから最新の原価・最低下代を取得し、
    『加盟店価格』『加盟店グループ管理』『注文者会社名』から会社ごとの販売価格（ロット単価）を取得する関数
    """
    raw_bytes = None
    used_label = ""
    sources = [
        (file_source, "Googleドライブ/アップロード: order_r"),
        (find_local_file(["order_r", "新原価"], ".xlsx"), "ローカル: order_r（マスター）.xlsx")
    ]
    for src, lbl in sources:
        if src is not None:
            if isinstance(src, bytes):
                raw_bytes = src
                used_label = lbl
                break
            elif hasattr(src, "getvalue"):
                raw_bytes = src.getvalue()
                used_label = lbl
                break
            elif isinstance(src, str) and os.path.exists(src):
                with open(src, "rb") as f:
                    raw_bytes = f.read()
                used_label = f"ローカル: {os.path.basename(src)}"
                break

    cost_map = {}
    gedai_map = {}
    wh_type_map = {}
    std_price_map = {}
    r_name_map = {}
    kamei_price_map = {}

    if raw_bytes is None:
        return cost_map, gedai_map, wh_type_map, std_price_map, r_name_map, kamei_price_map, "未検出"

    clean_io = fix_corrupted_xlsx_filters(raw_bytes)
    try:
        xls = pd.ExcelFile(clean_io)
        # 1. 『新原価』シートの読み取り（原価・最低下代の最優先マスター）
        if "新原価" in xls.sheet_names:
            df_sg = pd.read_excel(xls, sheet_name="新原価")
            for _, r in df_sg.iterrows():
                vals = r.values
                if len(vals) >= 2 and clean_str(vals[0]):
                    c_n = normalize_code(vals[0])
                    c_val = to_num(vals[1])
                    if c_val > 0:
                        cost_map[c_n] = c_val
                    if len(vals) >= 3 and clean_str(vals[2]):
                        wh_type_map[c_n] = clean_str(vals[2])
                    if len(vals) >= 4:
                        g_val = to_num(vals[3])
                        if 0 < g_val < 90000:
                            gedai_map[c_n] = g_val

        # 2. 『item』シートの読み取り
        if "item" in xls.sheet_names:
            df_it = pd.read_excel(xls, sheet_name="item")
            for _, r in df_it.iterrows():
                c_raw = clean_str(r.get("商品番号", ""))
                if not c_raw:
                    continue
                c_n = normalize_code(c_raw)
                p_val = to_num(r.get("販売価格", 0))
                if 0 < p_val < 90000:
                    std_price_map[c_n] = p_val
                c_val = to_num(r.get("原価", 0))
                if c_n not in cost_map and c_val > 0:
                    cost_map[c_n] = c_val
                nm = clean_str(r.get("商品名", ""))
                if nm:
                    nm_clean = re.sub(r'^.*?】', '', nm)
                    nm_clean = re.sub(r'[\(（]\d{13}[\)）]$', '', nm_clean).strip()
                    r_name_map[c_n] = nm_clean or nm

        # 3. グループ番号・加盟店コードから会社名を引く辞書を作成
        grp_to_comp = {"12": "【標準ランク】1.1倍", "13": "【標準ランク】1.15倍", "14": "【標準ランク】1.2倍", "15": "【標準ランク】1.25倍", "40": "【標準ランク】定価/上代"}
        if "加盟店グループ管理" in xls.sheet_names:
            df_gm = pd.read_excel(xls, sheet_name="加盟店グループ管理")
            for _, r in df_gm.iterrows():
                cnm = clean_str(r.get("加盟店名", ""))
                if cnm:
                    g_no = clean_str(r.get("グループ番号", "")).replace("_x000D_", "").strip()
                    k_cd = clean_str(r.get("加盟店コード", "")).strip()
                    if g_no:
                        grp_to_comp[g_no] = cnm
                    if k_cd:
                        grp_to_comp[k_cd] = cnm

        if "注文者会社名" in xls.sheet_names:
            df_cm = pd.read_excel(xls, sheet_name="注文者会社名")
            for _, r in df_cm.iterrows():
                cnm = clean_str(r.get("加盟店名", ""))
                if cnm:
                    g_no = clean_str(r.get("グループ番号", "")).strip()
                    k_cd = clean_str(r.get("加盟店コード", "")).strip()
                    if g_no and g_no not in grp_to_comp:
                        grp_to_comp[g_no] = cnm
                    if k_cd and k_cd not in grp_to_comp:
                        grp_to_comp[k_cd] = cnm

        # 4. 『加盟店価格』シートから商品別・会社別の販売価格（ロット単価）を一括取得
        if "加盟店価格" in xls.sheet_names:
            df_kp = pd.read_excel(xls, sheet_name="加盟店価格")
            for _, r in df_kp.iterrows():
                p_raw = clean_str(r.get("商品番号", "")) or clean_str(r.get("商品管理番号", ""))
                if not p_raw:
                    continue
                c_n = normalize_code(p_raw)
                u_price = int(round(to_num(r.get("ロット単価", 0))))
                if u_price <= 0 or u_price >= 90000:
                    continue
                lot_q = max(int(round(to_num(r.get("ロット数", 1)))), 1)
                comp_nm = clean_str(r.get("会社名", ""))
                g_seq = clean_str(r.get("グループ連番", ""))
                if not comp_nm and g_seq:
                    comp_nm = grp_to_comp.get(g_seq, f"グループ連番:{g_seq}")
                if not comp_nm:
                    continue

                # 最低下代が新原価シートになければ加盟店価格シートの最低下代で補完
                if c_n not in gedai_map:
                    g_val = to_num(r.get("最低下代", 0))
                    if 0 < g_val < 90000:
                        gedai_map[c_n] = g_val

                if c_n not in kamei_price_map:
                    kamei_price_map[c_n] = []
                kamei_price_map[c_n].append({
                    "会社名": comp_nm,
                    "グループ": g_seq,
                    "ロット数": lot_q,
                    "販売価格": u_price
                })
    except Exception:
        pass

    return cost_map, gedai_map, wh_type_map, std_price_map, r_name_map, kamei_price_map, f"{used_label} (新原価 {len(cost_map):,}件 / 加盟店価格 {len(kamei_price_map):,}品番)"

# =========================================================
# ロジザード在庫CSV（名古屋・青山）の直接読み込み補完処理
# =========================================================
def parse_roji_csv_files(roji_csv_list):
    """Googleドライブまたはローカルに『ロジザード在庫.csv』がある場合に読み取る関数"""
    free_map = {}
    sec_map = {}
    all_files = list(roji_csv_list or [])
    if not all_files:
        for d in LOCAL_SEARCH_DIRS:
            if os.path.exists(d):
                try:
                    for fn in os.listdir(d):
                        if "ロジザード" in fn and fn.lower().endswith(".csv"):
                            with open(os.path.join(d, fn), "rb") as f:
                                all_files.append((fn, f.read()))
                except Exception:
                    pass

    for fname, raw_b in all_files:
        for enc in ["cp932", "shift_jis", "utf-8-sig", "utf-8"]:
            try:
                df_r = pd.read_csv(io.BytesIO(raw_b), encoding=enc, low_memory=False)
                if not df_r.empty and len(df_r.columns) >= 2:
                    cols_s = [clean_str(c) for c in df_r.columns]
                    c_code = next((orig for orig, cs in zip(df_r.columns, cols_s) if any(k in cs for k in ["商品コード", "品番", "商品型番", "型番"])), None)
                    c_free = next((orig for orig, cs in zip(df_r.columns, cols_s) if any(k in cs for k in ["フリー", "有効", "可能", "在庫数"])), None)
                    c_sec = next((orig for orig, cs in zip(df_r.columns, cols_s) if any(k in cs for k in ["確保", "引当", "保留"])), None)
                    if c_code is not None and c_free is not None:
                        for _, rr in df_r.iterrows():
                            cn = normalize_code(rr[c_code])
                            if cn:
                                free_map[cn] = free_map.get(cn, 0) + int(to_num(rr[c_free]))
                                if c_sec is not None:
                                    sec_map[cn] = sec_map.get(cn, 0) + int(to_num(rr[c_sec]))
                    break
            except Exception:
                continue
    return free_map, sec_map

# ==============================================================================
# 【パート 2 / 3】yayoi_CSV 週次売上(先週vs今週)・新原価による粗利計算・バッジ判定エンジン
# ==============================================================================
def infer_date_from_filename_and_mtime(fname: str, mtime_str: str = ""):
    ref_year = 2026
    if mtime_str and len(mtime_str) >= 4 and mtime_str[:4].isdigit():
        ref_year = int(mtime_str[:4])

    # 親フォルダ名やパスに「2025年」「2024年」などが含まれる場合
    m_yr = re.search(r'(202[2-9])年?', fname)
    if m_yr:
        ref_year = int(m_yr.group(1))

    m = re.search(r'yayoi_(\d{2})(\d{2})', fname.lower())
    if m:
        mm, dd = int(m.group(1)), int(m.group(2))
        if 1 <= mm <= 12 and 1 <= dd <= 31:
            if not m_yr and mtime_str and len(mtime_str) >= 7:
                m_year = int(mtime_str[:4])
                m_month = int(mtime_str[5:7])
                if mm > m_month + 1:
                    ref_year = m_year - 1
            dt = pd.to_datetime(f"{ref_year}-{mm:02d}-{dd:02d}", errors="coerce")
            if pd.notna(dt):
                return dt, dt.strftime("%Y-%m-%d"), ref_year
    return pd.NaT, "", ref_year

def parse_yayoi_sales_csv(fname: str, raw_bytes: bytes, mtime_str: str = "") -> pd.DataFrame:
    if not raw_bytes:
        return pd.DataFrame()

    fn_dt, fn_date_fmt, fn_year = infer_date_from_filename_and_mtime(fname, mtime_str)

    df = pd.DataFrame()
    for enc in ["cp932", "shift_jis", "utf-8-sig", "utf-8"]:
        try:
            df = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, low_memory=False)
            if not df.empty and len(df.columns) >= 3:
                break
        except Exception:
            continue

    if df.empty:
        return pd.DataFrame()

    cols_str = [clean_str(c) for c in df.columns]

    def f_col(cands):
        for cand in cands:
            for orig_c, clean_c in zip(df.columns, cols_str):
                if cand == clean_c or cand in clean_c:
                    return orig_c
        return None

    c_code = f_col(["商品コード", "品番", "型番", "商品番号", "code"])
    c_cust = f_col(["得意先名称", "得意先名", "会社名", "販売先", "顧客名", "得意先略称", "cust"])
    c_qty = f_col(["数量", "売上数", "出荷数", "個数", "バラ数", "qty"])
    c_price = f_col(["単価", "売上単価", "納入単価", "税抜単価", "price"])
    c_amt = f_col(["金額", "売上金額", "税抜金額", "明細金額", "amount"])
    c_dt = f_col(["売上日", "出荷日", "伝票日付", "日付", "受注日", "date"])
    c_stf = f_col(["担当者名", "担当者", "入力者", "staff"])

    # 見出しなし弥生CSVの場合の自動列推定
    if c_code is None or c_qty is None:
        for enc in ["cp932", "shift_jis", "utf-8-sig", "utf-8"]:
            try:
                df_nohead = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, header=None, low_memory=False)
                if not df_nohead.empty and len(df_nohead.columns) >= 5:
                    df = df_nohead
                    break
            except Exception:
                continue

        sample_row = [clean_str(x) for x in df.iloc[0].values]
        code_col_idx = None
        date_col_idx = None
        for idx_c, val in enumerate(sample_row):
            if code_col_idx is None and re.match(r'^[A-Z]{2,4}-[0-9A-Z-]+$', val.upper()):
                code_col_idx = idx_c
            if date_col_idx is None and re.match(r'^(?:202[0-9]|R\d|令和\d)[-/\s年.]?[01]?[0-9][-/\s月.]?[0-3]?[0-9]', val):
                date_col_idx = idx_c

        if code_col_idx is not None:
            c_code = df.columns[code_col_idx]
            c_dt = df.columns[date_col_idx] if date_col_idx is not None else (df.columns[1] if len(df.columns) > 1 else None)
            c_cust = df.columns[5] if len(df.columns) > 5 else df.columns[0]
            c_qty = df.columns[min(code_col_idx + 4, len(df.columns) - 1)]
            for offset_q in [4, 5, 3, 2, 6, 7, 8, 1]:
                if code_col_idx + offset_q < len(df.columns):
                    v_test = clean_str(df.iloc[0, code_col_idx + offset_q]).replace(",", "")
                    if v_test.replace("-", "").isdigit():
                        c_qty = df.columns[code_col_idx + offset_q]
                        break
            c_price = None
            c_amt = None
            c_stf = None

    if c_code is None or c_qty is None:
        return pd.DataFrame()

    out = pd.DataFrame()
    out["code_norm"] = df[c_code].apply(normalize_code)
    out["cust"] = df[c_cust].apply(lambda x: clean_str(x) if clean_str(x) else "（得意先名なし）") if c_cust is not None else "（得意先名なし）"
    out["qty"] = df[c_qty].apply(to_num)
    out["price"] = df[c_price].apply(to_num) if c_price is not None else 0.0
    raw_amt = df[c_amt].apply(to_num) if c_amt is not None else pd.Series([0.0] * len(df))
    calc_amt = out["qty"] * out["price"]
    # 金額列が0または「単価と同じ数字（数量が掛けられていない）」場合は、数量×単価を優先採用
    out["amount"] = [
        ca if (ra <= 0 or (q > 1 and abs(ra - p) < 1 and ca > ra)) else ra
        for ra, ca, q, p in zip(raw_amt, calc_amt, out["qty"], out["price"])
    ]
    out["staff"] = df[c_stf].apply(clean_str) if c_stf is not None else ""

    if c_dt is not None:
        parsed_d = df[c_dt].apply(lambda v: parse_jp_date(v, fallback_year=fn_year))
        out["dt"] = [p[0] if pd.notna(p[0]) else fn_dt for p in parsed_d]
        out["date_fmt"] = [p[1] if p[1] else fn_date_fmt for p in parsed_d]
    else:
        out["dt"] = fn_dt
        out["date_fmt"] = fn_date_fmt

    out = out[(out["code_norm"] != "") & (~out["code_norm"].isin(["商品コード", "品番", "型番"])) & (out["qty"] != 0)].copy()

    def calc_ki(ts):
        if pd.isna(ts):
            return "今期(26/7〜)"
        fy = ts.year if ts.month >= 7 else ts.year - 1
        if fy >= 2026:
            return "今期(26/7〜)"
        elif fy == 2025:
            return "77期"
        elif fy == 2024:
            return "76期"
        elif fy == 2023:
            return "75期"
        elif fy == 2022:
            return "74期"
        return f"{fy}年度"

    out["ki"] = out["dt"].apply(calc_ki)
    return out

@st.cache_data(ttl=3600, show_spinner="🔥 売上データ(yayoi_CSV)から『先週vs今週の売れ数』『購入先ランキング』『売上・粗利』を計算中...")
def fetch_and_build_yayoi_sales_from_gdrive(up_sales_bytes, up_sales_name, target_codes_tuple):
    target_set = set(target_codes_tuple)
    dfs = []
    local_csv_count = 0

    if up_sales_bytes is not None and up_sales_name:
        df_u = parse_yayoi_sales_csv(up_sales_name, up_sales_bytes)
        if not df_u.empty:
            dfs.append(df_u)

    # 1. ローカルに G:\マイドライブ\売上ダッシュボード 等がある場合は直接高速スキャン
    for base_d in LOCAL_SEARCH_DIRS:
        if os.path.exists(base_d) and ("売上ダッシュボード" in base_d or "発注検討" in base_d):
            try:
                for root_d, _, files in os.walk(base_d):
                    for fn in files:
                        if fn.lower().endswith(".csv") and not fn.startswith("~$") and ("yayoi_" in fn.lower() or "売上" in fn):
                            full_p = os.path.join(root_d, fn)
                            mtime_iso = datetime.datetime.fromtimestamp(os.path.getmtime(full_p)).strftime("%Y-%m-%d")
                            with open(full_p, "rb") as f:
                                df_l = parse_yayoi_sales_csv(full_p, f.read(), mtime_iso)
                                if not df_l.empty:
                                    dfs.append(df_l[df_l["code_norm"].isin(target_set)])
                                    local_csv_count += 1
            except Exception:
                pass

    # 2. クラウド実行時（ローカルで見つからない場合）はGoogle Drive APIから並列ダウンロード
    creds_info = get_gdrive_creds_info()
    csv_file_metas = []

    if local_csv_count == 0 and creds_info is not None:
        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
            creds = service_account.Credentials.from_service_account_info(
                creds_info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
            )
            service = build("drive", "v3", credentials=creds)

            page_token = None
            q = "trashed = false and (name contains 'yayoi_' or name contains '弥生' or name contains '売上') and (mimeType contains 'csv' or name contains '.csv')"
            while True:
                res = service.files().list(
                    q=q,
                    pageSize=250,
                    pageToken=page_token,
                    fields="nextPageToken, files(id, name, modifiedTime, size)",
                    orderBy="modifiedTime desc",
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True
                ).execute()
                batch_files = res.get("files", [])
                for f in batch_files:
                    fn = f.get("name", "")
                    mt = f.get("modifiedTime", "")
                    if fn.lower().endswith(".csv") and not fn.startswith("~$"):
                        csv_file_metas.append((f["id"], fn, mt))
                page_token = res.get("nextPageToken")
                if not page_token or len(csv_file_metas) >= 800:
                    break
        except Exception:
            pass

    if csv_file_metas and creds_info is not None:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build

        def dl_and_parse_worker(meta_item):
            fid, fname, mtime_str = meta_item
            try:
                c_obj = service_account.Credentials.from_service_account_info(
                    creds_info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
                )
                srv = build("drive", "v3", credentials=c_obj, cache_discovery=False)
                raw_b = srv.files().get_media(fileId=fid, supportsAllDrives=True).execute()
                if raw_b:
                    df_p = parse_yayoi_sales_csv(fname, raw_b, mtime_str)
                    if not df_p.empty:
                        return df_p[df_p["code_norm"].isin(target_set)]
            except Exception:
                pass
            return pd.DataFrame()

        with ThreadPoolExecutor(max_workers=12) as executor:
            futures = [executor.submit(dl_and_parse_worker, m_item) for m_item in csv_file_metas]
            for fut in as_completed(futures):
                df_res = fut.result()
                if not df_res.empty:
                    dfs.append(df_res)

    total_files_cnt = local_csv_count + len(csv_file_metas)
    if not dfs:
        return {}, {}, {}, "", 0

    df_all = pd.concat(dfs, ignore_index=True)
    df_all = df_all[df_all["code_norm"].isin(target_set)].copy()
    if df_all.empty:
        return {}, {}, {}, "", total_files_cnt

    # ★ 最新売上日を基準にして「今週(直近7日間)」と「先週(その前の7日間)」を自動判定
    valid_dts = df_all["dt"].dropna()
    max_sales_dt = valid_dts.max() if not valid_dts.empty else pd.Timestamp.now()
    this_w_start = max_sales_dt - pd.Timedelta(days=6)
    prev_w_end = this_w_start - pd.Timedelta(days=1)
    prev_w_start = prev_w_end - pd.Timedelta(days=6)

    this_w_label = f"今週 ({this_w_start.strftime('%m/%d')}〜{max_sales_dt.strftime('%m/%d')})"
    prev_w_label = f"先週 ({prev_w_start.strftime('%m/%d')}〜{prev_w_end.strftime('%m/%d')})"

    mask_this_w = (df_all["dt"] >= this_w_start) & (df_all["dt"] <= max_sales_dt)
    mask_prev_w = (df_all["dt"] >= prev_w_start) & (df_all["dt"] <= prev_w_end)
    mask_konki = (df_all["ki"] == "今期(26/7〜)")

    df_this_w = df_all[mask_this_w]
    df_prev_w = df_all[mask_prev_w]
    df_konki = df_all[mask_konki]

    this_w_qty_map = df_this_w.groupby("code_norm")["qty"].sum().round().astype(int).to_dict() if not df_this_w.empty else {}
    prev_w_qty_map = df_prev_w.groupby("code_norm")["qty"].sum().round().astype(int).to_dict() if not df_prev_w.empty else {}
    konki_qty_map = df_konki.groupby("code_norm")["qty"].sum().round().astype(int).to_dict() if not df_konki.empty else {}
    konki_amt_map = df_konki.groupby("code_norm")["amount"].sum().round().astype(int).to_dict() if not df_konki.empty else {}

    def make_rank_list(sub_df):
        if sub_df.empty:
            return []
        c_grp = (
            sub_df.groupby("cust", as_index=False)
            .agg(qty=("qty", "sum"), amount=("amount", "sum"), cnt=("cust", "count"), last_date=("date_fmt", "max"))
            .sort_values("qty", ascending=False)
        )
        tot = c_grp["qty"].sum()
        return [
            {
                "得意先名": r["cust"],
                "数量": int(round(r["qty"])),
                "構成比": f"{round(r['qty'] / tot * 100, 1) if tot > 0 else 0}%",
                "金額": int(round(r["amount"])),
                "伝票回数": int(r["cnt"]),
                "最終売上日": str(r["last_date"] or "—")
            }
            for _, r in c_grp.head(50).iterrows()
        ]

    comp_map = {}
    det_map = {}

    for c_norm, grp_code in df_all.groupby("code_norm", sort=False):
        comp_map[c_norm] = {}
        comp_map[c_norm]["ALL"] = make_rank_list(grp_code)
        comp_map[c_norm]["THIS_WEEK"] = make_rank_list(grp_code[(grp_code["dt"] >= this_w_start) & (grp_code["dt"] <= max_sales_dt)])
        comp_map[c_norm]["PREV_WEEK"] = make_rank_list(grp_code[(grp_code["dt"] >= prev_w_start) & (grp_code["dt"] <= prev_w_end)])

        for ki_label, grp_ki in grp_code.groupby("ki", sort=False):
            comp_map[c_norm][ki_label] = make_rank_list(grp_ki)

        grp_sorted = grp_code.sort_values("date_fmt", ascending=False)
        det_map[c_norm] = [
            {
                "期": r["ki"],
                "日付": str(r["date_fmt"] or "—"),
                "得意先名": r["cust"],
                "数量": int(round(r["qty"])),
                "金額": int(round(r["amount"])),
                "担当者": str(r["staff"] or "")
            }
            for _, r in grp_sorted.head(60).iterrows()
        ]

    stats_pack = {
        "this_w_qty": this_w_qty_map,
        "prev_w_qty": prev_w_qty_map,
        "konki_qty": konki_qty_map,
        "konki_amt": konki_amt_map,
        "this_w_label": this_w_label,
        "prev_w_label": prev_w_label
    }
    konki_total_sum = sum(konki_qty_map.values())
    src_info = f"yayoi_CSV {total_files_cnt}ファイル連動済 ({this_w_label} / 今期売上: {konki_total_sum:,}個)"
    return comp_map, det_map, stats_pack, src_info, total_files_cnt

@st.cache_data
def load_company_master(file_bytes_or_path):
    loaded = []
    seen = set()
    sources = [file_bytes_or_path, find_local_file(["弥生会社名", "会社名", "得意先"], ".xlsx")]
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
    sources = [file_bytes_or_path, find_local_file(["弥生商品番号", "商品番号"], ".xlsx")]
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
    if "確保貼り付け" in all_sheets:
        df_k = all_sheets["確保貼り付け"].iloc[4:]
        for _, rk in df_k.iterrows():
            vals = rk.values
            if len(vals) > 14 and clean_str(vals[13]) and clean_str(vals[13]) not in ("行ラベル", "総計", "(空白)"):
                c_n = normalize_code(vals[13])
                sheet_yayoi_sum[c_n] = sheet_yayoi_sum.get(c_n, 0) + int(to_num(vals[14]))

    # ★ 発注検討表-修正.xlsx 内の total シートからは金額・単価・原価を一切引っ張らない（異常値混入防止）
    latest_ship_map = {}
    tot_cost_fallback = {}
    tot_avg_price_map = {}
    tot_sales_amt_map = {}

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
        y_sec_init = sheet_yayoi_sum.get(c_norm, cell_yayoi)

        free_val = (stock_roji + sec_roji + in_roji) - y_sec_init
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
            "今期出荷": 0,
            "先週売数": 0,
            "今週売数": 0,
            "新原価": 0,
            "最低下代": 0,
            "平均単価": round(tot_avg_price_map.get(c_norm, 0), 0),
            "売上高": tot_sales_amt_map.get(c_norm, 0),
            "粗利額": 0,
            "粗利率": 0.0,
            "在庫金額": 0,
            "ABCランク": "C",
            "新商品": False,
            "急増": False,
            "滞留": False,
            "バッジ表示": "",
            "1課発注点": p1_val,
            "2課発注点": p2_val,
            "在庫数_ロジザード": stock_roji,
            "入荷予定数_ロジザード": in_roji,
            "入荷予定日_ロジザード": ndate_val,
            "確保数_ロジザード": sec_roji,
            "在庫確保_弥生": y_sec_init,
            "確保数のズレ": sec_roji - y_sec_init,
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
        "tot_cost_fallback": tot_cost_fallback,
        "tot_avg_price_map": tot_avg_price_map,
        "tot_sales_amt_map": tot_sales_amt_map
    }
    return df_main, df_hist, prev_col_title, extra_meta, target_sheet_name, all_sheets

@st.cache_data
def load_yayoi_data(file_bytes_or_path, all_sheets_from_main=None):
    df = pd.DataFrame()
    used_src_name = ""
    sources = [
        (file_bytes_or_path, "Googleドライブ/アップロード (在庫確保)"),
        (find_local_file(["在庫確保"], ".xlsx"), "ローカル: 在庫確保.xlsx"),
        (find_local_file(["弥生確保リスト", "確保リスト"], ".xlsx"), "ローカル: 弥生確保リスト.xlsx")
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

    if df.empty and all_sheets_from_main and "確保貼り付け" in all_sheets_from_main:
        supp_rows = []
        df_k = all_sheets_from_main["確保貼り付け"].iloc[4:]
        for idx_k, rk in df_k.iterrows():
            vals = rk.values
            if len(vals) > 14 and clean_str(vals[13]) and clean_str(vals[13]) not in ("行ラベル", "総計", "(空白)"):
                c_raw2 = clean_str(vals[13])
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
            df = pd.DataFrame(supp_rows)
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
        df = df.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
    return df, used_src_name

# ---------------------------------------------------------
# ヘッダー ＆ Googleドライブ同期ボタン ＆ アップローダー
# ---------------------------------------------------------
head_col1, head_col2, head_col3 = st.columns([2.7, 1.1, 1.8])
with head_col1:
    st.markdown('<div class="app-header">📊 発注検討表 ＆ 売上・粗利・週次トレンド統合システム</div>', unsafe_allow_html=True)
with head_col2:
    if st.button("🔄 Googleドライブから最新データを同期", type="primary", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
with head_col3:
    with st.expander("📂 手動アップロード / GDrive設定", expanded=False):
        gdrive_url_input = st.text_input("☁️ ①発注検討表 GoogleドライブURL", value=GDRIVE_ORDER_URL)
        gdrive_kakuho_input = st.text_input("☁️ ②在庫確保.xlsx GoogleドライブURL (任意)", value=GDRIVE_KAKUHO_URL, placeholder="空欄でも自動検索します")
        gdrive_mr_input = st.text_input("☁️ ③order_r(マスター).xlsx URL (任意)", value=GDRIVE_MASTER_R_URL, placeholder="空欄でも自動検索します")
        u_c1, u_c2 = st.columns(2)
        up_main = u_c1.file_uploader("📋 発注検討表.xlsx", type=["xlsx"], key="up_m")
        up_yayoi = u_c2.file_uploader("📑 在庫確保.xlsx", type=["xlsx"], key="up_y")
        u_c3, u_c4 = st.columns(2)
        up_order_r = u_c3.file_uploader("💰 order_r（マスター）.xlsx", type=["xlsx"], key="up_or")
        up_sales = u_c4.file_uploader("🔥 売上yayoi_CSV (任意)", type=["csv", "xlsx"], key="up_sales")

gdrive_sync = sync_all_files_from_gdrive(gdrive_url_input, gdrive_kakuho_input, gdrive_mr_input)

if gdrive_sync["error_detail"] and up_main is None:
    st.warning(f"⚠️ **Googleドライブ未接続**\n{gdrive_sync['error_detail']}")

main_src = up_main if up_main is not None else (gdrive_sync["main_bytes"] if gdrive_sync["main_bytes"] is not None else find_local_file(["発注検討表"], ".xlsx"))
yayoi_src = up_yayoi if up_yayoi is not None else gdrive_sync["yayoi_bytes"]
order_r_src = up_order_r if up_order_r is not None else gdrive_sync["order_r_bytes"]
comp_src = gdrive_sync["comp_bytes"]
item_src = gdrive_sync["item_bytes"]

# 1. 発注検討表マスター・在庫確保・order_r（新原価）の読み込み
df_main, df_hist, prev_col_title, extra_meta, loaded_sheet_name, all_sheets_raw = load_excel_master(main_src)
df_yayoi, yayoi_src_label = load_yayoi_data(yayoi_src, all_sheets_raw)
if gdrive_sync["yayoi_filename"] and up_yayoi is None and "ローカル" not in yayoi_src_label:
    yayoi_src_label = gdrive_sync["yayoi_filename"]

cost_map, gedai_map, wh_type_map, std_price_map, r_name_map, kamei_price_map, cost_src_label = load_order_r_cost_master(order_r_src)
if gdrive_sync["order_r_filename"] and up_order_r is None and "ローカル" not in cost_src_label:
    cost_src_label = f"{gdrive_sync['order_r_filename']} (新原価 {len(cost_map):,}件)"

# 2. ロジザード在庫CSVが直接ドライブ/ローカルにある場合は最新在庫として補完
roji_csv_free, roji_csv_sec = parse_roji_csv_files(gdrive_sync.get("roji_csv_list", []))
if roji_csv_free:
    df_main["在庫数_ロジザード"] = df_main["品番_正規化"].map(roji_csv_free).fillna(df_main["在庫数_ロジザード"]).astype(int)
if roji_csv_sec:
    df_main["確保数_ロジザード"] = df_main["品番_正規化"].map(roji_csv_sec).fillna(df_main["確保数_ロジザード"]).astype(int)

master_comps = load_company_master(comp_src)
master_items = load_item_master(item_src)

if df_main.empty:
    st.error("⚠️ 『発注検討表.xlsx』が読み込めませんでした。Googleドライブの共有権限を確認するか、右上のエクスパンダーからアップロードしてください。")
    st.stop()

yayoi_dict = {}
yayoi_comp_search_map = {}
yayoi_cust_rank_map = {}

if not df_yayoi.empty:
    df_yayoi["品番_正規化"] = df_yayoi["商品コード"].apply(normalize_code)
    yayoi_all_totals = df_yayoi.groupby("品番_正規化")["数量"].sum().to_dict()

    df_main["在庫確保_弥生"] = df_main["品番_正規化"].map(yayoi_all_totals).fillna(0).astype(int)
    df_main["確保数のズレ"] = df_main["確保数_ロジザード"] - df_main["在庫確保_弥生"]
    df_main["入荷後のフリー在庫"] = (df_main["在庫数_ロジザード"] + df_main["確保数_ロジザード"] + df_main["入荷予定数_ロジザード"]) - df_main["在庫確保_弥生"]
    df_main["発注検討フラグ"] = df_main["入荷後のフリー在庫"] - (df_main["1課発注点"] + df_main["2課発注点"])

    for code_n, grp in df_yayoi.groupby("品番_正規化", sort=False):
        grp_sorted = grp.sort_values(by=["受注日_dt", "納入期日_dt"], ascending=[False, False], na_position="last")
        yayoi_dict[code_n] = grp_sorted[["受注番号", "得意先名", "数量", "受注日", "納入期日", "倉庫名", "備考", "担当者名"]].to_dict(orient="records")
        yayoi_comp_search_map[code_n] = " ".join(grp_sorted["得意先名"].unique())

        c_rank = grp.groupby("得意先名", as_index=False).agg(qty=("数量", "sum"), cnt=("受注番号", "count"), last_dt=("受注日", "max")).sort_values("qty", ascending=False)
        tot_q = c_rank["qty"].sum()
        yayoi_cust_rank_map[code_n] = [
            {
                "得意先名": r_c["得意先名"],
                "数量": int(r_c["qty"]),
                "構成比": f"{round(r_c['qty'] / tot_q * 100, 1) if tot_q > 0 else 0}%",
                "伝票回数": int(r_c["cnt"]),
                "最終受注日": str(r_c["last_dt"] or "—")
            }
            for _, r_c in c_rank.iterrows()
        ]

item_wh_map = {}
item_name_map = {}
for it in master_items:
    c_norm = normalize_code(it["code"])
    raw_wh = clean_str(it.get("wh", ""))
    item_wh_map[c_norm] = raw_wh
    item_wh_map[it["code"]] = raw_wh
    item_name_map[c_norm] = it["name"]
    item_name_map[it["code"]] = it["name"]

for c_n, r_nm in r_name_map.items():
    if c_n not in item_name_map or not item_name_map[c_n]:
        item_name_map[c_n] = r_nm

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

target_norm_codes = tuple(sorted(df_main["品番_正規化"].unique()))
up_sales_bytes = up_sales.getvalue() if up_sales is not None else None
up_sales_name = up_sales.name if up_sales is not None else ""

# 3. 売上データ(yayoi_CSV)から先週・今週・今期・過去期の購入先と売上金額を集計
sales_comp_map, sales_det_map, sales_stats, sales_src_info, yayoi_csv_cnt = fetch_and_build_yayoi_sales_from_gdrive(
    up_sales_bytes, up_sales_name, target_norm_codes
)

this_w_qty_map = sales_stats.get("this_w_qty", {})
prev_w_qty_map = sales_stats.get("prev_w_qty", {})
konki_qty_map = sales_stats.get("konki_qty", {})
konki_amt_map = sales_stats.get("konki_amt", {})
this_w_label = sales_stats.get("this_w_label", "今週(直近7日)")
prev_w_label = sales_stats.get("prev_w_label", "先週(前週7日)")

df_main["今期出荷"] = df_main["品番_正規化"].map(konki_qty_map).fillna(0).astype(int)
df_main["先週売数"] = df_main["品番_正規化"].map(prev_w_qty_map).fillna(0).astype(int)
df_main["今週売数"] = df_main["品番_正規化"].map(this_w_qty_map).fillna(0).astype(int)

# 4. 『今期（26/7〜）の売上だけ』で今期売上高・今期粗利額・今期粗利率を算出（過去期は一切混ぜない）
new_costs, min_gedais, avg_prices, sales_amts, gross_profits, margin_rates, stock_vals = [], [], [], [], [], [], []
is_new_list, is_surge_list, is_dead_list = [], [], []

for _, r in df_main.iterrows():
    cn = r["品番_正規化"]
    # ① 原価・下代は『order_r（マスター）.xlsx』の「新原価」「item」シートから取得
    u_cost = float(cost_map.get(cn, 0.0))
    raw_gedai = float(gedai_map.get(cn, 0.0))
    raw_std_p = float(std_price_map.get(cn, 0.0))

    # 標準販売単価（itemシートの販売価格 または 新原価シートの最低下代）
    if 0 < raw_std_p < 90000:
        u_gedai = raw_std_p
    elif 0 < raw_gedai < 90000:
        u_gedai = raw_gedai
    else:
        u_gedai = round(u_cost / 0.65, 0) if u_cost > 0 else 0.0

    # ② 今期（26/7〜）の売上数量と売上金額のみを使用！
    konki_q = int(r["今期出荷"])
    k77_q = int(r["77期"])
    csv_amt = float(konki_amt_map.get(cn, 0.0))
    csv_unit_p = (csv_amt / konki_q) if (konki_q > 0 and csv_amt > 0) else 0.0

    # 今期の実売平均単価（CSVの金額が原価以上なら実売単価、もしCSVに単価が入っていない・原価割れならorder_rの下代で補完）
    if csv_unit_p > u_cost and u_cost > 0:
        u_avg_p = csv_unit_p
    elif u_gedai > u_cost and u_cost > 0:
        u_avg_p = u_gedai
    elif u_cost > 0:
        u_avg_p = round(u_cost / 0.65, 0)
    else:
        u_avg_p = csv_unit_p if csv_unit_p > 0 else u_gedai

    # ③ 今期の売上高・今期の粗利額・今期の粗利率を算出（今期売れていない商品はすべて 0 にする）
    if konki_q > 0 and u_avg_p > 0:
        # CSVの実売合計金額が原価合計を超えていればそのまま使用、不足していれば個数×単価で算出
        s_amt = int(round(csv_amt)) if (csv_amt >= konki_q * u_cost and csv_amt > 0) else int(round(konki_q * u_avg_p))
        cost_total = int(round(konki_q * u_cost)) if u_cost > 0 else int(round(s_amt * 0.65))
        g_profit = max(0, s_amt - cost_total)
        m_rate = round((g_profit / s_amt) * 100, 1) if s_amt > 0 else 0.0
    else:
        s_amt = 0
        g_profit = 0
        m_rate = 0.0

    st_val = int(round(max(int(r["在庫数_ロジザード"]), 0) * u_cost))

    # 🆕 新商品 ／ ⚡ 急増 ／ ❄️ 滞留 の判定
    is_new = (int(r["74期"]) == 0 and int(r["75期"]) == 0 and (int(r["76期"]) > 0 or k77_q > 0 or konki_q > 0))
    avg_weekly = max((konki_q / 20.0) if konki_q > 0 else (k77_q / 52.0), 10.0)
    w_this = int(r["今週売数"])
    w_prev = int(r["先週売数"])
    is_surge = (w_this >= 30 and w_this >= max(w_prev * 1.5, avg_weekly * 1.8))
    is_dead = ((int(r["入荷後のフリー在庫"]) >= 1000 or st_val >= 300000) and int(r["在庫回転率"]) <= 15 and konki_q == 0)

    new_costs.append(int(round(u_cost)))
    min_gedais.append(int(round(u_gedai)))
    avg_prices.append(int(round(u_avg_p)))
    sales_amts.append(s_amt)
    gross_profits.append(g_profit)
    margin_rates.append(m_rate)
    stock_vals.append(st_val)
    is_new_list.append(is_new)
    is_surge_list.append(is_surge)
    is_dead_list.append(is_dead)

df_main["新原価"] = new_costs
df_main["最低下代"] = min_gedais
df_main["平均単価"] = avg_prices
df_main["売上高"] = sales_amts
df_main["粗利額"] = gross_profits
df_main["粗利率"] = margin_rates
df_main["在庫金額"] = stock_vals
df_main["新商品"] = is_new_list
df_main["急増"] = is_surge_list
df_main["滞留"] = is_dead_list

# 👑 利益貢献度ランク（ABC分析：粗利額上位75%=A, 92%=B, その他=C）
df_profit_sorted = df_main.sort_values("粗利額", ascending=False)
total_pos_profit = df_profit_sorted[df_profit_sorted["粗利額"] > 0]["粗利額"].sum()
abc_map = {}
cum_p = 0
for _, rp in df_profit_sorted.iterrows():
    cn = rp["品番_正規化"]
    gp = rp["粗利額"]
    if gp <= 0 or total_pos_profit <= 0:
        abc_map[cn] = "C"
    else:
        cum_p += gp
        ratio = cum_p / total_pos_profit
        if ratio <= 0.75:
            abc_map[cn] = "A"
        elif ratio <= 0.92:
            abc_map[cn] = "B"
        else:
            abc_map[cn] = "C"

df_main["ABCランク"] = df_main["品番_正規化"].map(abc_map).fillna("C")

def make_badge_str(row):
    b = []
    if row["ABCランク"] == "A":
        b.append("👑A")
    elif row["ABCランク"] == "B":
        b.append("B")
    if row["新商品"]:
        b.append("🆕新")
    if row["急増"]:
        b.append("⚡急増")
    if row["滞留"]:
        b.append("❄️滞留")
    return " ".join(b) if b else "C"

df_main["バッジ表示"] = df_main.apply(make_badge_str, axis=1)

sales_comp_search_str_map = {}
for c_n, ki_dict in sales_comp_map.items():
    all_c_list = ki_dict.get("ALL", [])
    sales_comp_search_str_map[c_n] = " ".join([x["得意先名"] for x in all_c_list])

df_main["関連会社検索用"] = (
    df_main["品番_正規化"].map(yayoi_comp_search_map).fillna("") + " " +
    df_main["品番_正規化"].map(sales_comp_search_str_map).fillna("")
)
m_df = df_main.copy()

unregistered_list = []
all_kakuho_check = {}
if not df_yayoi.empty:
    for c_n, grp in df_yayoi.groupby("品番_正規化"):
        all_kakuho_check[c_n] = int(grp["数量"].sum())
else:
    for c_n, q in extra_meta.get("sheet_yayoi_sum", {}).items():
        all_kakuho_check[c_n] = all_kakuho_check.get(c_n, 0) + q

for c_n, k_qty in all_kakuho_check.items():
    if k_qty > 0 and c_n not in set(target_norm_codes):
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
            "新原価": int(round(cost_map.get(c_n, 0))),
            "入数": extra_meta.get("lot_map", {}).get(c_n, 1)
        })
df_unreg = pd.DataFrame(unregistered_list)

hist_dict = {}
if df_hist is not None and not df_hist.empty:
    for p_code, grp in df_hist.groupby("品番", sort=False):
        hist_dict[p_code] = grp.sort_values("会議日", ascending=False)[["会議日", "決定事項"]].to_dict(orient="records")

src_status = f"✅ Googleドライブ同期完了 ({gdrive_sync['source_label']})" if gdrive_sync["main_bytes"] is not None else "📂 ローカル参照中"
zure_count = int((m_df["確保数のズレ"] != 0).sum())
total_gross_profit = int(m_df["粗利額"].sum())
total_stock_val = int(m_df["在庫金額"].sum())

st.caption(
    f"**{src_status}** ｜ 💰 原価元: **{cost_src_label}** ｜ 📑 確保元: **{yayoi_src_label or '未検出'}** ({len(df_yayoi):,}行) ｜ "
    f"🔥 売上: **{sales_src_info or 'yayoi_CSV未検出'}** ｜ "
    f"👑 総粗利: **¥{total_gross_profit:,}** ／ 倉庫在庫金額(原価): **¥{total_stock_val:,}** ｜ 🚨 要発注: **{int((m_df['発注検討フラグ'] < 0).sum())}** 件"
)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    f"📋 1. 発注検討表（週次売れ数・新原価・粗利・バッジ搭載）",
    "🏢 2. 会社（得意先）別・受注番号別 確保リスト",
    "💻 3. 弥生受注伝票（リアルタイム発注判定つき）",
    "🕒 4. 入荷予定 ＆ 過去の会議決定履歴",
    f"🚀 5. 営業用・大口注文シミュレーター ＆ 抜け防止 ({len(df_unreg)}件)"
])

# ==============================================================================
# 【パート 3 / 3】画面UI・先週vs今週クリックPOP・新原価＆粗利＆発注予定金額リアルタイム計算
# ==============================================================================
with tab1:
    center_list = sorted([x for x in m_df["商品センター"].unique() if x])
    centers_json = json.dumps(center_list, ensure_ascii=False)
    records_json = json.dumps(m_df.to_dict(orient="records"), ensure_ascii=False)
    yayoi_json = json.dumps(yayoi_dict, ensure_ascii=False)
    yayoi_rank_json = json.dumps(yayoi_cust_rank_map, ensure_ascii=False)
    sales_comp_json = json.dumps(sales_comp_map, ensure_ascii=False)
    sales_det_json = json.dumps(sales_det_map, ensure_ascii=False)
    kamei_json = json.dumps(kamei_price_map, ensure_ascii=False)
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
            display: grid; grid-template-columns: 2.6fr 1.0fr 2.9fr; gap: 14px;
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
            border-bottom: none; border-radius: 6px 6px 0 0; flex-wrap: wrap; gap: 8px;
        }}
        .toolbar-info {{ font-size: 12.5px; color: #1e293b; font-weight: bold; display: flex; gap: 14px; align-items: center; flex-wrap: wrap; }}
        .kpi-pill {{ background: #fff; border: 1px solid #94a3b8; padding: 3px 10px; border-radius: 5px; font-size: 12.5px; }}
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
        .w-badge {{ width: 92px; min-width: 92px; text-align: center; font-weight: bold; font-size: 11px; white-space: nowrap; }}
        .w-num {{ width: 72px; min-width: 72px; text-align: right; font-variant-numeric: tabular-nums; }}
        .w-money {{ width: 88px; min-width: 88px; text-align: right; font-variant-numeric: tabular-nums; }}
        .w-date {{ width: 95px; min-width: 95px; text-align: center; }}
        .w-prev-dec {{ width: 160px; min-width: 160px; max-width: 160px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .w-dec {{ width: 220px; min-width: 220px; }}

        .bg-cyan {{ background-color: #67e8f9 !important; }}
        .bg-flag-neg {{ background-color: #fecdd3 !important; color: #be123c; font-weight: bold; }}
        .bg-green {{ background-color: #a3e635 !important; }}
        .bg-pink-rate {{ background-color: #fbcfe8 !important; color: #be123c; font-weight: bold; }}
        .bg-yellow-ro {{ background-color: #fef08a !important; color: #1e293b; }}
        .bg-yellow-edit {{ background-color: #ffff00 !important; }}
        .bg-hist-ki {{ background-color: #f1f5f9; color: #334155; }}
        .bg-konki-hit {{ background-color: #ffedd5 !important; color: #c2410c; font-weight: bold; }}
        .bg-week-hit {{ background-color: #fef3c7 !important; color: #b45309; font-weight: bold; }}
        .bg-profit {{ background-color: #ecfdf5 !important; color: #047857; font-weight: bold; }}
        .bg-zure-diff {{ background-color: #ffe4e6 !important; color: #be123c; font-weight: bold; }}

        tbody tr:hover td {{ filter: brightness(0.95); }}
        .click-link {{ color: #0284c7; font-weight: bold; text-decoration: underline; cursor: pointer; display: block; width: 100%; }}
        .click-link-ki {{ color: #1e40af; font-weight: bold; text-decoration: underline; cursor: pointer; display: block; width: 100%; }}
        .click-link-konki {{ color: #c2410c; font-weight: bold; text-decoration: underline; cursor: pointer; display: block; width: 100%; }}
        .click-link-profit {{ color: #047857; font-weight: bold; text-decoration: underline; cursor: pointer; display: block; width: 100%; }}
        .click-link-zure {{ color: #be123c; font-weight: bold; text-decoration: underline; cursor: pointer; display: block; width: 100%; }}
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
        .summary-bar {{ display: flex; gap: 20px; flex-wrap: wrap; background: #f1f5f9; padding: 12px 18px; border-radius: 8px; border-left: 6px solid #2563eb; font-size: 14.5px; align-items: center; }}
        .summary-item {{ font-weight: bold; color: #1e293b; }}
        .summary-item span {{ color: #0284c7; font-size: 16.5px; margin-left: 4px; }}
        .mode-btn {{ padding: 6px 14px; border-radius: 6px; border: 1px solid #cbd5e1; font-weight: bold; cursor: pointer; font-size: 13.5px; background: #fff; color: #1e293b; }}
        .mode-btn.active {{ background: #ea580c; color: #fff; border-color: #c2410c; }}
        .quick-edit-box {{ background: #fffbeb; border: 2px solid #fde68a; padding: 12px 16px; border-radius: 8px; display: flex; flex-direction: column; gap: 10px; }}
        .stamp-btn {{ background: #e2e8f0; border: 1px solid #94a3b8; border-radius: 5px; padding: 4px 10px; font-size: 13px; cursor: pointer; font-weight: bold; }}
        .modal-grid {{ display: grid; grid-template-columns: 3.5fr 1.5fr; gap: 20px; }}
        .sub-table {{ width: 100%; border-collapse: collapse; }}
        .sub-table th {{ background: #e2e8f0; color: #1e293b; padding: 8px 10px; font-size: 14px; height: auto; position: static; border: 1px solid #cbd5e1; }}
        .sub-table td {{ padding: 8px 10px; font-size: 14px; height: auto; border: 1px solid #cbd5e1; }}
        .sec-title {{ font-size: 16px; font-weight: bold; color: #1e3a8a; margin-bottom: 8px; border-bottom: 2px solid #93c5fd; padding-bottom: 4px; }}
        .diag-box {{ padding: 14px 18px; border-radius: 8px; margin-bottom: 14px; font-size: 14.5px; line-height: 1.6; }}
    </style>
    </head>
    <body>
    <div class="instant-filter-bar">
        <div class="filter-group">
            <label class="filter-label">表示絞り込み・並び替え（全アイデア搭載・瞬時切替）</label>
            <select id="js-filter-mode" class="filter-select" onchange="applyInstantFilter()">
                <option value="all">すべて表示（Excelと同じ並び）</option>
                <option value="need_order">🚨 要発注（発注検討フラグがマイナスのみ）</option>
                <option value="need_order_a">👑 【最優先】Aランク主力商品で要発注のもの</option>
                <option value="this_week_sold">📅 今週（直近7日間）売れた商品（売れ数順）</option>
                <option value="surge">⚡ 【急増アラート】今週急激に売れている商品</option>
                <option value="new_prod">🆕 【新商品】76期以降に登場した新商品の売れ行き</option>
                <option value="profit_rank">💰 粗利額（利益貢献度）が高い順ランキング</option>
                <option value="dead_stock">❄️ 【滞留・過剰在庫】在庫が多く回転率が低い商品（在庫金額順）</option>
                <option value="konki_sold">🔥 今期(26/7〜)に売れている商品（多い順）</option>
                <option value="zure">⚖️ 確保数のズレがある商品のみ</option>
                <option value="incoming">🚢 入荷予定がある商品</option>
                <option value="secured">🏢 弥生確保がある商品</option>
            </select>
        </div>
        <div class="filter-group">
            <label class="filter-label">商品センター</label>
            <select id="js-filter-center" class="filter-select" onchange="applyInstantFilter()">
                <option value="すべて">すべて</option>
            </select>
        </div>
        <div class="filter-group">
            <label class="filter-label">🔍 品番・商品名・会社名（購入先/確保先）・バッジ(新/A/急増)・決定事項で検索</label>
            <input type="text" id="js-search-kw" class="filter-input" placeholder="例: ABO-08PK / 👑A / 🆕新 / ⚡急増 / 丸藤 / トップ産業" oninput="applyInstantFilter()">
        </div>
    </div>

    <div class="toolbar">
        <div class="toolbar-info">
            <span>表示: <b id="row-count">0</b> 件</span>
            <span class="kpi-pill">💰 決定事項の発注予定総額(新原価): <b id="planned-order-cost" style="color:#be123c; font-size:14px;">¥0</b></span>
            <span style="font-size:12px; color:#475569;">💡 「先週/今週売数」「今期」「74〜77期」「粗利額」をクリックすると販売先POPが開きます！</span>
        </div>
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
                    <th class="w-badge" style="background:#f8fafc;">ランク/<br>バッジ</th>
                    <th class="w-num bg-hist-ki">74期<br>実績<br><span style="color:#1d4ed8; font-size:10px;">🔍購入先</span></th>
                    <th class="w-num bg-hist-ki">75期<br>実績<br><span style="color:#1d4ed8; font-size:10px;">🔍購入先</span></th>
                    <th class="w-num bg-hist-ki">76期<br>実績<br><span style="color:#1d4ed8; font-size:10px;">🔍購入先</span></th>
                    <th class="w-num bg-hist-ki">77期<br>(前期)<br><span style="color:#1d4ed8; font-size:10px;">🔍購入先</span></th>
                    <th class="w-num" style="background:#ffedd5;">🔥今期<br>(26/7〜)<br><span style="color:#c2410c; font-size:10px;">🔍購入先</span></th>
                    <th class="w-num" style="background:#fef3c7;">先週売数<br><span style="color:#b45309; font-size:10px;">🔍前週7日</span></th>
                    <th class="w-num" style="background:#fde68a;">今週売数<br><span style="color:#b45309; font-size:10px;">🔍直近7日</span></th>
                    <th class="w-num" style="background:#ecfdf5;">新原価<br><span style="color:#047857; font-size:10px;">(order_r)</span></th>
                    <th class="w-money" style="background:#ecfdf5;">売上高<br><span style="color:#047857; font-size:10px;">(税抜)</span></th>
                    <th class="w-money" style="background:#d1fae5;">粗利額<br>(利益)<br><span style="color:#047857; font-size:10px;">🔍内訳</span></th>
                    <th class="w-num" style="background:#ecfdf5;">粗利率</th>
                    <th class="w-num" style="background:#e0f2fe;">1課発注点<br><span style="color:#0369a1; font-size:10px;">✏️手動修正可</span></th>
                    <th class="w-num" style="background:#e0f2fe;">2課発注点<br><span style="color:#0369a1; font-size:10px;">✏️手動修正可</span></th>
                    <th class="w-num bg-cyan">在庫数<br><span class="red-txt">ロジザード</span></th>
                    <th class="w-num bg-cyan">入荷予定<br>数<br><span class="red-txt">ロジザード</span></th>
                    <th class="w-date bg-cyan">入荷予定<br>日<br><span class="red-txt">ロジザード</span></th>
                    <th class="w-num bg-cyan">確保数<br><span class="red-txt">ロジザード<br>(在庫確保)</span></th>
                    <th class="w-num bg-cyan">在庫確保<br><span class="red-txt">(弥生)</span></th>
                    <th class="w-num bg-cyan">確保数の<br>ズレ<br><span class="red-txt">🔍ズレ一覧</span></th>
                    <th class="w-num bg-cyan">入荷後の<br>フリー在庫</th>
                    <th class="w-num">発注検討<br>フラグ<br><span class="red-txt">マイナスの<br>場合発注検討</span></th>
                    <th class="w-num">在庫回転<br>率</th>
                    <th class="w-num">リンベル<br>確保数</th>
                    <th class="w-prev-dec bg-yellow-ro">{prev_title_html}</th>
                    <th class="w-dec">決定事項<br><span style="font-size:10px; color:#b91c1c;">(数値入力で発注金額自動計算)</span></th>
                    <th class="w-date">発注日</th>
                </tr>
            </thead>
            <tbody id="table-body"></tbody>
        </table>
    </div>

    <div class="modal-overlay" id="detail-modal" onclick="closeModalOnBg(event)">
        <div class="modal-window">
<div class="modal-header">
                <div class="modal-title" id="modal-title-text">📦 品番詳細・内訳データ</div>
                <div style="display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
                    <button class="mode-btn" id="btn-mode-sales" onclick="switchModalMode('sales', '今期(26/7〜)')">🔥 購入先（売上・粗利）内訳</button>
                    <button class="mode-btn" id="btn-mode-kamei" onclick="switchModalMode('kamei', '')">💴 会社別の販売価格（加盟店価格）一覧</button>
                    <button class="mode-btn" id="btn-mode-zure" onclick="switchModalMode('zure', '')">⚖️ 確保ズレ診断＆全ズレ商品一覧</button>
                    <button class="mode-btn" id="btn-mode-yayoi" onclick="switchModalMode('yayoi', '')">🏢 弥生確保リスト（すべて）</button>
                    <button class="modal-close" onclick="closeModal()">✕ 閉じる</button>
                </div>
            </div>
            <div class="modal-body">
                <div class="summary-bar" id="modal-summary"></div>
                <div class="quick-edit-box">
                    <div style="font-size:14px; font-weight:bold; color:#92400e; display:flex; justify-content:space-between;">
                        <span>✏️ この場で決定事項・入荷予定日を入力（表＆発注予定金額にもリアルタイム連動）</span>
                        <span>
                            スタンプ:
                            <button class="stamp-btn" onclick="addStamp('様子見')">＋様子見</button>
                            <button class="stamp-btn" onclick="addStamp('発注済')">＋発注済</button>
                            <button class="stamp-btn" onclick="addStamp('確認中')">＋確認中</button>
                            <button class="stamp-btn" onclick="addStamp('ロジ修正済')">＋ロジ修正済</button>
                            <button class="stamp-btn" onclick="addStamp('廃盤予定')">＋廃盤予定</button>
                        </span>
                    </div>
                    <div style="display:flex; gap:12px; align-items:center;">
                        <label style="font-size:14px; font-weight:bold;">前回:</label>
                        <input type="text" id="modal-view-prev" class="cell-input" style="width:200px; height:32px; background:#fef08a; font-size:14px;" readonly>
                        <label style="font-size:14px; font-weight:bold;">入荷予定日:</label>
                        <input type="text" id="modal-input-ndate" class="cell-input" style="width:140px; height:32px; font-size:14px;" placeholder="例: 11/20" oninput="syncFromModal()">
                        <label style="font-size:14px; font-weight:bold;">決定事項:</label>
                        <input type="text" id="modal-input-dec" class="cell-input" style="flex:1; height:32px; font-size:14px;" placeholder="例: 3000個発注" oninput="syncFromModal()">
                        <label style="font-size:14px; font-weight:bold;">発注日:</label>
                        <input type="text" id="modal-input-odate" class="cell-input" style="width:130px; height:32px; font-size:14px;" oninput="syncFromModal()">
                    </div>
                </div>
                <div class="modal-grid">
                    <div>
                        <div class="sec-title" id="modal-sub-title">🏢 詳細リスト</div>
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
        const yayoiRankMap = {yayoi_rank_json};
        const salesCompMap = {sales_comp_json};
        const salesDetMap = {sales_det_json};
        const kameiPriceMap = {kamei_json};
        const histMap = {hist_json};
        const thisWeekLabel = "{this_w_label}";
        const prevWeekLabel = "{prev_w_label}";
        let currentModalIdx = null;
        let currentModalMode = "yayoi";
        let currentKiFilter = "今期(26/7〜)";
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

        // 会社名から加盟店価格（order_r）を曖昧一致で検索するヘルパー関数
        function findCompanyKameiPrice(codeNorm, targetCompName) {{
            const kList = kameiPriceMap[codeNorm] || [];
            if (!targetCompName || kList.length === 0) return null;
            const cleanTarget = String(targetCompName).replace(/株式会社|有限会社|\(株\)|（株）|㈱|㈲|\s| /g, "");
            if (!cleanTarget) return null;
            for (const item of kList) {{
                const cleanK = String(item["会社名"] || "").replace(/株式会社|有限会社|\(株\)|（株）|㈱|㈲|\s| /g, "");
                if (cleanK && (cleanTarget.includes(cleanK) || cleanK.includes(cleanTarget))) {{
                    return item["販売価格"];
                }}
            }}
            return null;
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

        function updatePlannedOrderCost() {{
            let totalOrderCost = 0;
            tableData.forEach(r => {{
                const decStr = String(r["決定事項"] || "").replace(/,/g, "");
                if (!decStr || decStr.includes("様子見") || decStr.includes("不要") || decStr.includes("廃盤")) return;
                const m = decStr.match(/(\\d+)/);
                if (m) {{
                    const q = Number(m[1]);
                    const cost = Number(r["新原価"] || 0);
                    if (q > 0 && cost > 0) totalOrderCost += (q * cost);
                }}
            }});
            const el = document.getElementById("planned-order-cost");
            if (el) el.textContent = "¥" + totalOrderCost.toLocaleString("ja-JP");
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
                if (fMode === "need_order_a" && !(Number(r["発注検討フラグ"]) < 0 && r["ABCランク"] === "A")) return false;
                if (fMode === "this_week_sold" && !(Number(r["今週売数"]) > 0)) return false;
                if (fMode === "surge" && !r["急増"]) return false;
                if (fMode === "new_prod" && !r["新商品"]) return false;
                if (fMode === "profit_rank" && !(Number(r["粗利額"]) > 0)) return false;
                if (fMode === "dead_stock" && !r["滞留"]) return false;
                if (fMode === "konki_sold" && !(Number(r["今期出荷"]) > 0)) return false;
                if (fMode === "zure" && !(Number(r["確保数のズレ"]) !== 0)) return false;
                if (fMode === "incoming" && !(Number(r["入荷予定数_ロジザード"]) > 0 || (r["入荷予定日_ロジザード"] && r["入荷予定日_ロジザード"] !== ""))) return false;
                if (fMode === "secured" && !(Number(r["在庫確保_弥生"]) > 0)) return false;

                if (fCenter !== "すべて" && r["商品センター"] !== fCenter) return false;

                if (kw) {{
                    const targetStr = [
                        r["品番"], r["商品区分"], r["バッジ表示"], r["前回決定事項"], r["決定事項"], r["関連会社検索用"]
                    ].join(" ").toLowerCase();
                    if (!targetStr.includes(kw)) return false;
                }}
                return true;
            }});

            if (fMode === "this_week_sold" || fMode === "surge") {{
                filtered.sort((a, b) => Number(b["今週売数"]) - Number(a["今週売数"]));
            }} else if (fMode === "profit_rank") {{
                filtered.sort((a, b) => Number(b["粗利額"]) - Number(a["粗利額"]));
            }} else if (fMode === "dead_stock") {{
                filtered.sort((a, b) => Number(b["在庫金額"]) - Number(a["在庫金額"]));
            }} else if (fMode === "konki_sold") {{
                filtered.sort((a, b) => Number(b["今期出荷"]) - Number(a["今期出荷"]));
            }}

            visibleIndices = filtered.map(r => r["_origIdx"]);
            renderTable(filtered);
            updatePlannedOrderCost();
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
                const zureLinkCls = (zureVal !== 0) ? "click-link-zure" : "click-link";
                const konkiVal = Number(r["今期出荷"] || 0);
                const konkiCls = (konkiVal > 0) ? "bg-konki-hit" : "bg-hist-ki";
                const prevWVal = Number(r["先週売数"] || 0);
                const thisWVal = Number(r["今週売数"] || 0);
                const thisWCls = (thisWVal > 0) ? "bg-week-hit" : "bg-hist-ki";
                const badgeBg = r["急増"] ? "background:#fef08a; color:#b91c1c;" : (r["新商品"] ? "background:#dbeafe; color:#1d4ed8;" : (r["ABCランク"]==="A" ? "background:#ffedd5; color:#c2410c;" : ""));

                html += `<tr>
                    <td class="w-no sticky-col-1">${{r["No"]}}</td>
                    <td class="w-center sticky-col-2">${{r["商品センター"]}}</td>
                    <td class="w-code sticky-col-3" title="${{r["品番"]}}"><span class="click-link" onclick="openModal(${{idx}}, 'yayoi', '')">${{r["品番"]}}</span></td>
                    <td class="w-name sticky-col-4" title="${{r["商品区分"]}}">${{r["商品区分"]}}</td>
                    <td class="w-badge" style="${{badgeBg}}">${{r["バッジ表示"]}}</td>
                    <td class="w-num bg-hist-ki"><span class="click-link-ki" onclick="openModal(${{idx}}, 'sales', '74期')">${{fmtNum(r["74期"])}}</span></td>
                    <td class="w-num bg-hist-ki"><span class="click-link-ki" onclick="openModal(${{idx}}, 'sales', '75期')">${{fmtNum(r["75期"])}}</span></td>
                    <td class="w-num bg-hist-ki"><span class="click-link-ki" onclick="openModal(${{idx}}, 'sales', '76期')">${{fmtNum(r["76期"])}}</span></td>
                    <td class="w-num bg-hist-ki"><span class="click-link-ki" style="color:#0f172a;" onclick="openModal(${{idx}}, 'sales', '77期')">${{fmtNum(r["77期"])}}</span></td>
                    <td class="w-num ${{konkiCls}}"><span class="click-link-konki" onclick="openModal(${{idx}}, 'sales', '今期(26/7〜)')">${{konkiVal > 0 ? "🔥 " + fmtNum(konkiVal) : "0"}}</span></td>
                    <td class="w-num bg-hist-ki"><span class="click-link-ki" onclick="openModal(${{idx}}, 'sales', 'PREV_WEEK')">${{fmtNum(prevWVal)}}</span></td>
                    <td class="w-num ${{thisWCls}}"><span class="click-link-konki" onclick="openModal(${{idx}}, 'sales', 'THIS_WEEK')">${{thisWVal > 0 ? "⚡ " + fmtNum(thisWVal) : "0"}}</span></td>
                    <td class="w-num" style="background:#f0fdf4; font-weight:bold;"><span class="click-link-profit" onclick="openModal(${{idx}}, 'kamei', '')">¥${{fmtNum(r["新原価"])}}</span></td>
                    <td class="w-money" style="background:#f0fdf4;"><span class="click-link-profit" onclick="openModal(${{idx}}, 'sales', '今期(26/7〜)')">¥${{fmtNum(r["売上高"])}}</span></td>
                    <td class="w-money bg-profit"><span class="click-link-profit" onclick="openModal(${{idx}}, 'sales', '今期(26/7〜)')">¥${{fmtNum(r["粗利額"])}}</span></td>
                    <td class="w-num" style="background:#f0fdf4; font-weight:bold;"><span class="click-link-profit" onclick="openModal(${{idx}}, 'sales', '今期(26/7〜)')">${{r["粗利率"]}}%</span></td>
                    <td class="w-num"><input type="number" class="cell-input-point" id="point1-${{idx}}" value="${{r["1課発注点"]}}" oninput="tableData[${{idx}}]['1課発注点']=Number(this.value); recalculateRow(${{idx}})" onkeydown="handleKeyDown(event, 'point1', ${{idx}})"></td>
                    <td class="w-num"><input type="number" class="cell-input-point" id="point2-${{idx}}" value="${{r["2課発注点"]}}" oninput="tableData[${{idx}}]['2課発注点']=Number(this.value); recalculateRow(${{idx}})" onkeydown="handleKeyDown(event, 'point2', ${{idx}})"></td>
                    <td class="w-num bg-cyan">${{fmtNum(r["在庫数_ロジザード"])}}</td>
                    <td class="w-num bg-cyan">${{fmtNum(r["入荷予定数_ロジザード"])}}</td>
                    <td class="w-date bg-cyan"><input type="text" class="cell-input" id="ndate-${{idx}}" value="${{r["入荷予定日_ロジザード"] || ""}}" placeholder="日付入力" oninput="updateCell(${{idx}}, '入荷予定日_ロジザード', this.value)" onkeydown="handleKeyDown(event, 'ndate', ${{idx}})"></td>
                    <td class="w-num bg-cyan">${{fmtNum(r["確保数_ロジザード"])}}</td>
                    <td class="w-num bg-cyan"><span class="click-link" onclick="openModal(${{idx}}, 'yayoi', '')">${{fmtNum(r["在庫確保_弥生"])}}</span></td>
                    <td class="w-num ${{zureCls}}" id="zure-${{idx}}"><span class="${{zureLinkCls}}" onclick="openModal(${{idx}}, 'zure', '')">${{fmtNum(r["確保数のズレ"])}}</span></td>
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
            updatePlannedOrderCost();
        }}

        function updateCell(idx, field, val) {{
            tableData[idx][field] = val;
            saveToStorage(tableData[idx]["品番"], tableData[idx]);
        }}

        function switchModalMode(mode, kiStr) {{
            if (currentModalIdx === null) return;
            openModal(currentModalIdx, mode, kiStr || "今期(26/7〜)");
        }}

        // 会社別販売価格（加盟店価格）テーブルのHTMLを生成する共通関数
        function buildKameiPriceTableHtml(codeNorm, unitCost, minGedai) {{
            const kList = kameiPriceMap[codeNorm] || [];
            if (kList.length === 0) {{
                return `<div style="padding:14px; background:#f8fafc; color:#64748b; border:1px solid #cbd5e1; border-radius:6px;">
                    この品番（${{codeNorm}}）は『order_r（マスター）.xlsx』の「加盟店価格」シートに個別登録がありません。（標準最低下代: ¥${{fmtNum(minGedai)}} ／ 商品原価: ¥${{fmtNum(unitCost)}}）
                </div>`;
            }}
            let html = `<table class="sub-table"><thead><tr>
                <th style="width:50px;">No</th>
                <th>会社名（加盟店名 / ランク名）</th>
                <th style="width:95px;">グループ連番</th>
                <th style="width:80px;">ロット数</th>
                <th style="width:130px; background:#fef3c7; color:#b45309;">販売価格(ロット単価)</th>
                <th style="width:110px;">商品原価</th>
                <th style="width:120px;">1個あたり粗利</th>
                <th style="width:95px;">粗利率</th>
                <th style="width:120px;">最低下代(¥${{fmtNum(minGedai)}})比</th>
            </tr></thead><tbody>`;
            kList.forEach((kp, idx) => {{
                const sp = Number(kp["販売価格"] || 0);
                const prof = sp - unitCost;
                const mRate = (sp > 0 && unitCost > 0) ? ((prof / sp) * 100).toFixed(1) + "%" : "—";
                const diffGedai = (minGedai > 0) ? (sp - minGedai) : 0;
                const diffStr = (minGedai > 0) ? ((diffGedai >= 0 ? "+" : "") + fmtNum(diffGedai) + "円") : "—";
                const diffColor = (diffGedai < 0) ? "color:#be123c; font-weight:bold;" : "color:#047857;";
                html += `<tr>
                    <td style="text-align:center;">${{idx + 1}}</td>
                    <td style="font-weight:bold; color:#1e293b;">${{kp["会社名"]}}</td>
                    <td style="text-align:center; color:#64748b;">${{kp["グループ"] || "—"}}</td>
                    <td style="text-align:center;">${{fmtNum(kp["ロット数"])}}</td>
                    <td style="text-align:right; font-weight:bold; color:#b45309; background:#fffbeb; font-size:15.5px;">¥${{fmtNum(sp)}}</td>
                    <td style="text-align:right;">¥${{fmtNum(unitCost)}}</td>
                    <td style="text-align:right; font-weight:bold; color:${{prof>=0?'#047857':'#be123c'}}; background:#ecfdf5;">¥${{fmtNum(prof)}}</td>
                    <td style="text-align:right; font-weight:bold;">${{mRate}}</td>
                    <td style="text-align:right; ${{diffColor}}">${{diffStr}}</td>
                </tr>`;
            }});
            html += `</tbody></table>`;
            return html;
        }}

        function openModal(idx, mode, kiStr) {{
            currentModalIdx = idx;
            currentModalMode = mode;
            currentKiFilter = kiStr || "今期(26/7〜)";
            const r = tableData[idx];
            const code = r["品番"];
            const codeNorm = r["品番_正規化"] || code;
            const unitCost = Number(r["新原価"] || 0);
            const minGedai = Number(r["最低下代"] || 0);

            document.getElementById("btn-mode-sales").className = (mode === 'sales') ? "mode-btn active" : "mode-btn";
            document.getElementById("btn-mode-kamei").className = (mode === 'kamei') ? "mode-btn active" : "mode-btn";
            document.getElementById("btn-mode-zure").className = (mode === 'zure') ? "mode-btn active" : "mode-btn";
            document.getElementById("btn-mode-yayoi").className = (mode === 'yayoi') ? "mode-btn active" : "mode-btn";

            document.getElementById("modal-title-text").innerHTML = `📦 品番: 【${{code}}】 ${{r["商品区分"]}} ｜ ${{r["バッジ表示"]}} （拠点: ${{r["商品センター"]}}）`;

            // ★ サマリーバーの一番目立つ位置に「商品原価（新原価）」と「最低下代」を常時表示！
            document.getElementById("modal-summary").innerHTML = `
                <div class="summary-item" style="background:#dcfce7; padding:4px 12px; border-radius:6px; border:1px solid #86efac;">
                    💰 商品原価(新原価): <span style="color:#047857; font-size:19px;">¥${{fmtNum(unitCost)}}</span>
                </div>
                <div class="summary-item" style="background:#fef9c3; padding:4px 12px; border-radius:6px; border:1px solid #fde047;">
                    🏷️ 最低下代: <span style="color:#b45309; font-size:19px;">¥${{fmtNum(minGedai)}}</span>
                </div>
                <div class="summary-item">先週/今週売数: <span style="color:#b45309;">${{fmtNum(r["先週売数"])}} → ⚡${{fmtNum(r["今週売数"])}} 個</span></div>
                <div class="summary-item">🔥今期(26/7〜)出荷: <span style="color:#ea580c;">${{fmtNum(r["今期出荷"])}} 個</span></div>
                <div class="summary-item">今期売上 / 今期粗利: <span style="color:#047857;">¥${{fmtNum(r["売上高"])}} / ¥${{fmtNum(r["粗利額"])}} (${{r["粗利率"]}}%)</span></div>
                <div class="summary-item">フリー在庫: <span>${{fmtNum(r["入荷後のフリー在庫"])}} 個</span></div>
            `;
            document.getElementById("modal-view-prev").value = r["前回決定事項"] || "";
            document.getElementById("modal-input-ndate").value = r["入荷予定日_ロジザード"] || "";
            document.getElementById("modal-input-dec").value = r["決定事項"] || "";
            document.getElementById("modal-input-odate").value = r["発注日"] || "";

            // モードA：会社別の販売価格（加盟店価格）一覧モード
            if (mode === 'kamei') {{
                const kCount = (kameiPriceMap[codeNorm] || []).length;
                document.getElementById("modal-sub-title").innerHTML = `💴 【${{code}}】会社ごとの販売価格一覧（order_r『加盟店価格』シート参照：全 ${{kCount}} 件）`;
                document.getElementById("modal-yayoi-area").innerHTML = buildKameiPriceTableHtml(codeNorm, unitCost, minGedai);
            }}
// モードB：売上実績・購入先内訳モード（原価・利益額・利益率・合計行つき！）
            else if (mode === 'sales') {{
                let kiTitle = currentKiFilter;
                if (currentKiFilter === 'THIS_WEEK') kiTitle = thisWeekLabel;
                else if (currentKiFilter === 'PREV_WEEK') kiTitle = prevWeekLabel;
                else if (currentKiFilter === 'ALL') kiTitle = "全期間累計";

                let kiBtnsHtml = `
                    <div style="display:flex; gap:8px; margin-bottom:12px; align-items:center; flex-wrap:wrap;">
                        <span style="font-size:13px; font-weight:bold; color:#334155;">期間切替:</span>
                        <button class="stamp-btn" style="${{currentKiFilter==='THIS_WEEK'?'background:#b45309;color:#fff;':''}}" onclick="openModal(${{idx}}, 'sales', 'THIS_WEEK')">⚡ ${{thisWeekLabel}} (${{fmtNum(r['今週売数'])}}個)</button>
                        <button class="stamp-btn" style="${{currentKiFilter==='PREV_WEEK'?'background:#b45309;color:#fff;':''}}" onclick="openModal(${{idx}}, 'sales', 'PREV_WEEK')">📅 ${{prevWeekLabel}} (${{fmtNum(r['先週売数'])}}個)</button>
                        <button class="stamp-btn" style="${{currentKiFilter==='今期(26/7〜)'?'background:#ea580c;color:#fff;':''}}" onclick="openModal(${{idx}}, 'sales', '今期(26/7〜)')">🔥 今期 26/7〜 (${{fmtNum(r['今期出荷'])}}個)</button>
                        <button class="stamp-btn" style="${{currentKiFilter==='77期'?'background:#1e40af;color:#fff;':''}}" onclick="openModal(${{idx}}, 'sales', '77期')">77期 (${{fmtNum(r['77期'])}}個)</button>
                        <button class="stamp-btn" style="${{currentKiFilter==='76期'?'background:#1e40af;color:#fff;':''}}" onclick="openModal(${{idx}}, 'sales', '76期')">76期 (${{fmtNum(r['76期'])}}個)</button>
                        <button class="stamp-btn" style="${{currentKiFilter==='75期'?'background:#1e40af;color:#fff;':''}}" onclick="openModal(${{idx}}, 'sales', '75期')">75期 (${{fmtNum(r['75期'])}}個)</button>
                        <button class="stamp-btn" style="${{currentKiFilter==='74期'?'background:#1e40af;color:#fff;':''}}" onclick="openModal(${{idx}}, 'sales', '74期')">74期 (${{fmtNum(r['74期'])}}個)</button>
                        <button class="stamp-btn" style="${{currentKiFilter==='ALL'?'background:#047857;color:#fff;':''}}" onclick="openModal(${{idx}}, 'sales', 'ALL')">💰 全期間累計</button>
                    </div>
                `;
                document.getElementById("modal-sub-title").innerHTML = `🔥 【${{code}}】どこで売れたか（得意先別 売上・原価・利益額・利益率：${{kiTitle}}）`;

                const codeSalesObj = salesCompMap[codeNorm] || {{}};
                const compList = codeSalesObj[currentKiFilter] || [];
                const detList = salesDetMap[codeNorm] || [];

                let sHtml = kiBtnsHtml;
                if (compList.length > 0) {{
                    let sumQty = 0, sumSales = 0, sumCost = 0, sumProfit = 0, sumCnt = 0;

                    sHtml += `<table class="sub-table"><thead><tr>
                        <th style="width:42px;">順位</th>
                        <th>購入先（得意先名）</th>
                        <th style="width:90px;">購入数量</th>
                        <th style="width:68px;">構成比</th>
                        <th style="width:115px; background:#fef3c7; color:#b45309;">販売価格<br>(加盟店価格)</th>
                        <th style="width:115px; background:#e0f2fe; color:#0369a1;">原価<br>(1個 / 原価計)</th>
                        <th style="width:105px;">売上金額</th>
                        <th style="width:125px; background:#dcfce7; color:#047857;">利益額<br>(合計 / 1個利益)</th>
                        <th style="width:80px; background:#dcfce7; color:#047857;">利益率</th>
                        <th style="width:68px;">伝票数</th>
                        <th style="width:100px;">最終売上日</th>
                    </tr></thead><tbody>`;

                    compList.forEach((c, cIdx) => {{
                        const cQty = Number(c["数量"] || 0);
                        const kPrice = findCompanyKameiPrice(codeNorm, c["得意先名"]);
                        const fallbackUnitP = kPrice || Number(r["平均単価"] || minGedai || 0);
                        const rawAmt = Number(c["金額"] || 0);
                        const cAmt = (rawAmt >= cQty * unitCost && rawAmt > 0) ? rawAmt : Math.round(cQty * fallbackUnitP);
                        const effUnitP = cQty > 0 ? Math.round(cAmt / cQty) : fallbackUnitP;
                        const dispUnitP = kPrice ? `¥${{fmtNum(kPrice)}} <span style="font-size:11px;color:#78350f;">(登録値)</span>` : `¥${{fmtNum(effUnitP)}}`;

                        const cCostTotal = Math.round(cQty * unitCost);
                        const unitProf = effUnitP - unitCost;
                        const cProf = (cAmt > 0 && unitCost > 0) ? Math.max(0, Math.round(cAmt - cCostTotal)) : 0;
                        const cMargin = (cAmt > 0) ? ((cProf / cAmt) * 100).toFixed(1) + "%" : "0.0%";

                        sumQty += cQty;
                        sumSales += cAmt;
                        sumCost += cCostTotal;
                        sumProfit += cProf;
                        sumCnt += Number(c["伝票回数"] || 0);

                        sHtml += `<tr>
                            <td style="text-align:center; font-weight:bold;">${{cIdx + 1}}</td>
                            <td style="font-weight:bold; color:#1e293b;">${{c["得意先名"]}}</td>
                            <td style="text-align:right; font-weight:bold; color:#ea580c; background:#fff7ed; font-size:15px;">${{fmtNum(cQty)}} 個</td>
                            <td style="text-align:right; font-weight:bold;">${{c["構成比"]}}</td>
                            <td style="text-align:right; font-weight:bold; color:#b45309; background:#fffbeb;">${{dispUnitP}}</td>
                            <td style="text-align:right; background:#f0f9ff; color:#0369a1; font-weight:bold;">
                                ¥${{fmtNum(unitCost)}}<br><span style="font-size:11px; color:#64748b; font-weight:normal;">(計 ¥${{fmtNum(cCostTotal)}})</span>
                            </td>
                            <td style="text-align:right; font-weight:bold;">¥${{fmtNum(cAmt)}}</td>
                            <td style="text-align:right; font-weight:bold; color:#047857; background:#ecfdf5; font-size:15px;">
                                ¥${{fmtNum(cProf)}}<br><span style="font-size:11px; color:#065f46; font-weight:normal;">(1個 +¥${{fmtNum(unitProf)}})</span>
                            </td>
                            <td style="text-align:right; font-weight:bold; color:#047857; background:#d1fae5; font-size:15px;">${{cMargin}}</td>
                            <td style="text-align:center;">${{fmtNum(c["伝票回数"])}} 回</td>
                            <td style="text-align:center;">${{c["最終売上日"]}}</td>
                        </tr>`;
                    }});

                    const totMargin = (sumSales > 0) ? ((sumProfit / sumSales) * 100).toFixed(1) + "%" : "0.0%";
                    sHtml += `<tr style="background:#f1f5f9; border-top:2px solid #475569; font-weight:bold;">
                        <td colspan="2" style="text-align:center; color:#0f172a;">合計（${{compList.length}} 社）</td>
                        <td style="text-align:right; color:#ea580c; font-size:15px;">${{fmtNum(sumQty)}} 個</td>
                        <td style="text-align:right;">100%</td>
                        <td style="text-align:center;">—</td>
                        <td style="text-align:right; color:#0369a1;">計 ¥${{fmtNum(sumCost)}}</td>
                        <td style="text-align:right; font-size:15px;">¥${{fmtNum(sumSales)}}</td>
                        <td style="text-align:right; color:#047857; font-size:15.5px; background:#bbf7d0;">¥${{fmtNum(sumProfit)}}</td>
                        <td style="text-align:right; color:#047857; font-size:15.5px; background:#bbf7d0;">${{totMargin}}</td>
                        <td style="text-align:center;">${{fmtNum(sumCnt)}} 回</td>
                        <td style="text-align:center;">—</td>
                    </tr>`;
                    sHtml += `</tbody></table>`;
                }} else {{
                    sHtml += `<div class="diag-box" style="background:#f8fafc; border:1px solid #cbd5e1; color:#334155;">
                        <b>📊 期間（${{kiTitle}}）のyayoi_CSV売上明細は 0 件です。</b><br>
                        上の「期間切替」ボタンで他の期間に切り替えるか、下の『会社ごとの販売価格（加盟店価格）一覧』をご確認ください。
                    </div>`;
                }}

                const kCount = (kameiPriceMap[codeNorm] || []).length;
                sHtml += `<div class="sec-title" style="margin-top:18px;">💴 この商品の『会社ごとの販売価格（order_r 加盟店価格シート）』登録一覧（全 ${{kCount}} 件）</div>`;
                sHtml += buildKameiPriceTableHtml(codeNorm, unitCost, minGedai);

                if (detList.length > 0) {{
                    sHtml += `<div class="sec-title" style="margin-top:18px;">📄 売上伝票明細（最新60件）</div>`;
                    sHtml += `<table class="sub-table"><thead><tr>
                        <th style="width:90px;">期間区分</th>
                        <th style="width:105px;">売上日</th>
                        <th>得意先名</th>
                        <th style="width:85px;">数量</th>
                        <th style="width:95px;">販売単価</th>
                        <th style="width:95px;">原価(1個)</th>
                        <th style="width:110px;">売上金額</th>
                        <th style="width:110px;">利益額</th>
                        <th style="width:80px;">利益率</th>
                        <th style="width:80px;">担当者</th>
                    </tr></thead><tbody>`;
                    detList.forEach(d => {{
                        const dQty = Number(d["数量"] || 0);
                        const kp = findCompanyKameiPrice(codeNorm, d["得意先名"]);
                        const fbP = kp || Number(r["平均単価"] || minGedai || 0);
                        const rawDAmt = Number(d["金額"] || 0);
                        const dAmt = (rawDAmt >= dQty * unitCost && rawDAmt > 0) ? rawDAmt : Math.round(dQty * fbP);
                        const dUnitP = dQty > 0 ? Math.round(dAmt / dQty) : fbP;
                        const dProf = Math.max(0, Math.round(dAmt - (dQty * unitCost)));
                        const dMar = dAmt > 0 ? ((dProf / dAmt) * 100).toFixed(1) + "%" : "0.0%";
                        sHtml += `<tr>
                            <td style="text-align:center; font-weight:bold;">${{d["期"]}}</td>
                            <td style="text-align:center; font-weight:bold;">${{d["日付"]}}</td>
                            <td style="font-weight:bold;">${{d["得意先名"]}}</td>
                            <td style="text-align:right; font-weight:bold; color:#0284c7;">${{fmtNum(dQty)}}</td>
                            <td style="text-align:right;">¥${{fmtNum(dUnitP)}}</td>
                            <td style="text-align:right; color:#0369a1;">¥${{fmtNum(unitCost)}}</td>
                            <td style="text-align:right; font-weight:bold;">¥${{fmtNum(dAmt)}}</td>
                            <td style="text-align:right; font-weight:bold; color:#047857; background:#ecfdf5;">¥${{fmtNum(dProf)}}</td>
                            <td style="text-align:right; font-weight:bold; color:#047857;">${{dMar}}</td>
                            <td style="text-align:center;">${{d["担当者"]}}</td>
                        </tr>`;
                    }});
                    sHtml += `</tbody></table>`;
                }}
                document.getElementById("modal-yayoi-area").innerHTML = sHtml;
            }}
            // モードC：確保数のズレ診断モード
            else if (mode === 'zure') {{
                document.getElementById("modal-sub-title").innerHTML = `⚖️ 【${{code}}】確保数のズレ原因診断 ＆ 全ズレ商品リスト`;
                const rojiSec = Number(r["確保数_ロジザード"] || 0);
                const yayoiSec = Number(r["在庫確保_弥生"] || 0);
                const zure = Number(r["確保数のズレ"] || 0);

                let diagHtml = "";
                if (zure === 0) {{
                    diagHtml = `<div class="diag-box" style="background:#f0fdf4; border:2px solid #86efac; color:#166534;">
                        <b>✅ この商品（${{code}}）の確保数は完全に一致しています（ズレ: 0個）</b><br>
                        ロジザード確保数（<b>${{fmtNum(rojiSec)}}</b> 個） ＝ 在庫確保(弥生)（<b>${{fmtNum(yayoiSec)}}</b> 個）
                    </div>`;
                }} else if (zure < 0) {{
                    diagHtml = `<div class="diag-box" style="background:#fff1f2; border:2px solid #fda4af; color:#9f1239;">
                        <b>🚨 在庫確保(弥生)の方が 【${{fmtNum(Math.abs(zure))}} 個】 多い状態です（ズレ: ${{fmtNum(zure)}}）</b><br>
                        ・ロジザード側の確保数: <b>${{fmtNum(rojiSec)}}</b> 個 ／ 弥生販売側の確保合計: <b>${{fmtNum(yayoiSec)}}</b> 個<br>
                        👉 <b>主な原因</b>：弥生販売で受注伝票（確保）を登録したあと、<b>ロジザード側で確保在庫への移動（引当）がまだ行われていない</b>か、2027年等の先行受注分をロジザード側でまだ確保枠に移していない可能性があります。
                    </div>`;
                }} else {{
                    diagHtml = `<div class="diag-box" style="background:#fffbeb; border:2px solid #fcd34d; color:#92400e;">
                        <b>⚠️ ロジザード確保の方が 【＋${{fmtNum(zure)}} 個】 多い状態です（ズレ: +${{fmtNum(zure)}}）</b><br>
                        ・ロジザード側の確保数: <b>${{fmtNum(rojiSec)}}</b> 個 ／ 弥生販売側の確保合計: <b>${{fmtNum(yayoiSec)}}</b> 個<br>
                        👉 <b>主な原因</b>：弥生販売側ではすでに出荷売上済（またはキャンセル消込済）なのに、<b>ロジザード側の確保枠から外し忘れている</b>か、リンベル確保分（<b>${{fmtNum(r["リンベル確保数"] || 0)}}</b> 個）などがロジザード側に入っている可能性があります。
                    </div>`;
                }}

                const allZureRows = tableData.filter(item => Number(item["確保数のズレ"] || 0) !== 0)
                    .sort((a, b) => Math.abs(Number(b["確保数のズレ"])) - Math.abs(Number(a["確保数のズレ"])));

                diagHtml += `<div class="sec-title">⚖️ 現在「確保数のズレ」が発生している全商品一覧（全 ${{allZureRows.length}} 件 ／ 品番クリックで切替）</div>`;
                diagHtml += `<table class="sub-table"><thead><tr>
                    <th style="width:65px;">拠点</th>
                    <th style="width:130px;">品番</th>
                    <th>商品名</th>
                    <th style="width:115px;">ロジ確保数</th>
                    <th style="width:115px;">在庫確保(弥生)</th>
                    <th style="width:115px;">確保数のズレ</th>
                    <th>ズレの主な要因</th>
                </tr></thead><tbody>`;

                allZureRows.forEach(zr => {{
                    const zVal = Number(zr["確保数のズレ"]);
                    const isCurr = (zr["_origIdx"] === idx) ? "background:#fef9c3;" : "";
                    const reasonTxt = (zVal < 0) ? "🚨 弥生確保が多い（ロジ未引当/先行受注）" : "⚠️ ロジ確保が多い（ロジ解除漏れ/リンベル等）";
                    diagHtml += `<tr style="${{isCurr}}">
                        <td style="text-align:center;">${{zr["商品センター"]}}</td>
                        <td><span class="click-link" onclick="openModal(${{zr['_origIdx']}}, 'zure', '')">${{zr["品番"]}}</span></td>
                        <td style="font-weight:bold;">${{zr["商品区分"]}}</td>
                        <td style="text-align:right;">${{fmtNum(zr["確保数_ロジザード"])}}</td>
                        <td style="text-align:right; font-weight:bold; color:#0284c7;">${{fmtNum(zr["在庫確保_弥生"])}}</td>
                        <td style="text-align:right; font-weight:bold; color:#be123c; background:#ffe4e6; font-size:15px;">${{fmtNum(zVal)}}</td>
                        <td style="font-size:13px; font-weight:bold;">${{reasonTxt}}</td>
                    </tr>`;
                }});
                diagHtml += `</tbody></table>`;
                document.getElementById("modal-yayoi-area").innerHTML = diagHtml;
            }}
            // モードD：弥生確保リスト（会社ごとの加盟店販売価格も表示）
            else {{
                document.getElementById("modal-sub-title").innerHTML = "🏢 得意先別・弥生確保リスト（会社別の加盟店販売価格つき）";
                const targetList = yayoiMap[codeNorm] || [];

                if (targetList.length === 0) {{
                    document.getElementById("modal-yayoi-area").innerHTML = `<div style="padding:16px; font-size:14.5px; color:#64748b; background:#f8fafc;">該当する受注確保データはありません。</div>` +
                        `<div class="sec-title" style="margin-top:18px;">💴 この商品の『会社ごとの販売価格（order_r 加盟店価格）』登録一覧</div>` +
                        buildKameiPriceTableHtml(codeNorm, unitCost, minGedai);
                }} else {{
                    let yHtml = `<table class="sub-table"><thead><tr>
                        <th>受注番号</th>
                        <th>得意先名</th>
                        <th style="width:100px;">確保数量</th>
                        <th style="width:125px; background:#fef3c7; color:#b45309;">販売価格(加盟店価格)</th>
                        <th>受注日</th>
                        <th>納入期日</th>
                        <th>倉庫名</th>
                        <th>担当者</th>
                        <th>備考</th>
                    </tr></thead><tbody>`;
                    targetList.forEach(y => {{
                        const kp = findCompanyKameiPrice(codeNorm, y["得意先名"]);
                        const kpStr = kp ? `¥${{fmtNum(kp)}}` : (minGedai > 0 ? `¥${{fmtNum(minGedai)}} (最低下代)` : "—");
                        yHtml += `<tr>
                            <td style="text-align:center; font-weight:bold;">${{y["受注番号"] || ""}}</td>
                            <td style="font-weight:bold;">${{y["得意先名"]}}</td>
                            <td style="text-align:right; font-weight:bold; color:#0284c7; font-size:15.5px;">${{fmtNum(y["数量"])}}</td>
                            <td style="text-align:right; font-weight:bold; color:#b45309; background:#fffbeb;">${{kpStr}}</td>
                            <td style="text-align:center; font-weight:bold;">${{y["受注日"]}}</td>
                            <td style="text-align:center;">${{y["納入期日"]}}</td>
                            <td style="text-align:center; font-weight:bold;">${{y["倉庫名"] || ""}}</td>
                            <td style="text-align:center;">${{y["担当者名"] || ""}}</td>
                            <td>${{y["備考"]}}</td>
                        </tr>`;
                    }});
                    yHtml += `</tbody></table>`;
                    yHtml += `<div class="sec-title" style="margin-top:18px;">💴 この商品の『会社ごとの販売価格（order_r 加盟店価格）』登録一覧</div>`;
                    yHtml += buildKameiPriceTableHtml(codeNorm, unitCost, minGedai);
                    document.getElementById("modal-yayoi-area").innerHTML = yHtml;
                }}
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
            const cols = ["No", "商品センター", "品番", "商品区分", "バッジ表示", "74期", "75期", "76期", "77期", "今期出荷", "先週売数", "今週売数", "新原価", "売上高", "粗利額", "粗利率", "1課発注点", "2課発注点", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "確保数_ロジザード", "在庫確保_弥生", "確保数のズレ", "入荷後のフリー在庫", "発注検討フラグ", "在庫回転率", "リンベル確保数", "前回決定事項", "決定事項", "発注日"];
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
        "lot": max(int(r_m.get("入数", 1)), 1),
        "cost": int(r_m.get("新原価", 0))
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
        st.dataframe(inc[["品番", "商品区分", "新原価", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "発注日", "前回決定事項", "決定事項"]], hide_index=True, use_container_width=True, height=520)
    with ch2:
        st.subheader("📜 過去の会議シート 決定事項 全履歴")
        h_kw = st.text_input("品番・決定内容で検索", placeholder="例: 発注 / 様子見")
        hv = df_hist.sort_values(["会議日", "品番"], ascending=[False, True]) if not df_hist.empty else pd.DataFrame()
        if not hv.empty and h_kw:
            hv = hv[hv["品番"].str.contains(h_kw, case=False, na=False) | hv["決定事項"].str.contains(h_kw, case=False, na=False)]
        st.dataframe(hv, hide_index=True, use_container_width=True, height=450)

with tab5:
    st.subheader("🚀 営業用：大量注文・確保依頼の即判定シミュレーター（発注概算金額つき）")
    st.caption("営業担当者が大口注文や確保依頼を受けた際、下の表に『品番』と『今回の注文数』を入れる（またはExcelからコピペする）だけで、発注要否と推奨発注金額（新原価ベース）を0.1秒で一括判定します。")

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
                u_c = int(round(cost_map.get(c_n, 0)))
                sim_rows.append({
                    "品番": c_raw,
                    "商品名": item_name_map.get(c_n, "マスター未登録（特注/OEM枝番）"),
                    "今回の注文数": q_order,
                    "【判定結果】": "🚨 マスター未登録（発注担当へ即確認！）",
                    "現在のフリー在庫": 0,
                    "注文後の残り在庫": -q_order,
                    "発注点(安全在庫)": 0,
                    "入数": extra_meta.get("lot_map", {}).get(c_n, 1),
                    "推奨発注数(入数切上)": q_order,
                    "新原価": u_c,
                    "発注予定金額(原価)": q_order * u_c
                })
            else:
                info = master_sim_map[c_n]
                avail = info["free"]
                after_q = avail - q_order
                margin = after_q - info["point"]
                lot = info["lot"]
                u_c = info["cost"]
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
                    "推奨発注数(入数切上)": rec_q,
                    "新原価": u_c,
                    "発注予定金額(原価)": rec_q * u_c
                })
        if sim_rows:
            st.markdown("#### 📋 判定結果一覧")
            st.dataframe(pd.DataFrame(sim_rows), hide_index=True, use_container_width=True)

    st.divider()
    st.subheader(f"🚨 発注抜け防止アラート：弥生確保があるのに発注検討マスターに未登録の商品（{len(df_unreg)}件）")
    st.caption("『在庫確保.xlsx』に受注残があるにもかかわらず、発注検討マスターに品番が登録されていないため定期発注チェックから漏れている一覧です。")
    if df_unreg.empty:
        st.success("✅ 現在、マスター未登録の確保商品は0件です。抜け漏れはありません。")
    else:
        st.dataframe(df_unreg.sort_values("差引フリー在庫"), hide_index=True, use_container_width=True)
