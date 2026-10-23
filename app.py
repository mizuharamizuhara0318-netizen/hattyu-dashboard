import streamlit as st
import pandas as pd
import openpyxl
import datetime
import os

st.set_page_config(page_title="発注検討表（Web管理版）", layout="wide")

st.markdown("""
<style>
    .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 100% !important; }
    th { background-color: #f2f2f2 !important; color: #000 !important; font-weight: bold !important; border: 1px solid #ccc !important; white-space: normal !important; text-align: center !important; }
    td { border: 1px solid #ccc !important; font-family: "Meiryo", sans-serif; font-size: 13px; }
    /* テーブル全体の表示を最適化 */
    div[data-testid="stDataEditor"] { width: 100% !important; }
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
SAVE_CSV = os.path.join(BASE_DIR, "meeting_decisions_v8.csv")

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
def load_yayoi_data(file_path):
    if not os.path.exists(file_path):
        return pd.DataFrame()
    df = pd.read_excel(file_path, sheet_name=0, header=4)
    df = df[df["商品コード"].notna()].copy()
    df["商品コード"] = df["商品コード"].astype(str).str.strip()
    df["得意先名"] = df["得意先名"].apply(lambda x: clean_str(x) if clean_str(x) else "不明")
    df["受注残数"] = pd.to_numeric(df["受注残数"], errors="coerce").fillna(0).astype(int)
    df["数量"] = pd.to_numeric(df["数量"], errors="coerce").fillna(0).astype(int)
    df["受注日"] = pd.to_datetime(df["受注日"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    df["納入期日"] = pd.to_datetime(df["納入期日"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
    df["備考"] = df["備考"].apply(clean_str)
    df["倉庫名"] = df["倉庫名"].apply(clean_str)
    df["商品名/摘要"] = df["商品名/摘要"].apply(clean_str)
    return df

@st.cache_data
def load_excel_master(file_path):
    if not os.path.exists(file_path):
        return pd.DataFrame(), pd.DataFrame()
    
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    target_sheet = "マスター (新) (3)" if "マスター (新) (3)" in wb.sheetnames else "80828"
    ws = wb[target_sheet]
    rows = safe_get_rows(ws)
    
    dec_80828 = {}
    prev_80828 = {}
    if "80828" in wb.sheetnames:
        for r in safe_get_rows(wb["80828"])[1:]:
            if len(r) > 26 and r[2]:
                code_k = str(r[2]).strip()
                if r[25]: prev_80828[code_k] = clean_str(r[25])
                if r[26]: dec_80828[code_k] = clean_str(r[26])

    items = []
    for r in rows[1:]:
        if len(r) < 27 or r[2] is None:
            continue
        hinban = clean_str(r[2])
        if not hinban:
            continue
            
        na_flag = is_na_val(r[14]) or is_na_val(r[25]) or clean_str(r[12]) != ""
        curr_dec = clean_str(r[29]) if len(r) > 29 else ""
        if not curr_dec and hinban in dec_80828:
            curr_dec = dec_80828[hinban]
            
        ndate = clean_str(r[18]) if len(r) > 18 else ""

        items.append({
            "No": clean_str(r[0]),
            "商品センター": clean_str(r[1]),
            "品番": hinban,
            "商品区分": clean_str(r[3]),
            "1課発注点": int(round(to_num(r[14]), 0)),
            "2課発注点": clean_str(r[15]),
            "在庫数_ロジザード": int(to_num(r[16])),
            "入荷予定数_ロジザード": int(to_num(r[17])),
            "入荷予定日_ロジザード": ndate,
            "確保数_ロジザード": int(to_num(r[19])),
            "在庫確保_弥生": int(to_num(r[20])),
            "在庫確保_丸藤": int(to_num(r[21])),
            "確保数のズレ": int(to_num(r[22])),
            "入荷後のフリー在庫": int(to_num(r[24])),
            "発注検討フラグ": int(round(to_num(r[25]), 0)),
            "在庫回転率": round(to_num(r[26]) * 100, 0),
            "リンベル確保数": clean_str(r[27]) if len(r) > 27 else "",
            "決定事項": curr_dec,
            "発注日": "",
            "72期": int(to_num(r[5])),
            "73期": int(to_num(r[6])),
            "74期": int(to_num(r[7])),
            "75期": int(to_num(r[9])),
            "76期": int(to_num(r[10])),
            "77期": int(to_num(r[11])),
            "NA行": na_flag
        })
    df_main = pd.DataFrame(items)

    date_sheets = [s for s in wb.sheetnames if s.isdigit() and len(s) == 5]
    hist_list = []
    for s in date_sheets:
        ws_d = wb[s]
        r_d = safe_get_rows(ws_d)
        if not r_d: continue
        header = r_d[0]
        h_col = next((i for i, c in enumerate(header) if c == "品番"), 2)
        k_cols = [i for i, c in enumerate(header) if c and "決定事項" in str(c) and "前回" not in str(c)]
        s_clean = s.lstrip("0")
        rest = s_clean[1:] if s_clean.startswith("8") else s_clean
        dt_str = f"2026-{rest[:2]}-{rest[2:]}" if len(rest) == 4 else (f"2026-0{rest[0]}-{rest[1:]}" if len(rest) == 3 else s)
        for r in r_d[1:]:
            if len(r) <= h_col or not r[h_col]: continue
            c_val = clean_str(r[h_col])
            for kc in k_cols:
                if kc < len(r) and clean_str(r[kc]) not in ("", "0"):
                    hist_list.append({"会議日": dt_str, "シート": s, "品番": c_val, "決定事項": clean_str(r[kc])})
    df_hist = pd.DataFrame(hist_list).drop_duplicates(subset=["会議日", "品番", "決定事項"])
    return df_main, df_hist

df_yayoi = load_yayoi_data(YAYOI_EXCEL)
df_main, df_hist = load_excel_master(MAIN_EXCEL)

if df_main.empty:
    st.error(f"⚠️ メインの『発注検討表』ファイルが見つかりません。(探索パス: {MAIN_EXCEL})")
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

if "master_df" not in st.session_state:
    st.session_state.master_df = df_main

m_df = st.session_state.master_df

st.markdown("### 📊 発注検討表（直接編集 ＆ 弥生確保連動システム）")

tab1, tab2, tab3 = st.tabs([
    "📋 1. 発注検討表（直接入力・詳細）",
    "🏢 2. 会社（得意先）別 確保・注文検索",
    "🕒 3. 入荷予定・過去の発注決定履歴"
])

with tab1:
    f_col1, f_col2, f_col3 = st.columns([2, 2, 3])
    filter_mode = f_col1.selectbox(
        "表示絞り込み",
        ["すべて表示（Excelと同じ並び）", "⚠️発注検討フラグがマイナスのみ", "確保数のズレがある商品", "入荷予定がある商品", "弥生確保がある商品"],
        index=0
    )
    center_opt = ["すべて"] + sorted([x for x in m_df["商品センター"].unique() if x])
    sel_center = f_col2.selectbox("商品センター", center_opt)
    search_kw = f_col3.text_input("🔍 品番・商品名・会社名で検索", placeholder="例: AHL-2203 / スタンドライト / 丸藤")

    v_df = m_df.copy()
    if filter_mode == "⚠️発注検討フラグがマイナスのみ":
        v_df = v_df[(v_df["発注検討フラグ"] < 0) & (~v_df["NA行"])]
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
            v_df["決定事項"].str.contains(search_kw, case=False, na=False)
        ]

    excel_cols = [
        "No", "商品センター", "品番", "商品区分", "1課発注点", "2課発注点",
        "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード",
        "確保数_ロジザード", "在庫確保_弥生", "在庫確保_丸藤", "確保数のズレ",
        "入荷後のフリー在庫", "発注検討フラグ", "在庫回転率", "リンベル確保数",
        "決定事項", "発注日"
    ]
    excel_cols = list(dict.fromkeys(excel_cols))

    # 数字列の幅を統一し、カンマ区切り（format="%d"）を設定
    edited_df = st.data_editor(
        v_df[excel_cols],
        column_config={
            "No": st.column_config.NumberColumn("No", width="small", format="%d"),
            "商品センター": st.column_config.TextColumn("商品センター", width="small"),
            "品番": st.column_config.TextColumn("品番", width="medium"),
            "商品区分": st.column_config.TextColumn("商品区分", width="large"),
            "1課発注点": st.column_config.NumberColumn("1課発注点", width="small", format="%d"),
            "2課発注点": st.column_config.TextColumn("2課発注点", width="small"),
            "在庫数_ロジザード": st.column_config.NumberColumn("在庫数 ロジザード", width="small", format="%d"),
            "入荷予定数_ロジザード": st.column_config.NumberColumn("入荷予定数 ロジザード", width="small", format="%d"),
            "入荷予定日_ロジザード": st.column_config.TextColumn("📝入荷予定日", width="medium"),
            "確保数_ロジザード": st.column_config.NumberColumn("確保数 ロジザード", width="small", format="%d"),
            "在庫確保_弥生": st.column_config.NumberColumn("在庫確保 弥生", width="small", format="%d"),
            "在庫確保_丸藤": st.column_config.NumberColumn("在庫確保 丸藤", width="small", format="%d"),
            "確保数のズレ": st.column_config.NumberColumn("確保数のズレ", width="small", format="%d"),
            "入荷後のフリー在庫": st.column_config.NumberColumn("入荷後 フリー在庫", width="small", format="%d"),
            "発注検討フラグ": st.column_config.NumberColumn("発注検討フラグ", width="small", format="%d"),
            "在庫回転率": st.column_config.NumberColumn("在庫回転率(%)", width="small", format="%d%%"),
            "リンベル確保数": st.column_config.TextColumn("リンベル確保数", width="small"),
            "決定事項": st.column_config.TextColumn("📝決定事項", width="large"),
            "発注日": st.column_config.TextColumn("📝発注日", width="small"),
        },
        disabled=[c for c in excel_cols if c not in ("入荷予定日_ロジザード", "決定事項", "発注日")],
        hide_index=True,
        use_container_width=True,
        height=550
    )

    if st.button("💾 編集した内容を保存する", type="primary"):
        for _, r in edited_df.iterrows():
            c = clean_str(r["品番"])
            if c:
                st.session_state.master_df.loc[st.session_state.master_df["品番"] == c, "入荷予定日_ロジザード"] = clean_str(r.get("入荷予定日_ロジザード", ""))
                st.session_state.master_df.loc[st.session_state.master_df["品番"] == c, "決定事項"] = clean_str(r.get("決定事項", ""))
                st.session_state.master_df.loc[st.session_state.master_df["品番"] == c, "発注日"] = clean_str(r.get("発注日", ""))
        st.session_state.master_df[["品番", "入荷予定日_ロジザード", "決定事項", "発注日"]].to_csv(SAVE_CSV, index=False, encoding="utf-8-sig")
        st.success("✅ 保存しました！")

    down_col1, _ = st.columns([2, 4])
    csv_data = m_df[excel_cols].to_csv(index=False, encoding="utf-8-sig")
    down_col1.download_button(
        label="📥 現在の発注検討表（入力内容含む）をCSVダウンロード",
        data=csv_data,
        file_name=f"発注検討表_出力_{datetime.date.today()}.csv",
        mime="text/csv",
        type="secondary"
    )

    st.divider()

    st.subheader("🔎 品番別 詳細ドリルダウン（どの会社が何個確保しているか・過去履歴）")
    all_codes = v_df["品番"].tolist() if not v_df.empty else m_df["品番"].tolist()
    
    sel_code = st.selectbox(
        "📌 確認・入力する【品番】を選択してください",
        options=all_codes,
        format_func=lambda x: f"{x} ｜ {m_df.loc[m_df['品番']==x, '商品区分'].values[0]} （在庫:{m_df.loc[m_df['品番']==x, '在庫数_ロジザード'].values[0]:,} / 確保:{m_df.loc[m_df['品番']==x, '確保数_ロジザード'].values[0]:,} / 発注検討:{m_df.loc[m_df['品番']==x, '発注検討フラグ'].values[0]:,}）"
    )

    if sel_code:
        row_info = m_df[m_df["品番"] == sel_code].iloc[0]
        col_left, col_right = st.columns([3, 2])
        with col_left:
            st.markdown(f"#### 🏢 『{sel_code}』の得意先別・弥生確保内訳")
            if not df_yayoi.empty:
                sub_y = df_yayoi[df_yayoi["商品コード"] == sel_code].copy()
                if not sub_y.empty:
                    sub_y = sub_y.fillna("")
                    st.dataframe(sub_y[["得意先名", "受注残数", "数量", "受注日", "納入期日", "倉庫名", "備考", "担当者名"]], hide_index=True, use_container_width=True)
                else:
                    st.info(f"『{sel_code}』の弥生確保明細はありません。")
        with col_right:
            st.markdown(f"#### 📜 『{sel_code}』の過去決定履歴")
            if not df_hist.empty:
                sub_h = df_hist[df_hist["品番"] == sel_code].sort_values("会議日", ascending=False)
                if not sub_h.empty:
                    st.dataframe(sub_h[["会議日", "決定事項"]], hide_index=True, use_container_width=True)
                else:
                    st.caption("過去の会議決定履歴なし")

with tab2:
    st.subheader("🏢 会社（得意先）別 確保・注文検索 ＆ 月別納期スケジュール")
    if df_yayoi.empty:
        st.warning("『弥生確保リスト.xlsx』が見つかりません。")
    else:
        comp_totals = df_yayoi.groupby("得意先名")["受注残数"].sum().sort_values(ascending=False)
        comp_opts = ["すべての会社を表示"] + [f"{c} （確保合計: {v:,} 個）" for c, v in comp_totals.items()]
        c1, c2 = st.columns(2)
        sel_comp = c1.selectbox("会社名（得意先）を選択", comp_opts)
        kw_comp = c2.text_input("会社名・品番・備考で絞り込み", placeholder="例: 丸藤 / トップ産業 / 11月")

        yv = df_yayoi.copy()
        if sel_comp != "すべての会社を表示":
            cname = sel_comp.split(" （確保合計:")[0]
            yv = yv[yv["得意先名"] == cname]
        if kw_comp:
            yv = yv[yv["得意先名"].str.contains(kw_comp, case=False, na=False) | yv["商品コード"].str.contains(kw_comp, case=False, na=False) | yv["備考"].str.contains(kw_comp, case=False, na=False)]

        yv_merged = pd.merge(yv, m_df[["品番", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "入荷後のフリー在庫", "発注検討フラグ", "決定事項"]], left_on="商品コード", right_on="品番", how="left")
        yv_merged = yv_merged.fillna("")
        for col in yv_merged.select_dtypes(include=['object']).columns:
            yv_merged[col] = yv_merged[col].apply(lambda x: "" if is_na_val(x) else str(x))

        if not yv_merged.empty and "納入期日" in yv_merged.columns:
            valid_dates = yv_merged[yv_merged["納入期日"] != ""]["納入期日"]
            if not valid_dates.empty:
                yv_merged["納入月"] = pd.to_datetime(valid_dates, errors="coerce").dt.strftime("%Y年%m月")
                monthly_sch = yv_merged[yv_merged["納入月"].notna()].groupby("納入月")["受注残数"].sum()
                if not monthly_sch.empty:
                    st.markdown("##### 📈 選択した会社のゲージ・月別 納入スケジュール（受注残数合計）")
                    st.bar_chart(monthly_sch)

        st.dataframe(yv_merged[["得意先名", "商品コード", "商品名/摘要", "受注残数", "受注日", "納入期日", "備考", "倉庫名", "入荷予定日_ロジザード", "決定事項"]], hide_index=True, use_container_width=True, height=400)

with tab3:
    ch1, ch2 = st.columns(2)
    with ch1:
        st.subheader("🚢 入荷予定あり・発注済みの商品一覧")
        inc = m_df[(m_df["入荷予定数_ロジザード"] > 0) | (m_df["発注日"] != "") | (m_df["入荷予定日_ロジザード"] != "")].copy()
        st.dataframe(inc[["品番", "商品区分", "在庫数_ロジザード", "入荷予定数_ロジザード", "入荷予定日_ロジザード", "発注日", "決定事項"]], hide_index=True, use_container_width=True, height=500)
    with ch2:
        st.subheader("📜 過去28回の会議シート 決定事項 全履歴")
        h_kw = st.text_input("品番・決定内容で検索", placeholder="例: 発注 / 様子見")
        hv = df_hist.sort_values(["会議日", "品番"], ascending=[False, True]) if not df_hist.empty else pd.DataFrame()
        if not hv.empty and h_kw:
            hv = hv[hv["品番"].str.contains(h_kw, case=False, na=False) | hv["決定事項"].str.contains(h_kw, case=False, na=False)]
        st.dataframe(hv, hide_index=True, use_container_width=True, height=450)