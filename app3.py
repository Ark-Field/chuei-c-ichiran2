import io
import os
import json
import base64
import requests
import pandas as pd
from datetime import datetime, timedelta, date
import streamlit as st

# ReportLab関連のインポート（PDF出力用）
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

# ==========================================
# 1. ページ基本設定 ＆ 濃紺シックデザインCSS
# ==========================================
st.set_page_config(page_title="受注一覧表② 管理システム", layout="wide", initial_sidebar_state="collapsed")

dark_navy_theme_css = """
    <style>
    .stApp {
        background-color: #F8FAFC;
    }
    h1, h2, h3 {
        color: #0F172A !important;
        font-family: 'Helvetica Neue', Arial, sans-serif;
    }
    div.stButton > button:first-child {
        background-color: #1E3A8A;
        color: white;
        border-radius: 6px;
        border: none;
        font-weight: bold;
        padding: 0.5rem 1rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        transition: background-color 0.2s;
    }
    div.stButton > button:first-child:hover {
        background-color: #1E40AF;
        color: white;
    }
    div.stDownloadButton > button:first-child {
        background-color: #0F172A;
        color: white;
        border-radius: 6px;
        border: none;
        font-weight: bold;
    }
    div.stDownloadButton > button:first-child:hover {
        background-color: #1E293B;
        color: white;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #E2E8F0;
        border-radius: 6px 6px 0px 0px;
        color: #334155;
        font-weight: bold;
        padding: 10px 18px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E3A8A !important;
        color: white !important;
    }
    .streamlit-expanderHeader {
        background-color: #F1F5F9;
        border-radius: 6px;
        color: #1E293B;
        font-weight: bold;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
"""
st.markdown(dark_navy_theme_css, unsafe_allow_html=True)

st.title("📊 受注一覧表②（判型グループ・納期/顧客順管理）")

# ==========================================
# 2. 日本語フォント登録処理（ReportLab用）
# ==========================================
def setup_japanese_font():
    local_font_path = "NotoSansJP-Regular.ttf"
    if os.path.exists(local_font_path):
        try:
            pdfmetrics.registerFont(TTFont("JPFont", local_font_path))
            return "JPFont"
        except Exception as e:
            print(f"同梱フォント登録エラー: {e}")

    win_font = os.path.join(
        os.environ.get("SystemRoot", "C:\\Windows"), "Fonts", "msgothic.ttc"
    )
    if os.path.exists(win_font):
        try:
            pdfmetrics.registerFont(TTFont("JPFont", win_font, subfontIndex=0))
            return "JPFont"
        except Exception as e:
            print(f"Windowsフォント登録エラー: {e}")

    return "Helvetica"

FONT_NAME = setup_japanese_font()

# ==========================================
# 3. Secrets設定
# ==========================================
if "pocket" in st.secrets:
    POCKET_API_KEY = st.secrets["pocket"].get("api_key", "XZtsg3OrDFOh0VayOI83j84Oc3YIzm7oujbI0T08")
    POCKET_BASE_URL = st.secrets["pocket"].get("base_url", "https://app070.at-pocket.com/arkf05_bb")
else:
    POCKET_API_KEY = "XZtsg3OrDFOh0VayOI83j84Oc3YIzm7oujbI0T08"
    POCKET_BASE_URL = "https://app070.at-pocket.com/arkf05_bb"

APP_ID = 34  # @pocket対象アプリID

# ==========================================
# 4. 便利関数
# ==========================================
def clean_val(val):
    if val is None:
        return ""
    if isinstance(val, (list, dict)):
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    if s.lower() in ["none", "nan", "nat", "null"]:
        return ""
    return s

def parse_date_safe(val):
    s = clean_val(val)
    if not s:
        return None
    try:
        if isinstance(val, (date, datetime)):
            return val.date() if isinstance(val, datetime) else val
        return datetime.strptime(s[:10].replace("/", "-"), "%Y-%m-%d").date()
    except ValueError:
        return None

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    output.seek(0)
    return output

# --- 受注一覧表②のPDF生成（判型グループ別・納期/顧客順・備考2段） ---
def generate_order_pdf_v2(df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A3), leftMargin=15, rightMargin=15, topMargin=15, bottomMargin=15)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Normal"], fontName=FONT_NAME, fontSize=15, leading=17, textColor=colors.HexColor("#1E3A8A"))
    group_header_style = ParagraphStyle("GroupHeader", parent=styles["Normal"], fontName=FONT_NAME, fontSize=11, leading=14, textColor=colors.HexColor("#1E3A8A"), spaceBefore=10, spaceAfter=4)
    header_style = ParagraphStyle("Header", parent=styles["Normal"], fontName=FONT_NAME, fontSize=7.5, leading=9, textColor=colors.white, alignment=1)
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontName=FONT_NAME, fontSize=7, leading=8.5, textColor=colors.HexColor("#0F172A"))

    elements = [
        Paragraph("<b>受注一覧表②（判型グループ別・納期/顧客順・備考2段）</b>", title_style),
        Spacer(1, 4),
        Paragraph(f"<font size=8 color='#64748B'><b>発行日: {date.today().strftime('%Y年%m月%d日')}</b></font>", styles["Normal"]),
        Spacer(1, 6),
    ]

    top_cols = [
        "受注番号", "納期", "出版社名", "書名", "刷数", "回数", "単価", "製本数", "製品預かり", "総頁数",
        "32P", "16P", "8P", "4P", "2P", "印刷所", "折本日", "折", "無線日", "三方日", "仕上日", "入品日",
        "表紙日付", "カード日付", "カバー日付", "帯日付"
    ]
    available_top_cols = [c for c in top_cols if c in df.columns]

    width_mapping = {
        "受注番号": 52, "納期": 58, "出版社名": 72, "書名": 120, "刷数": 28, "回数": 28,
        "単価": 38, "製本数": 45, "製品預かり": 45, "総頁数": 42, "32P": 28, "16P": 28,
        "8P": 28, "4P": 28, "2P": 28, "印刷所": 48, "折本日": 58, "折": 28,
        "無線日": 58, "三方日": 58, "仕上日": 58, "入品日": 58, "表紙日付": 58,
        "カード日付": 58, "カバー日付": 58, "帯日付": 58
    }
    col_widths = [width_mapping.get(c, 45) for c in available_top_cols]

    # 判型ごとにグループ化してテーブルを作成
    group_col = "判型"
    groups = df[group_col].unique() if group_col in df.columns else [""]

    for idx, g_val in enumerate(groups):
        if idx > 0:
            elements.append(PageBreak())
        
        g_name_display = g_val if g_val else "（判型未設定）"
        elements.append(Paragraph(f"<b>■ 判型: {g_name_display}</b>", group_header_style))
        
        sub_df = df[df[group_col] == g_val] if g_val else df

        table_data = []
        header_row = [Paragraph(f"<b>{c}</b>", header_style) for c in available_top_cols]
        table_data.append(header_row)

        for _, row in sub_df.iterrows():
            row_cells_1 = [Paragraph(f"<b>{str(row[c]) if pd.notna(row[c]) else ''}</b>", cell_style) for c in available_top_cols]
            table_data.append(row_cells_1)
            
            memo_val = str(row["備考"]) if "備考" in df.columns and pd.notna(row["備考"]) else ""
            memo_cell = Paragraph(f"<b>備考: {memo_val}</b>", cell_style)
            row_cells_2 = [memo_cell] + [Paragraph("", cell_style) for _ in range(len(available_top_cols) - 1)]
            table_data.append(row_cells_2)

        t = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        t_style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#64748B")),
        ]

        row_count = len(sub_df)
        for i in range(row_count):
            memo_row_idx = 1 + (i * 2) + 1
            t_style.append(("SPAN", (0, memo_row_idx), (-1, memo_row_idx)))
            t_style.append(("BACKGROUND", (0, memo_row_idx), (-1, memo_row_idx), colors.HexColor("#F1F5F9")))

        t.setStyle(TableStyle(t_style))
        elements.append(t)
        elements.append(Spacer(1, 8))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 5. データ取得関数
# ==========================================
query_params = st.query_params
url_order_no = query_params.get("ord", None)

def fetch_pocket_records(start_date=None, end_date=None, target_order_no=None, target_publisher=None, date_field="field-52"):
    url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records"
    headers = {"X-At-Pocket-API-Key": POCKET_API_KEY, "Accept": "application/json"}
    params = {}
    
    if target_order_no:
        params["query"] = f'field-2 = "{target_order_no}"'
    else:
        query_parts = []
        if start_date and end_date:
            start_str = start_date.strftime("%Y-%m-%d")
            end_str = end_date.strftime("%Y-%m-%d")
            query_parts.append(f'{date_field} >= "{start_str}" and {date_field} <= "{end_str}"')
        if target_publisher:
            query_parts.append(f'field-8 like "{target_publisher}"')
        if query_parts:
            params["query"] = " and ".join(query_parts)
    
    try:
        response = requests.get(url, headers=headers, params=params, timeout=20)
        if response.status_code != 200 and "query" in params:
            response = requests.get(url, headers=headers, timeout=20)
        
        if response.status_code == 200:
            result = response.json()
            raw_items = result.get("records", result.get("Data", result.get("Records", [])))
            
            parsed_records = []
            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                rec_body = item.get("record", {})
                if not isinstance(rec_body, dict):
                    rec_body = {}
                
                merged_body = {**item, **rec_body}
                merged_body["_recordId"] = str(
                    item.get("Id") or item.get("recordId") or rec_body.get("Id") or rec_body.get("recordId") or ""
                ).strip()
                parsed_records.append(merged_body)
            
            if start_date and end_date:
                start_str = start_date.strftime("%Y-%m-%d")
                end_str = end_date.strftime("%Y-%m-%d")
                filtered = []
                for r in parsed_records:
                    raw_d = clean_val(r.get(date_field, ""))
                    if not raw_d:
                        continue
                    d_str = raw_d.replace("/", "-")[:10]
                    if len(d_str) == 10 and start_str <= d_str <= end_str:
                        filtered.append(r)
                parsed_records = filtered

            return parsed_records
        else:
            st.error(f"❌ APIエラー ({response.status_code}): {response.text}")
    except Exception as e:
        st.error(f"🚨 通信例外エラー: {e}")
    return []

if "api_records" not in st.session_state or url_order_no:
    if url_order_no:
        with st.spinner(f"DBから受注番号: {url_order_no} のデータを取得中..."):
            st.session_state["api_records"] = fetch_pocket_records(target_order_no=url_order_no)
    else:
        default_start = date.today()
        default_end = default_start + timedelta(days=14)
        st.session_state["api_records"] = fetch_pocket_records(start_date=default_start, end_date=default_end, date_field="field-52")

# ==========================================
# 6. メイン画面
# ==========================================
st.header("📋 受注一覧表②（判型グループ・納期/顧客順管理）")

with st.expander("🔍 検索条件の変更・絞り込みはこちら", expanded=bool(not url_order_no)):
    search_date_type = st.radio("📅 期間検索の基準を選択", ["納期（field-52）", "受注日（field-1）"], horizontal=True)
    target_date_field = "field-52" if "納期" in search_date_type else "field-1"

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        t1_start = st.date_input("期間（開始）", value=date.today(), key="v2_s")
    with col_s2:
        t1_end = st.date_input("期間（終了）", value=date.today() + timedelta(days=14), key="v2_e")
    
    publisher_input = st.text_input("📖 出版社名で絞り込み", value="", key="v2_pub")
    search_clicked = st.button("🔄 検索実行", type="primary", key="v2_btn")

if search_clicked:
    st.query_params.clear()
    with st.spinner("@ポケットからデータを取得中..."):
        st.session_state["api_records"] = fetch_pocket_records(
            start_date=t1_start, 
            end_date=t1_end, 
            target_publisher=publisher_input if publisher_input else None,
            date_field=target_date_field
        )
    st.rerun()

api_records = st.session_state.get("api_records", None)

if api_records:
    st.success(f"✨ 対象レコード **{len(api_records)} 件** を取得しました。")
    
    flat_data = []
    for record_idx, r in enumerate(api_records):
        h_ord_no = clean_val(r.get("field-2", ""))
        h_ord_date = parse_date_safe(r.get("field-1"))
        h_due = parse_date_safe(r.get("field-52"))
        h_pub = clean_val(r.get("field-8", ""))
        h_title = clean_val(r.get("field-13", ""))
        h_format = clean_val(r.get("field-23", ""))  # 判型
        h_qty = clean_val(r.get("field-4", ""))
        h_price = clean_val(r.get("field-5", ""))
        h_cust_keeping = clean_val(r.get("field-81", ""))
        h_total_pages = clean_val(r.get("field-82", ""))
        h_p32 = clean_val(r.get("field-30", ""))
        h_p16 = clean_val(r.get("field-31", ""))
        h_p8 = clean_val(r.get("field-32", ""))
        h_p4 = clean_val(r.get("field-33", ""))
        h_p2 = clean_val(r.get("field-34", ""))
        h_printer = clean_val(r.get("field-14", ""))
        h_folding = clean_val(r.get("field-16", ""))
        h_memo = clean_val(r.get("field-45", ""))
        
        h_print_count = clean_val(r.get("field-28", ""))
        h_times_count = clean_val(r.get("field-96", ""))
        
        sub_90 = r.get("field-90", [])
        sub_91 = r.get("field-91", [])
        
        item_90 = sub_90[0] if isinstance(sub_90, list) and len(sub_90) > 0 else {}
        rec_90 = item_90.get("record", item_90) if isinstance(item_90, dict) else {}
        
        item_91 = sub_91[0] if isinstance(sub_91, list) and len(sub_91) > 0 else {}
        rec_91 = item_91.get("record", item_91) if isinstance(item_91, dict) else {}

        flat_data.append({
            "更新": False,
            "_record_idx": record_idx,
            "判型": h_format,
            "納期": h_due,
            "出版社名": h_pub,
            "受注番号": h_ord_no,
            "書名": h_title,
            "刷数": h_print_count,
            "回数": h_times_count,
            "単価": h_price,
            "製本数": h_qty,
            "製品預かり": h_cust_keeping,
            "総頁数": h_total_pages,
            "32P": h_p32,
            "16P": h_p16,
            "8P": h_p8,
            "4P": h_p4,
            "2P": h_p2,
            "印刷所": h_printer,
            "折本日": parse_date_safe(rec_90.get("field-90_3")),
            "折": h_folding,
            "無線日": parse_date_safe(rec_90.get("field-90_4")),
            "三方日": parse_date_safe(rec_90.get("field-90_5")),
            "仕上日": parse_date_safe(rec_90.get("field-90_6")),
            "入品日": parse_date_safe(rec_90.get("field-90_10")),
            "表紙日付": parse_date_safe(rec_91.get("field-91_1")),
            "カード日付": parse_date_safe(rec_91.get("field-91_9")),
            "カバー日付": parse_date_safe(rec_91.get("field-91_5")),
            "帯日付": parse_date_safe(rec_91.get("field-91_13")),
            "備考": h_memo
        })

    df_v2 = pd.DataFrame(flat_data)

    # 判型順 ➔ 納期順 ➔ 出版社名順 でソート
    sort_keys = ["判型", "納期", "出版社名"]
    df_v2 = df_v2.sort_values(by=sort_keys, ascending=[True, True, True], na_position="last").reset_index(drop=True)

    edited_v2_df = st.data_editor(
        df_v2,
        num_rows="fixed",
        key="grid_v2",
        use_container_width=True,
        column_config={
            "更新": st.column_config.CheckboxColumn("更新", required=True),
            "_record_idx": None,
            "判型": st.column_config.TextColumn("判型", disabled=True),
            "納期": st.column_config.DateColumn("納期", format="YYYY-MM-DD", disabled=True),
            "出版社名": st.column_config.TextColumn("出版社名", disabled=True),
            "受注番号": st.column_config.TextColumn("受注番号", disabled=True),
            "書名": st.column_config.TextColumn("書名", disabled=True),
            "刷数": st.column_config.TextColumn("刷数"),
            "回数": st.column_config.TextColumn("回数"),
            "単価": st.column_config.TextColumn("単価", disabled=True),
            "製本数": st.column_config.TextColumn("製本数", disabled=True),
            "製品預かり": st.column_config.TextColumn("製品預かり", disabled=True),
            "総頁数": st.column_config.TextColumn("総頁数", disabled=True),
            "32P": st.column_config.TextColumn("32P", disabled=True),
            "16P": st.column_config.TextColumn("16P", disabled=True),
            "8P": st.column_config.TextColumn("8P", disabled=True),
            "4P": st.column_config.TextColumn("4P", disabled=True),
            "2P": st.column_config.TextColumn("2P", disabled=True),
            "印刷所": st.column_config.TextColumn("印刷所", disabled=True),
            "折本日": st.column_config.DateColumn("折本日", format="YYYY-MM-DD"),
            "折": st.column_config.TextColumn("折", disabled=True),
            "無線日": st.column_config.DateColumn("無線日", format="YYYY-MM-DD"),
            "三方日": st.column_config.DateColumn("三方日", format="YYYY-MM-DD"),
            "仕上日": st.column_config.DateColumn("仕上日", format="YYYY-MM-DD"),
            "入品日": st.column_config.DateColumn("入品日", format="YYYY-MM-DD"),
            "表紙日付": st.column_config.DateColumn("表紙日付", format="YYYY-MM-DD"),
            "カード日付": st.column_config.DateColumn("カード日付", format="YYYY-MM-DD"),
            "カバー日付": st.column_config.DateColumn("カバー日付", format="YYYY-MM-DD"),
            "帯日付": st.column_config.DateColumn("帯日付", format="YYYY-MM-DD"),
            "備考": st.column_config.TextColumn("備考")
        }
    )

    st.markdown("---")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        excel_bytes = to_excel_bytes(edited_v2_df)
        st.download_button(
            label="📥 受注一覧表②を Excel でダウンロード",
            data=excel_bytes,
            file_name=f"受注一覧表2_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_d2:
        pdf_bytes = generate_order_pdf_v2(edited_v2_df)
        st.download_button(
            label="🖨️ 受注一覧表②を A3 PDF でダウンロード",
            data=pdf_bytes,
            file_name=f"受注一覧表2_{date.today().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )
else:
    st.warning("⚠️ 対象データが見つかりませんでした。")