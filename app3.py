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

st.title("📊 受注一覧表②（@pocket保存機能・赤網掛け対応）")

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

def date_to_str(val):
    if pd.isna(val) or not val:
        return ""
    if isinstance(val, (date, datetime)):
        return val.strftime("%Y-%m-%d")
    parsed = parse_date_safe(val)
    return parsed.strftime("%Y-%m-%d") if parsed else ""

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
    output.seek(0)
    return output

# --- @pocket データ更新用関数 ---
def update_pocket_record(record_id, update_data):
    if not record_id:
        return False, "レコードIDが不正です"
    url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records/{record_id}"
    headers = {
        "X-At-Pocket-API-Key": POCKET_API_KEY,
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    payload = {
        "apiVersion": 1.1,
        "record": update_data
    }
    try:
        response = requests.put(url, headers=headers, json=payload, timeout=20)
        if response.status_code in [200, 204]:
            return True, "成功"
        else:
            return False, f"HTTP {response.status_code}: {response.text}"
    except Exception as e:
        return False, str(e)

# --- 受注一覧表②のPDF生成（該当行全体を赤網掛け） ---
def generate_order_pdf_v2(df, group_type="publisher"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A3), leftMargin=15, rightMargin=15, topMargin=15, bottomMargin=15)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Normal"], fontName=FONT_NAME, fontSize=15, leading=17, textColor=colors.HexColor("#1E3A8A"))
    group_header_style = ParagraphStyle("GroupHeader", parent=styles["Normal"], fontName=FONT_NAME, fontSize=10, leading=13, textColor=colors.HexColor("#1E3A8A"), spaceBefore=6, spaceAfter=2)
    header_style = ParagraphStyle("Header", parent=styles["Normal"], fontName=FONT_NAME, fontSize=7, leading=8.5, textColor=colors.white, alignment=1)
    
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontName=FONT_NAME, fontSize=6.5, leading=8, textColor=colors.HexColor("#0F172A"))
    red_cell_style = ParagraphStyle("RedCell", parent=cell_style, textColor=colors.HexColor("#DC2626"))

    group_label_str = "顧客グループ別（出版社順）" if group_type == "publisher" else "判型グループ別（判型順）"
    elements = [
        Paragraph(f"<b>受注一覧表②（{group_label_str}・行全体赤網掛け強調）</b>", title_style),
        Spacer(1, 4),
        Paragraph(f"<font size=8 color='#64748B'><b>発行日: {date.today().strftime('%Y年%m月%d日')}</b></font>", styles["Normal"]),
        Spacer(1, 6),
    ]

    top_cols = [
        "受注番号", "納期", "納期上/中/下", "急ぎ", "出版社名", "書名", "刷数", "回数", "単価", "製本数", "製品預かり", "総頁数",
        "32P", "16P", "8P", "4P", "2P", "印刷所", "折本日", "折", "無線日", "三方日", "仕上日", "入品日",
        "表紙日付", "カード日付", "カバー日付", "帯日付"
    ]
    available_top_cols = [c for c in top_cols if c in df.columns]

    width_mapping = {
        "受注番号": 48, 
        "納期": 46, "納期上/中/下": 32, "急ぎ": 24, "折本日": 46, "無線日": 46, "三方日": 46, "仕上日": 46, "入品日": 46,
        "表紙日付": 46, "カード日付": 46, "カバー日付": 46, "帯日付": 46,
        "出版社名": 68, "書名": 110, "刷数": 26, "回数": 26,
        "単価": 36, "製本数": 42, "製品預かり": 42, "総頁数": 38, "32P": 26, "16P": 26,
        "8P": 26, "4P": 26, "2P": 26, "印刷所": 45, "折": 26
    }
    col_widths = [width_mapping.get(c, 40) for c in available_top_cols]

    group_col = "出版社名" if group_type == "publisher" else "判型"
    groups = df[group_col].unique() if group_col in df.columns else [""]

    for g_val in groups:
        g_name_display = g_val if g_val else "（未設定）"
        prefix_label = "■ 出版社: " if group_type == "publisher" else "■ 判型: "
        elements.append(Paragraph(f"<b>{prefix_label}{g_name_display} 御中</b>" if group_type == "publisher" else f"<b>{prefix_label}{g_name_display}</b>", group_header_style))
        
        sub_df = df[df[group_col] == g_val] if g_val else df

        table_data = []
        header_row = [Paragraph(f"<b>{c}</b>", header_style) for c in available_top_cols]
        table_data.append(header_row)

        t_style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#64748B")),
        ]

        for i, (_, row) in enumerate(sub_df.iterrows()):
            val_102_check = str(row.get("納期上/中/下", ""))
            val_103_check = str(row.get("急ぎ", ""))
            is_urgent = (val_102_check == "確定" or val_103_check == "急")

            row_cells_1 = []
            for c in available_top_cols:
                val_str = str(row[c]) if pd.notna(row[c]) else ""
                if is_urgent and (c in ["納期上/中/下", "急ぎ"]):
                    row_cells_1.append(Paragraph(f"<b>{val_str}</b>", red_cell_style))
                else:
                    row_cells_1.append(Paragraph(f"<b>{val_str}</b>", cell_style))
            
            table_data.append(row_cells_1)
            
            memo_val = str(row["備考"]) if "備考" in df.columns and pd.notna(row["備考"]) else ""
            memo_cell = Paragraph(f"<b>備考: {memo_val}</b>", cell_style)
            row_cells_2 = [memo_cell] + [Paragraph("", cell_style) for _ in range(len(available_top_cols) - 1)]
            table_data.append(row_cells_2)

            row_1_idx = 1 + (i * 2)
            memo_row_idx = row_1_idx + 1

            if is_urgent:
                t_style.append(("BACKGROUND", (0, row_1_idx), (-1, row_1_idx), colors.HexColor("#FEE2E2")))
                t_style.append(("BACKGROUND", (0, memo_row_idx), (-1, memo_row_idx), colors.HexColor("#FEE2E2")))
            else:
                t_style.append(("BACKGROUND", (0, memo_row_idx), (-1, memo_row_idx), colors.HexColor("#F1F5F9")))

            t_style.append(("SPAN", (0, memo_row_idx), (-1, memo_row_idx)))

        t = Table(table_data, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle(t_style))
        elements.append(t)
        elements.append(Spacer(1, 4))

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
st.header("📋 受注一覧表②（@pocket一括保存対応）")

with st.expander("🔍 検索条件・出力グループ軸の設定はこちら", expanded=bool(not url_order_no)):
    search_date_type = st.radio("📅 期間検索の基準を選択", ["納期（field-52）", "受注日（field-1）"], horizontal=True)
    target_date_field = "field-52" if "納期" in search_date_type else "field-1"

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        t1_start = st.date_input("期間（開始）", value=date.today(), key="v2_s")
    with col_s2:
        t1_end = st.date_input("期間（終了）", value=date.today() + timedelta(days=14), key="v2_e")
    
    publisher_input = st.text_input("📖 出版社名で絞り込み", value="", key="v2_pub")
    
    st.markdown("---")
    group_axis_choice = st.radio(
        "🗂️ 一覧表のグループ軸を選択（ソート順）",
        options=["① 顧客グループ別（出版社順 ➔ 納期・判型順）", "② 判型グループ別（判型順 ➔ 納期・顧客順）"],
        horizontal=True,
        key="group_axis_radio"
    )
    current_group_type = "publisher" if "①" in group_axis_choice else "format"

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
        rec_id = r.get("_recordId", "")
        h_ord_no = clean_val(r.get("field-2", ""))
        h_ord_date = parse_date_safe(r.get("field-1"))
        h_due = parse_date_safe(r.get("field-52"))
        
        val_102 = clean_val(r.get("field-102", ""))
        if val_102 not in ["", "確定", "上旬", "中旬", "下旬"]:
            val_102 = ""
        val_103 = clean_val(r.get("field-103", ""))
        if val_103 not in ["", "急"]:
            val_103 = ""

        h_pub = clean_val(r.get("field-8", ""))
        h_title = clean_val(r.get("field-13", ""))
        h_format = clean_val(r.get("field-23", ""))
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
        sub_90_id = str(item_90.get("Id") or item_90.get("recordId") or rec_90.get("Id") or "").strip()
        
        item_91 = sub_91[0] if isinstance(sub_91, list) and len(sub_91) > 0 else {}
        rec_91 = item_91.get("record", item_91) if isinstance(item_91, dict) else {}
        sub_91_id = str(item_91.get("Id") or item_91.get("recordId") or rec_91.get("Id") or "").strip()

        flat_data.append({
            "更新": False,
            "_record_idx": record_idx,
            "_recordId": rec_id,
            "_sub_90_id": sub_90_id, # 明細レコードID保持
            "_sub_91_id": sub_91_id, # 明細レコードID保持
            "出版社名": h_pub,
            "判型": h_format,
            "受注番号": h_ord_no,
            "納期": h_due,
            "納期上/中/下": val_102,
            "急ぎ": val_103,
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

    if current_group_type == "publisher":
        sort_keys = ["出版社名", "納期", "判型"]
    else:
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
            "_recordId": None,
            "_sub_90_id": None,
            "_sub_91_id": None,
            "出版社名": st.column_config.TextColumn("出版社名", disabled=True),
            "判型": st.column_config.TextColumn("判型", disabled=True),
            "受注番号": st.column_config.TextColumn("受注番号", disabled=True),
            "納期": st.column_config.DateColumn("納期", format="YYYY-MM-DD", disabled=True),
            "納期上/中/下": st.column_config.SelectboxColumn(
                "納期上/中/下",
                options=["", "確定", "上旬", "中旬", "下旬"],
                help="納期の上旬・中旬・下旬・確定を選択します (field-102)"
            ),
            "急ぎ": st.column_config.SelectboxColumn(
                "急ぎ",
                options=["", "急"],
                help="急ぎ案件の場合は「急」を選択します (field-103)"
            ),
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

    # --- @pocket一括保存ボタンの処理（サブテーブル・明細レコードID対応） ---
    st.markdown("---")
    col_save, col_d1, col_d2 = st.columns([1.5, 1.5, 1.5])
    
    with col_save:
        if st.button("💾 変更した行を @pocket に保存", type="primary", key="save_pocket_btn"):
            update_rows = edited_v2_df[edited_v2_df["更新"] == True]
            if len(update_rows) == 0:
                st.warning("⚠️ 「更新」にチェックが入っている行がありません。")
            else:
                success_count = 0
                error_messages = []
                with st.spinner("@pocketへデータを保存中..."):
                    for _, row in update_rows.iterrows():
                        r_id = row.get("_recordId", "")
                        sub_90_id = row.get("_sub_90_id", "")
                        sub_91_id = row.get("_sub_91_id", "")
                        
                        # field-90（工程日）のサブテーブル構造を作成（明細レコードIDを含める）
                        sub_90_item = {
                            "record": {
                                "field-90_3": date_to_str(row.get("折本日")),
                                "field-90_4": date_to_str(row.get("無線日")),
                                "field-90_5": date_to_str(row.get("三方日")),
                                "field-90_6": date_to_str(row.get("仕上日")),
                                "field-90_10": date_to_str(row.get("入品日"))
                            }
                        }
                        if sub_90_id:
                            sub_90_item["Id"] = sub_90_id

                        # field-91（付物日）のサブテーブル構造を作成（明細レコードIDを含める）
                        sub_91_item = {
                            "record": {
                                "field-91_1": date_to_str(row.get("表紙日付")),
                                "field-91_9": date_to_str(row.get("カード日付")),
                                "field-91_5": date_to_str(row.get("カバー日付")),
                                "field-91_13": date_to_str(row.get("帯日付"))
                            }
                        }
                        if sub_91_id:
                            sub_91_item["Id"] = sub_91_id

                        payload = {
                            "field-102": clean_val(row.get("納期上/中/下", "")),
                            "field-103": clean_val(row.get("急ぎ", "")),
                            "field-28": clean_val(row.get("刷数", "")),
                            "field-45": clean_val(row.get("備考", "")),
                            "field-90": [sub_90_item],
                            "field-91": [sub_91_item]
                        }
                        
                        ok, msg = update_pocket_record(r_id, payload)
                        if ok:
                            success_count += 1
                        else:
                            error_messages.append(f"ID {r_id}: {msg}")
                
                if success_count > 0:
                    st.success(f"✨ **{success_count} 件** のデータを @pocket へ正常に保存しました！")
                if error_messages:
                    for err in error_messages:
                        st.error(f"❌ 保存エラー: {err}")
                if success_count > 0:
                    st.rerun()

    with col_d1:
        excel_bytes = to_excel_bytes(edited_v2_df)
        st.download_button(
            label="📥 Excel でダウンロード",
            data=excel_bytes,
            file_name=f"受注一覧表2_{current_group_type}_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_d2:
        pdf_bytes = generate_order_pdf_v2(edited_v2_df, group_type=current_group_type)
        st.download_button(
            label=f"🖨️ A3 PDF ダウンロード（{group_axis_choice[:2]}）",
            data=pdf_bytes,
            file_name=f"受注一覧表2_{current_group_type}_{date.today().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )
else:
    st.warning("⚠️ 対象データが見つかりませんでした。")