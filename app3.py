import io
import os
import requests
import pandas as pd
from datetime import datetime, timedelta, date
import streamlit as st

# ReportLab関連のインポート（PDF出力用）
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

# openpyxl関連のインポート（きれいなExcel出力用）
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CM_TO_PT = 1.0 / 0.0352778
ROW1_HEIGHT_CM = 1.2
ROW1_HEIGHT_PT = ROW1_HEIGHT_CM * CM_TO_PT

# ==========================================
# 1. ページ基本設定 ＆ フルスクリーン切り替えCSS
# ==========================================
st.set_page_config(page_title="受注一覧表（②・③断裁用）管理システム", layout="wide", initial_sidebar_state="expanded")

if "fullscreen_mode" not in st.session_state:
    st.session_state["fullscreen_mode"] = False

if st.session_state["fullscreen_mode"]:
    dynamic_layout_css = """
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        [data-testid="stHeader"] {display: none !important;}
        [data-testid="stSidebar"] {display: none !important;}
        [data-testid="stMainBlockContainer"] {
            max-width: 100% !important;
            padding-top: 1rem !important;
            padding-right: 2rem !important;
            padding-left: 2rem !important;
        }
        .stApp { background-color: #F8FAFC; }
        </style>
    """
else:
    dynamic_layout_css = """
        <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .stApp { background-color: #F8FAFC; }
        </style>
    """

st.markdown(dynamic_layout_css, unsafe_allow_html=True)

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
# 4. 判型ソート順定義 ＆ 便利関数
# ==========================================
FORMAT_ORDER_LIST = [
    "新書変", "新書", "文庫", "I・G", "Jr新", "少文", "単行", "判",
    "A4", "B4", "A5", "B5", "A6", "B6", "A7", "B7", "B20", "AB判",
    "特Ｂ４", "菊倍判", "B6ワイド版", "菊六変形", "岩波4/6判", "大月4/6判",
    "勁草4/6判", "フォア文庫", "A3", "A4変", "Ａ４変横", "A4横", "B5変",
    "B4変", "６２×９２", "Ｌ判", "６２Ｘ９７", "理想4/6判", "A6変", "新評4/6",
    "Ｂ5横本", "A5変", "三秀4/6判", "N新書", "新書変形", "ディグ4/6", "菊変",
    "岩波菊", "新評4/6変", "規格外", "加瀬4/6判", "B6変", "白水U判", "白水Q判",
    "磯崎4/6変", "青土4/6", "加瀬4/6変", "A5横", "恒星社4/6", "三秀A5変",
    "三陽A5変", "加瀬4/6変", "NJｒ新書", "白水4/6",
]

FORMAT_RANK_MAP = {}
for idx, fmt_name in enumerate(FORMAT_ORDER_LIST):
    norm_name = fmt_name.strip().replace("\u3000", "")
    if norm_name and norm_name not in FORMAT_RANK_MAP:
        FORMAT_RANK_MAP[norm_name] = idx

UNRANKED_FORMAT_RANK = len(FORMAT_ORDER_LIST) + 1

def get_format_rank(fmt_val):
    s = clean_val(fmt_val).replace("\u3000", "")
    if not s:
        return UNRANKED_FORMAT_RANK + 1
    return FORMAT_RANK_MAP.get(s, UNRANKED_FORMAT_RANK)

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

def format_umu_mark(val):
    s = clean_val(val)
    if s:
        return f"{s}/〇"
    return "✖"

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

def format_date_pdf(val):
    if pd.isna(val) or not val:
        return ""
    if isinstance(val, (date, datetime)):
        return val.strftime("%m/%d")
    parsed = parse_date_safe(val)
    return parsed.strftime("%m/%d") if parsed else str(val)

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

def sort_order_dataframe(df, group_type="publisher"):
    if df is None or len(df) == 0:
        return df
    work_df = df.copy()
    work_df["_fmt_rank"] = work_df["判型"].apply(get_format_rank) if "判型" in work_df.columns else UNRANKED_FORMAT_RANK
    work_df["_due_sort"] = pd.to_datetime(work_df["納期"], errors="coerce") if "納期" in work_df.columns else pd.NaT

    if group_type == "format":
        sort_cols = ["_fmt_rank", "判型", "_due_sort", "出版社名"]
        asc_flags = [True, True, True, True]
    else:
        sort_cols = ["_due_sort", "_fmt_rank", "判型", "出版社名"]
        asc_flags = [True, True, True, True]

    existing_cols = [c for c in sort_cols if c in work_df.columns]
    existing_asc = [asc_flags[i] for i, c in enumerate(sort_cols) if c in work_df.columns]

    work_df = work_df.sort_values(by=existing_cols, ascending=existing_asc, na_position="last").reset_index(drop=True)
    return work_df.drop(columns=["_fmt_rank", "_due_sort"], errors="ignore")

# ==========================================
# ③断裁用 列定義（内部キー・表示名・PDF幅pt・Excel幅）
# ==========================================
CUTTING_COLS_SPEC = [
    ("受注番号", "受注番号", 40, 11),
    ("納期", "納期", 28, 8),
    ("判型", "判型", 28, 8),
    ("納期2", "納期2", 26, 7.5),
    ("急ぎ", "急ぎ", 22, 6.5),
    ("出版社名", "出版社名", 34, 10),
    ("書名", "書名", 115, 34),
    ("刷数", "刷数", 22, 6.5),
    ("回数", "回数", 22, 6.5),
    ("単価", "単価", 30, 8),
    ("製本数", "製本数", 34, 9),
    ("製品預かり", "製品預かり", 34, 10),
    ("総頁数", "総頁数", 28, 7.5),
    ("32P", "32P", 20, 6),
    ("16P", "16P", 20, 7),
    ("8P", "8P", 20, 8),
    ("4P", "4P", 20, 9),
    ("2P", "2P", 20, 10),
    ("印刷所", "印刷所", 38, 9),
    ("表紙有無_結合", "表紙有無 結合", 44, 11),
    ("表紙有無_手書", "表紙有無 結合", 44, 11),
    ("カード有無_結合", "カード有無 結合", 48, 12),
    ("カード有無_手書", "カード有無 結合", 44, 11),
    ("カバー有無_結合", "カバー有無 結合", 50, 13),
    ("カバー有無_手書", "カバー有無 結合", 44, 11),
    ("帯有無_結合", "帯有無 結合", 46, 11),
    ("帯有無_手書", "帯有無 結合", 44, 11),
]

V3_DISPLAY_LABEL_MAP = {k: lbl for k, lbl, _, _ in CUTTING_COLS_SPEC}
V3_PDF_WIDTH_MAP = {k: w for k, _, w, _ in CUTTING_COLS_SPEC}
V3_EXCEL_WIDTH_MAP = {k: w for k, _, _, w in CUTTING_COLS_SPEC}

# ==========================================
# PDF生成関数（一覧表② / ③断裁用・1行目高さ1.2cm）
# ==========================================
def generate_order_pdf(df, table_mode="v2", group_type="publisher"):
    df = sort_order_dataframe(df, group_type=group_type)
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A3), leftMargin=8, rightMargin=8, topMargin=14, bottomMargin=14)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("Title", parent=styles["Normal"], fontName=FONT_NAME, fontSize=14, leading=16, textColor=colors.HexColor("#1E3A8A"))
    group_header_style = ParagraphStyle("GroupHeader", parent=styles["Normal"], fontName=FONT_NAME, fontSize=10, leading=13, textColor=colors.HexColor("#1E3A8A"), spaceBefore=5, spaceAfter=2, leftIndent=0, alignment=0)

    header_style = ParagraphStyle("Header", parent=styles["Normal"], fontName=FONT_NAME, fontSize=7.5, leading=9.0, textColor=colors.white, alignment=1)
    cell_style = ParagraphStyle("Cell", parent=styles["Normal"], fontName=FONT_NAME, fontSize=7.5 if table_mode == "v2" else 7.0, leading=9.0 if table_mode == "v2" else 8.5, textColor=colors.HexColor("#0F172A"))
    red_cell_style = ParagraphStyle("RedCell", parent=cell_style, textColor=colors.HexColor("#DC2626"))
    memo_style = ParagraphStyle("MemoCell", parent=cell_style, fontSize=8.0 if table_mode == "v2" else 7.5, leading=10.0 if table_mode == "v2" else 9.5)

    mode_title = "受注一覧表②" if table_mode == "v2" else "受注一覧表③断裁用"
    group_label_str = "出版社一覧表（納期順・判型順）" if group_type == "publisher" else "判型一覧表（判型順・納期順）"
    elements = [
        Paragraph(f"<b>{mode_title}（{group_label_str}・行全体赤網掛け強調）</b>", title_style),
        Spacer(1, 4),
        Paragraph(f"<font size=8 color='#64748B'><b>発行日: {date.today().strftime('%Y年%m月%d日')}</b></font>", styles["Normal"]),
        Spacer(1, 4),
    ]

    if table_mode == "v2":
        top_cols = [
            "受注番号", "見本日", "配本日", "納期", "納期2", "急ぎ", "出版社名", "判型", "書名", "刷数", "回数", "単価", "製本数", "製品預かり", "総頁数",
            "32P", "16P", "8P", "4P", "2P", "印刷所", "刷本日", "折", "折本日", "無線日", "三方日", "仕上日", "納品日",
            "表紙日付", "カード日", "カバ日付", "帯日付"
        ]
        available_top_cols = [c for c in top_cols if c in df.columns]
        width_mapping = {
            "受注番号": 42,
            "見本日": 26, "配本日": 26, "納期": 26, "納期2": 26, "急ぎ": 22,
            "出版社名": 34, "判型": 28, "書名": 132,
            "刷数": 22, "回数": 22, "単価": 32, "製本数": 36, "製品預かり": 36, "総頁数": 30,
            "32P": 20, "16P": 20, "8P": 20, "4P": 20, "2P": 20,
            "印刷所": 40,
            "刷本日": 26, "折": 20, "折本日": 26, "無線日": 26, "三方日": 26, "仕上日": 26, "納品日": 26,
            "表紙日付": 26, "カード日": 26, "カバ日付": 26, "帯日付": 26
        }
        col_widths = [width_mapping.get(c, 26) for c in available_top_cols]
        date_cols_set = {
            "見本日", "配本日", "納期", "刷本日", "折本日", "無線日", "三方日", "仕上日", "納品日",
            "表紙日付", "カード日", "カバ日付", "帯日付"
        }
        label_map = {c: c for c in available_top_cols}
    else:
        available_top_cols = [k for k, _, _, _ in CUTTING_COLS_SPEC if k in df.columns]
        col_widths = [V3_PDF_WIDTH_MAP.get(c, 30) for c in available_top_cols]
        date_cols_set = {"納期"}
        label_map = V3_DISPLAY_LABEL_MAP

    group_col = "出版社名" if group_type == "publisher" else "判型"
    groups = df[group_col].unique() if group_col in df.columns else [""]

    for g_val in groups:
        g_name_display = g_val if g_val else "（未設定）"
        prefix_label = "■ 出版社: " if group_type == "publisher" else "■ 判型: "
        header_text = f"<b>{prefix_label}{g_name_display} 御中</b>" if group_type == "publisher" else f"<b>{prefix_label}{g_name_display}</b>"
        elements.append(Paragraph(header_text, group_header_style))

        sub_df = df[df[group_col] == g_val] if g_val else df

        table_data = []
        header_row = [Paragraph(f"<b>{label_map.get(c, c)}</b>", header_style) for c in available_top_cols]
        table_data.append(header_row)

        t_style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("LEFTPADDING", (0, 0), (-1, -1), 1.2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 1.2),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#64748B")),
        ]

        row_heights = [20 if table_mode == "v3" else 19]

        for i, (_, row) in enumerate(sub_df.iterrows()):
            val_102_check = str(row.get("納期2", ""))
            val_103_check = str(row.get("急ぎ", ""))
            is_urgent = (val_102_check == "確定" or val_103_check == "急")

            row_cells_1 = []
            for c in available_top_cols:
                if c.endswith("_手書"):
                    val_str = ""
                elif c in date_cols_set:
                    val_str = format_date_pdf(row.get(c, ""))
                else:
                    raw_val = row.get(c, "")
                    val_str = str(raw_val) if pd.notna(raw_val) else ""

                if is_urgent and (c in ["納期2", "急ぎ"]):
                    row_cells_1.append(Paragraph(f"<b>{val_str}</b>", red_cell_style))
                else:
                    row_cells_1.append(Paragraph(f"<b>{val_str}</b>", cell_style))

            table_data.append(row_cells_1)
            row_heights.append(ROW1_HEIGHT_CM * cm)

            memo_val = str(row["備考"]) if "備考" in df.columns and pd.notna(row["備考"]) else ""
            memo_html = memo_val.replace("\n", "<br/>")
            memo_cell = Paragraph(f"<b>備考: {memo_html}</b>", memo_style)
            row_cells_2 = [memo_cell] + [Paragraph("", cell_style) for _ in range(len(available_top_cols) - 1)]
            table_data.append(row_cells_2)
            row_heights.append(46 if table_mode == "v2" else 44)

            row_1_idx = 1 + (i * 2)
            memo_row_idx = row_1_idx + 1

            if is_urgent:
                t_style.append(("BACKGROUND", (0, row_1_idx), (-1, row_1_idx), colors.HexColor("#FEE2E2")))
                t_style.append(("BACKGROUND", (0, memo_row_idx), (-1, memo_row_idx), colors.HexColor("#FEE2E2")))
            else:
                t_style.append(("BACKGROUND", (0, memo_row_idx), (-1, memo_row_idx), colors.HexColor("#F1F5F9")))

            t_style.append(("SPAN", (0, memo_row_idx), (-1, memo_row_idx)))
            t_style.append(("VALIGN", (0, memo_row_idx), (-1, memo_row_idx), "TOP"))

        t = Table(table_data, colWidths=col_widths, rowHeights=row_heights, repeatRows=1, hAlign="LEFT")
        t.setStyle(TableStyle(t_style))
        elements.append(t)
        elements.append(Spacer(1, 4))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# Excel生成関数（一覧表② / ③断裁用・1行目高さ1.2cm）
# ==========================================
def generate_order_excel(df, table_mode="v2", group_type="publisher"):
    df = sort_order_dataframe(df, group_type=group_type)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "受注一覧表②" if table_mode == "v2" else "受注一覧表③断裁用"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Meiryo", size=14, bold=True, color="1E3A8A")
    font_sub = Font(name="Meiryo", size=9, color="64748B")
    font_group = Font(name="Meiryo", size=11, bold=True, color="1E3A8A")
    font_header = Font(name="Meiryo", size=10, bold=True, color="FFFFFF")
    font_cell = Font(name="Meiryo", size=10, color="0F172A")
    font_red_cell = Font(name="Meiryo", size=10, bold=True, color="DC2626")
    font_memo = Font(name="Meiryo", size=10, bold=True, color="0F172A")

    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_urgent = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fill_memo = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    border_thin = Border(
        left=Side(style="thin", color="64748B"), right=Side(style="thin", color="64748B"),
        top=Side(style="thin", color="64748B"), bottom=Side(style="thin", color="64748B")
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_memo_top = Alignment(horizontal="left", vertical="top", wrap_text=True)

    if table_mode == "v2":
        top_cols = [
            "受注番号", "見本日", "配本日", "納期", "納期2", "急ぎ", "出版社名", "判型", "書名", "刷数", "回数", "単価", "製本数", "製品預かり", "総頁数",
            "32P", "16P", "8P", "4P", "2P", "印刷所", "刷本日", "折", "折本日", "無線日", "三方日", "仕上日", "納品日",
            "表紙日付", "カード日", "カバ日付", "帯日付"
        ]
        available_top_cols = [c for c in top_cols if c in df.columns]
        excel_width_map = {
            "受注番号": 11,
            "見本日": 7.5, "配本日": 7.5, "納期": 7.5, "納期2": 7.5, "急ぎ": 6,
            "出版社名": 10,
            "判型": 8,
            "書名": 38,
            "刷数": 6, "回数": 6, "単価": 8, "製本数": 9, "製品預かり": 10, "総頁数": 8,
            "32P": 6, "16P": 6, "8P": 6, "4P": 6, "2P": 6,
            "印刷所": 11,
            "刷本日": 7.5, "折": 6, "折本日": 7.5, "無線日": 7.5, "三方日": 7.5, "仕上日": 7.5, "納品日": 7.5,
            "表紙日付": 7.5, "カード日": 7.5, "カバ日付": 7.5, "帯日付": 7.5
        }
        date_cols_set = {
            "見本日", "配本日", "納期", "刷本日", "折本日", "無線日", "三方日", "仕上日", "納品日",
            "表紙日付", "カード日", "カバ日付", "帯日付"
        }
        label_map = {c: c for c in available_top_cols}
    else:
        available_top_cols = [k for k, _, _, _ in CUTTING_COLS_SPEC if k in df.columns]
        excel_width_map = V3_EXCEL_WIDTH_MAP
        date_cols_set = {"納期"}
        label_map = V3_DISPLAY_LABEL_MAP

    max_col_idx = len(available_top_cols)
    mode_title = "受注一覧表②" if table_mode == "v2" else "受注一覧表③断裁用"
    group_title = "出版社一覧表（納期順・判型順）" if group_type == "publisher" else "判型一覧表（判型順・納期順）"

    ws.append([f"{mode_title}（{group_title}）"])
    ws.cell(row=1, column=1).font = font_title
    ws.append([f"発行日: {date.today().strftime('%Y年%m月%d日')}"])
    ws.cell(row=2, column=1).font = font_sub
    ws.append([])

    group_col = "出版社名" if group_type == "publisher" else "判型"
    groups = df[group_col].unique() if group_col in df.columns else [""]

    current_row = 4

    for g_val in groups:
        g_name_display = g_val if g_val else "（未設定）"
        prefix_label = "■ 出版社: " if group_type == "publisher" else "■ 判型: "
        suffix_str = " 御中" if group_type == "publisher" else ""

        group_cell = ws.cell(row=current_row, column=1, value=f"{prefix_label}{g_name_display}{suffix_str}")
        group_cell.font = font_group
        group_cell.alignment = Alignment(horizontal="left", vertical="center")
        current_row += 1

        header_row_idx = current_row
        ws.append([label_map.get(c, c) for c in available_top_cols])
        for col_i in range(1, max_col_idx + 1):
            cell = ws.cell(row=header_row_idx, column=col_i)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_thin
        current_row += 1

        sub_df = df[df[group_col] == g_val] if g_val else df

        for _, row in sub_df.iterrows():
            val_102_check = str(row.get("納期2", ""))
            val_103_check = str(row.get("急ぎ", ""))
            is_urgent = (val_102_check == "確定" or val_103_check == "急")

            data_row_idx = current_row
            row_vals = []
            for c in available_top_cols:
                if c.endswith("_手書"):
                    row_vals.append("")
                elif c in date_cols_set:
                    row_vals.append(format_date_pdf(row.get(c, "")))
                else:
                    raw_val = row.get(c, "")
                    row_vals.append(str(raw_val) if pd.notna(raw_val) else "")

            ws.append(row_vals)
            ws.row_dimensions[data_row_idx].height = ROW1_HEIGHT_PT
            for col_i in range(1, max_col_idx + 1):
                col_key = available_top_cols[col_i - 1]
                cell = ws.cell(row=data_row_idx, column=col_i)
                cell.font = font_red_cell if (is_urgent and col_key in ["納期2", "急ぎ"]) else font_cell
                if is_urgent:
                    cell.fill = fill_urgent
                cell.border = border_thin
                cell.alignment = align_left_wrap if col_key in ["書名", "出版社名"] else align_center
            current_row += 1

            memo_row_idx = current_row
            memo_val = str(row["備考"]) if "備考" in df.columns and pd.notna(row["備考"]) else ""
            ws.append([f"備考: {memo_val}"] + [""] * (max_col_idx - 1))
            ws.merge_cells(start_row=memo_row_idx, start_column=1, end_row=memo_row_idx, end_column=max_col_idx)
            ws.row_dimensions[memo_row_idx].height = 64

            for col_i in range(1, max_col_idx + 1):
                cell = ws.cell(row=memo_row_idx, column=col_i)
                cell.font = font_memo
                cell.fill = fill_urgent if is_urgent else fill_memo
                cell.border = border_thin
                cell.alignment = align_memo_top
            current_row += 1

        current_row += 1

    for col_i, col_key in enumerate(available_top_cols, start=1):
        col_letter = get_column_letter(col_i)
        ws.column_dimensions[col_letter].width = excel_width_map.get(col_key, 8)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

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
            query_parts.append(f'field-105 like "{target_publisher}"')
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
        today_date = date.today()
        default_start = date(today_date.year, today_date.month, 1)
        if today_date.month == 12:
            default_end = date(today_date.year + 1, 1, 1) - timedelta(days=1)
        else:
            default_end = date(today_date.year, today_date.month + 1, 1) - timedelta(days=1)

        st.session_state["api_records"] = fetch_pocket_records(start_date=default_start, end_date=default_end, target_publisher="岩波", date_field="field-52")

# ==========================================
# 6. サイドバー（検索条件の設定）
# ==========================================
with st.sidebar:
    st.header("⚙️ 検索条件の設定")
    search_date_type = st.radio("📅 期間検索の基準を選択", ["納期（field-52）", "受注日（field-1）"], horizontal=False)
    target_date_field = "field-52" if "納期" in search_date_type else "field-1"

    today_date = date.today()
    init_start = date(today_date.year, today_date.month, 1)
    if today_date.month == 12:
        init_end = date(today_date.year + 1, 1, 1) - timedelta(days=1)
    else:
        init_end = date(today_date.year, today_date.month + 1, 1) - timedelta(days=1)

    t1_start = st.date_input("期間（開始）", value=init_start, key="unified_s")
    t1_end = st.date_input("期間（終了）", value=init_end, key="unified_e")

    publisher_input = st.text_input("📖 出版社名で絞り込み（部分一致）", value="岩波", key="unified_pub")
    exclude_iwanami = st.checkbox("「岩波」を含む出版社（岩波〇〇など）をすべて除外する", value=False, key="unified_exc_iwanami")

    st.markdown("---")
    group_axis_choice = st.radio(
        "🗂️ グループ軸（ソート順）",
        options=["① 判型別一覧表（判型順・納期順）", "② 出版社一覧表（納期順・判型順）"],
        index=1,
        horizontal=False,
        key="group_axis_radio_unified"
    )
    current_group_type = "format" if "①" in group_axis_choice else "publisher"

    search_clicked = st.button("🔄 検索実行", type="primary", key="unified_btn", use_container_width=True)

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

# ==========================================
# 7. メイン画面（画面上の帳票切替＆表示）
# ==========================================
mode_choice = st.radio(
    "📑 表示する一覧表を選択",
    options=["受注一覧表②", "受注一覧表③断裁用"],
    horizontal=True,
    key="table_mode_selector"
)
current_table_mode = "v2" if "②" in mode_choice else "v3"

st.header(f"📋 {mode_choice}（{'各日付更新' if current_table_mode == 'v2' else '断裁・有無確認（閲覧用）'}）")

fs_toggle = st.toggle("🖥️ 画面を広々表示（フルスクリーンモード）", value=st.session_state["fullscreen_mode"], key="fs_main_switch_unified")
if fs_toggle != st.session_state["fullscreen_mode"]:
    st.session_state["fullscreen_mode"] = fs_toggle
    st.rerun()

st.markdown("---")

api_records = st.session_state.get("api_records", None)

if api_records:
    flat_data = []
    for record_idx, r in enumerate(api_records):
        rec_id = r.get("_recordId", "")
        detail_url = f"{POCKET_BASE_URL}/apps/{APP_ID}/records/{rec_id}/detail" if rec_id and rec_id.lower() != "none" else ""

        h_ord_no = clean_val(r.get("field-2", ""))
        h_mihon = parse_date_safe(r.get("field-54"))
        h_haibon = parse_date_safe(r.get("field-56"))
        h_due = parse_date_safe(r.get("field-52"))

        val_102 = clean_val(r.get("field-102", ""))
        if val_102 not in ["", "確定", "上旬", "中旬", "下旬"]:
            val_102 = ""
        val_103 = clean_val(r.get("field-103", ""))
        if val_103 not in ["", "急"]:
            val_103 = ""

        h_pub = clean_val(r.get("field-105", ""))
        if exclude_iwanami and "岩波" in h_pub:
            continue

        h_format = clean_val(r.get("field-23", ""))
        h_title = clean_val(r.get("field-13", ""))
        h_print_count = clean_val(r.get("field-28", ""))
        h_times_count = clean_val(r.get("field-96", ""))
        h_price = clean_val(r.get("field-5", ""))
        h_qty = clean_val(r.get("field-4", ""))
        h_cust_keeping = clean_val(r.get("field-81", ""))
        h_total_pages = clean_val(r.get("field-82", ""))
        h_p32 = clean_val(r.get("field-30", ""))
        h_p16 = clean_val(r.get("field-31", ""))
        h_p8 = clean_val(r.get("field-32", ""))
        h_p4 = clean_val(r.get("field-33", ""))
        h_p2 = clean_val(r.get("field-34", ""))
        h_printer = clean_val(r.get("field-14", ""))
        h_print_date = parse_date_safe(r.get("field-51"))
        h_folding = clean_val(r.get("field-16", ""))
        h_memo = clean_val(r.get("field-45", ""))

        sub_90 = r.get("field-90", [])
        sub_91 = r.get("field-91", [])

        item_90 = sub_90[0] if isinstance(sub_90, list) and len(sub_90) > 0 else {}
        rec_90 = item_90.get("record", item_90) if isinstance(item_90, dict) else {}
        sub_90_id = str(item_90.get("Id") or item_90.get("recordId") or rec_90.get("Id") or "").strip()

        item_91 = sub_91[0] if isinstance(sub_91, list) and len(sub_91) > 0 else {}
        rec_91 = item_91.get("record", item_91) if isinstance(item_91, dict) else {}
        sub_91_id = str(item_91.get("Id") or item_91.get("recordId") or rec_91.get("Id") or "").strip()

        base_row = {
            "更新": False,
            "詳細": detail_url,
            "_record_idx": record_idx,
            "_recordId": rec_id,
            "_sub_90_id": sub_90_id,
            "_sub_91_id": sub_91_id,
        }

        if current_table_mode == "v2":
            base_row.update({
                "受注番号": h_ord_no,
                "見本日": h_mihon,
                "配本日": h_haibon,
                "納期": h_due,
                "納期2": val_102,
                "急ぎ": val_103,
                "出版社名": h_pub,
                "判型": h_format,
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
                "刷本日": h_print_date,
                "折": h_folding,
                "折本日": parse_date_safe(rec_90.get("field-90_3")),
                "無線日": parse_date_safe(rec_90.get("field-90_4")),
                "三方日": parse_date_safe(rec_90.get("field-90_5")),
                "仕上日": parse_date_safe(rec_90.get("field-90_6")),
                "納品日": parse_date_safe(rec_90.get("field-90_10")),
                "表紙日付": parse_date_safe(rec_91.get("field-91_1")),
                "カード日": parse_date_safe(rec_91.get("field-91_9")),
                "カバ日付": parse_date_safe(rec_91.get("field-91_5")),
                "帯日付": parse_date_safe(rec_91.get("field-91_13")),
                "備考": h_memo
            })
        else:
            base_row.update({
                "受注番号": h_ord_no,
                "納期": h_due,
                "判型": h_format,
                "納期2": val_102,
                "急ぎ": val_103,
                "出版社名": h_pub,
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
                "表紙有無_結合": format_umu_mark(rec_91.get("field-91_2")),
                "表紙有無_手書": "",
                "カード有無_結合": format_umu_mark(rec_91.get("field-91_10")),
                "カード有無_手書": "",
                "カバー有無_結合": format_umu_mark(rec_91.get("field-91_6")),
                "カバー有無_手書": "",
                "帯有無_結合": format_umu_mark(rec_91.get("field-91_14")),
                "帯有無_手書": "",
                "備考": h_memo
            })

        flat_data.append(base_row)

    df_view = pd.DataFrame(flat_data)

    if len(df_view) > 0:
        st.success(f"✨ 対象レコード **{len(df_view)} 件** を取得しました（現在表示: **{mode_choice}**）。")

        df_view = sort_order_dataframe(df_view, group_type=current_group_type)

        common_col_cfg = {
            "更新": st.column_config.CheckboxColumn("更新", required=True, width="small"),
            "詳細": st.column_config.LinkColumn(
                "🔗 @DB詳細",
                help="@pocketの該当データ画面へジャンプします",
                display_text="📝 詳細",
                width="small"
            ),
            "_record_idx": None,
            "_recordId": None,
            "_sub_90_id": None,
            "_sub_91_id": None,
            "受注番号": st.column_config.TextColumn("受注番号", disabled=True, width=85),
            "納期": st.column_config.DateColumn("納期", format="MM/DD", disabled=True, width=65),
            "判型": st.column_config.TextColumn("判型", disabled=True, width=65),
            "納期2": st.column_config.SelectboxColumn(
                "納期2",
                options=["", "確定", "上旬", "中旬", "下旬"],
                width=65,
                help="納期の上旬・中旬・下旬・確定を選択します (field-102)"
            ),
            "急ぎ": st.column_config.SelectboxColumn(
                "急ぎ",
                options=["", "急"],
                width=55,
                help="急ぎ案件の場合は「急」を選択します (field-103)"
            ),
            "出版社名": st.column_config.TextColumn("出版社名", disabled=True, width=70),
            "書名": st.column_config.TextColumn("書名", disabled=True, width=260 if current_table_mode == "v2" else 220),
            "刷数": st.column_config.TextColumn("刷数", width=55),
            "回数": st.column_config.TextColumn("回数", width=55),
            "単価": st.column_config.TextColumn("単価", disabled=True, width=65),
            "製本数": st.column_config.TextColumn("製本数", disabled=True, width=70),
            "製品預かり": st.column_config.TextColumn("製品預かり", disabled=True, width=75),
            "総頁数": st.column_config.TextColumn("総頁数", disabled=True, width=65),
            "32P": st.column_config.TextColumn("32P", disabled=True, width=50),
            "16P": st.column_config.TextColumn("16P", disabled=True, width=55),
            "8P": st.column_config.TextColumn("8P", disabled=True, width=60),
            "4P": st.column_config.TextColumn("4P", disabled=True, width=65),
            "2P": st.column_config.TextColumn("2P", disabled=True, width=70),
            "印刷所": st.column_config.TextColumn("印刷所", disabled=True, width=70),
            "備考": st.column_config.TextColumn("備考", width="medium")
        }

        if current_table_mode == "v2":
            common_col_cfg.update({
                "見本日": st.column_config.DateColumn("見本日", format="MM/DD", width=65),
                "配本日": st.column_config.DateColumn("配本日", format="MM/DD", width=65),
                "刷本日": st.column_config.DateColumn("刷本日", format="MM/DD", width=65),
                "折": st.column_config.TextColumn("折", disabled=True, width="small"),
                "折本日": st.column_config.DateColumn("折本日", format="MM/DD", width=65),
                "無線日": st.column_config.DateColumn("無線日", format="MM/DD", width=65),
                "三方日": st.column_config.DateColumn("三方日", format="MM/DD", width=65),
                "仕上日": st.column_config.DateColumn("仕上日", format="MM/DD", width=65),
                "納品日": st.column_config.DateColumn("納品日", format="MM/DD", width=65),
                "表紙日付": st.column_config.DateColumn("表紙日付", format="MM/DD", width=65),
                "カード日": st.column_config.DateColumn("カード日", format="MM/DD", width=65),
                "カバ日付": st.column_config.DateColumn("カバ日付", format="MM/DD", width=65),
                "帯日付": st.column_config.DateColumn("帯日付", format="MM/DD", width=65),
            })
        else:
            common_col_cfg.update({
                "表紙有無_結合": st.column_config.TextColumn("表紙有無 結合", disabled=True, width=80),
                "表紙有無_手書": st.column_config.TextColumn("表紙有無 結合", disabled=True, width=80),
                "カード有無_結合": st.column_config.TextColumn("カード有無 結合", disabled=True, width=85),
                "カード有無_手書": st.column_config.TextColumn("カード有無 結合", disabled=True, width=85),
                "カバー有無_結合": st.column_config.TextColumn("カバー有無 結合", disabled=True, width=90),
                "カバー有無_手書": st.column_config.TextColumn("カバー有無 結合", disabled=True, width=90),
                "帯有無_結合": st.column_config.TextColumn("帯有無 結合", disabled=True, width=85),
                "帯有無_手書": st.column_config.TextColumn("帯有無 結合", disabled=True, width=85),
            })

        edited_df = st.data_editor(
            df_view,
            num_rows="fixed",
            key=f"grid_{current_table_mode}",
            use_container_width=False,
            column_config=common_col_cfg
        )

        # --- @pocket一括保存ボタンの処理 ---
        st.markdown("---")
        col_save, col_d1, col_d2 = st.columns([1.5, 1.5, 1.5])

        with col_save:
            if st.button("💾 変更した行を @pocket に保存", type="primary", key=f"save_pocket_btn_{current_table_mode}"):
                update_rows = edited_df[edited_df["更新"] == True]
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

                            if current_table_mode == "v2":
                                sub_90_item = {
                                    "record": {
                                        "field-90_3": date_to_str(row.get("折本日")),
                                        "field-90_4": date_to_str(row.get("無線日")),
                                        "field-90_5": date_to_str(row.get("三方日")),
                                        "field-90_6": date_to_str(row.get("仕上日")),
                                        "field-90_10": date_to_str(row.get("納品日"))
                                    }
                                }
                                if sub_90_id:
                                    sub_90_item["Id"] = sub_90_id

                                sub_91_item = {
                                    "record": {
                                        "field-91_1": date_to_str(row.get("表紙日付")),
                                        "field-91_9": date_to_str(row.get("カード日")),
                                        "field-91_5": date_to_str(row.get("カバ日付")),
                                        "field-91_13": date_to_str(row.get("帯日付"))
                                    }
                                }
                                if sub_91_id:
                                    sub_91_item["Id"] = sub_91_id

                                payload = {
                                    "field-2": clean_val(row.get("受注番号", "")),
                                    "field-54": date_to_str(row.get("見本日")),
                                    "field-56": date_to_str(row.get("配本日")),
                                    "field-102": clean_val(row.get("納期2", "")),
                                    "field-103": clean_val(row.get("急ぎ", "")),
                                    "field-28": clean_val(row.get("刷数", "")),
                                    "field-51": date_to_str(row.get("刷本日")),
                                    "field-45": clean_val(row.get("備考", "")),
                                    "field-90": [sub_90_item],
                                    "field-91": [sub_91_item]
                                }
                            else:
                                payload = {
                                    "field-2": clean_val(row.get("受注番号", "")),
                                    "field-102": clean_val(row.get("納期2", "")),
                                    "field-103": clean_val(row.get("急ぎ", "")),
                                    "field-28": clean_val(row.get("刷数", "")),
                                    "field-45": clean_val(row.get("備考", ""))
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

        file_prefix = "受注一覧表2" if current_table_mode == "v2" else "受注一覧表3断裁用"
        with col_d1:
            excel_bytes = generate_order_excel(edited_df, table_mode=current_table_mode, group_type=current_group_type)
            st.download_button(
                label="📥 Excel でダウンロード",
                data=excel_bytes,
                file_name=f"{file_prefix}_{current_group_type}_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_excel_{current_table_mode}_{current_group_type}"
            )
        with col_d2:
            pdf_bytes = generate_order_pdf(edited_df, table_mode=current_table_mode, group_type=current_group_type)
            st.download_button(
                label=f"🖨️ A3 PDF ダウンロード（{group_axis_choice[:2]}）",
                data=pdf_bytes,
                file_name=f"{file_prefix}_{current_group_type}_{date.today().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                key=f"dl_pdf_{current_table_mode}_{current_group_type}"
            )
    else:
        st.warning("⚠️ 条件に一致するデータが見つかりませんでした。")
else:
    st.warning("⚠️ 対象データが見つかりませんでした。")