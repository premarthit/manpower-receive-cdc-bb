import streamlit as st
import pandas as pd
import textwrap
from pathlib import Path
import math

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="Manpower Dashboard",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="collapsed",
)


st.markdown("""
<style>
/* XD selected chips: override Streamlit/BaseWeb primary color */
[data-testid="stMultiSelect"] [data-baseweb="tag"] {
    background-color: #dff2d8 !important;
    background: #dff2d8 !important;
    border-color: #b8d9b0 !important;
}
</style>
""", unsafe_allow_html=True)

EXCEL_FILE = Path(__file__).parent / "Project.xlsx"

SHIFT_INFO = {
    "A100": {"name": "01.00", "bg": "#DDF3D5", "accent": "#7ED957"},
    "A600": {"name": "06.00", "bg": "#FFF1B8", "accent": "#FFD54F"},
    "A900": {"name": "09.00", "bg": "#F9D7E4", "accent": "#F45BA7"},
}
SHIFT_CODES = list(SHIFT_INFO.keys())
WORKING_CODES = set(SHIFT_CODES)
LEADER_ROLES = {"หัวหน้าหน่วย", "ผู้ช่วยหัวหน้าหน่วย"}

SPECIAL_NAMES = ["เรียกรถ", "ควบคุมงาน", "ไล่ของ", "แยกเสื่อม", "ดูแลหน้างาน"]

# ============================================================
# EXCEL
# ============================================================
@st.cache_data
def get_sheet_year_month_map(excel_mtime_ns):
    """
    อ่านชื่อ Sheet รูปแบบ MM-YY
    เช่น 08-69 = สิงหาคม 2569 / ค.ศ. 2026
    """
    workbook = pd.ExcelFile(EXCEL_FILE)

    sheet_map = {}

    for sheet_name in workbook.sheet_names:
        name = str(sheet_name).strip()
        parts = name.split("-")

        if len(parts) != 2:
            continue

        try:
            month_num = int(parts[0])
            year_yy = int(parts[1])
        except ValueError:
            continue

        if not (1 <= month_num <= 12 and 0 <= year_yy <= 99):
            continue

        # YY เป็นเลขปี พ.ศ. แบบ 2 หลัก
        # 69 = พ.ศ. 2569 = ค.ศ. 2026
        # 70 = พ.ศ. 2570 = ค.ศ. 2027
        gregorian_year = 1957 + year_yy

        sheet_map[(gregorian_year, month_num)] = name

    # รองรับ Sheet เดิมแบบ TIME ถ้ามี
    if not sheet_map and "TIME" in workbook.sheet_names:
        sheet_map[(pd.Timestamp.today().year, 0)] = "TIME"

    return sheet_map


@st.cache_data
def load_sheet_data(excel_mtime_ns, sheet_name, sheet_year=None, sheet_month=None):
    """อ่านข้อมูลจาก Sheet และสร้าง mapping วันที่ -> column"""
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    raw_dates = df.iloc[2, 4:]
    date_columns = {}

    for i, raw in enumerate(raw_dates):
        if pd.isna(raw):
            continue

        col_idx = 4 + i

        # กรณี Excel เป็นวันที่จริง
        parsed = pd.to_datetime(raw, errors="coerce", dayfirst=True)

        if pd.notna(parsed):
            parsed_date = parsed.date()

            # ถ้า Sheet มีชื่อปี/เดือนชัดเจน ให้บังคับปี/เดือนตามชื่อ Sheet
            if sheet_year is not None and sheet_month is not None:
                parsed_date = parsed_date.replace(
                    year=sheet_year,
                    month=sheet_month,
                )

            date_columns[parsed_date] = col_idx
            continue

        # กรณีหัวตารางเป็นเลขวัน เช่น 1, 2, 3 ... 31
        try:
            day_num = int(float(raw))
        except (TypeError, ValueError):
            continue

        if sheet_year is not None and sheet_month is not None:
            try:
                parsed_date = pd.Timestamp(
                    year=sheet_year,
                    month=sheet_month,
                    day=day_num,
                ).date()
                date_columns[parsed_date] = col_idx
            except ValueError:
                pass

    employees = df.iloc[3:].copy()
    employees.columns = range(employees.shape[1])

    employees = employees[
        employees[1].notna() & employees[3].notna()
    ].copy()

    employees[1] = (
        employees[1]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )
    employees[3] = employees[3].astype(str).str.strip()
    employees[4] = employees[4].astype(str).str.strip()

    return employees, date_columns


if not EXCEL_FILE.exists():
    st.error("ไม่พบไฟล์ Project.xlsx")
    st.stop()

excel_mtime_ns = EXCEL_FILE.stat().st_mtime_ns
sheet_year_month_map = get_sheet_year_month_map(excel_mtime_ns)

if not sheet_year_month_map:
    st.error("ไม่พบ Sheet รูปแบบ MM-YY เช่น 08-69 หรือ 09-69")
    st.stop()

# สร้างวันที่ที่มีอยู่จากทุก Sheet เช่น 08-69, 09-69
available_dates = []

for (sheet_year, sheet_month), sheet_name in sorted(sheet_year_month_map.items()):
    if sheet_month == 0:
        _, sheet_dates = load_sheet_data(
            excel_mtime_ns,
            sheet_name,
            sheet_year=None,
            sheet_month=None,
        )
    else:
        _, sheet_dates = load_sheet_data(
            excel_mtime_ns,
            sheet_name,
            sheet_year=sheet_year,
            sheet_month=sheet_month,
        )

    available_dates.extend(sheet_dates.keys())

available_dates = sorted(set(available_dates))

if not available_dates:
    st.error("ไม่พบวันที่ใน Sheet")
    st.stop()

today = pd.Timestamp.today().date()

# เปิดปฏิทินครอบคลุมทุกปีที่มี Sheet อยู่ใน Excel
data_years = sorted(set(d.year for d in available_dates))

if data_years:
    calendar_min_date = pd.Timestamp(min(data_years), 1, 1).date()
    calendar_max_date = pd.Timestamp(max(data_years), 12, 31).date()
else:
    calendar_min_date = pd.Timestamp(today.year, 1, 1).date()
    calendar_max_date = pd.Timestamp(today.year, 12, 31).date()

# เปิดเว็บแล้วให้เลือกวันนี้ ถ้ามีข้อมูลวันนี้
default_date = (
    today
    if calendar_min_date <= today <= calendar_max_date
    else calendar_max_date
)


# ============================================================
# MODERN HEADER + TOP CONTROL BAR
# ============================================================

# ============================================================
# MODERN HEADER + TOP CONTROL BAR
# ============================================================
st.markdown(
    '''
    <div class="modern-header">
        <div class="modern-header-title">MANPOWER RECEIVE CDC BB</div>
        <div class="modern-header-sub">MANPOWER CONTROL CENTER</div>
    </div>
    ''',
    unsafe_allow_html=True,
)

workload_area, date_area = st.columns([3.2, 1.0], gap="large")

with workload_area:
    workload_title_col, workload_value_col = st.columns(
        [1.0, 1.6],
        gap="small"
    )

    with workload_title_col:
        st.markdown(
            '<div class="control-label-box">Workload</div>',
            unsafe_allow_html=True,
        )

    current_qty = int(st.session_state.get("order_qty", 0))

    with workload_value_col:
        workload_text = st.text_input(
            "Workload",
            value=f"{current_qty:,}",
            label_visibility="collapsed",
            key="workload_text",
        )
        clean_qty = workload_text.replace(",", "").strip()
        if clean_qty.isdigit():
            order_qty = int(clean_qty)
            st.session_state["order_qty"] = order_qty
        else:
            # หลังสร้าง widget แล้ว ห้ามแก้ st.session_state["workload_text"]
            # โดยตรง เพราะ Streamlit จะเกิด StreamlitAPIException
            order_qty = current_qty



with date_area:
    selected_date = st.date_input(
        "วันที่",
        value=default_date,
        min_value=calendar_min_date,
        max_value=calendar_max_date,
        format="DD/MM/YYYY",
        label_visibility="collapsed",
        key="dashboard_date_picker",
    )
    st.markdown("</div>", unsafe_allow_html=True)

# เลือก Sheet ตามเดือน + ปีของวันที่
# ตัวอย่าง:
# 24/08/2026 -> Sheet 08-69
# 24/09/2026 -> Sheet 09-69
thai_year_short = (selected_date.year + 543) % 100
expected_sheet = f"{selected_date.month:02d}-{thai_year_short:02d}"

selected_month_sheet = sheet_year_month_map.get(
    (selected_date.year, selected_date.month)
)

if selected_month_sheet is None:
    st.error(
        f"ไม่พบ Sheet '{expected_sheet}' "
        f"สำหรับวันที่ {selected_date.strftime('%d/%m/%Y')}"
    )
    st.stop()

employees, date_columns = load_sheet_data(
    excel_mtime_ns,
    selected_month_sheet,
    sheet_year=selected_date.year,
    sheet_month=selected_date.month,
)

if selected_date not in date_columns:
    st.warning(
        f"ไม่พบวันที่ {selected_date.strftime('%d/%m/%Y')} "
        f"ใน Sheet '{selected_month_sheet}'"
    )
    st.stop()

date_col = date_columns[selected_date]

# ผู้จัดการแผนกแสดงเฉพาะในกล่อง "ผู้จัดการแผนก"
manager_mask = employees[4].astype(str).str.strip().str.contains(
    "ผู้จัดการแผนก", na=False
)
manager_employees = employees[manager_mask].copy()
operational_employees = employees[~manager_mask].copy()

working = operational_employees[
    operational_employees[date_col].astype(str).str.strip().isin(WORKING_CODES)
].copy()

working["shift_code"] = working[date_col].astype(str).str.strip()
working["person_id"] = (
    working[1].astype(str).str.replace(".0", "", regex=False).str.strip()
)
working["nickname"] = working[3].astype(str).str.strip()
working["position"] = working[4].astype(str).str.strip()

person_ids = working["person_id"].tolist()
id_to_row = {row["person_id"]: row for _, row in working.iterrows()}

all_id_to_row = {
    str(row[1]).replace(".0", "").strip(): row
    for _, row in employees.iterrows()
}

# ============================================================
# COUNTS
# ============================================================
total_people = len(operational_employees)
working_people = len(working)

shift_counts = {
    code: int((working["shift_code"] == code).sum())
    for code in SHIFT_CODES
}

all_day_values = operational_employees[date_col].astype(str).str.strip()
h_count = int(all_day_values.isin({"H", "H400"}).sum())
vacation_count = int(all_day_values.isin({"พ", "พักร้อน"}).sum())
day_count = int(all_day_values.isin({"ว", "วัน"}).sum())

# ============================================================
# HELPERS
# ============================================================
def display_person(pid):
    row = id_to_row.get(pid)
    if row is not None:
        return str(row["nickname"])

    all_row = all_id_to_row.get(str(pid))
    if all_row is not None:
        return str(all_row[3]).strip()

    return str(pid)


def role_badge(pid):
    row = id_to_row.get(pid)
    if row is None:
        return ""
    pos = str(row["position"]).strip()
    if pos in LEADER_ROLES:
        return pos
    return ""



# ============================================================
# GLOBAL CSS
# ============================================================
st.markdown("""
<style>
/* ซ่อนแถบ Header ของ Streamlit (Deploy / เมนู 3 จุด) */
[data-testid="stHeader"] {
    display: none !important;
}
[data-testid="stToolbar"] {
    display: none !important;
}

.workload-align-spacer {
    height: 58px;
}
/* ทำให้กล่อง Workload และช่องตัวเลขเท่ากันจริงทั้งชั้นนอกและชั้นใน */
.workload-title-wrap,
[data-testid="stTextInput"] {
    height: 58px !important;
    min-height: 58px !important;
    margin: 0 !important;
    padding: 0 !important;
    box-sizing: border-box !important;
}

[data-testid="stTextInput"] > div {
    height: 58px !important;
    min-height: 58px !important;
    margin: 0 !important;
    padding: 0 !important;
    box-sizing: border-box !important;
    display: flex !important;
    align-items: stretch !important;
}

[data-testid="stTextInput"] input {
    width: 100% !important;
    height: 58px !important;
    min-height: 58px !important;
    margin: 0 !important;
    box-sizing: border-box !important;
    border: 3px solid #111111 !important;
    border-radius: 16px !important;
    background: #ffffff !important;
}


/* ให้ขอบบน/ขอบล่างของ Workload และช่องตัวเลขตรงกัน */
.workload-title-wrap,
.workload-title-wrap .workload-inline-title,
.st-key-workload_text,
.st-key-workload_text > div,
.st-key-workload_text [data-testid="stTextInput"],
.st-key-workload_text [data-testid="stTextInput"] > div,
.st-key-workload_text [data-baseweb="base-input"],
.st-key-workload_text input {
    height: 58px !important;
    min-height: 58px !important;
    max-height: 58px !important;
    margin-top: 0 !important;
    margin-bottom: 0 !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
    box-sizing: border-box !important;
}

.st-key-workload_text {
    padding-bottom: 0 !important;
}

.st-key-workload_text [data-testid="stTextInput"] input {
    width: 100% !important;
    border: 3px solid #111111 !important;
    border-radius: 16px !important;
    background: #ffffff !important;
    color: #111111 !important;
    font-size: 28px !important;
    font-weight: 900 !important;
    text-align: center !important;
    box-sizing: border-box !important;
}

.workload-inline-title {
    width: 100%;
    height: 58px;
    border: 3px solid #111111;
    border-radius: 16px;
    background: #ffffff !important;
    height: 58px;
    min-height: 58px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #111111 !important;
    font-size: 28px;
    font-weight: 900;
    box-sizing: border-box;
}

.workload-inline-unit {
    height: 58px;
    min-height: 58px;
    display: flex;
    align-items: center;
    justify-content: flex-start;
    color: #111111 !important;
    font-size: 22px;
    font-weight: 900;
    white-space: nowrap;
    position: relative;
    left: -150px;
}

.workload-inline-title + div {
    margin-top: 0 !important;
}

[data-testid="stTextInput"] input {
    height: 58px !important;
    min-height: 58px !important;
    box-sizing: border-box !important;
    color: #111111 !important;
    font-size: 30px !important;
    font-weight: 900 !important;
    text-align: center !important;
    background: #ffffff !important;
    border: 3px solid #111111 !important;
    border-radius: 16px !important;
    padding: 8px 12px !important;
}

/* เอาขอบดำซ้าย/ขวาของช่อง Workload ออก */
.st-key-workload_text input,
.st-key-workload_text [data-testid="stTextInput"] input {
    border: 0 !important;
    outline: none !important;
    box-shadow: none !important;
    background: #ffffff !important;
}

.order-dashboard-card {
    width: 100%;
    box-sizing: border-box;
    background: #ffffff !important;
    border: 4px solid #111111;
    border-radius: 18px;
    padding: 10px 14px 10px 14px;
    text-align: center;
    margin: 8px 0 10px auto;
}

.order-dashboard-card [data-testid="stTextInput"] {
    margin-top: 4px;
}
.order-dashboard-card [data-testid="stTextInput"] input {
    color: #111111 !important;
    font-size: 20px !important;
    font-weight: 900 !important;
    text-align: center !important;
    background: #ffffff !important;
    border: 2px solid #111111 !important;
    border-radius: 12px !important;
    padding: 8px 10px !important;
}
.order-dashboard-title {
    color: #111111 !important;
    font-size: 26px;
    font-weight: 900;
    line-height: 1.2;
    margin-bottom: 6px;
}

.order-dashboard-unit {
    color: #111111 !important;
    font-size: 20px;
    font-weight: 900;
    margin-top: 4px;
}

[data-testid="stNumberInput"] > div {
    background: #ffffff !important;
}

[data-testid="stNumberInput"] input {
    color: #111111 !important;
    font-size: 28px !important;
    font-weight: 900 !important;
    text-align: center !important;
    background: #ffffff !important;
    border: 2px solid #111111 !important;
    border-radius: 12px !important;
}
.date-dashboard-card {
    width: 100%;
    box-sizing: border-box;
    background: transparent !important;
    border: 0 !important;
    border-radius: 0;
    padding: 0 !important;
    text-align: center;
    margin: 0 0 10px auto;
    overflow: hidden !important;
}

.date-dashboard-title {
    color: #111111 !important;
    font-size: 28px;
    font-weight: 900;
    line-height: 1.2;
    margin-bottom: 6px;
}

[data-testid="stDateInput"],
[data-testid="stDateInput"] > div,
[data-testid="stDateInput"] > div > div {
    background: transparent !important;
    border: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
    box-shadow: none !important;
    overflow: hidden !important;
}

[data-testid="stDateInput"] input {
    color: #111111 !important;
    font-size: 28px !important;
    font-weight: 900 !important;
    text-align: center !important;
    background: #ffffff !important;
    border: 2px solid #111111 !important;
    border-radius: 12px !important;
    padding: 8px 10px !important;
}

[data-testid="stDateInput"] button {
    color: #111111 !important;
}

/* =========================================================
   THEME ของช่องเลือกทั้งหมดทั้งหน้า
   ========================================================= */
[data-baseweb="select"] {
    background: #eaf2f7 !important;
}

[data-baseweb="select"] > div {
    min-height: 40px !important;
    border: 2px solid #a9becb !important;
    border-radius: 9px !important;
    background: #eaf2f7 !important;
    box-shadow: none !important;
}

[data-baseweb="select"] input,
[data-baseweb="select"] span {
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 800 !important;
}

[data-baseweb="select"] svg {
    color: #52666f !important;
}

/* ชื่อที่เลือกแล้วทุกช่อง */
[data-baseweb="tag"] {
    background: #dff2d8 !important;
    background-color: #dff2d8 !important;
    background-image: none !important;
    color: #263238 !important;
    border-radius: 7px !important;
    padding: 4px 8px !important;
    font-weight: 900 !important;
}

[data-baseweb="tag"] span {
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 900 !important;
}

[data-baseweb="tag"] svg {
    color: #52666f !important;
}

/* ตัวเลือกใน dropdown */
[role="option"] {
    background: #f8fbfd !important;
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 700 !important;
}

[role="option"]:hover,
[role="option"][aria-selected="true"] {
    background: #dcecf5 !important;
    color: #263238 !important;
}

/* ช่องค้นหาข้อความ */
[data-testid="stTextInput"] input {
    background: #f8fbfd !important;
    border: 2px solid #c5d0d6 !important;
    border-radius: 9px !important;
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 700 !important;
}


/* =========================================================
   FORCE XD SELECTED NAME TAG = LIGHT GREEN
   ========================================================= */
div[data-baseweb="tag"],
div[data-baseweb="tag"][class],
.xd-cell div[data-baseweb="tag"],
.xd-cell div[data-baseweb="tag"][class] {
    background-color: #dff2d8 !important;
    background-image: none !important;
    color: #263238 !important;
    border: 1px solid #b8d9b0 !important;
    border-radius: 7px !important;
    box-shadow: none !important;
}

div[data-baseweb="tag"] *,
.xd-cell div[data-baseweb="tag"] * {
    color: #263238 !important;
}

div[data-baseweb="tag"] span,
.xd-cell div[data-baseweb="tag"] span {
    background-color: transparent !important;
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 900 !important;
}

div[data-baseweb="tag"] svg,
.xd-cell div[data-baseweb="tag"] svg {
    color: #52666f !important;
    fill: #52666f !important;
}

html, body, [data-testid="stAppViewContainer"] {
    background: #B5D3E8 !important;
}
[data-testid="stHeader"] {
    background: transparent !important;
}

.block-container {
    max-width: 1800px !important;
    width: 96vw !important;
    margin: 0 auto !important;
    padding-top: 0.8rem !important;
    padding-bottom: 1.5rem !important;
    background: transparent !important;
}

h1, h2, h3, p, span, div, label {
    color: #111111;
}

.top-title {
    width: fit-content;
    margin: 0 auto 8px auto;
    border: 4px solid #ffffff;
    border-radius: 18px;
    padding: 8px 28px;
    font-size: clamp(22px, 2vw, 34px);
    font-weight: 900;
    letter-spacing: .3px;
    background: #ffffff;
}


.shift-time {
    flex: 0 0 38%;
}

.red-stat-label {
    flex: 0 0 38%;
}

.section-title {
    font-size: 22px;
    font-weight: 900;
    margin: 12px 0 6px 0;
    color: #111111 !important;
}

.muted {
    font-size: 12px;
    color: #666666 !important;
    margin-bottom: 7px;
}

.role-badge-static {
    display: inline-block;
    background: #FF3B3B;
    color: #ffffff !important;
    border-radius: 6px;
    padding: 3px 7px;
    margin-top: 4px;
    font-size: 10px;
    font-weight: 900;
}

/* ===== ส่วนลากวาง ===== */
.sortable-component {
    width: 100% !important;
    display: grid !important;
    grid-template-columns: repeat(5, minmax(150px, 1fr)) !important;
    gap: 10px !important;
    padding: 8px !important;
    background: transparent !important;
    box-sizing: border-box !important;
    border-radius: 0 !important;
}

.sortable-container {
    min-width: 0 !important;
    min-height: 96px !important;
    border: 2px solid #111111 !important;
    border-radius: 14px !important;
    overflow: hidden !important;
    background: #f6f6f6 !important;
    box-sizing: border-box !important;
}

.sortable-container-header {
    min-height: 34px !important;
    padding: 6px 9px !important;
    font-size: 14px !important;
    font-weight: 900 !important;
    color: #111111 !important;
    background: #d8d8d8 !important;
}

.sortable-container-body {
    min-height: 58px !important;
    padding: 6px !important;
    background: #f0f0f0 !important;
}

.sortable-item {
    color: #111111 !important;
    background: #ffffff !important;
    border: 2px solid #111111 !important;
    border-radius: 8px !important;
    padding: 6px 8px !important;
    margin: 3px !important;
    font-size: 12px !important;
    font-weight: 800 !important;
}

/* กะ 06 / 09 / 01 ตามลำดับที่ component แสดงจริง */
.sortable-container:nth-child(1) .sortable-container-header,
.sortable-container:nth-child(1) .sortable-item { background: #FFF1B8 !important; }

.sortable-container:nth-child(2) .sortable-container-header,
.sortable-container:nth-child(2) .sortable-item { background: #F9D7E4 !important; }

.sortable-container:nth-child(3) .sortable-container-header,
.sortable-container:nth-child(3) .sortable-item { background: #DDF3D5 !important; }

/* XD: 5 ช่องต่อแถว */
.sortable-container:nth-child(4)  { grid-column:1!important; grid-row:2!important; }
.sortable-container:nth-child(5)  { grid-column:2!important; grid-row:2!important; }
.sortable-container:nth-child(6)  { grid-column:3!important; grid-row:2!important; }
.sortable-container:nth-child(7)  { grid-column:4!important; grid-row:2!important; }
.sortable-container:nth-child(8)  { grid-column:5!important; grid-row:2!important; }

.sortable-container:nth-child(9)  { grid-column:1!important; grid-row:3!important; }
.sortable-container:nth-child(10) { grid-column:2!important; grid-row:3!important; }
.sortable-container:nth-child(11) { grid-column:3!important; grid-row:3!important; }
.sortable-container:nth-child(12) { grid-column:4!important; grid-row:3!important; }
.sortable-container:nth-child(13) { grid-column:5!important; grid-row:3!important; }

.sortable-container:nth-child(14) { grid-column:1!important; grid-row:4!important; }
.sortable-container:nth-child(15) { grid-column:2!important; grid-row:4!important; }
.sortable-container:nth-child(16) { grid-column:3!important; grid-row:4!important; }
.sortable-container:nth-child(17) { grid-column:4!important; grid-row:4!important; }
.sortable-container:nth-child(18) { grid-column:5!important; grid-row:4!important; }

.sortable-container:nth-child(19) { grid-column:1!important; grid-row:5!important; }
.sortable-container:nth-child(20) { grid-column:2!important; grid-row:5!important; }
.sortable-container:nth-child(21) { grid-column:3!important; grid-row:5!important; }
.sortable-container:nth-child(22) { grid-column:4!important; grid-row:5!important; }
.sortable-container:nth-child(23) { grid-column:5!important; grid-row:5!important; }

@media (max-width: 1200px) {
    .block-container {
        width: 98vw !important;
    }
    .sortable-component {
        grid-template-columns: repeat(4, minmax(140px, 1fr)) !important;
    }
}

@media (max-width: 900px) {
    .block-container {
        width: 98vw !important;
    }
    .sortable-component {
        grid-template-columns: repeat(3, minmax(130px, 1fr)) !important;
    }
}

@media (max-width: 650px) {
    .sortable-component {
        grid-template-columns: repeat(2, minmax(120px, 1fr)) !important;
    }
}

/* =========================================================
   MODERN DASHBOARD LAYOUT
   ========================================================= */

/* ลดช่องว่างระหว่าง Workload/วันที่ กับกล่องสรุปด้านบน */
[data-testid="stHorizontalBlock"]:has(.st-key-workload_text) {
    margin-bottom: -42px !important;
}

.modern-summary-grid {
    margin-top: 8px !important;
}


/* เว้นระยะระหว่างกราฟสัดส่วน/แนวโน้มกับปุ่มรายชื่อ */
.st-key-toggle_shift_people_ button {
    margin-top: 4px !important;
}

.modern-header {
    background: linear-gradient(135deg, #234A8A 0%, #173B73 55%, #102F5D 100%);
    border-radius: 18px;
    padding: 18px 24px;
    color: #fff !important;
    box-shadow: 0 10px 28px rgba(35, 63, 128, 0.18);
    margin-bottom: 18px;
}
.modern-header-title {
    color: #fff !important;
    font-size: 32px;
    font-weight: 900;
}
.modern-header-sub {
    color: rgba(255,255,255,.84) !important;
    font-size: 13px;
    font-weight: 800;
    margin-top: 3px;
}
.control-title {
    font-size: 13px;
    font-weight: 900;
    color: #49628f !important;
    margin-bottom: 6px;
}
.control-label-box {
    height: 58px;
    display: flex;
    align-items: center;
    justify-content: center;
    border: 3px solid #1c2750;
    border-radius: 14px;
    background: #fff;
    font-size: 28px;
    font-weight: 900;
    box-sizing: border-box;
}
/* หน่วย "ชิ้น" อยู่ติดด้านขวาของช่อง Workload */
/* ให้ "ชิ้น" อ้างอิงจากขอบขวาของกล่องตัวเลขโดยตรง */
.st-key-workload_text {
    position: relative !important;
    width: 400px !important;
    max-width: 400px !important;
    min-width: 400px !important;
    display: inline-block !important;
}
.st-key-workload_text::after {
    content: "ชิ้น";
    position: absolute;
    left: calc(100% + 6px);
    top: 50%;
    transform: translateY(-50%);
    color: #1c2750 !important;
    font-size: 20px;
    font-weight: 900;
    white-space: nowrap;
    pointer-events: none;
}
.st-key-workload_text [data-testid="stTextInput"],
.st-key-workload_text [data-baseweb="base-input"],
.st-key-workload_text input {
    width: 100% !important;
    max-width: 400px !important;
    box-sizing: border-box !important;
}
[data-testid="stTextInput"] input {
    height: 58px !important;
    border: 3px solid #1c2750 !important;
    border-radius: 14px !important;
    background: #fff !important;
    color: #17213f !important;
    font-size: 28px !important;
    font-weight: 900 !important;
    text-align: center !important;
}
[data-testid="stDateInput"] input {
    height: 58px !important;
    border: 3px solid #1c2750 !important;
    border-radius: 14px !important;
    background: #fff !important;
    color: #17213f !important;
    font-size: 24px !important;
    font-weight: 900 !important;
    text-align: center !important;
}

.manager-dashboard-card {
    overflow: hidden;
}

.manager-row {
    min-height: 42px;
    border-radius: 18px;
    padding: 0 12px;
    margin-bottom: 8px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    font-size: 20px;
    font-weight: 900;
}

.manager-row span {
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.manager-row strong {
    white-space: nowrap;
    font-size: 20px;
    font-weight: 900;
}


.manager-row span,
.manager-row strong {
    font-size: 20px !important;
    font-weight: 900 !important;
}

.manager-green {
    background: #7ed957 !important;
    color: #111111 !important;
}

.manager-gray {
    background: #d9dee5 !important;
    color: #3f4650 !important;
}

.manager-red {
    background: #ff3b3b !important;
    color: #ffffff !important;
    animation: manager-red-blink 0.8s infinite alternate;
}

.manager-red span,
.manager-red strong {
    color: #ffffff !important;
}

@keyframes manager-red-blink {
    from {
        background: #ff3b3b;
        box-shadow: 0 0 4px rgba(255,59,59,.35);
    }
    to {
        background: #fff0f0;
        box-shadow:
            0 0 18px rgba(255,59,59,.95),
            0 0 30px rgba(255,59,59,.65),
            0 0 42px rgba(255,59,59,.35);
    }
}

.manager-empty {
    color: #667085 !important;
    font-size: 14px;
    font-weight: 800;
    padding: 10px 4px;
}

.modern-card {
    background: rgba(255,255,255,.93);
    border: 1px solid rgba(74,101,160,.18);
    border-radius: 18px;
    padding: 16px;
    box-shadow: 0 10px 26px rgba(61,88,144,.10);
    box-sizing: border-box;
}

.modern-summary-grid {
    display: grid;
    grid-template-columns: 1.15fr 0.95fr 1.15fr 1.15fr;
    gap: 18px;
    margin-top: 18px;
}
.modern-summary-grid .modern-card {
    min-width: 0;
}
.ratio-modern-body {
    display: grid;
    grid-template-columns: 0.9fr 1.1fr;
    gap: 28px;
    align-items: center;
}
.card-title {
    font-size: 22px;
    font-weight: 900;
    color: #23335f !important;
    margin-bottom: 12px;
}
.shift-mini-row, .status-mini-row {
    min-height: 54px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 14px;
    margin-bottom: 9px;
    font-size: 20px;
    font-weight: 900;
}
.green-soft { background: #dff2d8; }
.yellow-soft { background: #fff1b8; }
.pink-soft { background: #f7dbe7; }

.shift-green {
    background: #63c16b !important;
    color: #111111 !important;
}
.shift-yellow {
    background: #f4c542 !important;
    color: #111111 !important;
}
.shift-red {
    background: #ff4b4b !important;
    color: #ffffff !important;
    animation: shift-red-blink 0.8s infinite alternate;
}
.shift-red span,
.shift-red strong {
    color: #ffffff !important;
}
.shift-neutral {
    background: #e5e7eb !important;
    color: #111111 !important;
}

@keyframes shift-red-blink {
    from {
        background: #ff3b3b;
        box-shadow: 0 0 4px rgba(255,59,59,.35);
    }
    to {
        background: #fff0f0;
        box-shadow:
            0 0 18px rgba(255,59,59,.95),
            0 0 30px rgba(255,59,59,.65),
            0 0 42px rgba(255,59,59,.35);
    }
}
.status-mini-row {
    background: #ff5b5b;
    color: #fff !important;
}
.status-mini-row span, .status-mini-row strong { color: #fff !important; }

.working-kpi {
    height: 100px;
    border-radius: 16px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 38px;
    font-weight: 900;
    color: #fff !important;
    margin-bottom: 10px;
}
.stat-work-green { background: #63ba6d !important; }
.stat-work-orange { background: #f3bf73 !important; color: #fff !important; }
.stat-work-red {
    background: #ff3b3b !important;
    color: #ffffff !important;
    animation: shift-red-blink 0.8s infinite alternate;
}

@keyframes status-blink {
    from { filter: brightness(1); }
    to { filter: brightness(1.28); }
}
.working-text {
    font-size: 15px;
    font-weight: 900;
    color: #56647e !important;
    text-align: center;
    margin-bottom: 12px;
}
.total-kpi {
    background: #4e5865;
    color: #fff !important;
    border-radius: 12px;
    padding: 11px 12px;
    text-align: center;
    font-size: 18px;
    font-weight: 900;
}

.ratio-left-panel {
    background: rgba(255,255,255,.50);
    border-radius: 16px;
    padding: 14px 10px;
    min-height: 360px;
    box-sizing: border-box;
}
.ratio-left-content {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    gap: 36px;
    min-height: 325px;
}
.ratio-left-content .ratio-pie-modern {
    width: 270px !important;
    height: 270px !important;
    flex: 0 0 270px;
    margin: 0;
}
.ratio-left-content .ratio-modern-legend {
    height: auto;
    gap: 22px;
    font-size: 20px;
}
.ratio-left-content .legend-dot {
    width: 16px;
    height: 16px;
    margin-right: 10px;
}
.ratio-trend-title {
    font-size: 20px;
    font-weight: 900;
    color: #23335f !important;
    margin: 4px 0 8px 4px;
}

.manpower-line-card {
    height: 285px;
    background: rgba(255,255,255,.92);
    border-radius: 12px;
    padding: 10px 10px 8px 4px;
    box-sizing: border-box;
    display: flex;
    gap: 8px;
    overflow: hidden;
}

.manpower-line-yaxis {
    width: 38px;
    flex: 0 0 38px;
    height: 240px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    align-items: flex-end;
    padding: 5px 2px 18px 0;
    box-sizing: border-box;
    color: #61708d !important;
    font-size: 11px;
    font-weight: 800;
}

.manpower-line-plot {
    position: relative;
    flex: 1 1 auto;
    min-width: 0;
    height: 260px;
}

.manpower-line-svg {
    position: absolute !important;
    left: 0 !important;
    top: 0 !important;
    width: 100% !important;
    height: 225px !important;
    display: block !important;
    overflow: visible !important;
}

.manpower-line-grid {
    stroke: rgba(101,122,157,.22);
    stroke-width: 1;
    vector-effect: non-scaling-stroke;
}

.manpower-line-path {
    fill: none;
    stroke: #1878c9;
    stroke-width: 4;
    stroke-linecap: round;
    stroke-linejoin: round;
    vector-effect: non-scaling-stroke;
}

.manpower-line-point {
    fill: #1878c9;
    stroke: #ffffff;
    stroke-width: 2;
    vector-effect: non-scaling-stroke;
}

.manpower-line-point.today {
    fill: #0f58a8;
    stroke-width: 3;
}

.manpower-line-number {
    fill: #29456f;
    font-size: 12px;
    font-weight: 900;
    font-family: Arial, sans-serif;
}

.manpower-line-date-row {
    position: absolute !important;
    left: 0 !important;
    right: 0 !important;
    bottom: 0 !important;
    height: 34px !important;
    display: flex !important;
    align-items: flex-start !important;
    justify-content: space-between !important;
    padding: 0 3px !important;
    box-sizing: border-box !important;
}

.manpower-line-date {
    width: 11%;
    text-align: center;
    color: #53627d !important;
    font-size: 10px !important;
    font-weight: 800 !important;
    white-space: nowrap !important;
    transform: rotate(-30deg);
    transform-origin: top center;
}

.ratio-trend-card {
    width: 100%;
    display: grid;
    grid-template-columns: 1fr 1.25fr;
    gap: 18px;
    margin-top: 18px;
}

.ratio-card,
.manpower-trend-card {
    width: 100%;
    min-width: 0;
    background: rgba(255,255,255,.94);
    border: 1px solid rgba(74,101,160,.18);
    border-radius: 18px;
    padding: 16px 18px;
    box-shadow: 0 10px 26px rgba(61,88,144,.10);
    box-sizing: border-box;
}

.ratio-card-title,
.manpower-trend-title {
    font-size: 20px;
    font-weight: 900;
    color: #23335f !important;
    margin: 2px 0 12px 2px;
}

.ratio-card-body {
    display: flex;
    align-items: center;
    justify-content: flex-start;
    gap: 34px;
    min-height: 315px;
}

.ratio-card .ratio-pie-svg {
    width: 270px !important;
    height: 270px !important;
    flex: 0 0 270px !important;
    display: block !important;
}

.ratio-card .ratio-modern-legend {
    display: flex;
    flex-direction: column;
    justify-content: center;
    gap: 20px;
    font-size: 18px;
    font-weight: 900;
}

.ratio-card .legend-dot {
    width: 16px;
    height: 16px;
    margin-right: 10px;
}

.manpower-trend-card .manpower-line-card {
    width: 100%;
    height: 300px;
    background: rgba(255,255,255,.72);
    border-radius: 12px;
    padding: 10px 10px 8px 4px;
    box-sizing: border-box;
    display: flex;
    gap: 8px;
    overflow: hidden;
}

@media (max-width: 1000px) {
    .ratio-trend-card {
        grid-template-columns: 1fr;
    }
    .ratio-card-body {
        justify-content: center;
        flex-wrap: wrap;
    }
}

.ratio-pie-modern {
    width: 190px;
    height: 190px;
    margin: 0;
    flex: 0 0 190px;
}
.ratio-side-content .ratio-modern-legend {
    height: auto;
    gap: 14px;
    font-size: 16px;
}
.trend-side .manpower-line-card {
    width: 100%;
    height: 285px;
    padding: 10px 10px 8px 4px;
    background: rgba(255,255,255,.72);
}
.trend-side .trend-yaxis {
    height: 225px;
    padding-bottom: 30px;
}
.trend-side .trend-chart-area {
    height: 255px;
}
.trend-side .trend-bars {
    height: 235px;
}
.trend-side .trend-bar-wrap {
    height: 235px;
    min-width: 34px;
}
.trend-side .trend-value {
    font-size: 10px;
    margin-bottom: 4px;
}
.trend-side .trend-bar {
    min-width: 22px;
    max-width: 40px;
    width: 68%;
}
.trend-side .trend-label {
    font-size: 10px;
    height: 28px;
}
@media (max-width: 1200px) {
    .modern-summary-grid {
        grid-template-columns: 1fr 1fr !important;
    }
}

@media (max-width: 1000px) {
    .ratio-trend-card {
        grid-template-columns: 1fr;
    }
}
.ratio-pie-modern {
    width: 220px;
    height: 220px;
    border-radius: 50%;
    border: 5px solid #1d2c52;
    margin: 8px auto 4px auto;
}
.ratio-modern-legend {
    height: 100%;
    display: flex;
    flex-direction: column;
    justify-content: center;
    gap: 14px;
    font-size: 18px;
    font-weight: 900;
}
.legend-dot {
    display: inline-block;
    width: 14px;
    height: 14px;
    border-radius: 50%;
    margin-right: 8px;
    vertical-align: middle;
}
.dot-green { background:#7DBB7A; }
.dot-yellow { background:#E5C96A; }
.dot-pink { background:#D79AB2; }
.dot-red { background:#9AA8C4; }

</style>
""", unsafe_allow_html=True)


# ============================================================


# ============================================================
# ผู้จัดการแผนก
# ============================================================
def manager_status(row):
    value = str(row[date_col]).strip()

    manager_shift_times = {
        "A100": "01.00",
        "A200": "02.00",
        "A400": "04.00",
        "A500": "05.00",
        "A600": "06.00",
        "A800": "08.00",
        "A900": "09.00",
    }

    if value in manager_shift_times:
        return manager_shift_times[value], "manager-green"

    if value in {"H", "H400"}:
        return "H", "manager-red"

    if value in {"สล", "SL"}:
        return "สล", "manager-red"

    if value in {"พ", "พักร้อน"}:
        return "พักร้อน", "manager-gray"

    return (value if value else "-"), "manager-gray"


manager_rows_html = ""
for _, manager_row in manager_employees.iterrows():
    full_name = str(manager_row[2]).strip()
    status_text, status_class = manager_status(manager_row)
    manager_rows_html += (
        f'<div class="manager-row {status_class}">'
        f'<span>{full_name}</span>'
        f'<strong>{status_text}</strong>'
        f'</div>'
    )

if not manager_rows_html:
    manager_rows_html = '<div class="manager-empty">ไม่พบผู้จัดการแผนก</div>'

# MAIN SUMMARY CARDS
# ============================================================
# สัดส่วน 4 กลุ่มที่แสดงใน Pie Chart
# ใช้ผลรวมของ 4 กลุ่มเป็นฐาน เพื่อให้กราฟรวมกันครบ 100%
status_ratio_count = h_count + vacation_count + day_count
ratio_total = max(
    shift_counts["A100"]
    + shift_counts["A600"]
    + shift_counts["A900"]
    + status_ratio_count,
    1,
)

ratio_a100 = shift_counts["A100"] / ratio_total * 100
ratio_a600 = shift_counts["A600"] / ratio_total * 100
ratio_a900 = shift_counts["A900"] / ratio_total * 100
ratio_h = status_ratio_count / ratio_total * 100

# โทนสีซอฟต์ ลดความแสบตา
RATIO_GREEN = "#7DBB7A"
RATIO_YELLOW = "#E5C96A"
RATIO_PINK = "#D79AB2"
RATIO_STATUS = "#9AA8C4"

# กราฟวงกลมแบบ SVG เพื่อให้เส้นแบ่งเป็น "เส้น" จริง ไม่เป็นลิ่ม
RATIO_DIVIDER = "#1d2c52"

pie_values = [
    ("01.00", ratio_a100, RATIO_GREEN),
    ("06.00", ratio_a600, RATIO_YELLOW),
    ("09.00", ratio_a900, RATIO_PINK),
    ("H + พ + ตรงวัน", ratio_h, RATIO_STATUS),
]

# ปรับเศษทศนิยมให้กลุ่มสุดท้ายปิดวงกลมพอดี
pie_values[-1] = (
    pie_values[-1][0],
    max(0, 100.0 - sum(v for _, v, _ in pie_values[:-1])),
    pie_values[-1][2],
)

def polar(cx, cy, r, angle_deg):
    a = math.radians(angle_deg)
    return cx + r * math.cos(a), cy + r * math.sin(a)

cx = cy = 150
r = 142
pie_paths = []
angle = -90.0

for _, pct, fill in pie_values:
    extent = (pct / 100.0) * 360.0
    end_angle = angle + extent

    x1, y1 = polar(cx, cy, r, angle)
    x2, y2 = polar(cx, cy, r, end_angle)
    large_arc = 1 if extent > 180 else 0

    path = (
        f"M {cx},{cy} "
        f"L {x1:.2f},{y1:.2f} "
        f"A {r},{r} 0 {large_arc} 1 {x2:.2f},{y2:.2f} Z"
    )

    pie_paths.append(
        f'<path d="{path}" fill="{fill}" '
        f'stroke="{RATIO_DIVIDER}" stroke-width="3" '
        f'stroke-linejoin="round"/>'
    )

    angle = end_angle

pie_svg = (
    '<svg class="ratio-pie-svg" viewBox="0 0 300 300" '
    'xmlns="http://www.w3.org/2000/svg" '
    'role="img" aria-label="สัดส่วนกำลังคน">'
    + "".join(pie_paths)
    + f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" '
      f'stroke="{RATIO_DIVIDER}" stroke-width="5"/>'
    + '</svg>'
)

if working_people >= 39:
    work_class = "stat-work-green"
    work_text = "กำลังคนเพียงพอ"
elif 37 <= working_people <= 38:
    work_class = "stat-work-orange"
    work_text = "กำลังคนปานกลาง"
else:
    work_class = "stat-work-red"
    work_text = "กำลังคนน้อยต้องเฝ้าระวัง"

# เงื่อนไขสีของจำนวนคนแต่ละกะ
def shift_status_class(code, count):
    if code == "A100":
        if count < 18:
            return "shift-red"
        elif 18 <= count <= 19:
            return "shift-yellow"
        return "shift-green"

    if code == "A600":
        if count < 3:
            return "shift-red"
        elif 3 <= count <= 4:
            return "shift-yellow"
        return "shift-green"

    if code == "A900":
        if count < 13:
            return "shift-red"
        elif 13 <= count <= 14:
            return "shift-yellow"
        return "shift-green"

    return "shift-neutral"

shift_a100_class = shift_status_class("A100", shift_counts["A100"])
shift_a600_class = shift_status_class("A600", shift_counts["A600"])
shift_a900_class = shift_status_class("A900", shift_counts["A900"])

# กล่องสรุปด้านบน
summary_html = f"""<div class="modern-summary-grid">
<div class="modern-card">
<div class="card-title">กำลังคนตามกะ</div>
<div class="shift-mini-row {shift_a100_class}"><span>01.00</span><strong>{shift_counts["A100"]} คน</strong></div>
<div class="shift-mini-row {shift_a600_class}"><span>06.00</span><strong>{shift_counts["A600"]} คน</strong></div>
<div class="shift-mini-row {shift_a900_class}"><span>09.00</span><strong>{shift_counts["A900"]} คน</strong></div>
</div>
<div class="modern-card">
<div class="card-title">มาทำงาน</div>
<div class="working-kpi {work_class}">{working_people} คน</div>
<div class="working-text">{work_text}</div>
<div class="total-kpi">กำลังคนทั้งหมด: {total_people} คน</div>
</div>
<div class="modern-card">
<div class="card-title">สถานะ</div>
<div class="status-mini-row"><span>H</span><strong>{h_count} คน</strong></div>
<div class="status-mini-row"><span>พักร้อน</span><strong>{vacation_count} คน</strong></div>
<div class="status-mini-row"><span>ตรงวัน</span><strong>{day_count} คน</strong></div>
</div>
<div class="modern-card manager-dashboard-card">
<div class="card-title">ผู้จัดการแผนก</div>
{manager_rows_html}
</div>
</div>"""
st.markdown(summary_html, unsafe_allow_html=True)

# ------------------------------------------------------------
# ข้อมูลกราฟแนวโน้ม ±4 วัน
# ------------------------------------------------------------
trend_dates = []
trend_counts = []
trend_labels = []

for offset in range(-4, 5):
    d = selected_date + pd.Timedelta(days=offset)
    if d not in date_columns:
        continue

    col = date_columns[d]
    day_values = employees[col].astype(str).str.strip()
    count = int(day_values.isin(WORKING_CODES).sum())

    trend_dates.append(d)
    trend_counts.append(count)
    trend_labels.append(d.strftime("%d/%m"))

max_count = max(max(trend_counts, default=0), 1)
chart_w = 1000
chart_h = 240
left_pad = 18
right_pad = 18
top_pad = 18
bottom_pad = 24
plot_w = chart_w - left_pad - right_pad
plot_h = chart_h - top_pad - bottom_pad

points = []
n = len(trend_counts)

for idx, (count, d) in enumerate(zip(trend_counts, trend_dates)):
    x = left_pad if n <= 1 else left_pad + (plot_w * idx / (n - 1))
    y = top_pad + plot_h - (count / max_count * plot_h)
    points.append((x, y, count, d))

polyline = " ".join(f"{x:.2f},{y:.2f}" for x, y, _, _ in points)

svg_items = [
    f'<line class="manpower-line-grid" x1="{left_pad}" y1="{top_pad}" '
    f'x2="{chart_w-right_pad}" y2="{top_pad}"/>',
    f'<line class="manpower-line-grid" x1="{left_pad}" y1="{top_pad + plot_h*0.25:.2f}" '
    f'x2="{chart_w-right_pad}" y2="{top_pad + plot_h*0.25:.2f}"/>',
    f'<line class="manpower-line-grid" x1="{left_pad}" y1="{top_pad + plot_h*0.50:.2f}" '
    f'x2="{chart_w-right_pad}" y2="{top_pad + plot_h*0.50:.2f}"/>',
    f'<line class="manpower-line-grid" x1="{left_pad}" y1="{top_pad + plot_h*0.75:.2f}" '
    f'x2="{chart_w-right_pad}" y2="{top_pad + plot_h*0.75:.2f}"/>',
    f'<polyline class="manpower-line-path" points="{polyline}"/>',
]

for x, y, count, d in points:
    today_cls = " today" if d == selected_date else ""
    svg_items.append(
        f'<circle class="manpower-line-point{today_cls}" cx="{x:.2f}" '
        f'cy="{y:.2f}" r="{7 if d == selected_date else 5.5}"/>'
    )
    label_y = max(14, y - 10)
    svg_items.append(
        f'<text class="manpower-line-number" x="{x:.2f}" y="{label_y:.2f}" '
        f'text-anchor="middle">{count}</text>'
    )

date_html = "".join(
    f'<div class="manpower-line-date">{d.strftime("%d/%m")}</div>'
    for _, _, _, d in points
)

trend_svg = "".join(svg_items)

ratio_trend_html = f"""<div class="ratio-trend-card">

<div class="ratio-card">
<div class="ratio-card-title">สัดส่วนกำลังคน</div>
<div class="ratio-card-body">
{pie_svg}
<div class="ratio-modern-legend">
<div><span class="legend-dot dot-green"></span>01.00 — {shift_counts["A100"]} คน ({ratio_a100:.1f}%)</div>
<div><span class="legend-dot dot-yellow"></span>06.00 — {shift_counts["A600"]} คน ({ratio_a600:.1f}%)</div>
<div><span class="legend-dot dot-pink"></span>09.00 — {shift_counts["A900"]} คน ({ratio_a900:.1f}%)</div>
<div><span class="legend-dot dot-red"></span>H + พ + ตรงวัน — {status_ratio_count} คน ({ratio_h:.1f}%)</div>
</div>
</div>
</div>

<div class="manpower-trend-card">
<div class="manpower-trend-title">แนวโน้มกำลังคน</div>
<div class="manpower-line-card">
<div class="manpower-line-yaxis">
<span>{max_count}</span>
<span>{int(max_count*.75)}</span>
<span>{int(max_count*.50)}</span>
<span>{int(max_count*.25)}</span>
<span>0</span>
</div>
<div class="manpower-line-plot">
<svg class="manpower-line-svg" viewBox="0 0 {chart_w} {chart_h}" preserveAspectRatio="none">
{trend_svg}
</svg>
<div class="manpower-line-date-row">
{date_html}
</div>
</div>
</div>
</div>

</div>"""

st.markdown(ratio_trend_html, unsafe_allow_html=True)

# ============================================================
# DRAG DROP DATA
# ============================================================
reset_version_key = f"reset_version_{selected_date.isoformat()}"
if reset_version_key not in st.session_state:
    st.session_state[reset_version_key] = 0

widget_version = st.session_state[reset_version_key]
state_key = f"visual_assignments_{selected_date.isoformat()}"

def default_assignments():
    data = {
        "shift_A600": working.loc[
            working["shift_code"] == "A600", "person_id"
        ].tolist(),
        "shift_A900": working.loc[
            working["shift_code"] == "A900", "person_id"
        ].tolist(),
        **{f"XD{i}": [] for i in range(1, 21)},
        **{f"special_{name}": [] for name in SPECIAL_NAMES},
        "shift_A100": working.loc[
            working["shift_code"] == "A100", "person_id"
        ].tolist(),
    }
    return data

if state_key not in st.session_state:
    st.session_state[state_key] = default_assignments()

# IMPORTANT: component renders the last source container first in this setup.
# Put A100 last so visual order is A100, A600, A900, XD1...XD20, specials.



# ============================================================
# สถิติการมาทำงานรายเดือนของแต่ละคน
# ============================================================
month_dates = sorted(
    d for d in date_columns.keys()
    if d.year == selected_date.year and d.month == selected_date.month
)
month_date_cols = [date_columns[d] for d in month_dates]

def get_person_month_stats(pid):
    """นับจำนวนสถานะของคนในเดือนเดียวกับวันที่เลือก"""
    person_id = str(pid).replace(".0", "").strip()
    row = all_id_to_row.get(person_id)

    stats = {
        "01.00": 0,
        "06.00": 0,
        "09.00": 0,
        "พักร้อน (พ)": 0,
        "ขาด (ข)": 0,
        "ป่วย (ป)": 0,
        "H": 0,
    }

    if row is None:
        return stats

    for col in month_date_cols:
        value = str(row[col]).strip()

        if value == "A100":
            stats["01.00"] += 1
        elif value == "A600":
            stats["06.00"] += 1
        elif value == "A900":
            stats["09.00"] += 1
        elif value in {"พ", "พักร้อน"}:
            stats["พักร้อน (พ)"] += 1
        elif value == "ข":
            stats["ขาด (ข)"] += 1
        elif value == "ป":
            stats["ป่วย (ป)"] += 1
        elif value in {"H", "H400"}:
            stats["H"] += 1

    return stats


def render_person_card_button(pid, bg, subtext="", role_badge=""):
    nickname = display_person(pid)
    safe_pid = "".join(ch if ch.isalnum() else "_" for ch in str(pid))
    button_key = f"person_card_{safe_pid}"

    clicked = st.button(
        nickname,
        key=button_key,
        use_container_width=True,
    )

    badge_css = ""
    if role_badge:
        badge_css = f"""
        .st-key-{button_key} button {{
            position: relative !important;
            padding-bottom: 46px !important;
        }}
        .st-key-{button_key} button::after {{
            content: "{role_badge}";
            position: absolute;
            left: 50%;
            bottom: 7px;
            transform: translateX(-50%);
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: auto;
            padding: 3px 7px;
            border-radius: 5px;
            background: #ff3b3b;
            color: #ffffff !important;
            font-size: 11px;
            line-height: 1.05;
            font-weight: 900;
            white-space: nowrap;
            pointer-events: none;
        }}
        """

    subtext_css = ""
    if subtext:
        subtext_css = f"""
        .st-key-{button_key} button::before {{
            content: "{subtext}";
            position: absolute;
            left: 50%;
            bottom: 7px;
            transform: translateX(-50%);
            color: #5b6472 !important;
            font-size: 12px;
            line-height: 1.05;
            font-weight: 900;
            white-space: nowrap;
            pointer-events: none;
        }}
        """
        # H cards need space for the subtext without a red badge.
        if not role_badge:
            subtext_css = f"""
            .st-key-{button_key} button {{
                position: relative !important;
                padding-bottom: 28px !important;
            }}
            .st-key-{button_key} button::after {{
                content: "{subtext}";
                position: absolute;
                left: 50%;
                bottom: 7px;
                transform: translateX(-50%);
                color: #5b6472 !important;
                font-size: 12px;
                line-height: 1.05;
                font-weight: 900;
                white-space: nowrap;
                pointer-events: none;
            }}
            """

    st.markdown(
        f"""
        <style>
        .st-key-{button_key} {{
            width: 100% !important;
        }}
        .st-key-{button_key} button {{
            width: 100% !important;
            min-height: 64px !important;
            border: 0 !important;
            border-radius: 9px !important;
            background: {bg} !important;
            color: #111111 !important;
            font-size: 25px !important;
            font-weight: 900 !important;
            line-height: 1.15 !important;
            white-space: pre-line !important;
            box-shadow: none !important;
            padding: 10px 12px !important;
            margin-bottom: 10px !important;
            transition: all .15s ease !important;
        }}
        .st-key-{button_key} button:hover {{
            filter: brightness(0.98);
            transform: translateY(-1px);
            box-shadow: 0 3px 10px rgba(0,0,0,.12) !important;
        }}
        .st-key-{button_key} button p,
        .st-key-{button_key} button span,
        .st-key-{button_key} button div {{
            color: #111111 !important;
            font-size: 25px !important;
            font-weight: 900 !important;
            line-height: 1.15 !important;
            white-space: pre-line !important;
        }}
        {badge_css}
        {subtext_css}
        </style>
        """,
        unsafe_allow_html=True,
    )

    if clicked:
        show_person_month_stats(pid)

@st.dialog("สถิติการมาทำงาน")
def show_person_month_stats(pid):

    nickname = display_person(pid)
    stats = get_person_month_stats(pid)

    st.markdown(
        f"### {nickname}\n"
        f"เดือน **{selected_date.strftime('%m/%Y')}**"
    )

    stat_cols = st.columns(2)

    rows = [
        ("01.00", stats["01.00"]),
        ("06.00", stats["06.00"]),
        ("09.00", stats["09.00"]),
        ("พักร้อน (พ)", stats["พักร้อน (พ)"]),
        ("ขาด (ข)", stats["ขาด (ข)"]),
        ("ป่วย (ป)", stats["ป่วย (ป)"]),
        ("H", stats["H"]),
    ]

    for i, (label, value) in enumerate(rows):
        with stat_cols[i % 2]:
            st.metric(label, f"{value} วัน")

# ============================================================
# รายชื่อคนที่มาทำงานตามกะ
# ============================================================
if "show_shift_people" not in st.session_state:
    st.session_state["show_shift_people"] = True

toggle_label = (
    "🙈 ซ่อนรายชื่อคนที่มาทำงานตามกะ"
    if st.session_state["show_shift_people"]
    else "👥 แสดงรายชื่อคนที่มาทำงานตามกะ"
)

st.markdown(
    '<div style="height:6px;"></div>',
    unsafe_allow_html=True,
)

if st.button(
    toggle_label,
    key=f"toggle_shift_people_{selected_date.isoformat()}"
):
    st.session_state["show_shift_people"] = not st.session_state["show_shift_people"]
    st.rerun()

if st.session_state["show_shift_people"]:
    st.markdown("""
    <style>
    .shift-panel-title {
        border: 3px solid #ff4b4b;
        border-radius: 14px;
        padding: 14px 16px;
        margin: 10px 0 10px 0;
        background: #ffffff;
        color: #111111 !important;
        font-size: 18px;
        font-weight: 900;
    }


    /* ชื่อทุกคนเป็นปุ่มกดดูสถิติรายเดือน */
    [data-testid="stPopover"] > button {
        width: 100% !important;
        min-height: 64px !important;
        border: 0 !important;
        border-radius: 9px !important;
        background: transparent !important;
        color: #111111 !important;
        font-size: 18px !important;
        font-weight: 900 !important;
        box-shadow: none !important;
        padding: 10px 12px !important;
    }

    [data-testid="stPopover"] > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 3px 10px rgba(0,0,0,.12) !important;
    }

    .person-card-button-label {
        font-size: 18px;
        font-weight: 900;
        color: #111111 !important;
        text-align: center;
    }

    .shift-person-card {
        border-radius: 9px;
        min-height: 64px;
        padding: 10px 12px;
        box-sizing: border-box;
        color: #111111 !important;
        font-size: 18px;
        font-weight: 900;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        text-align: center;
        margin-bottom: 10px;
    }

    .shift-person-card * {
        color: #111111 !important;
    }

    
    .h-previous-shift {
        margin-top: 3px;
        font-size: 12px;
        font-weight: 900;
        color: #5b6472 !important;
    }

.shift-role-badge {
        display: inline-block;
        width: fit-content;
        margin-top: 7px;
        padding: 5px 10px;
        border-radius: 6px;
        background: #ff3b3b;
        color: #ffffff !important;
        font-size: 20px;
        line-height: 1.1;
        font-weight: 900;
        letter-spacing: 0.1px;
    }

    @media (max-width: 1000px) {
        .shift-person-card {
            min-height: 54px;
            font-size: 14px;
        }
    }
    </style>
    """, unsafe_allow_html=True)

    for code in ["A100", "A600", "A900"]:
        info = SHIFT_INFO[code]
        names = working.loc[
            working["shift_code"] == code, "person_id"
        ].tolist()

        # หัวกะเต็มความกว้าง
        st.markdown(
            f'<div class="shift-panel-title">กะ {info["name"]} — {len(names)} คน</div>',
            unsafe_allow_html=True,
        )

        # 6 คนต่อแถว ตามรูปตัวอย่าง
        for row_start in range(0, len(names), 6):
            row_names = names[row_start:row_start + 6]
            cols = st.columns(6, gap="small")

            for col, pid in zip(cols, row_names):
                with col:
                    row = id_to_row.get(pid)
                    nickname = display_person(pid)
                    position = str(row["position"]).strip() if row is not None else ""
                    role_badge = (
                        f'<div class="shift-role-badge">{position}</div>'
                        if position in LEADER_ROLES
                        else ""
                    )

                    render_person_card_button(
                        pid,
                        info["bg"],
                        role_badge=position if position in LEADER_ROLES else "",
                    )

    # ========================================================
    # H วันนี้: ใช้กะของ "เมื่อวาน" เป็นตัวกำหนดสี
    # ========================================================
    h_today_ids = operational_employees.loc[
        operational_employees[date_col].astype(str).str.strip().isin({"H", "H400"}),
        1
    ].astype(str).str.replace(".0", "", regex=False).str.strip().tolist()

    yesterday = selected_date - pd.Timedelta(days=1)
    yesterday_col = date_columns.get(yesterday)

    h_shift_groups = {
        "A100": [],
        "A600": [],
        "A900": [],
        "GRAY": [],
    }

    if yesterday_col is not None:
        yesterday_values = employees[yesterday_col].astype(str).str.strip()

        for pid in h_today_ids:
            rows = employees[
                employees[1].astype(str).str.replace(".0", "", regex=False).str.strip() == pid
            ]

            if rows.empty:
                h_shift_groups["OTHER"].append(pid)
                continue

            previous_shift = str(rows.iloc[0][yesterday_col]).strip()

            if previous_shift in SHIFT_INFO:
                h_shift_groups[previous_shift].append(pid)
            elif previous_shift in {"ป", "ว", "พ"}:
                # สีเทาใช้เฉพาะคนที่เมื่อวานมีค่า ป / ว / พ
                h_shift_groups["GRAY"].append((pid, previous_shift))
    else:
        h_shift_groups["GRAY"] = []

    total_h_people = len(h_today_ids)

    st.markdown(
        f'<div class="shift-panel-title">H — {total_h_people} คน</div>',
        unsafe_allow_html=True,
    )

    # แสดง H ตามสีกะของเมื่อวาน
    for previous_code in ["A100", "A600", "A900", "GRAY"]:
        h_names = h_shift_groups[previous_code]
        if not h_names:
            continue

        if previous_code in SHIFT_INFO:
            info = SHIFT_INFO[previous_code]
            bg = info["bg"]
            previous_label_map = {pid: info["name"] for pid in h_names}
            person_list = h_names
        else:
            bg = "#E5E7EB"
            previous_label_map = {pid: status for pid, status in h_names}
            person_list = [pid for pid, _ in h_names]

        for row_start in range(0, len(person_list), 6):
            row_names = person_list[row_start:row_start + 6]
            cols = st.columns(6, gap="small")

            for col, pid in zip(cols, row_names):
                with col:
                    nickname = display_person(pid)
                    row = all_id_to_row.get(str(pid))

                    position = str(row[4]).strip() if row is not None else ""

                    role_badge = (
                        f'<div class="shift-role-badge">{position}</div>'
                        if position in LEADER_ROLES
                        else ""
                    )

                    render_person_card_button(
                        pid,
                        bg,
                        subtext=previous_label_map[pid],
                        role_badge=position if position in LEADER_ROLES else "",
                    )

# ============================================================
# จัดกำลังคน — XD1-XD20
# ============================================================
st.markdown('<div class="section-title">จัดกำลังคน</div>', unsafe_allow_html=True)
st.markdown("""
<style>
.xd-board-wrap {
    background: transparent !important;
    border-radius: 0;
    padding: 0;
    margin-top: 8px;
}

.xd-cell {
    background: #ffffff;
    border: 2px solid #111111;
    border-radius: 12px;
    padding: 8px;
    margin-bottom: 14px;
}

.xd-cell-title {
    min-height: 36px;
    display: flex;
    align-items: center;
    justify-content: center;
    text-align: center;
    padding: 0 8px;
    border-radius: 7px;
    background: #e7ecef;
    color: #374151;
    font-size: 18px;
    font-weight: 900;
    margin-bottom: 7px;
}

.xd-cell [data-testid="stMultiSelect"] {
    margin-bottom: 0 !important;
}

.xd-cell [data-baseweb="select"] {
    background: #ffffff !important;
}

.xd-cell [data-baseweb="select"] > div {
    min-height: 42px !important;
    border: 2px solid #b7c3cb !important;
    border-radius: 10px !important;
    background: #eaf2f7 !important;
    box-shadow: none !important;
}

/* ข้อความในช่องเลือก */
.xd-cell [data-baseweb="select"] input,
.xd-cell [data-baseweb="select"] span {
    font-size: 16px !important;
    font-weight: 800 !important;
    color: #263238 !important;
}

/* ชื่อที่เลือกแล้ว: ฟ้าเทาอ่อน */
.xd-cell [data-baseweb="tag"],
.xd-cell [data-baseweb="tag"][class] {
    background: #dff2d8 !important;
    background-color: #dff2d8 !important;
    background-image: none !important;
    color: #263238 !important;
    border: 1px solid #b8d9b0 !important;
    border-radius: 7px !important;
    padding: 4px 8px !important;
    font-weight: 900 !important;
    box-shadow: none !important;
}

.xd-cell [data-baseweb="tag"] *,
.xd-cell [data-baseweb="tag"] > div {
    background: transparent !important;
    background-color: transparent !important;
}

.xd-cell [data-baseweb="tag"] span {
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 900 !important;
}

/* ปุ่มลบชื่อ */
.xd-cell [data-baseweb="tag"] svg {
    color: #52666f !important;
}

/* ลูกศร dropdown */
.xd-cell [data-baseweb="select"] svg {
    color: #52666f !important;
}

/* ตัวเลือกใน dropdown */
.xd-cell [role="option"] {
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 700 !important;
    background: #f8fbfd !important;
}

.xd-cell [role="option"][aria-selected="true"] {
    background: #dcecf5 !important;
}

/* dropdown */
.xd-cell [data-baseweb="popover"] {
    border-radius: 10px !important;
}

.xd-cell [role="option"] {
    color: #374151 !important;
    font-size: 14px !important;
}
</style>
""", unsafe_allow_html=True)


st.markdown("""
<style>
/* FINAL OVERRIDE: ชื่อที่เลือกใน XD = เขียวอ่อน */
.xd-cell [data-testid="stMultiSelect"] {
    --primary-color: #dff2d8 !important;
    --secondary-background-color: #dff2d8 !important;
    --background-color: #dff2d8 !important;
}

.xd-cell [data-testid="stMultiSelect"] [data-baseweb="tag"],
.xd-cell [data-testid="stMultiSelect"] [role="listitem"],
.xd-cell [data-testid="stMultiSelect"] [aria-label*="Remove"],
.xd-cell [data-testid="stMultiSelect"] [class*="tag"] {
    background: #dff2d8 !important;
    background-color: #dff2d8 !important;
    background-image: none !important;
    border: 1px solid #b8d9b0 !important;
    border-radius: 7px !important;
    color: #263238 !important;
    box-shadow: none !important;
}

.xd-cell [data-testid="stMultiSelect"] [data-baseweb="tag"] *,
.xd-cell [data-testid="stMultiSelect"] [role="listitem"] * {
    background: transparent !important;
    background-color: transparent !important;
    color: #263238 !important;
}

.xd-cell [data-testid="stMultiSelect"] [data-baseweb="tag"] span,
.xd-cell [data-testid="stMultiSelect"] [role="listitem"] span {
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 900 !important;
}

.xd-cell [data-testid="stMultiSelect"] [data-baseweb="tag"] svg,
.xd-cell [data-testid="stMultiSelect"] [role="listitem"] svg {
    color: #52666f !important;
    fill: #52666f !important;
}
</style>
""", unsafe_allow_html=True)

for i in range(1, 21):
    st.session_state[state_key].setdefault(f"XD{i}", [])

# ตัวเลือกในช่อง XD แสดงเฉพาะคนที่มาทำงานกะ 01.00
xd_options = working.loc[
    working["shift_code"] == "A100", "person_id"
].tolist()

# 5 กล่องต่อแถวจริง โดยใช้ Streamlit columns
st.markdown('<div class="xd-board-wrap">', unsafe_allow_html=True)

for row_start in range(1, 21, 5):
    cols = st.columns(5, gap="small")
    for offset, col in enumerate(cols):
        i = row_start + offset
        with col:
            current = st.session_state[state_key].get(f"XD{i}", [])

            # คนที่ถูกใส่ประตู XD อื่นแล้ว จะไม่แสดงในตัวเลือกของประตูนี้
            assigned_to_other_xds = set()
            for xd_no in range(1, 21):
                if xd_no != i:
                    assigned_to_other_xds.update(
                        st.session_state[state_key].get(f"XD{xd_no}", [])
                    )

            available_xd_options = [
                pid for pid in xd_options
                if pid not in assigned_to_other_xds or pid in current
            ]

            # กันค่าค้างจาก state ที่ไม่อยู่ในตัวเลือกปัจจุบัน
            current = [pid for pid in current if pid in available_xd_options]

            st.markdown(
                f'<div class="xd-cell"><div class="xd-cell-title">XD{i}</div>',
                unsafe_allow_html=True
            )

            selected_ids = st.multiselect(
                f"เลือกชื่อ XD{i}",
                options=available_xd_options,
                default=current,
                format_func=display_person,
                placeholder="Choose options",
                key=f"xd_select_{selected_date.isoformat()}_{widget_version}_{i}",
                label_visibility="collapsed",
            )

            st.session_state[state_key][f"XD{i}"] = selected_ids
            st.markdown("</div>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)

# ============================================================
# จุดปฏิบัติงานพิเศษ
# ============================================================
st.markdown('<div class="section-title">จุดปฏิบัติงานพิเศษ</div>', unsafe_allow_html=True)

# จุดปฏิบัติงานพิเศษทั้ง 5 จุด แสดง "ทุกคน" ที่มาทำงานกะ 01.00
# ไม่ตัดชื่อออกแม้ว่าคนนั้นจะถูกจัดลง XD ไปแล้ว
special_a100_ids = working.loc[
    working["shift_code"] == "A100", "person_id"
].tolist()

options = {
    pid: display_person(pid)
    for pid in special_a100_ids
}

special_cols = st.columns(5)
for col, name in zip(special_cols, SPECIAL_NAMES):
    with col:
        key = f"special_{name}"
        current = st.session_state[state_key].get(key, [])
        selected_ids = st.multiselect(
            name,
            options=list(options.keys()),
            default=[pid for pid in current if pid in options],
            format_func=lambda pid: options.get(pid, display_person(pid)),
            key=f"special_pick_{selected_date.isoformat()}_{widget_version}_{name}",
        )
        st.session_state[state_key][key] = selected_ids

if st.button("🔄 รีเซ็ตการจัดกำลังคนของวันที่เลือก", use_container_width=True):
    # เปลี่ยน widget version เพื่อบังคับให้ Streamlit สร้างช่องเลือกใหม่
    # จึงล้างค่าที่ค้างอยู่ใน XD1–XD20 และจุดปฏิบัติงานพิเศษได้จริง
    st.session_state[state_key] = default_assignments()
    st.session_state[reset_version_key] += 1
    st.rerun()

st.markdown("""
<style>
/* FINAL: สีชื่อที่เลือกใน XD เท่านั้น */
[class*="st-key-xd_select_"] [data-baseweb="tag"],
[class*="st-key-xd_select_"] [data-baseweb="tag"][class],
[class*="st-key-xd_select_"] [role="listitem"],
[class*="st-key-xd_select_"] [class*="tag"] {
    background: #dff2d8 !important;
    background-color: #dff2d8 !important;
    background-image: none !important;
    border: 1px solid #b8d9b0 !important;
    border-radius: 7px !important;
    color: #263238 !important;
    box-shadow: none !important;
}

[class*="st-key-xd_select_"] [data-baseweb="tag"] *,
[class*="st-key-xd_select_"] [role="listitem"] * {
    background: transparent !important;
    background-color: transparent !important;
    color: #263238 !important;
}

[class*="st-key-xd_select_"] [data-baseweb="tag"] span,
[class*="st-key-xd_select_"] [role="listitem"] span {
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 900 !important;
}

[class*="st-key-xd_select_"] [data-baseweb="tag"] svg,
[class*="st-key-xd_select_"] [role="listitem"] svg {
    color: #52666f !important;
    fill: #52666f !important;
}
</style>
""", unsafe_allow_html=True)



st.markdown("""
<style>
/* บังคับสี chip ที่เลือกแล้วใน XD ให้เป็นเขียวอ่อน */
[class*="st-key-xd_select_"] div[data-baseweb="tag"],
[class*="st-key-xd_select_"] span[data-baseweb="tag"],
[class*="st-key-xd_select_"] [data-baseweb="tag"] {
    background: #dff2d8 !important;
    background-color: #dff2d8 !important;
    background-image: none !important;
    color: #263238 !important;
    border: 1px solid #b8d9b0 !important;
    box-shadow: none !important;
}

[class*="st-key-xd_select_"] div[data-baseweb="tag"] > div,
[class*="st-key-xd_select_"] div[data-baseweb="tag"] > span,
[class*="st-key-xd_select_"] [data-baseweb="tag"] * {
    background: transparent !important;
    background-color: transparent !important;
    background-image: none !important;
    color: #263238 !important;
}

[class*="st-key-xd_select_"] [data-baseweb="tag"] span {
    color: #263238 !important;
    font-size: 15px !important;
    font-weight: 900 !important;
}

[class*="st-key-xd_select_"] [data-baseweb="tag"] svg {
    color: #52666f !important;
    fill: #52666f !important;
    background: transparent !important;
}
</style>
""", unsafe_allow_html=True)

