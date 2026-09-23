import base64
from datetime import date, datetime
import io
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
import streamlit as st
import streamlit.components.v1 as components
import libsql_client


# ---------------------------------------------------------
# TURSO CLOUD DATABASE CONNECTION
# ---------------------------------------------------------
def get_db_connection():
    try:
        url = st.secrets["https://kovil-kanakku-kovilproject.aws-ap-northeast-1.turso.io"]
        token = st.secrets[
            "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3ODc4OTY2NTIsImlkIjoiMDFhMDQ2ZjAtODgwMS03YjRkLTk2YjYtYTNmZDMxOTg3MTgyIiwia2lkIjoiT2hSME10YU5BLXp0a3BLNVYxWUV0UGtNSEEyNFQ3c3g3MWplZ3lSUWxpZyIsInJpZCI6IjY3MmE0MTRkLWEyOTQtNGY0MS04NDgxLWJkN2VjNDI0NDNhNSJ9.D5Lwwf5QXIl2xCBhDidOv2IyI_0rISVgCR-rNQktwPW6QovZqcmExe0hwHozb1bsWAxo_gSvNeH4X-s9ASvKDA"]
    except Exception:
        url = "https://kovil-kanakku-kovilproject.aws-ap-northeast-1.turso.io"
        token = "eyJhbGciOiJFZERTQSIsInR5cCI6IkpXVCJ9.eyJhIjoicnciLCJpYXQiOjE3ODc4OTY2NTIsImlkIjoiMDFhMDQ2ZjAtODgwMS03YjRkLTk2YjYtYTNmZDMxOTg3MTgyIiwia2lkIjoiT2hSME10YU5BLXp0a3BLNVYxWUV0UGtNSEEyNFQ3c3g3MWplZ3lSUWxpZyIsInJpZCI6IjY3MmE0MTRkLWEyOTQtNGY0MS04NDgxLWJkN2VjNDI0NDNhNSJ9.D5Lwwf5QXIl2xCBhDidOv2IyI_0rISVgCR-rNQktwPW6QovZqcmExe0hwHozb1bsWAxo_gSvNeH4X-s9ASvKDA"

    return libsql_client.create_client_sync(url=url, auth_token=token)


# ---------------------------------------------------------
# STREAMLIT PAGE CONFIG & SESSION INITIALIZATION
# ---------------------------------------------------------
st.set_page_config(
    page_title="அருள்மிகு பெத்தையா காடேரி அம்பிகை", page_icon="🛕", layout="wide"
)

query_params = st.query_params

if "logged_in" not in st.session_state:
    if query_params.get("logged_in") == "true":
        st.session_state.logged_in = True
        st.session_state.username = query_params.get("user", "admin")
    else:
        st.session_state.logged_in = False
        st.session_state.username = ""

# Mobile Viewport & Thermal Print Specific CSS
st.markdown(
    """
    <style>
    body, .stApp {
        overscroll-behavior-y: contain;
    }

    /* Thermal Print Styling */
    @media print {
        body * {
            visibility: hidden;
        }
        #thermal-receipt, #thermal-receipt * {
            visibility: visible;
        }
        #thermal-receipt {
            position: absolute;
            left: 0;
            top: 0;
            width: 80mm;
            font-family: monospace, sans-serif;
            font-size: 12px;
            padding: 5mm;
            color: black !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------
# BACKGROUND SETUP
# ---------------------------------------------------------
def set_local_background(image_path):
    with open(image_path, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode()

    bg_css = f"""
    <style>
    .stApp {{
        background-image: linear-gradient(rgba(255, 255, 255, 0.88), rgba(255, 255, 255, 0.88)), url("data:image/png;base64,{encoded_string}");
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }}
    </style>
    """
    st.markdown(bg_css, unsafe_allow_html=True)


if os.path.exists("bg.jpg"):
    set_local_background("bg.jpg")


# ---------------------------------------------------------
# DATABASE INITIALIZATION (TURSO CLOUD)
# ---------------------------------------------------------
def init_db():
    conn = get_db_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS donors (
            donor_id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            city TEXT,
            phone TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS receipts (
            receipt_no INTEGER PRIMARY KEY AUTOINCREMENT,
            donor_id TEXT,
            date TEXT,
            year INTEGER,
            category TEXT,
            name TEXT,
            city TEXT,
            amount REAL,
            phone TEXT,
            payment_method TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS expenses (
            expense_id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            category TEXT,
            title TEXT,
            amount REAL,
            remarks TEXT,
            bill_no TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT
        )
    """)

    try:
        conn.execute("ALTER TABLE expenses ADD COLUMN bill_no TEXT")
    except Exception:
        pass

    res = conn.execute("SELECT * FROM users WHERE username = 'admin'")
    if not res.rows:
        conn.execute(
            "INSERT INTO users (username, password) VALUES (?, ?)",
            ("admin", "kovil123"),
        )


try:
    init_db()
except Exception as e:
    pass


def generate_donor_id():
    conn = get_db_connection()
    res = conn.execute("SELECT donor_id FROM donors ORDER BY ROWID DESC LIMIT 1")
    if res.rows and res.rows[0][0] and str(res.rows[0][0]).startswith("DID-"):
        num = int(str(res.rows[0][0]).split("-")[1]) + 1
        return f"DID-{num}"
    else:
        return "DID-1001"


# ---------------------------------------------------------
# HELPER FUNCTIONS (PDF & EXCEL GENERATION)
# ---------------------------------------------------------
def render_thermal_receipt_html(data):
    r_no, r_date, r_cat, r_name, r_city, r_amt, r_phone, r_pay = data[:8]
    r_did = data[8] if len(data) > 8 and data[8] else "-"

    html_code = f"""
    <div id="thermal-receipt" style="width: 280px; font-family: 'Courier New', monospace; border: 1px dashed #000; padding: 10px; margin: auto; background: #fff; color: #000;">
        <div style="text-align: center; font-weight: bold; font-size: 14px;">
            அருள்மிகு பெத்தையா காடேரி அம்பிகை
        </div>
        <div style="text-align: center; font-size: 10px; margin-bottom: 5px;">
            மஞ்சள் நீராட்டு வெள்ளாள சமூக குலதெய்வ மண்டகப்படி
        </div>
        <hr style="border-top: 1px dashed #000;">
        <table style="width: 100%; font-size: 11px; text-align: left;">
            <tr><td><b>ரசீது எண்:</b> {r_no}</td><td style="text-align:right;"><b>தேதி:</b> {r_date}</td></tr>
            <tr><td colspan="2"><b>நிதியாளர் ID:</b> {r_did}</td></tr>
        </table>
        <hr style="border-top: 1px dashed #000;">
        <div style="font-size: 11px; line-height: 1.5;">
            <b>வரவு வகை:</b> {r_cat}<br>
            <b>பெயர்:</b> {r_name}<br>
            <b>ஊர்:</b> {r_city if r_city else '-'}<br>
            <b>கைபேசி:</b> {r_phone if r_phone else '-'}<br>
            <b>செலுத்திய முறை:</b> {r_pay if r_pay else 'Cash'}<br>
        </div>
        <hr style="border-top: 1px dashed #000;">
        <div style="text-align: center; font-size: 15px; font-weight: bold; margin: 5px 0;">
            தொகை: Rs. {r_amt:,.2f}/-
        </div>
        <hr style="border-top: 1px dashed #000;">
        <div style="text-align: center; font-size: 10px; margin-top: 5px;">
            நன்றி! அருள்மிகு பெத்தையா காடேரி அம்பிகை துணை!
        </div>
    </div>
    <br>
    <div style="text-align: center;">
        <button onclick="window.print()" style="background-color: #04AA6D; color: white; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-weight: bold;">
            🖨️ பிரிண்ட் செய்க (Print Now)
        </button>
    </div>
    """
    return html_code


def get_pdf_font():
    local_font = "NotoSansTamil-Regular.ttf"
    if os.path.exists(local_font):
        try:
            pdfmetrics.registerFont(TTFont("TamilFont", local_font))
            return "TamilFont"
        except Exception:
            pass

    system_fonts = [
        "C:\\Windows\\Fonts\\Nirmala.ttf",
        "C:\\Windows\\Fonts\\latha.ttf",
    ]
    for font_path in system_fonts:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont("TamilFont", font_path))
                return "TamilFont"
            except Exception:
                continue
    return "Helvetica"


def generate_receipt_pdf(
        title, name, city, amount, receipt_no, date_str, phone, pay_method, donor_id="-"
):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    font_name = get_pdf_font()

    c.setLineWidth(2)
    c.rect(20, 20, width - 40, height - 40)

    c.setFont(font_name, 20)
    c.drawCentredString(width / 2, height - 80, "அருள்மிகு பெத்தையா காடேரி அம்பிகை")
    c.setFont(font_name, 12)
    c.drawCentredString(width / 2, height - 105, "மஞ்சள் நீராட்டு வெள்ளாள சமூக குலதெய்வ மண்டகப்படி")

    c.setFont(font_name, 12)
    c.drawString(50, height - 165, f"ரசீது எண் :  {receipt_no}")
    c.drawString(width - 200, height - 165, f"தேதி :  {date_str}")
    c.drawString(50, height - 185, f"நிதியாளர் ID :  {donor_id if donor_id else '-'}")

    c.setFont(font_name, 16)
    c.drawCentredString(width / 2, height - 215, f"வரவு வகை: {title}")

    c.rect(40, height - 465, width - 80, 230)
    c.setFont(font_name, 13)
    c.drawString(60, height - 265, f"பெயர் (Name)           :   {name}")
    c.drawString(60, height - 305, f"கைபேசி எண் (Phone)     :   {phone if phone else 'N/A'}")
    c.drawString(60, height - 345, f"ஊர் / பகுதி (City)         :   {city}")
    c.drawString(60, height - 385, f"செலுத்திய முறை (Mode)  :   {pay_method if pay_method else 'Cash'}")
    c.drawString(60, height - 425, f"தொகை (Amount)           :   Rs. {amount:,.2f}/-")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer


def generate_combined_excel_report(df_combined, total_income, total_expense, net_balance):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "வரவு செலவு அறிக்கை"

    title_font = Font(name="Calibri", size=16, bold=True, color="800000")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    header_fill = PatternFill(start_color="4A0E17", end_color="4A0E17", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"), right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"), bottom=Side(style="thin", color="D3D3D3")
    )

    ws.merge_cells(f"A1:{openpyxl.utils.get_column_letter(len(df_combined.columns))}1")
    ws["A1"] = "அருள்மிகு பெத்தையா காடேரி அம்பிகை - ஒருங்கிணைந்த வரவு செலவு அறிக்கை"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center")

    headers = list(df_combined.columns)
    ws.append([])
    ws.append(headers)

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    current_row = 4
    for row in df_combined.itertuples(index=False):
        ws.append(list(row))
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_num)
            cell.border = thin_border
            if col_num in [len(headers) - 1, len(headers)]:
                cell.number_format = "₹#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="center")
        current_row += 1

    offset = len(headers) - 2
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=offset)
    ws.cell(row=current_row, column=1, value="மொத்த வரவு (Total Income)").font = bold_font
    ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="right")
    ws.cell(row=current_row, column=offset + 1, value=float(total_income)).font = bold_font
    ws.cell(row=current_row, column=offset + 1).number_format = "₹#,##0.00"

    current_row += 1
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=offset)
    ws.cell(row=current_row, column=1, value="மொத்த செலவு (Total Expense)").font = bold_font
    ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="right")
    ws.cell(row=current_row, column=offset + 2, value=float(total_expense)).font = bold_font
    ws.cell(row=current_row, column=offset + 2).number_format = "₹#,##0.00"

    current_row += 1
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=offset)
    ws.cell(row=current_row, column=1, value="நிகரக் கையிருப்பு (Net Balance)").font = bold_font
    ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="right")
    ws.cell(row=current_row, column=offset + 1, value=float(net_balance)).font = bold_font
    ws.cell(row=current_row, column=offset + 1).number_format = "₹#,##0.00"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def generate_yearly_matrix_excel(df_matrix, title_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ஆண்டு வாரிய மேட்ரிக்ஸ்"

    title_font = Font(name="Calibri", size=15, bold=True, color="800000")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=11, bold=True)
    header_fill = PatternFill(start_color="4A0E17", end_color="4A0E17", fill_type="solid")
    total_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"), right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"), bottom=Side(style="thin", color="D3D3D3")
    )

    num_cols = len(df_matrix.columns)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    ws["A1"] = f"அருள்மிகு பெத்தையா காடேரி அம்பிகை - {title_name}"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center")

    headers = list(df_matrix.columns)
    ws.append([])
    ws.append(headers)

    for col_num in range(1, num_cols + 1):
        c = ws.cell(row=3, column=col_num)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center")

    curr_row = 4
    for row in df_matrix.itertuples(index=False):
        row_list = list(row)
        ws.append(row_list)
        is_last_row = (curr_row == 3 + len(df_matrix))

        for c_idx in range(1, num_cols + 1):
            cell = ws.cell(row=curr_row, column=c_idx)
            cell.border = thin_border
            if is_last_row:
                cell.font = bold_font
                cell.fill = total_fill

            if c_idx > 3 or (is_last_row and c_idx >= 4):
                try:
                    val = float(cell.value)
                    cell.value = val
                    cell.number_format = "₹#,##0.00"
                    cell.alignment = Alignment(horizontal="right")
                except (ValueError, TypeError):
                    pass
            else:
                cell.alignment = Alignment(horizontal="center")
        curr_row += 1

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------
# AUTHENTICATION
# ---------------------------------------------------------
def login():
    st.title("🛕 அருள்மிகு பெத்தையா காடேரி அம்பிகை")
    login_tab, signup_tab = st.tabs(["🔓 உள்நுழைக (Login)", "📝 புதிய கணக்கு உருவாக்க (Sign Up)"])

    with login_tab:
        st.subheader("கணக்கு மேலாண்மை - உள்நுழைவு")
        with st.form("login_form"):
            username = st.text_input("பயனர் பெயர் (Username)")
            password = st.text_input("கடவுச்சொல் (Password)", type="password")
            submit = st.form_submit_button("🔓 உள்நுழைக (Login)", type="primary")

            if submit:
                if not username or not password:
                    st.warning("⚠️ பயனர் பெயர் மற்றும் கடவுச்சொல்லை உள்ளிடவும்!")
                else:
                    conn = get_db_connection()
                    res = conn.execute(
                        "SELECT * FROM users WHERE username = ? AND password = ?",
                        (username, password),
                    )
                    if res.rows:
                        st.session_state.logged_in = True
                        st.session_state.username = username
                        st.query_params["logged_in"] = "true"
                        st.query_params["user"] = username
                        st.rerun()
                    else:
                        st.error("❌ தவறான பயனர் பெயர் அல்லது கடவுச்சொல்!")

    with signup_tab:
        st.subheader("புதிய பயனர் கணக்கு உருவாக்க")
        with st.form("signup_form", clear_on_submit=True):
            new_username = st.text_input("புதிய பயனர் பெயர் (New Username)")
            new_password = st.text_input("புதிய கடவுச்சொல் (New Password)", type="password")
            confirm_password = st.text_input("கடவுச்சொல்லை உறுதிசெய்க (Confirm Password)", type="password")
            signup_submit = st.form_submit_button("📝 கணக்கு உருவாக்கு (Create Account)", type="primary")

            if signup_submit:
                if not new_username or not new_password:
                    st.warning("⚠️ அனைத்து விவரங்களையும் நிரப்பவும்!")
                elif new_password != confirm_password:
                    st.error("❌ கடவுச்சொற்கள் பொருந்தவில்லை!")
                else:
                    conn = get_db_connection()
                    res = conn.execute("SELECT * FROM users WHERE username = ?", (new_username,))
                    if res.rows:
                        st.error("❌ இந்த பயனர் பெயர் ஏற்கனவே உள்ளது!")
                    else:
                        conn.execute(
                            "INSERT INTO users (username, password) VALUES (?, ?)",
                            (new_username, new_password),
                        )
                        st.success("✅ புதிய கணக்கு உருவாக்கப்பட்டது! Login Tab-ல் உள்நுழையலாம்.")


# ---------------------------------------------------------
# MAIN APPLICATION INTERFACE
# ---------------------------------------------------------
if not st.session_state.logged_in:
    login()
else:
    st.sidebar.title(f"👤 வரவேற்பு: {st.session_state.username}")
    if st.sidebar.button("🚪 வெளியேறு (Logout)"):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.query_params.clear()
        st.rerun()

    st.title("🛕 அருள்மிகு பெத்தையா காடேரி அம்பிகை மஞ்சள் நீராட்டு வெள்ளாள சமூக குலதெய்வ மண்டகப்படி")

    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "📥 புதிய வரவு (Receipt)",
            "📤 புதிய செலவு (Expense)",
            "📊 அறிக்கைகள் (Reports)",
            "✏️ பதிவைத் திருத்த / நீக்க (Edit / Delete Entry)",
        ]
    )

    # TAB 1: RECEIPT
    with tab1:
        st.header("📥 புதிய வரவு பதிவு")

        donor_option = st.radio(
            "நிதியாளர் தேர்வு வகை:",
            ["🆕 புதிய நிதியாளர் (New Donor)", "🔍 ஏற்கனவே உள்ள நிதியாளர் (Existing Donor)"],
            horizontal=True
        )

        selected_donor_id = None
        ex_name, ex_city, ex_phone = "", "", ""

        if donor_option == "🔍 ஏற்கனவே உள்ள நிதியாளர் (Existing Donor)":
            conn = get_db_connection()
            res = conn.execute("SELECT donor_id, name, city, phone FROM donors ORDER BY donor_id ASC")
            existing_donors = res.rows

            if existing_donors:
                donor_dict = {f"{d[0]} - {d[1]} ({d[2] if d[2] else ''})": d for d in existing_donors}
                chosen = st.selectbox("நிதியாளரைத் தேர்ந்தெடுக்கவும் (Unique ID / Name):", list(donor_dict.keys()))
                if chosen:
                    d_data = donor_dict[chosen]
                    selected_donor_id = d_data[0]
                    ex_name, ex_city, ex_phone = d_data[1], d_data[2], d_data[3]
            else:
                st.info("முன்பு பதிவு செய்யப்பட்ட நிதியாளர்கள் எவரும் இல்லை. புதிய நிதியாளராகப் பதிவு செய்யவும்.")
                donor_option = "🆕 புதிய நிதியாளர் (New Donor)"

        with st.form("receipt_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                receipt_date = st.date_input(
                    "தேதி (Date)",
                    value=date.today(),
                    min_value=date(1900, 1, 1),
                    max_value=date(2100, 12, 31),
                )
                category = st.selectbox(
                    "வரவு வகை",
                    [
                        "நிரந்தர வைப்பு நிதியாளர்கள் வரவுகள்",
                        "காணிக்கையாளர்கள் வரவுகள்",
                        "சிறப்பு வைப்பு நிதியாளர்கள் வரவுகள்",
                        "சிறப்பு காணிக்கையாளர்கள் வரவுகள்",
                        "திருவிளக்கு பூஜை வரவுகள்",
                    ],
                )
                phone = st.text_input("கைபேசி எண் (Phone No)",
                                      value=ex_phone if donor_option == "🔍 ஏற்கனவே உள்ள நிதியாளர் (Existing Donor)" else "")
                payment_method = st.selectbox(
                    "பணம் செலுத்திய முறை (Payment Method)",
                    [
                        "Cash (பணம்)",
                        "GPay / PhonePe (UPI)",
                        "Bank Transfer (வங்கி மாற்றம்)",
                        "Money order (மணி ஆர்டர் )",
                    ],
                )

            with col2:
                name = st.text_input("பெயர் (Name)",
                                     value=ex_name if donor_option == "🔍 ஏற்கனவே உள்ள நிதியாளர் (Existing Donor)" else "")
                city = st.text_input("ஊர் / பகுதி (City)",
                                     value=ex_city if donor_option == "🔍 ஏற்கனவே உள்ள நிதியாளர் (Existing Donor)" else "")
                amount = st.number_input(
                    "தொகை (₹)", min_value=0.0, step=100.0, format="%.2f"
                )

            submit_rec = st.form_submit_button("💾 சேமி (Save Receipt)", type="primary")

            if submit_rec:
                if not name or amount <= 0:
                    st.warning("⚠️ தயவுசெய்து பெயர் மற்றும் சரியான தொகையை நிரப்பவும்!")
                else:
                    formatted_date = receipt_date.strftime("%d-%m-%Y")
                    rec_year = receipt_date.year

                    conn = get_db_connection()

                    if donor_option == "🆕 புதிய நிதியாளர் (New Donor)" or not selected_donor_id:
                        final_donor_id = generate_donor_id()
                        conn.execute(
                            "INSERT INTO donors (donor_id, name, city, phone) VALUES (?, ?, ?, ?)",
                            (final_donor_id, name, city, phone)
                        )
                    else:
                        final_donor_id = selected_donor_id
                        conn.execute(
                            "UPDATE donors SET name=?, city=?, phone=? WHERE donor_id=?",
                            (name, city, phone, final_donor_id)
                        )

                    conn.execute(
                        "INSERT INTO receipts (donor_id, date, year, category, name, city,"
                        " amount, phone, payment_method) VALUES (?, ?, ?,"
                        " ?, ?, ?, ?, ?, ?)",
                        (
                            final_donor_id,
                            formatted_date,
                            rec_year,
                            category,
                            name,
                            city,
                            amount,
                            phone,
                            payment_method,
                        ),
                    )

                    last_res = conn.execute("SELECT receipt_no FROM receipts ORDER BY receipt_no DESC LIMIT 1")
                    rec_id = last_res.rows[0][0] if last_res.rows else 1

                    st.session_state["last_receipt_id"] = rec_id
                    st.success(
                        f"✅ வரவு வெற்றிகரமாகச் சேமிக்கப்பட்டது! (ரசீது எண்: {rec_id} | நிதியாளர் ID: {final_donor_id})"
                    )

        st.divider()
        st.subheader("🖨️ ரசீது அச்சிடுதல் (Thermal & A4 PDF)")

        default_r_no = st.session_state.get("last_receipt_id", 1)
        r_no_input = st.number_input(
            "ரசீது எண் உள்ளிடவும்:",
            min_value=1,
            step=1,
            value=int(default_r_no),
        )

        if st.button("🔍 ரசீது தேடு"):
            conn = get_db_connection()
            res = conn.execute(
                "SELECT receipt_no, date, category, name, city, amount,"
                " phone, payment_method, donor_id FROM receipts WHERE receipt_no = ?",
                (r_no_input,),
            )
            st.session_state["active_receipt"] = res.rows[0] if res.rows else None

        if "active_receipt" in st.session_state and st.session_state["active_receipt"]:
            data = st.session_state["active_receipt"]
            col_p1, col_p2 = st.columns(2)

            with col_p1:
                st.markdown("### 📄 A4 Standard PDF")
                pdf_bytes = generate_receipt_pdf(
                    data[2], data[3], data[4], data[5], data[0], data[1], data[6], data[7],
                    donor_id=data[8] if len(data) > 8 else "-"
                )
                st.download_button(
                    label="📥 A4 PDF பதிவிறக்கு",
                    data=pdf_bytes,
                    file_name=f"Receipt_{data[0]}.pdf",
                    mime="application/pdf",
                )

            with col_p2:
                st.markdown("### 🧾 Thermal POS Receipt")
                thermal_html = render_thermal_receipt_html(data)
                components.html(thermal_html, height=320, scrolling=True)

    # TAB 2: EXPENSE
    with tab2:
        st.header("📤 புதிய செலவு பதிவு")
        with st.form("expense_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                expense_date = st.date_input(
                    "செலவு தேதி (Date)",
                    value=date.today(),
                    min_value=date(1900, 1, 1),
                    max_value=date(2100, 12, 31),
                )
                exp_category = st.selectbox(
                    "செலவு வகை",
                    [
                        "மின்சாரக் கட்டணம் (EB Bill)",
                        "பூஜைப் பொருட்கள்",
                        "அன்னதானம்",
                        "வேலை ஆள் கூலி",
                        "பராமரிப்புச் செலவு",
                        "இதர செலவுகள்",
                    ],
                )
                exp_bill_no = st.text_input("பில் / ரசீது எண் (Bill / Voucher No)")
                exp_title = st.text_input("விவரம் (Title)")
            with col2:
                exp_amount = st.number_input(
                    "செலவுத் தொகை (₹)",
                    min_value=0.0,
                    step=50.0,
                    format="%.2f",
                )
                exp_remarks = st.text_input("குறிப்பு (Remarks)")

            submit_exp = st.form_submit_button("💾 செலவைச் சேமி (Save Expense)", type="primary")

            if submit_exp:
                if not exp_title or exp_amount <= 0:
                    st.warning("⚠️ விவரம் மற்றும் சரியான தொகையை நிரப்பவும்!")
                else:
                    formatted_exp_date = expense_date.strftime("%d-%m-%Y")
                    conn = get_db_connection()
                    conn.execute(
                        "INSERT INTO expenses (date, category, title,"
                        " amount, remarks, bill_no) VALUES (?, ?, ?, ?, ?, ?)",
                        (
                            formatted_exp_date,
                            exp_category,
                            exp_title,
                            exp_amount,
                            exp_remarks,
                            exp_bill_no,
                        ),
                    )
                    st.success("✅ செலவு பதிவு வெற்றிகரமாக சேமிக்கப்பட்டது!")

    # TAB 3: REPORTS
    with tab3:
        st.header("📊 கணக்கு அறிக்கைகள்")

        rep_tab1, rep_tab2 = st.tabs(
            [
                "📑 ஒருங்கிணைந்த வரவு-செலவு அறிக்கை (Combined Income & Expense)",
                "👥 வரவு வகை வாரியான ஆண்டு அறிக்கை (Yearly Matrix Reports)"
            ]
        )

        with rep_tab1:
            st.subheader("📑 வரவு மற்றும் செலவு இணைந்த விரிவான அறிக்கை")

            conn = get_db_connection()
            res_r = conn.execute("SELECT DISTINCT year FROM receipts WHERE year IS NOT NULL ORDER BY year DESC")
            years_r = [str(r[0]) for r in res_r.rows if r[0]]

            res_e = conn.execute(
                "SELECT DISTINCT SUBSTR(date, 7, 4) FROM expenses WHERE date IS NOT NULL ORDER BY SUBSTR(date, 7, 4) DESC")
            years_e = [str(r[0]) for r in res_e.rows if r[0]]

            all_years_set = sorted(list(set(years_r + years_e)), reverse=True)
            avail_years = ["அனைத்து ஆண்டுகளும் (All Years)"] + all_years_set

            col_y1, col_y2 = st.columns(2)
            with col_y1:
                selected_year = st.selectbox("📅 ஆண்டு வடிகட்டி (Year Filter):", avail_years, key="comb_year")
            with col_y2:
                type_filter = st.selectbox("வகை வடிகட்டி (Type Filter):",
                                           ["அனைத்தும் (All)", "வரவு (Income)", "செலவு (Expense)"])

            # 1. வரவுகள் (Payment Method சேர்க்கப்பட்டுள்ளது)
            rec_query = "SELECT receipt_no, date, category, name, city, phone, payment_method, amount FROM receipts"
            rec_params = []
            if selected_year != "அனைத்து ஆண்டுகளும் (All Years)":
                rec_query += " WHERE year = ?"
                rec_params.append(int(selected_year))

            res_rec = conn.execute(rec_query, rec_params)
            rec_rows = res_rec.rows

            # 2. செலவுகள்
            exp_query = "SELECT expense_id, date, category, title, amount, bill_no FROM expenses"
            exp_params = []
            if selected_year != "அனைத்து ஆண்டுகளும் (All Years)":
                exp_query += " WHERE SUBSTR(date, 7, 4) = ?"
                exp_params.append(selected_year)

            res_exp = conn.execute(exp_query, exp_params)
            exp_rows = res_exp.rows

            combined_list = []

            if type_filter in ["அனைத்தும் (All)", "வரவு (Income)"]:
                for r in rec_rows:
                    combined_list.append({
                        "ID / பில் எண்": f"REC-{r[0]}",
                        "தேதி": r[1],
                        "பதிவு வகை": "வரவு (Income)",
                        "வரவு/செலவு பிரிவு": r[2],
                        "பெயர் / விவரம்": r[3],
                        "ஊர்": r[4] if r[4] else "-",
                        "கைபேசி எண்": r[5] if r[5] else "-",
                        "செலுத்திய முறை": r[6] if r[6] else "Cash",
                        "வரவுத் தொகை (₹)": float(r[7]),
                        "செலவுத் தொகை (₹)": 0.0
                    })

            if type_filter in ["அனைத்தும் (All)", "செலவு (Expense)"]:
                for e in exp_rows:
                    combined_list.append({
                        "ID / பில் எண்": f"Bill No: {e[5]}" if e[5] else f"EXP-{e[0]}",
                        "தேதி": e[1],
                        "பதிவு வகை": "செலவு (Expense)",
                        "வரவு/செலவு பிரிவு": e[2],
                        "பெயர் / விவரம்": e[3],
                        "ஊர்": "-",
                        "கைபேசி எண்": "-",
                        "செலுத்திய முறை": "-",
                        "வரவுத் தொகை (₹)": 0.0,
                        "செலவுத் தொகை (₹)": float(e[4])
                    })

            if combined_list:
                df_comb = pd.DataFrame(combined_list)

                tot_inc = df_comb["வரவுத் தொகை (₹)"].sum()
                tot_exp = df_comb["செலவுத் தொகை (₹)"].sum()
                net_bal = tot_inc - tot_exp

                m1, m2, m3 = st.columns(3)
                m1.metric("📥 மொத்த வரவு (Total Income)", f"₹{tot_inc:,.2f}")
                m2.metric("📤 மொத்த செலவு (Total Expense)", f"₹{tot_exp:,.2f}")
                m3.metric("💰 நிகர கையிருப்பு (Net Balance)", f"₹{net_bal:,.2f}")

                st.divider()
                st.dataframe(df_comb, use_container_width=True)

                comb_excel = generate_combined_excel_report(df_comb, tot_inc, tot_exp, net_bal)
                st.download_button(
                    label="📊 இந்த அறிக்கையை Excel ஆக பதிவிறக்கு",
                    data=comb_excel,
                    file_name=f"Summary_Income_Expense_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.info("பதிவுகள் எதுவும் கிடைக்கவில்லை.")

        with rep_tab2:
            st.subheader("👥 வரவு வகை வாரியான ஆண்டு அறிக்கை (Yearly Matrix Report)")

            try:
                conn = get_db_connection()
                res = conn.execute("""
                    SELECT 
                        r.year, 
                        r.category, 
                        COALESCE(d.name, r.name) as name, 
                        COALESCE(d.city, r.city) as city, 
                        COALESCE(d.phone, r.phone) as phone,
                        COALESCE(r.payment_method, 'Cash') as payment_method,
                        r.amount
                    FROM receipts r
                    LEFT JOIN donors d ON r.donor_id = d.donor_id
                """)
                df_all_rec = pd.DataFrame(res.rows, columns=["year", "category", "name", "city", "phone", "amount"])
            except Exception as e:
                df_all_rec = pd.DataFrame()

            if isinstance(df_all_rec, pd.DataFrame) and not df_all_rec.empty:
                cat_options = ["அனைத்து வரவு வகைகளும் (All Categories)"] + list(
                    df_all_rec["category"].dropna().unique())
                selected_cat = st.selectbox("வரவு வகையைத் தேர்ந்தெடுக்கவும்:", cat_options)

                df_filtered = df_all_rec if selected_cat == "அனைத்து வரவு வகைகளும் (All Categories)" else df_all_rec[
                    df_all_rec["category"] == selected_cat]

                if not df_filtered.empty:
                    # Index-ல் phone சேர்க்கப்பட்டு donor_id நீக்கப்பட்டுள்ளது
                    pivot_df = df_filtered.pivot_table(
                        index=["name", "city", "phone", "payment_method"],
                        columns="year",
                        values="amount",
                        aggfunc="sum",
                        fill_value=0.0
                    ).reset_index()

                    year_cols = [col for col in pivot_df.columns if
                                 isinstance(col, (int, float, str)) and str(col).isdigit()]
                    year_cols_sorted = sorted(year_cols, key=lambda x: int(x))

                    pivot_df["மொத்தத் தொகை (Total ₹)"] = pivot_df[year_cols_sorted].sum(axis=1)

                    rename_dict = {
                        "name": "பெயர்",
                        "city": "ஊர்",
                        "phone": "கைபேசி எண்"
                        "payment_method": "செலுத்திய முறை"
                    }
                    for y in year_cols_sorted:
                        rename_dict[y] = f"{int(y)} தொகை (₹)"

                    pivot_df.rename(columns=rename_dict, inplace=True)

                    grand_total_row = {
                        "பெயர்": "மொத்தம் (Grand Total)",
                        "ஊர்": "-",
                        "கைபேசி எண்": "-"
                        "செலுத்திய முறை": "-"
                    }

                    for y in year_cols_sorted:
                        grand_total_row[f"{int(y)} தொகை (₹)"] = pivot_df[f"{int(y)} தொகை (₹)"].sum()

                    grand_total_row["மொத்தத் தொகை (Total ₹)"] = pivot_df["மொத்தத் தொகை (Total ₹)"].sum()

                    df_final_display = pd.concat([pivot_df, pd.DataFrame([grand_total_row])], ignore_index=True)

                    st.markdown(f"### 📋 {selected_cat} - மேட்ரிக்ஸ் அட்டவணை")
                    st.dataframe(df_final_display, use_container_width=True)

                    matrix_excel = generate_yearly_matrix_excel(df_final_display, selected_cat)
                    st.download_button(
                        label=f"📊 {selected_cat} - Excel அறிக்கையைப் பதிவிறக்கு",
                        data=matrix_excel,
                        file_name=f"Yearly_Matrix_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                else:
                    st.info("தேர்ந்தெடுக்கப்பட்ட வகைக்குத் தரவு எதுவும் இல்லை.")
            else:
                st.info("வரவுப் பதிவுகள் எதுவும் இதுவரை இல்லை.")

    # TAB 4: EDIT & DELETE ENTRY
    with tab4:
        st.header("✏️ பதிவைத் திருத்துதல் / 🗑️ பதிவை நீக்குதல்")

        edit_type = st.radio(
            "எதை திருத்த/நீக்க வேண்டும்?:",
            ["📥 வரவு (Receipt)", "📤 செலவு (Expense)"],
            horizontal=True,
        )

        edit_id = st.number_input(
            "தேவையான எண் (Receipt No / Expense ID):",
            min_value=1,
            step=1,
        )

        if st.button("🔍 தகவலைக் கொண்டுவா"):
            conn = get_db_connection()
            if "வரவு" in edit_type:
                res = conn.execute(
                    "SELECT receipt_no, date, category, name, city, amount,"
                    " phone, payment_method, donor_id FROM receipts WHERE receipt_no = ?",
                    (edit_id,),
                )
            else:
                res = conn.execute(
                    "SELECT expense_id, date, category, title, amount,"
                    " remarks, bill_no FROM expenses WHERE expense_id = ?",
                    (edit_id,),
                )
            st.session_state["edit_data"] = res.rows[0] if res.rows else None

        if "edit_data" in st.session_state and st.session_state["edit_data"]:
            data = st.session_state["edit_data"]
            st.info(f"எண் {data[0]}-ன் பழைய தகவல்கள் மாற்றத்திற்குத் தயார்.")

            with st.form("edit_form"):
                try:
                    p_date = datetime.strptime(data[1], "%d-%m-%Y").date()
                except Exception:
                    p_date = date.today()

                e_date = st.date_input(
                    "புதிய தேதி:",
                    value=p_date,
                    min_value=date(1900, 1, 1),
                    max_value=date(2100, 12, 31),
                )

                if "வரவு" in edit_type:
                    cat_list = [
                        "நிரந்தர வைப்பு நிதியாளர்கள் வரவுகள்",
                        "காணிக்கையாளர்கள் வரவுகள்",
                        "சிறப்பு வைப்பு நிதியாளர்கள் வரவுகள்",
                        "சிறப்பு காணிக்கையாளர்கள் வரவுகள்",
                        "திருவிளக்கு பூஜை வரவுகள்",
                    ]
                    e_cat = st.selectbox(
                        "வரவு வகை:",
                        cat_list,
                        index=cat_list.index(data[2]) if data[2] in cat_list else 0,
                    )
                    e_name = st.text_input("பெயர்:", value=data[3])
                    e_city = st.text_input("ஊர்:", value=data[4])
                    e_amt = st.number_input(
                        "தொகை (₹):", value=float(data[5]), step=50.0
                    )
                    e_phone = st.text_input("கைபேசி:", value=data[6] or "")

                    pay_methods = [
                        "Cash (பணம்)",
                        "GPay / PhonePe (UPI)",
                        "Bank Transfer (வங்கி மாற்றம்)",
                        "Money order (மணி ஆர்டர் )",
                    ]
                    cur_method = data[7] if len(data) > 7 and data[7] else "Cash (பணம்)"
                    e_method = st.selectbox(
                        "செலுத்திய முறை:",
                        pay_methods,
                        index=pay_methods.index(cur_method) if cur_method in pay_methods else 0,
                    )
                else:
                    cat_list = [
                        "மின்சாரக் கட்டணம் (EB Bill)",
                        "பூஜைப் பொருட்கள்",
                        "அன்னதானம்",
                        "வேலை ஆள் கூலி",
                        "பராமரிப்புச் செலவு",
                        "இதர செலவுகள்",
                    ]
                    e_cat = st.selectbox(
                        "செலவு வகை:",
                        cat_list,
                        index=cat_list.index(data[2]) if data[2] in cat_list else 0,
                    )
                    e_bill_no = st.text_input("பில் / ரசீது எண் (Bill / Voucher No):",
                                              value=data[6] if len(data) > 6 and data[6] else "")
                    e_title = st.text_input("விவரம்:", value=data[3])
                    e_amt = st.number_input(
                        "தொகை (₹):", value=float(data[4]), step=50.0
                    )
                    e_remarks = st.text_input("குறிப்பு:", value=data[5] or "")

                update_btn = st.form_submit_button("🔄 மாற்றங்களை சேமி (Update)", type="primary")

                if update_btn:
                    formatted_u_date = e_date.strftime("%d-%m-%Y")
                    u_year = e_date.year
                    conn = get_db_connection()
                    if "வரவு" in edit_type:
                        conn.execute(
                            "UPDATE receipts SET date=?, year=?, category=?,"
                            " name=?, city=?, amount=?, phone=?,"
                            " payment_method=? WHERE receipt_no=?",
                            (
                                formatted_u_date,
                                u_year,
                                e_cat,
                                e_name,
                                e_city,
                                e_amt,
                                e_phone,
                                e_method,
                                data[0],
                            ),
                        )
                    else:
                        conn.execute(
                            "UPDATE expenses SET date=?, category=?,"
                            " title=?, amount=?, remarks=?, bill_no=? WHERE"
                            " expense_id=?",
                            (
                                formatted_u_date,
                                e_cat,
                                e_title,
                                e_amt,
                                e_remarks,
                                e_bill_no,
                                data[0],
                            ),
                        )
                    st.success("✅ தகவல்கள் வெற்றிகரமாக புதுப்பிக்கப்பட்டன!")
                    st.session_state.pop("edit_data", None)
                    st.rerun()

            st.divider()
            st.subheader("🗑️ பதிவை நீக்கு (Delete Record)")

            confirm_del = st.checkbox("⚠️ இந்த பதிவை நிரந்தரமாக நீக்க விரும்புகிறேன்.")
            if st.button("🗑️ பதிவை நீக்கு (Delete)", type="secondary"):
                if confirm_del:
                    conn = get_db_connection()
                    if "வரவு" in edit_type:
                        conn.execute("DELETE FROM receipts WHERE receipt_no = ?", (data[0],))
                    else:
                        conn.execute("DELETE FROM expenses WHERE expense_id = ?", (data[0],))
                    st.success(f"🗑️ எண் {data[0]} வெற்றிகரமாக நீக்கப்பட்டது!")
                    st.session_state.pop("edit_data", None)
                    st.rerun()
                else:
                    st.warning("⚠️ பதிவை நீக்க மேலே உள்ள உறுதிப்படுத்தல் பெட்டியை (Checkbox) தேர்வு செய்யவும்!")
