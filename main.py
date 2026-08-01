from datetime import datetime
import io
import os
import sqlite3
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
import streamlit as st
import base64


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


# உங்கள் படப் பாதையைக் கொடுக்கவும் (bg.jpg உங்கள் project folder-ல் இருக்க வேண்டும்)
if os.path.exists("bg.jpg"):
    set_local_background("bg.jpg")

# ---------------------------------------------------------
# DATABASE SETUP (SQLite)
# ---------------------------------------------------------
DB_NAME = "kovil_kanakku.db"


def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS receipts (
            receipt_no INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            category TEXT,
            name TEXT,
            city TEXT,
            amount REAL,
            phone TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS expenses (
            expense_id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            category TEXT,
            title TEXT,
            amount REAL,
            remarks TEXT
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT
        )
    """)

    # Check if 'phone' column exists in receipts table
    cursor.execute("PRAGMA table_info(receipts)")
    columns = [column[1] for column in cursor.fetchall()]
    if "phone" not in columns:
        cursor.execute("ALTER TABLE receipts ADD COLUMN phone TEXT")

    # Admin User Verification
    cursor.execute("SELECT * FROM users WHERE username = 'admin'")
    if not cursor.fetchone():
        cursor.execute(
            "INSERT INTO users (username, password) VALUES (?, ?)",
            ("admin", "kovil123"),
        )
    conn.commit()
    conn.close()


init_db()


# ---------------------------------------------------------
# HELPER FUNCTIONS (PDF & EXCEL GENERATION)
# ---------------------------------------------------------
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
    title, name, city, amount, receipt_no, date_str, phone
):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    font_name = get_pdf_font()

    c.setLineWidth(2)
    c.rect(20, 20, width - 40, height - 40)

    c.setFont(font_name, 20)
    c.drawCentredString(
        width / 2, height - 80, "அருள்மிகு பெத்தையா காடேரி அம்பிகை"
    )
    c.setFont(font_name, 12)
    c.drawCentredString(
        width / 2, height - 105, "மஞ்சள் நீராட்டு வெள்ளாள சமூக குலதெய்வ  மண்டகப்படி"
    )

    c.setFont(font_name, 12)
    c.drawString(50, height - 175, f"ரசீது எண் :  {receipt_no}")
    c.drawString(width - 180, height - 175, f"தேதி :  {date_str}")

    c.setFont(font_name, 16)
    c.drawCentredString(width / 2, height - 210, f"வரவு வகை: {title}")

    c.rect(40, height - 440, width - 80, 200)
    c.setFont(font_name, 13)
    c.drawString(60, height - 260, f"பெயர் (Name)     :   {name}")
    c.drawString(
        60, height - 300, f"கைபேசி எண் (Phone)     :   {phone if phone else 'N/A'}"
    )
    c.drawString(60, height - 340, f"ஊர் / பகுதி (City)        :   {city}")
    c.drawString(
        60, height - 400, f"தொகை (Amount)           :   Rs. {amount}/-"
    )

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer


def generate_excel_report(rows, report_type, category, total_amt):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "அறிக்கை"

    title_font = Font(name="Calibri", size=16, bold=True, color="800000")
    sub_font = Font(name="Calibri", size=11, italic=True)
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=11)
    bold_font = Font(name="Calibri", size=11, bold=True)

    header_fill = PatternFill(
        start_color="4A0E17", end_color="4A0E17", fill_type="solid"
    )
    total_fill = PatternFill(
        start_color="EAEAEA", end_color="EAEAEA", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )

    ws.merge_cells("A1:F1")
    ws["A1"] = "அருள்மிகு பெத்தையா காடேரி அம்பிகை"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:F2")
    ws["A2"] = "மஞ்சள் நீராட்டு வெள்ளாள சமூக குலதெய்வ  மண்டகப்படி"
    ws["A2"].font = sub_font
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A3:F3")
    ws["A3"] = (
        f"அறிக்கை வகை: {report_type} | பிரிவு: {category} | தேதி:"
        f" {datetime.now().strftime('%d-%m-%Y')}"
    )
    ws["A3"].font = bold_font
    ws["A3"].alignment = Alignment(horizontal="center")

    ws.append([])
    headers = [
        "எண்",
        "தேதி",
        "வகை",
        "கைபேசி எண்",
        "பெயர் / விவரம்",
        "தொகை (₹)",
    ]
    ws.append(headers)

    header_row = 5
    for col_num in range(1, 7):
        cell = ws.cell(row=header_row, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    current_row = 6
    for row in rows:
        p_val = row[5] if len(row) > 5 and row[5] else "-"
        ws.append([row[0], row[1], row[2], p_val, row[3], row[4]])
        for col_num in range(1, 7):
            cell = ws.cell(row=current_row, column=col_num)
            cell.font = data_font
            cell.border = thin_border
            if col_num == 6:
                cell.number_format = "₹#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            elif col_num in [1, 2, 4]:
                cell.alignment = Alignment(horizontal="center")
        current_row += 1

    ws.merge_cells(
        start_row=current_row,
        start_column=1,
        end_row=current_row,
        end_column=5,
    )
    ws.cell(row=current_row, column=1, value="மொத்தம்").font = bold_font
    ws.cell(row=current_row, column=1).alignment = Alignment(
        horizontal="right"
    )

    tot_cell = ws.cell(row=current_row, column=6, value=total_amt)
    tot_cell.font = bold_font
    tot_cell.number_format = "₹#,##0.00"
    tot_cell.alignment = Alignment(horizontal="right")

    for col_num in range(1, 7):
        c = ws.cell(row=current_row, column=col_num)
        c.fill = total_fill
        c.border = thin_border

    column_widths = {"A": 10, "B": 15, "C": 22, "D": 18, "E": 32, "F": 20}
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------
# STREAMLIT UI INTERFACE
# ---------------------------------------------------------
st.set_page_config(
    page_title="அருள்மிகு பெத்தையா காடேரி அம்பிகை", page_icon="🛕", layout="wide"
)
# PWA / Mobile App Support
pwa_code = """
<script>
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', function() {
      navigator.serviceWorker.register('/sw.js').then(function(registration) {
        console.log('ServiceWorker registration successful');
      });
    });
  }
</script>
"""
st.markdown(pwa_code, unsafe_allow_html=True)

# Authentication Session State
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False


def login():
    st.title("🛕 அருள்மிகு பெத்தையா காடேரி அம்பிகை")
    st.subheader("கணக்கு மேலாண்மை - உள்நுழைவு")

    with st.form("login_form"):
        username = st.text_input("பயனர் பெயர் (Username)")
        password = st.text_input("கடவுச்சொல் (Password)", type="password")
        submit = st.form_submit_button("🔓 உள்நுழைக (Login)")

        if submit:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM users WHERE username = ? AND password = ?",
                (username, password),
            )
            user = cursor.fetchone()
            conn.close()

            if user:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.rerun()
            else:
                st.error("❌ தவறான பயனர் பெயர் அல்லது கடவுச்சொல்!")


if not st.session_state.logged_in:
    login()
else:
    # Sidebar Header & Logout
    st.sidebar.title(f"👤 வரவேற்பு: {st.session_state.username}")
    if st.sidebar.button("🚪 வெளியேறு (Logout)"):
        st.session_state.logged_in = False
        st.rerun()

    st.title("🛕 அருள்மிகு பெத்தையா காடேரி அம்பிகை கணக்கு மேலாண்மை")

    # 3 TABS CREATION HERE
    tab1, tab2, tab3 = st.tabs(
        [
            "📥 புதிய வரவு (Receipt)",
            "📤 புதிய செலவு (Expense)",
            "📊 அறிக்கைகள் (Reports)",
        ]
    )

    # =========================================================
    # TAB 1: RECEIPT (புதிய வரவு)
    # =========================================================
    with tab1:
        st.header("📥 புதிய வரவு பதிவு")

        with st.form("receipt_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                category = st.selectbox(
                    "வரவு வகை",
                    [
                        "சிறப்பு வைப்பு நிதியாளர்கள்",
                        "சிறப்பு காணிக்கையாளர்கள்",
                        "காணிக்கையாளர்கள் வரவுகள்",
                        "காணிக்கையாளர்கள்",
                    ],
                )
                phone = st.text_input("கைபேசி எண் (Phone No)")
                name = st.text_input("பெயர் (Name)")
            with col2:
                city = st.text_input("ஊர் / பகுதி (City)")
                amount = st.number_input(
                    "தொகை (₹)", min_value=0.0, step=100.0, format="%.2f"
                )

            submit_rec = st.form_submit_button(
                "💾 சேமி (Save Receipt)", type="primary"
            )

            if submit_rec:
                if not name or amount <= 0:
                    st.warning(
                        "⚠️ தயவுசெய்து பெயர் மற்றும் சரியான தொகையை நிரப்பவும்!"
                    )
                else:
                    conn = sqlite3.connect(DB_NAME)
                    cursor = conn.cursor()
                    curr_date = datetime.now().strftime("%d-%m-%Y")
                    cursor.execute(
                        "INSERT INTO receipts (date, category, name, city,"
                        " amount, phone) VALUES (?, ?, ?, ?, ?, ?)",
                        (curr_date, category, name, city, amount, phone),
                    )
                    conn.commit()
                    rec_id = cursor.lastrowid
                    conn.close()
                    st.session_state["last_receipt_id"] = rec_id
                    st.success(
                        "✅ வரவு வெற்றிகரமாகச் சேமிக்கப்பட்டது! (ரசீது எண்:"
                        f" {rec_id})"
                    )

        st.divider()
        st.subheader("🖨️ ரசீது பதிவிறக்கம் (PDF Download)")

        default_r_no = st.session_state.get("last_receipt_id", 1)
        r_no_input = st.number_input(
            "ரசீது எண் உள்ளிடவும்:",
            min_value=1,
            step=1,
            value=int(default_r_no),
        )

        if st.button("🔍 ரசீது தேடு"):
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute(
                "SELECT receipt_no, date, category, name, city, amount, phone"
                " FROM receipts WHERE receipt_no = ?",
                (r_no_input,),
            )
            data = cursor.fetchone()
            conn.close()

            if data:
                pdf_bytes = generate_receipt_pdf(
                    data[2],
                    data[3],
                    data[4],
                    data[5],
                    data[0],
                    data[1],
                    data[6],
                )
                st.download_button(
                    label="📥 PDF ரசீது பதிவிறக்கு",
                    data=pdf_bytes,
                    file_name=f"Receipt_{data[0]}.pdf",
                    mime="application/pdf",
                )
            else:
                st.error("❌ இந்த ரசீது எண் காணப்படவில்லை!")

    # =========================================================
    # TAB 2: EXPENSE (புதிய செலவு)
    # =========================================================
    with tab2:
        st.header("📤 புதிய செலவு பதிவு")

        with st.form("expense_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
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
                exp_title = st.text_input("விவரம் (Title)")
            with col2:
                exp_amount = st.number_input(
                    "செலவுத் தொகை (₹)", min_value=0.0, step=50.0, format="%.2f"
                )
                exp_remarks = st.text_input("குறிப்பு (Remarks)")

            submit_exp = st.form_submit_button(
                "💾 செலவைச் சேமி (Save Expense)", type="primary"
            )

            if submit_exp:
                if not exp_title or exp_amount <= 0:
                    st.warning("⚠️ விவரம் மற்றும் சரியான தொகையை நிரப்பவும்!")
                else:
                    conn = sqlite3.connect(DB_NAME)
                    cursor = conn.cursor()
                    curr_date = datetime.now().strftime("%d-%m-%Y")
                    cursor.execute(
                        "INSERT INTO expenses (date, category, title, amount,"
                        " remarks) VALUES (?, ?, ?, ?, ?)",
                        (
                            curr_date,
                            exp_category,
                            exp_title,
                            exp_amount,
                            exp_remarks,
                        ),
                    )
                    conn.commit()
                    conn.close()
                    st.success("✅ செலவு பதிவு வெற்றிகரமாக சேமிக்கப்பட்டது!")

    # =========================================================
    # TAB 3: REPORTS (அறிக்கைகள் & Delete Option)
    # =========================================================
    with tab3:
        st.header("📊 கணக்கு அறிக்கைகள்")

        view_type = st.radio(
            "அறிக்கை வகை:",
            ["📥 வரவு பட்டியல் (Income)", "📤 செலவு பட்டியல் (Expense)"],
            horizontal=True,
        )
        cat_filter = st.selectbox(
            "வகை வடிகட்டி (Filter):",
            [
                "அனைத்தும் (All)",
                "சிறப்பு வைப்பு நிதியாளர்கள்",
                "சிறப்பு காணிக்கையாளர்கள்",
                "காணிக்கையாளர்கள் வரவுகள்",
                "காணிக்கையாளர்கள்",
                "மின்சாரக் கட்டணம் (EB Bill)",
                "பூஜைப் பொருட்கள்",
                "அன்னதானம்",
                "வேலை ஆள் கூலி",
                "பராமரிப்புச் செலவு",
                "இதர செலவுகள்",
            ],
        )

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        if "வரவு" in view_type:
            query = (
                "SELECT receipt_no, date, category, name, amount, phone FROM"
                " receipts"
            )
            if cat_filter != "அனைத்தும் (All)":
                query += f" WHERE category = '{cat_filter}'"
            query += " ORDER BY receipt_no DESC"
        else:
            query = (
                "SELECT expense_id, date, category, title, amount, remarks"
                " FROM expenses"
            )
            if cat_filter != "அனைத்தும் (All)":
                query += f" WHERE category = '{cat_filter}'"
            query += " ORDER BY expense_id DESC"

        cursor.execute(query)
        rows = cursor.fetchall()

        cursor.execute("SELECT SUM(amount) FROM receipts")
        tot_v = cursor.fetchone()[0] or 0.0
        cursor.execute("SELECT SUM(amount) FROM expenses")
        tot_s = cursor.fetchone()[0] or 0.0
        conn.close()

        filtered_tot = sum(r[4] for r in rows) if rows else 0.0

        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.metric(
                label="தேர்ந்தெடுக்கப்பட்ட வகைத் தொகை",
                value=f"₹{filtered_tot:,.2f}",
            )
        with col_m2:
            st.metric(
                label="மொத்தக் கையிருப்பு (Balance)",
                value=f"₹{(tot_v - tot_s):,.2f}",
            )

        st.divider()

        # DELETE SECTION
        with st.expander("🗑️ பதிவை நீக்க (Delete Entry)"):
            st.warning(
                "⚠️ கவனிக்க: நீக்கப்பட்ட பதிவு டேட்டாபேஸிலிருந்து நிரந்தரமாக"
                " அழிக்கப்படும்!"
            )
            delete_id = st.number_input(
                "நீக்க வேண்டிய எண் (ID / Receipt No) உள்ளிடவும்:",
                min_value=1,
                step=1,
            )

            if st.button("❌ பதிவை நீக்கு (Confirm Delete)", type="primary"):
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                if "வரவு" in view_type:
                    cursor.execute(
                        "DELETE FROM receipts WHERE receipt_no = ?",
                        (delete_id,),
                    )
                else:
                    cursor.execute(
                        "DELETE FROM expenses WHERE expense_id = ?",
                        (delete_id,),
                    )

                if cursor.rowcount > 0:
                    conn.commit()
                    conn.close()
                    st.success(
                        f"✅ எண் {delete_id} பதிவு வெற்றிகரமாக நீக்கப்பட்டது!"
                    )
                    st.rerun()
                else:
                    conn.close()
                    st.error(f"❌ எண் {delete_id} காணப்படவில்லை!")

        st.divider()

        # DATA TABLE DISPLAY
        st.dataframe(
            rows,
            column_config={
                "0": "எண்",
                "1": "தேதி",
                "2": "வகை",
                "3": "பெயர்/விவரம்",
                "4": "தொகை (₹)",
                "5": "கைபேசி/குறிப்பு",
            },
            use_container_width=True,
        )

        if rows:
            excel_data = generate_excel_report(
                rows, view_type, cat_filter, filtered_tot
            )
            st.download_button(
                label="📊 Excel அறிக்கையாகப் பதிவிறக்கு",
                data=excel_data,
                file_name=f"Kovil_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
