import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import sqlite3
from indic_transliteration import sanscript
from indic_transliteration.sanscript import transliterate
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------------------------------------------------------
# 1. PAGE CONFIG
# ---------------------------------------------------------
st.set_page_config(page_title="கோவில் கணக்கு மேலாண்மை", layout="wide", page_icon="🛕")

# ---------------------------------------------------------
# 2. SQLITE DATABASE SETUP
# ---------------------------------------------------------
DB_NAME = "kovil_kanakku.db"


def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    # Receipts (Varavu) table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS receipts (
            receipt_no INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            category TEXT,
            name TEXT,
            city TEXT,
            amount REAL
        )
    ''')
    # Expenses (Selavu) table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS expenses (
            expense_id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            category TEXT,
            title TEXT,
            amount REAL,
            remarks TEXT
        )
    ''')
    # Users table for Login
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT
        )
    ''')
    # Create default admin if not exists
    cursor.execute("SELECT * FROM users WHERE username = 'admin'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO users (username, password) VALUES (?, ?)", ('admin', 'kovil123'))

    conn.commit()
    conn.close()


init_db()

# ---------------------------------------------------------
# 3. AUTHENTICATION (LOGIN SYSTEM)
# ---------------------------------------------------------
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False


def check_login(username, password):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password))
    user = cursor.fetchone()
    conn.close()
    return user is not None


def update_password(username, new_password):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("UPDATE users SET password = ? WHERE username = ?", (new_password, username))
    conn.commit()
    conn.close()


if not st.session_state['logged_in']:
    st.markdown("<h1 style='text-align: center;'>🛕 அருள்மிகு பெத்தையா காடேரி அம்பிகை</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center;'>கணக்கு மேலாண்மை - உள்நுழைவு (Login)</h3>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("🔐 தொடர லாகின் செய்யவும்")
        user_input = st.text_input("பயனர் பெயர் (Username)")
        pass_input = st.text_input("கடவுச்சொல் (Password)", type="password")

        if st.button("🔓 உள்நுழைக (Login)", use_container_width=True):
            if check_login(user_input, pass_input):
                st.session_state['logged_in'] = True
                st.session_state['username'] = user_input
                st.success("✅ வெற்றிகரமாக உள்நுழைந்துவிட்டீர்கள்!")
                st.rerun()
            else:
                st.error("❌ தவறான பயனர் பெயர் அல்லது கடவுச்சொல்!")
    st.stop()

# ---------------------------------------------------------
# 4. SIDEBAR & LOGOUT
# ---------------------------------------------------------
st.sidebar.title(f"👤 வரவேற்கிறோம், {st.session_state['username']}!")
if st.sidebar.button("🚪 வெளியேறு (Logout)"):
    st.session_state['logged_in'] = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🔑 பாஸ்வேர்ட் மாற்ற")
new_pass = st.sidebar.text_input("புதிய கடவுச்சொல்", type="password")
if st.sidebar.button("பாஸ்வேர்ட் மாற்று"):
    if new_pass:
        update_password(st.session_state['username'], new_pass)
        st.sidebar.success("✅ பாஸ்வேர்ட் மாற்றப்பட்டது!")
    else:
        st.sidebar.warning("⚠️ புதிய பாஸ்வேர்ட் டைப் செய்யவும்!")


# ---------------------------------------------------------
# 5. ACCURATE TANGLISH TO TAMIL CONVERTER
# ---------------------------------------------------------
def to_tamil(text):
    if not text:
        return ""

    special_dict = {
        "muthu": "முத்து", "ramesh": "ரமேஷ்", "suresh": "சுரேஷ்", "kumar": "குமார்",
        "raja": "ராஜா", "karthik": "கார்த்திக்", "murugan": "முருகன்", "selvam": "செல்வம்",
        "madurai": "மதுரை", "melur": "மேலூர்", "chennai": "சென்னை", "kovil": "கோவில்",
        "poojai": "பூஜை", "current bill": "மின்சாரக் கட்டணம்", "pal": "பால்", "flower": "பூக்கள்"
    }

    clean_text = text.strip().lower()
    if clean_text in special_dict:
        return special_dict[clean_text]

    try:
        converted = transliterate(text, sanscript.ITRANS, sanscript.TAMIL)
        return converted
    except Exception:
        return text


# ---------------------------------------------------------
# 6. TAMIL FONT HELPER
# ---------------------------------------------------------
def get_pdf_font():
    font_name = 'Helvetica'
    font_paths = [
        "C:\\Windows\\Fonts\\Nirmala.ttf",
        "C:\\Windows\\Fonts\\latha.ttf",
        "Bamini.ttf",
        "NotoSansTamil-Regular.ttf"
    ]
    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('TamilFont', path))
                return 'TamilFont'
            except:
                continue
    return font_name


# ---------------------------------------------------------
# 7. PDF GENERATION
# ---------------------------------------------------------
def generate_receipt_pdf(title, name, city, amount, receipt_no):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    font_name = get_pdf_font()

    c.setLineWidth(2)
    c.rect(20, 20, width - 40, height - 40)

    c.setFont(font_name, 20)
    c.drawCentredString(width / 2, height - 80, "அருள்மிகு பெத்தையா காடேரி அம்பிகை ")
    c.setFont(font_name, 12)
    c.drawCentredString(width / 2, height - 105, "நற்பணி மன்றம் மற்றும் தளபதி தளபதிகள்")
    c.drawCentredString(width / 2, height - 125, "மேலூர் மேல வீதி, மதுரை - 625020.")
    c.line(40, height - 145, width - 40, height - 145)

    current_date = datetime.now().strftime("%d-%m-%Y")
    c.setFont(font_name, 12)
    c.drawString(50, height - 175, f"ரசீது எண் :  {receipt_no}")
    c.drawString(width - 180, height - 175, f"தேதி :  {current_date}")

    c.setFont(font_name, 16)
    c.drawCentredString(width / 2, height - 210, title)

    c.rect(40, height - 420, width - 80, 180)
    c.setFont(font_name, 13)
    c.drawString(60, height - 260, f"பக்தர் பெயர் (Name)     :   {name}")
    c.drawString(60, height - 300, f"ஊர் / பகுதி (City)        :   {city}")
    c.drawString(60, height - 380, f"தொகை (Amount)           :   Rs. {amount}/-")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def generate_full_report_pdf(df, title_text):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(A4))
    width, height = landscape(A4)
    font_name = get_pdf_font()

    c.setFont(font_name, 16)
    c.drawCentredString(width / 2, height - 40, f"அருள்மிகு பெத்தையா காடேரி அம்பிகை - {title_text}")
    c.setFont(font_name, 10)
    c.drawCentredString(width / 2, height - 55, f"உருவாக்கப்பட்ட தேதி: {datetime.now().strftime('%d-%m-%Y %I:%M %p')}")
    c.line(30, height - 65, width - 30, height - 65)

    y = height - 90
    c.setFont(font_name, 10)

    # Table headers based on columns
    cols = df.columns.tolist()
    x_positions = [40, 120, 220, 360, 560, 720]

    for idx, col in enumerate(cols[:6]):
        c.drawString(x_positions[idx], y, str(col))

    c.line(30, y - 5, width - 30, y - 5)
    y -= 25

    total_amt = 0
    c.setFont(font_name, 10)
    for _, row in df.iterrows():
        if y < 50:
            c.showPage()
            y = height - 50
            c.setFont(font_name, 10)

        for idx, col in enumerate(cols[:6]):
            val = str(row[col]) if pd.notnull(row[col]) else ""
            c.drawString(x_positions[idx], y, val)

        if "தொகை (₹)" in df.columns:
            total_amt += row["தொகை (₹)"]
        y -= 20

    c.line(30, y + 10, width - 30, y + 10)
    if "தொகை (₹)" in df.columns:
        c.setFont(font_name, 11)
        c.drawString(550, y - 10, "மொத்தம்:")
        c.drawString(720, y - 10, f"Rs. {total_amt:,.2f}")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------
# 8. STREAMLIT MAIN APP UI
# ---------------------------------------------------------
st.title("🛕 அருள்மிகு பெத்தையா காடேரி அம்பிகை - வரவு செலவு கணக்கு")

tab1, tab2, tab3, tab4 = st.tabs([
    "📥 புதிய வரவு (Receipt)",
    "📤 புதிய செலவு (Expense)",
    "📊 வரவு-செலவு அறிக்கைகள் (Reports)",
    "🗑️ பதிவு நீக்குதல் (Delete)"
])

# --- TAB 1: RECEIPT (VARAVU) ---
with tab1:
    st.subheader("📝 புதிய வரவு பதிவு")
    col1, col2 = st.columns(2)

    with col1:
        category = st.selectbox("வரவு வகை", ["வைப்பு நிதி", "காணிக்கை", "சிறப்பு நன்கொடை", "இதர வரவு"])
        name_eng = st.text_input("பக்தர் பெயர் (English / Tanglish):", placeholder="எ.கா: muthu / ramesh",
                                 key="v_name_eng")
        name_tam_auto = to_tamil(name_eng) if name_eng else ""
        name_final = st.text_input("தமிழ் பெயர்:", value=name_tam_auto, key="v_name_tam")

    with col2:
        city_eng = st.text_input("ஊர் / பகுதி (English / Tanglish):", placeholder="எ.கா: madurai / melur",
                                 key="v_city_eng")
        city_tam_auto = to_tamil(city_eng) if city_eng else ""
        city_final = st.text_input("தமிழ் ஊர்:", value=city_tam_auto, key="v_city_tam")
        amount = st.number_input("தொகை (₹)", min_value=1.0, step=10.0, key="v_amount")

    st.markdown("---")

    if st.button("💾 வரவைச் சேமி & PDF உருவாக்கு", key="btn_save_varavu"):
        if not name_final or amount <= 0:
            st.warning("⚠️ தயவுசெய்து பெயர் மற்றும் தொகையை நிரப்பவும்!")
        else:
            try:
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                current_date = datetime.now().strftime("%d-%m-%Y")
                cursor.execute('''
                    INSERT INTO receipts (date, category, name, city, amount)
                    VALUES (?, ?, ?, ?, ?)
                ''', (current_date, category, name_final, city_final, amount))
                conn.commit()
                receipt_no = cursor.lastrowid
                conn.close()

                st.success(f"✅ வரவு சேமிக்கப்பட்டது! ரசீது எண்: {receipt_no}")
                pdf_bytes = generate_receipt_pdf(category, name_final, city_final, amount, receipt_no)
                st.download_button(
                    label="📄 தனி நபர் PDF ரசீது டவுன்லோட் செய்",
                    data=pdf_bytes,
                    file_name=f"Receipt_{receipt_no}.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"❌ சேமிப்பதில் எர்ரர்: {e}")

# --- TAB 2: EXPENSE (SELAVU) ---
with tab2:
    st.subheader("💸 புதிய செலவு பதிவு")
    col_e1, col_e2 = st.columns(2)

    with col_e1:
        exp_category = st.selectbox("செலவு வகை", [
            "மின்சாரக் கட்டணம் (EB Bill)",
            "பூஜைப் பொருட்கள்",
            "அன்னதானம்",
            "வேலை ஆள் கூலி",
            "பராமரிப்புச் செலவு",
            "திருவிழாச் செலவு",
            "இதர செலவுகள்"
        ])
        title_eng = st.text_input("செலவு விவரம் (English / Tanglish):", placeholder="எ.கா: EB Bill / Poojai Porutkal",
                                  key="e_title_eng")
        title_tam_auto = to_tamil(title_eng) if title_eng else ""
        title_final = st.text_input("தமிழ் விவரம்:", value=title_tam_auto, key="e_title_tam")

    with col_e2:
        exp_amount = st.number_input("செலவுத் தொகை (₹)", min_value=1.0, step=10.0, key="e_amount")
        remarks = st.text_area("கூடுதல் குறிப்பு (Optional):", placeholder="எ.கா: பில் எண் / யாருக்கு கொடுக்கப்பட்டது",
                               key="e_remarks")

    st.markdown("---")

    if st.button("💾 செலவைச் சேமி", key="btn_save_selavu"):
        if not title_final or exp_amount <= 0:
            st.warning("⚠️ தயவுசெய்து செலவு விவரம் மற்றும் தொகையை நிரப்பவும்!")
        else:
            try:
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                current_date = datetime.now().strftime("%d-%m-%Y")
                cursor.execute('''
                    INSERT INTO expenses (date, category, title, amount, remarks)
                    VALUES (?, ?, ?, ?, ?)
                ''', (current_date, exp_category, title_final, exp_amount, remarks))
                conn.commit()
                exp_id = cursor.lastrowid
                conn.close()

                st.success(f"✅ செலவு பதிவு செய்யப்பட்டது! பதிவு எண்: {exp_id}")
            except Exception as e:
                st.error(f"❌ சேமிப்பதில் எர்ரர்: {e}")

# --- TAB 3: REPORTS (VARAVU & SELAVU) ---
with tab3:
    st.subheader("📊 வரவு - செலவு கணக்கு அறிக்கைகள்")

    report_type = st.radio("பார்க்க வேண்டிய அறிக்கை:",
                           ["💰 மொத்த கணக்கு சுருக்கம் (Overview)", "📥 வரவு பட்டியல் (Income)",
                            "📤 செலவு பட்டியல் (Expenses)"], horizontal=True)

    conn = sqlite3.connect(DB_NAME)
    df_v = pd.read_sql_query(
        "SELECT receipt_no as 'எண்', date as 'தேதி', category as 'வகை', name as 'விவரம்/பெயர்', city as 'ஊர்', amount as 'தொகை (₹)' FROM receipts",
        conn)
    df_s = pd.read_sql_query(
        "SELECT expense_id as 'எண்', date as 'தேதி', category as 'வகை', title as 'விவரம்/பெயர்', remarks as 'குறிப்பு', amount as 'தொகை (₹)' FROM expenses",
        conn)
    conn.close()

    tot_v = df_v['தொகை (₹)'].sum() if not df_v.empty else 0.0
    tot_s = df_s['தொகை (₹)'].sum() if not df_s.empty else 0.0
    net_balance = tot_v - tot_s

    # --- OVERVIEW ---
    if report_type == "💰 மொத்த கணக்கு சுருக்கம் (Overview)":
        st.markdown("### 🏛️ கோவில் நிதி நிலை சுருக்கம்")
        m1, m2, m3 = st.columns(3)
        m1.metric("📥 மொத்த வரவு (Total Income)", f"₹ {tot_v:,.2f}")
        m2.metric("📤 மொத்த செலவு (Total Expenses)", f"₹ {tot_s:,.2f}")

        # Color coding for balance
        if net_balance >= 0:
            m3.metric("💵 கையிருப்பு (Net Balance)", f"₹ {net_balance:,.2f}",
                      delta=f"₹ {net_balance:,.2f} (பரிசீலனையில்)")
        else:
            m3.metric("⚠️ கையிருப்பு (Deficit)", f"₹ {net_balance:,.2f}", delta=f"₹ {net_balance:,.2f} (பற்றாக்குறை)")

        st.markdown("---")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.markdown("#### 📥 சமீபத்திய வரவுகள்")
            st.dataframe(df_v.tail(5), use_container_width=True)
        with col_c2:
            st.markdown("#### 📤 சமீபத்திய செலவுகள்")
            st.dataframe(df_s.tail(5), use_container_width=True)

    # --- VARAVU REPORT ---
    elif report_type == "📥 வரவு பட்டியல் (Income)":
        st.markdown("### 📥 வரவு பதிவேடு")
        if not df_v.empty:
            st.dataframe(df_v.sort_values(by='எண்', ascending=False), use_container_width=True)
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                pdf_v_bytes = generate_full_report_pdf(df_v, "வரவு அறிக்கை")
                st.download_button("🖨️ வரவு PDF Report", data=pdf_v_bytes,
                                   file_name=f"Varavu_Report_{datetime.now().strftime('%d_%m_%Y')}.pdf",
                                   mime="application/pdf")
            with col_p2:
                buffer_v_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_v_excel, engine='openpyxl') as writer:
                    df_v.to_excel(writer, index=False, sheet_name='Varavu')
                st.download_button("📥 வரவு Excel Report", data=buffer_v_excel.getvalue(),
                                   file_name=f"Varavu_{datetime.now().strftime('%d_%m_%Y')}.xlsx")
        else:
            st.info("வரவு பதிவுகள் எதுவுமில்லை.")

    # --- SELAVU REPORT ---
    elif report_type == "📤 செலவு பட்டியல் (Expenses)":
        st.markdown("### 📤 செலவு பதிவேடு")
        if not df_s.empty:
            st.dataframe(df_s.sort_values(by='எண்', ascending=False), use_container_width=True)
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                pdf_s_bytes = generate_full_report_pdf(df_s, "செலவு அறிக்கை")
                st.download_button("🖨️ செலவு PDF Report", data=pdf_s_bytes,
                                   file_name=f"Selavu_Report_{datetime.now().strftime('%d_%m_%Y')}.pdf",
                                   mime="application/pdf")
            with col_s2:
                buffer_s_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_s_excel, engine='openpyxl') as writer:
                    df_s.to_excel(writer, index=False, sheet_name='Selavu')
                st.download_button("📥 செலவு Excel Report", data=buffer_s_excel.getvalue(),
                                   file_name=f"Selavu_{datetime.now().strftime('%d_%m_%Y')}.xlsx")
        else:
            st.info("செலவு பதிவுகள் எதுவுமில்லை.")

# --- TAB 4: DELETE ---
with tab4:
    st.subheader("🗑️ தவறான பதிவை நீக்குதல்")
    del_type = st.radio("எந்த பதிவை நீக்க வேண்டும்?", ["வரவு (Receipt)", "செலவு (Expense)"], horizontal=True)

    conn = sqlite3.connect(DB_NAME)
    if del_type == "வரவு (Receipt)":
        df_del = pd.read_sql_query("SELECT receipt_no as id, name, amount FROM receipts ORDER BY receipt_no DESC", conn)
    else:
        df_del = pd.read_sql_query(
            "SELECT expense_id as id, title as name, amount FROM expenses ORDER BY expense_id DESC", conn)
    conn.close()

    if not df_del.empty:
        item_list = [f"பதிவு எண்: {row['id']} - {row['name']} (₹{row['amount']})" for _, row in df_del.iterrows()]
        selected_item = st.selectbox("நீக்க வேண்டிய பதிவைத் தேர்ந்தெடுக்கவும்:", item_list)
        target_id = int(selected_item.split("-")[0].replace("பதிவு எண்:", "").strip())

        if st.button("❌ இந்த பதிவை நீக்கு"):
            try:
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                if del_type == "வரவு (Receipt)":
                    cursor.execute("DELETE FROM receipts WHERE receipt_no = ?", (target_id,))
                else:
                    cursor.execute("DELETE FROM expenses WHERE expense_id = ?", (target_id,))
                conn.commit()
                conn.close()
                st.success(f"✅ பதிவு எண் {target_id} நீக்கப்பட்டது!")
                st.rerun()
            except Exception as e:
                st.error(f"நீக்குவதில் எர்ரர்: {e}")
    else:
        st.info("நீக்குவதற்கு பதிவுகள் இல்லை.")