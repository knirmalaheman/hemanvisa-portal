import streamlit as st
import pandas as pd
import os
import io
import datetime
import base64
import json

EXCEL_FILE = "Heman_Report.xlsx"
ADMIN_FILE = "admin_config.txt"
PERMS_FILE = "worker_permissions.json"
NOTICE_FILE = "admin_notice.txt"
LOGINS_FILE = "worker_logins_status.json"
ATTENDANCE_FILE = "worker_attendance.json"
LOGO_FILES = ["pngTShirtPrintCricket.jpg", "logo.jpg", "logo.png", "logo.jpeg"]

st.set_page_config(page_title="HemanVisa Portal", page_icon="✈️", layout="wide")

# 🌟 Perfect IST Timezone Helper (Cloud Server UTC + 5:30 hours for Indian Standard Time)
def get_ist_now():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)

def get_admin_creds():
    if os.path.exists(ADMIN_FILE):
        with open(ADMIN_FILE, "r") as f:
            lines = f.read().splitlines()
            if len(lines) >= 2:
                return lines[0].strip(), lines[1].strip()
    return "admin", "admin123"

def save_admin_creds(new_user, new_pass):
    with open(ADMIN_FILE, "w") as f:
        f.write(f"{new_user}\n{new_pass}")

def load_permissions():
    if os.path.exists(PERMS_FILE):
        try:
            with open(PERMS_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_permissions(perms_dict):
    with open(PERMS_FILE, "w") as f:
        json.dump(perms_dict, f, indent=4)

def get_admin_notice():
    if os.path.exists(NOTICE_FILE):
        try:
            with open(NOTICE_FILE, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            return ""
    return ""

def save_admin_notice(notice_text):
    with open(NOTICE_FILE, "w", encoding="utf-8") as f:
        f.write(notice_text.strip())

def load_login_tracking():
    if os.path.exists(LOGINS_FILE):
        try:
            with open(LOGINS_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def load_attendance():
    if os.path.exists(ATTENDANCE_FILE):
        try:
            with open(ATTENDANCE_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def save_attendance(att_list):
    with open(ATTENDANCE_FILE, "w") as f:
        json.dump(att_list, f, indent=4)

def record_worker_login(worker_name):
    if not worker_name:
        return
    now_dt = get_ist_now()
    today_str = now_dt.strftime("%Y-%m-%d")
    time_str = now_dt.strftime("%I:%M:%S %p")
    full_timestamp = f"{today_str} {time_str}"
    
    logins = load_login_tracking()
    logins[worker_name.strip()] = {
        "status": "Logged In",
        "last_login": full_timestamp
    }
    with open(LOGINS_FILE, "w") as f:
        json.dump(logins, f, indent=4)
        
    attendance = load_attendance()
    found = False
    for entry in attendance:
        if entry.get("Worker", "").lower() == worker_name.strip().lower() and entry.get("Date") == today_str:
            if not entry.get("Login_Time") or entry.get("Login_Time") == "-":
                entry["Login_Time"] = time_str
            entry["Status"] = "Active"
            found = True
            break
    if not found:
        attendance.append({
            "Worker": worker_name.strip(),
            "Date": today_str,
            "Login_Time": time_str,
            "Logout_Time": "-",
            "Status": "Active"
        })
    save_attendance(attendance)

def record_worker_logout(worker_name):
    if not worker_name:
        return
    now_dt = get_ist_now()
    today_str = now_dt.strftime("%Y-%m-%d")
    time_str = now_dt.strftime("%I:%M:%S %p")
    
    logins = load_login_tracking()
    if worker_name.strip() in logins:
        logins[worker_name.strip()]["status"] = "Not Logged In"
    with open(LOGINS_FILE, "w") as f:
        json.dump(logins, f, indent=4)
        
    attendance = load_attendance()
    for entry in attendance:
        if entry.get("Worker", "").lower() == worker_name.strip().lower() and entry.get("Date") == today_str:
            entry["Logout_Time"] = time_str
            entry["Status"] = "Logged Out"
            break
    save_attendance(attendance)

def get_sheet_names():
    if os.path.exists(EXCEL_FILE):
        try:
            xls = pd.ExcelFile(EXCEL_FILE)
            return xls.sheet_names
        except Exception:
            return ["Sheet1"]
    return ["Sheet1"]

def load_data(sheet_name=0):
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, dtype=str)
            df.columns = df.columns.str.strip()
            df = df.fillna("")
            if "Unnamed: 5" in df.columns:
                df = df.rename(columns={"Unnamed: 5": "Category"})
            
            if "Updated_By" not in df.columns:
                df["Updated_By"] = ""
            if "Updated_At" not in df.columns:
                df["Updated_At"] = ""
            if "Action_Type" not in df.columns:
                df["Action_Type"] = ""
            if "Delete_Reason" not in df.columns:
                df["Delete_Reason"] = ""
                
            return df
        except Exception:
            return None
    return None

def save_data(df_to_save, sheet_name=0):
    if "Category" in df_to_save.columns:
        df_to_save = df_to_save.rename(columns={"Category": ""})
    
    sheets_data = {}
    if os.path.exists(EXCEL_FILE):
        try:
            xls = pd.ExcelFile(EXCEL_FILE)
            for s in xls.sheet_names:
                sheets_data[s] = pd.read_excel(EXCEL_FILE, sheet_name=s, dtype=str)
        except Exception:
            pass
    
    sheets_data[sheet_name] = df_to_save

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        for s, df_val in sheets_data.items():
            df_val.to_excel(writer, sheet_name=str(s), index=False)

def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

def find_login_user_column(df):
    for col in df.columns:
        if col.strip() in ["username", "Case Worker", "DS-160 Case Worker"]:
            return col
    for col in df.columns:
        if col.strip().lower() in ["cases_worker", "user", "case worker"]:
            return col
    return df.columns[0]

def find_display_username_column(df):
    for col in df.columns:
        if col.strip() == "USERNAME":
            return col
    for col in df.columns:
        if col.strip().lower() == "username":
            return col
    return find_login_user_column(df)

def find_password_column(df):
    for col in df.columns:
        if col.strip().lower() in ["password", "pass"]:
            return col
    return df.columns[1] if len(df.columns) > 1 else df.columns[0]


# 🌟 Session State & Query Params Initialization
qp = st.query_params

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["username"] = ""
    st.session_state["role"] = ""

if "attendance_recorded" not in st.session_state:
    st.session_state["attendance_recorded"] = False

if not st.session_state["logged_in"] and qp.get("logged_in") == "true":
    st.session_state["logged_in"] = True
    st.session_state["username"] = qp.get("username", "")
    st.session_state["role"] = qp.get("role", "worker")
    if qp.get("role") == "worker" and not st.session_state["attendance_recorded"]:
        record_worker_login(qp.get("username", ""))
        st.session_state["attendance_recorded"] = True

if "open_dialog" not in st.session_state:
    st.session_state["open_dialog"] = None

if "selected_row_idx" not in st.session_state:
    st.session_state["selected_row_idx"] = None

# Handle Hyperlink Click from URL Query Params securely
if st.session_state["logged_in"] and "edit_row" in qp:
    try:
        r_val = int(qp["edit_row"])
        st.session_state["selected_row_idx"] = r_val
        st.session_state["open_dialog"] = "edit_case_popup"
        st.session_state["is_editing_popup"] = False
        st.session_state["is_delete_request_mode"] = False
        if "edit_row" in st.query_params:
            del st.query_params["edit_row"]
    except:
        if "edit_row" in st.query_params:
            del st.query_params["edit_row"]
elif not st.session_state["logged_in"] and "edit_row" in qp:
    qp.clear()

if "active_sheet" not in st.session_state:
    def get_initial_sheets():
        if os.path.exists(EXCEL_FILE):
            try:
                xls = pd.ExcelFile(EXCEL_FILE)
                return xls.sheet_names
            except Exception:
                return ["Sheet1"]
        return ["Sheet1"]
    sheets_init = get_initial_sheets()
    saved_sheet = qp.get("active_sheet")
    if saved_sheet in sheets_init:
        st.session_state["active_sheet"] = saved_sheet
    else:
        st.session_state["active_sheet"] = sheets_init[0] if sheets_init else "Sheet1"

if st.session_state.get("logged_in") and st.session_state.get("role") == "worker":
    if not st.session_state.get("attendance_recorded", False):
        record_worker_login(st.session_state["username"])
        st.session_state["attendance_recorded"] = True

# 🎨 Multi-Color Executive SaaS Styling & Professional Palette
st.markdown("""
<style>
    [data-testid="stHeader"] {
        display: none !important;
    }

    .sticky-top-banner {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        z-index: 999999;
        background: #ffffff;
        background-image: radial-gradient(#d0e1f9 1px, transparent 1px), linear-gradient(to right, rgba(29,111,165,0.03), rgba(243,156,18,0.03));
        background-size: 20px 20px;
        padding: 8px 28px;
        border-bottom: 3px solid #0f172a;
        box-shadow: 0 4px 20px rgba(15, 23, 42, 0.12);
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        box-sizing: border-box;
    }
    .sticky-top-banner img {
        max-height: 74px;
        width: auto;
    }
    .banner-middle {
        text-align: center;
        flex: 1;
        padding: 0 15px;
    }
    .world-map-title {
        font-size: 20px;
        font-weight: 800;
        color: #0f172a;
        letter-spacing: 1px;
        text-transform: uppercase;
        margin-bottom: 2px;
    }
    .world-map-subtitle {
        font-size: 12px;
        font-weight: 700;
        color: #d97706;
        letter-spacing: 1.5px;
    }
    .banner-info {
        display: flex;
        align-items: center;
        gap: 16px;
        flex-wrap: wrap;
        font-size: 12px;
        font-weight: 600;
        color: #1e293b;
    }
    .banner-info a {
        color: #1e293b;
        text-decoration: none;
    }
    .social-icons {
        display: flex;
        gap: 8px;
        align-items: center;
    }
    .social-icons img {
        height: 22px;
        width: 22px;
        border-radius: 4px;
        max-height: 22px !important;
    }

    .block-container {
        padding-top: 112px !important;
        padding-bottom: 2.5rem;
        max-width: 1400px !important;
    }

    .hero-section {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 55%, #334155 100%);
        color: white;
        padding: 52px 32px;
        border-radius: 14px;
        text-align: center;
        margin-bottom: 26px;
        box-shadow: 0 10px 28px rgba(15, 23, 42, 0.22);
    }
    .hero-title {
        font-size: 36px;
        font-weight: 800;
        margin-bottom: 14px;
        letter-spacing: 0.4px;
    }
    .hero-subtitle {
        font-size: 16px;
        opacity: 0.92;
        margin-bottom: 20px;
        line-height: 1.65;
        max-width: 800px;
        margin-left: auto;
        margin-right: auto;
    }

    .corporate-card {
        background: #ffffff;
        padding: 22px 24px;
        border-radius: 12px;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.06);
        border: 1px solid #e2e8f0;
        border-top: 4px solid #0f172a;
        height: 100%;
        margin-bottom: 14px;
    }
    .corporate-card h3 {
        color: #0f172a;
        font-size: 18px;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .corporate-card p {
        color: #475569;
        font-size: 14px;
        line-height: 1.55;
        margin: 0;
    }

    .marquee-container {
        background: linear-gradient(90deg, #0f172a 0%, #334155 100%);
        color: white;
        padding: 12px 18px;
        border-radius: 10px;
        font-weight: 700;
        font-size: 14px;
        margin-bottom: 20px !important;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.2);
        overflow: hidden;
        white-space: nowrap;
        border-left: 5px solid #d97706;
    }

    .sticky-control-panel {
        position: sticky;
        top: 96px;
        z-index: 99998;
        background-color: #ffffff;
        padding: 8px 0 10px 0;
        border-bottom: 1px solid #e2e8f0;
        margin-bottom: 14px;
    }

    .kpi-card {
        background: linear-gradient(135deg, #0f172a 0%, #334155 100%);
        color: white;
        padding: 18px 14px;
        border-radius: 12px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.18);
        text-align: center;
        border-bottom: 4px solid #d97706;
        margin-bottom: 12px;
        height: 100px;
        width: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        box-sizing: border-box;
    }
    .kpi-card.gold {
        background: linear-gradient(135deg, #92400e 0%, #b45309 100%);
        border-bottom: 4px solid #0f172a;
    }
    .kpi-card.green {
        background: linear-gradient(135deg, #065f46 0%, #047857 100%);
        border-bottom: 4px solid #fde047;
    }
    .kpi-card.purple {
        background: linear-gradient(135deg, #581c87 0%, #6b21a8 100%);
        border-bottom: 4px solid #fde047;
    }
    .kpi-title {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 1px;
        opacity: 0.92;
        margin-bottom: 6px;
        line-height: 1.2;
        font-weight: 700;
    }
    .kpi-value {
        font-size: 26px;
        font-weight: 800;
        margin: 0;
        line-height: 1.1;
    }

    .section-banner {
        background: linear-gradient(90deg, #f1f5f9 0%, #e2e8f0 100%);
        border-left: 5px solid #0f172a;
        padding: 11px 18px;
        border-radius: 8px;
        margin: 8px 0px;
        font-weight: 800;
        color: #0f172a;
        font-size: 16px;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
    }

    .action-group-title {
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #334155;
        margin: 10px 0 6px 2px;
    }

    /* 🎨 Distinct Professional Button Palettes (No More All-Blue) */
    .stButton > button, [data-testid="stDownloadButton"] > button {
        width: 100% !important;
        min-height: 46px !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        padding: 10px 16px !important;
        color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        box-shadow: 0 4px 10px rgba(15, 23, 42, 0.15);
        margin-top: 4px !important;
        margin-bottom: 8px !important;
        transition: all 0.22s ease;
        white-space: nowrap !important;
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%) !important;
    }

    /* 🟢 Teal / Emerald for New Cases & Login Status */
    button[aria-label*="New / Updated"], button[aria-label*="Login Status"] {
        background: linear-gradient(135deg, #0f766e 0%, #14b8a6 100%) !important;
    }
    /* 🔴 Crimson / Red for Delete Requests */
    button[aria-label*="Delete Requests"] {
        background: linear-gradient(135deg, #991b1b 0%, #dc2626 100%) !important;
    }
    /* 🔵 Slate / Navy for Sync & Attendance */
    button[aria-label*="Sync"], button[aria-label*="Attendance"] {
        background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%) !important;
    }
    /* 🟠 Warm Amber for Notices */
    button[aria-label*="Notice"] {
        background: linear-gradient(135deg, #b45309 0%, #d97706 100%) !important;
    }
    /* 🟣 Deep Purple for Users & Permissions */
    button[aria-label*="Users & Permissions"] {
        background: linear-gradient(135deg, #581c87 0%, #7e22ce 100%) !important;
    }
    /* ⚙️ Charcoal / Dark Slate for Passwords & Refresh */
    button[aria-label*="Password"], button[aria-label*="Refresh"] {
        background: linear-gradient(135deg, #334155 0%, #475569 100%) !important;
    }
    /* 🚪 Logout Button Accent */
    button[aria-label*="Logout"] {
        background: linear-gradient(135deg, #991b1b 0%, #b91c1c 100%) !important;
    }

    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 7px 16px rgba(15, 23, 42, 0.28);
        opacity: 0.92;
    }

    .username-hyperlink {
        color: #2563eb !important;
        text-decoration: underline !important;
        font-weight: 700 !important;
        cursor: pointer !important;
    }
    .username-hyperlink:hover {
        color: #0f172a !important;
    }

    div[data-testid="stRadio"] > div {
        flex-direction: row !important;
        flex-wrap: wrap !important;
        gap: 12px !important;
        background: #f8fafc;
        padding: 8px 14px;
        border-radius: 10px;
        border: 1px solid #e2e8f0;
    }

    div[data-testid="stExpander"] {
        width: 100% !important;
        max-width: 950px !important;
    }
    [data-testid="stVerticalBlock"] {
        gap: 0.45rem !important;
    }
    hr {
        margin-top: 0.9rem !important;
        margin-bottom: 0.9rem !important;
    }
</style>
""", unsafe_allow_html=True)

found_logo = next((f for f in LOGO_FILES if os.path.exists(f)), None)
logo_html = ""
if found_logo:
    ext = found_logo.split(".")[-1].lower()
    mime_type = "image/png" if ext == "png" else "image/jpeg"
    img_b64 = get_base64_image(found_logo)
    logo_html = f'<img src="data:{mime_type};base64,{img_b64}" alt="HemanVisa Logo">'
else:
    logo_html = '<h2 style="margin:0; color:#0b3c65;">HemanVisa</h2>'

st.markdown(
    f"""
    <div class="sticky-top-banner">
        <div>{logo_html}</div>
        <div class="banner-middle">
            <div class="world-map-title">🌐 GLOBAL VISA & IMMIGRATION EXPERTS</div>
            <div class="world-map-subtitle">✈️ Your Gateway to Study, Work & Travel Worldwide</div>
        </div>
        <div class="banner-info">
            <div>✉️ <a href="mailto:services@hemanvisa.com">services@hemanvisa.com</a></div>
            <div>📞 <a href="tel:+919533840632">+91 95338 40632</a></div>
            <div class="social-icons">
                <a href="https://facebook.com" target="_blank"><img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="FB"></a>
                <a href="https://instagram.com" target="_blank"><img src="https://cdn-icons-png.flaticon.com/512/733/733558.png" alt="IG"></a>
                <a href="https://youtube.com" target="_blank"><img src="
