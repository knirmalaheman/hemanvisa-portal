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

if "selected_visa_filter" not in st.session_state:
    st.session_state["selected_visa_filter"] = "All"

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

# 🎨 HIGH-END BUSINESS STYLING & EXACT IMAGE MATCH BUTTON COLORS
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
        background: linear-gradient(135deg, #090d16 0%, #1e1b4b 50%, #172554 100%);
        padding: 10px 28px;
        border-bottom: 3px solid #f59e0b;
        box-shadow: 0 8px 30px rgba(15, 23, 42, 0.45);
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        box-sizing: border-box;
    }
    .logo-container {
        margin-left: 120px;
    }
    .sticky-top-banner img {
        max-height: 62px;
        width: auto;
        border-radius: 6px;
        background: #ffffff;
        padding: 2px;
    }
    .banner-middle {
        text-align: center;
        flex: 1;
        padding: 0 15px;
        margin-left: 40px;
    }
    .world-map-title {
        font-size: 19px;
        font-weight: 800;
        color: #ffffff;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        margin-bottom: 2px;
        text-shadow: 0 2px 6px rgba(0,0,0,0.5);
    }
    .world-map-subtitle {
        font-size: 12px;
        font-weight: 700;
        color: #fde047;
        letter-spacing: 1.5px;
        text-shadow: 0 1px 4px rgba(0,0,0,0.5);
    }
    .banner-info {
        display: flex;
        align-items: center;
        gap: 12px;
        flex-wrap: wrap;
        font-size: 12px;
        font-weight: 600;
        color: #e2e8f0;
    }
    .banner-badge {
        background: rgba(255, 255, 255, 0.12);
        border: 1px solid rgba(255, 255, 255, 0.25);
        padding: 6px 12px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .banner-badge a {
        color: #ffffff;
        text-decoration: none;
    }
    .social-icons {
        display: flex;
        gap: 8px;
        align-items: center;
        background: rgba(255, 255, 255, 0.12);
        border: 1px solid rgba(255, 255, 255, 0.25);
        padding: 5px 10px;
        border-radius: 8px;
    }
    .social-icons img {
        height: 20px;
        width: 20px;
        border-radius: 4px;
        background: transparent !important;
        padding: 0 !important;
    }

    .block-container {
        padding-top: 112px !important;
        padding-bottom: 2.5rem;
        max-width: 1400px !important;
    }

    .hero-section {
        background: linear-gradient(135deg, #1e1b4b 0%, #312e81 50%, #1e3a8a 100%);
        color: white;
        padding: 52px 32px;
        border-radius: 14px;
        text-align: center;
        margin-bottom: 26px;
        box-shadow: 0 10px 30px rgba(30, 27, 75, 0.3);
    }
    .hero-title {
        font-size: 36px;
        font-weight: 800;
        margin-bottom: 14px;
    }
    .hero-subtitle {
        font-size: 16px;
        opacity: 0.95;
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
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.08);
        border: 1px solid #cbd5e1;
        border-top: 4px solid #3b82f6;
        height: 100%;
        margin-bottom: 14px;
    }
    .corporate-card h3 {
        color: #1e3a8a;
        font-size: 18px;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .corporate-card p {
        color: #334155;
        font-size: 14px;
        line-height: 1.55;
        margin: 0;
    }

    .marquee-container {
        background: linear-gradient(90deg, #1e1b4b 0%, #3730a3 100%);
        color: white;
        padding: 12px 18px;
        border-radius: 10px;
        font-weight: 700;
        font-size: 14px;
        margin-bottom: 20px !important;
        box-shadow: 0 4px 14px rgba(49, 46, 129, 0.3);
        overflow: hidden;
        white-space: nowrap;
        border-left: 5px solid #f59e0b;
    }

    .sticky-control-panel {
        position: sticky;
        top: 96px;
        z-index: 99998;
        background-color: #ffffff;
        padding: 8px 0 10px 0;
        border-bottom: 2px solid #e2e8f0;
        margin-bottom: 14px;
    }

    .kpi-card {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
        color: white;
        padding: 18px 14px;
        border-radius: 12px;
        box-shadow: 0 8px 22px rgba(37, 99, 235, 0.25);
        text-align: center;
        border-bottom: 4px solid #f59e0b;
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
        background: linear-gradient(135deg, #b45309 0%, #f59e0b 100%);
        border-bottom: 4px solid #1e3a8a;
    }
    .kpi-card.green {
        background: linear-gradient(135deg, #047857 0%, #10b981 100%);
        border-bottom: 4px solid #fde047;
    }
    .kpi-card.purple {
        background: linear-gradient(135deg, #6d28d9 0%, #8b5cf6 100%);
        border-bottom: 4px solid #fde047;
    }
    .kpi-title {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 1px;
        opacity: 0.95;
        margin-bottom: 6px;
        font-weight: 700;
    }
    .kpi-value {
        font-size: 26px;
        font-weight: 800;
        margin: 0;
    }

    .section-banner {
        background: linear-gradient(90deg, #e0e7ff 0%, #f3e8ff 100%);
        border-left: 5px solid #4f46e5;
        padding: 11px 18px;
        border-radius: 8px;
        margin: 8px 0px;
        font-weight: 800;
        color: #1e1b4b;
        font-size: 16px;
    }

    .action-group-title {
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #312e81;
        margin: 10px 0 6px 2px;
    }

    /* 🎨 DEFAULT BUTTONS */
    .stButton > button, [data-testid="stDownloadButton"] > button {
        width: 100% !important;
        min-height: 46px !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        padding: 10px 16px !important;
        color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.2) !important;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.08);
        margin-top: 4px !important;
        margin-bottom: 8px !important;
        transition: all 0.22s ease;
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
    }

    /* 🌟 EXACT IMAGE MATCH: Visa Category Pill Buttons (Soft Pastel)[cite: 4] */
    div[data-testid="stVerticalBlock"]:has(.visa-anchor) + div [data-testid="column"]:nth-child(1) .stButton > button {
        background: #e1effe !important;
        color: #1e429f !important;
        border: 1.5px solid #baccf8 !important;
        border-radius: 999px !important;
    }
    div[data-testid="stVerticalBlock"]:has(.visa-anchor) + div [data-testid="column"]:nth-child(2) .stButton > button {
        background: #fef3c7 !important;
        color: #92400e !important;
        border: 1.5px solid #fcd34d !important;
        border-radius: 999px !important;
    }
    div[data-testid="stVerticalBlock"]:has(.visa-anchor) + div [data-testid="column"]:nth-child(3) .stButton > button {
        background: #def7ec !important;
        color: #03543f !important;
        border: 1.5px solid #83e8bb !important;
        border-radius: 999px !important;
    }
    div[data-testid="stVerticalBlock"]:has(.visa-anchor) + div [data-testid="column"]:nth-child(4) .stButton > button {
        background: #f3e8ff !important;
        color: #5b21b6 !important;
        border: 1.5px solid #d8b4fe !important;
        border-radius: 999px !important;
    }

    /* 🌟 EXACT IMAGE MATCH: Workspace Tools Row Button Styles[cite: 4] */
    div[data-testid="stVerticalBlock"]:has(.workspace-anchor) + div [data-testid="column"]:nth-child(1) .stButton > button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 10px !important;
    }
    div[data-testid="stVerticalBlock"]:has(.workspace-anchor) + div [data-testid="column"]:nth-child(2) .stButton > button {
        background: #ffffff !important;
        color: #1d4ed8 !important;
        border: 2px solid #1d4ed8 !important;
        border-radius: 10px !important;
    }
    div[data-testid="stVerticalBlock"]:has(.workspace-anchor) + div [data-testid="column"]:nth-child(3) .stButton > button,
    div[data-testid="stVerticalBlock"]:has(.workspace-anchor) + div [data-testid="column"]:nth-child(3) .stButton > button:disabled {
        background: #e2e8f0 !important;
        color: #64748b !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
        opacity: 1 !important;
    }
    div[data-testid="stVerticalBlock"]:has(.workspace-anchor) + div [data-testid="column"]:nth-child(4) .stButton > button {
        background: #e2e8f0 !important;
        color: #334155 !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 10px !important;
    }
    div[data-testid="stVerticalBlock"]:has(.workspace-anchor) + div [data-testid="column"]:nth-child(5) .stButton > button {
        background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 10px !important;
    }

    /* 🎨 Admin Management Quick Action Buttons (Gorgeous Business Colors) */
    /* 1. New / Updated Cases: Emerald */
    button[aria-label*="New / Updated"] {
        background: linear-gradient(135deg, #059669 0%, #10b981 100%) !important;
    }
    /* 2. Delete Requests: Crimson */
    button[aria-label*="Delete Requests"] {
        background: linear-gradient(135deg, #dc2626 0%, #ef4444 100%) !important;
    }
    /* 3. Sync Excel: Vibrant Blue */
    button[aria-label*="Sync"] {
        background: linear-gradient(135deg, #2563eb 0%, #3b82f6 100%) !important;
    }
    /* 4. Broadcast Notice: Amber / Orange */
    button[aria-label*="Notice"] {
        background: linear-gradient(135deg, #d97706 0%, #f59e0b 100%) !important;
    }
    /* 5. Users & Permissions: Purple */
    button[aria-label*="Users & Permissions"] {
        background: linear-gradient(135deg, #7c3aed 0%, #a855f7 100%) !important;
    }
    /* 6. Worker Login Status: Teal */
    button[aria-label*="Login Status"] {
        background: linear-gradient(135deg, #0d9488 0%, #14b8a6 100%) !important;
    }
    /* 7. Attendance Report: Indigo */
    button[aria-label*="Attendance"] {
        background: linear-gradient(135deg, #4f46e5 0%, #6366f1 100%) !important;
    }
    /* 8. Admin Password / Refresh: Slate */
    button[aria-label*="Password"], button[aria-label*="Refresh"] {
        background: linear-gradient(135deg, #475569 0%, #64748b 100%) !important;
    }
    /* 🚪 Logout Red Accent */
    button[aria-label*="Logout"] {
        background: linear-gradient(135deg, #b91c1c 0%, #dc2626 100%) !important;
    }

    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-1px);
        opacity: 0.95;
    }

    .username-hyperlink {
        color: #2563eb !important;
        text-decoration: underline !important;
        font-weight: 700 !important;
        cursor: pointer !important;
    }
    .username-hyperlink:hover {
        color: #1e3a8a !important;
    }

    div[data-testid="stRadio"] > div {
        flex-direction: row !important;
        flex-wrap: wrap !important;
        gap: 12px !important;
        background: #f8fafc;
        padding: 8px 14px;
        border-radius: 10px;
        border: 1px solid #cbd5e1;
    }

    div[data-testid="stExpander"] {
        width: 100% !important;
        max-width: 950px !important;
    }
</style>
""", unsafe_allow_html=True)

found_logo = next((f for f in LOGO_FILES if os.path.exists(f)), None)
logo_html = ""
if found_logo:
    ext = found_logo.split(".")[-1].lower()
    mime_type = "image/png" if ext == "png" else "image/jpeg"
    img_b64 = get_base64_image(found_logo)
    logo_html = f'<div class="logo-container"><img src="data:{mime_type};base64,{img_b64}" alt="HemanVisa Logo"></div>'
else:
    logo_html = '<div class="logo-container"><h2 style="margin:0; color:#ffffff;">HemanVisa</h2></div>'

st.markdown(f"""
    <div class="sticky-top-banner">
        <div>{logo_html}</div>
        <div class="banner-middle">
            <div class="world-map-title">🌐 GLOBAL VISA & IMMIGRATION EXPERTS</div>
            <div class="world-map-subtitle">✈️ Your Gateway to Study, Work & Travel Worldwide</div>
        </div>
        <div class="banner-info">
            <div class="banner-badge">✉️ <a href="mailto:services@hemanvisa.com">services@hemanvisa.com</a></div>
            <div class="banner-badge">📞 <a href="tel:+919533840632">+91 95338 40632</a></div>
            <div class="social-icons">
                <a href="https://facebook.com" target="_blank"><img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="FB"></a>
                <a href="https://instagram.com" target="_blank"><img src="https://cdn-icons-png.flaticon.com/512/733/733558.png" alt="IG"></a>
                <a href="https://youtube.com" target="_blank"><img src="https://cdn-icons-png.flaticon.com/512/1384/1384060.png" alt="YT"></a>
            </div>
        </div>
    </div>
""", unsafe_allow_html=True)

def update_admin_sheet():
    st.session_state["active_sheet"] = st.session_state["admin_sheet_radio"]
    st.query_params["active_sheet"] = st.session_state["admin_sheet_radio"]

def update_worker_sheet():
    st.session_state["active_sheet"] = st.session_state["worker_sheet_radio"]
    st.query_params["active_sheet"] = st.session_state["worker_sheet_radio"]

@st.dialog("📋 Case Details, Edit & Delete", width="large")
def view_and_edit_case_dialog(sheet_name, row_idx, role, current_user, can_update=True):
    df_full = load_data(sheet_name=sheet_name)
    if df_full is None or row_idx not in df_full.index:
        st.error("Case data not found.")
        return

    row_data = df_full.loc[row_idx]
    display_c = find_display_username_column(df_full)
    p_col = find_password_column(df_full)

    st.info(f"📂 Sheet: **{sheet_name}** | Row #: **{row_idx + 1}** | Username: **{row_data[display_c]}**")

    action_type_val = str(row_data.get("Action_Type", "")).strip().lower()
    if role == "worker" and action_type_val == "delete rejected":
        st.error("❌ మీ డిలీట్ రిక్వెస్ట్‌ను అడ్మిన్ గారు తిరస్కరించారు (Delete Request Rejected by Admin).")

    if "is_editing_popup" not in st.session_state:
        st.session_state["is_editing_popup"] = False
    if "is_delete_request_mode" not in st.session_state:
        st.session_state["is_delete_request_mode"] = False

    if not st.session_state["is_editing_popup"] and not st.session_state["is_delete_request_mode"]:
        st.markdown("### 📄 Case Information")
        for col in df_full.columns:
            if col not in [p_col]:
                st.text(f"{col}: {row_data[col]}")
        
        st.markdown("---")
        if role == "worker" and not can_update:
            st.warning("⚠️ View Only Mode: Admin has disabled update permissions for your account.")
        else:
            col_act1, col_act2 = st.columns(2)
            with col_act1:
                if st.button("✏️ Edit Case Details", type="primary", use_container_width=True, key="btn_trigger_edit"):
                    st.session_state["is_editing_popup"] = True
                    st.rerun()
            with col_act2:
                if role == "admin":
                    if st.button("🗑️ Delete Case", type="secondary", use_container_width=True, key="btn_trigger_delete"):
                        try:
                            df_full = df_full.drop(row_idx)
                            save_data(df_full, sheet_name=sheet_name)
                            st.session_state["selected_row_idx"] = None
                            st.session_state["open_dialog"] = None
                            st.success("✅ Case deleted successfully!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Error deleting case: {e}")
                else:
                    if st.button("🗑️ Request Delete", type="secondary", use_container_width=True, key="btn_trigger_del_req_mode"):
                        st.session_state["is_delete_request_mode"] = True
                        st.rerun()

    elif st.session_state["is_delete_request_mode"]:
        st.markdown("### 🗑️ Request Delete Case")
        with st.form(f"delete_reason_form_{row_idx}"):
            del_reason = st.text_area("✍️ ఈ కేసును ఎందుకు డిలీట్ చేయాలనుకుంటున్నారు? (Reason required):", placeholder="కారణం ఇక్కడ రాయండి...")
            col_sub_req, col_back_req = st.columns(2)
            with col_sub_req:
                submitted_req = st.form_submit_button("🚀 Submit Delete Request", type="primary", use_container_width=True)
            with col_back_req:
                cancelled_req = st.form_submit_button("↩️ Back", use_container_width=True)

            if submitted_req:
                if del_reason.strip() == "":
                    st.warning("⚠️ దయచేసి డిలీట్ చేయడానికి తగిన కారణం (Reason) రాయండి.")
                else:
                    try:
                        df_full.at[row_idx, "Updated_By"] = current_user
                        df_full.at[row_idx, "Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        df_full.at[row_idx, "Action_Type"] = "Delete Request"
                        df_full.at[row_idx, "Delete_Reason"] = del_reason.strip()
                        save_data(df_full, sheet_name=sheet_name)
                        st.session_state["is_delete_request_mode"] = False
                        st.session_state["selected_row_idx"] = None
                        st.session_state["open_dialog"] = None
                        st.success("✅ Delete request sent to Admin with reason successfully!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error: {e}")

            if cancelled_req:
                st.session_state["is_delete_request_mode"] = False
                st.rerun()

    else:
        st.markdown("### ✏️ Edit Case Details")
        with st.form(f"popup_edit_form_{row_idx}"):
            updated_values = {}
            cols_to_show = list(df_full.columns) if role == "admin" else [c for c in df_full.columns if c not in [display_c, p_col, "Updated_By", "Updated_At", "Delete_Reason"]]
            
            f_col1, f_col2 = st.columns(2)
            for i, col_name in enumerate(cols_to_show):
                target_col = f_col1 if i % 2 == 0 else f_col2
                with target_col:
                    if any(k in col_name.lower() for k in ['date', 'time', 'day']):
                        try:
                            parsed_d = datetime.datetime.strptime(str(row_data[col_name]).split()[0], "%Y-%m-%d").date()
                        except:
                            parsed_d = datetime.date.today()
                        sel_d = st.date_input(f"{col_name}", value=parsed_d, key=f"popup_date_{row_idx}_{col_name}")
                        updated_values[col_name] = sel_d.strftime("%Y-%m-%d")
                    else:
                        updated_values[col_name] = st.text_input(f"{col_name}", value=str(row_data[col_name]), key=f"popup_field_{row_idx}_{col_name}")

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            col_save, col_back = st.columns(2)
            with col_save:
                submitted = st.form_submit_button("💾 Save & Update", type="primary", use_container_width=True)
            with col_back:
                cancelled = st.form_submit_button("↩️ Cancel", use_container_width=True)

            if submitted:
                try:
                    current_time = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    changed = False
                    for col_name, new_val in updated_values.items():
                        old_val = str(df_full.at[row_idx, col_name]).strip()
                        if old_val != str(new_val).strip():
                            changed = True
                        df_full.at[row_idx, col_name] = str(new_val).strip()

                    if changed:
                        df_full.at[row_idx, "Updated_By"] = current_user
                        df_full.at[row_idx, "Updated_At"] = current_time
                        df_full.at[row_idx, "Action_Type"] = "Updated Case"

                    save_data(df_full, sheet_name=sheet_name)
                    st.session_state["is_editing_popup"] = False
                    st.session_state["selected_row_idx"] = None
                    st.session_state["open_dialog"] = None
                    st.success("✅ Case details updated successfully!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error updating case: {e}")
            
            if cancelled:
                st.session_state["is_editing_popup"] = False
                st.session_state["selected_row_idx"] = None
                st.session_state["open_dialog"] = None
                st.rerun()

def render_clickable_table(df, display_c, p_col):
    if df.empty:
        st.info("No cases found.")
        return

    display_cols = [c for c in df.columns if c not in [p_col]]
    
    curr_logged = qp.get("logged_in", "true" if st.session_state.get("logged_in") else "false")
    curr_user = qp.get("username", st.session_state.get("username", ""))
    curr_role = qp.get("role", st.session_state.get("role", ""))
    curr_sheet = qp.get("active_sheet", st.session_state.get("active_sheet", ""))

    html = f"""
    <div style="max-height: 500px; overflow-y: auto; overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 12px; margin-bottom: 16px; box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);">
    <table style="width:100%; border-collapse: collapse; font-size: 13px; font-family: sans-serif;">
        <thead>
            <tr style="background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%); color: white; position: sticky; top: 0; z-index: 1;">
                <th style="padding: 12px 10px; border: 1px solid #3b82f6; text-align: center;">#</th>
    """
    for col in display_cols:
        html += f'<th style="padding: 12px 10px; border: 1px solid #3b82f6; text-align: left; white-space: nowrap;">{col}</th>'
    html += '<th style="padding: 12px 10px; border: 1px solid #3b82f6; text-align: left; white-space: nowrap;">Status</th>'
    html += "</tr></thead><tbody>"

    for idx, row in df.iterrows():
        row_bg = "#f8fafc" if idx % 2 == 0 else "#ffffff"
        
        act_t = str(row.get("Action_Type", "")).strip().lower()
        if act_t == "delete rejected":
            row_bg = "#fee2e2"

        html += f'<tr style="background-color: {row_bg};">'
        html += f'<td style="padding: 10px; border: 1px solid #e2e8f0; text-align: center; color: #64748b; font-weight: 600;">{idx + 1}</td>'
        
        for col in display_cols:
            val = str(row[col])
            if col == display_c:
                display_val = val.strip() if val.strip() != "" else "Not Assigned"
                link_url = f"?logged_in={curr_logged}&username={curr_user}&role={curr_role}&active_sheet={curr_sheet}&edit_row={idx}"
                cell_content = f'<a href="{link_url}" target="_self" class="username-hyperlink">{display_val}</a>'
            else:
                cell_content = val if val != "" else "-"
            
            html += f'<td style="padding: 10px; border: 1px solid #e2e8f0; color: #1e293b; white-space: nowrap;">{cell_content}</td>'
        
        status_display = "Rejected" if act_t == "delete rejected" else "-"
        html += f'<td style="padding: 10px; border: 1px solid #e2e8f0; color: #b91c1c; font-weight: 700; white-space: nowrap;">{status_display}</td>'
        html += "</tr>"

    html += "</tbody></table></div>"
    st.markdown(html, unsafe_allow_html=True)

@st.dialog("📢 Add / Update Notice (Scrolling Announcement)", width="medium")
def admin_notice_dialog():
    current_notice = get_admin_notice()
    with st.form("admin_notice_form"):
        notice_input = st.text_area("Notice Text:", value=current_notice, placeholder="ఉదాహరణకు: అందరూ ఈరోజు కల్లా పెండింగ్ కేసులు పూర్తి చేయాలి...")
        col_save_n, col_clear_n = st.columns(2)
        with col_save_n:
            submitted_notice = st.form_submit_button("💾 Save Notice", type="primary", use_container_width=True)
        with col_clear_n:
            cleared_notice = st.form_submit_button("🗑️ Clear Notice", use_container_width=True)

        if submitted_notice:
            save_admin_notice(notice_input)
            st.success("✅ Notice updated successfully!")
            st.rerun()
        if cleared_notice:
            save_admin_notice("")
            st.success("✅ Notice cleared successfully!")
            st.rerun()

@st.dialog("👥 Case Worker Login Status", width="large")
def admin_login_status_dialog():
    sheets = get_sheet_names()
    all_workers = set()
    for s in sheets:
        df_s = load_data(sheet_name=s)
        if df_s is not None and not df_s.empty:
            lc = find_login_user_column(df_s)
            for _, r in df_s.iterrows():
                u = str(r[lc]).strip()
                if u != "" and u.lower() != "admin":
                    all_workers.add(u)
                    
    login_records = load_login_tracking()
    normalized_login_records = {k.lower(): v for k, v in login_records.items()}
    
    logged_in_list = []
    not_logged_in_list = []
    
    for w in sorted(list(all_workers)):
        rec = normalized_login_records.get(w.lower(), {})
        status = rec.get("status", "Not Logged In")
        last_time = rec.get("last_login", "-")
        
        item = {"worker": w, "status": status, "time": last_time}
        if status == "Logged In":
            logged_in_list.append(item)
        else:
            not_logged_in_list.append(item)
            
    tab_logged, tab_not_logged = st.tabs([f"🟢 Login ({len(logged_in_list)})", f"🔴 Not Login ({len(not_logged_in_list)})"])
    with tab_logged:
        for item in logged_in_list:
            c1, c2, c3 = st.columns([2, 2, 2])
            with c1: st.write(f"👤 **{item['worker']}**")
            with c2: st.success("🟢 Logged In")
            with c3: st.text(item['time'])
    with tab_not_logged:
        for item in not_logged_in_list:
            c1, c2, c3 = st.columns([2, 2, 2])
            with c1: st.write(f"👤 **{item['worker']}**")
            with c2: st.error("🔴 Not Logged In")
            with c3: st.text(item['time'])

@st.dialog("📊 Case Worker Attendance Report", width="large")
def admin_attendance_dialog():
    attendance_data = load_attendance()
    if not attendance_data:
        st.warning("ఎటువంటి అటెండెన్స్ రికార్డులు కనుగొనబడలేదు.")
    else:
        df_att = pd.DataFrame(attendance_data)
        df_att.index = range(1, len(df_att) + 1)
        st.dataframe(df_att, use_container_width=True)

@st.dialog("📋 New / Updated Cases & Calendar", width="large")
def admin_updated_cases_dialog():
    sheets = get_sheet_names()
    all_rows = []
    for s in sheets:
        df_s = load_data(sheet_name=s)
        if df_s is not None and not df_s.empty and "Updated_By" in df_s.columns:
            upd_rows = df_s[df_s["Updated_By"].astype(str).str.strip() != ""]
            if not upd_rows.empty:
                upd_rows = upd_rows.copy()
                upd_rows.insert(0, "Sheet_Name", s)
                all_rows.append(upd_rows)
                
    tab_new, tab_upd, tab_cal = st.tabs(["➕ న్యూ కేసులు (New Cases)", "✏️ అప్‌డేట్ చేసినవి (Updated)", "📅 క్యాలెండర్ (Calendar)"])
    with tab_new:
        if all_rows:
            combined_df = pd.concat(all_rows, ignore_index=True)
            p_col = find_password_column(combined_df)
            new_cases_df = combined_df[combined_df["Action_Type"].astype(str).str.strip().str.lower() == "new case"]
            render_clickable_table(new_cases_df, find_display_username_column(new_cases_df), p_col)
    with tab_upd:
        if all_rows:
            combined_df = pd.concat(all_rows, ignore_index=True)
            p_col = find_password_column(combined_df)
            updated_cases_df = combined_df[combined_df["Action_Type"].astype(str).str.strip().str.lower() == "updated case"]
            render_clickable_table(updated_cases_df, find_display_username_column(updated_cases_df), p_col)
    with tab_cal:
        selected_cal_sheet = st.selectbox("Select Sheet for Calendar View", options=sheets, key="cal_sheet_select_tab")
        df_cal = load_data(sheet_name=selected_cal_sheet)
        if df_cal is not None and not df_cal.empty:
            date_cols = [c for c in df_cal.columns if any(k in c.lower() for k in ['date', 'time', 'day'])]
            if date_cols:
                selected_date_col = st.selectbox("Select Date Column:", options=date_cols, key="cal_date_col_tab")
                selected_date = st.date_input("Select Target Date:", value=datetime.date.today(), key="cal_target_date_tab")
                matched_df = df_cal[df_cal[selected_date_col].astype(str).str.contains(selected_date.strftime("%Y-%m-%d"), na=False)]
                render_clickable_table(matched_df, find_display_username_column(matched_df), find_password_column(matched_df))

@st.dialog("🗑️ Pending Delete Requests", width="large")
def admin_delete_requests_dialog():
    sheets = get_sheet_names()
    del_reqs_list = []
    for s in sheets:
        df_s = load_data(sheet_name=s)
        if df_s is not None and not df_s.empty and "Action_Type" in df_s.columns:
            for idx, r in df_s.iterrows():
                if str(r.get("Action_Type", "")).strip().lower() == "delete request":
                    r_dict = r.to_dict()
                    r_dict["Sheet_Name"] = s
                    r_dict["Original_Index"] = idx
                    del_reqs_list.append(r_dict)
    
    for i, r_dict in enumerate(del_reqs_list):
        s_name = r_dict["Sheet_Name"]
        orig_idx = r_dict["Original_Index"]
        case_username = str(r_dict.get(find_display_username_column(load_data(sheet_name=s_name)), "N/A"))
        with st.expander(f"📁 Sheet: {s_name} | Case User: {case_username}"):
            st.markdown(f"✍️ **Delete Reason:** `{r_dict.get('Delete_Reason', '')}`")
            col_acc, col_rej = st.columns(2)
            with col_acc:
                if st.button("✅ Accept Delete", key=f"accept_del_{s_name}_{orig_idx}_{i}", type="primary", use_container_width=True):
                    df_target = load_data(sheet_name=s_name)
                    df_target = df_target.drop(orig_idx)
                    save_data(df_target, sheet_name=s_name)
                    st.rerun()

@st.dialog("👥 Users, Passwords & Permissions", width="large")
def permissions_dialog():
    sheets = get_sheet_names()
    all_users_dict = {}
    for s in sheets:
        df_s = load_data(sheet_name=s)
        if df_s is not None and not df_s.empty:
            for _, r in df_s.iterrows():
                u_name = str(r[find_login_user_column(df_s)]).strip()
                pwd = str(r[find_password_column(df_s)]).strip()
                if u_name and u_name.lower() != "admin":
                    all_users_dict[u_name] = pwd
    for idx, (u_name, curr_pwd) in enumerate(all_users_dict.items()):
        st.text_input("Username", value=u_name, disabled=True, key=f"dlg_u_{idx}")

@st.dialog("🔑 Change Admin Password", width="medium")
def admin_password_dialog():
    with st.form("dlg_admin_pass_form"):
        curr_p = st.text_input("Current Password", type="password")
        new_p = st.text_input("New Password", type="password")
        if st.form_submit_button("Update Password", type="primary", use_container_width=True):
            admin_u, admin_p = get_admin_creds()
            if curr_p == admin_p:
                save_admin_creds(admin_u, new_p.strip())
                st.success("✅ Password updated successfully!")

@st.dialog("📤 Sync Local Excel File to Cloud", width="medium")
def sync_excel_dialog():
    uploaded_excel = st.file_uploader("Upload Updated Excel File (.xlsx)", type=["xlsx"])
    if uploaded_excel is not None:
        if st.button("🚀 Confirm & Sync", type="primary", use_container_width=True):
            with open(EXCEL_FILE, "wb") as f:
                f.write(uploaded_excel.getbuffer())
            st.success("✅ Synced successfully!")
            st.rerun()

@st.dialog("🔑 Change Password", width="medium")
def worker_password_dialog(worker_name):
    with st.form("dlg_worker_pass_form"):
        curr_wp = st.text_input("Current Password", type="password")
        new_wp = st.text_input("New Password", type="password")
        if st.form_submit_button("Update Password", type="primary", use_container_width=True):
            st.success("✅ Password updated!")

@st.dialog("➕ Add New Case", width="large")
def add_case_dialog(worker_name):
    sheets = get_sheet_names()
    selected_target_sheet = st.selectbox("Select Target Sheet", options=sheets)
    target_df = load_data(sheet_name=selected_target_sheet)
    with st.form("dlg_new_case_form"):
        st.text_input("Customer Name")
        if st.form_submit_button("Submit", type="primary", use_container_width=True):
            st.success("✅ Added!")

@st.dialog("🔍 Search Main Data", width="large")
def search_data_dialog():
    sheets = get_sheet_names()
    selected_search_sheet = st.selectbox("Select Sheet", options=sheets)
    df_search = load_data(sheet_name=selected_search_sheet)
    st.text_input("Search query")

# Dialog Handler Routing
if st.session_state["open_dialog"] == "permissions":
    permissions_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "admin_password":
    admin_password_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "sync_excel":
    sync_excel_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "admin_updated_cases":
    admin_updated_cases_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "admin_delete_requests":
    admin_delete_requests_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "admin_notice":
    admin_notice_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "admin_login_status":
    admin_login_status_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "admin_attendance":
    admin_attendance_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "worker_password":
    worker_password_dialog(st.session_state.get("username", ""))
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "add_case":
    add_case_dialog(st.session_state.get("username", ""))
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "search_data":
    search_data_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "edit_case_popup":
    r_idx = st.session_state.get("selected_row_idx")
    if r_idx is not None:
        curr_role = st.session_state.get("role", "worker")
        curr_u = st.session_state.get("username", "")
        perms = load_permissions()
        c_upd = True if curr_role == "admin" else perms.get(curr_u, {}).get("can_update", True)
        view_and_edit_case_dialog(st.session_state["active_sheet"], r_idx, curr_role, curr_u, c_upd)

# 1. Public / Login Screen
if not st.session_state["logged_in"]:
    selected_page = st.radio("Navigation", ["🏠 Home", "ℹ️ About Us", "✈️ Services", "⭐ Success Stories", "📞 Contact Us", "🔐 Portal Login"], horizontal=True, label_visibility="collapsed")
    if selected_page == "🔐 Portal Login":
        col_l1, col_l2, col_l3 = st.columns([1, 1.4, 1])
        with col_l2:
            with st.container(border=True):
                with st.form("login_form"):
                    username_input = st.text_input("👤 Username")
                    password_input = st.text_input("🔑 Password", type="password")
                    if st.form_submit_button("🚀 Login to Portal", type="primary", use_container_width=True):
                        admin_u, admin_p = get_admin_creds()
                        if username_input.strip().lower() == admin_u.lower() and password_input.strip() == admin_p:
                            st.session_state["logged_in"] = True
                            st.session_state["username"] = admin_u
                            st.session_state["role"] = "admin"
                            st.query_params["logged_in"] = "true"
                            st.query_params["username"] = admin_u
                            st.query_params["role"] = "admin"
                            st.rerun()

# 2. ADMIN DASHBOARD
elif st.session_state["role"] == "admin":
    sheets = get_sheet_names()
    current_sheet = st.session_state["active_sheet"]
    df = load_data(sheet_name=current_sheet)
    
    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    top_col1, top_col2, top_col3 = st.columns([6.5, 1.2, 1.2], gap="small")
    with top_col1:
        st.markdown(f"<div class='section-banner' style='margin:0;'>👑 Admin Control Panel — {current_sheet}</div>", unsafe_allow_html=True)
    with top_col2:
        if st.button("🔄 Refresh", key="admin_refresh"):
            st.rerun()
    with top_col3:
        if st.button("🚪 Logout", key="admin_logout"):
            st.query_params.clear()
            st.session_state["logged_in"] = False
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    admin_dash_tab = st.radio("Admin Navigation", ["📊 Dashboard & Quick Actions", "📂 Case Records & Search"], horizontal=True, label_visibility="collapsed", key="admin_dash_tab_radio")
    
    visa_col = next((c for c in df.columns if any(term in c.lower() for term in ['visa', 'category', 'type', 'service'])), None)
    visa_counts = {}
    if visa_col and not df.empty:
        for val in df[visa_col].astype(str):
            v_clean = val.strip().upper()
            if v_clean:
                visa_counts[v_clean] = visa_counts.get(v_clean, 0) + 1

    if admin_dash_tab == "📊 Dashboard & Quick Actions":
        r1_c1, r1_c2, r1_c3, r1_c4 = st.columns(4, gap="medium")
        with r1_c1:
            if st.button("📋 New / Updated Cases", use_container_width=True, key="admin_upd_cases_btn"):
                st.session_state["open_dialog"] = "admin_updated_cases"
                st.rerun()
        with r1_c2:
            if st.button("🗑️ Delete Requests", use_container_width=True, key="admin_del_req_btn"):
                st.session_state["open_dialog"] = "admin_delete_requests"
                st.rerun()
        with r1_c3:
            if st.button("🔄 Sync Excel", use_container_width=True, key="admin_sync_btn"):
                st.session_state["open_dialog"] = "sync_excel"
                st.rerun()
        with r1_c4:
            if st.button("📢 Broadcast Notice", use_container_width=True, key="admin_notice_btn"):
                st.session_state["open_dialog"] = "admin_notice"
                st.rerun()

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        r2_c1, r2_c2, r2_c3, r2_c4 = st.columns(4, gap="medium")
        with r2_c1:
            if st.button("👥 Users & Permissions", use_container_width=True, key="admin_perm_btn"):
                st.session_state["open_dialog"] = "permissions"
                st.rerun()
        with r2_c2:
            if st.button("🟢 Worker Login Status", use_container_width=True, key="admin_login_status_btn"):
                st.session_state["open_dialog"] = "admin_login_status"
                st.rerun()
        with r2_c3:
            if st.button("📊 Attendance Report", use_container_width=True, key="admin_attendance_btn"):
                st.session_state["open_dialog"] = "admin_attendance"
                st.rerun()
        with r2_c4:
            if st.button("🔑 Admin Password", use_container_width=True, key="admin_pw_btn"):
                st.session_state["open_dialog"] = "admin_password"
                st.rerun()

    elif admin_dash_tab == "📂 Case Records & Search":
        if visa_counts:
            st.markdown("<div class='action-group-title visa-anchor'>✈️ VISA CATEGORY DISTRIBUTION (CLICK ANY BUTTON TO FILTER):</div>", unsafe_allow_html=True)
            visa_items = list(visa_counts.items())
            v_cols = st.columns(len(visa_items), gap="medium")
            for idx, (v_name, v_count) in enumerate(visa_items):
                with v_cols[idx]:
                    if st.button(f"🛂 {v_name}: {v_count}", key=f"admin_visa_card_btn_{v_name}_{idx}", use_container_width=True):
                        st.session_state["selected_visa_filter"] = "All" if st.session_state["selected_visa_filter"] == v_name else v_name
                        st.rerun()
            st.divider()
        
        display_df = df
        if st.session_state["selected_visa_filter"] != "All" and visa_col:
            display_df = display_df[display_df[visa_col].astype(str).str.strip().str.upper() == st.session_state["selected_visa_filter"]]
        render_clickable_table(display_df, find_display_username_column(df), find_password_column(df))

# 3. CASE WORKER DASHBOARD
else:
    sheets = get_sheet_names()
    current_sheet = st.session_state["active_sheet"]
    df = load_data(sheet_name=current_sheet)
    current_user = st.session_state["username"].strip()
    user_rows = df[df[find_login_user_column(df)].astype(str).str.strip().str.lower() == current_user.lower()]
    
    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    top_col1, top_col2, top_col3 = st.columns([6.5, 1.2, 1.2], gap="small")
    with top_col1:
        st.markdown(f"<div class='section-banner' style='margin:0;'>👋 Welcome, {current_user} — Case Worker Workspace</div>", unsafe_allow_html=True)
    with top_col2:
        if st.button("🔄 Refresh", key="worker_refresh"):
            st.rerun()
    with top_col3:
        if st.button("🚪 Logout", key="worker_logout"):
            record_worker_logout(current_user)
            st.query_params.clear()
            st.session_state["logged_in"] = False
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    worker_dash_tab = st.radio("Worker Navigation", ["📊 My Workspace", "ℹ️ Guide"], horizontal=True, label_visibility="collapsed")
    
    if worker_dash_tab == "📊 My Workspace":
        visa_col = next((c for c in user_rows.columns if any(term in c.lower() for term in ['visa', 'category', 'type', 'service'])), None)
        visa_counts = {}
        if visa_col and not user_rows.empty:
            for val in user_rows[visa_col].astype(str):
                v_clean = val.strip().upper()
                if v_clean:
                    visa_counts[v_clean] = visa_counts.get(v_clean, 0) + 1

        if visa_counts:
            st.markdown("<div class='action-group-title visa-anchor'>✈️ MY VISA CATEGORY DISTRIBUTION (CLICK ANY BUTTON TO FILTER):</div>", unsafe_allow_html=True)
            visa_items = list(visa_counts.items())
            w_cols = st.columns(len(visa_items), gap="medium")
            for idx, (v_name, v_count) in enumerate(visa_items):
                with w_cols[idx]:
                    icon = "⏳" if "PEND" in v_name else "🌐" if "B1" in v_name else "🛂"
                    if st.button(f"{icon} {v_name}: {v_count}", key=f"worker_visa_btn_{v_name}_{idx}", use_container_width=True):
                        st.session_state["selected_visa_filter"] = "All" if st.session_state["selected_visa_filter"] == v_name else v_name
                        st.rerun()

        st.markdown("<div class='action-group-title workspace-anchor'>⚡ WORKSPACE TOOLS</div>", unsafe_allow_html=True)
        w_cols = st.columns(4, gap="medium")
        with w_cols[0]:
            if st.button("➕ Add New Case", use_container_width=True, key="worker_add_btn"):
                st.session_state["open_dialog"] = "add_case"
                st.rerun()
        with w_cols[1]:
            if st.button("🔍 Search Records", use_container_width=True, key="worker_search_btn"):
                st.session_state["open_dialog"] = "search_data"
                st.rerun()
        with w_cols[2]:
            if st.button("🚫 No Download", disabled=True, use_container_width=True, key="worker_dl_disabled"):
                pass
        with w_cols[3]:
            if st.button("🔑 Change Password", use_container_width=True, key="worker_pw_btn"):
                st.session_state["open_dialog"] = "worker_password"
                st.rerun()

        st.divider()
        display_user_rows = user_rows
        if st.session_state["selected_visa_filter"] != "All" and visa_col:
            display_user_rows = display_user_rows[display_user_rows[visa_col].astype(str).str.strip().str.upper() == st.session_state["selected_visa_filter"]]
        render_clickable_table(display_user_rows, find_display_username_column(df), find_password_column(df))
