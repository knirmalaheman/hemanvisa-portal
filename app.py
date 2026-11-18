import streamlit as st
import pandas as pd
import os
import io
import datetime
import base64
import json
import math

EXCEL_FILE = "Heman_Report.xlsx"
ADMIN_FILE = "admin_config.txt"
PERMS_FILE = "worker_permissions.json"
NOTICE_FILE = "admin_notice.txt"
LOGINS_FILE = "worker_logins_status.json"
ATTENDANCE_FILE = "worker_attendance.json"
PDF_DIR = "uploaded_appointment_pdfs"
LOGO_FILES = ["pngTShirtPrintCricket.jpg", "logo.jpg", "logo.png", "logo.jpeg"]

if not os.path.exists(PDF_DIR):
    os.makedirs(PDF_DIR)

st.set_page_config(page_title="HemanVisa Portal", page_icon="✈️", layout="wide")

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
    for k in list(logins.keys()):
        if k.lower() == worker_name.strip().lower():
            logins[k]["status"] = "Not Logged In"
    with open(LOGINS_FILE, "w") as f:
        json.dump(logins, f, indent=4)
        
    attendance = load_attendance()
    for entry in attendance:
        if entry.get("Worker", "").lower() == worker_name.strip().lower() and entry.get("Date") == today_str:
            entry["Logout_Time"] = time_str
            entry["Status"] = "Logged Out"
            break
    save_attendance(attendance)

# 🌟 STRICT COLUMN REORDERING: Case, Username, Password, No of App always first
def reorder_dataframe_columns(df):
    if df is None or df.empty:
        return df
    cols = list(df.columns)
    
    # Priority columns required at the very beginning
    priority_keys = ["case", "username", "password", "no of app", "no of application"]
    found_priority = []
    
    for pk in priority_keys:
        for c in cols:
            c_clean = c.strip().lower().replace("_", " ")
            if c_clean == pk and c not in found_priority:
                found_priority.append(c)
                
    for c in cols:
        c_clean = c.strip().lower().replace("_", " ")
        if any(pk in c_clean for pk in priority_keys) and c not in found_priority:
            found_priority.append(c)
            
    remaining_cols = [c for c in cols if c not in found_priority]
    return df[found_priority + remaining_cols]

@st.cache_data(ttl=600)
def load_all_sheets_cached(file_mtime):
    if os.path.exists(EXCEL_FILE):
        try:
            all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
            processed = {}
            for s_name, df in all_sheets.items():
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
                if "Booking_Status" not in df.columns:
                    df["Booking_Status"] = "Active"
                if "Appointment_Entry" not in df.columns:
                    df["Appointment_Entry"] = ""
                if "Appointment_PDF" not in df.columns:
                    df["Appointment_PDF"] = ""
                if "Preferred_Location" not in df.columns:
                    df["Preferred_Location"] = ""
                if "Preferred_Dates" not in df.columns:
                    df["Preferred_Dates"] = ""
                if "Prior_Notice" not in df.columns:
                    df["Prior_Notice"] = ""
                
                df = reorder_dataframe_columns(df)
                processed[s_name] = df
            return processed
        except Exception:
            return {}
    return {}

@st.cache_data(ttl=600)
def get_worker_export_bytes(file_mtime, current_user, current_sheet):
    all_sheets = load_all_sheets_cached(file_mtime)
    worker_buffer = io.BytesIO()
    with pd.ExcelWriter(worker_buffer, engine='openpyxl') as writer:
        has_user_data = False
        for s, df_s in all_sheets.items():
            if df_s is not None and not df_s.empty:
                login_cs = find_login_user_column(df_s)
                u_rows_s = df_s[df_s[login_cs].astype(str).str.strip().str.lower() == current_user.lower()]
                if not u_rows_s.empty:
                    has_user_data = True
                    u_rows_s.to_excel(writer, sheet_name=s, index=False)
        if not has_user_data and current_sheet in all_sheets:
            df_cur = all_sheets[current_sheet]
            login_cs = find_login_user_column(df_cur)
            u_rows_cur = df_cur[df_cur[login_cs].astype(str).str.strip().str.lower() == current_user.lower()]
            u_rows_cur.to_excel(writer, sheet_name=current_sheet, index=False)
    return worker_buffer.getvalue()

def _get_file_mtime():
    return os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0

def get_sheet_names():
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    if all_sheets:
        return list(all_sheets.keys())
    return ["Sheet1"]

def load_data(sheet_name=0):
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    if not all_sheets:
        return None
    if isinstance(sheet_name, int):
        keys = list(all_sheets.keys())
        if 0 <= sheet_name < len(keys):
            return all_sheets[keys[sheet_name]].copy()
        return None
    if sheet_name in all_sheets:
        return all_sheets[sheet_name].copy()
    return None

def save_data(df_to_save, sheet_name=0):
    if "Category" in df_to_save.columns:
        df_to_save = df_to_save.rename(columns={"Category": ""})
    
    sheets_data = {}
    if os.path.exists(EXCEL_FILE):
        try:
            sheets_data = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
        except Exception:
            pass
    
    sheets_data[sheet_name] = df_to_save

    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        for s, df_val in sheets_data.items():
            df_val = reorder_dataframe_columns(df_val)
            df_val.to_excel(writer, sheet_name=str(s), index=False)
            
    st.cache_data.clear()

@st.cache_data
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
        if col.strip().upper() == "USERNAME":
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

def find_visa_type_column(df):
    for c in df.columns:
        if any(term in c.lower() for term in ['visa', 'category', 'type', 'service']):
            return c
    return None

def get_all_workers_list():
    sheets = get_sheet_names()
    all_workers = set()
    for s in sheets:
        df_s = load_data(sheet_name=s)
        if df_s is not None and not df_s.empty:
            lc = find_login_user_column(df_s)
            users = df_s[lc].astype(str).str.strip()
            for u in users.unique():
                if u != "" and u.lower() != "admin":
                    all_workers.add(u)
    return sorted(list(all_workers))

def are_all_workers_blocked():
    workers = get_all_workers_list()
    if not workers:
        return False
    perms = load_permissions()
    norm_perms = {k.lower(): v for k, v in perms.items()}
    return all(norm_perms.get(w.lower(), {}).get("block_login", False) for w in workers)

def is_worker_blocked(worker_name):
    if not worker_name:
        return False
    perms = load_permissions()
    for k, v in perms.items():
        if k.lower() == worker_name.strip().lower():
            return v.get("block_login", False)
    return False

def toggle_all_workers_block(block_status: bool):
    all_workers = get_all_workers_list()
    current_perms = load_permissions()
    for w in all_workers:
        matched_key = next((k for k in current_perms if k.lower() == w.lower()), w)
        if matched_key not in current_perms:
            current_perms[matched_key] = {"can_update": True, "can_add": True, "can_search": True, "can_download": False, "can_sync": False, "block_login": block_status}
        else:
            current_perms[matched_key]["block_login"] = block_status
        if block_status:
            record_worker_logout(w)
    save_permissions(current_perms)

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

if "force_logged_out_msg" not in st.session_state:
    st.session_state["force_logged_out_msg"] = False

if not st.session_state["logged_in"] and qp.get("logged_in") == "true":
    qp_role = qp.get("role", "worker")
    qp_user = qp.get("username", "")
    if qp_role == "worker" and is_worker_blocked(qp_user):
        record_worker_logout(qp_user)
        qp.clear()
        st.session_state["force_logged_out_msg"] = True
    else:
        st.session_state["logged_in"] = True
        st.session_state["username"] = qp_user
        st.session_state["role"] = qp_role
        if qp_role == "worker" and not st.session_state["attendance_recorded"]:
            record_worker_login(qp_user)
            st.session_state["attendance_recorded"] = True

if st.session_state.get("logged_in") and st.session_state.get("role") == "worker":
    curr_w = st.session_state.get("username", "")
    if is_worker_blocked(curr_w):
        record_worker_logout(curr_w)
        st.query_params.clear()
        st.session_state["logged_in"] = False
        st.session_state["username"] = ""
        st.session_state["role"] = ""
        st.session_state["attendance_recorded"] = False
        st.session_state["open_dialog"] = None
        st.session_state["selected_row_idx"] = None
        st.session_state["force_logged_out_msg"] = True
        st.rerun()

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
        st.session_state["is_reschedule_mode"] = False
        if "edit_row" in st.query_params:
            del st.query_params["edit_row"]
    except:
        if "edit_row" in st.query_params:
            del st.query_params["edit_row"]
elif not st.session_state["logged_in"] and "edit_row" in qp:
    qp.clear()

if "active_sheet" not in st.session_state:
    sheets_init = get_sheet_names()
    saved_sheet = qp.get("active_sheet")
    if saved_sheet in sheets_init:
        st.session_state["active_sheet"] = saved_sheet
    else:
        st.session_state["active_sheet"] = sheets_init[0] if sheets_init else "Sheet1"

if st.session_state.get("logged_in") and st.session_state.get("role") == "worker":
    if not st.session_state.get("attendance_recorded", False):
        record_worker_login(st.session_state["username"])
        st.session_state["attendance_recorded"] = True

# 🌟 STYLING
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
        background: linear-gradient(135deg, #070f24 0%, #0f1c3f 50%, #172554 100%);
        padding: 10px 28px;
        border-bottom: 3px solid #d4af37;
        box-shadow: 0 10px 35px rgba(7, 15, 36, 0.65);
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
        border-radius: 8px;
        background: #ffffff;
        padding: 3px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
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
        letter-spacing: 1.5px;
        text-transform: uppercase;
        margin-bottom: 2px;
        text-shadow: 0 2px 8px rgba(0,0,0,0.6);
    }
    .world-map-subtitle {
        font-size: 12px;
        font-weight: 700;
        color: #fde047;
        letter-spacing: 1.8px;
        text-shadow: 0 1px 4px rgba(0,0,0,0.6);
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
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.2);
        padding: 6px 14px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
        backdrop-filter: blur(8px);
    }
    .banner-info a {
        color: #ffffff;
        text-decoration: none;
    }
    .social-icons {
        display: flex;
        gap: 8px;
        align-items: center;
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.2);
        padding: 5px 12px;
        border-radius: 8px;
    }
    .social-icons img {
        height: 20px;
        width: 20px;
        border-radius: 4px;
        max-height: 20px !important;
        background: transparent !important;
        padding: 0 !important;
    }
    .block-container {
        padding-top: 112px !important;
        padding-bottom: 2.5rem;
        max-width: 1400px !important;
    }
    .hero-section {
        background: linear-gradient(135deg, #0f1c3f 0%, #172554 50%, #0a1128 100%);
        color: white;
        padding: 52px 32px;
        border-radius: 16px;
        text-align: center;
        margin-bottom: 26px;
        box-shadow: 0 12px 35px rgba(15, 28, 63, 0.4);
        border: 1px solid rgba(255,255,255,0.1);
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
        padding: 24px;
        border-radius: 14px;
        box-shadow: 0 8px 25px rgba(15, 23, 42, 0.06);
        border: 1px solid #e2e8f0;
        border-top: 4px solid #172554;
        height: 100%;
        margin-bottom: 14px;
    }
    .marquee-container {
        background: linear-gradient(90deg, #1e293b 0%, #334155 100%);
        color: white;
        padding: 12px 18px;
        border-radius: 10px;
        font-weight: 700;
        font-size: 14px;
        margin-bottom: 20px !important;
        box-shadow: 0 6px 18px rgba(15, 28, 63, 0.2);
        border-left: 5px solid #d4af37;
    }
    .sticky-control-panel {
        position: sticky;
        top: 96px;
        z-index: 99998;
        background-color: #ffffff;
        padding: 10px 0 12px 0;
        border-bottom: 2px solid #e2e8f0;
        margin-bottom: 16px;
    }
    .kpi-card {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        color: white;
        padding: 22px 16px;
        border-radius: 14px;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.15);
        text-align: center;
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-bottom: 4px solid #64748b;
        margin-bottom: 12px;
        height: 115px;
        width: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        box-sizing: border-box;
    }
    .kpi-card.gold {
        background: linear-gradient(135deg, #92400e 0%, #b45309 100%);
        border-bottom: 4px solid #d97706;
    }
    .kpi-card.green {
        background: linear-gradient(135deg, #065f46 0%, #047857 100%);
        border-bottom: 4px solid #059669;
    }
    .kpi-card.purple {
        background: linear-gradient(135deg, #3730a3 0%, #4338ca 100%);
        border-bottom: 4px solid #6366f1;
    }
    .kpi-title {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 1.4px;
        opacity: 0.9;
        margin-bottom: 6px;
        font-weight: 700;
        color: #f1f5f9;
    }
    .kpi-value {
        font-size: 28px;
        font-weight: 800;
        margin: 0;
        line-height: 1.1;
    }
    .section-banner {
        background: linear-gradient(90deg, #f8fafc 0%, #f1f5f9 100%);
        border-left: 5px solid #1e293b;
        padding: 12px 18px;
        border-radius: 8px;
        margin: 8px 0px;
        font-weight: 800;
        color: #1e293b;
        font-size: 16px;
        border: 1px solid #e2e8f0;
    }
    .action-group-title {
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.9px;
        color: #334155;
        margin: 12px 0 8px 2px;
    }
    .stButton > button, [data-testid="stDownloadButton"] > button {
        width: 100% !important;
        min-height: 44px !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-size: 13.5px !important;
        padding: 10px 16px !important;
        color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.2) !important;
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%) !important;
    }
    .username-hyperlink {
        color: #1e293b !important;
        text-decoration: underline !important;
        font-weight: 700 !important;
        cursor: pointer !important;
    }
    div[data-testid="stRadio"] > div {
        flex-direction: row !important;
        flex-wrap: wrap !important;
        gap: 12px !important;
        background: #f8fafc;
        padding: 10px 16px;
        border-radius: 10px;
        border: 1px solid #cbd5e1;
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

# 🌟 ENHANCED DIALOG (Fixed re-trigger loop & ordered view)
@st.dialog("📋 Case Details, Edit & Appointment PDF", width="large")
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
        st.error("❌ Your delete request was rejected by Admin.")

    if "is_editing_popup" not in st.session_state:
        st.session_state["is_editing_popup"] = False
    if "is_delete_request_mode" not in st.session_state:
        st.session_state["is_delete_request_mode"] = False
    if "is_reschedule_mode" not in st.session_state:
        st.session_state["is_reschedule_mode"] = False

    is_booked_sheet = "booked" in sheet_name.lower()

    if not st.session_state["is_editing_popup"] and not st.session_state["is_delete_request_mode"] and not st.session_state["is_reschedule_mode"]:
        st.markdown("### 📄 Case Information")
        
        # 🌟 Reorder columns in case info view so Case, Username, Password, No of App come first
        df_full = reorder_dataframe_columns(df_full)
        row_data = df_full.loc[row_idx]
        
        for col in df_full.columns:
            if col not in [p_col, "Appointment_PDF"]:
                st.text(f"{col}: {row_data[col]}")
        
        curr_b_status = str(row_data.get("Booking_Status", "Active")).strip()
        st.markdown(f"📌 **Status:** `{curr_b_status}`")
        
        app_entry = str(row_data.get("Appointment_Entry", "")).strip()
        if app_entry:
            st.markdown(f"📝 **Current Appointment Entry:** `{app_entry}`")

        pref_loc = str(row_data.get("Preferred_Location", "")).strip()
        pref_dates = str(row_data.get("Preferred_Dates", "")).strip()
        prior_notice = str(row_data.get("Prior_Notice", "")).strip()
        if pref_loc or pref_dates or prior_notice:
            st.markdown(f"🔄 **Reschedule Info:** Location: `{pref_loc}` | Dates: `{pref_dates}` | Prior Notice: `{prior_notice}`")

        pdf_path = str(row_data.get("Appointment_PDF", "")).strip()
        if pdf_path and os.path.exists(pdf_path):
            st.success("📎 Appointment PDF is attached.")
            with open(pdf_path, "rb") as pf:
                pdf_bytes = pf.read()
            st.download_button(
                label="📥 Download Appointment PDF",
                data=pdf_bytes,
                file_name=os.path.basename(pdf_path),
                mime="application/pdf",
                key=f"dl_pdf_view_{row_idx}",
                use_container_width=True
            )
        else:
            st.info("ℹ️ No Appointment PDF uploaded yet.")

        st.markdown("---")
        if role == "worker" and not can_update:
            st.warning("⚠️ View Only Mode: Admin has disabled update permissions for your account.")
        else:
            if is_booked_sheet:
                col_act1, col_act2, col_act3 = st.columns(3)
                with col_act1:
                    if st.button("✏️ Edit Case", type="primary", use_container_width=True, key="btn_trigger_edit"):
                        st.session_state["is_editing_popup"] = True
                        st.rerun()
                with col_act2:
                    if st.button("🔄 Reschedule", type="secondary", use_container_width=True, key="btn_trigger_reschedule"):
                        st.session_state["is_reschedule_mode"] = True
                        st.rerun()
                with col_act3:
                    if role == "admin":
                        if st.button("🗑️ Delete", type="secondary", use_container_width=True, key="btn_trigger_delete"):
                            try:
                                df_full = df_full.drop(row_idx)
                                save_data(df_full, sheet_name=sheet_name)
                                st.session_state["selected_row_idx"] = None
                                st.session_state["open_dialog"] = None
                                st.success("✅ Case deleted successfully!")
                                st.rerun()
                            except Exception as e:
                                st.error(f"❌ Error: {e}")
                    else:
                        if st.button("🗑️ Request Delete", type="secondary", use_container_width=True, key="btn_trigger_del_req_mode"):
                            st.session_state["is_delete_request_mode"] = True
                            st.rerun()
            else:
                col_act1, col_act2 = st.columns(2)
                with col_act1:
                    if st.button("✏️ Edit Case & Update Booking", type="primary", use_container_width=True, key="btn_trigger_edit"):
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
                                st.error(f"❌ Error: {e}")
                    else:
                        if st.button("🗑️ Request Delete", type="secondary", use_container_width=True, key="btn_trigger_del_req_mode"):
                            st.session_state["is_delete_request_mode"] = True
                            st.rerun()

    elif st.session_state["is_reschedule_mode"]:
        st.markdown("### 🔄 Reschedule Appointment Case")
        with st.form(f"reschedule_form_{row_idx}"):
            pref_loc_input = st.text_input("Preferred Location", value=str(row_data.get("Preferred_Location", "")), placeholder="Enter preferred location...")
            pref_dates_input = st.text_input("Preferred Dates", value=str(row_data.get("Preferred_Dates", "")), placeholder="Enter preferred dates...")
            prior_notice_input = st.text_input("Prior Notice", value=str(row_data.get("Prior_Notice", "")), placeholder="Enter prior notice details...")
            
            col_res_save, col_res_back = st.columns(2)
            with col_res_save:
                submitted_res = st.form_submit_button("💾 Save Reschedule", type="primary", use_container_width=True)
            with col_res_back:
                cancelled_res = st.form_submit_button("↩️ Back", use_container_width=True)

            if submitted_res:
                if not pref_loc_input.strip() or not pref_dates_input.strip() or not prior_notice_input.strip():
                    st.warning("⚠️ Please fill all 3 fields (Preferred Location, Preferred Dates, and Prior Notice) before saving.")
                else:
                    try:
                        df_full.at[row_idx, "Preferred_Location"] = pref_loc_input.strip()
                        df_full.at[row_idx, "Preferred_Dates"] = pref_dates_input.strip()
                        df_full.at[row_idx, "Prior_Notice"] = prior_notice_input.strip()
                        df_full.at[row_idx, "Updated_By"] = current_user
                        df_full.at[row_idx, "Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        df_full.at[row_idx, "Action_Type"] = "Rescheduled Case"
                        
                        save_data(df_full, sheet_name=sheet_name)
                        st.session_state["is_reschedule_mode"] = False
                        st.session_state["selected_row_idx"] = None
                        st.session_state["open_dialog"] = None
                        st.success("✅ Case successfully rescheduled and saved!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error: {e}")

            if cancelled_res:
                st.session_state["is_reschedule_mode"] = False
                st.rerun()

    elif st.session_state["is_delete_request_mode"]:
        st.markdown("### 🗑️ Request Delete Case")
        with st.form(f"delete_reason_form_{row_idx}"):
            del_reason = st.text_area("✍️ Reason for deletion (Required):", placeholder="Type your reason here...")
            col_sub_req, col_back_req = st.columns(2)
            with col_sub_req:
                submitted_req = st.form_submit_button("🚀 Submit Delete Request", type="primary", use_container_width=True)
            with col_back_req:
                cancelled_req = st.form_submit_button("↩️ Back", use_container_width=True)

            if submitted_req:
                if del_reason.strip() == "":
                    st.warning("⚠️ Please provide a reason for deletion.")
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
        st.markdown("### ✏️ Edit Case & Booking Status")
        with st.form(f"popup_edit_form_{row_idx}"):
            updated_values = {}
            cols_to_show = list(df_full.columns) if role == "admin" else [c for c in df_full.columns if c not in [display_c, p_col, "Updated_By", "Updated_At", "Delete_Reason", "Appointment_PDF", "Appointment_Entry", "Booking_Status", "Preferred_Location", "Preferred_Dates", "Prior_Notice"]]
            
            f_col1, f_col2 = st.columns(2)
            for i, col_name in enumerate(cols_to_show):
                target_col = f_col1 if i % 2 == 0 else f_col2
                with target_col:
                    c_lower = col_name.lower()
                    if 'date' in c_lower and 'prefer' not in c_lower:
                        try:
                            parsed_d = datetime.datetime.strptime(str(row_data[col_name]).split()[0], "%Y-%m-%d").date()
                        except:
                            parsed_d = datetime.date.today()
                        sel_d = st.date_input(f"{col_name}", value=parsed_d, key=f"popup_date_{row_idx}_{col_name}")
                        updated_values[col_name] = sel_d.strftime("%Y-%m-%d")
                    else:
                        updated_values[col_name] = st.text_input(f"{col_name}", value=str(row_data[col_name]), key=f"popup_field_{row_idx}_{col_name}")

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            curr_bs = str(row_data.get("Booking_Status", "Active")).strip()
            opts = ["Active", "Booked"]
            b_idx = opts.index(curr_bs) if curr_bs in opts else 0
            booking_status_choice = st.selectbox("📌 Select Status", options=opts, index=b_idx, key=f"popup_bs_{row_idx}")

            app_entry_val = st.text_input("📝 Current Appointment Entry", value=str(row_data.get("Appointment_Entry", "")), key=f"popup_app_entry_{row_idx}")

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            uploaded_pdf_file = st.file_uploader("📎 Upload Appointment PDF (Required if Booked)", type=["pdf"], key=f"popup_pdf_up_{row_idx}")

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            col_save, col_back = st.columns(2)
            with col_save:
                submitted = st.form_submit_button("💾 Save & Update", type="primary", use_container_width=True)
            with col_back:
                cancelled = st.form_submit_button("↩️ Cancel", use_container_width=True)

            if submitted:
                try:
                    if booking_status_choice == "Booked" and not uploaded_pdf_file and not str(row_data.get("Appointment_PDF", "")).strip():
                        st.warning("⚠️ Please upload the Appointment PDF when selecting Booked status.")
                    else:
                        current_time = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        changed = False
                        
                        for col_name, new_val in updated_values.items():
                            old_val = str(df_full.at[row_idx, col_name]).strip()
                            if old_val != str(new_val).strip():
                                changed = True
                            df_full.at[row_idx, col_name] = str(new_val).strip()

                        if str(df_full.at[row_idx, "Booking_Status"]) != booking_status_choice:
                            df_full.at[row_idx, "Booking_Status"] = booking_status_choice
                            changed = True

                        if str(df_full.at[row_idx, "Appointment_Entry"]) != app_entry_val.strip():
                            df_full.at[row_idx, "Appointment_Entry"] = app_entry_val.strip()
                            changed = True

                        if uploaded_pdf_file is not None:
                            safe_uname = str(row_data.get(display_c, "case")).strip().replace(" ", "_")
                            pdf_filename = f"Appointment_{safe_uname}_{row_idx}_{int(datetime.datetime.now().timestamp())}.pdf"
                            full_pdf_path = os.path.join(PDF_DIR, pdf_filename)
                            with open(full_pdf_path, "wb") as pf:
                                pf.write(uploaded_pdf_file.getbuffer())
                            df_full.at[row_idx, "Appointment_PDF"] = full_pdf_path
                            changed = True

                        if changed:
                            df_full.at[row_idx, "Updated_By"] = current_user
                            df_full.at[row_idx, "Updated_At"] = current_time
                            df_full.at[row_idx, "Action_Type"] = "Updated Case"

                        target_save_sheet = sheet_name
                        if booking_status_choice == "Booked":
                            sheets_all = get_sheet_names()
                            booked_sheet_name = next((s for s in sheets_all if "booked" in s.lower()), None)
                            if booked_sheet_name and booked_sheet_name != sheet_name:
                                target_save_sheet = booked_sheet_name
                                df_target_booked = load_data(sheet_name=booked_sheet_name)
                                row_to_transfer = df_full.loc[row_idx].copy()
                                
                                for col in df_target_booked.columns:
                                    if col not in row_to_transfer:
                                        row_to_transfer[col] = ""
                                
                                df_target_booked = pd.concat([df_target_booked, pd.DataFrame([row_to_transfer])], ignore_index=True)
                                save_data(df_target_booked, sheet_name=booked_sheet_name)
                                df_full = df_full.drop(row_idx)

                        save_data(df_full, sheet_name=sheet_name)
                        # Clear popup states completely to prevent re-opening loop
                        st.session_state["is_editing_popup"] = False
                        st.session_state["selected_row_idx"] = None
                        st.session_state["open_dialog"] = None
                        st.success("✅ Case successfully updated and saved!")
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ Error updating case: {e}")
            
            if cancelled:
                st.session_state["is_editing_popup"] = False
                st.session_state["selected_row_idx"] = None
                st.session_state["open_dialog"] = None
                st.rerun()

# 🌟 ALL ADMIN & WORKER DIALOGS & RENDERERS
@st.dialog("📥 Appointment PDFs Download Center", width="large")
def admin_pdf_download_dialog():
    st.info("💡 Download all uploaded Appointment PDFs across all sheets:")
    sheets = get_sheet_names()
    pdf_records = []
    for s in sheets:
        df_s = load_data(sheet_name=s)
        if df_s is not None and not df_s.empty and "Appointment_PDF" in df_s.columns:
            disp_c = find_display_username_column(df_s)
            login_c = find_login_user_column(df_s)
            for idx, r in df_s.iterrows():
                p_path = str(r.get("Appointment_PDF", "")).strip()
                if p_path and os.path.exists(p_path):
                    pdf_records.append({
                        "Sheet": s, "Row": idx + 1, "Username": str(r.get(disp_c, "N/A")),
                        "Worker": str(r.get(login_c, "N/A")), "Path": p_path, "Filename": os.path.basename(p_path)
                    })
    if not pdf_records:
        st.warning("No Appointment PDFs uploaded yet.")
    else:
        for item in pdf_records:
            with st.container(border=True):
                c_info, c_dl = st.columns([3, 1])
                with c_info:
                    st.write(f"📁 **Sheet:** {item['Sheet']} &nbsp;|&nbsp; 👤 **User:** {item['Username']}")
                    st.caption(f"File: {item['Filename']}")
                with c_dl:
                    with open(item['Path'], "rb") as f:
                        st.download_button("📥 Download", data=f.read(), file_name=item['Filename'], mime="application/pdf", key=f"dl_pdf_{item['Sheet']}_{item['Row']}", use_container_width=True)

def render_clickable_table(df, display_c, p_col, table_key="main", page_size=100):
    if df.empty:
        st.info("No cases found.")
        return

    df = reorder_dataframe_columns(df)
    total_rows = len(df)
    total_pages = max(1, math.ceil(total_rows / page_size))
    page_state_key = f"page_num_{table_key}"
    
    if page_state_key not in st.session_state:
        st.session_state[page_state_key] = 1

    curr_page = st.session_state[page_state_key]
    start_idx = (curr_page - 1) * page_size
    end_idx = min(start_idx + page_size, total_rows)
    page_df = df.iloc[start_idx:end_idx]

    if total_pages > 1:
        p_c1, p_c2, p_c3, p_c4 = st.columns([1.2, 2.6, 1.2, 1.2])
        with p_c1:
            if st.button("⬅️ Previous", key=f"prev_{table_key}", disabled=(curr_page <= 1), use_container_width=True):
                st.session_state[page_state_key] = curr_page - 1
                st.rerun()
        with p_c2:
            st.markdown(f"<div style='text-align:center; padding-top:10px; font-weight:700;'>Page {curr_page} of {total_pages}</div>", unsafe_allow_html=True)
        with p_c3:
            if st.button("Next ➡️", key=f"next_{table_key}", disabled=(curr_page >= total_pages), use_container_width=True):
                st.session_state[page_state_key] = curr_page + 1
                st.rerun()

    display_cols = [c for c in page_df.columns if c not in [p_col, "Appointment_PDF"]]
    curr_logged = qp.get("logged_in", "true" if st.session_state.get("logged_in") else "false")
    curr_user = qp.get("username", st.session_state.get("username", ""))
    curr_role = qp.get("role", st.session_state.get("role", ""))
    curr_sheet = qp.get("active_sheet", st.session_state.get("active_sheet", ""))

    html_parts = [
        '<div style="max-height: 500px; overflow-y: auto; overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 12px; margin-bottom: 16px;">',
        '<table style="width:100%; border-collapse: collapse; font-size: 13px; font-family: sans-serif;">',
        '<thead><tr style="background: linear-gradient(135deg, #1e293b 0%, #334155 100%); color: white; position: sticky; top: 0; z-index: 1;">',
        '<th style="padding: 12px 10px; border: 1px solid #334155; text-align: center;">#</th>'
    ]
    for col in display_cols:
        html_parts.append(f'<th style="padding: 12px 10px; border: 1px solid #334155; text-align: left; white-space: nowrap;">{col}</th>')
    html_parts.append('<th style="padding: 12px 10px; border: 1px solid #334155; text-align: left; white-space: nowrap;">Status</th>')
    html_parts.append("</tr></thead><tbody>")

    base_link = f"?logged_in={curr_logged}&username={curr_user}&role={curr_role}&active_sheet={curr_sheet}&edit_row="
    indices = page_df.index.tolist()
    act_col_vals = page_df["Action_Type"].astype(str).str.strip().str.lower().tolist() if "Action_Type" in page_df.columns else [""] * len(page_df)
    col_data = {col: page_df[col].astype(str).tolist() for col in display_cols}

    for i, idx in enumerate(indices):
        act_t = act_col_vals[i]
        row_bg = "#fee2e2" if act_t == "delete rejected" else ("#f8fafc" if idx % 2 == 0 else "#ffffff")
        html_parts.append(f'<tr style="background-color: {row_bg};">')
        html_parts.append(f'<td style="padding: 10px; border: 1px solid #e2e8f0; text-align: center; color: #64748b; font-weight: 600;">{idx + 1}</td>')
        
        for col in display_cols:
            val = col_data[col][i]
            if col == display_c:
                display_val = val.strip() if val.strip() != "" else "Not Assigned"
                cell_content = f'<a href="{base_link}{idx}" target="_self" class="username-hyperlink">{display_val}</a>'
            else:
                cell_content = val if val != "" else "-"
            html_parts.append(f'<td style="padding: 10px; border: 1px solid #e2e8f0; color: #1e293b; white-space: nowrap;">{cell_content}</td>')
        
        status_display = "Rejected" if act_t == "delete rejected" else "-"
        html_parts.append(f'<td style="padding: 10px; border: 1px solid #e2e8f0; color: #b91c1c; font-weight: 700; white-space: nowrap;">{status_display}</td>')
        html_parts.append("</tr>")

    html_parts.append("</tbody></table></div>")
    st.markdown("".join(html_parts), unsafe_allow_html=True)

@st.dialog("📢 Add / Update Notice (Scrolling Announcement)", width="medium")
def admin_notice_dialog():
    current_notice = get_admin_notice()
    with st.form("admin_notice_form"):
        notice_input = st.text_area("Notice Text:", value=current_notice)
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
    all_workers = get_all_workers_list()
    login_records = load_login_tracking()
    normalized_login_records = {k.lower(): v for k, v in login_records.items()}
    
    logged_in_list = []
    not_logged_in_list = []
    
    for w in all_workers:
        rec = normalized_login_records.get(w.lower(), {})
        status = rec.get("status", "Not Logged In")
        last_time = rec.get("last_login", "-")
        item = {"worker": w, "status": status, "time": last_time}
        if status == "Logged In":
            logged_in_list.append(item)
        else:
            not_logged_in_list.append(item)
            
    tab_logged, tab_not_logged = st.tabs([f"🟢 Logged In ({len(logged_in_list)})", f"🔴 Not Logged In ({len(not_logged_in_list)})"])
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
        st.warning("No attendance records found.")
    else:
        df_att = pd.DataFrame(attendance_data)
        df_att.index = range(1, len(df_att) + 1)
        st.dataframe(df_att, use_container_width=True)
        att_buffer = io.BytesIO()
        with pd.ExcelWriter(att_buffer, engine='openpyxl') as writer:
            df_att.to_excel(writer, sheet_name="Attendance", index=False)
        st.download_button("📥 Download Attendance Report (.xlsx)", data=att_buffer.getvalue(), file_name="Attendance.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

@st.dialog("📊 Case Worker & Visa Type Booked Cases Report", width="large")
def admin_worker_wise_report_dialog():
    sheets = get_sheet_names()
    target_sheet = next((s for s in sheets if "booked" in s.lower()), sheets[0] if sheets else "Sheet1")
    today_str = get_ist_now().strftime("%Y-%m-%d")
    
    today_worker_stats = {}
    total_worker_stats = {}
    visa_type_stats = {}
    
    df_s = load_data(sheet_name=target_sheet)
    if df_s is not None and not df_s.empty:
        login_c = find_login_user_column(df_s)
        visa_c = find_visa_type_column(df_s)
        date_cols = [c for c in df_s.columns if any(k in c.lower() for k in ['date', 'time', 'day', 'created'])]
        
        for _, r in df_s.iterrows():
            w_name = str(r.get(login_c, "Not Assigned")).strip()
            if w_name == "" or w_name.lower() == "nan":
                w_name = "Not Assigned"
            
            v_type = str(r.get(visa_c, "Unspecified")).strip().upper() if visa_c else "Unspecified"
            if v_type == "" or v_type == "NAN":
                v_type = "Unspecified"
            
            total_worker_stats.setdefault(w_name, {"Total Cases": 0, "Visa_Breakdown": {}})
            total_worker_stats[w_name]["Total Cases"] += 1
            total_worker_stats[w_name]["Visa_Breakdown"][v_type] = total_worker_stats[w_name]["Visa_Breakdown"].get(v_type, 0) + 1
            visa_type_stats[v_type] = visa_type_stats.get(v_type, 0) + 1
            
            is_today = any(today_str in str(r.get(dc, "")) for dc in date_cols) or (today_str in str(r.get("Updated_At", "")))
            if is_today:
                today_worker_stats.setdefault(w_name, {"Today Cases": 0, "Visa_Breakdown": {}})
                today_worker_stats[w_name]["Today Cases"] += 1

    tab_today, tab_total, tab_visa = st.tabs(["📅 Today's Booked Cases", "📈 Total Booked Cases", "✈️ Visa Type Summary"])
    with tab_today:
        today_list = [{"Case Worker": w, "Today Booked Cases": d.get("Today Cases", 0), "Active Sheet": target_sheet} for w, d in today_worker_stats.items() if d.get("Today Cases", 0) > 0]
        df_today = pd.DataFrame(today_list)
        if not df_today.empty:
            df_today.index = range(1, len(df_today) + 1)
            st.dataframe(df_today, use_container_width=True)
        else:
            st.info("No booked cases today.")
    with tab_total:
        total_list = [{"Case Worker": w, "Total Booked Cases": d["Total Cases"], "Active Sheet": target_sheet} for w, d in total_worker_stats.items()]
        df_total = pd.DataFrame(total_list)
        if not df_total.empty:
            df_total.index = range(1, len(df_total) + 1)
            st.dataframe(df_total, use_container_width=True)
    with tab_visa:
        visa_list = [{"Visa Type": v, "Total Booked Cases": c} for v, c in visa_type_stats.items()]
        df_visa = pd.DataFrame(visa_list)
        if not df_visa.empty:
            df_visa.index = range(1, len(df_visa) + 1)
            st.dataframe(df_visa, use_container_width=True)

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
                
    tab_new, tab_upd = st.tabs(["➕ New Cases", "✏️ Updated Cases"])
    with tab_new:
        if all_rows:
            comb = pd.concat(all_rows, ignore_index=True)
            pc = find_password_column(comb)
            new_df = comb[comb["Action_Type"].astype(str).str.strip().str.lower() == "new case"]
            render_clickable_table(new_df, find_display_username_column(new_df), pc, table_key="dlg_new")
    with tab_upd:
        if all_rows:
            comb = pd.concat(all_rows, ignore_index=True)
            pc = find_password_column(comb)
            upd_df = comb[comb["Action_Type"].astype(str).str.strip().str.lower() == "updated case"]
            render_clickable_table(upd_df, find_display_username_column(upd_df), pc, table_key="dlg_upd")

@st.dialog("🗑️ Pending Delete Requests", width="large")
def admin_delete_requests_dialog():
    sheets = get_sheet_names()
    del_reqs = []
    for s in sheets:
        df_s = load_data(sheet_name=s)
        if df_s is not None and not df_s.empty and "Action_Type" in df_s.columns:
            for idx, r in df_s.iterrows():
                if str(r.get("Action_Type", "")).strip().lower() == "delete request":
                    rd = r.to_dict()
                    rd["Sheet_Name"] = s
                    rd["Original_Index"] = idx
                    del_reqs.append(rd)
    
    if not del_reqs:
        st.info("No pending delete requests.")
    else:
        for i, rd in enumerate(del_reqs):
            s_name = rd["Sheet_Name"]
            orig_idx = rd["Original_Index"]
            with st.container(border=True):
                st.write(f"Sheet: {s_name} | Reason: {rd.get('Delete_Reason', '')}")
                c_acc, c_rej = st.columns(2)
                with c_acc:
                    if st.button("✅ Accept", key=f"acc_{s_name}_{orig_idx}_{i}", type="primary"):
                        dt = load_data(sheet_name=s_name)
                        dt = dt.drop(orig_idx)
                        save_data(dt, sheet_name=s_name)
                        st.success("Deleted!")
                        st.rerun()
                with c_rej:
                    if st.button("❌ Reject", key=f"rej_{s_name}_{orig_idx}_{i}"):
                        dt = load_data(sheet_name=s_name)
                        dt.at[orig_idx, "Action_Type"] = "Delete Rejected"
                        save_data(dt, sheet_name=s_name)
                        st.success("Rejected!")
                        st.rerun()

@st.dialog("👥 Users, Passwords & Permissions", width="large")
def permissions_dialog():
    all_blocked_now = are_all_workers_blocked()
    if all_blocked_now:
        if st.button("🔓 Activate All Workers", use_container_width=True):
            toggle_all_workers_block(False)
            st.rerun()
    else:
        if st.button("🔒 Block All Workers", use_container_width=True):
            toggle_all_workers_block(True)
            st.rerun()

@st.dialog("🔑 Change Admin Password", width="medium")
def admin_password_dialog():
    with st.form("dlg_admin_pass_form"):
        curr_p = st.text_input("Current Password", type="password")
        new_p = st.text_input("New Password", type="password")
        if st.form_submit_button("Update Password", type="primary", use_container_width=True):
            admin_u, admin_p = get_admin_creds()
            if curr_p == admin_p and new_p.strip():
                save_admin_creds(admin_u, new_p.strip())
                st.success("Updated successfully!")

@st.dialog("📤 Sync Local Excel File to Cloud", width="medium")
def sync_excel_dialog():
    up_ex = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
    if up_ex and st.button("Sync", type="primary", use_container_width=True):
        with open(EXCEL_FILE, "wb") as f:
            f.write(up_ex.getbuffer())
        st.cache_data.clear()
        st.success("Synced!")
        st.rerun()

@st.dialog("🔑 Change Password", width="medium")
def worker_password_dialog(worker_name):
    with st.form("dlg_worker_pass_form"):
        curr_wp = st.text_input("Current Password", type="password")
        new_wp = st.text_input("New Password", type="password")
        if st.form_submit_button("Update Password", type="primary", use_container_width=True):
            st.success("Password updated!")

@st.dialog("➕ Add New Case", width="large")
def add_case_dialog(worker_name):
    sheets = get_sheet_names()
    target_sheet = st.selectbox("Select Target Sheet", options=sheets)
    target_df = load_data(sheet_name=target_sheet)
    t_login_col = find_login_user_column(target_df)
    t_pass_col = find_password_column(target_df)
    
    with st.form("dlg_new_case_form", clear_on_submit=True):
        form_data = {}
        cols = [c for c in target_df.columns if c not in [t_login_col, t_pass_col, "Updated_By", "Updated_At", "Action_Type", "Delete_Reason", "Booking_Status", "Appointment_Entry", "Appointment_PDF", "Preferred_Location", "Preferred_Dates", "Prior_Notice"]]
        for col_name in cols:
            form_data[col_name] = st.text_input(f"{col_name}")
        
        if st.form_submit_button("Submit", type="primary", use_container_width=True):
            new_row = {col: "" for col in target_df.columns}
            new_row[t_login_col] = worker_name
            new_row["Updated_By"] = worker_name
            new_row["Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
            new_row["Action_Type"] = "New Case"
            for k, v in form_data.items():
                new_row[k] = v
            updated_target_df = pd.concat([target_df, pd.DataFrame([new_row])], ignore_index=True)
            save_data(updated_target_df, sheet_name=target_sheet)
            st.session_state["open_dialog"] = None
            st.success("Added successfully!")
            st.rerun()

@st.dialog("🔍 Search Main Data", width="large")
def search_data_dialog():
    sheets = get_sheet_names()
    selected_search_sheet = st.selectbox("Select Sheet", options=sheets)
    df_search = load_data(sheet_name=selected_search_sheet)
    search_query = st.text_input("Search query:").strip()
    if search_query:
        mask = df_search.apply(lambda row: row.astype(str).str.contains(search_query, case=False, na=False).any(), axis=1)
        st.dataframe(df_search[mask], use_container_width=True)

# Dialog Controller
if st.session_state["open_dialog"] == "permissions":
    permissions_dialog()
elif st.session_state["open_dialog"] == "admin_password":
    admin_password_dialog()
elif st.session_state["open_dialog"] == "sync_excel":
    sync_excel_dialog()
elif st.session_state["open_dialog"] == "admin_updated_cases":
    admin_updated_cases_dialog()
elif st.session_state["open_dialog"] == "admin_delete_requests":
    admin_delete_requests_dialog()
elif st.session_state["open_dialog"] == "admin_notice":
    admin_notice_dialog()
elif st.session_state["open_dialog"] == "admin_login_status":
    admin_login_status_dialog()
elif st.session_state["open_dialog"] == "admin_attendance":
    admin_attendance_dialog()
elif st.session_state["open_dialog"] == "admin_worker_report":
    admin_worker_wise_report_dialog()
elif st.session_state["open_dialog"] == "admin_pdf_downloads":
    admin_pdf_download_dialog()
elif st.session_state["open_dialog"] == "worker_password":
    worker_password_dialog(st.session_state.get("username", ""))
elif st.session_state["open_dialog"] == "add_case":
    add_case_dialog(st.session_state.get("username", ""))
elif st.session_state["open_dialog"] == "search_data":
    search_data_dialog()
elif st.session_state["open_dialog"] == "edit_case_popup":
    r_idx = st.session_state.get("selected_row_idx")
    if r_idx is not None:
        curr_role = st.session_state.get("role", "worker")
        curr_u = st.session_state.get("username", "")
        perms = load_permissions()
        c_upd = True if curr_role == "admin" else perms.get(curr_u, {}).get("can_update", True)
        view_and_edit_case_dialog(st.session_state["active_sheet"], r_idx, curr_role, curr_u, c_upd)

if not st.session_state["logged_in"]:
    default_nav_idx = 5 if st.session_state.get("force_logged_out_msg") else 0
    selected_page = st.radio(
        "Navigation",
        ["🏠 Home", "ℹ️ About Us", "✈️ Services", "⭐ Success Stories", "📞 Contact Us", "🔐 Portal Login"],
        index=default_nav_idx,
        horizontal=True,
        label_visibility="collapsed"
    )
    
    if selected_page == "🔐 Portal Login":
        col_l1, col_l2, col_l3 = st.columns([1, 1.4, 1])
        with col_l2:
            with st.container(border=True):
                st.markdown("<h3 style='text-align:center;'>🔐 Portal Login</h3>", unsafe_allow_html=True)
                sheets = get_sheet_names()
                with st.form("login_form"):
                    username_input = st.text_input("👤 Username")
                    password_input = st.text_input("🔑 Password", type="password")
                    submit_btn = st.form_submit_button("🚀 Login", type="primary", use_container_width=True)
                    
                    if submit_btn:
                        u_in = username_input.strip()
                        p_in = password_input.strip()
                        admin_u, admin_p = get_admin_creds()
                        
                        if u_in.lower() == admin_u.lower() and p_in == admin_p:
                            st.session_state["logged_in"] = True
                            st.session_state["username"] = admin_u
                            st.session_state["role"] = "admin"
                            st.rerun()
                        else:
                            logged_in_success = False
                            matched_sheet = ""
                            actual_matched_username = ""
                            for s in sheets:
                                df_check = load_data(sheet_name=s)
                                if df_check is not None and not df_check.empty:
                                    login_c = find_login_user_column(df_check)
                                    p_col = find_password_column(df_check)
                                    for _, r in df_check.iterrows():
                                        if str(r[login_c]).strip().lower() == u_in.lower() and str(r[p_col]).strip() == p_in:
                                            logged_in_success = True
                                            matched_sheet = s
                                            actual_matched_username = str(r[login_c]).strip()
                                            break
                                if logged_in_success:
                                    break
                            
                            if logged_in_success:
                                record_worker_login(actual_matched_username)
                                st.session_state["logged_in"] = True
                                st.session_state["username"] = actual_matched_username
                                st.session_state["role"] = "worker"
                                st.session_state["active_sheet"] = matched_sheet
                                st.session_state["attendance_recorded"] = True
                                st.rerun()
                            else:
                                st.error("❌ Incorrect Username or Password!")

elif st.session_state["role"] == "admin":
    sheets = get_sheet_names()
    current_sheet = st.session_state["active_sheet"]
    if current_sheet not in sheets:
        current_sheet = sheets[0]
        st.session_state["active_sheet"] = current_sheet

    df = load_data(sheet_name=current_sheet)
    login_c = find_login_user_column(df)
    display_c = find_display_username_column(df)
    p_col = find_password_column(df)
    
    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    top_col1, top_col2, top_col3 = st.columns([6.5, 1.2, 1.2], gap="small")
    with top_col1:
        st.markdown(f"<div class='section-banner' style='margin:0;'>👑 Admin Panel — {current_sheet}</div>", unsafe_allow_html=True)
    with top_col2:
        if st.button("🔄 Refresh", key="admin_refresh"):
            st.cache_data.clear()
            st.rerun()
    with top_col3:
        if st.button("🚪 Logout", key="admin_logout"):
            st.session_state["logged_in"] = False
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    admin_dash_tab = st.radio("Admin Nav", ["📊 Dashboard", "📂 Records"], horizontal=True, label_visibility="collapsed")
    
    if admin_dash_tab == "📊 Dashboard":
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.markdown(f"<div class='kpi-card'><div class='kpi-title'>Total Cases</div><div class='kpi-value'>{len(df)}</div></div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='kpi-card green'><div class='kpi-title'>Updated</div><div class='kpi-value'>{len(df[df['Updated_By'] != ''])}</div></div>", unsafe_allow_html=True)
        
        st.markdown("<div class='action-group-title'>Quick Tools</div>", unsafe_allow_html=True)
        b1, b2, b3, b4 = st.columns(4)
        with b1:
            if st.button("New / Updated"): st.session_state["open_dialog"] = "admin_updated_cases"; st.rerun()
        with b2:
            if st.button("Worker Report"): st.session_state["open_dialog"] = "admin_worker_report"; st.rerun()
        with b3:
            if st.button("Permissions"): st.session_state["open_dialog"] = "permissions"; st.rerun()
        with b4:
            if st.button("Sync Excel"): st.session_state["open_dialog"] = "sync_excel"; st.rerun()
            
    else:
        render_clickable_table(df, display_c, p_col, table_key="admin_main")

else:
    sheets = get_sheet_names()
    current_sheet = st.session_state["active_sheet"]
    if current_sheet not in sheets:
        current_sheet = sheets[0]
        st.session_state["active_sheet"] = current_sheet

    df = load_data(sheet_name=current_sheet)
    current_user = st.session_state["username"].strip()
    
    login_c = find_login_user_column(df)
    display_c = find_display_username_column(df)
    p_col = find_password_column(df)
    
    user_rows = df[df[login_c].astype(str).str.strip().str.lower() == current_user.lower()]

    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    tc1, tc2, tc3 = st.columns([6.5, 1.2, 1.2])
    with tc1: st.markdown(f"<div class='section-banner' style='margin:0;'>👋 {current_user} Workspace</div>", unsafe_allow_html=True)
    with tc2: 
        if st.button("🔄 Refresh"): st.cache_data.clear(); st.rerun()
    with tc3: 
        if st.button("🚪 Logout"): record_worker_logout(current_user); st.session_state["logged_in"] = False; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    w1, w2, w3 = st.columns(3)
    with w1:
        if st.button("➕ Add Case"): st.session_state["open_dialog"] = "add_case"; st.rerun()
    with w2:
        if st.button("🔍 Search"): st.session_state["open_dialog"] = "search_data"; st.rerun()
    with w3:
        if st.button("🔑 Change Password"): st.session_state["open_dialog"] = "worker_password"; st.rerun()

    st.divider()
    render_clickable_table(user_rows, display_c, p_col, table_key=f"worker_{current_user}")
