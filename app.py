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
PDF_BASE_DIR = "uploaded_appointment_pdfs"
LOGO_FILES = ["pngTShirtPrintCricket.jpg", "logo.jpg", "logo.png", "logo.jpeg"]

if not os.path.exists(PDF_BASE_DIR):
    os.makedirs(PDF_BASE_DIR)

st.set_page_config(
    page_title="HemanVisa Portal", 
    page_icon="✈️", 
    layout="wide", 
    initial_sidebar_state="expanded"  # <-- ఇది సైడ్‌బార్ ఎప్పుడూ ఓపెన్‌గా ఉండేలా చేస్తుంది
)

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

def reorder_dataframe_columns(df):
    if df is None or df.empty:
        return df
    cols_to_drop = [c for c in df.columns if c.strip().lower() in ["ds-160 case worker", "ds160 case worker"]]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop, errors="ignore")

    cols = list(df.columns)
    desired_order_prefixes = ["case", "username", "password", "no of app", "lead"]
    found_priority = []
    
    for prefix in desired_order_prefixes:
        for c in cols:
            c_low = c.strip().lower().replace("_", " ")
            if (c_low == prefix or prefix in c_low) and c not in found_priority:
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
                
                if "Appointment_Entry" not in df.columns and "Current Appointment" in df.columns:
                    df = df.rename(columns={"Current Appointment": "Appointment_Entry"})
                elif "Appointment_Entry" not in df.columns:
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
                    u_rows_s = reorder_dataframe_columns(u_rows_s)
                    u_rows_s.to_excel(writer, sheet_name=s, index=False)
        if not has_user_data and current_sheet in all_sheets:
            df_cur = all_sheets[current_sheet]
            login_cs = find_login_user_column(df_cur)
            u_rows_cur = df_cur[df_cur[login_cs].astype(str).str.strip().str.lower() == current_user.lower()]
            u_rows_cur = reorder_dataframe_columns(u_rows_cur)
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
    
    df_to_save = reorder_dataframe_columns(df_to_save)
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
        if col.strip().lower() in ["username", "case worker", "cases_worker", "user"]:
            if col.strip() != "USERNAME":
                return col
    for col in df.columns:
        if col.strip().lower() in ["case worker", "cases_worker", "user"]:
            return col
    return df.columns[0]

def find_display_username_column(df):
    for col in df.columns:
        if col.strip() == "USERNAME":
            return col
    for col in df.columns:
        if col.strip().lower() == "username" and col.strip() != find_login_user_column(df):
            return col
    return df.columns[0]

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

# --- STYLING SETUP ---
st.markdown("""
<style>
    [data-testid="stHeader"] {
        display: none !important;
    }
    [data-testid="collapsedControl"] {
        display: none !important;
    }
    .sticky-top-banner {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        z-index: 99999;
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
        margin-left: 50px;
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
        margin-left: 20px;
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
        transition: all 0.25s ease;
    }
    .banner-badge:hover {
        background: rgba(255, 255, 255, 0.18);
        border-color: #d4af37;
        transform: translateY(-1px);
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
        backdrop-filter: blur(8px);
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
        letter-spacing: 0.5px;
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
        transition: all 0.2s ease;
    }
    .corporate-card:hover {
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.1);
        transform: translateY(-2px);
    }
    .corporate-card h3 {
        color: #172554;
        font-size: 18px;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .corporate-card p {
        color: #475569;
        font-size: 14px;
        line-height: 1.6;
        margin: 0;
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
        overflow: hidden;
        white-space: nowrap;
        border-left: 5px solid #d4af37;
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
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    .kpi-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 12px 28px rgba(15, 23, 42, 0.25);
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
        line-height: 1.2;
        font-weight: 700;
        color: #f1f5f9;
    }
    .kpi-value {
        font-size: 28px;
        font-weight: 800;
        margin: 0;
        line-height: 1.1;
        letter-spacing: 0.5px;
        text-shadow: 0 2px 4px rgba(0,0,0,0.25);
    }

    div[data-testid="stElementContainer"]:has(.visa-color-0) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-0) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%) !important;
        border-bottom: 3px solid #60a5fa !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-1) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-1) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #065f46 0%, #059669 100%) !important;
        border-bottom: 3px solid #34d399 !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-2) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-2) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #92400e 0%, #d97706 100%) !important;
        border-bottom: 3px solid #fbbf24 !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-3) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-3) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #4338ca 0%, #6366f1 100%) !important;
        border-bottom: 3px solid #818cf8 !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-4) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-4) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #831843 0%, #db2777 100%) !important;
        border-bottom: 3px solid #f472b6 !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-5) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-5) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #115e59 0%, #0d9488 100%) !important;
        border-bottom: 3px solid #2dd4bf !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-6) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-6) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #334155 100%, #475569 0%) !important;
        border-bottom: 3px solid #94a3b8 !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-7) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-7) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #7f1d1d 0%, #b91c1c 100%) !important;
        border-bottom: 3px solid #f87171 !important;
        color: #ffffff !important;
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
        box-shadow: 0 3px 12px rgba(15, 23, 42, 0.04);
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
        box-shadow: 0 3px 10px rgba(15, 23, 42, 0.12);
        margin-top: 4px !important;
        margin-bottom: 8px !important;
        transition: all 0.2s ease;
        white-space: nowrap !important;
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%) !important;
    }
    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.2);
        opacity: 0.95;
    }
    .username-hyperlink {
        color: #1e293b !important;
        text-decoration: underline !important;
        font-weight: 700 !important;
        cursor: pointer !important;
    }
    .username-hyperlink:hover {
        color: #2563eb !important;
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
