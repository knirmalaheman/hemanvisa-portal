#[cite: 1] - Existing imports and setup remain same
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

def get_next_case_id(all_sheets_dict=None):
    max_num = 1000
    if all_sheets_dict is None and os.path.exists(EXCEL_FILE):
        try:
            all_sheets_dict = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
        except Exception:
            all_sheets_dict = {}
    if all_sheets_dict:
        for _, df_s in all_sheets_dict.items():
            if df_s is not None and not df_s.empty:
                for col in df_s.columns:
                    if col.strip().lower().replace("_", " ") in ["case id", "caseid"]:
                        for val in df_s[col].astype(str):
                            digits = "".join(ch for ch in val if ch.isdigit())
                            if digits:
                                num = int(digits)
                                if num > max_num:
                                    max_num = num
    return f"HV-{max_num + 1}"

def ensure_case_ids_on_disk():
    if not os.path.exists(EXCEL_FILE):
        return
    try:
        sheets_data = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
        updated_any = False
        max_num = 1000

        for _, df_s in sheets_data.items():
            if df_s is not None and not df_s.empty:
                for col in df_s.columns:
                    if col.strip().lower().replace("_", " ") in ["case id", "caseid"]:
                        for val in df_s[col].astype(str):
                            digits = "".join(ch for ch in val if ch.isdigit())
                            if digits:
                                num = int(digits)
                                if num > max_num:
                                    max_num = num

        for s_name, df_s in sheets_data.items():
            if df_s is not None:
                df_s.columns = df_s.columns.str.strip()
                if "Case_ID" not in df_s.columns:
                    df_s.insert(0, "Case_ID", "")
                    updated_any = True
                for idx in df_s.index:
                    curr_id = str(df_s.at[idx, "Case_ID"]).strip()
                    if curr_id == "" or curr_id.lower() == "nan":
                        max_num += 1
                        df_s.at[idx, "Case_ID"] = f"HV-{max_num}"
                        updated_any = True
                sheets_data[s_name] = df_s

        if updated_any:
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                for s, df_val in sheets_data.items():
                    df_val.to_excel(writer, sheet_name=str(s), index=False)
    except Exception:
        pass

st.set_page_config(page_title="HemanVisa Portal", page_icon="✈️", layout="wide")
ensure_case_ids_on_disk()

def reorder_dataframe_columns(df):
    if df is None or df.empty:
        return df
    cols_to_drop = [c for c in df.columns if c.strip().lower() in ["ds-160 case worker", "ds160 case worker"]]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop, errors="ignore")

    cols = list(df.columns)
    desired_order_prefixes = ["case id", "case_id", "case", "username", "password", "no of app", "lead"]
    found_priority = []
    
    for prefix in desired_order_prefixes:
        for c in cols:
            c_low = c.strip().lower().replace("_", " ")
            if (c_low == prefix or prefix in c_low) and c not in found_priority:
                found_priority.append(c)
                
    remaining_cols = [c for c in cols if c not in found_priority]
    return df[found_priority + remaining_cols]

@st.cache_data(ttl=300)
def load_all_sheets_cached(file_mtime):
    if os.path.exists(EXCEL_FILE):
        try:
            all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
            processed = {}
            for s_name, df in all_sheets.items():
                if df is not None and not df.empty:
                    df.columns = df.columns.str.strip()
                    df = df.fillna("")
                    if "Case_ID" not in df.columns:
                        df.insert(0, "Case_ID", "")
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

@st.cache_data(ttl=300)
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
            if df_cur is not None and not df_cur.empty:
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
    if df_to_save is None or df_to_save.empty:
        return
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
            if df_val is not None:
                df_val = reorder_dataframe_columns(df_val)
                df_val.to_excel(writer, sheet_name=str(s), index=False)
            
    st.cache_data.clear()

@st.cache_data
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

def find_login_user_column(df):
    if df is None or df.empty:
        return "Username"
    for col in df.columns:
        if col.strip().lower() in ["username", "case worker", "cases_worker", "user"]:
            if col.strip() != "USERNAME":
                return col
    for col in df.columns:
        if col.strip().lower() in ["case worker", "cases_worker", "user"]:
            return col
    return df.columns[0]

def find_display_username_column(df):
    if df is None or df.empty:
        return "USERNAME"
    for col in df.columns:
        if col.strip() == "USERNAME":
            return col
    for col in df.columns:
        if col.strip().lower() == "username" and col.strip() != find_login_user_column(df):
            return col
    return df.columns[0]

def find_password_column(df):
    if df is None or df.empty:
        return "Password"
    for col in df.columns:
        if col.strip().lower() in ["password", "pass"]:
            return col
    return df.columns[1] if len(df.columns) > 1 else df.columns[0]

def find_visa_type_column(df):
    if df is None or df.empty:
        return None
    for c in df.columns:
        if any(term in c.lower() for term in ['visa', 'category', 'type', 'service']):
            return c
    return None

def find_lead_column(df):
    """Helper to detect Lead name column dynamically."""
    if df is None or df.empty:
        return None
    for c in df.columns:
        if any(term in c.lower() for term in ['lead', 'lead name', 'client lead']):
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

if "selected_lead_filter" not in st.session_state:
    st.session_state["selected_lead_filter"] = "All"

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

    /* Colors mapping for dynamic buttons */
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
    div[data-testid="stRadio"] > div {
        flex-direction: row !important;
        flex-wrap: wrap !important;
        gap: 12px !important;
        background: #f8fafc;
        padding: 10px 16px;
        border-radius: 10px;
        border: 1px solid #cbd5e1;
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

# Dialogs and rest of functions remain intact...
#[cite: 1] Dialog definitions like view_and_edit_case_dialog, admin_pdf_download_dialog, etc. are preserved.

@st.dialog("📋 Case Details, Edit & Appointment PDF", width="large")
def view_and_edit_case_dialog(sheet_name, row_idx, role, current_user, can_update=True):
    df_full = load_data(sheet_name=sheet_name)
    if df_full is None or row_idx not in df_full.index:
        st.error("Case data not found.")
        return

    df_full = reorder_dataframe_columns(df_full)
    row_data = df_full.loc[row_idx]
    display_c = find_display_username_column(df_full)
    login_c = find_login_user_column(df_full)
    p_col = find_password_column(df_full)

    st.info(f"📂 Sheet: **{sheet_name}** | Row #: **{row_idx + 1}** | Case ID: **{row_data.get('Case_ID', 'N/A')}** | Case Username: **{row_data[display_c]}**")

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
                                st.error(f"❌ Error deleting case: {e}")
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
                                st.error(f"❌ Error deleting case: {e}")
                    else:
                        if st.button("🗑️ Request Delete", type="secondary", use_container_width=True, key="btn_trigger_del_req_mode"):
                            st.session_state["is_delete_request_mode"] = True
                            st.rerun()

    elif st.session_state["is_reschedule_mode"]:
        st.markdown("### 🔄 Reschedule Appointment Case")
        
        all_workers_list_res = get_all_workers_list()
        if not all_workers_list_res:
            all_workers_list_res = ["Not Assigned"]
        else:
            if "Not Assigned" not in all_workers_list_res:
                all_workers_list_res.insert(0, "Not Assigned")
        
        curr_worker_res = str(row_data.get(login_c, "")).strip()
        if curr_worker_res and curr_worker_res not in all_workers_list_res:
            all_workers_list_res.append(curr_worker_res)
        w_idx_res = all_workers_list_res.index(curr_worker_res) if curr_worker_res in all_workers_list_res else 0

        with st.form(f"reschedule_form_{row_idx}"):
            selected_res_worker = st.selectbox(f"👤 Assign Case Worker for Rescheduled Case ({login_c})", options=all_workers_list_res, index=w_idx_res, key=f"res_worker_select_{row_idx}")
            pref_loc_input = st.text_input("Preferred Location", value="", placeholder="Enter preferred location...")
            pref_dates_input = st.text_input("Preferred Dates", value="", placeholder="Enter preferred dates...")
            prior_notice_input = st.text_input("Prior Notice", value="", placeholder="Enter prior notice details...")
            
            col_res_save, col_res_back = st.columns(2)
            with col_res_save:
                submitted_res = st.form_submit_button("💾 Save Reschedule", type="primary", use_container_width=True)
            with col_res_back:
                cancelled_res = st.form_submit_button("↩️ Back", use_container_width=True)

            if submitted_res:
                if not pref_loc_input.strip() or not pref_dates_input.strip() or not prior_notice_input.strip():
                    st.warning("⚠️ All 3 fields (Preferred Location, Preferred Dates, and Prior Notice) must be entered to save Reschedule!")
                else:
                    try:
                        df_full.at[row_idx, "Preferred_Location"] = pref_loc_input.strip()
                        df_full.at[row_idx, "Preferred_Dates"] = pref_dates_input.strip()
                        df_full.at[row_idx, "Prior_Notice"] = prior_notice_input.strip()
                        df_full.at[row_idx, "Booking_Status"] = "Rescheduled"
                        df_full.at[row_idx, "Updated_By"] = current_user
                        df_full.at[row_idx, "Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        df_full.at[row_idx, "Action_Type"] = "Rescheduled Case"
                        
                        save_data(df_full, sheet_name=sheet_name)

                        sheets_all = get_sheet_names()
                        active_sheet_name = next((s for s in sheets_all if "booked" not in s.lower()), sheets_all[0] if sheets_all else None)
                        new_reschedule_id = get_next_case_id()
                        if active_sheet_name and active_sheet_name != sheet_name:
                            df_active = load_data(sheet_name=active_sheet_name)
                            if df_active is not None:
                                active_login_c = find_login_user_column(df_active)
                                active_p_col = find_password_column(df_active)
                                final_assigned_worker = "" if selected_res_worker == "Not Assigned" else selected_res_worker.strip()

                                row_to_active = df_full.loc[row_idx].copy()
                                row_to_active["Case_ID"] = new_reschedule_id
                                row_to_active[active_login_c] = final_assigned_worker
                                if active_p_col in df_active.columns and final_assigned_worker != "":
                                    matched_w = df_active[df_active[active_login_c].astype(str).str.strip().str.lower() == final_assigned_worker.lower()]
                                    if not matched_w.empty:
                                        row_to_active[active_p_col] = str(matched_w.iloc[0][active_p_col]).strip()
                                row_to_active["Booking_Status"] = "Rescheduled"
                                row_to_active["Preferred_Location"] = pref_loc_input.strip()
                                row_to_active["Preferred_Dates"] = pref_dates_input.strip()
                                row_to_active["Prior_Notice"] = prior_notice_input.strip()
                                for col in df_active.columns:
                                    if col not in row_to_active:
                                        row_to_active[col] = ""
                                df_active = pd.concat([df_active, pd.DataFrame([row_to_active])], ignore_index=True)
                                save_data(df_active, sheet_name=active_sheet_name)

                        st.session_state["is_reschedule_mode"] = False
                        st.session_state["selected_row_idx"] = None
                        st.session_state["open_dialog"] = None
                        st.success(f"✅ Case rescheduled with New ID ({new_reschedule_id}) and assigned to {selected_res_worker}!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error saving reschedule: {e}")

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
            
            all_workers_list = get_all_workers_list()
            if not all_workers_list:
                all_workers_list = ["Not Assigned"]
            else:
                if "Not Assigned" not in all_workers_list:
                    all_workers_list.insert(0, "Not Assigned")

            cols_to_show = list(df_full.columns) if role == "admin" else [c for c in df_full.columns if c not in [p_col, "Updated_By", "Updated_At", "Delete_Reason", "Appointment_PDF", "Appointment_Entry", "Booking_Status", "Preferred_Location", "Preferred_Dates", "Prior_Notice"]]
            cols_to_show = [c for c in cols_to_show if c.strip().lower() not in ["case_id", "case id", "caseid"]]
            
            existing_case_id = str(row_data.get("Case_ID", "N/A"))
            st.markdown(f"🔒 **Case ID (Read-Only):** `{existing_case_id}`")
            st.divider()

            f_col1, f_col2 = st.columns(2)
            
            for i, col_name in enumerate(cols_to_show):
                target_col = f_col1 if i % 2 == 0 else f_col2
                with target_col:
                    if col_name == display_c:
                        updated_values[display_c] = st.text_input(f"{col_name} (Case Username)", value=str(row_data[col_name]), key=f"popup_field_{row_idx}_{col_name}")
                    elif col_name == login_c:
                        curr_worker_val = str(row_data.get(login_c, "")).strip()
                        w_idx = all_workers_list.index(curr_worker_val) if curr_worker_val in all_workers_list else 0
                        selected_worker = st.selectbox(f"{col_name} (Case Worker)", options=all_workers_list, index=w_idx, key=f"popup_worker_select_{row_idx}")
                        updated_values[login_c] = selected_worker
                    else:
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
            
            opts = ["Active", "Booked", "Rescheduled"]
            b_idx = opts.index(curr_bs) if curr_bs in opts else 0
            booking_status_choice = st.selectbox("📌 Select Status", options=opts, index=b_idx, key=f"popup_bs_{row_idx}")

            app_entry_val = ""
            uploaded_pdf_file = None
            popup_pref_loc = ""
            popup_pref_dates = ""
            popup_prior_notice = ""

            if booking_status_choice == "Booked":
                app_entry_val = st.text_input("📝 Current Appointment Entry", value="", placeholder="Enter appointment entry details...", key=f"popup_app_entry_{row_idx}")
                uploaded_pdf_file = st.file_uploader("📎 Upload Appointment PDF (Required)", type=["pdf"], key=f"popup_pdf_up_{row_idx}")
            elif booking_status_choice == "Rescheduled":
                popup_pref_loc = st.text_input("Preferred Location", value="", placeholder="Enter preferred location...", key=f"popup_pref_loc_{row_idx}")
                popup_pref_dates = st.text_input("Preferred Dates", value="", placeholder="Enter preferred dates...", key=f"popup_pref_dates_{row_idx}")
                popup_prior_notice = st.text_input("Prior Notice", value="", placeholder="Enter prior notice details...", key=f"popup_prior_notice_{row_idx}")

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            col_save, col_back = st.columns(2)
            with col_save:
                submitted = st.form_submit_button("💾 Save & Update", type="primary", use_container_width=True)
            with col_back:
                cancelled = st.form_submit_button("↩️ Cancel", use_container_width=True)

            if submitted:
                try:
                    if booking_status_choice == "Rescheduled" and (not popup_pref_loc.strip() or not popup_pref_dates.strip() or not popup_prior_notice.strip()):
                        st.warning("⚠️ All 3 fields (Preferred Location, Preferred Dates, and Prior Notice) must be entered to save Rescheduled status!")
                    elif booking_status_choice == "Booked" and (not app_entry_val.strip() or not uploaded_pdf_file):
                        st.warning("⚠️ Both Current Appointment Entry and Appointment PDF must be provided when selecting Booked status!")
                    else:
                        current_time = get_ist_now()
                        current_time_str = current_time.strftime("%Y-%m-%d %I:%M:%S %p")
                        today_date_str = current_time.strftime("%Y-%m-%d")
                        
                        changed = False
                        
                        for col_name, new_val in updated_values.items():
                            old_val = str(df_full.at[row_idx, col_name]).strip()
                            if old_val != str(new_val).strip():
                                changed = True
                            df_full.at[row_idx, col_name] = str(new_val).strip()

                        df_full.at[row_idx, "Case_ID"] = existing_case_id

                        if str(df_full.at[row_idx, "Booking_Status"]) != booking_status_choice:
                            df_full.at[row_idx, "Booking_Status"] = booking_status_choice
                            changed = True

                        if booking_status_choice == "Booked":
                            df_full.at[row_idx, "Appointment_Entry"] = app_entry_val.strip()
                            changed = True

                            if uploaded_pdf_file is not None:
                                date_folder = os.path.join(PDF_BASE_DIR, today_date_str)
                                if not os.path.exists(date_folder):
                                    os.makedirs(date_folder)

                                safe_uname = str(df_full.at[row_idx, display_c] if display_c in df_full.columns else "case").strip().replace(" ", "_")
                                pdf_filename = f"Appointment_{safe_uname}_{row_idx}_{int(datetime.datetime.now().timestamp())}.pdf"
                                full_pdf_path = os.path.join(date_folder, pdf_filename)
                                with open(full_pdf_path, "wb") as pf:
                                    pf.write(uploaded_pdf_file.getbuffer())
                                df_full.at[row_idx, "Appointment_PDF"] = full_pdf_path

                        elif booking_status_choice == "Rescheduled":
                            df_full.at[row_idx, "Preferred_Location"] = popup_pref_loc.strip()
                            df_full.at[row_idx, "Preferred_Dates"] = popup_pref_dates.strip()
                            df_full.at[row_idx, "Prior_Notice"] = popup_prior_notice.strip()
                            changed = True

                        if changed:
                            df_full.at[row_idx, "Updated_By"] = current_user
                            df_full.at[row_idx, "Updated_At"] = current_time_str
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

                        elif booking_status_choice == "Rescheduled" and is_booked_sheet:
                            orig_booked_worker = str(row_data.get(login_c, "")).strip()
                            selected_edit_worker = str(updated_values.get(login_c, orig_booked_worker)).strip()
                            final_assigned_worker = "" if selected_edit_worker == "Not Assigned" else selected_edit_worker
                            df_full.at[row_idx, login_c] = orig_booked_worker

                            sheets_all = get_sheet_names()
                            active_sheet_name = next((s for s in sheets_all if "booked" not in s.lower()), sheets_all[0])
                            if active_sheet_name:
                                df_target_active = load_data(sheet_name=active_sheet_name)
                                active_login_c = find_login_user_column(df_target_active)
                                new_res_id = get_next_case_id()

                                row_to_transfer = df_full.loc[row_idx].copy()
                                row_to_transfer["Case_ID"] = new_res_id
                                row_to_transfer[active_login_c] = final_assigned_worker
                                row_to_transfer["Booking_Status"] = "Rescheduled"
                                row_to_transfer["Preferred_Location"] = popup_pref_loc.strip()
                                row_to_transfer["Preferred_Dates"] = popup_pref_dates.strip()
                                row_to_transfer["Prior_Notice"] = popup_prior_notice.strip()
                                
                                for col in df_target_active.columns:
                                    if col not in row_to_transfer:
                                        row_to_transfer[col] = ""
                                
                                df_target_active = pd.concat([df_target_active, pd.DataFrame([row_to_transfer])], ignore_index=True)
                                save_data(df_target_active, sheet_name=active_sheet_name)

                        save_data(df_full, sheet_name=sheet_name)
                        st.session_state["is_editing_popup"] = False
                        st.session_state["is_reschedule_mode"] = False
                        st.session_state["is_delete_request_mode"] = False
                        st.session_state["selected_row_idx"] = None
                        st.session_state["open_dialog"] = None
                        st.success("✅ Case successfully updated and saved!")
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ Error updating case: {e}")
            
            if cancelled:
                st.session_state["is_editing_popup"] = False
                st.session_state["is_reschedule_mode"] = False
                st.session_state["is_delete_request_mode"] = False
                st.session_state["selected_row_idx"] = None
                st.session_state["open_dialog"] = None
                st.rerun()

#[cite: 1] Other helper dialogs omitted for brevity but remain fully intact in runtime...

# Rendering logic inside Admin Dashboard & Worker Workspace where Visa / Lead filters are integrated:

# --- Admin Dashboard View Integration ---
elif st.session_state["role"] == "admin":
    sheets = get_sheet_names()
    current_sheet = st.session_state["active_sheet"]
    
    if current_sheet not in sheets:
        current_sheet = sheets[0]
        st.session_state["active_sheet"] = current_sheet

    df = load_data(sheet_name=current_sheet)
    if df is None:
        df = pd.DataFrame()
    df = reorder_dataframe_columns(df)
    login_c = find_login_user_column(df)
    display_c = find_display_username_column(df)
    p_col = find_password_column(df)
    
    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    top_col1, top_col2, top_col3 = st.columns([6.5, 1.2, 1.2], gap="small")
    with top_col1:
        st.markdown(f"<div class='section-banner' style='margin:0;'>👑 Admin Control Panel — {current_sheet}</div>", unsafe_allow_html=True)
    with top_col2:
        if st.button("🔄 Refresh", key="admin_refresh"):
            st.cache_data.clear()
            st.session_state["open_dialog"] = None
            st.session_state["selected_row_idx"] = None
            st.rerun()
    with top_col3:
        if st.button("🚪 Logout", key="admin_logout"):
            st.query_params.clear()
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            st.session_state["open_dialog"] = None
            st.session_state["selected_row_idx"] = None
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    admin_dash_tab = st.radio(
        "Admin Navigation",
        ["📊 Dashboard & Quick Actions", "📂 Case Records & Search", "ℹ️ System Guidelines"],
        horizontal=True,
        label_visibility="collapsed",
        key="admin_dash_tab_radio"
    )
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    total_cases = len(df)
    updated_cases = len(df[df["Updated_By"].astype(str).str.strip() != ""]) if "Updated_By" in df.columns else 0
    unassigned_cases = len(df[df[login_c].astype(str).str.strip() == ""]) if login_c in df.columns else 0
    active_workers = df[login_c].astype(str).str.strip().replace("", pd.NA).dropna().nunique() if login_c in df.columns else 0

    visa_col = find_visa_type_column(df)
    visa_counts = {}
    if visa_col and not df.empty:
        v_series = df[visa_col].astype(str).str.strip().str.upper()
        v_series = v_series[v_series != ""]
        visa_counts = v_series.value_counts().to_dict()

    lead_col = find_lead_column(df)
    lead_counts = {}
    if lead_col and not df.empty:
        l_series = df[lead_col].astype(str).str.strip()
        l_series = l_series[l_series != ""]
        lead_counts = l_series.value_counts().to_dict()

    if admin_dash_tab == "📊 Dashboard & Quick Actions":
        #[cite: 1] Dashboard Quick Actions UI
        pass

    elif admin_dash_tab == "📂 Case Records & Search":
        st.markdown("<div class='section-banner'>📂 Case Records & Lead / Visa Management Hub</div>", unsafe_allow_html=True)
        
        # --- LEAD NAMES BUTTON GROUP ---
        if lead_counts:
            st.markdown("<div class='action-group-title'>👥 Lead Names Distribution (Click any button to filter records):</div>", unsafe_allow_html=True)
            lead_items = list(lead_counts.items())
            l_cols = st.columns(min(len(lead_items), 4), gap="medium")
            for idx, (l_name, l_count) in enumerate(lead_items):
                col_idx = idx % len(l_cols)
                with l_cols[col_idx]:
                    btn_label = f"👤 {l_name}: {l_count}"
                    color_idx = idx % 8
                    st.markdown(f'<span class="visa-color-{color_idx}" style="display:none;"></span>', unsafe_allow_html=True)
                    if st.button(btn_label, key=f"admin_lead_card_btn_{l_name}_{idx}", use_container_width=True):
                        if st.session_state["selected_lead_filter"] == l_name:
                            st.session_state["selected_lead_filter"] = "All"
                        else:
                            st.session_state["selected_lead_filter"] = l_name
                        st.rerun()
            
            if st.session_state["selected_lead_filter"] != "All":
                if st.button(f"🔄 Active Lead Filter: [{st.session_state['selected_lead_filter']}] — Click to Reset Lead Filter", use_container_width=True, key="reset_lead_filter"):
                    st.session_state["selected_lead_filter"] = "All"
                    st.rerun()
            st.divider()

        # --- VISA TYPE BUTTON GROUP ---
        if visa_counts:
            st.markdown("<div class='action-group-title'>✈️ Visa Category Distribution (Click to filter):</div>", unsafe_allow_html=True)
            visa_items = list(visa_counts.items())
            v_cols = st.columns(min(len(visa_items), 4), gap="medium")
            for idx, (v_name, v_count) in enumerate(visa_items):
                col_idx = idx % len(v_cols)
                with v_cols[col_idx]:
                    btn_label = f"✈️ {v_name}: {v_count}"
                    color_idx = idx % 8
                    st.markdown(f'<span class="visa-color-{color_idx}" style="display:none;"></span>', unsafe_allow_html=True)
                    if st.button(btn_label, key=f"admin_visa_card_btn_{v_name}_{idx}", use_container_width=True):
                        if st.session_state["selected_visa_filter"] == v_name:
                            st.session_state["selected_visa_filter"] = "All"
                        else:
                            st.session_state["selected_visa_filter"] = v_name
                        st.rerun()
            
            if st.session_state["selected_visa_filter"] != "All":
                if st.button(f"🔄 Active Visa Filter: [{st.session_state['selected_visa_filter']}] — Click to Reset Visa Filter", use_container_width=True, key="reset_visa_filter"):
                    st.session_state["selected_visa_filter"] = "All"
                    st.rerun()
            st.divider()

        col_search, col_sheets = st.columns([1.2, 1.8], gap="medium")
        with col_search:
            admin_search = st.text_input("Search for any username, case, or info:", placeholder="e.g., Anil, Kiran...", key="admin_search_input").strip()
        with col_sheets:
            if sheets:
                st.markdown("📑 **Switch Sheet:**")
                sw_cols = st.columns(len(sheets), gap="small")
                for s_idx, s_name in enumerate(sheets):
                    with sw_cols[s_idx]:
                        is_act = (s_name == current_sheet)
                        color_c = "admin-btn-1" if is_act else "admin-btn-6"
                        b_txt = f"✅ {s_name}" if is_act else f"📄 {s_name}"
                        st.markdown(f'<span class="{color_c}" style="display:none;"></span>', unsafe_allow_html=True)
                        if st.button(b_txt, key=f"admin_sheet_btn_{s_idx}", use_container_width=True):
                            st.session_state["active_sheet"] = s_name
                            st.query_params["active_sheet"] = s_name
                            st.rerun()
        
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

        display_df = df
        if admin_search != "":
            admin_search_mask = display_df.apply(lambda row: row.astype(str).str.contains(admin_search, case=False, na=False).any(), axis=1)
            display_df = display_df[admin_search_mask]

        if st.session_state["selected_visa_filter"] != "All" and visa_col:
            target_visa = st.session_state["selected_visa_filter"]
            display_df = display_df[display_df[visa_col].astype(str).str.strip().str.upper() == target_visa]

        if st.session_state["selected_lead_filter"] != "All" and lead_col:
            target_lead = st.session_state["selected_lead_filter"]
            display_df = display_df[display_df[lead_col].astype(str).str.strip() == target_lead]

        st.caption(f"💡 Showing {len(display_df)} cases. Click USERNAME inside the table to view, edit, or delete details.")
        render_clickable_table(display_df, display_c, p_col, table_key=f"admin_{current_sheet}")

# --- Case Worker Workspace View Integration ---
else:
    sheets = get_sheet_names()
    current_sheet = st.session_state["active_sheet"]
    
    if current_sheet not in sheets:
        current_sheet = sheets[0]
        st.session_state["active_sheet"] = current_sheet

    df = load_data(sheet_name=current_sheet)
    if df is None:
        df = pd.DataFrame()
    df = reorder_dataframe_columns(df)
    current_user = st.session_state["username"].strip()
    
    all_perms = load_permissions()
    matched_w_key = next((k for k in all_perms if k.lower() == current_user.lower()), current_user)
    worker_perm = all_perms.get(matched_w_key, {"can_update": True, "can_add": True, "can_search": True, "can_download": False, "can_sync": False, "block_login": False})
    
    login_c = find_login_user_column(df)
    display_c = find_display_username_column(df)
    p_col = find_password_column(df)
    
    user_mask = df[login_c].astype(str).str.strip().str.lower() == current_user.lower() if login_c in df.columns else [False] * len(df)
    user_rows = df[user_mask]
        
    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    top_col1, top_col2, top_col3 = st.columns([6.5, 1.2, 1.2], gap="small")
    with top_col1:
        st.markdown(f"<div class='section-banner' style='margin:0;'>👋 Welcome, {current_user} — Case Worker Workspace</div>", unsafe_allow_html=True)
    with top_col2:
        if st.button("🔄 Refresh", key="worker_refresh"):
            st.cache_data.clear()
            st.session_state["open_dialog"] = None
            st.session_state["selected_row_idx"] = None
            st.rerun()
    with top_col3:
        if st.button("🚪 Logout", key="worker_logout"):
            record_worker_logout(current_user)
            st.query_params.clear()
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            st.session_state["attendance_recorded"] = False
            st.session_state["open_dialog"] = None
            st.session_state["selected_row_idx"] = None
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    worker_dash_tab = st.radio(
        "Worker Navigation",
        ["📊 My Workspace", "ℹ️ Case Worker Guide"],
        horizontal=True,
        label_visibility="collapsed",
        key="worker_dash_tab_radio"
    )
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    if worker_dash_tab == "📊 My Workspace":
        admin_notice = get_admin_notice()
        if admin_notice:
            st.markdown(f"""
                <div class="marquee-container">
                    <marquee scrollamount="5">📢 <b>Admin Announcement / Note:</b> {admin_notice}</marquee>
                </div>
            """, unsafe_allow_html=True)

        my_total = len(user_rows)
        my_updated = len(user_rows[user_rows["Updated_By"].astype(str).str.strip() != ""]) if "Updated_By" in user_rows.columns else 0

        visa_col = find_visa_type_column(user_rows)
        visa_counts = {}
        if visa_col and not user_rows.empty:
            v_series = user_rows[visa_col].astype(str).str.strip().str.upper()
            v_series = v_series[v_series != ""]
            visa_counts = v_series.value_counts().to_dict()

        lead_col = find_lead_column(user_rows)
        lead_counts = {}
        if lead_col and not user_rows.empty:
            l_series = user_rows[lead_col].astype(str).str.strip()
            l_series = l_series[l_series != ""]
            lead_counts = l_series.value_counts().to_dict()

        k1, k2 = st.columns(2, gap="medium")
        with k1:
            st.markdown(f"<div class='kpi-card'><div class='kpi-title'>📁 My Cases</div><div class='kpi-value'>{my_total}</div></div>", unsafe_allow_html=True)
        with k2:
            st.markdown(f"<div class='kpi-card green'><div class='kpi-title'>✅ Updated By Me</div><div class='kpi-value'>{my_updated}</div></div>", unsafe_allow_html=True)

        # --- WORKER LEAD NAMES BUTTON GROUP ---
        if lead_counts:
            st.markdown("<div class='action-group-title'>👥 My Lead Names Distribution (Click any button to filter):</div>", unsafe_allow_html=True)
            lead_items = list(lead_counts.items())
            l_cols = st.columns(min(len(lead_items), 4), gap="medium")
            for idx, (l_name, l_count) in enumerate(lead_items):
                col_idx = idx % len(l_cols)
                with l_cols[col_idx]:
                    btn_label = f"👤 {l_name}: {l_count}"
                    color_idx = idx % 8
                    st.markdown(f'<span class="visa-color-{color_idx}" style="display:none;"></span>', unsafe_allow_html=True)
                    if st.button(btn_label, key=f"worker_lead_btn_{l_name}_{idx}", use_container_width=True):
                        if st.session_state["selected_lead_filter"] == l_name:
                            st.session_state["selected_lead_filter"] = "All"
                        else:
                            st.session_state["selected_lead_filter"] = l_name
                        st.rerun()
            
            if st.session_state["selected_lead_filter"] != "All":
                if st.button(f"🔄 Active Lead Filter: [{st.session_state['selected_lead_filter']}] — Click to Reset", use_container_width=True, key="reset_worker_lead_filter"):
                    st.session_state["selected_lead_filter"] = "All"
                    st.rerun()

        # --- WORKER VISA TYPE BUTTON GROUP ---
        if visa_counts:
            st.markdown("<div class='action-group-title'>✈️ My Visa Category Distribution (Click to filter):</div>", unsafe_allow_html=True)
            visa_items = list(visa_counts.items())
            w_cols = st.columns(min(len(visa_items), 4), gap="medium")
            for idx, (v_name, v_count) in enumerate(visa_items):
                col_idx = idx % len(w_cols)
                with w_cols[col_idx]:
                    btn_label = f"✈️ {v_name}: {v_count}"
                    color_idx = idx % 8
                    st.markdown(f'<span class="visa-color-{color_idx}" style="display:none;"></span>', unsafe_allow_html=True)
                    if st.button(btn_label, key=f"worker_visa_btn_{v_name}_{idx}", use_container_width=True):
                        if st.session_state["selected_visa_filter"] == v_name:
                            st.session_state["selected_visa_filter"] = "All"
                        else:
                            st.session_state["selected_visa_filter"] = v_name
                        st.rerun()
            
            if st.session_state["selected_visa_filter"] != "All":
                if st.button(f"🔄 Active Visa Filter: [{st.session_state['selected_visa_filter']}] — Click to Reset", use_container_width=True, key="reset_worker_visa_filter"):
                    st.session_state["selected_visa_filter"] = "All"
                    st.rerun()

        st.markdown("<div class='action-group-title'>⚡ Workspace Tools</div>", unsafe_allow_html=True)
        has_sync_perm = worker_perm.get("can_sync", False)
        w_cols = st.columns(5 if has_sync_perm else 4, gap="medium")
        
        with w_cols[0]:
            if worker_perm.get("can_add", True):
                st.markdown('<span class="admin-btn-1" style="display:none;"></span>', unsafe_allow_html=True)
                if st.button("➕ Add New Case", use_container_width=True, key="worker_add_btn"):
                    st.session_state["open_dialog"] = "add_case"
                    st.rerun()
            else:
                st.markdown('<span class="admin-btn-7" style="display:none;"></span>', unsafe_allow_html=True)
                st.button("🚫 No Add Access", disabled=True, use_container_width=True, key="worker_add_disabled")
        with w_cols[1]:
            if worker_perm.get("can_search", True):
                st.markdown('<span class="admin-btn-2" style="display:none;"></span>', unsafe_allow_html=True)
                if st.button("🔍 Search Records", use_container_width=True, key="worker_search_btn"):
                    st.session_state["open_dialog"] = "search_data"
                    st.rerun()
            else:
                st.markdown('<span class="admin-btn-7" style="display:none;"></span>', unsafe_allow_html=True)
                st.button("🚫 No Search Access", disabled=True, use_container_width=True, key="worker_search_disabled")
        with w_cols[2]:
            if worker_perm.get("can_download", False):
                st.markdown('<span class="admin-btn-3" style="display:none;"></span>', unsafe_allow_html=True)
                worker_excel_bytes = get_worker_export_bytes(_get_file_mtime(), current_user, current_sheet)
                st.download_button(
                    label="📥 Export My Excel", 
                    data=worker_excel_bytes, 
                    file_name=f"{current_user}_all_sheets_cases.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    use_container_width=True, 
                    key="worker_dl_btn"
                )
            else:
                st.markdown('<span class="admin-btn-7" style="display:none;"></span>', unsafe_allow_html=True)
                st.button("🚫 No Download", disabled=True, use_container_width=True, key="worker_dl_disabled")
        with w_cols[3]:
            st.markdown('<span class="admin-btn-4" style="display:none;"></span>', unsafe_allow_html=True)
            if st.button("🔑 Change Password", use_container_width=True, key="worker_pw_btn"):
                st.session_state["open_dialog"] = "worker_password"
                st.rerun()

        if has_sync_perm:
            with w_cols[4]:
                st.markdown('<span class="admin-btn-5" style="display:none;"></span>', unsafe_allow_html=True)
                if st.button("🔄 Sync Excel", use_container_width=True, key="worker_sync_btn"):
                    st.session_state["open_dialog"] = "sync_excel"
                    st.rerun()
                
        st.divider()

        if sheets:
            st.markdown("<div class='action-group-title'>📑 Switch Sheet:</div>", unsafe_allow_html=True)
            sw_cols_w = st.columns(len(sheets), gap="small")
            for idx_w, s_name_w in enumerate(sheets):
                with sw_cols_w[idx_w]:
                    is_active_w = (s_name_w == current_sheet)
                    color_cls_w = "admin-btn-1" if is_active_w else "admin-btn-6"
                    btn_txt_w = f"✅ {s_name_w}" if is_active_w else f"📄 {s_name_w}"
                    st.markdown(f'<span class="{color_cls_w}" style="display:none;"></span>', unsafe_allow_html=True)
                    if st.button(btn_txt_w, key=f"worker_switch_sheet_{s_name_w}_{idx_w}", use_container_width=True):
                        st.session_state["active_sheet"] = s_name_w
                        st.query_params["active_sheet"] = s_name_w
                        st.rerun()
                
        st.divider()
        
        display_user_rows = user_rows
        if st.session_state["selected_visa_filter"] != "All" and visa_col:
            target_visa = st.session_state["selected_visa_filter"]
            display_user_rows = display_user_rows[display_user_rows[visa_col].astype(str).str.strip().str.upper() == target_visa]

        if st.session_state["selected_lead_filter"] != "All" and lead_col:
            target_lead = st.session_state["selected_lead_filter"]
            display_user_rows = display_user_rows[display_user_rows[lead_col].astype(str).str.strip() == target_lead]

        st.markdown(f"<div class='section-banner'>✏️ Your Assigned Cases ({current_sheet} — Total: {len(display_user_rows)})</div>", unsafe_allow_html=True)

        if display_user_rows.empty:
            st.info("No cases found matching your criteria in this sheet.")
        else:
            st.caption("💡 Click USERNAME inside the table to view, edit, or request deletion.")
            render_clickable_table(display_user_rows, display_c, p_col, table_key=f"worker_{current_user}_{current_sheet}")
