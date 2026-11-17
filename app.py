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

def find_visa_type_column(df):
    for c in df.columns:
        if any(term in c.lower() for term in ['visa', 'category', 'type', 'service']):
            return c
    return None

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
        transition: all 0.2s ease;
    }
    .banner-badge:hover {
        background: rgba(255, 255, 255, 0.22);
        border-color: #f59e0b;
    }
    .banner-info a {
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
        letter-spacing: 0.4px;
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
        background: linear-gradient(90deg, #e0e7ff 0%, #f3e8ff 100%);
        border-left: 5px solid #4f46e5;
        padding: 11px 18px;
        border-radius: 8px;
        margin: 8px 0px;
        font-weight: 800;
        color: #1e1b4b;
        font-size: 16px;
        box-shadow: 0 3px 10px rgba(79, 70, 229, 0.08);
    }
    .action-group-title {
        font-size: 13px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #312e81;
        margin: 10px 0 6px 2px;
    }
    .stButton > button, [data-testid="stDownloadButton"] > button {
        width: 100% !important;
        min-height: 46px !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        padding: 10px 16px !important;
        color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.2) !important;
        box-shadow: 0 4px 14px rgba(30, 58, 138, 0.2);
        margin-top: 4px !important;
        margin-bottom: 8px !important;
        transition: all 0.22s ease;
        white-space: nowrap !important;
        background: linear-gradient(135deg, #2563eb 100%, #1d4ed8 0%) !important;
    }
    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.35);
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
        st.error("❌ Your delete request was rejected by Admin.")

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
    st.info("💡 Add or update the scrolling marquee notice for all case workers:")
    
    with st.form("admin_notice_form"):
        notice_input = st.text_area("Notice Text:", value=current_notice, placeholder="e.g., All case workers must complete pending cases by today...")
        col_save_n, col_clear_n = st.columns(2)
        with col_save_n:
            submitted_notice = st.form_submit_button("💾 Save Notice", type="primary", use_container_width=True)
        with col_clear_n:
            cleared_notice = st.form_submit_button("🗑️ Clear Notice", use_container_width=True)

        if submitted_notice:
            save_admin_notice(notice_input)
            st.success("✅ Notice updated successfully for all case workers!")
            st.rerun()
        if cleared_notice:
            save_admin_notice("")
            st.success("✅ Notice cleared successfully!")
            st.rerun()

@st.dialog("👥 Case Worker Login Status", width="large")
def admin_login_status_dialog():
    st.info("💡 View real-time login and active status of all case workers:")
    
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
        
        item = {
            "worker": w,
            "status": status,
            "time": last_time
        }
        if status == "Logged In":
            logged_in_list.append(item)
        else:
            not_logged_in_list.append(item)
            
    tab_logged, tab_not_logged = st.tabs([f"🟢 Logged In ({len(logged_in_list)})", f"🔴 Not Logged In ({len(not_logged_in_list)})"])
    
    with tab_logged:
        if not logged_in_list:
            st.warning("Currently no case workers are logged in.")
        else:
            h1, h2, h3 = st.columns([2, 2, 2])
            with h1: st.markdown("**Case Worker Name**")
            with h2: st.markdown("**Login Status**")
            with h3: st.markdown("**Last Login Timestamp**")
            st.divider()
            for item in logged_in_list:
                c1, c2, c3 = st.columns([2, 2, 2])
                with c1: st.write(f"👤 **{item['worker']}**")
                with c2: st.success("🟢 Logged In")
                with c3: st.text(item['time'])
                
    with tab_not_logged:
        if not not_logged_in_list:
            st.success("All case workers are currently logged in!")
        else:
            h1, h2, h3 = st.columns([2, 2, 2])
            with h1: st.markdown("**Case Worker Name**")
            with h2: st.markdown("**Login Status**")
            with h3: st.markdown("**Last Login Timestamp**")
            st.divider()
            for item in not_logged_in_list:
                c1, c2, c3 = st.columns([2, 2, 2])
                with c1: st.write(f"👤 **{item['worker']}**")
                with c2: st.error("🔴 Not Logged In")
                with c3: st.text(item['time'])

@st.dialog("📊 Case Worker Attendance Report", width="large")
def admin_attendance_dialog():
    st.info("💡 Daily login and logout attendance logs for case workers:")
    
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
            
        st.download_button(
            label="📥 Download Attendance Report (.xlsx)",
            data=att_buffer.getvalue(),
            file_name="Case_Worker_Attendance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="download_attendance_excel"
        )

# 🌟 Worker Report Dialog with Today's, Total, and Visa Type-wise Breakdown
@st.dialog("📊 Case Worker & Visa Type Booked Cases Report", width="large")
def admin_worker_wise_report_dialog():
    st.info("💡 Booked cases report broken down by Case Worker, Today's activity, Total cases, and Visa Types:")
    
    sheets = get_sheet_names()
    today_str = get_ist_now().strftime("%Y-%m-%d")
    
    today_worker_stats = {}
    total_worker_stats = {}
    visa_type_stats = {}
    
    for s in sheets:
        df_s = load_data(sheet_name=s)
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
                
                # Total Cases Stats (Worker-wise)
                if w_name not in total_worker_stats:
                    total_worker_stats[w_name] = {"Total Cases": 0, "Visa_Breakdown": {}, "Sheets": set()}
                total_worker_stats[w_name]["Total Cases"] += 1
                total_worker_stats[w_name]["Sheets"].add(s)
                total_worker_stats[w_name]["Visa_Breakdown"][v_type] = total_worker_stats[w_name]["Visa_Breakdown"].get(v_type, 0) + 1
                
                # Visa Type Stats (Global)
                if v_type not in visa_type_stats:
                    visa_type_stats[v_type] = 0
                visa_type_stats[v_type] += 1
                
                # Today's Cases Check
                is_today = False
                for dc in date_cols:
                    val_str = str(r.get(dc, "")).strip()
                    if today_str in val_str:
                        is_today = True
                        break
                
                if not is_today and today_str in str(r.get("Updated_At", "")):
                    is_today = True

                if is_today:
                    if w_name not in today_worker_stats:
                        today_worker_stats[w_name] = {"Today Cases": 0, "Visa_Breakdown": {}, "Sheets": set()}
                    today_worker_stats[w_name]["Today Cases"] += 1
                    today_worker_stats[w_name]["Sheets"].add(s)
                    today_worker_stats[w_name]["Visa_Breakdown"][v_type] = today_worker_stats[w_name]["Visa_Breakdown"].get(v_type, 0) + 1

    tab_today, tab_total, tab_visa = st.tabs([
        "📅 Today's Booked Cases", 
        "📈 Total Booked Cases", 
        "✈️ Visa Type-wise Summary"
    ])
    
    with tab_today:
        if not today_worker_stats:
            st.warning(f"No cases booked today ({today_str}).")
        else:
            today_list = []
            for w, data in today_worker_stats.items():
                breakdown_str = ", ".join([f"{v}: {cnt}" for v, cnt in data["Visa_Breakdown"].items()])
                today_list.append({
                    "Case Worker": w, 
                    "Today Booked Cases": data["Today Cases"], 
                    "Visa Type Breakdown": breakdown_str,
                    "Active Sheets": ", ".join(list(data["Sheets"]))
                })
            df_today = pd.DataFrame(today_list)
            df_today.index = range(1, len(df_today) + 1)
            st.success(f"Total Cases Booked Today: {df_today['Today Booked Cases'].sum()}")
            st.dataframe(df_today, use_container_width=True)
            
            buf_today = io.BytesIO()
            with pd.ExcelWriter(buf_today, engine='openpyxl') as writer:
                df_today.to_excel(writer, sheet_name="Today_Booked_Report", index=False)
            st.download_button("📥 Download Today's Report (.xlsx)", data=buf_today.getvalue(), file_name=f"Today_Booked_Cases_{today_str}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_today_rep")

    with tab_total:
        if not total_worker_stats:
            st.warning("No records found.")
        else:
            total_list = []
            for w, data in total_worker_stats.items():
                breakdown_str = ", ".join([f"{v}: {cnt}" for v, cnt in data["Visa_Breakdown"].items()])
                total_list.append({
                    "Case Worker": w, 
                    "Total Booked Cases": data["Total Cases"], 
                    "Visa Type Breakdown": breakdown_str,
                    "Active Sheets": ", ".join(list(data["Sheets"]))
                })
            df_total = pd.DataFrame(total_list)
            df_total.index = range(1, len(df_total) + 1)
            st.success(f"Total Booked Cases across all workers: {df_total['Total Booked Cases'].sum()}")
            st.dataframe(df_total, use_container_width=True)
            
            buf_total = io.BytesIO()
            with pd.ExcelWriter(buf_total, engine='openpyxl') as writer:
                df_total.to_excel(writer, sheet_name="Total_Booked_Report", index=False)
            st.download_button("📥 Download Total Report (.xlsx)", data=buf_total.getvalue(), file_name="Total_Booked_Cases_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_total_rep")

    with tab_visa:
        if not visa_type_stats:
            st.warning("No visa type data found.")
        else:
            visa_list = [{"Visa Type / Category": v, "Total Booked Cases": count} for v, count in visa_type_stats.items()]
            df_visa = pd.DataFrame(visa_list)
            df_visa = df_visa.sort_values(by="Total Booked Cases", ascending=False)
            df_visa.index = range(1, len(df_visa) + 1)
            st.success(f"Total Visa Categories: {len(df_visa)}")
            st.dataframe(df_visa, use_container_width=True)
            
            buf_visa = io.BytesIO()
            with pd.ExcelWriter(buf_visa, engine='openpyxl') as writer:
                df_visa.to_excel(writer, sheet_name="Visa_Type_Summary", index=False)
            st.download_button("📥 Download Visa Type Summary (.xlsx)", data=buf_visa.getvalue(), file_name="Visa_Type_Wise_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_visa_rep")

@st.dialog("📋 New / Updated Cases & Calendar", width="large")
def admin_updated_cases_dialog():
    st.info("💡 View newly added cases, updated cases, and date-wise calendar records:")
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
                
    tab_new, tab_upd, tab_cal = st.tabs(["➕ New Cases", "✏️ Updated Cases", "📅 Calendar View"])
    
    with tab_new:
        if not all_rows:
            st.info("No new cases found.")
        else:
            combined_df = pd.concat(all_rows, ignore_index=True)
            p_col = find_password_column(combined_df)
            view_df = combined_df.drop(columns=[p_col], errors="ignore")
            new_cases_df = view_df[view_df["Action_Type"].astype(str).str.strip().str.lower() == "new case"]
            if new_cases_df.empty:
                st.info("No new cases found.")
            else:
                st.success(f"Total New Cases: {len(new_cases_df)}")
                render_clickable_table(new_cases_df, find_display_username_column(new_cases_df), p_col)
                
    with tab_upd:
        if not all_rows:
            st.info("No updated cases found.")
        else:
            combined_df = pd.concat(all_rows, ignore_index=True)
            p_col = find_password_column(combined_df)
            view_df = combined_df.drop(columns=[p_col], errors="ignore")
            updated_cases_df = view_df[view_df["Action_Type"].astype(str).str.strip().str.lower() == "updated case"]
            if updated_cases_df.empty:
                st.info("No updated cases found.")
            else:
                st.success(f"Total Updated Cases: {len(updated_cases_df)}")
                render_clickable_table(updated_cases_df, find_display_username_column(updated_cases_df), p_col)

    with tab_cal:
        st.write("📅 **Filter records by specific date:**")
        selected_cal_sheet = st.selectbox("Select Sheet for Calendar View", options=sheets, key="cal_sheet_select_tab")
        df_cal = load_data(sheet_name=selected_cal_sheet)
        if df_cal is not None and not df_cal.empty:
            date_cols = [c for c in df_cal.columns if any(k in c.lower() for k in ['date', 'time', 'day'])]
            if date_cols:
                selected_date_col = st.selectbox("Select Date Column:", options=date_cols, key="cal_date_col_tab")
                selected_date = st.date_input("Select Target Date:", value=datetime.date.today(), key="cal_target_date_tab")
                date_str = selected_date.strftime("%Y-%m-%d")
                
                mask = df_cal[selected_date_col].astype(str).str.contains(date_str, na=False)
                matched_df = df_cal[mask]
                st.write(f"📅 Records for date **{date_str}** ({len(matched_df)} found):")
                if matched_df.empty:
                    st.info(f"No records found for date ({date_str}).")
                else:
                    pc = find_password_column(matched_df)
                    display_c = find_display_username_column(matched_df)
                    render_clickable_table(matched_df, display_c, pc)
            else:
                st.warning("No date column found in this sheet.")

@st.dialog("🗑️ Pending Delete Requests", width="large")
def admin_delete_requests_dialog():
    st.info("💡 Delete requests submitted by case workers along with their reasons:")
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
    
    if not del_reqs_list:
        st.info("No pending delete requests.")
    else:
        st.success(f"Total Pending Delete Requests: {len(del_reqs_list)}")
        for i, r_dict in enumerate(del_reqs_list):
            s_name = r_dict["Sheet_Name"]
            orig_idx = r_dict["Original_Index"]
            worker_name = r_dict.get("Updated_By", "Unknown")
            req_time = r_dict.get("Updated_At", "")
            
            df_temp = load_data(sheet_name=s_name)
            disp_col = find_display_username_column(df_temp) if df_temp is not None else "USERNAME"
            case_username = str(r_dict.get(disp_col, "N/A"))
            delete_reason = str(r_dict.get("Delete_Reason", "Reason not specified"))

            with st.expander(f"📁 Sheet: {s_name} | Case User: {case_username} | Worker: {worker_name}"):
                st.markdown(f"👤 **Case Username:** `{case_username}`")
                st.markdown(f"✍️ **Delete Reason:** `{delete_reason}`")
                st.caption(f"Requested Time: {req_time}")
                
                col_acc, col_rej = st.columns(2)
                with col_acc:
                    if st.button("✅ Accept Delete (Permanent Delete)", key=f"accept_del_{s_name}_{orig_idx}_{i}", type="primary", use_container_width=True):
                        df_target = load_data(sheet_name=s_name)
                        if df_target is not None and orig_idx in df_target.index:
                            df_target = df_target.drop(orig_idx)
                            save_data(df_target, sheet_name=s_name)
                            st.success("✅ Delete request accepted! Case permanently deleted.")
                            st.rerun()
                with col_rej:
                    if st.button("❌ Reject Delete (Restore Case)", key=f"reject_del_{s_name}_{orig_idx}_{i}", use_container_width=True):
                        df_target = load_data(sheet_name=s_name)
                        if df_target is not None and orig_idx in df_target.index:
                            df_target.at[orig_idx, "Action_Type"] = "Delete Rejected"
                            save_data(df_target, sheet_name=s_name)
                            st.success("❌ Delete request rejected. Marked as Rejected for worker.")
                            st.rerun()

# 🌟 Professional Card-Style Permissions Dialog with Individual "Update" Button per Case Worker
@st.dialog("👥 Users, Passwords & Permissions (All Sheets)", width="large")
def permissions_dialog():
    st.markdown("""
        <div style="background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%); border-left: 5px solid #10b981; padding: 12px 18px; border-radius: 8px; margin-bottom: 16px; color: #065f46; font-size: 14px; font-weight: 600;">
            💡 Manage case worker credentials and access permissions below. Click the <b>"💾 Update"</b> button on any worker's card to save their changes immediately.
        </div>
    """, unsafe_allow_html=True)

    sheets = get_sheet_names()
    all_users_dict = {}
    for s in sheets:
        df_s = load_data(sheet_name=s)
        if df_s is not None and not df_s.empty:
            login_c = find_login_user_column(df_s)
            p_col = find_password_column(df_s)
            if login_c in df_s.columns and p_col in df_s.columns:
                for _, r in df_s.iterrows():
                    u_name = str(r[login_c]).strip()
                    pwd = str(r[p_col]).strip()
                    if u_name != "" and u_name.lower() != "admin":
                        if u_name not in all_users_dict:
                            all_users_dict[u_name] = pwd

    current_perms = load_permissions()
    
    if not all_users_dict:
        st.warning("No case workers found.")
    else:
        for idx, (u_name, curr_pwd) in enumerate(all_users_dict.items()):
            w_perm = current_perms.get(u_name, {"can_update": True, "can_add": True, "can_search": True, "can_download": False, "can_sync": False})
            
            with st.container(border=True):
                st.markdown(f"""
                    <div style="background: linear-gradient(90deg, #eff6ff 0%, #f8fafc 100%); border-left: 4px solid #2563eb; padding: 8px 14px; border-radius: 6px; margin-bottom: 10px;">
                        <span style="font-size: 15px; font-weight: 800; color: #1e3a8a;">👤 Case Worker: {u_name}</span>
                    </div>
                """, unsafe_allow_html=True)
                
                c1, c2, c3, c4, c5, c6, c_btn = st.columns([2.2, 1.1, 1.1, 1.1, 1.1, 1.1, 1.5], gap="small")
                
                with c1:
                    new_pwd_val = st.text_input("Password", value=curr_pwd, type="default", key=f"dlg_p_{idx}")
                with c2:
                    st.markdown("<div style='height: 26px;'></div>", unsafe_allow_html=True)
                    new_upd_val = st.checkbox("Update", value=w_perm.get("can_update", True), key=f"dlg_upd_{idx}")
                with c3:
                    st.markdown("<div style='height: 26px;'></div>", unsafe_allow_html=True)
                    new_add_val = st.checkbox("Add", value=w_perm.get("can_add", True), key=f"dlg_add_{idx}")
                with c4:
                    st.markdown("<div style='height: 26px;'></div>", unsafe_allow_html=True)
                    new_sch_val = st.checkbox("Search", value=w_perm.get("can_search", True), key=f"dlg_sch_{idx}")
                with c5:
                    st.markdown("<div style='height: 26px;'></div>", unsafe_allow_html=True)
                    new_dl_val = st.checkbox("Download", value=w_perm.get("can_download", False), key=f"dlg_dl_{idx}")
                with c6:
                    st.markdown("<div style='height: 26px;'></div>", unsafe_allow_html=True)
                    new_sync_val = st.checkbox("Sync", value=w_perm.get("can_sync", False), key=f"dlg_sync_{idx}")
                with c_btn:
                    st.markdown("<div style='height: 22px;'></div>", unsafe_allow_html=True)
                    if st.button("💾 Update", key=f"dlg_btn_upd_{idx}", type="primary", use_container_width=True):
                        try:
                            all_sheets_data = {}
                            xls = pd.ExcelFile(EXCEL_FILE)
                            for s in xls.sheet_names:
                                df_s = pd.read_excel(EXCEL_FILE, sheet_name=s, dtype=str)
                                lc = find_login_user_column(df_s)
                                pc = find_password_column(df_s)
                                if lc in df_s.columns and pc in df_s.columns:
                                    mask = df_s[lc].astype(str).str.strip().str.lower() == u_name.lower()
                                    if mask.any():
                                        df_s.loc[mask, pc] = new_pwd_val.strip()
                                all_sheets_data[s] = df_s
                            
                            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                                for s, df_val in all_sheets_data.items():
                                    df_val.to_excel(writer, sheet_name=str(s), index=False)
                            
                            current_perms[u_name] = {
                                "can_update": new_upd_val,
                                "can_add": new_add_val,
                                "can_search": new_sch_val,
                                "can_download": new_dl_val,
                                "can_sync": new_sync_val
                            }
                            save_permissions(current_perms)
                            st.success(f"✅ Updated '{u_name}' successfully!")
                        except Exception as e:
                            st.error(f"❌ Error updating '{u_name}': {e}")

@st.dialog("🔑 Change Admin Password", width="medium")
def admin_password_dialog():
    with st.form("dlg_admin_pass_form"):
        curr_p = st.text_input("Current Password", type="password")
        new_p = st.text_input("New Password", type="password")
        conf_p = st.text_input("Confirm Password", type="password")
        
        if st.form_submit_button("Update Admin Password", type="primary", use_container_width=True):
            admin_u, admin_p = get_admin_creds()
            if curr_p != admin_p:
                st.error("❌ Incorrect current password!")
            elif new_p.strip() == "":
                st.warning("⚠️ New password cannot be empty.")
            elif new_p != conf_p:
                st.error("❌ New passwords do not match!")
            else:
                save_admin_creds(admin_u, new_p.strip())
                st.success("✅ Admin password updated successfully!")

@st.dialog("📤 Sync Local Excel File to Cloud", width="medium")
def sync_excel_dialog():
    st.info("💡 Upload your updated local Excel file here to instantly synchronize it with the cloud.")
    uploaded_excel = st.file_uploader("Upload Updated Excel File (.xlsx)", type=["xlsx"])
    
    if uploaded_excel is not None:
        if st.button("🚀 Confirm & Sync to Cloud", type="primary", use_container_width=True):
            try:
                with open(EXCEL_FILE, "wb") as f:
                    f.write(uploaded_excel.getbuffer())
                st.success("✅ Cloud data updated successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {e}")

@st.dialog("🔑 Change Password", width="medium")
def worker_password_dialog(worker_name):
    sheets = get_sheet_names()
    worker_current_pass = ""
    for s in sheets:
        df_chk = load_data(sheet_name=s)
        if df_chk is not None and not df_chk.empty:
            lc = find_login_user_column(df_chk)
            pc = find_password_column(df_chk)
            matched_r = df_chk[df_chk[lc].astype(str).str.strip().str.lower() == worker_name.lower()]
            if not matched_r.empty:
                worker_current_pass = str(matched_r.iloc[0][pc]).strip()
                break

    with st.form("dlg_worker_pass_form"):
        curr_wp = st.text_input("Current Password", type="password")
        new_wp = st.text_input("New Password", type="password")
        conf_wp = st.text_input("Confirm Password", type="password")
        
        if st.form_submit_button("Update Password", type="primary", use_container_width=True):
            if curr_wp.strip() != worker_current_pass:
                st.error("❌ Incorrect current password!")
            elif new_wp.strip() == "":
                st.warning("⚠️ New password cannot be empty.")
            elif new_wp != conf_wp:
                st.error("❌ New passwords do not match!")
            else:
                try:
                    all_sheets_data = {}
                    xls = pd.ExcelFile(EXCEL_FILE)
                    for s in xls.sheet_names:
                        df_s = pd.read_excel(EXCEL_FILE, sheet_name=s, dtype=str)
                        lc = find_login_user_column(df_s)
                        pc = find_password_column(df_s)
                        if lc in df_s.columns and pc in df_s.columns:
                            mask = df_s[lc].astype(str).str.strip().str.lower() == worker_name.lower()
                            if mask.any():
                                df_s.loc[mask, pc] = new_wp.strip()
                        all_sheets_data[s] = df_s
                    
                    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                        for s, df_val in all_sheets_data.items():
                            df_val.to_excel(writer, sheet_name=str(s), index=False)
                    st.success("✅ Password updated successfully!")
                except Exception as e:
                    st.error(f"❌ Error: {e}")

@st.dialog("➕ Add New Case", width="large")
def add_case_dialog(worker_name):
    sheets = get_sheet_names()
    st.info("💡 Select the sheet where you want to add a new case:")
    
    selected_target_sheet = st.selectbox("Select Target Sheet for New Case", options=sheets, key="dlg_target_sheet")
    target_df = load_data(sheet_name=selected_target_sheet)
    t_login_col = find_login_user_column(target_df)
    t_pass_col = find_password_column(target_df)
    
    available_usernames = sorted(list(set(target_df[t_login_col].astype(str).str.strip().tolist())))
    available_usernames = [u for u in available_usernames if u != ""]
    available_usernames.insert(0, "Not Assigned")
    if worker_name not in available_usernames and worker_name != "":
        available_usernames.insert(1, worker_name)
    
    with st.form("dlg_new_case_form", clear_on_submit=True):
        assigned_username = st.selectbox("Assign Username", options=available_usernames, index=available_usernames.index(worker_name) if worker_name in available_usernames else 0)
        st.divider()
        st.write(f"📂 **{selected_target_sheet}** Sheet Headings:")
        
        form_data = {}
        cols = [c for c in target_df.columns if c not in [t_login_col, t_pass_col, "Updated_By", "Updated_At", "Action_Type", "Delete_Reason"]]
        for col_name in cols:
            if any(k in col_name.lower() for k in ['date', 'time', 'day']):
                sel_date = st.date_input(f"{col_name}", value=datetime.date.today(), key=f"add_date_{col_name}")
                form_data[col_name] = sel_date.strftime("%Y-%m-%d")
            else:
                form_data[col_name] = st.text_input(f"{col_name}", key=f"add_txt_{col_name}")
        
        if st.form_submit_button("🚀 Submit New Case", type="primary", use_container_width=True):
            try:
                new_row = {col: "" for col in target_df.columns}
                final_assigned_user = "" if assigned_username == "Not Assigned" else assigned_username.strip()
                new_row[t_login_col] = final_assigned_user
                
                current_pass = "admin123"
                matched_pass_rows = target_df[target_df[t_login_col].astype(str).str.strip().str.lower() == final_assigned_user.lower()]
                if not matched_pass_rows.empty:
                    current_pass = str(matched_pass_rows.iloc[0][t_pass_col]).strip()

                new_row[t_pass_col] = current_pass if final_assigned_user != "" else "admin123"
                new_row["Updated_By"] = worker_name
                new_row["Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                new_row["Action_Type"] = "New Case"
                new_row["Delete_Reason"] = ""
                
                for k, v in form_data.items():
                    new_row[k] = v
                    
                updated_target_df = pd.concat([target_df, pd.DataFrame([new_row])], ignore_index=True)
                save_data(updated_target_df, sheet_name=selected_target_sheet)
                st.success("✅ New case added successfully!")
            except Exception as e:
                st.error(f"❌ Error: {e}")

@st.dialog("🔍 Search Main Data", width="large")
def search_data_dialog():
    sheets = get_sheet_names()
    selected_search_sheet = st.selectbox("Select Sheet to Search:", options=sheets, key="dlg_search_sheet")
    df_search = load_data(sheet_name=selected_search_sheet)
    if df_search is not None and not df_search.empty:
        display_c = find_display_username_column(df_search)
        p_col = find_password_column(df_search)
        
        search_query = st.text_input("Type any branch name or person name:", placeholder="e.g., ALLAGADDA, Adoni...", key="dlg_search_input").strip()
        
        if search_query != "":
            main_search_mask = df_search.apply(lambda row: row.astype(str).str.contains(search_query, case=False, na=False).any(), axis=1)
            found_data = df_search[main_search_mask]
            
            if found_data.empty:
                st.warning(f"No details found for '{search_query}'.")
            else:
                st.success(f"Results for '{search_query}' ({len(found_data)} rows):")
                view_only_df = found_data.drop(columns=[display_c, p_col, "Updated_By", "Updated_At", "Action_Type", "Delete_Reason"], errors="ignore")
                view_only_df.index = range(1, len(view_only_df) + 1)
                st.dataframe(view_only_df, use_container_width=True)

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
elif st.session_state["open_dialog"] == "admin_worker_report":
    admin_worker_wise_report_dialog()
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

if not st.session_state["logged_in"]:
    selected_page = st.radio(
        "Navigation",
        ["🏠 Home", "ℹ️ About Us", "✈️ Services", "⭐ Success Stories", "📞 Contact Us", "🔐 Portal Login"],
        horizontal=True,
        label_visibility="collapsed"
    )
    
    st.markdown("<div style='height: 15px;'></div>", unsafe_allow_html=True)

    if selected_page == "🏠 Home":
        st.markdown("""
            <div class="hero-section">
                <div class="hero-title">🌍 Welcome to HemanVisa Global Experts</div>
                <div class="hero-subtitle">Your trusted partner for international study, work visas, permanent residency, and seamless global mobility. We empower your dreams with expert guidance and transparent processing.</div>
            </div>
        """, unsafe_allow_html=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("""
                <div class="corporate-card">
                    <h3>🎓 Study Abroad</h3>
                    <p>Secure admissions in top global universities with complete visa documentation and interview preparation support.</p>
                </div>
            """, unsafe_allow_html=True)
        with c2:
            st.markdown("""
                <div class="corporate-card">
                    <h3>💼 Work & Skilled Visas</h3>
                    <p>Expand your career internationally with employer-sponsored work visas and skilled migration pathways.</p>
                </div>
            """, unsafe_allow_html=True)
        with c3:
            st.markdown("""
                <div class="corporate-card">
                    <h3>✈️ Permanent Residency</h3>
                    <p>Settle abroad permanently with hassle-free PR processing, point-based assessments, and legal consulting.</p>
                </div>
            """, unsafe_allow_html=True)

    elif selected_page == "ℹ️ About Us":
        st.markdown("<div class='section-banner'>ℹ️ About HemanVisa & Immigration Services</div>", unsafe_allow_html=True)
        col_a1, col_a2 = st.columns([1.5, 1])
        with col_a1:
            st.write("""
                HemanVisa is a premier global immigration and visa consulting firm dedicated to turning international aspirations into reality. 
                With years of operational excellence, our seasoned case workers and migration experts deliver personalized solutions for students, professionals, and families.
                
                Our core mission is to maintain absolute transparency, precision in documentation, and end-to-end case tracking so every client experiences a smooth journey.
            """)
        with col_a2:
            st.info("🌟 **Our Core Values:**\n* Transparency & Trust\n* 100% Document Verification\n* Dedicated Case Workers\n* Timely Processing")

    elif selected_page == "✈️ Services":
        st.markdown("<div class='section-banner'>✈️ Our Premium Global Services</div>", unsafe_allow_html=True)
        s1, s2 = st.columns(2)
        with s1:
            st.markdown("""
                <div class="corporate-card">
                    <h3>📋 DS-160 & Visa Filing</h3>
                    <p>Professional assistance with visa application forms, document verification, embassy appointment booking, and mock interview coaching.</p>
                </div>
                <div class="corporate-card">
                    <h3>🏖️ Tourist & Business Visas</h3>
                    <p>Quick turnaround tourist visa processing for USA, UK, Canada, Europe, Australia, and rest of the world.</p>
                </div>
            """, unsafe_allow_html=True)
        with s2:
            st.markdown("""
                <div class="corporate-card">
                    <h3>📚 Student Career Counseling</h3>
                    <p>Course selection, university applications, scholarship guidance, and financial documentation support.</p>
                </div>
                <div class="corporate-card">
                    <h3>🏢 Corporate Immigration</h3>
                    <p>Seamless mobility solutions for multinational enterprises transferring talent across international borders.</p>
                </div>
            """, unsafe_allow_html=True)

    elif selected_page == "⭐ Success Stories":
        st.markdown("<div class='section-banner'>⭐ Client Success Stories & Testimonials</div>", unsafe_allow_html=True)
        t1, t2, t3 = st.columns(3)
        with t1:
            st.markdown("""
                <div class="corporate-card">
                    <p>"HemanVisa made my Canada study visa process extremely smooth. The case workers were responsive and professional!"</p>
                    <b style="color:#2563eb;">— Rajesh K. (Canada PR)</b>
                </div>
            """, unsafe_allow_html=True)
        with t2:
            st.markdown("""
                <div class="corporate-card">
                    <p>"Extremely grateful for the prompt documentation and regular status updates. Got my US work visa approved without any hassle."</p>
                    <b style="color:#2563eb;">— Sneha R. (USA Work Visa)</b>
                </div>
            """, unsafe_allow_html=True)
        with t3:
            st.markdown("""
                <div class="corporate-card">
                    <p>"Best consulting firm with transparent pricing and dedicated case officers. Highly recommended for anyone planning abroad."</p>
                    <b style="color:#2563eb;">— Anil M. (UK Student Visa)</b>
                </div>
            """, unsafe_allow_html=True)

    elif selected_page == "📞 Contact Us":
        st.markdown("<div class='section-banner'>📞 Get in Touch with Our Experts</div>", unsafe_allow_html=True)
        c_col1, c_col2 = st.columns(2)
        with c_col1:
            st.write("""
                **Corporate Headquarters:**  
                🌐 Global Visa & Immigration Experts  
                ✉️ Email: services@hemanvisa.com  
                📞 Phone: +91 95338 40632  
                
                **Office Hours:**  
                Monday – Saturday: 9:30 AM – 7:00 PM  
            """)
        with c_col2:
            with st.form("contact_form"):
                st.subheader("Send Us a Message")
                st.text_input("Your Name")
                st.text_input("Email / Phone")
                st.text_area("Your Inquiry / Visa Requirement")
                if st.form_submit_button("Submit Inquiry", type="primary", use_container_width=True):
                    st.success("✅ Thank you! Our expert will contact you shortly.")

    elif selected_page == "🔐 Portal Login":
        col_l1, col_l2, col_l3 = st.columns([1, 1.4, 1])
        with col_l2:
            with st.container(border=True):
                st.markdown("<h3 style='text-align:center; color:#1e3a8a;'>🔐 Portal Login</h3>", unsafe_allow_html=True)
                st.markdown("<p style='text-align:center; font-size:13px; color:gray;'>Admin & Case Worker Secure Access</p>", unsafe_allow_html=True)
                sheets = get_sheet_names()
                if not sheets:
                    st.error(f"'{EXCEL_FILE}' file not found.")
                else:
                    with st.form("login_form"):
                        username_input = st.text_input("👤 Username", placeholder="Type your username")
                        password_input = st.text_input("🔑 Password", type="password", placeholder="Type your password")
                        submit_btn = st.form_submit_button("🚀 Login to Portal", type="primary", use_container_width=True)
                        
                        if submit_btn:
                            u_in = username_input.strip()
                            p_in = password_input.strip()
                            admin_u, admin_p = get_admin_creds()
                            
                            if u_in.lower() == admin_u.lower() and p_in == admin_p:
                                st.session_state["logged_in"] = True
                                st.session_state["username"] = admin_u
                                st.session_state["role"] = "admin"
                                st.session_state["open_dialog"] = None
                                st.session_state["selected_row_idx"] = None
                                st.query_params.clear()
                                st.query_params["logged_in"] = "true"
                                st.query_params["username"] = admin_u
                                st.query_params["role"] = "admin"
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
                                            db_user = str(r[login_c]).strip()
                                            db_pass = str(r[p_col]).strip()
                                            if db_user.lower() == u_in.lower() and db_pass == p_in:
                                                logged_in_success = True
                                                matched_sheet = s
                                                actual_matched_username = db_user
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
                                    st.session_state["open_dialog"] = None
                                    st.session_state["selected_row_idx"] = None
                                    st.query_params.clear()
                                    st.query_params["logged_in"] = "true"
                                    st.query_params["username"] = actual_matched_username
                                    st.query_params["role"] = "worker"
                                    st.query_params["active_sheet"] = matched_sheet
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
        st.markdown(f"<div class='section-banner' style='margin:0;'>👑 Admin Control Panel — {current_sheet}</div>", unsafe_allow_html=True)
    with top_col2:
        if st.button("🔄 Refresh", key="admin_refresh"):
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
    updated_cases = len(df[df["Updated_By"].astype(str).str.strip() != ""])
    unassigned_cases = len(df[df[login_c].astype(str).str.strip() == ""])
    active_workers = df[login_c].astype(str).str.strip().replace("", pd.NA).dropna().nunique()

    visa_col = find_visa_type_column(df)
    visa_counts = {}
    if visa_col and not df.empty:
        for val in df[visa_col].astype(str):
            v_clean = val.strip().upper()
            if v_clean != "":
                visa_counts[v_clean] = visa_counts.get(v_clean, 0) + 1

    if admin_dash_tab == "📊 Dashboard & Quick Actions":
        k1, k2, k3, k4 = st.columns(4, gap="medium")
        with k1:
            st.markdown(f"<div class='kpi-card'><div class='kpi-title'>📋 Total Cases</div><div class='kpi-value'>{total_cases}</div></div>", unsafe_allow_html=True)
        with k2:
            st.markdown(f"<div class='kpi-card green'><div class='kpi-title'>✅ Updated Cases</div><div class='kpi-value'>{updated_cases}</div></div>", unsafe_allow_html=True)
        with k3:
            st.markdown(f"<div class='kpi-card gold'><div class='kpi-title'>⚠️ Unassigned</div><div class='kpi-value'>{unassigned_cases}</div></div>", unsafe_allow_html=True)
        with k4:
            st.markdown(f"<div class='kpi-card purple'><div class='kpi-title'>👥 Workers</div><div class='kpi-value'>{active_workers}</div></div>", unsafe_allow_html=True)

        st.markdown("<div class='action-group-title'>⚡ Case & Request Management</div>", unsafe_allow_html=True)
        r1_c1, r1_c2, r1_c3, r1_c4, r1_c5 = st.columns(5, gap="small")
        with r1_c1:
            if st.button("📋 New / Updated", use_container_width=True, key="admin_upd_cases_btn"):
                st.session_state["open_dialog"] = "admin_updated_cases"
                st.rerun()
        with r1_c2:
            if st.button("📊 Worker Report", use_container_width=True, key="admin_worker_rep_btn"):
                st.session_state["open_dialog"] = "admin_worker_report"
                st.rerun()
        with r1_c3:
            if st.button("🗑️ Delete Req", use_container_width=True, key="admin_del_req_btn"):
                st.session_state["open_dialog"] = "admin_delete_requests"
                st.rerun()
        with r1_c4:
            if st.button("🔄 Sync Excel", use_container_width=True, key="admin_sync_btn"):
                st.session_state["open_dialog"] = "sync_excel"
                st.rerun()
        with r1_c5:
            if st.button("📢 Notice", use_container_width=True, key="admin_notice_btn"):
                st.session_state["open_dialog"] = "admin_notice"
                st.rerun()

        st.markdown("<div class='action-group-title'>🛡️ Team, Attendance & Security</div>", unsafe_allow_html=True)
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

        st.divider()

    elif admin_dash_tab == "📂 Case Records & Search":
        st.markdown("<div class='section-banner'>📂 Case Records & Visa Type Management Hub</div>", unsafe_allow_html=True)
        
        if visa_counts:
            st.markdown("<div class='action-group-title'>✈️ Visa Category Distribution (Click any card to filter):</div>", unsafe_allow_html=True)
            visa_items = list(visa_counts.items())
            mid = (len(visa_items) + 1) // 2
            rows = [visa_items[:mid], visa_items[mid:]]
            
            for row_items in rows:
                if row_items:
                    v_cols = st.columns(len(row_items), gap="medium")
                    for idx, (v_name, v_count) in enumerate(row_items):
                        with v_cols[idx]:
                            btn_label = f"✈️ {v_name}: {v_count}"
                            if st.button(btn_label, key=f"admin_visa_card_btn_{v_name}_{idx}", use_container_width=True):
                                if st.session_state["selected_visa_filter"] == v_name:
                                    st.session_state["selected_visa_filter"] = "All"
                                else:
                                    st.session_state["selected_visa_filter"] = v_name
                                st.rerun()
            
            if st.session_state["selected_visa_filter"] != "All":
                if st.button(f"🔄 Active Filter: [{st.session_state['selected_visa_filter']}] — Click to Reset (Show All Cases)", use_container_width=True, key="reset_visa_filter"):
                    st.session_state["selected_visa_filter"] = "All"
                    st.rerun()
            st.divider()

        col_search, col_sheets = st.columns([1.2, 1.8], gap="medium")
        with col_search:
            admin_search = st.text_input("Search for any username, case, or info:", placeholder="e.g., Anil, Kiran...", key="admin_search_input").strip()
        with col_sheets:
            if sheets:
                st.markdown("📑 **Sheets Switch:**")
                st.radio("Select Sheet Admin", options=sheets, index=sheets.index(current_sheet) if current_sheet in sheets else 0, horizontal=True, label_visibility="collapsed", key="admin_sheet_radio", on_change=update_admin_sheet)
        
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

        display_df = df
        if admin_search != "":
            admin_search_mask = display_df.apply(lambda row: row.astype(str).str.contains(admin_search, case=False, na=False).any(), axis=1)
            display_df = display_df[admin_search_mask]

        if st.session_state["selected_visa_filter"] != "All" and visa_col:
            target_visa = st.session_state["selected_visa_filter"]
            display_df = display_df[display_df[visa_col].astype(str).str.strip().str.upper() == target_visa]

        st.caption(f"💡 Showing {len(display_df)} cases. Click USERNAME inside the table to view, edit, or delete details.")
        render_clickable_table(display_df, display_c, p_col)
                    
        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as f:
                excel_bytes = f.read()
            st.download_button(
                label="📥 Download Complete Excel File (All Sheets)", 
                data=excel_bytes, 
                file_name=EXCEL_FILE, 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True, 
                key="admin_dl_all_btn"
            )

    elif admin_dash_tab == "ℹ️ System Guidelines":
        st.markdown("<div class='section-banner'>ℹ️ HemanVisa Admin Operations & Workflow Guide</div>", unsafe_allow_html=True)
        g1, g2 = st.columns(2, gap="medium")
        with g1:
            st.markdown("""
                <div class="corporate-card">
                    <h3>👥 User & Permission Management</h3>
                    <p>Assign specific privileges (Update, Add, Search, Download, Sync) to case workers to maintain strict security and accountability across all sheets.</p>
                </div>
                <div class="corporate-card">
                    <h3>🗑️ Delete Request Approvals</h3>
                    <p>When case workers request case deletions with reasons, review them in the 'Delete Requests' section to either permanently delete or reject the request.</p>
                </div>
            """, unsafe_allow_html=True)
        with g2:
            st.markdown("""
                <div class="corporate-card">
                    <h3>📊 Attendance & Login Tracking</h3>
                    <p>Monitor real-time worker logins, active status, and automated daily attendance logs with instant Excel report exports.</p>
                </div>
                <div class="corporate-card">
                    <h3>📢 Live Announcements</h3>
                    <p>Broadcast scrolling marquee announcements directly to all active case worker screens using the 'Broadcast Notice' feature.</p>
                </div>
            """, unsafe_allow_html=True)

else:
    sheets = get_sheet_names()
    current_sheet = st.session_state["active_sheet"]
    
    if current_sheet not in sheets:
        current_sheet = sheets[0]
        st.session_state["active_sheet"] = current_sheet

    df = load_data(sheet_name=current_sheet)
    current_user = st.session_state["username"].strip()
    
    all_perms = load_permissions()
    worker_perm = all_perms.get(current_user, {"can_update": True, "can_add": True, "can_search": True, "can_download": False, "can_sync": False})
    
    login_c = find_login_user_column(df)
    display_c = find_display_username_column(df)
    p_col = find_password_column(df)
    
    user_mask = df[login_c].astype(str).str.strip().str.lower() == current_user.lower()
    user_rows = df[user_mask]
        
    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    top_col1, top_col2, top_col3 = st.columns([6.5, 1.2, 1.2], gap="small")
    with top_col1:
        st.markdown(f"<div class='section-banner' style='margin:0;'>👋 Welcome, {current_user} — Case Worker Workspace</div>", unsafe_allow_html=True)
    with top_col2:
        if st.button("🔄 Refresh", key="worker_refresh"):
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
        my_updated = len(user_rows[user_rows["Updated_By"].astype(str).str.strip() != ""])

        visa_col = find_visa_type_column(user_rows)
        visa_counts = {}
        if visa_col and not user_rows.empty:
            for val in user_rows[visa_col].astype(str):
                v_clean = val.strip().upper()
                if v_clean != "":
                    visa_counts[v_clean] = visa_counts.get(v_clean, 0) + 1

        k1, k2 = st.columns(2, gap="medium")
        with k1:
            st.markdown(f"<div class='kpi-card'><div class='kpi-title'>📁 My Cases</div><div class='kpi-value'>{my_total}</div></div>", unsafe_allow_html=True)
        with k2:
            st.markdown(f"<div class='kpi-card green'><div class='kpi-title'>✅ Updated By Me</div><div class='kpi-value'>{my_updated}</div></div>", unsafe_allow_html=True)

        if visa_counts:
            st.markdown("<div class='action-group-title'>✈️ My Visa Category Distribution (Click any button to filter):</div>", unsafe_allow_html=True)
            visa_items = list(visa_counts.items())
            
            if visa_items:
                w_cols = st.columns(len(visa_items), gap="medium")
                for idx, (v_name, v_count) in enumerate(visa_items):
                    with w_cols[idx]:
                        btn_label = f"✈️ {v_name}: {v_count}"
                        if st.button(btn_label, key=f"worker_visa_btn_{v_name}_{idx}", use_container_width=True):
                            if st.session_state["selected_visa_filter"] == v_name:
                                st.session_state["selected_visa_filter"] = "All"
                            else:
                                st.session_state["selected_visa_filter"] = v_name
                            st.rerun()
            
            if st.session_state["selected_visa_filter"] != "All":
                if st.button(f"🔄 Active Filter: [{st.session_state['selected_visa_filter']}] — Click to Reset (Show All My Cases)", use_container_width=True, key="reset_worker_visa_filter"):
                    st.session_state["selected_visa_filter"] = "All"
                    st.rerun()

        st.markdown("<div class='action-group-title'>⚡ Workspace Tools</div>", unsafe_allow_html=True)
        has_sync_perm = worker_perm.get("can_sync", False)
        w_cols = st.columns(5 if has_sync_perm else 4, gap="medium")
        
        with w_cols[0]:
            if worker_perm.get("can_add", True):
                if st.button("➕ Add New Case", use_container_width=True, key="worker_add_btn"):
                    st.session_state["open_dialog"] = "add_case"
                    st.rerun()
            else:
                st.button("🚫 No Add Access", disabled=True, use_container_width=True, key="worker_add_disabled")
        with w_cols[1]:
            if worker_perm.get("can_search", True):
                if st.button("🔍 Search Records", use_container_width=True, key="worker_search_btn"):
                    st.session_state["open_dialog"] = "search_data"
                    st.rerun()
            else:
                st.button("🚫 No Search Access", disabled=True, use_container_width=True, key="worker_search_disabled")
        with w_cols[2]:
            if worker_perm.get("can_download", False):
                worker_buffer = io.BytesIO()
                with pd.ExcelWriter(worker_buffer, engine='openpyxl') as writer:
                    has_user_data = False
                    for s in sheets:
                        df_s = load_data(sheet_name=s)
                        if df_s is not None and not df_s.empty:
                            login_cs = find_login_user_column(df_s)
                            u_rows_s = df_s[df_s[login_cs].astype(str).str.strip().str.lower() == current_user.lower()]
                            if not u_rows_s.empty:
                                has_user_data = True
                                u_rows_s.to_excel(writer, sheet_name=s, index=False)
                    
                    if not has_user_data:
                        user_rows.to_excel(writer, sheet_name=current_sheet, index=False)
                
                st.download_button(
                    label="📥 Export My Excel", 
                    data=worker_buffer.getvalue(), 
                    file_name=f"{current_user}_all_sheets_cases.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    use_container_width=True, 
                    key="worker_dl_btn"
                )
            else:
                st.button("🚫 No Download", disabled=True, use_container_width=True, key="worker_dl_disabled")
        with w_cols[3]:
            if st.button("🔑 Change Password", use_container_width=True, key="worker_pw_btn"):
                st.session_state["open_dialog"] = "worker_password"
                st.rerun()

        if has_sync_perm:
            with w_cols[4]:
                if st.button("🔄 Sync Excel", use_container_width=True, key="worker_sync_btn"):
                    st.session_state["open_dialog"] = "sync_excel"
                    st.rerun()
                
        st.divider()

        if sheets:
            st.write("📑 **Sheets Switch:**")
            st.radio("Select Worker Sheet", options=sheets, index=sheets.index(current_sheet) if current_sheet in sheets else 0, horizontal=True, label_visibility="collapsed", key="worker_sheet_radio", on_change=update_worker_sheet)
                
        st.divider()
        
        display_user_rows = user_rows
        if st.session_state["selected_visa_filter"] != "All" and visa_col:
            target_visa = st.session_state["selected_visa_filter"]
            display_user_rows = display_user_rows[display_user_rows[visa_col].astype(str).str.strip().str.upper() == target_visa]

        st.markdown(f"<div class='section-banner'>✏️ Your Assigned Cases ({current_sheet} — Total: {len(display_user_rows)})</div>", unsafe_allow_html=True)

        if display_user_rows.empty:
            st.info("No cases found matching your criteria in this sheet.")
        else:
            st.caption("💡 Click USERNAME inside the table to view, edit, or request deletion.")
            render_clickable_table(display_user_rows, display_c, p_col)

    elif worker_dash_tab == "ℹ️ Case Worker Guide":
        st.markdown("<div class='section-banner'>ℹ️ Case Worker Quick Guide & Instructions</div>", unsafe_allow_html=True)
        w_g1, w_g2 = st.columns(2, gap="medium")
        with w_g1:
            st.markdown("""
                <div class="corporate-card">
                    <h3>✏️ Editing & Updating Cases</h3>
                    <p>Click on any USERNAME inside your assigned cases table to open the pop-up modal. You can review all details and update information instantly.</p>
                </div>
                <div class="corporate-card">
                    <h3>➕ Adding New Cases</h3>
                    <p>Use the 'Add New Case' button in the toolbar to create new records directly in your designated sheet with automatic timestamping.</p>
                </div>
            """, unsafe_allow_html=True)
        with w_g2:
            st.markdown("""
                <div class="corporate-card">
                    <h3>🗑️ Requesting Case Deletions</h3>
                    <p>If a case needs removal, click 'Request Delete' and provide a mandatory reason. The admin will review and approve or reject your request.</p>
                </div>
                <div class="corporate-card">
                    <h3>🔒 Security & Password</h3>
                    <p>Keep your account secure by changing your password periodically using the 'Change Password' button in the toolbar.</p>
                </div>
            """, unsafe_allow_html=True)
