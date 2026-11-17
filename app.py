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

# 🎨 CSS STYLING
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
        margin-left: 20px;
    }
    .sticky-top-banner img {
        max-height: 55px;
        width: auto;
        border-radius: 6px;
        background: #ffffff;
        padding: 2px;
    }
    .banner-middle {
        text-align: center;
        flex: 1;
        padding: 0 15px;
    }
    .world-map-title {
        font-size: 18px;
        font-weight: 800;
        color: #ffffff;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        margin-bottom: 2px;
    }
    .world-map-subtitle {
        font-size: 11px;
        font-weight: 700;
        color: #fde047;
        letter-spacing: 1.5px;
    }
    .banner-info {
        display: flex;
        align-items: center;
        gap: 10px;
        flex-wrap: wrap;
        font-size: 12px;
        font-weight: 600;
        color: #e2e8f0;
    }
    .banner-badge {
        background: rgba(255, 255, 255, 0.12);
        border: 1px solid rgba(255, 255, 255, 0.25);
        padding: 5px 10px;
        border-radius: 8px;
    }
    .banner-badge a {
        color: #ffffff;
        text-decoration: none;
    }

    .block-container {
        padding-top: 105px !important;
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
    }
    .hero-title {
        font-size: 36px;
        font-weight: 800;
        margin-bottom: 14px;
    }
    .hero-subtitle {
        font-size: 16px;
        opacity: 0.95;
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

    .sticky-control-panel {
        position: sticky;
        top: 86px;
        z-index: 99998;
        background-color: #ffffff;
        padding: 8px 0 10px 0;
        border-bottom: 2px solid #e2e8f0;
        margin-bottom: 14px;
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

    /* 🎨 VISA CATEGORY PILL BUTTONS (Soft Pastel) */
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
</style>
""", unsafe_allow_html=True)

found_logo = next((f for f in LOGO_FILES if os.path.exists(f)), None)
logo_html = ""
if found_logo:
    ext = found_logo.split(".")[-1].lower()
    mime_type = "image/png" if ext == "png" else "image/jpeg"
    img_b64 = get_base64_image(found_logo) if 'get_base64_image' in globals() else ""
    try:
        with open(found_logo, "rb") as img_file:
            img_b64 = base64.b64encode(img_file.read()).decode()
        logo_html = f'<div class="logo-container"><img src="data:{mime_type};base64,{img_b64}" alt="HemanVisa Logo"></div>'
    except:
        logo_html = '<div class="logo-container"><h3 style="margin:0; color:#ffffff;">HemanVisa</h3></div>'
else:
    logo_html = '<div class="logo-container"><h3 style="margin:0; color:#ffffff;">HemanVisa</h3></div>'

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
        col_act1, col_act2 = st.columns(2)
        with col_act1:
            if st.button("✏️ Edit Case Details", type="primary", use_container_width=True, key="btn_trigger_edit"):
                st.session_state["is_editing_popup"] = True
                st.rerun()
        with col_act2:
            if role == "admin":
                if st.button("🗑️ Delete Case", type="secondary", use_container_width=True, key="btn_trigger_delete"):
                    df_full = df_full.drop(row_idx)
                    save_data(df_full, sheet_name=sheet_name)
                    st.session_state["selected_row_idx"] = None
                    st.session_state["open_dialog"] = None
                    st.success("✅ Case deleted successfully!")
                    st.rerun()
            else:
                if st.button("🗑️ Request Delete", type="secondary", use_container_width=True, key="btn_trigger_del_req_mode"):
                    st.session_state["is_delete_request_mode"] = True
                    st.rerun()

    elif st.session_state["is_delete_request_mode"]:
        st.markdown("### 🗑️ Request Delete Case")
        with st.form(f"delete_reason_form_{row_idx}"):
            del_reason = st.text_area("✍️ కారణం ఇక్కడ రాయండి:", placeholder="Reason required...")
            if st.form_submit_button("🚀 Submit Delete Request", type="primary", use_container_width=True):
                if del_reason.strip():
                    df_full.at[row_idx, "Updated_By"] = current_user
                    df_full.at[row_idx, "Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    df_full.at[row_idx, "Action_Type"] = "Delete Request"
                    df_full.at[row_idx, "Delete_Reason"] = del_reason.strip()
                    save_data(df_full, sheet_name=sheet_name)
                    st.session_state["is_delete_request_mode"] = False
                    st.session_state["selected_row_idx"] = None
                    st.session_state["open_dialog"] = None
                    st.success("✅ Delete request sent!")
                    st.rerun()

    else:
        st.markdown("### ✏️ Edit Case Details")
        with st.form(f"popup_edit_form_{row_idx}"):
            updated_values = {}
            for col_name in df_full.columns:
                if col_name not in [display_c, p_col, "Updated_By", "Updated_At", "Delete_Reason"]:
                    updated_values[col_name] = st.text_input(f"{col_name}", value=str(row_data[col_name]))
            if st.form_submit_button("💾 Save & Update", type="primary", use_container_width=True):
                for col_name, new_val in updated_values.items():
                    df_full.at[row_idx, col_name] = str(new_val).strip()
                df_full.at[row_idx, "Updated_By"] = current_user
                df_full.at[row_idx, "Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                df_full.at[row_idx, "Action_Type"] = "Updated Case"
                save_data(df_full, sheet_name=sheet_name)
                st.session_state["is_editing_popup"] = False
                st.session_state["selected_row_idx"] = None
                st.session_state["open_dialog"] = None
                st.success("✅ Updated successfully!")
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
    <div style="max-height: 500px; overflow-y: auto; overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 12px; margin-bottom: 16px;">
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
        html += '<td style="padding: 10px; border: 1px solid #e2e8f0; color: #b91c1c; font-weight: 700; white-space: nowrap;">-</td>'
        html += "</tr>"
    html += "</tbody></table></div>"
    st.markdown(html, unsafe_allow_html=True)

# Dialog Handlers
if st.session_state["open_dialog"] == "edit_case_popup":
    r_idx = st.session_state.get("selected_row_idx")
    if r_idx is not None:
        view_and_edit_case_dialog(st.session_state["active_sheet"], r_idx, st.session_state.get("role", "worker"), st.session_state.get("username", ""))

# 1. Login Screen
if not st.session_state["logged_in"]:
    selected_page = st.radio("Navigation", ["🏠 Home", "🔐 Portal Login"], horizontal=True, label_visibility="collapsed")
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
    
    if admin_dash_tab == "📊 Dashboard & Quick Actions":
        r1_c1, r1_c2, r1_c3, r1_c4 = st.columns(4, gap="medium")
        with r1_c1:
            st.button("📋 New Cases", use_container_width=True)
        with r1_c2:
            st.button("🗑️ Delete Requests", use_container_width=True)
        with r1_c3:
            st.button("🔄 Sync Excel", use_container_width=True)
        with r1_c4:
            st.button("📢 Broadcast Notice", use_container_width=True)

    elif admin_dash_tab == "📂 Case Records & Search":
        render_clickable_table(df, find_display_username_column(df), find_password_column(df))

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

    worker_dash_tab = st.radio("Worker Navigation", ["📊 My Workspace"], horizontal=True, label_visibility="collapsed")
    if worker_dash_tab == "📊 My Workspace":
        render_clickable_table(user_rows, find_display_username_column(df), find_password_column(df))
