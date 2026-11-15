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
LOGO_FILES = ["pngTShirtPrintCricket.jpg", "logo.jpg", "logo.png", "logo.jpeg"]

st.set_page_config(page_title="HemanVisa Portal", page_icon="✈️", layout="wide")

# 🌟 Session State & Query Params Initialization
qp = st.query_params

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["username"] = ""
    st.session_state["role"] = ""

if not st.session_state["logged_in"] and qp.get("logged_in") == "true":
    st.session_state["logged_in"] = True
    st.session_state["username"] = qp.get("username", "")
    st.session_state["role"] = qp.get("role", "worker")

if "open_dialog" not in st.session_state:
    st.session_state["open_dialog"] = None

if "selected_row_idx" not in st.session_state:
    st.session_state["selected_row_idx"] = None

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

# 🎨 Modern Professional CSS
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
        padding: 6px 20px;
        border-bottom: 3px solid #1d6fa5;
        box-shadow: 0 4px 15px rgba(11, 60, 101, 0.15);
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        box-sizing: border-box;
    }
    .sticky-top-banner img {
        max-height: 85px;
        width: auto;
    }
    .banner-middle {
        text-align: center;
        flex: 1;
        padding: 0 15px;
    }
    .world-map-title {
        font-size: 19px;
        font-weight: 800;
        color: #0b3c65;
        letter-spacing: 1px;
        text-transform: uppercase;
        margin-bottom: 1px;
    }
    .world-map-subtitle {
        font-size: 12px;
        font-weight: 700;
        color: #d35400;
        letter-spacing: 1.5px;
    }
    .banner-info {
        display: flex;
        align-items: center;
        gap: 15px;
        flex-wrap: wrap;
        font-size: 12px;
        font-weight: 600;
        color: #0b3c65;
    }
    .banner-info a {
        color: #0b3c65;
        text-decoration: none;
    }
    .social-icons {
        display: flex;
        gap: 6px;
        align-items: center;
    }
    .social-icons img {
        height: 22px;
        width: 22px;
        border-radius: 4px;
        max-height: 22px !important;
    }

    @media (prefers-color-scheme: dark) {
        .sticky-top-banner {
            background-color: #0e1117;
            border-bottom: 3px solid #4da6ff;
        }
        .world-map-title {
            color: #ffffff;
        }
        .banner-info {
            color: #e6edf3;
        }
        .banner-info a {
            color: #e6edf3;
        }
    }

    .block-container {
        padding-top: 105px !important;
        padding-bottom: 2rem;
    }

    .sticky-control-panel {
        position: sticky;
        top: 98px;
        z-index: 99998;
        background-color: #ffffff;
        padding-top: 2px;
        padding-bottom: 6px;
        border-bottom: 2px solid #e1e4e8;
        margin-bottom: 10px;
    }
    @media (prefers-color-scheme: dark) {
        .sticky-control-panel {
            background-color: #0e1117;
            border-bottom: 2px solid #30363d;
        }
    }

    .kpi-card {
        background: linear-gradient(135deg, #0b3c65 0%, #1d6fa5 100%);
        color: white;
        padding: 10px 8px;
        border-radius: 8px;
        box-shadow: 0 3px 8px rgba(0,0,0,0.1);
        text-align: center;
        border-left: 4px solid #f39c12;
        margin-bottom: 5px;
        height: 85px;
        width: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        box-sizing: border-box;
    }
    .kpi-card.gold {
        background: linear-gradient(135deg, #d35400 0%, #f39c12 100%);
        border-left: 4px solid #0b3c65;
    }
    .kpi-card.green {
        background: linear-gradient(135deg, #117a65 0%, #1abc9c 100%);
        border-left: 4px solid #f1c40f;
    }
    .kpi-card.purple {
        background: linear-gradient(135deg, #4a235a 0%, #8e44ad 100%);
        border-left: 4px solid #f39c12;
    }
    .kpi-title {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        opacity: 0.95;
        margin-bottom: 4px;
        line-height: 1.2;
    }
    .kpi-value {
        font-size: 20px;
        font-weight: 800;
        margin: 0;
        line-height: 1.1;
    }

    .section-banner {
        background-color: #eef5fc;
        border-left: 5px solid #0b3c65;
        padding: 10px 15px;
        border-radius: 6px;
        margin: 10px 0px;
        font-weight: 700;
        color: #0b3c65;
        font-size: 18px;
    }
    @media (prefers-color-scheme: dark) {
        .section-banner {
            background-color: #1c2530;
            border-left: 5px solid #4da6ff;
            color: #e6edf3;
        }
    }

    .stButton > button, [data-testid="stDownloadButton"] > button {
        width: 100% !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        padding: 12px 15px !important;
        color: #ffffff !important;
        background: linear-gradient(135deg, #0b3c65 0%, #1d6fa5 100%) !important;
        border: none !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
        margin-top: 5px !important;
        margin-bottom: 10px !important;
    }

    div[data-testid="stExpander"] {
        width: 100% !important;
        max-width: 950px !important;
    }
    div.row-widget.stRadio > div {
        flex-direction: row;
        gap: 12px;
        background: rgba(29, 111, 165, 0.06);
        padding: 8px 14px;
        border-radius: 8px;
    }
    [data-testid="stVerticalBlock"] {
        gap: 0.3rem !important;
    }
    hr {
        margin-top: 0.8rem !important;
        margin-bottom: 0.8rem !important;
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

st.markdown(f"""
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

def find_username_column(df):
    for col in df.columns:
        if col.strip().lower() in ["username", "cases_worker", "user"]:
            return col
    return df.columns[0]

def find_password_column(df):
    for col in df.columns:
        if col.strip().lower() in ["password", "pass"]:
            return col
    return df.columns[1] if len(df.columns) > 1 else df.columns[0]

# 🌟 Case Details Edit Popup Dialog (For Both Admin & Case Worker)
@st.dialog("✏️ Edit Case Details", width="large")
def edit_single_case_dialog(sheet_name, row_idx, role, current_user, can_update=True):
    df_full = load_data(sheet_name=sheet_name)
    if df_full is None or row_idx not in df_full.index:
        st.error("Case data not found.")
        return

    row_data = df_full.loc[row_idx]
    u_col = find_username_column(df_full)
    p_col = find_password_column(df_full)

    st.info(f"📂 Sheet: **{sheet_name}** | Row #: **{row_idx + 1}**")

    if role == "worker" and not can_update:
        st.warning("⚠️ View Only Mode: Admin has disabled update permissions for your account.")
        for col in df_full.columns:
            if col not in [u_col, p_col]:
                st.text_input(col, value=str(row_data[col]), disabled=True, key=f"view_only_{col}")
        return

    with st.form("single_case_edit_form"):
        updated_values = {}
        cols_to_show = list(df_full.columns) if role == "admin" else [c for c in df_full.columns if c not in [u_col, p_col, "Updated_By", "Updated_At"]]
        
        f_col1, f_col2 = st.columns(2)
        for i, col_name in enumerate(cols_to_show):
            target_col = f_col1 if i % 2 == 0 else f_col2
            with target_col:
                updated_values[col_name] = st.text_input(f"{col_name}", value=str(row_data[col_name]))

        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.form_submit_button("💾 Save & Update Case", type="primary", use_container_width=True):
            try:
                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                changed = False
                for col_name, new_val in updated_values.items():
                    old_val = str(df_full.at[row_idx, col_name]).strip()
                    if old_val != str(new_val).strip():
                        changed = True
                    df_full.at[row_idx, col_name] = str(new_val).strip()

                if changed and role == "worker":
                    df_full.at[row_idx, "Updated_By"] = current_user
                    df_full.at[row_idx, "Updated_At"] = current_time

                save_data(df_full, sheet_name=sheet_name)
                st.session_state["selected_row_idx"] = None
                st.success("✅ Case details updated successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error updating case: {e}")

# 🌟 Other Modal Dialogs
@st.dialog("👥 Users, Passwords & Permissions (All Sheets)", width="large")
def permissions_dialog():
    st.info("💡 Modify permissions for each case worker, then click the **💾 Save** button:")
    sheets = get_sheet_names()
    all_users_dict = {}
    for s in sheets:
        df_s = load_data(sheet_name=s)
        if df_s is not None and not df_s.empty:
            u_col = find_username_column(df_s)
            p_col = find_password_column(df_s)
            if u_col in df_s.columns and p_col in df_s.columns:
                for _, r in df_s.iterrows():
                    u_name = str(r[u_col]).strip()
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
            
            col_u, col_p, col_upd, col_add, col_sch, col_dl, col_sync, col_btn = st.columns([1.4, 1.4, 0.85, 0.85, 0.85, 0.85, 0.85, 1.0])
            
            with col_u:
                st.text_input("Username", value=u_name, disabled=True, key=f"dlg_u_{idx}", label_visibility="collapsed")
            with col_p:
                new_pwd_val = st.text_input("Password", value=curr_pwd, type="default", key=f"dlg_p_{idx}", label_visibility="collapsed")
            with col_upd:
                new_upd_val = st.checkbox("Update", value=w_perm.get("can_update", True), key=f"dlg_upd_{idx}")
            with col_add:
                new_add_val = st.checkbox("Add", value=w_perm.get("can_add", True), key=f"dlg_add_{idx}")
            with col_sch:
                new_sch_val = st.checkbox("Search", value=w_perm.get("can_search", True), key=f"dlg_sch_{idx}")
            with col_dl:
                new_dl_val = st.checkbox("Download", value=w_perm.get("can_download", False), key=f"dlg_dl_{idx}")
            with col_sync:
                new_sync_val = st.checkbox("Sync", value=w_perm.get("can_sync", False), key=f"dlg_sync_{idx}")
            with col_btn:
                if st.button("💾 Save", key=f"dlg_btn_{idx}", type="primary", use_container_width=True):
                    try:
                        all_sheets_data = {}
                        xls = pd.ExcelFile(EXCEL_FILE)
                        for s in xls.sheet_names:
                            df_s = pd.read_excel(EXCEL_FILE, sheet_name=s, dtype=str)
                            uc = find_username_column(df_s)
                            pc = find_password_column(df_s)
                            if uc in df_s.columns and pc in df_s.columns:
                                mask = df_s[uc].astype(str).str.strip().str.lower() == u_name.lower()
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
                        st.success(f"✅ Details updated for '{u_name}'!")
                    except Exception as e:
                        st.error(f"❌ Error: {e}")
            st.write("")

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
    st.info("💡 మీ లోకల్ సిస్టమ్‌లో సవరించిన ఎక్సెల్ ఫైల్‌ను ఇక్కడ అప్‌లోడ్ చేస్తే, అది వెంటనే క్లౌడ్‌లో అప్‌డేట్ అవుతుంది.")
    uploaded_excel = st.file_uploader("Upload Updated Excel File (.xlsx)", type=["xlsx"])
    
    if uploaded_excel is not None:
        if st.button("🚀 Confirm & Sync to Cloud", type="primary", use_container_width=True):
            try:
                with open(EXCEL_FILE, "wb") as f:
                    f.write(uploaded_excel.getbuffer())
                st.success("✅ క్లౌడ్ డేటా విజయవంతంగా అప్‌డేట్ చేయబడింది!")
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
            uc = find_username_column(df_chk)
            pc = find_password_column(df_chk)
            matched_r = df_chk[df_chk[uc].astype(str).str.strip().str.lower() == worker_name.lower()]
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
                        uc = find_username_column(df_s)
                        pc = find_password_column(df_s)
                        if uc in df_s.columns and pc in df_s.columns:
                            mask = df_s[uc].astype(str).str.strip().str.lower() == worker_name.lower()
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
    t_user_col = find_username_column(target_df)
    t_pass_col = find_password_column(target_df)
    
    available_usernames = sorted(list(set(target_df[t_user_col].astype(str).str.strip().tolist())))
    available_usernames = [u for u in available_usernames if u != ""]
    available_usernames.insert(0, "Not Assigned")
    if worker_name not in available_usernames and worker_name != "":
        available_usernames.insert(1, worker_name)
    
    with st.form("dlg_new_case_form", clear_on_submit=True):
        assigned_username = st.selectbox("Assign Username", options=available_usernames, index=available_usernames.index(worker_name) if worker_name in available_usernames else 0)
        st.divider()
        st.write(f"📂 **{selected_target_sheet}** Sheet Headings:")
        
        form_data = {}
        cols = [c for c in target_df.columns if c not in [t_user_col, t_pass_col, "Updated_By", "Updated_At"]]
        for col_name in cols:
            form_data[col_name] = st.text_input(f"{col_name}")
        
        if st.form_submit_button("🚀 Submit New Case", type="primary", use_container_width=True):
            try:
                new_row = {col: "" for col in target_df.columns}
                final_assigned_user = "" if assigned_username == "Not Assigned" else assigned_username.strip()
                new_row[t_user_col] = final_assigned_user
                
                current_pass = "admin123"
                matched_pass_rows = target_df[target_df[t_user_col].astype(str).str.strip().str.lower() == final_assigned_user.lower()]
                if not matched_pass_rows.empty:
                    current_pass = str(matched_pass_rows.iloc[0][t_pass_col]).strip()

                new_row[t_pass_col] = current_pass if final_assigned_user != "" else "admin123"
                new_row["Updated_By"] = worker_name
                new_row["Updated_At"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
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
        u_col = find_username_column(df_search)
        p_col = find_password_column(df_search)
        
        search_query = st.text_input("Type any branch name or person name:", placeholder="e.g., ALLAGADDA, Adoni...", key="dlg_search_input").strip()
        
        if search_query != "":
            main_search_mask = df_search.apply(lambda row: row.astype(str).str.contains(search_query, case=False, na=False).any(), axis=1)
            found_data = df_search[main_search_mask]
            
            if found_data.empty:
                st.warning(f"No details found for '{search_query}'.")
            else:
                st.success(f"Results for '{search_query}' ({len(found_data)} rows):")
                view_only_df = found_data.drop(columns=[u_col, p_col, "Updated_By", "Updated_At"], errors="ignore")
                view_only_df.index = range(1, len(view_only_df) + 1)
                st.dataframe(view_only_df, use_container_width=True)

@st.dialog("📋 Updated Cases & Editors History", width="large")
def updated_cases_history_dialog():
    sheets = get_sheet_names()
    selected_sh = st.selectbox("Select Sheet:", options=sheets, key="dlg_hist_sheet")
    df_h = load_data(sheet_name=selected_sh)
    if df_h is not None and not df_h.empty:
        if "Updated_By" in df_h.columns:
            updated_rows = df_h[df_h["Updated_By"].astype(str).str.strip() != ""]
            if updated_rows.empty:
                st.warning("No cases have been updated yet in this sheet.")
            else:
                st.success(f"Found {len(updated_rows)} updated cases in '{selected_sh}':")
                display_hist = updated_rows[["Updated_By", "Updated_At"] + [c for c in updated_rows.columns if c not in ["Updated_By", "Updated_At"]]]
                display_hist.index = range(1, len(display_hist) + 1)
                st.dataframe(display_hist, use_container_width=True)

# 🌟 Handle Dialog Execution securely via Session State
if st.session_state["open_dialog"] == "permissions":
    permissions_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "admin_password":
    admin_password_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "sync_excel":
    sync_excel_dialog()
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
elif st.session_state["open_dialog"] == "updated_history":
    updated_cases_history_dialog()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "edit_case_popup":
    r_idx = st.session_state.get("selected_row_idx")
    if r_idx is not None:
        curr_role = st.session_state.get("role", "worker")
        curr_u = st.session_state.get("username", "")
        perms = load_permissions()
        c_upd = True if curr_role == "admin" else perms.get(curr_u, {}).get("can_update", True)
        edit_single_case_dialog(st.session_state["active_sheet"], r_idx, curr_role, curr_u, c_upd)
    st.session_state["open_dialog"] = None

# 1. Login Screen
if not st.session_state["logged_in"]:
    col_l1, col_l2, col_l3 = st.columns([1, 1.4, 1])
    with col_l2:
        with st.container(border=True):
            st.markdown("<h3 style='text-align:center; color:#0b3c65;'>🔐 Portal Login</h3>", unsafe_allow_html=True)
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
                        
                        if u_in == admin_u and p_in == admin_p:
                            st.session_state["logged_in"] = True
                            st.session_state["username"] = admin_u
                            st.session_state["role"] = "admin"
                            st.query_params["logged_in"] = "true"
                            st.query_params["username"] = admin_u
                            st.query_params["role"] = "admin"
                            st.rerun()
                        else:
                            logged_in_success = False
                            matched_sheet = ""
                            for s in sheets:
                                df_check = load_data(sheet_name=s)
                                if df_check is not None and not df_check.empty:
                                    u_col = find_username_column(df_check)
                                    p_col = find_password_column(df_check)
                                    user_rows = df_check[(df_check[u_col].astype(str).str.strip().str.lower() == u_in.lower()) & 
                                                         (df_check[p_col].astype(str).str.strip() == p_in)]
                                    if not user_rows.empty:
                                        logged_in_success = True
                                        matched_sheet = s
                                        break
                            
                            if logged_in_success:
                                st.session_state["logged_in"] = True
                                st.session_state["username"] = u_in
                                st.session_state["role"] = "worker"
                                st.session_state["active_sheet"] = matched_sheet
                                st.query_params["logged_in"] = "true"
                                st.query_params["username"] = u_in
                                st.query_params["role"] = "worker"
                                st.query_params["active_sheet"] = matched_sheet
                                st.rerun()
                            else:
                                st.error("❌ Incorrect Username or Password!")

# 2. ADMIN DASHBOARD
elif st.session_state["role"] == "admin":
    sheets = get_sheet_names()
    current_sheet = st.session_state["active_sheet"]
    
    if current_sheet not in sheets:
        current_sheet = sheets[0]
        st.session_state["active_sheet"] = current_sheet

    df = load_data(sheet_name=current_sheet)
    u_col = find_username_column(df)
    p_col = find_password_column(df)
    
    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    top_col1, top_col2, top_col3 = st.columns([6, 1, 1], gap="small")
    with top_col1:
        st.markdown(f"<div class='section-banner' style='margin:0;'>👑 Admin Control Panel — {current_sheet}</div>", unsafe_allow_html=True)
    with top_col2:
        if st.button("🔄 Refresh", key="admin_refresh"):
            st.rerun()
    with top_col3:
        if st.button("🚪 Logout", key="admin_logout"):
            st.query_params.clear()
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    total_cases = len(df)
    updated_cases = len(df[df["Updated_By"].astype(str).str.strip() != ""])
    unassigned_cases = len(df[df[u_col].astype(str).str.strip() == ""])
    active_workers = df[u_col].astype(str).str.strip().replace("", pd.NA).dropna().nunique()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"<div class='kpi-card'><div class='kpi-title'>📋 Total Cases</div><div class='kpi-value'>{total_cases}</div></div>", unsafe_allow_html=True)
    with k2:
        st.markdown(f"<div class='kpi-card green'><div class='kpi-title'>✅ Updated Cases</div><div class='kpi-value'>{updated_cases}</div></div>", unsafe_allow_html=True)
    with k3:
        st.markdown(f"<div class='kpi-card gold'><div class='kpi-title'>⚠️ Unassigned Cases</div><div class='kpi-value'>{unassigned_cases}</div></div>", unsafe_allow_html=True)
    with k4:
        st.markdown(f"<div class='kpi-card purple'><div class='kpi-title'>👥 Case Workers</div><div class='kpi-value'>{active_workers}</div></div>", unsafe_allow_html=True)

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    col_btn_1, col_btn_2, col_btn_3, col_btn_4 = st.columns(4)
    with col_btn_1:
        if st.button("👥 Users, Passwords & Permissions", use_container_width=True, key="admin_perm_btn"):
            st.session_state["open_dialog"] = "permissions"
            st.rerun()
    with col_btn_2:
        if st.button("🔑 Change Admin Password", use_container_width=True, key="admin_pw_btn"):
            st.session_state["open_dialog"] = "admin_password"
            st.rerun()
    with col_btn_3:
        if st.button("🔄 Sync Local Excel", use_container_width=True, key="admin_sync_btn"):
            st.session_state["open_dialog"] = "sync_excel"
            st.rerun()
    with col_btn_4:
        if st.button("📋 View Updated Cases", use_container_width=True, key="admin_upd_hist_btn"):
            st.session_state["open_dialog"] = "updated_history"
            st.rerun()

    st.divider()
    
    st.markdown(f"<div class='section-banner'>🔍 Admin Search & Sheet Management ({current_sheet})</div>", unsafe_allow_html=True)
    col_s1, col_s2 = st.columns([2, 1])
    with col_s1:
        admin_search = st.text_input("Search for any username, case, or info:", placeholder="e.g., Anil, Kiran...", key="admin_search_input").strip()
    with col_s2:
        if sheets:
            st.write("📑 **Sheets Switch:**")
            st.radio("Select Sheet Admin", options=sheets, index=sheets.index(current_sheet) if current_sheet in sheets else 0, horizontal=True, label_visibility="collapsed", key="admin_sheet_radio", on_change=update_admin_sheet)
    
    if admin_search != "":
        admin_search_mask = df.apply(lambda row: row.astype(str).str.contains(admin_search, case=False, na=False).any(), axis=1)
        display_df = df[admin_search_mask]
        st.info(f"Results for '{admin_search}' ({len(display_df)} rows):")
    else:
        display_df = df

    st.markdown("---")
    st.subheader("✏️ Select Case to Edit")
    
    if not display_df.empty:
        row_indices = list(display_df.index)
        col_sel1, col_sel2 = st.columns([3, 1])
        with col_sel1:
            selected_admin_row = st.selectbox(
                "Choose a case row to edit:",
                options=row_indices,
                format_func=lambda idx: f"Row {idx+1} | " + " | ".join([f"{c}: {str(display_df.loc[idx, c])}" for c in display_df.columns[:2]]),
                key="admin_case_selectbox",
                label_visibility="collapsed"
            )
        with col_sel2:
            if st.button("✏️ Edit Selected Case", type="primary", use_container_width=True, key="admin_open_popup_btn"):
                st.session_state["selected_row_idx"] = selected_admin_row
                st.session_state["open_dialog"] = "edit_case_popup"
                st.rerun()

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    st.dataframe(display_df, use_container_width=True, height=380, key="admin_full_table")
                
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

# 3. CASE WORKER DASHBOARD
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
    
    user_col = find_username_column(df)
    pass_col = find_password_column(df)
    
    user_mask = df[user_col].astype(str).str.strip().str.lower() == current_user.lower()
    user_rows = df[user_mask]
        
    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    top_col1, top_col2, top_col3 = st.columns([6, 1, 1], gap="small")
    with top_col1:
        st.markdown(f"<div class='section-banner' style='margin:0;'>👑 Welcome, {current_user} — Case Worker Workspace</div>", unsafe_allow_html=True)
    with top_col2:
        if st.button("🔄 Refresh", key="worker_refresh"):
            st.rerun()
    with top_col3:
        if st.button("🚪 Logout", key="worker_logout"):
            st.query_params.clear()
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    my_total = len(user_rows)
    my_updated = len(user_rows[user_rows["Updated_By"].astype(str).str.strip() != ""])

    visa_col = None
    for c in user_rows.columns:
        if any(term in c.lower() for term in ['visa', 'category', 'type', 'service']):
            visa_col = c
            break

    visa_counts = {}
    if visa_col and not user_rows.empty:
        for val in user_rows[visa_col].astype(str):
            v_clean = val.strip().upper()
            if v_clean != "":
                visa_counts[v_clean] = visa_counts.get(v_clean, 0) + 1

    total_metrics = 2 + len(visa_counts)
    kpi_cols = st.columns(total_metrics if total_metrics > 0 else 2)

    with kpi_cols[0]:
        st.markdown(f"<div class='kpi-card'><div class='kpi-title'>📁 My Cases</div><div class='kpi-value'>{my_total}</div></div>", unsafe_allow_html=True)
    with kpi_cols[1]:
        st.markdown(f"<div class='kpi-card green'><div class='kpi-title'>✅ Updated By Me</div><div class='kpi-value'>{my_updated}</div></div>", unsafe_allow_html=True)

    colors_list = ["gold", "purple", "green", "blue"]
    for idx, (v_name, v_count) in enumerate(visa_counts.items()):
        c_class = colors_list[idx % len(colors_list)]
        with kpi_cols[2 + idx]:
            st.markdown(f"<div class='kpi-card {c_class}'><div class='kpi-title'>✈️ {v_name}</div><div class='kpi-value'>{v_count}</div></div>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

    has_sync_perm = worker_perm.get("can_sync", False)
    w_cols = st.columns(5 if has_sync_perm else 4)
    
    with w_cols[0]:
        if st.button("🔑 Change Password", use_container_width=True, key="worker_pw_btn"):
            st.session_state["open_dialog"] = "worker_password"
            st.rerun()
    with w_cols[1]:
        if worker_perm.get("can_add", True):
            if st.button("➕ Add New Case", use_container_width=True, key="worker_add_btn"):
                st.session_state["open_dialog"] = "add_case"
                st.rerun()
        else:
            st.button("🚫 No Add Permission", disabled=True, use_container_width=True, key="worker_add_disabled")
    with w_cols[2]:
        if worker_perm.get("can_search", True):
            if st.button("🔍 Search Main Data", use_container_width=True, key="worker_search_btn"):
                st.session_state["open_dialog"] = "search_data"
                st.rerun()
        else:
            st.button("🚫 No Search Permission", disabled=True, use_container_width=True, key="worker_search_disabled")
    with w_cols[3]:
        if worker_perm.get("can_download", False):
            worker_buffer = io.BytesIO()
            with pd.ExcelWriter(worker_buffer, engine='openpyxl') as writer:
                has_user_data = False
                for s in sheets:
                    df_s = load_data(sheet_name=s)
                    if df_s is not None and not df_s.empty:
                        uc_s = find_username_column(df_s)
                        u_rows_s = df_s[df_s[uc_s].astype(str).str.strip().str.lower() == current_user.lower()]
                        if not u_rows_s.empty:
                            has_user_data = True
                            u_rows_s.to_excel(writer, sheet_name=s, index=False)
                
                if not has_user_data:
                    user_rows.to_excel(writer, sheet_name=current_sheet, index=False)
            
            st.download_button(
                label="📥 Download Excel", 
                data=worker_buffer.getvalue(), 
                file_name=f"{current_user}_all_sheets_cases.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True, 
                key="worker_dl_btn"
            )
        else:
            st.button("🚫 No Download Permission", disabled=True, use_container_width=True, key="worker_dl_disabled")
            
    if has_sync_perm:
        with w_cols[4]:
            if st.button("🔄 Sync Local Excel", use_container_width=True, key="worker_sync_btn"):
                st.session_state["open_dialog"] = "sync_excel"
                st.rerun()
            
    st.divider()

    if sheets:
        st.write("📑 **Sheets Switch:**")
        st.radio("Select Worker Sheet", options=sheets, index=sheets.index(current_sheet) if current_sheet in sheets else 0, horizontal=True, label_visibility="collapsed", key="worker_sheet_radio", on_change=update_worker_sheet)
            
    st.divider()
    
    st.markdown(f"<div class='section-banner'>✏️ Your Login Cases ({current_sheet} — Total: {len(user_rows)})</div>", unsafe_allow_html=True)
    
    if not user_rows.empty:
        worker_row_indices = list(user_rows.index)
        col_wsel1, col_wsel2 = st.columns([3, 1])
        with col_wsel1:
            selected_worker_row = st.selectbox(
                "Choose your case row to edit:",
                options=worker_row_indices,
                format_func=lambda idx: f"Row {idx+1} | " + " | ".join([f"{c}: {str(user_rows.loc[idx, c])}" for c in user_rows.columns[:2] if c not in [user_col, pass_col]]),
                key="worker_case_selectbox",
                label_visibility="collapsed"
            )
        with col_wsel2:
            if st.button("✏️ Edit Selected Case", type="primary", use_container_width=True, key="worker_open_popup_btn"):
                st.session_state["selected_row_idx"] = selected_worker_row
                st.session_state["open_dialog"] = "edit_case_popup"
                st.rerun()

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    display_worker_df = user_rows.drop(columns=[user_col, pass_col, "Updated_By", "Updated_At"], errors="ignore")
    st.dataframe(display_worker_df, use_container_width=True, height=380, key="worker_full_table")
