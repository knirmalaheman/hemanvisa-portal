import streamlit as st
import pandas as pd
import os
import io
import datetime
import base64
import json
import math
import re

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

st.set_page_config(page_title="HemanVisa Portal", page_icon="✈️", layout="wide")


def get_ist_now():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)


def _get_file_mtime():
    return os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0


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
        "last_login": full_timestamp,
    }
    with open(LOGINS_FILE, "w") as f:
        json.dump(logins, f, indent=4)

    attendance = load_attendance()
    found = False
    for entry in attendance:
        if (
            entry.get("Worker", "").lower() == worker_name.strip().lower()
            and entry.get("Date") == today_str
        ):
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
            "Status": "Active",
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
        if (
            entry.get("Worker", "").lower() == worker_name.strip().lower()
            and entry.get("Date") == today_str
        ):
            entry["Logout_Time"] = time_str
            entry["Status"] = "Logged Out"
            break
    save_attendance(attendance)


def reorder_dataframe_columns(df):
    if df is None or df.empty:
        return df
    cols_to_drop = [
        c
        for c in df.columns
        if c.strip().lower() in ["ds-160 case worker", "ds160 case worker"]
    ]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop, errors="ignore")

    cols = list(df.columns)
    found_priority = []
    if "Case_ID" in cols:
        found_priority.append("Case_ID")

    desired_order_prefixes = ["case", "username", "password", "no of app", "lead"]
    for prefix in desired_order_prefixes:
        for c in cols:
            c_low = c.strip().lower().replace("_", " ")
            if (c_low == prefix or prefix in c_low) and c not in found_priority:
                found_priority.append(c)

    remaining_cols = [c for c in cols if c not in found_priority]
    return df[found_priority + remaining_cols]


@st.cache_data(ttl=600, show_spinner=False)
def load_all_sheets_cached(file_mtime):
    if os.path.exists(EXCEL_FILE):
        try:
            all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
            processed = {}
            max_num = 1000
            seen_ids = set()

            for s_name, df in all_sheets.items():
                df.columns = df.columns.str.strip()
                if "Case_ID" in df.columns:
                    for val in df["Case_ID"].fillna("").astype(str):
                        m = re.search(r"(\d+)", val.strip())
                        if m:
                            num = int(m.group(1))
                            if num > max_num:
                                max_num = num

            for s_name, df in all_sheets.items():
                df.columns = df.columns.str.strip()
                df = df.fillna("")
                if "Case_ID" not in df.columns:
                    df.insert(0, "Case_ID", "")
                for idx in df.index:
                    cid = str(df.at[idx, "Case_ID"]).strip()
                    if cid == "" or cid.lower() == "nan" or cid in seen_ids:
                        max_num += 1
                        new_cid = f"HV-{max_num}"
                        df.at[idx, "Case_ID"] = new_cid
                        seen_ids.add(new_cid)
                    else:
                        seen_ids.add(cid)

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

                if (
                    "Appointment_Entry" not in df.columns
                    and "Current Appointment" in df.columns
                ):
                    df = df.rename(
                        columns={"Current Appointment": "Appointment_Entry"}
                    )
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


def ensure_all_cases_have_unique_ids(force=False):
    if not os.path.exists(EXCEL_FILE):
        return
    mtime = _get_file_mtime()
    last_checked_key = "_last_id_sync_mtime"
    if not force and st.session_state.get(last_checked_key) == mtime:
        return
    try:
        all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
        modified = False
        max_num = 1000
        seen_ids = set()

        for s_name, df in all_sheets.items():
            df.columns = df.columns.str.strip()
            if "Case_ID" in df.columns:
                for val in df["Case_ID"].fillna("").astype(str):
                    m = re.search(r"(\d+)", val.strip())
                    if m:
                        num = int(m.group(1))
                        if num > max_num:
                            max_num = num

        for s_name, df in all_sheets.items():
            df.columns = df.columns.str.strip()
            df = df.fillna("")
            if "Case_ID" not in df.columns:
                df.insert(0, "Case_ID", "")
                modified = True
            for idx in df.index:
                cid = str(df.at[idx, "Case_ID"]).strip()
                if cid == "" or cid.lower() == "nan" or cid in seen_ids:
                    max_num += 1
                    new_cid = f"HV-{max_num}"
                    df.at[idx, "Case_ID"] = new_cid
                    seen_ids.add(new_cid)
                    modified = True
                else:
                    seen_ids.add(cid)
            all_sheets[s_name] = reorder_dataframe_columns(df)

        if modified:
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                for s, df_val in all_sheets.items():
                    df_val.to_excel(writer, sheet_name=str(s), index=False)
            st.cache_data.clear()
        st.session_state[last_checked_key] = _get_file_mtime()
    except Exception:
        pass


def get_sheet_names():
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    if all_sheets:
        return list(all_sheets.keys())
    return ["Sheet1"]


def get_default_active_sheet(sheets_list=None):
    if sheets_list is None:
        sheets_list = get_sheet_names()
    if not sheets_list:
        return "Sheet1"
    return next(
        (s for s in sheets_list if "booked" not in s.lower()), sheets_list[0]
    )


def generate_next_case_id(extra_reserved_ids=None):
    all_sheets_dict = load_all_sheets_cached(_get_file_mtime())
    max_num = 1000
    for s_name, df_s in all_sheets_dict.items():
        if df_s is not None and not df_s.empty and "Case_ID" in df_s.columns:
            for val in df_s["Case_ID"].astype(str):
                m = re.search(r"(\d+)", val.strip())
                if m:
                    num = int(m.group(1))
                    if num > max_num:
                        max_num = num
    if extra_reserved_ids:
        for cid in extra_reserved_ids:
            m = re.search(r"(\d+)", str(cid).strip())
            if m:
                num = int(m.group(1))
                if num > max_num:
                    max_num = num
    return f"HV-{max_num + 1}"


def append_rescheduled_to_active_sheet(
    source_row,
    pref_loc,
    pref_dates,
    prior_notice,
    updated_by,
    updated_at,
    assigned_worker=None,
):
    sheets_all = get_sheet_names()
    active_sheet_name = get_default_active_sheet(sheets_all)
    if not active_sheet_name:
        return None, None
    df_active = load_data(sheet_name=active_sheet_name)
    if df_active is None:
        return None, None

    new_rescheduled_id = generate_next_case_id()
    row_to_active = source_row.copy()
    row_to_active["Case_ID"] = new_rescheduled_id

    if assigned_worker is not None:
        clean_worker = "" if str(assigned_worker).strip() in ["", "Not Assigned"] else str(assigned_worker).strip()
        login_c_act = find_login_user_column(df_active)
        row_to_active[login_c_act] = clean_worker
        for col in df_active.columns:
            c_low = col.strip().lower()
            if c_low in ["case worker", "cases_worker", "worker", "assigned_to", "assigned to"]:
                row_to_active[col] = clean_worker

    row_to_active["Booking_Status"] = "Rescheduled"
    row_to_active["Preferred_Location"] = pref_loc.strip()
    row_to_active["Preferred_Dates"] = pref_dates.strip()
    row_to_active["Prior_Notice"] = prior_notice.strip()
    row_to_active["Updated_By"] = updated_by
    row_to_active["Updated_At"] = updated_at
    row_to_active["Action_Type"] = "Rescheduled Case"

    for col in df_active.columns:
        if col not in row_to_active:
            row_to_active[col] = ""

    df_active_ordered = pd.concat(
        [df_active, pd.DataFrame([row_to_active])], ignore_index=True
    )
    save_data(df_active_ordered, sheet_name=active_sheet_name)
    return new_rescheduled_id, active_sheet_name


@st.cache_data(ttl=600, show_spinner=False)
def get_worker_export_bytes(file_mtime, current_user, current_sheet):
    all_sheets = load_all_sheets_cached(file_mtime)
    worker_buffer = io.BytesIO()
    with pd.ExcelWriter(worker_buffer, engine="openpyxl") as writer:
        has_user_data = False
        for s, df_s in all_sheets.items():
            if df_s is not None and not df_s.empty:
                login_cs = find_login_user_column(df_s)
                u_rows_s = df_s[
                    df_s[login_cs].astype(str).str.strip().str.lower()
                    == current_user.lower()
                ]
                if not u_rows_s.empty:
                    has_user_data = True
                    u_rows_s = reorder_dataframe_columns(u_rows_s)
                    u_rows_s.to_excel(writer, sheet_name=s, index=False)
        if not has_user_data and current_sheet in all_sheets:
            df_cur = all_sheets[current_sheet]
            login_cs = find_login_user_column(df_cur)
            u_rows_cur = df_cur[
                df_cur[login_cs].astype(str).str.strip().str.lower()
                == current_user.lower()
            ]
            u_rows_cur = reorder_dataframe_columns(u_rows_cur)
            u_rows_cur.to_excel(writer, sheet_name=current_sheet, index=False)
    return worker_buffer.getvalue()


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

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        for s, df_val in sheets_data.items():
            df_val = reorder_dataframe_columns(df_val)
            df_val.to_excel(writer, sheet_name=str(s), index=False)

    st.cache_data.clear()
    st.session_state["_last_id_sync_mtime"] = _get_file_mtime()


@st.cache_data(show_spinner=False)
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


def find_ds160_columns(df):
    ds_cols = []
    for col in df.columns:
        c_norm = col.strip().lower().replace("-", "").replace("_", "").replace(" ", "")
        if "ds160" in c_norm or "ds 160" in col.strip().lower():
            ds_cols.append(col)
    return ds_cols


def find_password_column(df):
    for col in df.columns:
        if col.strip().lower() in ["password", "pass"]:
            return col
    return df.columns[1] if len(df.columns) > 1 else df.columns[0]


def get_target_search_columns(df):
    target_cols = set()
    if df is None or df.empty:
        return []
    for c in df.columns:
        c_low = c.strip().lower().replace("_", " ")
        if (
            "case id" in c_low
            or "username" in c_low
            or "user" in c_low
            or "case worker" in c_low
            or "worker" in c_low
            or "password" in c_low
            or c_low == "pass"
        ):
            target_cols.add(c)
    for ds_c in find_ds160_columns(df):
        target_cols.add(ds_c)
    if not target_cols:
        target_cols.add(find_display_username_column(df))
    return [c for c in df.columns if c in target_cols]


def search_username_ds160_password_worker(df, query_text):
    if df is None or df.empty or not query_text:
        return (
            pd.Series([True] * len(df), index=df.index)
            if df is not None
            else pd.Series([], dtype=bool)
        )

    q = query_text.strip().lower()
    target_cols = get_target_search_columns(df)

    mask = pd.Series([False] * len(df), index=df.index)
    for col in target_cols:
        col_mask = (
            df[col]
            .astype(str)
            .str.lower()
            .str.contains(q, case=False, na=False, regex=False)
        )
        mask = mask | col_mask
    return mask


def search_across_all_sheets_by_fields(query_text, lead_filter="All"):
    if not query_text:
        return pd.DataFrame()
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    matched_dfs = []
    for s_name, df_s in all_sheets.items():
        if df_s is None or df_s.empty:
            continue
        df_curr = reorder_dataframe_columns(df_s.copy())
        if lead_filter and lead_filter != "All":
            login_c_s = find_login_user_column(df_curr)
            lead_col_s = next(
                (c for c in df_curr.columns if "lead" in c.strip().lower()), login_c_s
            )
            if lead_col_s in df_curr.columns:
                df_curr = df_curr[
                    df_curr[lead_col_s]
                    .astype(str)
                    .str.strip()
                    .str.lower()
                    .str.contains(lead_filter.lower(), na=False)
                ]
        if df_curr.empty:
            continue
        m = search_username_ds160_password_worker(df_curr, query_text)
        res_s = df_curr[m]
        if not res_s.empty:
            res_s = res_s.copy()
            res_s["Sheet_Name"] = s_name
            matched_dfs.append(res_s)
    if not matched_dfs:
        return pd.DataFrame()
    return pd.concat(matched_dfs, ignore_index=True)


def search_main_data_any(df, query_text):
    if df is None or df.empty or not query_text:
        return (
            pd.Series([True] * len(df), index=df.index)
            if df is not None
            else pd.Series([], dtype=bool)
        )

    q = query_text.strip().lower()
    p_col = find_password_column(df)
    excluded_cols = {
        p_col,
        "Appointment_PDF",
        "Updated_By",
        "Updated_At",
        "Action_Type",
        "Delete_Reason",
    }
    search_cols = [
        c
        for c in df.columns
        if c not in excluded_cols and "password" not in c.strip().lower()
    ]
    if not search_cols:
        search_cols = [c for c in df.columns if c != p_col]

    mask = pd.Series([False] * len(df), index=df.index)
    for col in search_cols:
        col_mask = (
            df[col]
            .astype(str)
            .str.lower()
            .str.contains(q, case=False, na=False, regex=False)
        )
        mask = mask | col_mask
    return mask


def find_visa_type_column(df):
    for c in df.columns:
        if any(term in c.lower() for term in ["visa", "category", "type", "service"]):
            return c
    return None


def get_all_workers_list():
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    all_workers = set()
    for _, df_s in all_sheets.items():
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
    return all(
        norm_perms.get(w.lower(), {}).get("block_login", False) for w in workers
    )


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
        matched_key = next(
            (k for k in current_perms if k.lower() == w.lower()), w
        )
        if matched_key not in current_perms:
            current_perms[matched_key] = {
                "can_update": True,
                "can_add": True,
                "can_search": True,
                "can_download": False,
                "can_sync": False,
                "block_login": block_status,
            }
        else:
            current_perms[matched_key]["block_login"] = block_status
        if block_status:
            record_worker_logout(w)
    save_permissions(current_perms)


def is_not_assigned_value(val):
    s = str(val).strip().lower()
    return s in ["", "nan", "none", "not assigned", "unassigned", "-"]


def is_today_updated_row(row, today_str):
    upd_at = str(row.get("Updated_At", "")).strip()
    upd_by = str(row.get("Updated_By", "")).strip()
    if today_str in upd_at and upd_by != "":
        return True
    return False


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
        if "active_sheet" in qp and qp["active_sheet"] in get_sheet_names():
            st.session_state["active_sheet"] = qp["active_sheet"]
        st.session_state["selected_row_idx"] = r_val
        st.session_state["open_dialog"] = "edit_case_popup"
        st.session_state["is_editing_popup"] = False
        st.session_state["is_delete_request_mode"] = False
        st.session_state["is_reschedule_mode"] = False
        if "edit_row" in st.query_params:
            del st.query_params["edit_row"]
    except Exception:
        if "edit_row" in st.query_params:
            del st.query_params["edit_row"]
elif not st.session_state["logged_in"] and "edit_row" in qp:
    qp.clear()

sheets_init = get_sheet_names()
default_s_init = get_default_active_sheet(sheets_init)

if "active_sheet" not in st.session_state:
    st.session_state["active_sheet"] = default_s_init
    if st.session_state.get("logged_in"):
        st.query_params["active_sheet"] = default_s_init

if "dash_sheet" not in st.session_state:
    st.session_state["dash_sheet"] = default_s_init

if "records_sheet" not in st.session_state:
    st.session_state["records_sheet"] = default_s_init

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
        padding: 8px 24px;
        border-bottom: 3px solid #d4af37;
        box-shadow: 0 8px 25px rgba(7, 15, 36, 0.55);
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
        max-height: 54px;
        width: auto;
        border-radius: 8px;
        background: #ffffff;
        padding: 3px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    }
    .banner-middle {
        text-align: center;
        flex: 1;
        padding: 0 12px;
    }
    .world-map-title {
        font-size: 18px;
        font-weight: 800;
        color: #ffffff;
        letter-spacing: 1.4px;
        text-transform: uppercase;
        margin-bottom: 2px;
    }
    .world-map-subtitle {
        font-size: 11.5px;
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
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.2);
        padding: 5px 12px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        gap: 6px;
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
        padding: 5px 10px;
        border-radius: 8px;
    }
    .social-icons img {
        height: 18px;
        width: 18px;
        border-radius: 4px;
        max-height: 18px !important;
        background: transparent !important;
        padding: 0 !important;
    }
    .block-container {
        padding-top: 95px !important;
        padding-bottom: 2rem;
        max-width: 1400px !important;
    }
    .hero-section {
        background: linear-gradient(135deg, #0f1c3f 0%, #172554 50%, #0a1128 100%);
        color: white;
        padding: 48px 30px;
        border-radius: 16px;
        text-align: center;
        margin-bottom: 24px;
        box-shadow: 0 12px 35px rgba(15, 28, 63, 0.35);
    }
    .hero-title {
        font-size: 34px;
        font-weight: 800;
        margin-bottom: 12px;
    }
    .hero-subtitle {
        font-size: 15.5px;
        opacity: 0.95;
        line-height: 1.6;
        max-width: 800px;
        margin: 0 auto;
    }
    .corporate-card {
        background: #ffffff;
        padding: 22px;
        border-radius: 14px;
        box-shadow: 0 6px 20px rgba(15, 23, 42, 0.06);
        border: 1px solid #e2e8f0;
        border-top: 4px solid #172554;
        height: 100%;
        margin-bottom: 12px;
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
        padding: 10px 16px;
        border-radius: 10px;
        font-weight: 700;
        font-size: 13.5px;
        margin-bottom: 14px !important;
        overflow: hidden;
        white-space: nowrap;
        border-left: 5px solid #d4af37;
    }
    .sticky-control-panel {
        position: sticky;
        top: 78px;
        z-index: 99990;
        background-color: #ffffff;
        padding: 6px 0 8px 0;
        border-bottom: 2px solid #e2e8f0;
        margin-bottom: 10px;
    }
    .workspace-header-bar {
        background: linear-gradient(135deg, #0f1c3f 0%, #1e293b 100%);
        color: #ffffff;
        padding: 11px 18px;
        border-radius: 11px;
        border-left: 5px solid #d4af37;
        font-size: 15.5px;
        font-weight: 800;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .workspace-sheet-pill {
        background: rgba(212, 175, 55, 0.18);
        border: 1px solid #d4af37;
        color: #fde047;
        padding: 3px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 700;
    }
    .kpi-card {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        color: white;
        padding: 14px 10px;
        border-radius: 13px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.14);
        text-align: center;
        border-bottom: 4px solid #64748b;
        margin-bottom: 6px;
        height: 92px;
        width: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
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
    .kpi-card.teal {
        background: linear-gradient(135deg, #115e59 0%, #0f766e 100%);
        border-bottom: 4px solid #14b8a6;
    }
    .kpi-card.rose {
        background: linear-gradient(135deg, #881337 0%, #be123c 100%);
        border-bottom: 4px solid #fb7185;
    }
    .kpi-title {
        font-size: 10.5px;
        text-transform: uppercase;
        letter-spacing: 0.9px;
        opacity: 0.92;
        margin-bottom: 5px;
        font-weight: 700;
        color: #f1f5f9;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 100%;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: 800;
        margin: 0;
        line-height: 1.1;
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
        background: linear-gradient(135deg, #334155 0%, #475569 100%) !important;
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
        padding: 11px 16px;
        border-radius: 8px;
        margin: 8px 0px;
        font-weight: 800;
        color: #1e293b;
        font-size: 15.5px;
        border: 1px solid #e2e8f0;
    }
    .action-group-title {
        font-size: 12.5px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: #334155;
        margin: 8px 0 4px 2px;
    }
    .stButton > button, [data-testid="stDownloadButton"] > button {
        width: 100% !important;
        min-height: 40px !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-size: 13px !important;
        padding: 6px 10px !important;
        color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.2) !important;
        box-shadow: 0 3px 8px rgba(15, 23, 42, 0.1);
        margin-top: 2px !important;
        margin-bottom: 6px !important;
        white-space: nowrap !important;
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%) !important;
    }
    .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
        opacity: 0.94;
    }
    .username-hyperlink {
        color: #1e293b !important;
        text-decoration: underline !important;
        font-weight: 700 !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] {
        display: flex !important;
        flex-direction: row !important;
        flex-wrap: wrap !important;
        gap: 8px !important;
        background: #f1f5f9 !important;
        padding: 6px !important;
        border-radius: 12px !important;
        border: 1px solid #cbd5e1 !important;
        width: fit-content !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label {
        background: transparent !important;
        padding: 7px 16px !important;
        border-radius: 9px !important;
        margin: 0 !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label > div:first-child,
    div[data-testid="stRadio"] > div[role="radiogroup"] label > span:first-child,
    div[data-testid="stRadio"] > div[role="radiogroup"] label input[type="radio"] {
        display: none !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label p {
        font-size: 13.5px !important;
        font-weight: 700 !important;
        color: #475569 !important;
        margin: 0 !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input:checked) {
        background: linear-gradient(135deg, #0f1c3f 0%, #1e3a8a 100%) !important;
        border-bottom: 2.5px solid #d4af37 !important;
    }
    div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input:checked) p {
        color: #ffffff !important;
    }
    .worker-hero-banner {
        background: linear-gradient(135deg, #0f1c3f 0%, #1e3a8a 45%, #065f46 100%);
        color: #ffffff;
        padding: 13px 18px;
        border-radius: 12px;
        border-left: 5px solid #fbbf24;
        border-bottom: 2px solid rgba(251, 191, 36, 0.45);
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.16);
        display: flex;
        align-items: center;
        justify-content: space-between;
        flex-wrap: wrap;
        gap: 10px;
        margin-bottom: 8px;
    }
    .worker-hero-title {
        font-size: 16px;
        font-weight: 800;
        letter-spacing: 0.4px;
    }
    .worker-pill-group {
        display: flex;
        gap: 8px;
        flex-wrap: wrap;
        align-items: center;
    }
    .worker-pill-gold {
        background: rgba(251, 191, 36, 0.2);
        border: 1px solid #fbbf24;
        color: #fef08a;
        padding: 4px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
    }
    .worker-pill-cyan {
        background: rgba(56, 189, 248, 0.2);
        border: 1px solid #38bdf8;
        color: #bae6fd;
        padding: 4px 12px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 700;
    }
    [data-testid="stVerticalBlock"] {
        gap: 0.35rem !important;
    }
    hr {
        margin-top: 0.75rem !important;
        margin-bottom: 0.75rem !important;
    }
</style>
""", unsafe_allow_html=True)

found_logo = next((f for f in LOGO_FILES if os.path.exists(f)), None)
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


def render_clickable_table(
    df,
    display_c,
    p_col,
    table_key="main",
    sheet_for_link="",
    page_size=100,
    allow_edit=True,
):
    if df.empty:
        st.info("No cases found.")
        return

    df = reorder_dataframe_columns(df)
    total_rows = len(df)
    total_pages = max(1, math.ceil(total_rows / page_size))
    page_state_key = f"page_num_{table_key}"

    if page_state_key not in st.session_state:
        st.session_state[page_state_key] = 1
    if st.session_state[page_state_key] > total_pages:
        st.session_state[page_state_key] = total_pages
    if st.session_state[page_state_key] < 1:
        st.session_state[page_state_key] = 1

    curr_page = st.session_state[page_state_key]
    start_idx = (curr_page - 1) * page_size
    end_idx = min(start_idx + page_size, total_rows)
    page_df = df.iloc[start_idx:end_idx]

    if total_pages > 1:
        p_c1, p_c2, p_c3, p_c4 = st.columns([1.2, 2.6, 1.2, 1.2], gap="small")
        with p_c1:
            if st.button(
                "⬅️ Previous",
                key=f"prev_{table_key}",
                disabled=(curr_page <= 1),
                use_container_width=True,
            ):
                st.session_state[page_state_key] = curr_page - 1
                st.rerun()
        with p_c2:
            st.markdown(
                f"<div style='text-align:center; padding-top:10px; font-weight:700; color:#1e293b; font-size:13.5px;'>"
                f"Page {curr_page} of {total_pages} &nbsp;|&nbsp; Showing {start_idx + 1}–{end_idx} of {total_rows} cases</div>",
                unsafe_allow_html=True,
            )
        with p_c3:
            if st.button(
                "Next ➡️",
                key=f"next_{table_key}",
                disabled=(curr_page >= total_pages),
                use_container_width=True,
            ):
                st.session_state[page_state_key] = curr_page + 1
                st.rerun()
        with p_c4:
            jump_page = st.number_input(
                "Page",
                min_value=1,
                max_value=total_pages,
                value=curr_page,
                label_visibility="collapsed",
                key=f"jump_{table_key}",
            )
            if jump_page != curr_page:
                st.session_state[page_state_key] = int(jump_page)
                st.rerun()

    cols_to_hide = (
        ["Appointment_PDF"] if allow_edit else [p_col, "Appointment_PDF"]
    )
    display_cols = [c for c in page_df.columns if c not in cols_to_hide]

    curr_logged = qp.get(
        "logged_in", "true" if st.session_state.get("logged_in") else "false"
    )
    curr_user = qp.get("username", st.session_state.get("username", ""))
    curr_role = qp.get("role", st.session_state.get("role", ""))
    curr_sheet = (
        sheet_for_link
        or qp.get("active_sheet", st.session_state.get("active_sheet", ""))
    )

    html_parts = [
        '<div style="max-height: 500px; overflow-y: auto; overflow-x: auto; border: 1px solid #cbd5e1; border-radius: 12px; margin-bottom: 14px;">',
        '<table style="width:100%; border-collapse: collapse; font-size: 13px; font-family: sans-serif;">',
        '<thead><tr style="background: linear-gradient(135deg, #1e293b 0%, #334155 100%); color: white; position: sticky; top: 0; z-index: 1;">',
        '<th style="padding: 11px 10px; border: 1px solid #334155; text-align: center;">#</th>',
    ]
    for col in display_cols:
        html_parts.append(
            f'<th style="padding: 11px 10px; border: 1px solid #334155; text-align: left; white-space: nowrap;">{col}</th>'
        )
    html_parts.append(
        '<th style="padding: 11px 10px; border: 1px solid #334155; text-align: left; white-space: nowrap;">Status</th>'
    )
    html_parts.append("</tr></thead><tbody>")

    base_link = f"?logged_in={curr_logged}&username={curr_user}&role={curr_role}&active_sheet={curr_sheet}&edit_row="
    indices = page_df.index.tolist()
    act_col_vals = (
        page_df["Action_Type"].astype(str).str.strip().str.lower().tolist()
        if "Action_Type" in page_df.columns
        else [""] * len(page_df)
    )
    booking_status_vals = (
        page_df["Booking_Status"].astype(str).str.strip().tolist()
        if "Booking_Status" in page_df.columns
        else ["Active"] * len(page_df)
    )
    col_data = {col: page_df[col].astype(str).tolist() for col in display_cols}

    for i, idx in enumerate(indices):
        serial_no = start_idx + i + 1
        act_t = act_col_vals[i]
        b_status = booking_status_vals[i]
        row_bg = (
            "#fee2e2"
            if act_t == "delete rejected"
            else ("#f8fafc" if i % 2 == 0 else "#ffffff")
        )
        html_parts.append(f'<tr style="background-color: {row_bg};">')
        html_parts.append(
            f'<td style="padding: 9px 10px; border: 1px solid #e2e8f0; text-align: center; color: #64748b; font-weight: 700;">{serial_no}</td>'
        )

        for col in display_cols:
            val = col_data[col][i]
            if col == "Case_ID":
                cell_content = f'<span style="font-weight:800; color:#1e3a8a; background:#eff6ff; padding:2px 8px; border-radius:6px; border:1px solid #bfdbfe;">{val if val != "" else "-"}</span>'
            elif col == display_c:
                display_val = (
                    val.strip() if val.strip() != "" else "Not Assigned"
                )
                if allow_edit:
                    cell_content = f'<a href="{base_link}{idx}" target="_self" class="username-hyperlink">{display_val}</a>'
                else:
                    cell_content = f'<span style="font-weight:700; color:#1e293b;">{display_val}</span>'
            else:
                cell_content = val if val != "" else "-"
            html_parts.append(
                f'<td style="padding: 9px 10px; border: 1px solid #e2e8f0; color: #1e293b; white-space: nowrap;">{cell_content}</td>'
            )

        if act_t == "delete rejected":
            status_display = "Rejected"
            status_color = "#b91c1c"
        elif b_status.lower() == "rescheduled":
            status_display = "Rescheduled"
            status_color = "#d97706"
        elif b_status.lower() == "booked":
            status_display = "Booked"
            status_color = "#059669"
        else:
            status_display = "Active"
            status_color = "#1e293b"

        html_parts.append(
            f'<td style="padding: 9px 10px; border: 1px solid #e2e8f0; color: {status_color}; font-weight: 700; white-space: nowrap;">{status_display}</td>'
        )
        html_parts.append("</tr>")

    html_parts.append("</tbody></table></div>")
    st.markdown("".join(html_parts), unsafe_allow_html=True)


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
    case_id_val = str(row_data.get("Case_ID", "N/A")).strip()

    st.info(
        f"🆔 Case ID: **{case_id_val}** | 📂 Sheet: **{sheet_name}** | Row #: **{row_idx + 1}** | Case Username: **{row_data[display_c]}**"
    )

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

    if (
        not st.session_state["is_editing_popup"]
        and not st.session_state["is_delete_request_mode"]
        and not st.session_state["is_reschedule_mode"]
    ):
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
            st.markdown(
                f"🔄 **Reschedule Info:** Location: `{pref_loc}` | Dates: `{pref_dates}` | Prior Notice: `{prior_notice}`"
            )

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
                use_container_width=True,
            )
        else:
            st.info("ℹ️ No Appointment PDF uploaded yet.")

        st.markdown("---")
        if role == "worker" and not can_update:
            st.warning(
                "⚠️ View Only Mode: Admin has disabled update permissions for your account."
            )
        else:
            if is_booked_sheet:
                col_act1, col_act2, col_act3 = st.columns(3)
                with col_act1:
                    if st.button(
                        "✏️ Edit Case",
                        type="primary",
                        use_container_width=True,
                        key="btn_trigger_edit",
                    ):
                        st.session_state["is_editing_popup"] = True
                        st.rerun()
                with col_act2:
                    if st.button(
                        "🔄 Reschedule",
                        type="secondary",
                        use_container_width=True,
                        key="btn_trigger_reschedule",
                    ):
                        st.session_state["is_reschedule_mode"] = True
                        st.rerun()
                with col_act3:
                    if role == "admin":
                        if st.button(
                            "🗑️ Delete",
                            type="secondary",
                            use_container_width=True,
                            key="btn_trigger_delete",
                        ):
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
                        if st.button(
                            "🗑️ Request Delete",
                            type="secondary",
                            use_container_width=True,
                            key="btn_trigger_del_req_mode",
                        ):
                            st.session_state["is_delete_request_mode"] = True
                            st.rerun()
            else:
                col_act1, col_act2 = st.columns(2)
                with col_act1:
                    if st.button(
                        "✏️ Edit Case & Update Booking",
                        type="primary",
                        use_container_width=True,
                        key="btn_trigger_edit",
                    ):
                        st.session_state["is_editing_popup"] = True
                        st.rerun()
                with col_act2:
                    if role == "admin":
                        if st.button(
                            "🗑️ Delete Case",
                            type="secondary",
                            use_container_width=True,
                            key="btn_trigger_delete",
                        ):
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
                        if st.button(
                            "🗑️ Request Delete",
                            type="secondary",
                            use_container_width=True,
                            key="btn_trigger_del_req_mode",
                        ):
                            st.session_state["is_delete_request_mode"] = True
                            st.rerun()

    elif st.session_state["is_reschedule_mode"]:
        st.markdown(
            f"### 🔄 Reschedule Appointment Case (Booked ID stays: `{case_id_val}`)"
        )
        all_workers_res = get_all_workers_list()
        worker_opts_res = ["Not Assigned"] + [
            w for w in all_workers_res if w.lower() != "not assigned"
        ]
        curr_w_val = str(row_data.get(login_c, "")).strip()
        def_w_idx = (
            worker_opts_res.index(curr_w_val)
            if curr_w_val in worker_opts_res
            else (
                worker_opts_res.index(current_user)
                if current_user in worker_opts_res
                else 0
            )
        )
        with st.form(f"reschedule_form_{row_idx}"):
            assign_worker_res = st.selectbox(
                "👷 Allocate Rescheduled Case to Case Worker (Active Sheet)",
                options=worker_opts_res,
                index=def_w_idx,
            )
            pref_loc_input = st.text_input(
                "Preferred Location",
                value="",
                placeholder="Enter preferred location...",
            )
            pref_dates_input = st.text_input(
                "Preferred Dates",
                value="",
                placeholder="Enter preferred dates...",
            )
            prior_notice_input = st.text_input(
                "Prior Notice",
                value="",
                placeholder="Enter prior notice details...",
            )

            col_res_save, col_res_back = st.columns(2)
            with col_res_save:
                submitted_res = st.form_submit_button(
                    "💾 Save Reschedule", type="primary", use_container_width=True
                )
            with col_res_back:
                cancelled_res = st.form_submit_button(
                    "↩️ Back", use_container_width=True
                )

            if submitted_res:
                if (
                    not pref_loc_input.strip()
                    or not pref_dates_input.strip()
                    or not prior_notice_input.strip()
                ):
                    st.warning(
                        "⚠️ All 3 fields (Preferred Location, Preferred Dates, and Prior Notice) must be entered to save Reschedule!"
                    )
                else:
                    try:
                        ts_now = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        booked_snapshot_row = df_full.loc[row_idx].copy()

                        df_full.at[row_idx, "Case_ID"] = case_id_val
                        df_full.at[row_idx, "Preferred_Location"] = (
                            pref_loc_input.strip()
                        )
                        df_full.at[row_idx, "Preferred_Dates"] = (
                            pref_dates_input.strip()
                        )
                        df_full.at[row_idx, "Prior_Notice"] = (
                            prior_notice_input.strip()
                        )
                        df_full.at[row_idx, "Booking_Status"] = (
                            "Booked" if is_booked_sheet else "Rescheduled"
                        )
                        df_full.at[row_idx, "Updated_By"] = current_user
                        df_full.at[row_idx, "Updated_At"] = ts_now
                        df_full.at[row_idx, "Action_Type"] = "Rescheduled Case"

                        save_data(df_full, sheet_name=sheet_name)

                        new_rescheduled_id = None
                        target_act_sheet = None
                        if is_booked_sheet:
                            new_rescheduled_id, target_act_sheet = (
                                append_rescheduled_to_active_sheet(
                                    source_row=booked_snapshot_row,
                                    pref_loc=pref_loc_input,
                                    pref_dates=pref_dates_input,
                                    prior_notice=prior_notice_input,
                                    updated_by=current_user,
                                    updated_at=ts_now,
                                    assigned_worker=assign_worker_res,
                                )
                            )

                        st.session_state["is_reschedule_mode"] = False
                        st.session_state["selected_row_idx"] = None
                        st.session_state["open_dialog"] = None
                        if new_rescheduled_id:
                            st.success(
                                f"✅ Booked sheet keeps ID {case_id_val} (Status: Booked) & {target_act_sheet} created new ID {new_rescheduled_id} assigned to {assign_worker_res}!"
                            )
                        else:
                            st.success("✅ Case successfully rescheduled!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error saving reschedule: {e}")

            if cancelled_res:
                st.session_state["is_reschedule_mode"] = False
                st.rerun()

    elif st.session_state["is_delete_request_mode"]:
        st.markdown("### 🗑️ Request Delete Case")
        with st.form(f"delete_reason_form_{row_idx}"):
            del_reason = st.text_area(
                "✍️ Reason for deletion (Required):",
                placeholder="Type your reason here...",
            )
            col_sub_req, col_back_req = st.columns(2)
            with col_sub_req:
                submitted_req = st.form_submit_button(
                    "🚀 Submit Delete Request",
                    type="primary",
                    use_container_width=True,
                )
            with col_back_req:
                cancelled_req = st.form_submit_button(
                    "↩️ Back", use_container_width=True
                )

            if submitted_req:
                if del_reason.strip() == "":
                    st.warning("⚠️ Please provide a reason for deletion.")
                else:
                    try:
                        df_full.at[row_idx, "Updated_By"] = current_user
                        df_full.at[row_idx, "Updated_At"] = (
                            get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        )
                        df_full.at[row_idx, "Action_Type"] = "Delete Request"
                        df_full.at[row_idx, "Delete_Reason"] = del_reason.strip()
                        save_data(df_full, sheet_name=sheet_name)
                        st.session_state["is_delete_request_mode"] = False
                        st.session_state["selected_row_idx"] = None
                        st.session_state["open_dialog"] = None
                        st.success(
                            "✅ Delete request sent to Admin with reason successfully!"
                        )
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
            elif "Not Assigned" not in all_workers_list:
                all_workers_list.insert(0, "Not Assigned")

            cols_to_show = (
                list(df_full.columns)
                if role == "admin"
                else [
                    c
                    for c in df_full.columns
                    if c
                    not in [
                        p_col,
                        "Updated_By",
                        "Updated_At",
                        "Delete_Reason",
                        "Appointment_PDF",
                        "Appointment_Entry",
                        "Booking_Status",
                        "Preferred_Location",
                        "Preferred_Dates",
                        "Prior_Notice",
                    ]
                ]
            )

            f_col1, f_col2 = st.columns(2)

            for i, col_name in enumerate(cols_to_show):
                target_col = f_col1 if i % 2 == 0 else f_col2
                with target_col:
                    if col_name == "Case_ID":
                        st.text_input(
                            "Case_ID (Locked / Non-Editable)",
                            value=case_id_val,
                            disabled=True,
                            key=f"popup_locked_cid_{row_idx}",
                        )
                    elif col_name == display_c:
                        updated_values[display_c] = st.text_input(
                            f"{col_name} (Case Username)",
                            value=str(row_data[col_name]),
                            key=f"popup_field_{row_idx}_{col_name}",
                        )
                    elif col_name == login_c:
                        curr_worker_val = str(row_data.get(login_c, "")).strip()
                        w_idx = (
                            all_workers_list.index(curr_worker_val)
                            if curr_worker_val in all_workers_list
                            else 0
                        )
                        selected_worker = st.selectbox(
                            f"{col_name} (Case Worker)",
                            options=all_workers_list,
                            index=w_idx,
                            key=f"popup_worker_select_{row_idx}",
                        )
                        updated_values[login_c] = selected_worker
                    else:
                        c_lower = col_name.lower()
                        if "date" in c_lower and "prefer" not in c_lower:
                            try:
                                parsed_d = datetime.datetime.strptime(
                                    str(row_data[col_name]).split()[0],
                                    "%Y-%m-%d",
                                ).date()
                            except Exception:
                                parsed_d = datetime.date.today()
                            sel_d = st.date_input(
                                f"{col_name}",
                                value=parsed_d,
                                key=f"popup_date_{row_idx}_{col_name}",
                            )
                            updated_values[col_name] = sel_d.strftime("%Y-%m-%d")
                        else:
                            updated_values[col_name] = st.text_input(
                                f"{col_name}",
                                value=str(row_data[col_name]),
                                key=f"popup_field_{row_idx}_{col_name}",
                            )

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            curr_bs = str(row_data.get("Booking_Status", "Active")).strip()

            opts = ["Active", "Booked", "Rescheduled"]
            b_idx = opts.index(curr_bs) if curr_bs in opts else 0
            booking_status_choice = st.selectbox(
                "📌 Select Status",
                options=opts,
                index=b_idx,
                key=f"popup_bs_{row_idx}",
            )

            app_entry_val = ""
            uploaded_pdf_file = None
            popup_pref_loc = ""
            popup_pref_dates = ""
            popup_prior_notice = ""

            if booking_status_choice == "Booked":
                app_entry_val = st.text_input(
                    "📝 Current Appointment Entry",
                    value="",
                    placeholder="Enter appointment entry details...",
                    key=f"popup_app_entry_{row_idx}",
                )
                uploaded_pdf_file = st.file_uploader(
                    "📎 Upload Appointment PDF (Required)",
                    type=["pdf"],
                    key=f"popup_pdf_up_{row_idx}",
                )
            elif booking_status_choice == "Rescheduled":
                popup_pref_loc = st.text_input(
                    "Preferred Location",
                    value="",
                    placeholder="Enter preferred location...",
                    key=f"popup_pref_loc_{row_idx}",
                )
                popup_pref_dates = st.text_input(
                    "Preferred Dates",
                    value="",
                    placeholder="Enter preferred dates...",
                    key=f"popup_pref_dates_{row_idx}",
                )
                popup_prior_notice = st.text_input(
                    "Prior Notice",
                    value="",
                    placeholder="Enter prior notice details...",
                    key=f"popup_prior_notice_{row_idx}",
                )

            st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
            col_save, col_back = st.columns(2)
            with col_save:
                submitted = st.form_submit_button(
                    "💾 Save & Update", type="primary", use_container_width=True
                )
            with col_back:
                cancelled = st.form_submit_button(
                    "↩️ Cancel", use_container_width=True
                )

            if submitted:
                try:
                    if booking_status_choice == "Rescheduled" and (
                        not popup_pref_loc.strip()
                        or not popup_pref_dates.strip()
                        or not popup_prior_notice.strip()
                    ):
                        st.warning(
                            "⚠️ All 3 fields (Preferred Location, Preferred Dates, and Prior Notice) must be entered to save Rescheduled status!"
                        )
                    elif booking_status_choice == "Booked" and (
                        not app_entry_val.strip() or not uploaded_pdf_file
                    ):
                        st.warning(
                            "⚠️ Both Current Appointment Entry and Appointment PDF must be provided when selecting Booked status!"
                        )
                    else:
                        current_time = get_ist_now()
                        current_time_str = current_time.strftime(
                            "%Y-%m-%d %I:%M:%S %p"
                        )
                        today_date_str = current_time.strftime("%Y-%m-%d")

                        changed = False

                        for col_name, new_val in updated_values.items():
                            if col_name == "Case_ID":
                                continue
                            old_val = str(df_full.at[row_idx, col_name]).strip()
                            if old_val != str(new_val).strip():
                                changed = True
                            df_full.at[row_idx, col_name] = str(new_val).strip()

                        df_full.at[row_idx, "Case_ID"] = case_id_val

                        if (
                            str(df_full.at[row_idx, "Booking_Status"])
                            != booking_status_choice
                        ):
                            df_full.at[row_idx, "Booking_Status"] = (
                                booking_status_choice
                            )
                            changed = True

                        if booking_status_choice == "Booked":
                            df_full.at[row_idx, "Appointment_Entry"] = (
                                app_entry_val.strip()
                            )
                            changed = True

                            if uploaded_pdf_file is not None:
                                date_folder = os.path.join(
                                    PDF_BASE_DIR, today_date_str
                                )
                                if not os.path.exists(date_folder):
                                    os.makedirs(date_folder)

                                safe_uname = (
                                    str(
                                        df_full.at[row_idx, display_c]
                                        if display_c in df_full.columns
                                        else "case"
                                    )
                                    .strip()
                                    .replace(" ", "_")
                                )
                                pdf_filename = f"Appointment_{case_id_val}_{safe_uname}_{int(datetime.datetime.now().timestamp())}.pdf"
                                full_pdf_path = os.path.join(
                                    date_folder, pdf_filename
                                )
                                with open(full_pdf_path, "wb") as pf:
                                    pf.write(uploaded_pdf_file.getbuffer())
                                df_full.at[row_idx, "Appointment_PDF"] = (
                                    full_pdf_path
                                )

                        elif booking_status_choice == "Rescheduled":
                            df_full.at[row_idx, "Preferred_Location"] = (
                                popup_pref_loc.strip()
                            )
                            df_full.at[row_idx, "Preferred_Dates"] = (
                                popup_pref_dates.strip()
                            )
                            df_full.at[row_idx, "Prior_Notice"] = (
                                popup_prior_notice.strip()
                            )
                            changed = True

                        if changed:
                            df_full.at[row_idx, "Updated_By"] = current_user
                            df_full.at[row_idx, "Updated_At"] = current_time_str
                            df_full.at[row_idx, "Action_Type"] = (
                                "Rescheduled Case"
                                if booking_status_choice == "Rescheduled"
                                else "Updated Case"
                            )

                        if booking_status_choice == "Booked":
                            sheets_all = get_sheet_names()
                            booked_sheet_name = next(
                                (s for s in sheets_all if "booked" in s.lower()),
                                None,
                            )
                            if booked_sheet_name and booked_sheet_name != sheet_name:
                                df_target_booked = load_data(
                                    sheet_name=booked_sheet_name
                                )
                                row_to_transfer = df_full.loc[row_idx].copy()

                                for col in df_target_booked.columns:
                                    if col not in row_to_transfer:
                                        row_to_transfer[col] = ""

                                df_target_booked = pd.concat(
                                    [
                                        df_target_booked,
                                        pd.DataFrame([row_to_transfer]),
                                    ],
                                    ignore_index=True,
                                )
                                save_data(
                                    df_target_booked, sheet_name=booked_sheet_name
                                )
                                df_full = df_full.drop(row_idx)

                        save_data(df_full, sheet_name=sheet_name)

                        new_resched_id = None
                        target_act_sheet_p = None
                        if (
                            booking_status_choice == "Rescheduled"
                            and is_booked_sheet
                        ):
                            active_worker_pick = str(
                                df_full.at[row_idx, login_c]
                            ).strip()
                            df_full.at[row_idx, "Case_ID"] = case_id_val
                            df_full.at[row_idx, "Booking_Status"] = "Booked"
                            save_data(df_full, sheet_name=sheet_name)
                            new_resched_id, target_act_sheet_p = (
                                append_rescheduled_to_active_sheet(
                                    source_row=df_full.loc[row_idx],
                                    pref_loc=popup_pref_loc,
                                    pref_dates=popup_pref_dates,
                                    prior_notice=popup_prior_notice,
                                    updated_by=current_user,
                                    updated_at=current_time_str,
                                    assigned_worker=active_worker_pick,
                                )
                            )

                        st.session_state["is_editing_popup"] = False
                        st.session_state["is_reschedule_mode"] = False
                        st.session_state["is_delete_request_mode"] = False
                        st.session_state["selected_row_idx"] = None
                        st.session_state["open_dialog"] = None
                        if new_resched_id:
                            st.success(
                                f"✅ Kept in Booked sheet ({case_id_val}, Status: Booked) & added to {target_act_sheet_p} with new ID: {new_resched_id} (Worker: {active_worker_pick or 'Not Assigned'})"
                            )
                        else:
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


@st.dialog("📌 KPI Case Details", width="large")
def admin_kpi_cases_dialog(sheet_name, filter_mode, lead_filter="All"):
    today_str = get_ist_now().strftime("%Y-%m-%d")
    df_s = load_data(sheet_name=sheet_name)
    if df_s is None or df_s.empty:
        st.info("No cases found.")
        return

    df_s = reorder_dataframe_columns(df_s)
    login_c = find_login_user_column(df_s)
    display_c = find_display_username_column(df_s)
    p_col = find_password_column(df_s)
    lead_col = next(
        (c for c in df_s.columns if "lead" in c.strip().lower()), login_c
    )

    if lead_filter and lead_filter != "All" and lead_col in df_s.columns:
        df_s = df_s[
            df_s[lead_col]
            .astype(str)
            .str.strip()
            .str.lower()
            .str.contains(lead_filter.lower(), na=False)
        ]

    if filter_mode == "today_updated":
        st.info(
            f"🕒 Today Updated Cases ({today_str}) in **{sheet_name}** | Lead: **{lead_filter}**"
        )
        matched = df_s[
            df_s.apply(lambda r: is_today_updated_row(r, today_str), axis=1)
        ]
    elif filter_mode == "not_assigned":
        st.info(
            f"👤 Not Assigned Cases in **{sheet_name}** | Lead: **{lead_filter}**"
        )
        matched = df_s[df_s[login_c].apply(is_not_assigned_value)]
    else:
        matched = df_s

    if matched.empty:
        st.warning("No matching cases found for this KPI.")
    else:
        st.success(f"Total Matching Cases: {len(matched)}")
        render_clickable_table(
            matched,
            display_c,
            p_col,
            table_key=f"kpi_dlg_{filter_mode}_{sheet_name}",
            sheet_for_link=sheet_name,
            allow_edit=True,
        )


@st.dialog(
    "📥 Appointment PDFs Download Center (Date-Wise Folders & Search)",
    width="large",
)
def admin_pdf_download_dialog():
    st.info(
        "💡 Uploaded Appointment PDFs organized date-wise in respective folders with overall Case Username search:"
    )

    if not os.path.exists(PDF_BASE_DIR):
        os.makedirs(PDF_BASE_DIR)

    all_sheets = load_all_sheets_cached(_get_file_mtime())
    pdf_info_map = {}
    for s, df_s in all_sheets.items():
        if (
            df_s is not None
            and not df_s.empty
            and "Appointment_PDF" in df_s.columns
        ):
            disp_c = find_display_username_column(df_s)
            login_c = find_login_user_column(df_s)
            for _, r in df_s.iterrows():
                p_path = str(r.get("Appointment_PDF", "")).strip()
                if p_path:
                    abs_p = os.path.abspath(p_path)
                    pdf_info_map[abs_p] = {
                        "case_id": str(r.get("Case_ID", "N/A")),
                        "case_user": str(r.get(disp_c, "N/A")),
                        "worker": str(r.get(login_c, "N/A")),
                        "sheet": s,
                    }

    search_query = (
        st.text_input(
            "🔍 Search PDFs by Case ID, Case Username or Filename across all folders:",
            placeholder="Type Case ID or username here...",
            key="pdf_overall_search",
        )
        .strip()
        .lower()
    )

    date_folders = sorted(
        [
            d
            for d in os.listdir(PDF_BASE_DIR)
            if os.path.isdir(os.path.join(PDF_BASE_DIR, d))
        ],
        reverse=True,
    )
    root_files = [
        f
        for f in os.listdir(PDF_BASE_DIR)
        if f.lower().endswith(".pdf")
        and os.path.isfile(os.path.join(PDF_BASE_DIR, f))
    ]
    if root_files:
        date_folders.insert(0, "Unsorted / General")

    if not date_folders:
        st.warning("No date folders or Appointment PDFs found yet.")
    else:
        found_any = False
        for d_folder in date_folders:
            folder_path = (
                PDF_BASE_DIR
                if d_folder == "Unsorted / General"
                else os.path.join(PDF_BASE_DIR, d_folder)
            )
            files_in_folder = (
                [
                    f
                    for f in os.listdir(folder_path)
                    if f.lower().endswith(".pdf")
                    and os.path.isfile(os.path.join(folder_path, f))
                ]
                if os.path.isdir(folder_path)
                else []
            )

            if files_in_folder:
                filtered_files = []
                for f_name in files_in_folder:
                    f_full_path = os.path.join(folder_path, f_name)
                    abs_full_p = os.path.abspath(f_full_path)
                    meta = pdf_info_map.get(
                        abs_full_p,
                        {
                            "case_id": "N/A",
                            "case_user": "N/A",
                            "worker": "N/A",
                            "sheet": "N/A",
                        },
                    )

                    if (
                        search_query == ""
                        or search_query in meta["case_id"].lower()
                        or search_query in meta["case_user"].lower()
                        or search_query in f_name.lower()
                        or search_query in meta["worker"].lower()
                    ):
                        filtered_files.append((f_name, f_full_path, meta))

                if filtered_files:
                    found_any = True
                    expander_label = (
                        f"📁 Folder: {d_folder} ({len(filtered_files)} PDFs matching)"
                        if search_query != ""
                        else f"📁 Folder: {d_folder} ({len(filtered_files)} PDFs)"
                    )
                    with st.expander(
                        expander_label, expanded=(search_query != "")
                    ):
                        for f_name, f_full_path, meta in filtered_files:
                            with st.container(border=True):
                                c_info, c_dl, c_del = st.columns([2.6, 1.1, 1.1])
                                with c_info:
                                    st.write(f"📄 **File:** `{f_name}`")
                                    st.markdown(
                                        f"🆔 **Case ID:** `{meta['case_id']}` &nbsp;|&nbsp; 👤 **Case Username:** `{meta['case_user']}` &nbsp;|&nbsp; 👷 **Worker:** `{meta['worker']}`"
                                    )
                                    st.caption(
                                        f"Sheet: {meta['sheet']} | Path: {f_full_path}"
                                    )
                                with c_dl:
                                    with open(f_full_path, "rb") as f:
                                        b_data = f.read()
                                    st.download_button(
                                        label="📥 Download",
                                        data=b_data,
                                        file_name=f_name,
                                        mime="application/pdf",
                                        key=f"dl_pdf_{d_folder}_{f_name}_{os.path.getmtime(f_full_path)}",
                                        use_container_width=True,
                                    )
                                with c_del:
                                    if st.button(
                                        "🗑️ Delete",
                                        key=f"del_pdf_{d_folder}_{f_name}_{os.path.getmtime(f_full_path)}",
                                        use_container_width=True,
                                    ):
                                        try:
                                            if os.path.exists(f_full_path):
                                                os.remove(f_full_path)
                                            st.success(
                                                f"✅ Deleted '{f_name}' successfully!"
                                            )
                                            st.rerun()
                                        except Exception as e:
                                            st.error(
                                                f"❌ Error deleting file: {e}"
                                            )

        if search_query != "" and not found_any:
            st.warning(
                f"No PDFs found matching Case ID, username, or filename: '{search_query}'"
            )


@st.dialog("📢 Add / Update Notice (Scrolling Announcement)", width="medium")
def admin_notice_dialog():
    current_notice = get_admin_notice()
    st.info(
        "💡 Add or update the scrolling marquee notice for all case workers:"
    )

    with st.form("admin_notice_form"):
        notice_input = st.text_area(
            "Notice Text:",
            value=current_notice,
            placeholder="e.g., All case workers must complete pending cases by today...",
        )
        col_save_n, col_clear_n = st.columns(2)
        with col_save_n:
            submitted_notice = st.form_submit_button(
                "💾 Save Notice", type="primary", use_container_width=True
            )
        with col_clear_n:
            cleared_notice = st.form_submit_button(
                "🗑️ Clear Notice", use_container_width=True
            )

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

    tab_logged, tab_not_logged = st.tabs([
        f"🟢 Logged In ({len(logged_in_list)})",
        f"🔴 Not Logged In ({len(not_logged_in_list)})",
    ])

    with tab_logged:
        if not logged_in_list:
            st.warning("Currently no case workers are logged in.")
        else:
            for item in logged_in_list:
                c1, c2, c3 = st.columns([2, 2, 2])
                with c1:
                    st.write(f"👤 **{item['worker']}**")
                with c2:
                    st.success("🟢 Logged In")
                with c3:
                    st.text(item["time"])

    with tab_not_logged:
        if not not_logged_in_list:
            st.success("All case workers are currently logged in!")
        else:
            for item in not_logged_in_list:
                c1, c2, c3 = st.columns([2, 2, 2])
                with c1:
                    st.write(f"👤 **{item['worker']}**")
                with c2:
                    st.error("🔴 Not Logged In")
                with c3:
                    st.text(item["time"])


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
        with pd.ExcelWriter(att_buffer, engine="openpyxl") as writer:
            df_att.to_excel(writer, sheet_name="Attendance", index=False)

        st.download_button(
            label="📥 Download Attendance Report (.xlsx)",
            data=att_buffer.getvalue(),
            file_name="Case_Worker_Attendance.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="download_attendance_excel",
        )


@st.dialog(
    "📊 Case Worker & Visa Type Booked Cases Report", width="large"
)
def admin_worker_wise_report_dialog():
    st.info(
        "💡 Booked cases report broken down by Case Worker, Today's activity, Total cases, and Visa Types:"
    )
    sheets = get_sheet_names()
    target_sheet = next(
        (s for s in sheets if "booked" in s.lower()),
        sheets[0] if sheets else "Sheet1",
    )
    today_str = get_ist_now().strftime("%Y-%m-%d")

    today_worker_stats = {}
    total_worker_stats = {}
    visa_type_stats = {}

    df_s = load_data(sheet_name=target_sheet)
    if df_s is not None and not df_s.empty:
        login_c = find_login_user_column(df_s)
        visa_c = find_visa_type_column(df_s)
        date_cols = [
            c
            for c in df_s.columns
            if any(k in c.lower() for k in ["date", "time", "day", "created"])
        ]

        for _, r in df_s.iterrows():
            w_name = str(r.get(login_c, "Not Assigned")).strip()
            if w_name == "" or w_name.lower() == "nan":
                w_name = "Not Assigned"

            v_type = (
                str(r.get(visa_c, "Unspecified")).strip().upper()
                if visa_c
                else "Unspecified"
            )
            if v_type == "" or v_type == "NAN":
                v_type = "Unspecified"

            if w_name not in total_worker_stats:
                total_worker_stats[w_name] = {
                    "Total Cases": 0,
                    "Visa_Breakdown": {},
                    "Sheets": {target_sheet},
                }
            total_worker_stats[w_name]["Total Cases"] += 1
            total_worker_stats[w_name]["Visa_Breakdown"][v_type] = (
                total_worker_stats[w_name]["Visa_Breakdown"].get(v_type, 0) + 1
            )

            visa_type_stats[v_type] = visa_type_stats.get(v_type, 0) + 1

            is_today = False
            for dc in date_cols:
                if today_str in str(r.get(dc, "")).strip():
                    is_today = True
                    break

            if not is_today and today_str in str(r.get("Updated_At", "")):
                is_today = True

            if is_today:
                if w_name not in today_worker_stats:
                    today_worker_stats[w_name] = {
                        "Today Cases": 0,
                        "Visa_Breakdown": {},
                        "Sheets": {target_sheet},
                    }
                today_worker_stats[w_name]["Today Cases"] += 1
                today_worker_stats[w_name]["Visa_Breakdown"][v_type] = (
                    today_worker_stats[w_name]["Visa_Breakdown"].get(v_type, 0)
                    + 1
                )

    tab_today, tab_total, tab_visa = st.tabs([
        "📅 Today's Booked Cases",
        "📈 Total Booked Cases",
        "✈️ Visa Type-wise Summary",
    ])

    with tab_today:
        if not today_worker_stats:
            st.warning(
                f"No cases booked today ({today_str}) in '{target_sheet}'."
            )
        else:
            today_list = []
            for w, data in today_worker_stats.items():
                breakdown_str = ", ".join(
                    [f"{v}: {cnt}" for v, cnt in data["Visa_Breakdown"].items()]
                )
                today_list.append({
                    "Case Worker": w,
                    "Today Booked Cases": data["Today Cases"],
                    "Visa Type Breakdown": breakdown_str,
                    "Active Sheet": target_sheet,
                })
            df_today = pd.DataFrame(today_list)
            df_today.index = range(1, len(df_today) + 1)
            st.success(
                f"Total Cases Booked Today in '{target_sheet}': {df_today['Today Booked Cases'].sum()}"
            )
            st.dataframe(df_today, use_container_width=True)

    with tab_total:
        if not total_worker_stats:
            st.warning(f"No records found in '{target_sheet}'.")
        else:
            total_list = []
            for w, data in total_worker_stats.items():
                breakdown_str = ", ".join(
                    [f"{v}: {cnt}" for v, cnt in data["Visa_Breakdown"].items()]
                )
                total_list.append({
                    "Case Worker": w,
                    "Total Booked Cases": data["Total Cases"],
                    "Visa Type Breakdown": breakdown_str,
                    "Active Sheet": target_sheet,
                })
            df_total = pd.DataFrame(total_list)
            df_total.index = range(1, len(df_total) + 1)
            st.success(
                f"Total Booked Cases in '{target_sheet}': {df_total['Total Booked Cases'].sum()}"
            )
            st.dataframe(df_total, use_container_width=True)

    with tab_visa:
        if not visa_type_stats:
            st.warning(f"No visa type data found in '{target_sheet}'.")
        else:
            visa_list = [
                {"Visa Type / Category": v, "Total Booked Cases": count}
                for v, count in visa_type_stats.items()
            ]
            df_visa = pd.DataFrame(visa_list).sort_values(
                by="Total Booked Cases", ascending=False
            )
            df_visa.index = range(1, len(df_visa) + 1)
            st.dataframe(df_visa, use_container_width=True)


@st.dialog("📋 New / Updated Cases & Calendar", width="large")
def admin_updated_cases_dialog():
    st.info(
        "💡 View newly added cases, updated cases, and date-wise calendar records:"
    )
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    all_rows = []

    for s, df_s in all_sheets.items():
        if (
            df_s is not None
            and not df_s.empty
            and "Updated_By" in df_s.columns
        ):
            upd_rows = df_s[
                df_s["Updated_By"].astype(str).str.strip() != ""
            ]
            if not upd_rows.empty:
                upd_rows = upd_rows.copy()
                upd_rows.insert(0, "Sheet_Name", s)
                all_rows.append(upd_rows)

    tab_new, tab_upd, tab_cal = st.tabs([
        "➕ New Cases",
        "✏️ Updated Cases",
        "📅 Calendar View",
    ])

    with tab_new:
        if not all_rows:
            st.info("No new cases found.")
        else:
            combined_df = pd.concat(all_rows, ignore_index=True)
            p_col = find_password_column(combined_df)
            view_df = combined_df.drop(columns=[p_col], errors="ignore")
            new_cases_df = view_df[
                view_df["Action_Type"].astype(str).str.strip().str.lower()
                == "new case"
            ]
            if new_cases_df.empty:
                st.info("No new cases found.")
            else:
                st.success(f"Total New Cases: {len(new_cases_df)}")
                render_clickable_table(
                    new_cases_df,
                    find_display_username_column(new_cases_df),
                    p_col,
                    table_key="dlg_new_cases",
                )

    with tab_upd:
        if not all_rows:
            st.info("No updated cases found.")
        else:
            combined_df = pd.concat(all_rows, ignore_index=True)
            p_col = find_password_column(combined_df)
            view_df = combined_df.drop(columns=[p_col], errors="ignore")
            updated_cases_df = view_df[
                view_df["Action_Type"].astype(str).str.strip().str.lower()
                == "updated case"
            ]
            if updated_cases_df.empty:
                st.info("No updated cases found.")
            else:
                st.success(f"Total Updated Cases: {len(updated_cases_df)}")
                render_clickable_table(
                    updated_cases_df,
                    find_display_username_column(updated_cases_df),
                    p_col,
                    table_key="dlg_upd_cases",
                )

    with tab_cal:
        sheets = list(all_sheets.keys())
        selected_cal_sheet = st.selectbox(
            "Select Sheet for Calendar View",
            options=sheets,
            key="cal_sheet_select_tab",
        )
        df_cal = all_sheets.get(selected_cal_sheet)
        if df_cal is not None and not df_cal.empty:
            date_cols = [
                c
                for c in df_cal.columns
                if any(k in c.lower() for k in ["date", "time", "day"])
                and "prefer" not in c.lower()
            ]
            if date_cols:
                selected_date_col = st.selectbox(
                    "Select Date Column:",
                    options=date_cols,
                    key="cal_date_col_tab",
                )
                selected_date = st.date_input(
                    "Select Target Date:",
                    value=datetime.date.today(),
                    key="cal_target_date_tab",
                )
                date_str = selected_date.strftime("%Y-%m-%d")
                matched_df = df_cal[
                    df_cal[selected_date_col]
                    .astype(str)
                    .str.contains(date_str, na=False)
                ]
                if matched_df.empty:
                    st.info(f"No records found for date ({date_str}).")
                else:
                    pc = find_password_column(matched_df)
                    display_c = find_display_username_column(matched_df)
                    render_clickable_table(
                        matched_df,
                        display_c,
                        pc,
                        table_key="dlg_cal_cases",
                        sheet_for_link=selected_cal_sheet,
                    )


@st.dialog("🗑️ Pending Delete Requests", width="large")
def admin_delete_requests_dialog():
    st.info(
        "💡 Delete requests submitted by case workers along with their reasons:"
    )
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    del_reqs_list = []

    for s, df_s in all_sheets.items():
        if (
            df_s is not None
            and not df_s.empty
            and "Action_Type" in df_s.columns
        ):
            for idx, r in df_s.iterrows():
                if (
                    str(r.get("Action_Type", "")).strip().lower()
                    == "delete request"
                ):
                    r_dict = r.to_dict()
                    r_dict["Sheet_Name"] = s
                    r_dict["Original_Index"] = idx
                    del_reqs_list.append(r_dict)

    if not del_reqs_list:
        st.info("No pending delete requests.")
    else:
        for i, r_dict in enumerate(del_reqs_list):
            s_name = r_dict["Sheet_Name"]
            orig_idx = r_dict["Original_Index"]
            worker_name = r_dict.get("Updated_By", "Unknown")
            req_time = r_dict.get("Updated_At", "")
            case_id_del = r_dict.get("Case_ID", "N/A")
            df_temp = all_sheets.get(s_name)
            disp_col = (
                find_display_username_column(df_temp)
                if df_temp is not None
                else "USERNAME"
            )
            case_username = str(r_dict.get(disp_col, "N/A"))
            delete_reason = str(
                r_dict.get("Delete_Reason", "Reason not specified")
            )

            with st.expander(
                f"🆔 {case_id_del} | 📁 Sheet: {s_name} | Case User: {case_username} | Worker: {worker_name}"
            ):
                st.markdown(
                    f"🆔 **Case ID:** `{case_id_del}` &nbsp;|&nbsp; 👤 **Case Username:** `{case_username}`"
                )
                st.markdown(f"✍️ **Delete Reason:** `{delete_reason}`")
                st.caption(f"Requested Time: {req_time}")
                col_acc, col_rej = st.columns(2)
                with col_acc:
                    if st.button(
                        "✅ Accept Delete (Permanent Delete)",
                        key=f"accept_del_{s_name}_{orig_idx}_{i}",
                        type="primary",
                        use_container_width=True,
                    ):
                        df_target = load_data(sheet_name=s_name)
                        if df_target is not None and orig_idx in df_target.index:
                            df_target = df_target.drop(orig_idx)
                            save_data(df_target, sheet_name=s_name)
                            st.rerun()
                with col_rej:
                    if st.button(
                        "❌ Reject Delete (Restore Case)",
                        key=f"reject_del_{s_name}_{orig_idx}_{i}",
                        use_container_width=True,
                    ):
                        df_target = load_data(sheet_name=s_name)
                        if df_target is not None and orig_idx in df_target.index:
                            df_target.at[orig_idx, "Action_Type"] = (
                                "Delete Rejected"
                            )
                            save_data(df_target, sheet_name=s_name)
                            st.rerun()


@st.dialog("👥 Users, Passwords & Permissions (All Sheets)", width="large")
def permissions_dialog():
    st.info(
        "💡 Manage case worker status (Active / Blocked), passwords, and permissions."
    )
    all_blocked_now = are_all_workers_blocked()
    if all_blocked_now:
        if st.button(
            "🔓 Activate All Workers (Currently: All Blocked)",
            use_container_width=True,
            key="dlg_toggle_all_btn",
        ):
            toggle_all_workers_block(False)
            st.rerun()
    else:
        if st.button(
            "🔒 Block All Workers (Currently: Active)",
            use_container_width=True,
            key="dlg_toggle_all_btn",
        ):
            toggle_all_workers_block(True)
            st.rerun()

    st.divider()
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    all_users_dict = {}
    for _, df_s in all_sheets.items():
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
    for idx, (u_name, curr_pwd) in enumerate(all_users_dict.items()):
        matched_key = next(
            (k for k in current_perms if k.lower() == u_name.lower()), u_name
        )
        w_perm = current_perms.get(
            matched_key,
            {
                "can_update": True,
                "can_add": True,
                "can_search": True,
                "can_download": False,
                "can_sync": False,
                "block_login": False,
            },
        )
        is_blocked = w_perm.get("block_login", False)

        with st.container(border=True):
            hdr_c1, hdr_c2 = st.columns([2.5, 1.5], gap="small")
            with hdr_c1:
                status_badge = "🔴 BLOCKED" if is_blocked else "🟢 ACTIVE"
                st.markdown(f"**👤 {u_name} | Status: {status_badge}**")
            with hdr_c2:
                if is_blocked:
                    if st.button(
                        "🔓 Activate Worker",
                        key=f"dlg_toggle_status_{idx}",
                        use_container_width=True,
                    ):
                        w_perm["block_login"] = False
                        current_perms[matched_key] = w_perm
                        save_permissions(current_perms)
                        st.rerun()
                else:
                    if st.button(
                        "🔒 Block Worker",
                        key=f"dlg_toggle_status_{idx}",
                        use_container_width=True,
                    ):
                        w_perm["block_login"] = True
                        current_perms[matched_key] = w_perm
                        save_permissions(current_perms)
                        record_worker_logout(u_name)
                        st.rerun()

            c1, c2, c3, c4, c5, c6, c_btn = st.columns(
                [2.1, 0.9, 0.9, 0.9, 0.9, 0.9, 1.4], gap="small"
            )
            with c1:
                new_pwd_val = st.text_input(
                    "Password", value=curr_pwd, key=f"dlg_p_{idx}"
                )
            with c2:
                new_upd_val = st.checkbox(
                    "Update",
                    value=w_perm.get("can_update", True),
                    key=f"dlg_upd_{idx}",
                )
            with c3:
                new_add_val = st.checkbox(
                    "Add",
                    value=w_perm.get("can_add", True),
                    key=f"dlg_add_{idx}",
                )
            with c4:
                new_sch_val = st.checkbox(
                    "Search",
                    value=w_perm.get("can_search", True),
                    key=f"dlg_sch_{idx}",
                )
            with c5:
                new_dl_val = st.checkbox(
                    "Download",
                    value=w_perm.get("can_download", False),
                    key=f"dlg_dl_{idx}",
                )
            with c6:
                new_sync_val = st.checkbox(
                    "Sync",
                    value=w_perm.get("can_sync", False),
                    key=f"dlg_sync_{idx}",
                )
            with c_btn:
                if st.button(
                    "💾 Save",
                    key=f"dlg_btn_upd_{idx}",
                    type="primary",
                    use_container_width=True,
                ):
                    try:
                        all_sheets_data = pd.read_excel(
                            EXCEL_FILE, sheet_name=None, dtype=str
                        )
                        for s, df_s in all_sheets_data.items():
                            lc = find_login_user_column(df_s)
                            pc = find_password_column(df_s)
                            if lc in df_s.columns and pc in df_s.columns:
                                mask = (
                                    df_s[lc]
                                    .astype(str)
                                    .str.strip()
                                    .str.lower()
                                    == u_name.lower()
                                )
                                if mask.any():
                                    df_s.loc[mask, pc] = new_pwd_val.strip()
                            all_sheets_data[s] = reorder_dataframe_columns(
                                df_s
                            )

                        with pd.ExcelWriter(
                            EXCEL_FILE, engine="openpyxl"
                        ) as writer:
                            for s, df_val in all_sheets_data.items():
                                df_val.to_excel(
                                    writer, sheet_name=str(s), index=False
                                )

                        current_perms[matched_key] = {
                            "can_update": new_upd_val,
                            "can_add": new_add_val,
                            "can_search": new_sch_val,
                            "can_download": new_dl_val,
                            "can_sync": new_sync_val,
                            "block_login": is_blocked,
                        }
                        save_permissions(current_perms)
                        st.cache_data.clear()
                        st.success(f"✅ Saved changes for '{u_name}'!")
                    except Exception as e:
                        st.error(f"❌ Error updating '{u_name}': {e}")


@st.dialog("🔑 Change Admin Password", width="medium")
def admin_password_dialog():
    with st.form("dlg_admin_pass_form"):
        curr_p = st.text_input("Current Password", type="password")
        new_p = st.text_input("New Password", type="password")
        conf_p = st.text_input("Confirm Password", type="password")
        if st.form_submit_button(
            "Update Admin Password", type="primary", use_container_width=True
        ):
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
    st.info(
        "💡 Upload your updated local Excel file here to instantly synchronize it with the cloud."
    )
    uploaded_excel = st.file_uploader(
        "Upload Updated Excel File (.xlsx)", type=["xlsx"]
    )
    if uploaded_excel is not None:
        if st.button(
            "🚀 Confirm & Sync to Cloud",
            type="primary",
            use_container_width=True,
        ):
            try:
                with open(EXCEL_FILE, "wb") as f:
                    f.write(uploaded_excel.getbuffer())
                st.cache_data.clear()
                ensure_all_cases_have_unique_ids(force=True)
                st.success("✅ Cloud data updated successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {e}")


@st.dialog("🔑 Change Password", width="medium")
def worker_password_dialog(worker_name):
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    worker_current_pass = ""
    for _, df_chk in all_sheets.items():
        if df_chk is not None and not df_chk.empty:
            lc = find_login_user_column(df_chk)
            pc = find_password_column(df_chk)
            matched_r = df_chk[
                df_chk[lc].astype(str).str.strip().str.lower()
                == worker_name.lower()
            ]
            if not matched_r.empty:
                worker_current_pass = str(matched_r.iloc[0][pc]).strip()
                break

    with st.form("dlg_worker_pass_form"):
        curr_wp = st.text_input("Current Password", type="password")
        new_wp = st.text_input("New Password", type="password")
        conf_wp = st.text_input("Confirm Password", type="password")
        if st.form_submit_button(
            "Update Password", type="primary", use_container_width=True
        ):
            if curr_wp.strip() != worker_current_pass:
                st.error("❌ Incorrect current password!")
            elif new_wp.strip() == "":
                st.warning("⚠️ New password cannot be empty.")
            elif new_wp != conf_wp:
                st.error("❌ New passwords do not match!")
            else:
                try:
                    all_sheets_data = pd.read_excel(
                        EXCEL_FILE, sheet_name=None, dtype=str
                    )
                    for s, df_s in all_sheets_data.items():
                        lc = find_login_user_column(df_s)
                        pc = find_password_column(df_s)
                        if lc in df_s.columns and pc in df_s.columns:
                            mask = (
                                df_s[lc].astype(str).str.strip().str.lower()
                                == worker_name.lower()
                            )
                            if mask.any():
                                df_s.loc[mask, pc] = new_wp.strip()
                        all_sheets_data[s] = reorder_dataframe_columns(df_s)
                    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                        for s, df_val in all_sheets_data.items():
                            df_val.to_excel(
                                writer, sheet_name=str(s), index=False
                            )
                    st.cache_data.clear()
                    st.success("✅ Password updated successfully!")
                except Exception as e:
                    st.error(f"❌ Error: {e}")


@st.dialog("➕ Add New Case", width="large")
def add_case_dialog(worker_name):
    sheets = get_sheet_names()
    default_idx = 0
    for idx_s, s_nm in enumerate(sheets):
        if "booked" not in s_nm.lower():
            default_idx = idx_s
            break

    selected_target_sheet = st.selectbox(
        "Select Target Sheet for New Case",
        options=sheets,
        index=default_idx,
        key="dlg_target_sheet",
    )
    target_df = load_data(sheet_name=selected_target_sheet)
    t_login_col = find_login_user_column(target_df)
    t_pass_col = find_password_column(target_df)
    preview_next_id = generate_next_case_id()

    available_usernames = sorted(
        list(set(target_df[t_login_col].astype(str).str.strip().tolist()))
    )
    available_usernames = [u for u in available_usernames if u != ""]
    available_usernames.insert(0, "Not Assigned")
    if worker_name not in available_usernames and worker_name != "":
        available_usernames.insert(1, worker_name)

    with st.form("dlg_new_case_form", clear_on_submit=True):
        st.text_input(
            "🆔 Auto-Generated Case ID (Locked)",
            value=preview_next_id,
            disabled=True,
        )
        assigned_username = st.selectbox(
            "Assign Username",
            options=available_usernames,
            index=(
                available_usernames.index(worker_name)
                if worker_name in available_usernames
                else 0
            ),
        )
        form_data = {}
        cols = [
            c
            for c in target_df.columns
            if c
            not in [
                "Case_ID",
                t_login_col,
                t_pass_col,
                "Updated_By",
                "Updated_At",
                "Action_Type",
                "Delete_Reason",
                "Booking_Status",
                "Appointment_Entry",
                "Appointment_PDF",
                "Preferred_Location",
                "Preferred_Dates",
                "Prior_Notice",
            ]
        ]
        for col_name in cols:
            c_lower = col_name.lower()
            if "date" in c_lower and "prefer" not in c_lower:
                sel_date = st.date_input(
                    f"{col_name}",
                    value=datetime.date.today(),
                    key=f"add_date_{col_name}",
                )
                form_data[col_name] = sel_date.strftime("%Y-%m-%d")
            else:
                form_data[col_name] = st.text_input(
                    f"{col_name}", key=f"add_txt_{col_name}"
                )

        if st.form_submit_button(
            "🚀 Submit New Case", type="primary", use_container_width=True
        ):
            try:
                final_new_case_id = generate_next_case_id()
                new_row = {col: "" for col in target_df.columns}
                new_row["Case_ID"] = final_new_case_id
                final_assigned_user = (
                    ""
                    if assigned_username == "Not Assigned"
                    else assigned_username.strip()
                )
                new_row[t_login_col] = final_assigned_user

                current_pass = "admin123"
                matched_pass_rows = target_df[
                    target_df[t_login_col].astype(str).str.strip().str.lower()
                    == final_assigned_user.lower()
                ]
                if not matched_pass_rows.empty:
                    current_pass = str(
                        matched_pass_rows.iloc[0][t_pass_col]
                    ).strip()

                new_row[t_pass_col] = (
                    current_pass if final_assigned_user != "" else "admin123"
                )
                new_row["Updated_By"] = worker_name
                new_row["Updated_At"] = get_ist_now().strftime(
                    "%Y-%m-%d %I:%M:%S %p"
                )
                new_row["Action_Type"] = "New Case"
                new_row["Booking_Status"] = "Active"

                for k, v in form_data.items():
                    new_row[k] = v

                updated_target_df = pd.concat(
                    [target_df, pd.DataFrame([new_row])], ignore_index=True
                )
                save_data(updated_target_df, sheet_name=selected_target_sheet)
                st.session_state["open_dialog"] = None
                st.success(
                    f"✅ New case added successfully with Case ID: {final_new_case_id}"
                )
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {e}")


@st.dialog("🔍 Search Records (Read-Only)", width="large")
def search_data_dialog():
    sheets = get_sheet_names()
    selected_search_sheet = st.selectbox(
        "Select Sheet to Search:", options=sheets, key="dlg_search_sheet"
    )
    df_search = load_data(sheet_name=selected_search_sheet)
    if df_search is not None and not df_search.empty:
        p_col = find_password_column(df_search)
        search_query = (
            st.text_input(
                "Search by Case ID, Case Username, DS160, Case Password, or Case Worker:",
                placeholder="Type Case ID, Username, DS160, Password, or Case Worker...",
                key="dlg_search_input",
            )
            .strip()
        )
        if search_query != "":
            found_data = df_search[
                search_username_ds160_password_worker(df_search, search_query)
            ]
            if found_data.empty:
                st.warning(f"No matching records found for '{search_query}'.")
            else:
                view_only_df = found_data.drop(
                    columns=[
                        p_col,
                        "Updated_By",
                        "Updated_At",
                        "Action_Type",
                        "Delete_Reason",
                        "Appointment_PDF",
                    ],
                    errors="ignore",
                )
                view_only_df = reorder_dataframe_columns(view_only_df)
                view_only_df.index = range(1, len(view_only_df) + 1)
                st.dataframe(view_only_df, use_container_width=True)


dialog_map = {
    "permissions": permissions_dialog,
    "admin_password": admin_password_dialog,
    "sync_excel": sync_excel_dialog,
    "admin_updated_cases": admin_updated_cases_dialog,
    "admin_delete_requests": admin_delete_requests_dialog,
    "admin_notice": admin_notice_dialog,
    "admin_login_status": admin_login_status_dialog,
    "admin_attendance": admin_attendance_dialog,
    "admin_worker_report": admin_worker_wise_report_dialog,
    "admin_pdf_downloads": admin_pdf_download_dialog,
    "worker_password": lambda: worker_password_dialog(
        st.session_state.get("username", "")
    ),
    "add_case": lambda: add_case_dialog(st.session_state.get("username", "")),
    "search_data": search_data_dialog,
}

if st.session_state["open_dialog"] in dialog_map:
    dialog_map[st.session_state["open_dialog"]]()
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "admin_kpi_popup":
    kpi_meta = st.session_state.get("admin_kpi_popup_meta", {})
    admin_kpi_cases_dialog(
        sheet_name=kpi_meta.get("sheet", default_s_init),
        filter_mode=kpi_meta.get("mode", "today_updated"),
        lead_filter=kpi_meta.get("lead", "All"),
    )
    st.session_state["open_dialog"] = None
elif st.session_state["open_dialog"] == "edit_case_popup":
    r_idx = st.session_state.get("selected_row_idx")
    if r_idx is not None:
        curr_role = st.session_state.get("role", "worker")
        curr_u = st.session_state.get("username", "")
        perms = load_permissions()
        c_upd = (
            True
            if curr_role == "admin"
            else perms.get(curr_u, {}).get("can_update", True)
        )
        view_and_edit_case_dialog(
            st.session_state["active_sheet"], r_idx, curr_role, curr_u, c_upd
        )

if not st.session_state["logged_in"]:
    default_nav_idx = 5 if st.session_state.get("force_logged_out_msg") else 0
    selected_page = st.radio(
        "Navigation",
        [
            "🏠 Home",
            "ℹ️ About Us",
            "✈️ Services",
            "⭐ Success Stories",
            "📞 Contact Us",
            "🔐 Portal Login",
        ],
        index=default_nav_idx,
        horizontal=True,
        label_visibility="collapsed",
    )

    if selected_page == "🏠 Home":
        st.markdown("""
            <div class="hero-section">
                <div class="hero-title">🌍 Welcome to HemanVisa Global Experts</div>
                <div class="hero-subtitle">Your trusted partner for international study, work visas, permanent residency, and seamless global mobility.</div>
            </div>
        """, unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(
                '<div class="corporate-card"><h3>🎓 Study Abroad</h3><p>Secure admissions in top global universities with complete visa documentation and interview support.</p></div>',
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown(
                '<div class="corporate-card"><h3>💼 Work & Skilled Visas</h3><p>Expand your career internationally with employer-sponsored work visas and skilled migration pathways.</p></div>',
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                '<div class="corporate-card"><h3>✈️ Permanent Residency</h3><p>Settle abroad permanently with hassle-free PR processing and legal consulting.</p></div>',
                unsafe_allow_html=True,
            )

    elif selected_page == "ℹ️ About Us":
        st.markdown(
            "<div class='section-banner'>ℹ️ About HemanVisa & Immigration Services</div>",
            unsafe_allow_html=True,
        )
        st.write(
            "HemanVisa is a global immigration and visa consulting firm dedicated to turning international aspirations into reality."
        )

    elif selected_page == "✈️ Services":
        st.markdown(
            "<div class='section-banner'>✈️ Our Premium Global Services</div>",
            unsafe_allow_html=True,
        )
        s1, s2 = st.columns(2)
        with s1:
            st.markdown(
                '<div class="corporate-card"><h3>📋 DS-160 & Visa Filing</h3><p>Professional assistance with visa application forms, document verification, and appointment booking.</p></div>',
                unsafe_allow_html=True,
            )
        with s2:
            st.markdown(
                '<div class="corporate-card"><h3>📚 Student Career Counseling</h3><p>Course selection, university applications, scholarship guidance, and financial support.</p></div>',
                unsafe_allow_html=True,
            )

    elif selected_page == "⭐ Success Stories":
        st.markdown(
            "<div class='section-banner'>⭐ Client Success Stories</div>",
            unsafe_allow_html=True,
        )
        st.info("Trusted by hundreds of applicants across USA, UK, Canada, and Europe.")

    elif selected_page == "📞 Contact Us":
        st.markdown(
            "<div class='section-banner'>📞 Get in Touch with Our Experts</div>",
            unsafe_allow_html=True,
        )
        with st.form("contact_form"):
            st.text_input("Your Name")
            st.text_input("Email / Phone")
            st.text_area("Your Inquiry / Visa Requirement")
            if st.form_submit_button(
                "Submit Inquiry", type="primary", use_container_width=True
            ):
                st.success("✅ Thank you! Our expert will contact you shortly.")

    elif selected_page == "🔐 Portal Login":
        col_l1, col_l2, col_l3 = st.columns([1, 1.4, 1])
        with col_l2:
            if st.session_state.get("force_logged_out_msg"):
                st.error(
                    "⚠️ Your session was terminated because Admin blocked your account."
                )
                st.session_state["force_logged_out_msg"] = False
            with st.container(border=True):
                st.markdown(
                    "<h3 style='text-align:center; color:#1e293b;'>🔐 Portal Login</h3>",
                    unsafe_allow_html=True,
                )
                sheets = get_sheet_names()
                with st.form("login_form"):
                    username_input = st.text_input(
                        "👤 Username", placeholder="Type your username"
                    )
                    password_input = st.text_input(
                        "🔑 Password",
                        type="password",
                        placeholder="Type your password",
                    )
                    submit_btn = st.form_submit_button(
                        "🚀 Login to Portal",
                        type="primary",
                        use_container_width=True,
                    )
                    if submit_btn:
                        u_in = username_input.strip()
                        p_in = password_input.strip()
                        admin_u, admin_p = get_admin_creds()
                        default_act_sheet = get_default_active_sheet(sheets)

                        if u_in.lower() == admin_u.lower() and p_in == admin_p:
                            st.session_state["logged_in"] = True
                            st.session_state["username"] = admin_u
                            st.session_state["role"] = "admin"
                            st.session_state["active_sheet"] = default_act_sheet
                            st.session_state["dash_sheet"] = default_act_sheet
                            st.session_state["records_sheet"] = default_act_sheet
                            st.session_state["admin_active_view"] = "dashboard"
                            st.query_params.clear()
                            st.query_params["logged_in"] = "true"
                            st.query_params["username"] = admin_u
                            st.query_params["role"] = "admin"
                            st.query_params["active_sheet"] = default_act_sheet
                            st.rerun()
                        elif is_worker_blocked(u_in):
                            st.error(
                                "❌ Your login access has been disabled by Admin."
                            )
                        else:
                            logged_in_success = False
                            actual_matched_username = ""
                            all_sheets_chk = load_all_sheets_cached(
                                _get_file_mtime()
                            )
                            for _, df_check in all_sheets_chk.items():
                                if df_check is not None and not df_check.empty:
                                    login_c = find_login_user_column(df_check)
                                    p_col = find_password_column(df_check)
                                    for _, r in df_check.iterrows():
                                        db_user = str(r[login_c]).strip()
                                        db_pass = str(r[p_col]).strip()
                                        if (
                                            db_user.lower() == u_in.lower()
                                            and db_pass == p_in
                                        ):
                                            logged_in_success = True
                                            actual_matched_username = db_user
                                            break
                                if logged_in_success:
                                    break

                            if logged_in_success:
                                record_worker_login(actual_matched_username)
                                st.session_state["logged_in"] = True
                                st.session_state["username"] = (
                                    actual_matched_username
                                )
                                st.session_state["role"] = "worker"
                                st.session_state["active_sheet"] = (
                                    default_act_sheet
                                )
                                st.session_state["dash_sheet"] = (
                                    default_act_sheet
                                )
                                st.session_state["records_sheet"] = (
                                    default_act_sheet
                                )
                                st.session_state["attendance_recorded"] = True
                                st.query_params.clear()
                                st.query_params["logged_in"] = "true"
                                st.query_params["username"] = (
                                    actual_matched_username
                                )
                                st.query_params["role"] = "worker"
                                st.query_params["active_sheet"] = (
                                    default_act_sheet
                                )
                                st.rerun()
                            else:
                                st.error("❌ Incorrect Username or Password!")

elif st.session_state["role"] == "admin":
    all_sheets_map = load_all_sheets_cached(_get_file_mtime())
    sheets = list(all_sheets_map.keys()) if all_sheets_map else ["Sheet1"]
    default_act_s = get_default_active_sheet(sheets)
    current_sheet = st.session_state.get("active_sheet", default_act_s)
    if current_sheet not in sheets:
        current_sheet = default_act_s
        st.session_state["active_sheet"] = current_sheet

    active_admin_view = st.session_state.get("admin_active_view", "dashboard")
    if active_admin_view not in ["dashboard", "records"]:
        active_admin_view = "dashboard"
        st.session_state["admin_active_view"] = "dashboard"

    header_sheet_label = (
        st.session_state.get("dash_sheet", current_sheet)
        if active_admin_view == "dashboard"
        else st.session_state.get("records_sheet", current_sheet)
    )

    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    top_col1, top_col2, top_col3 = st.columns([6.6, 1.2, 1.2], gap="small")
    with top_col1:
        st.markdown(
            f"<div class='workspace-header-bar'><span>👑 Admin Control Panel</span><span class='workspace-sheet-pill'>📂 {header_sheet_label}</span></div>",
            unsafe_allow_html=True,
        )
    with top_col2:
        if st.button("🔄 Refresh", key="admin_refresh"):
            st.cache_data.clear()
            st.session_state.pop("_last_id_sync_mtime", None)
            act_s = get_default_active_sheet(get_sheet_names())
            st.session_state["active_sheet"] = act_s
            st.session_state["dash_sheet"] = act_s
            st.session_state["records_sheet"] = act_s
            st.session_state["selected_visa_filter"] = "All"
            st.session_state["records_selected_lead"] = "All"
            st.session_state["open_dialog"] = None
            st.session_state["selected_row_idx"] = None
            st.query_params["active_sheet"] = act_s
            st.rerun()
    with top_col3:
        if st.button("🚪 Logout", key="admin_logout"):
            st.query_params.clear()
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            st.session_state["admin_active_view"] = "dashboard"
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    t_c1, t_c2 = st.columns([1.4, 1.4], gap="small")
    with t_c1:
        is_d_act = active_admin_view == "dashboard"
        st.markdown(
            f'<span class="{"admin-btn-0" if is_d_act else "admin-btn-6"}" style="display:none;"></span>',
            unsafe_allow_html=True,
        )
        if st.button(
            "📊 Dashboard & Quick Actions",
            key="admin_tab_dash_btn",
            use_container_width=True,
        ):
            st.session_state["admin_active_view"] = "dashboard"
            st.rerun()
    with t_c2:
        is_r_act = active_admin_view == "records"
        st.markdown(
            f'<span class="{"admin-btn-0" if is_r_act else "admin-btn-6"}" style="display:none;"></span>',
            unsafe_allow_html=True,
        )
        if st.button(
            "📂 Case Records & Search",
            key="admin_tab_rec_btn",
            use_container_width=True,
        ):
            st.session_state["admin_active_view"] = "records"
            st.rerun()

    if "admin_selected_lead" not in st.session_state:
        st.session_state["admin_selected_lead"] = "All"

    if active_admin_view == "dashboard":
        dash_top_l, dash_top_r = st.columns([6.2, 1.8], gap="small")
        with dash_top_l:
            st.markdown(
                "<div class='section-banner'>📊 Admin Dashboard & Quick Actions</div>",
                unsafe_allow_html=True,
            )
        with dash_top_r:
            if os.path.exists(EXCEL_FILE):
                with open(EXCEL_FILE, "rb") as f_dl_corner:
                    st.download_button(
                        label="📥 Download Excel",
                        data=f_dl_corner.read(),
                        file_name=EXCEL_FILE,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="admin_dash_corner_dl_btn",
                    )

        dash_s_name = st.session_state.get("dash_sheet", default_act_s)
        if dash_s_name not in sheets:
            dash_s_name = default_act_s
        df_d_base = all_sheets_map.get(dash_s_name, pd.DataFrame()).copy()
        df_d_base = reorder_dataframe_columns(df_d_base)
        login_c_d = find_login_user_column(df_d_base)
        lead_col_d = next(
            (c for c in df_d_base.columns if "lead" in c.strip().lower()),
            login_c_d,
        )

        dash_search_query = st.text_input(
            "🔍 Search All Sheets (Case ID, Username, DS160, Password & Case Worker):",
            placeholder="Type Case ID, Username, DS160, Case Password, or Case Worker...",
            key="dash_overall_search_input",
        ).strip()

        st.markdown(
            "<div class='action-group-title'>👥 FILTER BY LEAD NAME:</div>",
            unsafe_allow_html=True,
        )
        lead_options = ["All", "Kiran", "Shanthi"]
        l_btn_cols = st.columns([1.2, 1.2, 1.2, 4.4], gap="small")
        for l_idx, l_name in enumerate(lead_options):
            with l_btn_cols[l_idx]:
                is_l_act = st.session_state["admin_selected_lead"] == l_name
                l_cls = "admin-btn-0" if is_l_act else "admin-btn-6"
                l_lbl = f"✅ {l_name}" if is_l_act else l_name
                st.markdown(
                    f'<span class="{l_cls}" style="display:none;"></span>',
                    unsafe_allow_html=True,
                )
                if st.button(
                    l_lbl,
                    key=f"dash_lead_btn_{l_name}_{l_idx}",
                    use_container_width=True,
                ):
                    st.session_state["admin_selected_lead"] = l_name
                    st.rerun()

        if sheets:
            st.markdown(
                "<div class='action-group-title'>📑 SWITCH SHEET (DASHBOARD FILTER):</div>",
                unsafe_allow_html=True,
            )
            sw_cols = st.columns(len(sheets), gap="small")
            for idx, s_name in enumerate(sheets):
                with sw_cols[idx]:
                    is_active = s_name == dash_s_name
                    color_cls = "admin-btn-1" if is_active else "admin-btn-6"
                    btn_txt = f"✅ {s_name}" if is_active else f"📄 {s_name}"
                    st.markdown(
                        f'<span class="{color_cls}" style="display:none;"></span>',
                        unsafe_allow_html=True,
                    )
                    if st.button(
                        btn_txt,
                        key=f"dash_only_switch_sheet_{s_name}_{idx}",
                        use_container_width=True,
                    ):
                        st.session_state["dash_sheet"] = s_name
                        st.rerun()

        if (
            st.session_state["admin_selected_lead"] != "All"
            and not df_d_base.empty
            and lead_col_d in df_d_base.columns
        ):
            sel_l = st.session_state["admin_selected_lead"].lower()
            df_dash = df_d_base[
                df_d_base[lead_col_d]
                .astype(str)
                .str.strip()
                .str.lower()
                .str.contains(sel_l, na=False)
            ]
        else:
            df_dash = df_d_base

        if dash_search_query != "":
            all_matched = search_across_all_sheets_by_fields(
                dash_search_query,
                lead_filter=st.session_state["admin_selected_lead"],
            )
            if all_matched.empty:
                st.warning(
                    f"No matching records found across all sheets for '{dash_search_query}'."
                )
            else:
                p_col_m = find_password_column(all_matched)
                disp_m = find_display_username_column(all_matched)
                render_clickable_table(
                    all_matched,
                    disp_m,
                    p_col_m,
                    table_key="dash_global_search",
                    allow_edit=False,
                )

        total_cases_d = len(df_dash) if df_dash is not None else 0
        today_str_d = get_ist_now().strftime("%Y-%m-%d")
        today_count_d = 0
        today_updated_count_d = 0
        not_assigned_count_d = 0

        if df_dash is not None and not df_dash.empty:
            date_cols_d = [
                c
                for c in df_dash.columns
                if any(k in c.lower() for k in ["date", "time", "day", "created"])
            ]
            for _, r in df_dash.iterrows():
                if any(
                    today_str_d in str(r.get(dc, "")).strip()
                    for dc in date_cols_d
                ) or today_str_d in str(r.get("Updated_At", "")):
                    today_count_d += 1
                if is_today_updated_row(r, today_str_d):
                    today_updated_count_d += 1
                if is_not_assigned_value(r.get(login_c_d, "")):
                    not_assigned_count_d += 1

        del_reqs_count_d = 0
        for _, df_s in all_sheets_map.items():
            if (
                df_s is not None
                and not df_s.empty
                and "Action_Type" in df_s.columns
            ):
                del_reqs_count_d += int(
                    (
                        df_s["Action_Type"].astype(str).str.strip().str.lower()
                        == "delete request"
                    ).sum()
                )

        all_workers_d = get_all_workers_list()
        login_recs_d = load_login_tracking()
        norm_logins_d = {k.lower(): v for k, v in login_recs_d.items()}
        logged_in_cnt_d = sum(
            1
            for w in all_workers_d
            if norm_logins_d.get(w.lower(), {}).get("status") == "Logged In"
        )

        st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
        k_cols = st.columns(6, gap="small")
        with k_cols[0]:
            st.markdown(
                f"<div class='kpi-card'><div class='kpi-title'>Total ({dash_s_name})</div><div class='kpi-value'>{total_cases_d}</div></div>",
                unsafe_allow_html=True,
            )
        with k_cols[1]:
            st.markdown(
                f"<div class='kpi-card green'><div class='kpi-title'>Today's Cases</div><div class='kpi-value'>{today_count_d}</div></div>",
                unsafe_allow_html=True,
            )
        with k_cols[2]:
            st.markdown(
                f"<div class='kpi-card teal'><div class='kpi-title'>Today Updated</div><div class='kpi-value'>{today_updated_count_d}</div></div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                '<span class="admin-btn-5" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                f"🕒 View ({today_updated_count_d})",
                key="kpi_btn_today_updated",
                use_container_width=True,
            ):
                st.session_state["admin_kpi_popup_meta"] = {
                    "sheet": dash_s_name,
                    "mode": "today_updated",
                    "lead": st.session_state["admin_selected_lead"],
                }
                st.session_state["open_dialog"] = "admin_kpi_popup"
                st.rerun()
        with k_cols[3]:
            st.markdown(
                f"<div class='kpi-card rose'><div class='kpi-title'>Not Assigned</div><div class='kpi-value'>{not_assigned_count_d}</div></div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                '<span class="admin-btn-7" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                f"👤 View ({not_assigned_count_d})",
                key="kpi_btn_not_assigned",
                use_container_width=True,
            ):
                st.session_state["admin_kpi_popup_meta"] = {
                    "sheet": dash_s_name,
                    "mode": "not_assigned",
                    "lead": st.session_state["admin_selected_lead"],
                }
                st.session_state["open_dialog"] = "admin_kpi_popup"
                st.rerun()
        with k_cols[4]:
            st.markdown(
                f"<div class='kpi-card purple'><div class='kpi-title'>Workers Online</div><div class='kpi-value'>{logged_in_cnt_d} / {len(all_workers_d)}</div></div>",
                unsafe_allow_html=True,
            )
        with k_cols[5]:
            st.markdown(
                f"<div class='kpi-card gold'><div class='kpi-title'>Delete Reqs</div><div class='kpi-value'>{del_reqs_count_d}</div></div>",
                unsafe_allow_html=True,
            )

        st.markdown(
            "<div class='action-group-title'>⚡ QUICK ACTIONS & REPORTS:</div>",
            unsafe_allow_html=True,
        )
        q_r1_1, q_r1_2, q_r1_3, q_r1_4 = st.columns(4, gap="small")
        with q_r1_1:
            st.markdown(
                '<span class="admin-btn-0" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "👥 Users & Permissions",
                key="q_btn_perms",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "permissions"
                st.rerun()
        with q_r1_2:
            st.markdown(
                '<span class="admin-btn-1" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "📊 Worker Booked Report",
                key="q_btn_wrep",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "admin_worker_report"
                st.rerun()
        with q_r1_3:
            st.markdown(
                '<span class="admin-btn-2" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "📋 New / Updated & Cal",
                key="q_btn_upd",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "admin_updated_cases"
                st.rerun()
        with q_r1_4:
            st.markdown(
                '<span class="admin-btn-3" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                f"🗑️ Delete Reqs ({del_reqs_count_d})",
                key="q_btn_delreqs",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "admin_delete_requests"
                st.rerun()

        q_r2_1, q_r2_2, q_r2_3, q_r2_4 = st.columns(4, gap="small")
        with q_r2_1:
            st.markdown(
                '<span class="admin-btn-4" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "📥 Appointment PDFs",
                key="q_btn_pdfs",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "admin_pdf_downloads"
                st.rerun()
        with q_r2_2:
            st.markdown(
                '<span class="admin-btn-5" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "🟢 Login Status",
                key="q_btn_lstatus",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "admin_login_status"
                st.rerun()
        with q_r2_3:
            st.markdown(
                '<span class="admin-btn-6" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "🕒 Attendance Log",
                key="q_btn_att",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "admin_attendance"
                st.rerun()
        with q_r2_4:
            st.markdown(
                '<span class="admin-btn-7" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "📢 Notice Board",
                key="q_btn_notice",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "admin_notice"
                st.rerun()

        q_r3_1, q_r3_2, q_r3_3 = st.columns([1.3, 1.3, 1.4], gap="small")
        with q_r3_1:
            st.markdown(
                '<span class="admin-btn-1" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "📤 Sync Excel to Cloud",
                key="q_btn_sync",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "sync_excel"
                st.rerun()
        with q_r3_2:
            st.markdown(
                '<span class="admin-btn-6" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "🔑 Change Admin Pass",
                key="q_btn_apass",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "admin_password"
                st.rerun()
        with q_r3_3:
            all_blk_d = are_all_workers_blocked()
            st.markdown(
                f'<span class="{"admin-btn-1" if all_blk_d else "admin-btn-7"}" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "🔓 Unblock All Workers"
                if all_blk_d
                else "🔒 Block All Workers",
                key="q_btn_blkall",
                use_container_width=True,
            ):
                toggle_all_workers_block(not all_blk_d)
                st.rerun()

    elif active_admin_view == "records":
        rec_s_name = st.session_state.get("records_sheet", default_act_s)
        if rec_s_name not in sheets:
            rec_s_name = default_act_s
        df_r = all_sheets_map.get(rec_s_name, pd.DataFrame()).copy()
        df_r = reorder_dataframe_columns(df_r)
        login_c_r = find_login_user_column(df_r)
        display_c_r = find_display_username_column(df_r)
        p_col_r = find_password_column(df_r)
        visa_col_r = find_visa_type_column(df_r)
        lead_col_r = next(
            (c for c in df_r.columns if "lead" in c.strip().lower()),
            login_c_r,
        )

        if "records_selected_lead" not in st.session_state:
            st.session_state["records_selected_lead"] = "All"

        visa_counts_r = {}
        if visa_col_r and not df_r.empty:
            v_series = df_r[visa_col_r].astype(str).str.strip().str.upper()
            v_series = v_series[v_series != ""]
            visa_counts_r = v_series.value_counts().to_dict()

        st.markdown(
            "<div class='section-banner'>📂 Case Records & Visa Type Management Hub</div>",
            unsafe_allow_html=True,
        )

        st.markdown(
            "<div class='action-group-title'>👥 FILTER BY LEAD NAME:</div>",
            unsafe_allow_html=True,
        )
        lead_options_r = ["All", "Kiran", "Shanthi"]
        r_l_cols = st.columns([1.2, 1.2, 1.2, 4.4], gap="small")
        for l_idx_r, l_name_r in enumerate(lead_options_r):
            with r_l_cols[l_idx_r]:
                is_rl_act = st.session_state["records_selected_lead"] == l_name_r
                rl_cls = "admin-btn-0" if is_rl_act else "admin-btn-6"
                rl_lbl = f"✅ {l_name_r}" if is_rl_act else l_name_r
                st.markdown(
                    f'<span class="{rl_cls}" style="display:none;"></span>',
                    unsafe_allow_html=True,
                )
                if st.button(
                    rl_lbl,
                    key=f"records_lead_btn_{l_name_r}_{l_idx_r}",
                    use_container_width=True,
                ):
                    st.session_state["records_selected_lead"] = l_name_r
                    st.session_state[f"page_num_admin_{rec_s_name}"] = 1
                    st.rerun()

        if visa_counts_r:
            visa_items = list(visa_counts_r.items())
            w_cols = st.columns(min(len(visa_items), 6), gap="small")
            for idx, (v_name, v_count) in enumerate(visa_items):
                with w_cols[idx % len(w_cols)]:
                    st.markdown(
                        f'<span class="visa-color-{idx % 8}" style="display:none;"></span>',
                        unsafe_allow_html=True,
                    )
                    if st.button(
                        f"✈️ {v_name}: {v_count}",
                        key=f"admin_visa_card_btn_{rec_s_name}_{v_name}_{idx}",
                        use_container_width=True,
                    ):
                        st.session_state["selected_visa_filter"] = (
                            "All"
                            if st.session_state["selected_visa_filter"] == v_name
                            else v_name
                        )
                        st.session_state[f"page_num_admin_{rec_s_name}"] = 1
                        st.rerun()

            if st.session_state["selected_visa_filter"] != "All":
                if st.button(
                    f"🔄 Active Filter: [{st.session_state['selected_visa_filter']}] — Click to Reset",
                    use_container_width=True,
                    key="reset_visa_filter",
                ):
                    st.session_state["selected_visa_filter"] = "All"
                    st.session_state[f"page_num_admin_{rec_s_name}"] = 1
                    st.rerun()

        col_search, col_sheets = st.columns([1.2, 1.8], gap="medium")
        with col_search:
            admin_search = st.text_input(
                "Search main data (any field):",
                placeholder="Type Case ID, Username, DS160, Visa Type, Lead...",
                key="admin_search_input",
            ).strip()
        with col_sheets:
            if sheets:
                st.markdown("📑 **Switch Sheet:**")
                sw_cols_r = st.columns(len(sheets), gap="small")
                for s_idx, s_name in enumerate(sheets):
                    with sw_cols_r[s_idx]:
                        is_act = s_name == rec_s_name
                        color_c = "admin-btn-1" if is_act else "admin-btn-6"
                        b_txt = f"✅ {s_name}" if is_act else f"📄 {s_name}"
                        st.markdown(
                            f'<span class="{color_c}" style="display:none;"></span>',
                            unsafe_allow_html=True,
                        )
                        if st.button(
                            b_txt,
                            key=f"records_only_sheet_btn_{s_idx}",
                            use_container_width=True,
                        ):
                            st.session_state["records_sheet"] = s_name
                            st.session_state["active_sheet"] = s_name
                            st.session_state["selected_visa_filter"] = "All"
                            st.rerun()

        display_df = df_r
        if (
            st.session_state["records_selected_lead"] != "All"
            and not display_df.empty
            and lead_col_r in display_df.columns
        ):
            sel_l_r = st.session_state["records_selected_lead"].lower()
            display_df = display_df[
                display_df[lead_col_r]
                .astype(str)
                .str.strip()
                .str.lower()
                .str.contains(sel_l_r, na=False)
            ]

        if admin_search != "":
            display_df = display_df[
                search_main_data_any(display_df, admin_search)
            ]

        if st.session_state["selected_visa_filter"] != "All" and visa_col_r:
            target_visa = st.session_state["selected_visa_filter"]
            display_df = display_df[
                display_df[visa_col_r].astype(str).str.strip().str.upper()
                == target_visa
            ]

        render_clickable_table(
            display_df,
            display_c_r,
            p_col_r,
            table_key=f"admin_{rec_s_name}",
            sheet_for_link=rec_s_name,
            allow_edit=True,
        )

        if os.path.exists(EXCEL_FILE):
            with open(EXCEL_FILE, "rb") as f:
                excel_bytes = f.read()
            st.download_button(
                label="📥 Download Complete Excel File (All Sheets)",
                data=excel_bytes,
                file_name=EXCEL_FILE,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="admin_dl_all_btn",
            )

else:
    current_user = st.session_state["username"]
    all_sheets_map_w = load_all_sheets_cached(_get_file_mtime())
    sheets = (
        list(all_sheets_map_w.keys()) if all_sheets_map_w else ["Sheet1"]
    )
    default_act_s_w = get_default_active_sheet(sheets)
    current_sheet = st.session_state.get("active_sheet", default_act_s_w)
    if current_sheet not in sheets:
        current_sheet = default_act_s_w
        st.session_state["active_sheet"] = current_sheet

    perms = load_permissions()
    user_perms = perms.get(current_user, {})
    can_update = user_perms.get("can_update", True)
    can_add = user_perms.get("can_add", True)
    can_search = user_perms.get("can_search", True)
    can_download = user_perms.get("can_download", False)

    if "worker_selected_lead" not in st.session_state:
        st.session_state["worker_selected_lead"] = "All"

    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    w_top1, w_top2, w_top3, w_top4 = st.columns([5.4, 1.2, 1.2, 1.2], gap="small")
    with w_top1:
        st.markdown(
            f"<div class='worker-hero-banner'><span class='worker-hero-title'>👷 Case Worker: {current_user}</span><div class='worker-pill-group'><span class='worker-pill-gold'>📂 {current_sheet}</span><span class='worker-pill-cyan'>🎯 Lead: {st.session_state['worker_selected_lead']}</span></div></div>",
            unsafe_allow_html=True,
        )
    with w_top2:
        if can_download:
            w_dl_bytes = get_worker_export_bytes(
                _get_file_mtime(), current_user, current_sheet
            )
            st.download_button(
                label="📥 Excel",
                data=w_dl_bytes,
                file_name=f"{current_user}_Cases.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="worker_top_excel_dl",
            )
        else:
            st.button(
                "📥 Excel",
                key="worker_top_excel_dl_disabled",
                disabled=True,
                use_container_width=True,
            )
    with w_top3:
        if st.button("🔄 Refresh", key="worker_refresh", use_container_width=True):
            st.cache_data.clear()
            st.session_state.pop("_last_id_sync_mtime", None)
            act_s = get_default_active_sheet(get_sheet_names())
            st.session_state["active_sheet"] = act_s
            st.session_state["worker_selected_lead"] = "All"
            st.session_state["open_dialog"] = None
            st.session_state["selected_row_idx"] = None
            st.query_params["active_sheet"] = act_s
            st.rerun()
    with w_top4:
        if st.button("🚪 Logout", key="worker_logout", use_container_width=True):
            record_worker_logout(current_user)
            st.session_state["attendance_recorded"] = False
            st.query_params.clear()
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    notice_msg = get_admin_notice()
    if notice_msg:
        st.markdown(
            f'<div class="marquee-container"><marquee behavior="scroll" direction="left" scrollamount="6">📢 ADMIN NOTICE: {notice_msg}</marquee></div>',
            unsafe_allow_html=True,
        )

    w_c1, w_c2, w_c3, w_c4 = st.columns([1.3, 1.3, 1.4, 1.4], gap="small")
    with w_c1:
        if can_add:
            st.markdown(
                '<span class="admin-btn-1" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "➕ Add New Case",
                key="w_btn_add",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "add_case"
                st.rerun()
        else:
            st.button(
                "➕ Add New Case",
                key="w_btn_add_dis",
                disabled=True,
                use_container_width=True,
            )
    with w_c2:
        if can_search:
            st.markdown(
                '<span class="admin-btn-0" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                "🔍 Search Records",
                key="w_btn_search",
                use_container_width=True,
            ):
                st.session_state["open_dialog"] = "search_data"
                st.rerun()
        else:
            st.button(
                "🔍 Search Records",
                key="w_btn_search_dis",
                disabled=True,
                use_container_width=True,
            )
    with w_c3:
        st.markdown(
            '<span class="admin-btn-2" style="display:none;"></span>',
            unsafe_allow_html=True,
        )
        if st.button(
            "🔑 Change Password",
            key="w_btn_pass",
            use_container_width=True,
        ):
            st.session_state["open_dialog"] = "worker_password"
            st.rerun()
    with w_c4:
        if can_download:
            st.markdown(
                '<span class="admin-btn-5" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            w_dl_bytes_bar = get_worker_export_bytes(
                _get_file_mtime(), current_user, current_sheet
            )
            st.download_button(
                label="📥 Download Excel",
                data=w_dl_bytes_bar,
                file_name=f"{current_user}_Cases.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="worker_bar_excel_dl",
            )
        else:
            st.button(
                "📥 Download Excel",
                key="worker_bar_excel_dl_dis",
                disabled=True,
                use_container_width=True,
            )

    if sheets:
        st.markdown(
            "<div class='action-group-title'>📑 SWITCH SHEET:</div>",
            unsafe_allow_html=True,
        )
        w_sw_cols = st.columns(len(sheets), gap="small")
        palette_w = ["admin-btn-0", "admin-btn-1", "admin-btn-2", "admin-btn-3", "admin-btn-4", "admin-btn-5"]
        for s_idx_w, s_name_w in enumerate(sheets):
            with w_sw_cols[s_idx_w]:
                is_act_w = s_name_w == current_sheet
                c_cls_w = palette_w[s_idx_w % len(palette_w)] if is_act_w else "admin-btn-6"
                b_txt_w = f"✅ {s_name_w}" if is_act_w else f"📄 {s_name_w}"
                st.markdown(
                    f'<span class="{c_cls_w}" style="display:none;"></span>',
                    unsafe_allow_html=True,
                )
                if st.button(
                    b_txt_w,
                    key=f"worker_switch_sheet_{s_idx_w}",
                    use_container_width=True,
                ):
                    st.session_state["active_sheet"] = s_name_w
                    st.session_state[f"page_num_worker_{s_name_w}"] = 1
                    st.query_params["active_sheet"] = s_name_w
                    st.rerun()

    df_w = all_sheets_map_w.get(current_sheet, pd.DataFrame()).copy()
    df_w = reorder_dataframe_columns(df_w)
    login_c_w = find_login_user_column(df_w)
    display_c_w = find_display_username_column(df_w)
    p_col_w = find_password_column(df_w)
    lead_col_w = next(
        (c for c in df_w.columns if "lead" in c.strip().lower()),
        login_c_w,
    )

    user_rows_w = (
        df_w[
            df_w[login_c_w].astype(str).str.strip().str.lower()
            == current_user.lower()
        ].copy()
        if not df_w.empty
        else pd.DataFrame()
    )

    st.markdown(
        "<div class='action-group-title'>👥 FILTER BY LEAD NAME:</div>",
        unsafe_allow_html=True,
    )
    lead_options_w = ["All", "Kiran", "Shanthi"]
    wl_cols = st.columns([1.2, 1.2, 1.2, 4.4], gap="small")
    for l_idx_w, l_name_w in enumerate(lead_options_w):
        with wl_cols[l_idx_w]:
            is_wl_act = st.session_state["worker_selected_lead"] == l_name_w
            wl_cls = "admin-btn-0" if is_wl_act else "admin-btn-6"
            wl_lbl = f"✅ {l_name_w}" if is_wl_act else l_name_w
            st.markdown(
                f'<span class="{wl_cls}" style="display:none;"></span>',
                unsafe_allow_html=True,
            )
            if st.button(
                wl_lbl,
                key=f"worker_lead_btn_{l_name_w}_{l_idx_w}",
                use_container_width=True,
            ):
                st.session_state["worker_selected_lead"] = l_name_w
                st.session_state[f"page_num_worker_{current_sheet}"] = 1
                st.rerun()

    worker_search_q = st.text_input(
        "🔍 Search your records:",
        placeholder="Type Case ID, Username, DS160, Lead...",
        key="worker_records_search_input",
    ).strip()

    if (
        st.session_state["worker_selected_lead"] != "All"
        and not user_rows_w.empty
        and lead_col_w in user_rows_w.columns
    ):
        sel_lw = st.session_state["worker_selected_lead"].lower()
        user_rows_w = user_rows_w[
            user_rows_w[lead_col_w]
            .astype(str)
            .str.strip()
            .str.lower()
            .str.contains(sel_lw, na=False)
        ]

    if worker_search_q != "" and not user_rows_w.empty:
        user_rows_w = user_rows_w[
            search_main_data_any(user_rows_w, worker_search_q)
        ]

    total_w_cases = len(user_rows_w) if not user_rows_w.empty else 0
    today_str_w = get_ist_now().strftime("%Y-%m-%d")
    today_w_cases = 0
    booked_resched_w = 0

    if not user_rows_w.empty:
        date_cols_w = [
            c
            for c in user_rows_w.columns
            if any(k in c.lower() for k in ["date", "time", "day", "created"])
        ]
        for _, r in user_rows_w.iterrows():
            if any(
                today_str_w in str(r.get(dc, "")).strip() for dc in date_cols_w
            ) or today_str_w in str(r.get("Updated_At", "")):
                today_w_cases += 1

            bs_val = str(r.get("Booking_Status", "Active")).strip().lower()
            if bs_val in ["booked", "rescheduled"]:
                booked_resched_w += 1

    st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
    wk1, wk2, wk3 = st.columns(3, gap="medium")
    with wk1:
        st.markdown(
            f"<div class='kpi-card teal'><div class='kpi-title'>Total Cases ({current_sheet})</div><div class='kpi-value'>{total_w_cases}</div></div>",
            unsafe_allow_html=True,
        )
    with wk2:
        st.markdown(
            f"<div class='kpi-card green'><div class='kpi-title'>Today's Cases</div><div class='kpi-value'>{today_w_cases}</div></div>",
            unsafe_allow_html=True,
        )
    with wk3:
        st.markdown(
            f"<div class='kpi-card gold'><div class='kpi-title'>Booked / Rescheduled</div><div class='kpi-value'>{booked_resched_w}</div></div>",
            unsafe_allow_html=True,
        )

    st.markdown(
        f"<div class='section-banner'>📋 Your Assigned Cases — {current_sheet} ({len(user_rows_w)})</div>",
        unsafe_allow_html=True,
    )
    render_clickable_table(
        user_rows_w,
        display_c_w,
        p_col_w,
        table_key=f"worker_{current_sheet}",
        sheet_for_link=current_sheet,
        allow_edit=True,
    )
