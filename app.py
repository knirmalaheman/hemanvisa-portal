import streamlit as st
import pandas as pd
import os
import io
import datetime
import base64
import json
import math
import re
import zipfile

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


def _read_excel_sheets_safely(excel_path):
    if not os.path.exists(excel_path):
        return {}
    try:
        if os.path.getsize(excel_path) == 0:
            raise zipfile.BadZipFile("0-byte Excel file")
        return pd.read_excel(excel_path, sheet_name=None, dtype=str)
    except Exception:
        try:
            stamp = get_ist_now().strftime("%Y%m%d_%H%M%S")
            os.replace(excel_path, f"{excel_path}.corrupt_{stamp}.bak")
        except Exception:
            pass
        return {}


def _write_excel_sheets_safely(excel_path, sheets_dict):
    parent_dir = os.path.dirname(os.path.abspath(excel_path))
    if parent_dir and not os.path.exists(parent_dir):
        os.makedirs(parent_dir, exist_ok=True)
    tmp_path = f"{excel_path}.tmp"
    with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
        for s, df_val in sheets_dict.items():
            ordered_df = reorder_dataframe_columns(df_val)
            ordered_df.to_excel(writer, sheet_name=str(s), index=False)
    os.replace(tmp_path, excel_path)


@st.cache_data(ttl=600, show_spinner=False)
def load_all_sheets_cached(file_mtime):
    if os.path.exists(EXCEL_FILE):
        try:
            all_sheets = _read_excel_sheets_safely(EXCEL_FILE)
            if not all_sheets:
                return {}
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
        all_sheets = _read_excel_sheets_safely(EXCEL_FILE)
        if not all_sheets:
            return
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
            _write_excel_sheets_safely(EXCEL_FILE, all_sheets)
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


def _resolve_target_active_sheet(sheets_all, assigned_worker=None):
    if assigned_worker and str(assigned_worker).strip().lower() == "anil":
        for s_name in sheets_all:
            s_low = s_name.lower()
            if "anil" in s_low and "active" in s_low:
                return s_name
        for s_name in sheets_all:
            if "anil" in s_name.lower() and "booked" not in s_name.lower():
                return s_name
    return get_default_active_sheet(sheets_all)


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
    active_sheet_name = _resolve_target_active_sheet(sheets_all, assigned_worker=assigned_worker)
    if not active_sheet_name:
        return None, None
    df_active = load_data(sheet_name=active_sheet_name)
    if df_active is None:
        return None, None

    new_rescheduled_id = generate_next_case_id()
    row_to_active = source_row.copy()
    row_to_active["Case_ID"] = new_rescheduled_id

    if assigned_worker is not None:
        raw_w = str(assigned_worker).strip()
        clean_worker = "" if raw_w.lower() in ["", "not assigned"] else raw_w
        login_c_act = find_login_user_column(df_active)
        if clean_worker and not df_active.empty and login_c_act in df_active.columns:
            existing_vals = [str(x).strip() for x in df_active[login_c_act].dropna().unique() if str(x).strip()]
            exact_ci = next((ev for ev in existing_vals if ev.lower() == clean_worker.lower()), None)
            if exact_ci:
                clean_worker = exact_ci
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

    sheets_data = _read_excel_sheets_safely(EXCEL_FILE)
    df_to_save = reorder_dataframe_columns(df_to_save)
    sheets_data[sheet_name] = df_to_save
    _write_excel_sheets_safely(EXCEL_FILE, sheets_data)

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
                if u != "" and u.lower() not in ["", "admin", "not assigned", "nan", "none"]:
                    all_workers.add(u)
    perms_map = load_permissions()
    for p_user in perms_map.keys():
        u_clean = str(p_user).strip()
        if u_clean and u_clean.lower() not in ["admin", "not assigned"]:
            all_workers.add(u_clean)
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
