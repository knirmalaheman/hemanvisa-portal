import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import requests
import os
import io
import datetime
import base64
import json
import math
import re
import urllib.parse

EXCEL_FILE = "Heman_Report.xlsx"
ADMIN_FILE = "admin_config.txt"
PERMS_FILE = "worker_permissions.json"
NOTICE_FILE = "admin_notice.txt"
LOGINS_FILE = "worker_logins_status.json"
ATTENDANCE_FILE = "worker_attendance.json"
PDF_BASE_DIR = "uploaded_appointment_pdfs"
LOGO_FILES = ["pngTShirtPrintCricket.jpg", "logo.jpg", "logo.png", "logo.jpeg"]

SYSTEM_SHEETS = {"App_Logins", "App_Attendance", "App_Permissions", "App_Config"}

if not os.path.exists(PDF_BASE_DIR):
    os.makedirs(PDF_BASE_DIR)

st.set_page_config(page_title="HemanVisa Portal", page_icon="✈️", layout="wide")


@st.cache_data(ttl=3600, show_spinner=False)
def get_webapp_url():
    try:
        url = st.secrets.get("GSHEET_WEBAPP_URL", "")
        return str(url).strip() if url else ""
    except Exception:
        return ""


def _rows_to_df(raw_rows):
    if not raw_rows or not isinstance(raw_rows, list) or len(raw_rows) == 0:
        return pd.DataFrame()
    headers = [
        str(h).strip() if str(h).strip() != "" else f"Col_{i}"
        for i, h in enumerate(raw_rows[0])
    ]
    seen = {}
    unique_headers = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            unique_headers.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            unique_headers.append(h)
    data_rows = raw_rows[1:] if len(raw_rows) > 1 else []
    n_cols = len(unique_headers)
    cleaned_rows = []
    for r in data_rows:
        row_list = list(r) if isinstance(r, list) else [r]
        if len(row_list) < n_cols:
            row_list = row_list + [""] * (n_cols - len(row_list))
        elif len(row_list) > n_cols:
            row_list = row_list[:n_cols]
        cleaned_rows.append(["" if v is None else str(v) for v in row_list])
    return pd.DataFrame(cleaned_rows, columns=unique_headers, dtype=str)


def _fetch_sheets_from_webapp():
    url = get_webapp_url()
    if not url:
        return None
    try:
        resp = requests.get(url, timeout=12, allow_redirects=True)
        resp.raise_for_status()
        text_body = resp.text.strip()
        if not text_body:
            return None
        payload = json.loads(text_body)
        if isinstance(payload, dict) and isinstance(payload.get("sheets"), dict):
            sheets_dict = {}
            for s_name, s_rows in payload["sheets"].items():
                sheets_dict[str(s_name).strip()] = _rows_to_df(s_rows)
            return sheets_dict if sheets_dict else None
        if isinstance(payload, list):
            df_main = _rows_to_df(payload)
            return {"Active Cases": df_main}
        return None
    except Exception as e:
        st.session_state["_last_webapp_fetch_error"] = str(e)
        return None


def _push_sheet_to_webapp(sheet_name, df_to_push, all_sheets_dict=None):
    url = get_webapp_url()
    if not url:
        st.error("GSHEET_WEBAPP_URL కనిపించడం లేదు!")
        return
    try:
        df_clean = df_to_push.fillna("").astype(str)
        rows_2d = [df_clean.columns.tolist()] + df_clean.values.tolist()
        body = {
            "action": "save_sheet",
            "sheet_name": str(sheet_name),
            "data": rows_2d,
        }
        if all_sheets_dict:
            all_payload = {}
            for s_k, df_k in all_sheets_dict.items():
                df_kc = df_k.fillna("").astype(str)
                all_payload[str(s_k)] = [df_kc.columns.tolist()] + df_kc.values.tolist()
            body["sheets"] = all_payload

        resp = requests.post(url, json=body, timeout=12, allow_redirects=True)
        resp.raise_for_status()
        res_json = resp.json()
        if isinstance(res_json, dict) and res_json.get("status") != "success":
            st.error(f"Google Sheet Save Error ({sheet_name}): {res_json}")
    except Exception as e:
        st.error(f"Google Sheet Push Error ({sheet_name}): {e}")


def get_ist_now():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)


def _get_file_mtime():
    sync_ver = st.session_state.get("_gsheet_sync_version", 0)
    f_mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
    return f"{f_mtime}_{sync_ver}"


def _read_sheet_rows(sheet_name, sheets_cache=None):
    sheets = sheets_cache if isinstance(sheets_cache, dict) else load_all_sheets_cached(_get_file_mtime())
    df = sheets.get(sheet_name) if isinstance(sheets, dict) else None
    if df is None or df.empty:
        return []
    return df.fillna("").astype(str).to_dict(orient="records")


def _write_sheet_records(sheet_name, columns, records):
    if records:
        df_new = pd.DataFrame(records)
        for col in columns:
            if col not in df_new.columns:
                df_new[col] = ""
        df_new = df_new[columns].fillna("").astype(str)
    else:
        df_new = pd.DataFrame(columns=columns, dtype=str)

    _push_sheet_to_webapp(sheet_name, df_new, all_sheets_dict=None)
    st.session_state["_gsheet_sync_version"] = (
        st.session_state.get("_gsheet_sync_version", 0) + 1
    )
    st.cache_data.clear()


def get_admin_creds(sheets_cache=None):
    rows = _read_sheet_rows("App_Config", sheets_cache)
    for r in rows:
        if r.get("Key", "").strip() == "admin_creds":
            u = r.get("Value1", "").strip() or "admin"
            p = r.get("Value2", "").strip() or "admin123"
            return u, p
    return "admin", "admin123"


def save_admin_creds(new_user, new_pass):
    rows = _read_sheet_rows("App_Config")
    updated = False
    for r in rows:
        if r.get("Key", "").strip() == "admin_creds":
            r["Value1"] = str(new_user).strip()
            r["Value2"] = str(new_pass).strip()
            updated = True
    if not updated:
        rows.append({
            "Key": "admin_creds",
            "Value1": str(new_user).strip(),
            "Value2": str(new_pass).strip(),
        })
    _write_sheet_records("App_Config", ["Key", "Value1", "Value2"], rows)


def get_admin_notice(sheets_cache=None):
    rows = _read_sheet_rows("App_Config", sheets_cache)
    for r in rows:
        if r.get("Key", "").strip() == "admin_notice":
            return r.get("Value1", "").strip()
    return ""


def save_admin_notice(notice_text):
    rows = _read_sheet_rows("App_Config")
    updated = False
    for r in rows:
        if r.get("Key", "").strip() == "admin_notice":
            r["Value1"] = str(notice_text).strip()
            updated = True
    if not updated:
        rows.append({
            "Key": "admin_notice",
            "Value1": str(notice_text).strip(),
            "Value2": "",
        })
    _write_sheet_records("App_Config", ["Key", "Value1", "Value2"], rows)


def load_permissions(sheets_cache=None):
    rows = _read_sheet_rows("App_Permissions", sheets_cache)
    perms = {}
    for r in rows:
        w = r.get("Worker", "").strip()
        if w:
            perms[w] = {
                "can_edit": str(r.get("can_edit", "True")).strip().lower()
                in ["true", "1", "yes"],
                "can_update": str(r.get("can_update", r.get("can_edit", "True"))).strip().lower()
                in ["true", "1", "yes"],
                "can_delete_req": str(r.get("can_delete_req", "True")).strip().lower()
                in ["true", "1", "yes"],
                "can_reschedule": str(r.get("can_reschedule", "True")).strip().lower()
                in ["true", "1", "yes"],
                "can_add": str(r.get("can_add", "True")).strip().lower()
                in ["true", "1", "yes"],
                "can_search": str(r.get("can_search", "True")).strip().lower()
                in ["true", "1", "yes"],
                "can_download": str(r.get("can_download", "False")).strip().lower()
                in ["true", "1", "yes"],
                "can_sync": str(r.get("can_sync", "False")).strip().lower()
                in ["true", "1", "yes"],
                "block_login": str(r.get("block_login", "False")).strip().lower()
                in ["true", "1", "yes"],
            }
    return perms


def save_permissions(perms_dict):
    cols = [
        "Worker",
        "can_edit",
        "can_update",
        "can_delete_req",
        "can_reschedule",
        "can_add",
        "can_search",
        "can_download",
        "can_sync",
        "block_login",
    ]
    records = []
    for w, p in (perms_dict or {}).items():
        if str(w).strip():
            records.append({
                "Worker": str(w).strip(),
                "can_edit": str(p.get("can_edit", p.get("can_update", True))),
                "can_update": str(p.get("can_update", p.get("can_edit", True))),
                "can_delete_req": str(p.get("can_delete_req", True)),
                "can_reschedule": str(p.get("can_reschedule", True)),
                "can_add": str(p.get("can_add", True)),
                "can_search": str(p.get("can_search", True)),
                "can_download": str(p.get("can_download", False)),
                "can_sync": str(p.get("can_sync", False)),
                "block_login": str(p.get("block_login", False)),
            })
    _write_sheet_records("App_Permissions", cols, records)


def _normalize_iso_timestamp_to_ist(val_str):
    s = str(val_str).strip()
    if not s or s == "-":
        return ""
    iso_match = re.match(
        r"^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2}:\d{2})(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$",
        s,
        re.IGNORECASE,
    )
    if iso_match:
        d_part, t_part = iso_match.groups()
        try:
            dt_utc = datetime.datetime.strptime(f"{d_part} {t_part}", "%Y-%m-%d %H:%M:%S")
            if d_part == "1899-12-30":
                dt_ist_time = dt_utc + datetime.timedelta(hours=5, minutes=30)
                return dt_ist_time.strftime("%I:%M:%S %p")
            dt_ist = dt_utc + datetime.timedelta(hours=5, minutes=30)
            return dt_ist.strftime("%Y-%m-%d %I:%M:%S %p")
        except Exception:
            pass
    return s


def _normalize_date_ist(val_str):
    s = str(val_str).strip()
    if not s or s == "-":
        return ""
    iso_match = re.match(
        r"^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2}:\d{2})(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$",
        s,
        re.IGNORECASE,
    )
    if iso_match:
        d_part, t_part = iso_match.groups()
        try:
            dt_utc = datetime.datetime.strptime(f"{d_part} {t_part}", "%Y-%m-%d %H:%M:%S")
            dt_ist = dt_utc + datetime.timedelta(hours=5, minutes=30)
            return dt_ist.strftime("%Y-%m-%d")
        except Exception:
            pass
    return s[:10] if len(s) >= 10 else s


def _normalize_time_ist(val_str):
    s = str(val_str).strip()
    if not s or s == "-":
        return "-"
    norm_full = _normalize_iso_timestamp_to_ist(s)
    if " " in norm_full and re.match(r"^\d{4}-\d{2}-\d{2}\s+", norm_full):
        return norm_full.split(" ", 1)[1]
    return norm_full


def load_login_tracking(sheets_cache=None):
    rows = _read_sheet_rows("App_Logins", sheets_cache)
    logins = {}
    for r in rows:
        w = r.get("Worker", "").strip()
        if w:
            raw_last = r.get("last_login", "-").strip() or "-"
            clean_last = _normalize_iso_timestamp_to_ist(raw_last) if raw_last != "-" else "-"
            logins[w] = {
                "status": r.get("status", "Not Logged In").strip() or "Not Logged In",
                "last_login": clean_last or "-",
            }
    return logins


def _save_login_tracking(logins_dict):
    cols = ["Worker", "status", "last_login"]
    records = []
    for w, info in (logins_dict or {}).items():
        if str(w).strip():
            raw_last = str(info.get("last_login", "-")).strip() or "-"
            clean_last = _normalize_iso_timestamp_to_ist(raw_last) if raw_last != "-" else "-"
            records.append({
                "Worker": str(w).strip(),
                "status": str(info.get("status", "Not Logged In")),
                "last_login": clean_last or "-",
            })
    _write_sheet_records("App_Logins", cols, records)


ATTENDANCE_COLS = ["S.No", "Worker", "Date", "Login_Time", "Logout_Time", "Duration", "Status"]


def _parse_time_12_or_24(t_str):
    s = str(t_str).strip()
    if not s or s == "-":
        return None
    norm_t = _normalize_time_ist(s)
    if not norm_t or norm_t == "-":
        return None
    for fmt in ("%I:%M:%S %p", "%I:%M %p", "%H:%M:%S", "%H:%M"):
        try:
            return datetime.datetime.strptime(norm_t, fmt).time()
        except Exception:
            continue
    return None


def _compute_duration_seconds(date_str, login_str, logout_str, is_active=False):
    t_in = _parse_time_12_or_24(login_str)
    if not t_in:
        return 0
    try:
        d_clean = _normalize_date_ist(date_str)
        d_base = (
            datetime.datetime.strptime(d_clean, "%Y-%m-%d").date()
            if d_clean
            else get_ist_now().date()
        )
    except Exception:
        d_base = get_ist_now().date()

    dt_in = datetime.datetime.combine(d_base, t_in)

    t_out = _parse_time_12_or_24(logout_str)
    if t_out:
        dt_out = datetime.datetime.combine(d_base, t_out)
        if dt_out < dt_in:
            dt_out = dt_out + datetime.timedelta(days=1)
        return max(0, int((dt_out - dt_in).total_seconds()))

    if is_active:
        now_ist = get_ist_now()
        dt_now = datetime.datetime.combine(d_base, now_ist.time())
        if d_base < now_ist.date():
            dt_now = now_ist
        elif dt_now < dt_in:
            dt_now = dt_now + datetime.timedelta(days=1)
        return max(0, int((dt_now - dt_in).total_seconds()))

    return 0


def _format_seconds_hhmm(total_seconds):
    sec = max(0, int(total_seconds))
    hrs = sec // 3600
    mins = (sec % 3600) // 60
    return f"{hrs}h {mins}m"


def load_attendance(sheets_cache=None):
    rows = _read_sheet_rows("App_Attendance", sheets_cache)
    cleaned = []
    sno = 1
    for r in rows:
        w = str(r.get("Worker", "")).strip()
        if w:
            d_clean = _normalize_date_ist(r.get("Date", ""))
            l_clean = _normalize_time_ist(r.get("Login_Time", "-"))
            lo_clean = _normalize_time_ist(r.get("Logout_Time", "-"))
            raw_st = str(r.get("Status", "")).strip()
            st_clean = "Logged Out" if lo_clean != "-" else (raw_st if raw_st else "Active")
            is_act = st_clean.lower() in ["active", "logged in", "online"]
            dur_sec = _compute_duration_seconds(d_clean, l_clean, lo_clean, is_active=is_act)
            dur_str = _format_seconds_hhmm(dur_sec) if (l_clean != "-" and (lo_clean != "-" or is_act)) else "0h 0m"
            cleaned.append({
                "S.No": sno,
                "Worker": w,
                "Date": d_clean,
                "Login_Time": l_clean,
                "Logout_Time": lo_clean,
                "Duration": dur_str,
                "Status": st_clean,
                "_duration_sec": dur_sec,
            })
            sno += 1
    return cleaned


def save_attendance(att_list):
    records = []
    sno = 1
    for item in (att_list or []):
        w = str(item.get("Worker", "")).strip()
        if w:
            d_clean = _normalize_date_ist(item.get("Date", ""))
            l_clean = _normalize_time_ist(item.get("Login_Time", "-"))
            lo_clean = _normalize_time_ist(item.get("Logout_Time", "-"))
            raw_st = str(item.get("Status", "")).strip()
            st_clean = "Logged Out" if lo_clean != "-" else (raw_st if raw_st else "Active")
            is_act = st_clean.lower() in ["active", "logged in", "online"]
            dur_sec = _compute_duration_seconds(d_clean, l_clean, lo_clean, is_active=is_act)
            dur_str = _format_seconds_hhmm(dur_sec) if (l_clean != "-" and (lo_clean != "-" or is_act)) else "0h 0m"
            records.append({
                "S.No": sno,
                "Worker": w,
                "Date": d_clean,
                "Login_Time": l_clean,
                "Logout_Time": lo_clean,
                "Duration": dur_str,
                "Status": st_clean,
            })
            sno += 1
    _write_sheet_records("App_Attendance", ATTENDANCE_COLS, records)


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
    _save_login_tracking(logins)

    # ప్రతి కొత్త login ఒక స్వతంత్ర session గా యాడ్ అవుతుంది (పాత sessions duration చెక్కుచెదరదు)
    attendance = load_attendance()
    attendance.append({
        "Worker": worker_name.strip(),
        "Date": today_str,
        "Login_Time": time_str,
        "Logout_Time": "-",
        "Duration": "0h 0m",
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
        if k.strip().lower() == worker_name.strip().lower():
            logins[k]["status"] = "Not Logged In"
    _save_login_tracking(logins)

    attendance = load_attendance()
    updated_any = False

    # 1. నేటి open session (Logout_Time == "-") ని మాత్రమే క్లోజ్ చేసి ఆ session duration లెక్కించడం
    for entry in reversed(attendance):
        entry_date = _normalize_date_ist(entry.get("Date", ""))
        if (
            entry.get("Worker", "").strip().lower() == worker_name.strip().lower()
            and entry_date == today_str
            and _normalize_time_ist(entry.get("Logout_Time", "-")) == "-"
        ):
            entry["Date"] = today_str
            entry["Logout_Time"] = time_str
            entry["Status"] = "Logged Out"
            dur_sec = _compute_duration_seconds(today_str, entry.get("Login_Time", "-"), time_str, is_active=False)
            entry["Duration"] = _format_seconds_hhmm(dur_sec)
            updated_any = True
            break

    # 2. ఏదైనా open session (Logout_Time == "-") ఉంటే క్లోజ్ చేయడం
    if not updated_any:
        for entry in reversed(attendance):
            if (
                entry.get("Worker", "").strip().lower() == worker_name.strip().lower()
                and _normalize_time_ist(entry.get("Logout_Time", "-")) == "-"
            ):
                e_date = _normalize_date_ist(entry.get("Date", "")) or today_str
                entry["Logout_Time"] = time_str
                entry["Status"] = "Logged Out"
                dur_sec = _compute_duration_seconds(e_date, entry.get("Login_Time", "-"), time_str, is_active=False)
                entry["Duration"] = _format_seconds_hhmm(dur_sec)
                updated_any = True
                break

    if not updated_any:
        attendance.append({
            "Worker": worker_name.strip(),
            "Date": today_str,
            "Login_Time": "-",
            "Logout_Time": time_str,
            "Duration": "0h 0m",
            "Status": "Logged Out",
        })
    save_attendance(attendance)


def build_attendance_summary_tables(attendance_data):
    if not attendance_data:
        return pd.DataFrame(), pd.DataFrame()

    worker_date_map = {}
    worker_overall_map = {}

    for row in attendance_data:
        w = str(row.get("Worker", "")).strip() or "Unknown"
        d = str(row.get("Date", "")).strip() or "Unknown"
        l_t = str(row.get("Login_Time", "-")).strip() or "-"
        lo_t = str(row.get("Logout_Time", "-")).strip() or "-"
        st = str(row.get("Status", "")).strip() or ("Logged Out" if lo_t != "-" else "Active")
        sec = int(row.get("_duration_sec", 0) or 0)

        wd_key = (w, d)
        if wd_key not in worker_date_map:
            worker_date_map[wd_key] = {
                "Worker": w,
                "Date": d,
                "Sessions": 0,
                "First_Login": l_t if l_t != "-" else None,
                "Last_Logout": lo_t if lo_t != "-" else None,
                "Total_Seconds": 0,
                "Current_Status": st,
            }
        wd = worker_date_map[wd_key]
        wd["Sessions"] += 1
        if l_t != "-" and wd["First_Login"] is None:
            wd["First_Login"] = l_t
        if lo_t != "-":
            wd["Last_Logout"] = lo_t
        wd["Total_Seconds"] += sec
        wd["Current_Status"] = st

        if w not in worker_overall_map:
            worker_overall_map[w] = {
                "Worker": w,
                "Active_Days": set(),
                "Total_Sessions": 0,
                "Last_Login": l_t if l_t != "-" else "-",
                "Last_Logout": lo_t if lo_t != "-" else "-",
                "Total_Seconds": 0,
                "Current_Status": st,
            }
        wo = worker_overall_map[w]
        if d and d != "Unknown":
            wo["Active_Days"].add(d)
        wo["Total_Sessions"] += 1
        if l_t != "-":
            wo["Last_Login"] = l_t
        if lo_t != "-":
            wo["Last_Logout"] = lo_t
        wo["Total_Seconds"] += sec
        wo["Current_Status"] = st

    day_rows = []
    for idx_d, (_, wd) in enumerate(worker_date_map.items(), start=1):
        tot_sec = wd["Total_Seconds"]
        day_rows.append({
            "S.No": idx_d,
            "Worker": wd["Worker"],
            "Date": wd["Date"],
            "Sessions": wd["Sessions"],
            "First_Login": wd["First_Login"] or "-",
            "Last_Logout": wd["Last_Logout"] or "-",
            "Total_Hours": round(tot_sec / 3600.0, 2),
            "Total_Duration": _format_seconds_hhmm(tot_sec),
            "Status": wd["Current_Status"],
        })

    overall_rows = []
    grand_sec = 0
    grand_sessions = 0
    for idx_o, (_, wo) in enumerate(worker_overall_map.items(), start=1):
        tot_sec = wo["Total_Seconds"]
        grand_sec += tot_sec
        grand_sessions += wo["Total_Sessions"]
        overall_rows.append({
            "S.No": idx_o,
            "Worker": wo["Worker"],
            "Days_Worked": len(wo["Active_Days"]),
            "Total_Sessions": wo["Total_Sessions"],
            "Last_Login": wo["Last_Login"],
            "Last_Logout": wo["Last_Logout"],
            "Total_Hours": round(tot_sec / 3600.0, 2),
            "Total_Duration": _format_seconds_hhmm(tot_sec),
            "Status": wo["Current_Status"],
        })

    if overall_rows:
        overall_rows.append({
            "S.No": len(overall_rows) + 1,
            "Worker": "OVERALL TOTAL",
            "Days_Worked": "-",
            "Total_Sessions": grand_sessions,
            "Last_Login": "-",
            "Last_Logout": "-",
            "Total_Hours": round(grand_sec / 3600.0, 2),
            "Total_Duration": _format_seconds_hhmm(grand_sec),
            "Status": "-",
        })

    df_day = pd.DataFrame(day_rows)
    df_overall = pd.DataFrame(overall_rows)
    return df_day, df_overall


def close_case_popup_state():
    st.session_state["is_editing_popup"] = False
    st.session_state["is_delete_request_mode"] = False
    st.session_state["is_reschedule_mode"] = False
    st.session_state["selected_row_idx"] = None
    st.session_state["open_dialog"] = None
    if "edit_row" in st.query_params:
        try:
            del st.query_params["edit_row"]
        except Exception:
            pass


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

    if "Sheet_Name" in cols:
        found_priority.append("Sheet_Name")
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


def _detect_case_username_col(df):
    if df is None or df.empty:
        return None
    for col in df.columns:
        if col.strip() == "USERNAME":
            return col
    login_c = None
    for col in df.columns:
        if col.strip().lower() in ["username", "case worker", "cases_worker", "user"]:
            if col.strip() != "USERNAME":
                login_c = col
                break
    for col in df.columns:
        if col.strip().lower() == "username" and col != login_c:
            return col
    for col in df.columns:
        c_low = col.strip().lower().replace("_", " ")
        if "case" in c_low and "user" in c_low:
            return col
    return None


def _is_valid_case_username(val):
    s = str(val).strip().lower()
    return s not in ["", "nan", "none", "not assigned", "unassigned", "-"]


def _assign_case_username_based_ids(all_sheets):
    max_num = 1000
    username_to_id = {}
    modified = False

    for s_name, df in all_sheets.items():
        if df is None or str(s_name).strip() in SYSTEM_SHEETS:
            continue
        df.columns = df.columns.str.strip()
        disp_col = _detect_case_username_col(df)
        if "Case_ID" in df.columns:
            for idx in df.index:
                cid = str(df.at[idx, "Case_ID"]).strip()
                if cid and cid.lower() != "nan":
                    m = re.search(r"(\d+)", cid)
                    if m:
                        num = int(m.group(1))
                        if num > max_num:
                            max_num = num
                    if disp_col and disp_col in df.columns:
                        u_val = str(df.at[idx, disp_col]).strip()
                        if _is_valid_case_username(u_val):
                            u_key = u_val.lower()
                            if u_key not in username_to_id:
                                username_to_id[u_key] = cid

    for s_name, df in all_sheets.items():
        if df is None or str(s_name).strip() in SYSTEM_SHEETS:
            continue
        df.columns = df.columns.str.strip()
        df = df.fillna("")
        if "Case_ID" not in df.columns:
            df.insert(0, "Case_ID", "")
            modified = True

        disp_col = _detect_case_username_col(df)
        for idx in df.index:
            cid = str(df.at[idx, "Case_ID"]).strip()
            if cid == "" or cid.lower() == "nan":
                u_key = ""
                if disp_col and disp_col in df.columns:
                    u_val = str(df.at[idx, disp_col]).strip()
                    if _is_valid_case_username(u_val):
                        u_key = u_val.lower()

                if u_key and u_key in username_to_id:
                    df.at[idx, "Case_ID"] = username_to_id[u_key]
                else:
                    max_num += 1
                    new_cid = f"HM{max_num}"
                    df.at[idx, "Case_ID"] = new_cid
                    if u_key:
                        username_to_id[u_key] = new_cid
                modified = True

        all_sheets[s_name] = df

    return all_sheets, modified


def _normalize_sheet_df(df):
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
    if "Issue_Remark" not in df.columns:
        df["Issue_Remark"] = ""

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

    return reorder_dataframe_columns(df)


@st.cache_data(ttl=900, show_spinner=False)
def load_all_sheets_cached(file_mtime):
    remote_sheets = _fetch_sheets_from_webapp()
    all_sheets = None

    if remote_sheets is not None:
        if len(remote_sheets) == 1 and os.path.exists(EXCEL_FILE):
            try:
                local_raw = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
                remote_key = list(remote_sheets.keys())[0]
                remote_df = remote_sheets[remote_key]
                merged = {}
                if remote_df is not None and not remote_df.empty:
                    merged[remote_key] = remote_df
                for l_s, l_df in local_raw.items():
                    if str(l_s) not in merged:
                        merged[str(l_s)] = l_df
                all_sheets = merged if merged else remote_sheets
            except Exception:
                all_sheets = remote_sheets
        else:
            all_sheets = remote_sheets
    elif os.path.exists(EXCEL_FILE):
        try:
            all_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
        except Exception:
            return {}

    if not all_sheets:
        return {}

    try:
        all_sheets, _ = _assign_case_username_based_ids(all_sheets)
        processed = {}
        for s_name, df in all_sheets.items():
            if df is None:
                continue
            s_clean = str(s_name).strip()
            if s_clean in SYSTEM_SHEETS:
                df_sys = df.copy()
                df_sys.columns = df_sys.columns.str.strip()
                processed[s_clean] = df_sys.fillna("").astype(str)
            else:
                processed[s_clean] = _normalize_sheet_df(df)
        return processed
    except Exception:
        return {}


def ensure_all_cases_have_unique_ids(force=False):
    mtime = _get_file_mtime()
    last_checked_key = "_last_id_sync_mtime"
    if not force and st.session_state.get(last_checked_key) == mtime:
        return
    try:
        all_sheets = load_all_sheets_cached(mtime)
        if not all_sheets:
            return
        all_sheets, modified = _assign_case_username_based_ids(all_sheets)
        for s_name, df in all_sheets.items():
            all_sheets[s_name] = reorder_dataframe_columns(df)

        if modified:
            with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                for s, df_val in all_sheets.items():
                    df_val.to_excel(writer, sheet_name=str(s), index=False)
            first_s = list(all_sheets.keys())[0]
            _push_sheet_to_webapp(first_s, all_sheets[first_s], all_sheets)
            st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
            st.cache_data.clear()
        st.session_state[last_checked_key] = _get_file_mtime()
    except Exception:
        pass


def get_sheet_names(sheets_cache=None):
    all_sheets = sheets_cache if isinstance(sheets_cache, dict) else load_all_sheets_cached(_get_file_mtime())
    if all_sheets:
        visible = [s for s in all_sheets.keys() if str(s).strip() not in SYSTEM_SHEETS]
        return visible if visible else list(all_sheets.keys())
    return ["Sheet1"]


def get_default_active_sheet(sheets_list=None):
    if sheets_list is None:
        sheets_list = get_sheet_names()
    filtered = [s for s in sheets_list if str(s).strip() not in SYSTEM_SHEETS]
    pool = filtered if filtered else sheets_list
    if not pool:
        return "Sheet1"
    return next(
        (
            s
            for s in pool
            if "booked" not in s.lower() and "issue" not in s.lower()
        ),
        pool[0],
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
    return f"HM{max_num + 1}"


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
    row_to_active = {col: "" for col in df_active.columns}

    for col in df_active.columns:
        if col in source_row.index:
            row_to_active[col] = str(source_row[col])

    row_to_active["Case_ID"] = new_rescheduled_id

    if assigned_worker is not None:
        raw_w = str(assigned_worker).strip()
        clean_worker = "" if raw_w.lower() in ["", "not assigned"] else raw_w
        login_c_act = find_login_user_column(df_active)
        row_to_active[login_c_act] = clean_worker
        for col in df_active.columns:
            c_low = col.strip().lower().replace("_", " ")
            if c_low in ["case worker", "cases worker", "worker", "assigned to", "user"]:
                row_to_active[col] = clean_worker

    row_to_active["Booking_Status"] = "Rescheduled"
    row_to_active["Preferred_Location"] = pref_loc.strip()
    row_to_active["Preferred_Dates"] = pref_dates.strip()
    row_to_active["Prior_Notice"] = prior_notice.strip()
    row_to_active["Updated_By"] = updated_by
    row_to_active["Updated_At"] = updated_at
    row_to_active["Action_Type"] = "Rescheduled Case"

    df_active_ordered = pd.concat(
        [df_active, pd.DataFrame([row_to_active])], ignore_index=True
    )
    save_data(df_active_ordered, sheet_name=active_sheet_name)
    return new_rescheduled_id, active_sheet_name


@st.cache_data(ttl=900, show_spinner=False)
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

    target_sheet_key = str(sheet_name)
    if isinstance(sheet_name, int) and sheets_data:
        keys = list(sheets_data.keys())
        if 0 <= sheet_name < len(keys):
            target_sheet_key = str(keys[sheet_name])

    df_to_save = reorder_dataframe_columns(df_to_save)
    sheets_data[target_sheet_key] = df_to_save

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        for s, df_val in sheets_data.items():
            df_val = reorder_dataframe_columns(df_val)
            df_val.to_excel(writer, sheet_name=str(s), index=False)

    _push_sheet_to_webapp(target_sheet_key, df_to_save, sheets_data)
    st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
    st.cache_data.clear()
    st.session_state["_last_id_sync_mtime"] = _get_file_mtime()


@st.cache_data(ttl=86400, show_spinner=False)
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
        if df_s is None or df_s.empty or str(s_name).strip() in SYSTEM_SHEETS:
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
            res_s["__orig_idx__"] = res_s.index
            res_s["Sheet_Name"] = s_name
            res_s = reorder_dataframe_columns(res_s)
            matched_dfs.append(res_s)
    if not matched_dfs:
        return pd.DataFrame()
    return reorder_dataframe_columns(pd.concat(matched_dfs, ignore_index=True))


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


def get_all_workers_list(sheets_cache=None, perms_cache=None):
    all_sheets = sheets_cache if isinstance(sheets_cache, dict) else load_all_sheets_cached(_get_file_mtime())
    all_workers = set()
    for s_n, df_s in all_sheets.items():
        if df_s is not None and not df_s.empty and str(s_n).strip() not in SYSTEM_SHEETS:
            lc = find_login_user_column(df_s)
            users = df_s[lc].astype(str).str.strip()
            for u in users.unique():
                if u != "" and u.lower() not in ["", "admin", "not assigned", "nan", "none"]:
                    all_workers.add(u)
    perms_map = perms_cache if isinstance(perms_cache, dict) else load_permissions(all_sheets)
    for p_user in perms_map.keys():
        u_clean = str(p_user).strip()
        if u_clean and u_clean.lower() not in ["admin", "not assigned"]:
            all_workers.add(u_clean)
    for extra_w in ["Kalyan", "Anil"]:
        if not any(w.lower() == extra_w.lower() for w in all_workers):
            all_workers.add(extra_w)
    return sorted(list(all_workers))


def are_all_workers_blocked(sheets_cache=None, perms_cache=None):
    workers = get_all_workers_list(sheets_cache, perms_cache)
    if not workers:
        return False
    perms = perms_cache if isinstance(perms_cache, dict) else load_permissions(sheets_cache)
    norm_perms = {k.lower(): v for k, v in perms.items()}
    return all(
        norm_perms.get(w.lower(), {}).get("block_login", False) for w in workers
    )


def is_worker_blocked(worker_name, sheets_cache=None, perms_cache=None):
    if not worker_name:
        return False
    perms = perms_cache if isinstance(perms_cache, dict) else load_permissions(sheets_cache)
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


def enable_right_cursor_navigation_in_popup():
    components.html(
        """
        <script>
        (function() {
            const doc = window.parent.document;

            function getPopupInputs() {
                const dialog = doc.querySelector('[data-testid="stDialog"]');
                if (!dialog) return [];
                const elements = Array.from(
                    dialog.querySelectorAll('input:not([type="file"]):not([type="hidden"]):not([disabled]), select:not([disabled]), textarea:not([disabled])')
                );
                return elements.filter(el => {
                    const rect = el.getBoundingClientRect();
                    return rect.width > 0 && rect.height > 0 && !el.closest('[data-testid="stSidebar"]');
                });
            }

            function moveCursorRight(currentEl) {
                const inputs = getPopupInputs();
                if (!inputs.length) return;
                const idx = inputs.indexOf(currentEl);
                if (idx === -1) return;

                const curRect = currentEl.getBoundingClientRect();
                const curMidY = curRect.top + curRect.height / 2;

                let rightCandidate = null;
                let bestRightScore = Infinity;

                for (let i = 0; i < inputs.length; i++) {
                    if (i === idx) continue;
                    const cand = inputs[i];
                    const r = cand.getBoundingClientRect();
                    const midY = r.top + r.height / 2;
                    const dy = Math.abs(midY - curMidY);
                    const dx = r.left - curRect.right;

                    if (dx > -10 && dy < 36) {
                        const score = Math.max(0, dx) + dy * 4;
                        if (score < bestRightScore) {
                            bestRightScore = score;
                            rightCandidate = cand;
                        }
                    }
                }

                const nextEl = rightCandidate || inputs[(idx + 1) % inputs.length];
                if (nextEl) {
                    nextEl.focus();
                    if (typeof nextEl.select === "function" && nextEl.tagName === "INPUT") {
                        try { nextEl.select(); } catch (e) {}
                    }
                }
            }

            if (!doc.__hemanPopupRightNavBound) {
                doc.addEventListener("keydown", function(e) {
                    if (e.key !== "Enter" || e.shiftKey || e.ctrlKey || e.altKey || e.metaKey) return;
                    const dialog = doc.querySelector('[data-testid="stDialog"]');
                    if (!dialog) return;
                    const active = doc.activeElement;
                    if (!active || !dialog.contains(active)) return;
                    if (active.tagName === "TEXTAREA" || active.tagName === "BUTTON") return;

                    e.preventDefault();
                    e.stopPropagation();
                    moveCursorRight(active);
                }, true);
                doc.__hemanPopupRightNavBound = true;
            }
        })();
        </script>
        """,
        height=0,
        width=0,
    )


# Single fast load per render cycle
all_sheets_root = load_all_sheets_cached(_get_file_mtime())
perms_root = load_permissions(all_sheets_root)

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

if "is_editing_popup" not in st.session_state:
    st.session_state["is_editing_popup"] = False

if "is_delete_request_mode" not in st.session_state:
    st.session_state["is_delete_request_mode"] = False

if "is_reschedule_mode" not in st.session_state:
    st.session_state["is_reschedule_mode"] = False

if not st.session_state["logged_in"] and qp.get("logged_in") == "true":
    qp_role = qp.get("role", "worker")
    qp_user = qp.get("username", "")
    if qp_role == "worker" and is_worker_blocked(qp_user, all_sheets_root, perms_root):
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
    if is_worker_blocked(curr_w, all_sheets_root, perms_root):
        record_worker_logout(curr_w)
        st.query_params.clear()
        st.session_state["logged_in"] = False
        st.session_state["username"] = ""
        st.session_state["role"] = ""
        st.session_state["attendance_recorded"] = False
        close_case_popup_state()
        st.session_state["force_logged_out_msg"] = True
        st.rerun()

if "open_dialog" not in st.session_state:
    st.session_state["open_dialog"] = None

if "selected_row_idx" not in st.session_state:
    st.session_state["selected_row_idx"] = None

sheets_init = get_sheet_names(all_sheets_root)
default_s_init = get_default_active_sheet(sheets_init)

if st.session_state["logged_in"] and "edit_row" in qp:
    try:
        r_val = int(qp["edit_row"])
        if "active_sheet" in qp and qp["active_sheet"] in sheets_init:
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
    [data-testid="stDialog"] {
        z-index: 100005 !important;
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
    div[data-testid="stElementContainer"]:
    div[data-testid="stElementContainer"]:has(.visa-color-2) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-2) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #92400e 0%, #d97706 100%) !important;
        border-bottom: 3px solid #fbbf24 !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-3) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-3) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #3730a3 0%, #4f46e5 100%) !important;
        border-bottom: 3px solid #818cf8 !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-4) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-4) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #115e59 0%, #0d9488 100%) !important;
        border-bottom: 3px solid #2dd4bf !important;
        color: #ffffff !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-5) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-5) + div[data-testid="stElementContainer"] button {
        background: linear-gradient(135deg, #881337 0%, #e11d48 100%) !important;
        border-bottom: 3px solid #fb7185 !important;
        color: #ffffff !important;
    }

    div[data-testid="stElementContainer"]:has(.visa-color-0) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.visa-color-1) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.visa-color-2) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.visa-color-3) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.visa-color-4) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.visa-color-5) + div[data-testid="stElementContainer"] button {
        width: 100% !important;
        height: 44px !important;
        padding: 4px 10px !important;
        font-size: 13px !important;
        font-weight: 800 !important;
        border-radius: 10px !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        box-shadow: 0 3px 8px rgba(15, 23, 42, 0.14) !important;
        transition: transform 0.12s ease, filter 0.12s ease !important;
    }
    div[data-testid="stElementContainer"]:has(.visa-color-0) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.visa-color-1) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.visa-color-2) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.visa-color-3) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.visa-color-4) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.visa-color-5) + div[data-testid="stElementContainer"] button:hover {
        transform: translateY(-1px) !important;
        filter: brightness(1.08) !important;
    }
    div[data-testid="stElementContainer"]:has(.admin-btn-0) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-1) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-2) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-3) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-4) + div[data-testid="stElementContainer"] button,
    div[data-testid="stElementContainer"]:has(.admin-btn-5) + div[data-testid="stElementContainer"] button {
        width: 100% !important;
        height: 68px !important;
        padding: 6px 8px !important;
        font-size: 12.5px !important;
        font-weight: 800 !important;
        border-radius: 12px !important;
        line-height: 1.25 !important;
        white-space: normal !important;
        box-shadow: 0 5px 14px rgba(15, 23, 42, 0.18) !important;
        transition: transform 0.12s ease, filter 0.12s ease !important;
    }
    div[data-testid="stElementContainer"]:has(.admin-btn-0) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.admin-btn-1) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.admin-btn-2) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.admin-btn-3) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.admin-btn-4) + div[data-testid="stElementContainer"] button:hover,
    div[data-testid="stElementContainer"]:has(.admin-btn-5) + div[data-testid="stElementContainer"] button:hover {
        transform: translateY(-2px) !important;
        filter: brightness(1.08) !important;
    }

    .edit-box-banner {
        background: linear-gradient(135deg, #0f1c3f 0%, #172554 100%);
        color: white;
        padding: 16px 22px;
        border-radius: 12px;
        margin-bottom: 16px;
        border-left: 5px solid #d4af37;
    }
    .edit-box-banner h4 {
        margin: 0 0 4px 0;
        font-size: 19px;
        font-weight: 800;
    }
    .edit-box-banner p {
        margin: 0;
        font-size: 13px;
        opacity: 0.88;
    }
    .footer {
        text-align: center;
        color: #475569;
        font-size: 13px;
        padding: 20px;
        border-top: 1px solid #e2e8f0;
        margin-top: 32px;
    }
    div.stButton > button:focus {
        outline: none !important;
    }
</style>
""", unsafe_allow_html=True)

logo_b64 = ""
for l_file in LOGO_FILES:
    if os.path.exists(l_file):
        logo_b64 = get_base64_image(l_file)
        break

logo_img_tag = (
    f'<img src="data:image/jpeg;base64,{logo_b64}" alt="HemanVisa Logo">'
    if logo_b64
    else '<div style="background:#d4af37;color:#070f24;padding:8px 14px;border-radius:8px;font-weight:800;">HV</div>'
)

st.markdown(f"""
<div class="sticky-top-banner">
    <div class="logo-container">
        {logo_img_tag}
    </div>
    <div class="banner-middle">
        <div class="world-map-title">HEMANVISA GLOBAL SERVICES</div>
        <div class="world-map-subtitle">GLOBAL IMMIGRATION &amp; VISA CONSULTANCY</div>
    </div>
    <div class="banner-info">
        <div class="banner-badge">📧 <a href="mailto:info@hemanvisa.com">info@hemanvisa.com</a></div>
        <div class="banner-badge">🌐 <a href="https://hemanvisa.com" target="_blank">www.hemanvisa.com</a></div>
        <div class="social-icons">
            <a href="https://www.instagram.com" target="_blank" title="Instagram"><img src="https://cdn-icons-png.flaticon.com/512/174/174855.png" alt="Instagram"></a>
            <a href="https://www.facebook.com" target="_blank" title="Facebook"><img src="https://cdn-icons-png.flaticon.com/512/733/733547.png" alt="Facebook"></a>
            <a href="https://www.linkedin.com" target="_blank" title="LinkedIn"><img src="https://cdn-icons-png.flaticon.com/512/174/174857.png" alt="LinkedIn"></a>
            <a href="https://www.youtube.com" target="_blank" title="YouTube"><img src="https://cdn-icons-png.flaticon.com/512/1384/1384060.png" alt="YouTube"></a>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

try:
    from zoneinfo import ZoneInfo
    _IST_TZ = ZoneInfo("Asia/Kolkata")
except Exception:
    _IST_TZ = None


def get_ist_now():
    if _IST_TZ is not None:
        return datetime.datetime.now(_IST_TZ).replace(tzinfo=None)
    return datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)


def _parse_any_dt_to_ist(val_str):
    s = str(val_str).strip()
    if not s or s == "-":
        return None
    for fmt in (
        "%Y-%m-%d %I:%M:%S %p",
        "%Y-%m-%d %I:%M %p",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    ):
        try:
            return datetime.datetime.strptime(s, fmt)
        except Exception:
            continue
    iso_match = re.match(
        r"^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2}:\d{2})(?:\.\d+)?(?:Z|[+-]\d{2}:?\d{2})?$",
        s,
        re.IGNORECASE,
    )
    if iso_match:
        d_part, t_part = iso_match.groups()
        try:
            dt_utc = datetime.datetime.strptime(f"{d_part} {t_part}", "%Y-%m-%d %H:%M:%S")
            return dt_utc + datetime.timedelta(hours=5, minutes=30)
        except Exception:
            pass
    return None


def _compute_duration_seconds(date_str, login_str, logout_str, is_active=False):
    t_in = _parse_time_12_or_24(login_str)
    if not t_in:
        return 0
    try:
        d_clean = _normalize_date_ist(date_str)
        d_base = (
            datetime.datetime.strptime(d_clean, "%Y-%m-%d").date()
            if d_clean
            else get_ist_now().date()
        )
    except Exception:
        d_base = get_ist_now().date()

    dt_in = datetime.datetime.combine(d_base, t_in)

    t_out = _parse_time_12_or_24(logout_str)
    if t_out:
        dt_out = datetime.datetime.combine(d_base, t_out)
        if dt_out < dt_in:
            dt_out = dt_out + datetime.timedelta(days=1)
        return max(0, int((dt_out - dt_in).total_seconds()))

    if is_active:
        now_ist = get_ist_now()
        dt_now = datetime.datetime.combine(d_base, now_ist.time())
        if d_base < now_ist.date():
            dt_now = now_ist
        elif dt_now < dt_in:
            dt_now = dt_now + datetime.timedelta(days=1)
        return max(0, int((dt_now - dt_in).total_seconds()))

    return 0


def _format_seconds_hhmm(total_seconds):
    sec = max(0, int(total_seconds))
    hrs = sec // 3600
    mins = (sec % 3600) // 60
    return f"{hrs}h {mins}m"


@st.cache_data(ttl=60, show_spinner=False)
def _load_attendance_cached(file_mtime):
    rows = _read_sheet_rows("App_Attendance")
    cleaned = []
    sno = 1
    for r in rows:
        w = str(r.get("Worker", "")).strip()
        if w:
            d_clean = _normalize_date_ist(r.get("Date", ""))
            l_clean = _normalize_time_ist(r.get("Login_Time", "-"))
            lo_clean = _normalize_time_ist(r.get("Logout_Time", "-"))
            raw_st = str(r.get("Status", "")).strip()
            st_clean = "Logged Out" if lo_clean != "-" else (raw_st if raw_st else "Active")
            is_act = st_clean.lower() in ["active", "logged in", "online"]
            dur_sec = _compute_duration_seconds(d_clean, l_clean, lo_clean, is_active=is_act)
            dur_str = _format_seconds_hhmm(dur_sec) if (l_clean != "-" and (lo_clean != "-" or is_act)) else "0h 0m"
            cleaned.append({
                "S.No": sno,
                "Worker": w,
                "Date": d_clean,
                "Login_Time": l_clean,
                "Logout_Time": lo_clean,
                "Duration": dur_str,
                "Status": st_clean,
                "_duration_sec": dur_sec,
            })
            sno += 1
    return cleaned


def load_attendance(sheets_cache=None):
    return _load_attendance_cached(_get_file_mtime())


def save_attendance(att_list):
    records = []
    sno = 1
    for item in (att_list or []):
        w = str(item.get("Worker", "")).strip()
        if w:
            d_clean = _normalize_date_ist(item.get("Date", ""))
            l_clean = _normalize_time_ist(item.get("Login_Time", "-"))
            lo_clean = _normalize_time_ist(item.get("Logout_Time", "-"))
            raw_st = str(item.get("Status", "")).strip()
            st_clean = "Logged Out" if lo_clean != "-" else (raw_st if raw_st else "Active")
            is_act = st_clean.lower() in ["active", "logged in", "online"]
            dur_sec = _compute_duration_seconds(d_clean, l_clean, lo_clean, is_active=is_act)
            dur_str = _format_seconds_hhmm(dur_sec) if (l_clean != "-" and (lo_clean != "-" or is_act)) else "0h 0m"
            records.append({
                "S.No": sno,
                "Worker": w,
                "Date": d_clean,
                "Login_Time": l_clean,
                "Logout_Time": lo_clean,
                "Duration": dur_str,
                "Status": st_clean,
            })
            sno += 1
    _write_sheet_records("App_Attendance", ATTENDANCE_COLS, records)


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
    _save_login_tracking(logins)

    attendance = load_attendance()
    attendance.append({
        "Worker": worker_name.strip(),
        "Date": today_str,
        "Login_Time": time_str,
        "Logout_Time": "-",
        "Duration": "0h 0m",
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
        if k.strip().lower() == worker_name.strip().lower():
            logins[k]["status"] = "Not Logged In"
    _save_login_tracking(logins)

    attendance = load_attendance()
    updated_any = False

    for entry in reversed(attendance):
        entry_date = _normalize_date_ist(entry.get("Date", ""))
        if (
            entry.get("Worker", "").strip().lower() == worker_name.strip().lower()
            and entry_date == today_str
            and _normalize_time_ist(entry.get("Logout_Time", "-")) == "-"
        ):
            entry["Date"] = today_str
            entry["Logout_Time"] = time_str
            entry["Status"] = "Logged Out"
            dur_sec = _compute_duration_seconds(today_str, entry.get("Login_Time", "-"), time_str, is_active=False)
            entry["Duration"] = _format_seconds_hhmm(dur_sec)
            updated_any = True
            break

    if not updated_any:
        for entry in reversed(attendance):
            if (
                entry.get("Worker", "").strip().lower() == worker_name.strip().lower()
                and _normalize_time_ist(entry.get("Logout_Time", "-")) == "-"
            ):
                e_date = _normalize_date_ist(entry.get("Date", "")) or today_str
                entry["Logout_Time"] = time_str
                entry["Status"] = "Logged Out"
                dur_sec = _compute_duration_seconds(e_date, entry.get("Login_Time", "-"), time_str, is_active=False)
                entry["Duration"] = _format_seconds_hhmm(dur_sec)
                updated_any = True
                break

    if not updated_any:
        attendance.append({
            "Worker": worker_name.strip(),
            "Date": today_str,
            "Login_Time": "-",
            "Logout_Time": time_str,
            "Duration": "0h 0m",
            "Status": "Logged Out",
        })
    save_attendance(attendance)


def build_attendance_summary_tables(attendance_data):
    if not attendance_data:
        return pd.DataFrame(), pd.DataFrame()

    worker_date_map = {}
    worker_overall_map = {}

    for row in attendance_data:
        w = str(row.get("Worker", "")).strip() or "Unknown"
        d = str(row.get("Date", "")).strip() or "Unknown"
        l_t = str(row.get("Login_Time", "-")).strip() or "-"
        lo_t = str(row.get("Logout_Time", "-")).strip() or "-"
        st = str(row.get("Status", "")).strip() or ("Logged Out" if lo_t != "-" else "Active")
        sec = int(row.get("_duration_sec", 0) or 0)

        t_in = _parse_time_12_or_24(l_t)
        t_out = _parse_time_12_or_24(lo_t)

        wd_key = (w, d)
        if wd_key not in worker_date_map:
            worker_date_map[wd_key] = {
                "Worker": w,
                "Date": d,
                "Sessions": 0,
                "First_Login": l_t if l_t != "-" else None,
                "_first_login_t": t_in,
                "Last_Logout": lo_t if lo_t != "-" else None,
                "_last_logout_t": t_out,
                "Total_Seconds": 0,
                "Current_Status": st,
            }
        wd = worker_date_map[wd_key]
        wd["Sessions"] += 1
        if t_in is not None:
            if wd["_first_login_t"] is None or t_in < wd["_first_login_t"]:
                wd["_first_login_t"] = t_in
                wd["First_Login"] = l_t
        if t_out is not None:
            if wd["_last_logout_t"] is None or t_out >= wd["_last_logout_t"]:
                wd["_last_logout_t"] = t_out
                wd["Last_Logout"] = lo_t
        wd["Total_Seconds"] += sec
        wd["Current_Status"] = st

        if w not in worker_overall_map:
            worker_overall_map[w] = {
                "Worker": w,
                "Active_Days": set(),
                "Total_Sessions": 0,
                "Last_Login": l_t if l_t != "-" else "-",
                "Last_Logout": lo_t if lo_t != "-" else "-",
                "Total_Seconds": 0,
                "Current_Status": st,
            }
        wo = worker_overall_map[w]
        if d and d != "Unknown":
            wo["Active_Days"].add(d)
        wo["Total_Sessions"] += 1
        if l_t != "-":
            wo["Last_Login"] = l_t
        if lo_t != "-":
            wo["Last_Logout"] = lo_t
        wo["Total_Seconds"] += sec
        wo["Current_Status"] = st

    day_rows = []
    for idx_d, (_, wd) in enumerate(worker_date_map.items(), start=1):
        tot_sec = wd["Total_Seconds"]
        day_rows.append({
            "S.No": idx_d,
            "Worker": wd["Worker"],
            "Date": wd["Date"],
            "Sessions": wd["Sessions"],
            "First_Login": wd["First_Login"] or "-",
            "Last_Logout": wd["Last_Logout"] or "-",
            "Total_Hours": round(tot_sec / 3600.0, 2),
            "Total_Duration": _format_seconds_hhmm(tot_sec),
            "Status": wd["Current_Status"],
        })

    overall_rows = []
    grand_sec = 0
    grand_sessions = 0
    for idx_o, (_, wo) in enumerate(worker_overall_map.items(), start=1):
        tot_sec = wo["Total_Seconds"]
        grand_sec += tot_sec
        grand_sessions += wo["Total_Sessions"]
        overall_rows.append({
            "S.No": idx_o,
            "Worker": wo["Worker"],
            "Days_Worked": len(wo["Active_Days"]),
            "Total_Sessions": wo["Total_Sessions"],
            "Last_Login": wo["Last_Login"],
            "Last_Logout": wo["Last_Logout"],
            "Total_Hours": round(tot_sec / 3600.0, 2),
            "Total_Duration": _format_seconds_hhmm(tot_sec),
            "Status": wo["Current_Status"],
        })

    if overall_rows:
        overall_rows.append({
            "S.No": len(overall_rows) + 1,
            "Worker": "OVERALL TOTAL",
            "Days_Worked": "-",
            "Total_Sessions": grand_sessions,
            "Last_Login": "-",
            "Last_Logout": "-",
            "Total_Hours": round(grand_sec / 3600.0, 2),
            "Total_Duration": _format_seconds_hhmm(grand_sec),
            "Status": "-",
        })

    df_day = pd.DataFrame(day_rows)
    df_overall = pd.DataFrame(overall_rows)
    return df_day, df_overall


def _reset_sno_df(df):
    if df is None or df.empty:
        return df
    out = df.copy().reset_index(drop=True)
    out.insert(0, "S.No", range(1, len(out) + 1)) if "S.No" not in out.columns else out.__setitem__("S.No", range(1, len(out) + 1))
    return out


# ==========================================
# HOME PAGE (NOT LOGGED IN)
# ==========================================
if not st.session_state["logged_in"]:
    if st.session_state.get("force_logged_out_msg"):
        st.warning("🔒 అడ్మిన్ మీ లాగిన్ యాక్సెస్‌ను తాత్కాలికంగా నిలిపివేశారు. మీ సెషన్ లాగ్ అవుట్ చేయబడింది.")
        st.session_state["force_logged_out_msg"] = False

    st.markdown("""
    <div class="hero-section">
        <div class="hero-title">✈️ Welcome to HemanVisa Global Services</div>
        <div class="hero-subtitle">
            Seamless Visa Processing, Appointment Management &amp; Global Immigration Assistance.
            Sign in below to access your personalized case dashboard.
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
        <div class="corporate-card">
            <h3>🌍 Global Visa Expertise</h3>
            <p>End-to-end guidance for tourist, business, work, and student visas across major international destinations.</p>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
        <div class="corporate-card">
            <h3>📅 Smart Appointment Tracking</h3>
            <p>Real-time monitoring of visa slots, rescheduling workflows, and instant document status updates.</p>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown("""
        <div class="corporate-card">
            <h3>🛡️ Secure Case Operations</h3>
            <p>Role-based access control, automated case IDs, and synchronized reporting for team efficiency.</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    _, login_col, _ = st.columns([1, 1.4, 1])
    with login_col:
        st.markdown("### 🔐 Portal Sign In")
        login_mode = st.radio("Select Login Role", ["Case Worker", "Admin"], horizontal=True)

        if login_mode == "Admin":
            admin_u_input = st.text_input("Admin Username", key="admin_login_u")
            admin_p_input = st.text_input("Admin Password", type="password", key="admin_login_p")
            if st.button("Sign In as Admin", type="primary", use_container_width=True):
                expected_u, expected_p = get_admin_creds(all_sheets_root)
                if admin_u_input.strip() == expected_u and admin_p_input == expected_p:
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = expected_u
                    st.session_state["role"] = "admin"
                    st.query_params["logged_in"] = "true"
                    st.query_params["role"] = "admin"
                    st.query_params["username"] = expected_u
                    st.rerun()
                else:
                    st.error("❌ Invalid Admin Username or Password.")
        else:
            all_workers_list = get_all_workers_list(all_sheets_root, perms_root)
            worker_options = all_workers_list if all_workers_list else ["Keerthana", "Sai", "Kalyan", "Anil"]
            selected_worker = st.selectbox("Select Case Worker", worker_options, key="worker_login_select")
            if st.button("Sign In as Case Worker", type="primary", use_container_width=True):
                if is_worker_blocked(selected_worker, all_sheets_root, perms_root):
                    st.error("🔒 మీ ఖాతా (Login) ప్రస్తుతం అడ్మిన్ ద్వారా బ్లాక్ చేయబడింది.")
                else:
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = selected_worker
                    st.session_state["role"] = "worker"
                    st.query_params["logged_in"] = "true"
                    st.query_params["role"] = "worker"
                    st.query_params["username"] = selected_worker
                    record_worker_login(selected_worker)
                    st.session_state["attendance_recorded"] = True
                    st.rerun()

    st.markdown('<div class="footer">© HemanVisa Global Services. All Rights Reserved.</div>', unsafe_allow_html=True)
    st.stop()


# ==========================================
# LOGGED IN WORKSPACE (ADMIN & WORKER)
# ==========================================
current_user = st.session_state.get("username", "")
current_role = st.session_state.get("role", "worker")
is_admin = current_role == "admin"

sheets_list = get_sheet_names(all_sheets_root)
if st.session_state.get("active_sheet") not in sheets_list:
    st.session_state["active_sheet"] = get_default_active_sheet(sheets_list)

active_sheet_name = st.session_state["active_sheet"]
df_current = load_data(sheet_name=active_sheet_name)
if df_current is None:
    df_current = pd.DataFrame()

notice_text = get_admin_notice(all_sheets_root)
if notice_text:
    st.markdown(f"""
    <div class="marquee-container">
        <marquee behavior="scroll" direction="left" scrollamount="6">📢 {notice_text}</marquee>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
hdr_c1, hdr_c2, hdr_c3 = st.columns([4.5, 2.5, 1.5])
with hdr_c1:
    role_badge = "ADMIN PORTAL" if is_admin else f"CASE WORKER • {current_user}"
    st.markdown(f"""
    <div class="workspace-header-bar">
        <span>📌 {role_badge}</span>
        <span class="workspace-sheet-pill">Sheet: {active_sheet_name}</span>
    </div>
    """, unsafe_allow_html=True)
with hdr_c2:
    chosen_sheet = st.selectbox(
        "Active Sheet",
        sheets_list,
        index=sheets_list.index(active_sheet_name) if active_sheet_name in sheets_list else 0,
        key="topbar_sheet_selector",
        label_visibility="collapsed",
    )
    if chosen_sheet != active_sheet_name:
        st.session_state["active_sheet"] = chosen_sheet
        st.session_state["dash_sheet"] = chosen_sheet
        st.session_state["records_sheet"] = chosen_sheet
        st.query_params["active_sheet"] = chosen_sheet
        close_case_popup_state()
        st.rerun()
with hdr_c3:
    if st.button("🚪 Logout", use_container_width=True, key="topbar_logout_btn"):
        if not is_admin and current_user:
            record_worker_logout(current_user)
        close_case_popup_state()
        st.query_params.clear()
        st.session_state["logged_in"] = False
        st.session_state["username"] = ""
        st.session_state["role"] = ""
        st.session_state["attendance_recorded"] = False
        st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

enable_right_cursor_navigation_in_popup()

worker_perms = perms_root.get(current_user, {}) if not is_admin else {}
can_edit_w = True if is_admin else worker_perms.get("can_edit", worker_perms.get("can_update", True))
can_del_req_w = True if is_admin else worker_perms.get("can_delete_req", True)
can_resched_w = True if is_admin else worker_perms.get("can_reschedule", True)
can_add_w = True if is_admin else worker_perms.get("can_add", True)
can_search_w = True if is_admin else worker_perms.get("can_search", True)
can_download_w = True if is_admin else worker_perms.get("can_download", False)
can_sync_w = True if is_admin else worker_perms.get("can_sync", False)


@st.dialog("📋 Case Details & Actions", width="large")
def render_case_popup_dialog():
    r_idx = st.session_state.get("selected_row_idx")
    df_pop = load_data(sheet_name=st.session_state.get("active_sheet", active_sheet_name))
    if df_pop is None or df_pop.empty or r_idx is None or r_idx not in df_pop.index:
        st.warning("Case record not found.")
        if st.button("Close", key="pop_not_found_close"):
            close_case_popup_state()
            st.rerun()
        return

    row_s = df_pop.loc[r_idx]
    st.markdown(f"""
    <div class="edit-box-banner">
        <h4>Case ID: {row_s.get("Case_ID", "-")}</h4>
        <p>Sheet: {st.session_state.get("active_sheet", active_sheet_name)} | Status: {row_s.get("Booking_Status", "Active")}</p>
    </div>
    """, unsafe_allow_html=True)

    info_cols = st.columns(2)
    for i, col_name in enumerate(df_pop.columns):
        with info_cols[i % 2]:
            st.text_input(col_name, value=str(row_s.get(col_name, "")), disabled=True, key=f"view_field_{r_idx}_{col_name}")

    act_c1, act_c2, act_c3, act_c4 = st.columns(4)
    with act_c1:
        if can_edit_w and st.button("✏️ Edit Case", key=f"btn_edit_mode_{r_idx}", use_container_width=True):
            st.session_state["is_editing_popup"] = True
            st.session_state["is_delete_request_mode"] = False
            st.session_state["is_reschedule_mode"] = False
    with act_c2:
        if can_resched_w and st.button("🔄 Reschedule", key=f"btn_resched_mode_{r_idx}", use_container_width=True):
            st.session_state["is_reschedule_mode"] = True
            st.session_state["is_editing_popup"] = False
            st.session_state["is_delete_request_mode"] = False
    with act_c3:
        if can_del_req_w and st.button("🗑️ Delete Req", key=f"btn_del_mode_{r_idx}", use_container_width=True):
            st.session_state["is_delete_request_mode"] = True
            st.session_state["is_editing_popup"] = False
            st.session_state["is_reschedule_mode"] = False
    with act_c4:
        if st.button("❌ Close", key=f"btn_close_pop_{r_idx}", use_container_width=True):
            close_case_popup_state()
            st.rerun()

    if st.session_state.get("is_editing_popup") and can_edit_w:
        st.markdown("#### ✏️ Edit Fields")
        updated_vals = {}
        editable_cols = [c for c in df_pop.columns if c not in ["Case_ID", "Updated_By", "Updated_At"]]
        for idx_e, col_e in enumerate(editable_cols):
            updated_vals[col_e] = st.text_input(f"Edit {col_e}", value=str(row_s.get(col_e, "")), key=f"edit_inp_{r_idx}_{idx_e}")
        if st.button("💾 Save Changes", type="primary", key=f"save_edit_btn_{r_idx}"):
            now_ist = get_ist_now()
            for col_e, val_e in updated_vals.items():
                df_pop.at[r_idx, col_e] = val_e
            df_pop.at[r_idx, "Updated_By"] = current_user
            df_pop.at[r_idx, "Updated_At"] = f"{now_ist.strftime('%Y-%m-%d')} {now_ist.strftime('%I:%M:%S %p')}"
            df_pop.at[r_idx, "Action_Type"] = "Updated"
            save_data(df_pop, sheet_name=st.session_state.get("active_sheet", active_sheet_name))
            close_case_popup_state()
            st.success("✅ Saved successfully!")
            st.rerun()

    if st.session_state.get("is_reschedule_mode") and can_resched_w:
        st.markdown("#### 🔄 Reschedule Details")
        pref_loc = st.text_input("Preferred Location", value=str(row_s.get("Preferred_Location", "")), key=f"res_loc_{r_idx}")
        pref_dates = st.text_input("Preferred Dates", value=str(row_s.get("Preferred_Dates", "")), key=f"res_dates_{r_idx}")
        prior_notice = st.text_input("Prior Notice", value=str(row_s.get("Prior_Notice", "")), key=f"res_notice_{r_idx}")
        if st.button("✅ Confirm Reschedule", type="primary", key=f"confirm_res_{r_idx}"):
            now_ist = get_ist_now()
            ts_str = f"{now_ist.strftime('%Y-%m-%d')} {now_ist.strftime('%I:%M:%S %p')}"
            df_pop.at[r_idx, "Booking_Status"] = "Rescheduled"
            df_pop.at[r_idx, "Preferred_Location"] = pref_loc
            df_pop.at[r_idx, "Preferred_Dates"] = pref_dates
            df_pop.at[r_idx, "Prior_Notice"] = prior_notice
            df_pop.at[r_idx, "Updated_By"] = current_user
            df_pop.at[r_idx, "Updated_At"] = ts_str
            df_pop.at[r_idx, "Action_Type"] = "Rescheduled"
            save_data(df_pop, sheet_name=st.session_state.get("active_sheet", active_sheet_name))
            append_rescheduled_to_active_sheet(row_s, pref_loc, pref_dates, prior_notice, current_user, ts_str)
            close_case_popup_state()
            st.success("✅ Rescheduled!")
            st.rerun()

    if st.session_state.get("is_delete_request_mode") and can_del_req_w:
        st.markdown("#### 🗑️ Delete Request")
        del_reason = st.text_input("Reason for Deletion", value=str(row_s.get("Delete_Reason", "")), key=f"del_reason_{r_idx}")
        if st.button("📨 Submit Delete Request", type="primary", key=f"sub_del_{r_idx}"):
            if not del_reason.strip():
                st.warning("Please enter a deletion reason.")
            else:
                now_ist = get_ist_now()
                df_pop.at[r_idx, "Delete_Reason"] = del_reason.strip()
                df_pop.at[r_idx, "Action_Type"] = "Delete Requested"
                df_pop.at[r_idx, "Updated_By"] = current_user
                df_pop.at[r_idx, "Updated_At"] = f"{now_ist.strftime('%Y-%m-%d')} {now_ist.strftime('%I:%M:%S %p')}"
                save_data(df_pop, sheet_name=st.session_state.get("active_sheet", active_sheet_name))
                close_case_popup_state()
                st.success("✅ Delete request submitted!")
                st.rerun()


if st.session_state.get("open_dialog") == "edit_case_popup":
    render_case_popup_dialog()


# ==========================================
# TABS SECTION
# ==========================================
if is_admin:
    tabs = st.tabs(["📊 Dashboard", "📁 All Cases", "➕ Add Case", "🕒 Attendance & Logins", "⚙️ Admin Settings"])
else:
    tabs = st.tabs(["📊 Dashboard", "📁 My Cases", "➕ Add Case", "🕒 My Attendance"])

# --- TAB 0: DASHBOARD ---
with tabs[0]:
    if not df_current.empty:
        login_col_d = find_login_user_column(df_current)
        if is_admin:
            df_dash = df_current.copy()
        else:
            df_dash = df_current[
                df_current[login_col_d].astype(str).str.strip().str.lower() == current_user.lower()
            ].copy()

        total_cases_cnt = len(df_dash)
        booked_cnt = int(
            df_dash["Booking_Status"].astype(str).str.lower().eq("booked").sum()
        ) if "Booking_Status" in df_dash.columns else 0
        resched_cnt = int(
            df_dash["Booking_Status"].astype(str).str.lower().eq("rescheduled").sum()
        ) if "Booking_Status" in df_dash.columns else 0
        del_req_cnt = int(
            df_dash["Action_Type"].astype(str).str.lower().str.contains("delete", na=False).sum()
        ) if "Action_Type" in df_dash.columns else 0

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Cases</div><div class="kpi-value">{total_cases_cnt}</div></div>', unsafe_allow_html=True)
        with k2:
            st.markdown(f'<div class="kpi-card green"><div class="kpi-title">Booked</div><div class="kpi-value">{booked_cnt}</div></div>', unsafe_allow_html=True)
        with k3:
            st.markdown(f'<div class="kpi-card gold"><div class="kpi-title">Rescheduled</div><div class="kpi-value">{resched_cnt}</div></div>', unsafe_allow_html=True)
        with k4:
            st.markdown(f'<div class="kpi-card rose"><div class="kpi-title">Delete Requests</div><div class="kpi-value">{del_req_cnt}</div></div>', unsafe_allow_html=True)

        visa_col_d = find_visa_type_column(df_dash)
        if visa_col_d:
            st.markdown("#### Visa Categories")
            visa_vals = sorted([v for v in df_dash[visa_col_d].astype(str).str.strip().unique() if v])
            v_cols = st.columns(min(6, max(1, len(visa_vals) + 1)))
            with v_cols[0]:
                st.markdown('<div class="visa-color-0"></div>', unsafe_allow_html=True)
                if st.button("All Visas", key="visa_btn_all"):
                    st.session_state["selected_visa_filter"] = "All"
            for idx_v, v_name in enumerate(visa_vals[:5], start=1):
                with v_cols[idx_v]:
                    st.markdown(f'<div class="visa-color-{idx_v % 6}"></div>', unsafe_allow_html=True)
                    if st.button(v_name, key=f"visa_btn_{idx_v}"):
                        st.session_state["selected_visa_filter"] = v_name

            sel_visa = st.session_state.get("selected_visa_filter", "All")
            if sel_visa != "All":
                df_dash = df_dash[df_dash[visa_col_d].astype(str).str.strip() == sel_visa]

        st.dataframe(_reset_sno_df(df_dash), use_container_width=True, hide_index=True)
    else:
        st.info("No records available in this sheet.")

# --- TAB 1: CASES LIST & POPUP TRIGGER ---
with tabs[1]:
    if not df_current.empty:
        login_col_c = find_login_user_column(df_current)
        if is_admin:
            df_view = df_current.copy()
        else:
            df_view = df_current[
                df_current[login_col_c].astype(str).str.strip().str.lower() == current_user.lower()
            ].copy()

        if can_search_w:
            q_txt = st.text_input("🔍 Search Case ID, Username, Password, DS-160, Worker...", key="cases_search_box")
            if q_txt.strip():
                mask_q = search_username_ds160_password_worker(df_view, q_txt)
                df_view = df_view[mask_q]

        if not df_view.empty:
            st.dataframe(_reset_sno_df(df_view), use_container_width=True, hide_index=True)

            st.markdown("#### Select Case to Open Popup")
            row_options = {
                f"{idx + 1}. {r.get('Case_ID', 'HM')} | {r.get(find_display_username_column(df_view), '-')}": orig_i
                for idx, (orig_i, r) in enumerate(df_view.iterrows())
            }
            chosen_case_label = st.selectbox("Choose Case", list(row_options.keys()), key="popup_case_selector")
            if st.button("📂 Open Case Popup", type="primary", key="open_case_popup_btn"):
                st.session_state["selected_row_idx"] = row_options[chosen_case_label]
                st.session_state["open_dialog"] = "edit_case_popup"
                st.session_state["is_editing_popup"] = False
                st.session_state["is_delete_request_mode"] = False
                st.session_state["is_reschedule_mode"] = False
                st.rerun()

        if can_download_w and not is_admin:
            dl_bytes = get_worker_export_bytes(_get_file_mtime(), current_user, active_sheet_name)
            st.download_button(
                "📥 Download My Cases (Excel)",
                data=dl_bytes,
                file_name=f"{current_user}_cases.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="worker_dl_btn",
            )
    else:
        st.info("No cases found.")

# --- TAB 2: ADD CASE ---
with tabs[2]:
    if not can_add_w:
        st.warning("🔒 You do not have permission to add new cases.")
    else:
        st.markdown("### ➕ Add New Case")
        new_case_id = generate_next_case_id()
        st.text_input("Generated Case ID", value=new_case_id, disabled=True, key="add_case_id_disp")

        base_cols = [c for c in df_current.columns if c not in ["Case_ID", "Updated_By", "Updated_At", "Action_Type"]]
        if not base_cols:
            base_cols = ["USERNAME", "PASSWORD", "Lead", "Booking_Status"]

        new_record = {"Case_ID": new_case_id}
        for i_c, col_n in enumerate(base_cols):
            default_val = current_user if (not is_admin and col_n == find_login_user_column(df_current)) else ""
            new_record[col_n] = st.text_input(col_n, value=default_val, key=f"new_case_inp_{i_c}")

        if st.button("💾 Save New Case", type="primary", key="save_new_case_btn"):
            now_ist = get_ist_now()
            new_record["Updated_By"] = current_user
            new_record["Updated_At"] = f"{now_ist.strftime('%Y-%m-%d')} {now_ist.strftime('%I:%M:%S %p')}"
            new_record["Action_Type"] = "Added"
            df_updated = pd.concat([df_current, pd.DataFrame([new_record])], ignore_index=True)
            save_data(df_updated, sheet_name=active_sheet_name)
            st.success(f"✅ Case {new_case_id} added successfully!")
            st.rerun()

# --- TAB 3: ATTENDANCE & LOGINS ---
with tabs[3]:
    att_all = load_attendance()
    if is_admin:
        att_filtered = att_all
    else:
        att_filtered = [r for r in att_all if str(r.get("Worker", "")).strip().lower() == current_user.lower()]

    df_day_sum, df_overall_sum = build_attendance_summary_tables(att_filtered)

    st.markdown("### 🕒 Session-wise Login / Logout & Duration Log")
    if att_filtered:
        df_sessions_disp = pd.DataFrame([
            {
                "S.No": idx_s,
                "Worker": r.get("Worker", ""),
                "Date": r.get("Date", ""),
                "Login_Time": r.get("Login_Time", "-"),
                "Logout_Time": r.get("Logout_Time", "-"),
                "Duration": r.get("Duration", "0h 0m"),
                "Status": r.get("Status", ""),
            }
            for idx_s, r in enumerate(att_filtered, start=1)
        ])
        st.dataframe(df_sessions_disp, use_container_width=True, hide_index=True)
    else:
        st.info("No login/logout sessions recorded yet.")

    st.markdown("### 📅 Daily Worker Duration Summary (All Sessions Combined)")
    if not df_day_sum.empty:
        st.dataframe(_reset_sno_df(df_day_sum), use_container_width=True, hide_index=True)
    else:
        st.info("No daily summary available.")

    if is_admin:
        st.markdown("### 📊 Overall Worker Summary")
        if not df_overall_sum.empty:
            st.dataframe(df_overall_sum, use_container_width=True, hide_index=True)

        st.markdown("### 🟢 Current Live Login Status")
        logins_map = load_login_tracking()
        if logins_map:
            live_rows = [
                {
                    "S.No": idx_l,
                    "Worker": w_k,
                    "Status": v_info.get("status", "Not Logged In"),
                    "Last_Login": v_info.get("last_login", "-"),
                }
                for idx_l, (w_k, v_info) in enumerate(logins_map.items(), start=1)
            ]
            st.dataframe(pd.DataFrame(live_rows), use_container_width=True, hide_index=True)

# --- TAB 4: ADMIN SETTINGS ---
if is_admin:
    with tabs[4]:
        st.markdown("### ⚙️ Admin Settings")

        st.markdown("#### 📢 Scrolling Announcement Notice")
        curr_notice = get_admin_notice(all_sheets_root)
        new_notice = st.text_input("Notice Text", value=curr_notice, key="admin_notice_input")
        if st.button("Save Notice", key="save_notice_btn"):
            save_admin_notice(new_notice)
            st.success("✅ Notice updated!")
            st.rerun()

        st.markdown("#### 🔐 Admin Credentials")
        curr_au, curr_ap = get_admin_creds(all_sheets_root)
        new_au = st.text_input("New Admin Username", value=curr_au, key="new_admin_u")
        new_ap = st.text_input("New Admin Password", value=curr_ap, type="password", key="new_admin_p")
        if st.button("Update Admin Credentials", key="update_admin_creds_btn"):
            save_admin_creds(new_au, new_ap)
            st.success("✅ Admin credentials updated!")
            st.rerun()

        st.markdown("#### 🛂 Worker Access & Login Block Control")
        all_w_list = get_all_workers_list(all_sheets_root, perms_root)
        b_c1, b_c2 = st.columns(2)
        with b_c1:
            if st.button("🔒 Block All Workers", key="block_all_w_btn", use_container_width=True):
                toggle_all_workers_block(True)
                st.success("✅ All workers blocked.")
                st.rerun()
        with b_c2:
            if st.button("🔓 Unblock All Workers", key="unblock_all_w_btn", use_container_width=True):
                toggle_all_workers_block(False)
                st.success("✅ All workers unblocked.")
                st.rerun()

        if all_w_list:
            sel_w_perm = st.selectbox("Select Worker to Configure", all_w_list, key="perm_worker_select")
            w_p = perms_root.get(sel_w_perm, {})
            p_edit = st.checkbox("Can Edit", value=w_p.get("can_edit", True), key="p_edit_chk")
            p_res = st.checkbox("Can Reschedule", value=w_p.get("can_reschedule", True), key="p_res_chk")
            p_del = st.checkbox("Can Delete Request", value=w_p.get("can_delete_req", True), key="p_del_chk")
            p_add = st.checkbox("Can Add Case", value=w_p.get("can_add", True), key="p_add_chk")
            p_dl = st.checkbox("Can Download Excel", value=w_p.get("can_download", False), key="p_dl_chk")
            p_blk = st.checkbox("Block Login", value=w_p.get("block_login", False), key="p_blk_chk")

            if st.button("💾 Save Worker Permissions", type="primary", key="save_worker_perms_btn"):
                updated_perms = dict(perms_root)
                updated_perms[sel_w_perm] = {
                    "can_edit": p_edit,
                    "can_update": p_edit,
                    "can_reschedule": p_res,
                    "can_delete_req": p_del,
                    "can_add": p_add,
                    "can_search": True,
                    "can_download": p_dl,
                    "can_sync": False,
                    "block_login": p_blk,
                }
                if p_blk:
                    record_worker_logout(sel_w_perm)
                save_permissions(updated_perms)
                st.success(f"✅ Permissions saved for {sel_w_perm}!")
                st.rerun()

        if can_sync_w and st.button("🔄 Sync Case IDs Across All Sheets", key="sync_all_ids_btn"):
            ensure_all_cases_have_unique_ids(force=True)
            st.success("✅ Case IDs synced!")
            st.rerun()

st.markdown('<div class="footer">© HemanVisa Global Services. All Rights Reserved.</div>', unsafe_allow_html=True)
