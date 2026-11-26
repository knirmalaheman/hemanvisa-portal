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


def get_webapp_url():
    try:
        url = st.secrets.get("GSHEET_WEBAPP_URL", "")
        return str(url).strip() if url else ""
    except Exception:
        return ""
def _rows_to_df(raw_rows):
    if not raw_rows or not isinstance(raw_rows, list):
        return pd.DataFrame()
    if len(raw_rows) == 0:
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
    cleaned_rows = []
    for r in data_rows:
        row_list = list(r) if isinstance(r, list) else [r]
        if len(row_list) < len(unique_headers):
            row_list = row_list + [""] * (len(unique_headers) - len(row_list))
        elif len(row_list) > len(unique_headers):
            row_list = row_list[: len(unique_headers)]
        cleaned_rows.append(["" if v is None else str(v) for v in row_list])
    return pd.DataFrame(cleaned_rows, columns=unique_headers, dtype=str)


def _fetch_sheets_from_webapp():
    url = get_webapp_url()
    if not url:
        return None
    try:
        resp = requests.get(url, timeout=45, allow_redirects=True)
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

        resp = requests.post(url, json=body, timeout=45, allow_redirects=True)
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


def _read_sheet_rows(sheet_name):
    sheets = load_all_sheets_cached(_get_file_mtime())
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


def get_admin_creds():
    rows = _read_sheet_rows("App_Config")
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


def get_admin_notice():
    rows = _read_sheet_rows("App_Config")
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


def load_permissions():
    rows = _read_sheet_rows("App_Permissions")
    perms = {}
    for r in rows:
        w = r.get("Worker", "").strip()
        if w:
            perms[w] = {
                "can_edit": str(r.get("can_edit", "True")).strip().lower()
                in ["true", "1", "yes"],
                "can_delete_req": str(r.get("can_delete_req", "True")).strip().lower()
                in ["true", "1", "yes"],
                "can_reschedule": str(r.get("can_reschedule", "True")).strip().lower()
                in ["true", "1", "yes"],
                "can_add": str(r.get("can_add", "True")).strip().lower()
                in ["true", "1", "yes"],
            }
    return perms


def save_permissions(perms_dict):
    cols = ["Worker", "can_edit", "can_delete_req", "can_reschedule", "can_add"]
    records = []
    for w, p in (perms_dict or {}).items():
        if str(w).strip():
            records.append({
                "Worker": str(w).strip(),
                "can_edit": str(p.get("can_edit", True)),
                "can_delete_req": str(p.get("can_delete_req", True)),
                "can_reschedule": str(p.get("can_reschedule", True)),
                "can_add": str(p.get("can_add", True)),
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


def load_login_tracking():
    rows = _read_sheet_rows("App_Logins")
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


ATTENDANCE_COLS = ["Worker", "Date", "Login_Time", "Logout_Time", "Duration", "Status"]


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


def load_attendance():
    rows = _read_sheet_rows("App_Attendance")
    cleaned = []
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
                "Worker": w,
                "Date": d_clean,
                "Login_Time": l_clean,
                "Logout_Time": lo_clean,
                "Duration": dur_str,
                "Status": st_clean,
                "_duration_sec": dur_sec,
            })
    return cleaned


def save_attendance(att_list):
    records = []
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
                "Worker": w,
                "Date": d_clean,
                "Login_Time": l_clean,
                "Logout_Time": lo_clean,
                "Duration": dur_str,
                "Status": st_clean,
            })
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
    for entry in attendance:
        if (
            entry.get("Worker", "").strip().lower() == worker_name.strip().lower()
            and _normalize_time_ist(entry.get("Logout_Time", "-")) == "-"
            and str(entry.get("Status", "")).strip().lower() in ["active", "logged in", ""]
        ):
            entry["Logout_Time"] = time_str
            entry["Status"] = "Logged Out"

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
            updated_any = True
            break

    if not updated_any:
        for entry in reversed(attendance):
            if (
                entry.get("Worker", "").strip().lower() == worker_name.strip().lower()
                and _normalize_time_ist(entry.get("Logout_Time", "-")) == "-"
            ):
                entry["Logout_Time"] = time_str
                entry["Status"] = "Logged Out"
                updated_any = True
                break

    if not updated_any:
        for entry in reversed(attendance):
            entry_date = _normalize_date_ist(entry.get("Date", ""))
            if (
                entry.get("Worker", "").strip().lower() == worker_name.strip().lower()
                and entry_date == today_str
            ):
                entry["Date"] = today_str
                entry["Logout_Time"] = time_str
                entry["Status"] = "Logged Out"
                updated_any = True
                break

    if not updated_any:
        for entry in reversed(attendance):
            if entry.get("Worker", "").strip().lower() == worker_name.strip().lower():
                entry["Logout_Time"] = time_str
                entry["Status"] = "Logged Out"
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
    for _, wd in worker_date_map.items():
        tot_sec = wd["Total_Seconds"]
        day_rows.append({
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
    for _, wo in worker_overall_map.items():
        tot_sec = wo["Total_Seconds"]
        grand_sec += tot_sec
        grand_sessions += wo["Total_Sessions"]
        overall_rows.append({
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
    if not df_day.empty:
        df_day.index = range(1, len(df_day) + 1)
    if not df_overall.empty:
        df_overall.index = range(1, len(df_overall) + 1)
    return df_day, df_overall
