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

try:
    from zoneinfo import ZoneInfo
    IST_TZ = ZoneInfo("Asia/Kolkata")
except Exception:
    IST_TZ = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

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
def _cached_webapp_url():
    try:
        url = st.secrets.get("GSHEET_WEBAPP_URL", "")
        return str(url).strip() if url else ""
    except Exception:
        return ""


def get_webapp_url():
    return _cached_webapp_url()


def _rows_to_df(raw_rows):
    if not raw_rows or not isinstance(raw_rows, list) or len(raw_rows) == 0:
        return pd.DataFrame(columns=["Case_ID", "Username", "USERNAME", "Password", "Lead"], dtype=str)
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
        resp = requests.get(url, timeout=15, allow_redirects=True)
        resp.raise_for_status()
        text_body = resp.text.strip()
        if not text_body:
            return None
        payload = json.loads(text_body)
        if isinstance(payload, dict) and isinstance(payload.get("sheets"), dict):
            sheets_dict = {}
            for s_name, s_rows in payload["sheets"].items():
                clean_s_name = str(s_name).strip()
                if clean_s_name:
                    sheets_dict[clean_s_name] = _rows_to_df(s_rows)
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

        resp = requests.post(url, json=body, timeout=20, allow_redirects=True)
        resp.raise_for_status()
        res_json = resp.json()
        if isinstance(res_json, dict) and res_json.get("status") != "success":
            st.error(f"Google Sheet Save Error ({sheet_name}): {res_json}")
    except Exception as e:
        st.error(f"Google Sheet Push Error ({sheet_name}): {e}")


def get_ist_now():
    try:
        now_aware = datetime.datetime.now(datetime.timezone.utc).astimezone(IST_TZ)
        return now_aware.replace(tzinfo=None)
    except Exception:
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
                "can_update": str(r.get("can_update", r.get("can_edit", "True"))).strip().lower()
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
        "Worker", "can_edit", "can_delete_req", "can_reschedule", "can_add",
        "can_update", "can_search", "can_download", "can_sync", "block_login",
    ]
    records = []
    for w, p in (perms_dict or {}).items():
        if str(w).strip():
            records.append({
                "Worker": str(w).strip(),
                "can_edit": str(p.get("can_edit", True)),
                "can_delete_req": str(p.get("can_delete_req", True)),
                "can_reschedule": str(p.get("can_reschedule", True)),
                "can_add": str(p.get("can_add", True)),
                "can_update": str(p.get("can_update", p.get("can_edit", True))),
                "can_search": str(p.get("can_search", True)),
                "can_download": str(p.get("can_download", False)),
                "can_sync": str(p.get("can_sync", False)),
                "block_login": str(p.get("block_login", False)),
            })
    _write_sheet_records("App_Permissions", cols, records)


def _parse_iso_to_ist_datetime(val_str):
    s = str(val_str).strip()
    if not s or s == "-":
        return None
    m = re.match(
        r"^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2}:\d{2})(?:\.\d+)?(Z|[+-]\d{2}:?\d{2})?$",
        s,
        re.IGNORECASE,
    )
    if not m:
        return None
    d_part, t_part, tz_part = m.groups()
    try:
        dt_naive = datetime.datetime.strptime(f"{d_part} {t_part}", "%Y-%m-%d %H:%M:%S")
        if tz_part:
            tz_norm = tz_part.upper()
            if tz_norm == "Z":
                dt_aware = dt_naive.replace(tzinfo=datetime.timezone.utc)
            else:
                sign = 1 if tz_norm.startswith("+") else -1
                digits = re.sub(r"[^0-9]", "", tz_norm[1:])
                off_h = int(digits[:2]) if len(digits) >= 2 else 0
                off_m = int(digits[2:4]) if len(digits) >= 4 else 0
                offset = datetime.timezone(datetime.timedelta(hours=sign * off_h, minutes=sign * off_m))
                dt_aware = dt_naive.replace(tzinfo=offset)
            return dt_aware.astimezone(IST_TZ).replace(tzinfo=None), True
        return dt_naive, False
    except Exception:
        return None


def _normalize_iso_timestamp_to_ist(val_str):
    s = str(val_str).strip()
    if not s or s == "-":
        return ""
    parsed = _parse_iso_to_ist_datetime(s)
    if parsed:
        dt_val, was_tz_aware = parsed
        d_part = dt_val.strftime("%Y-%m-%d")
        if d_part == "1899-12-30":
            return dt_val.strftime("%I:%M:%S %p")
        if was_tz_aware:
            return dt_val.strftime("%Y-%m-%d %I:%M:%S %p")
        return (dt_val + datetime.timedelta(hours=5, minutes=30)).strftime("%Y-%m-%d %I:%M:%S %p")
    return s


def _normalize_date_ist(val_str):
    s = str(val_str).strip()
    if not s or s == "-":
        return ""
    parsed = _parse_iso_to_ist_datetime(s)
    if parsed:
        dt_val, was_tz_aware = parsed
        if was_tz_aware:
            return dt_val.strftime("%Y-%m-%d")
        return (dt_val + datetime.timedelta(hours=5, minutes=30)).strftime("%Y-%m-%d")
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


def format_duration(minutes: float) -> str:
    """Convert minutes into 'Xh YYm' format."""
    if pd.isna(minutes) or minutes < 0:
        minutes = 0.0
    total_mins = int(round(minutes))
    hours = total_mins // 60
    mins = total_mins % 60
    return f"{hours}h {mins:02d}m"


def _format_seconds_hhmm(total_seconds):
    sec = max(0, int(total_seconds))
    hrs = sec // 3600
    mins = (sec % 3600) // 60
    return f"{hrs}h {mins:02d}m"


def build_attendance_reports(raw_logs_df: pd.DataFrame, now_dt: datetime.datetime | None = None):
    """
    Build detailed session logs, day-wise worker hours, and worker-wise/overall summary.
    raw_logs_df expected columns: Worker, Login_DT, Logout_DT (optional/None if active).
    """
    if now_dt is None:
        now_dt = get_ist_now()

    if raw_logs_df is None or raw_logs_df.empty:
        empty_sessions = pd.DataFrame(
            columns=[
                "Date",
                "Worker",
                "Login_Time",
                "Logout_Time",
                "Duration_Hours",
                "Duration",
                "Status",
            ]
        )
        empty_day = pd.DataFrame(
            columns=[
                "Date",
                "Worker",
                "Total_Sessions",
                "First_Login",
                "Last_Logout",
                "Total_Hours",
                "Total_Duration",
            ]
        )
        empty_summary = pd.DataFrame(
            columns=[
                "Worker",
                "Days_Worked",
                "Total_Sessions",
                "Last_Login",
                "Last_Logout",
                "Total_Hours",
                "Total_Duration",
                "Status",
            ]
        )
        return empty_sessions, empty_day, empty_summary, 0.0, "0h 00m", 0

    df = raw_logs_df.copy()
    df["Worker"] = df["Worker"].astype(str).str.strip()
    df["Login_DT"] = pd.to_datetime(df["Login_DT"], errors="coerce")
    df["Logout_DT"] = pd.to_datetime(df["Logout_DT"], errors="coerce")

    df = df.dropna(subset=["Worker", "Login_DT"]).copy()
    df["Is_Active"] = df["Logout_DT"].isna()
    df["Effective_Logout_DT"] = df["Logout_DT"].fillna(pd.Timestamp(now_dt))

    df["Duration_Mins"] = (
        (df["Effective_Logout_DT"] - df["Login_DT"]).dt.total_seconds() / 60.0
    ).clip(lower=0.0)

    df["Date"] = df["Login_DT"].dt.strftime("%Y-%m-%d")
    df["Login_Time"] = df["Login_DT"].dt.strftime("%I:%M:%S %p")
    df["Logout_Time"] = df.apply(
        lambda r: "-" if r["Is_Active"] else r["Logout_DT"].strftime("%I:%M:%S %p"),
        axis=1,
    )
    df["Status"] = df["Is_Active"].map({True: "Active", False: "Completed"})
    df["Duration_Hours"] = (df["Duration_Mins"] / 60.0).round(2)
    df["Duration"] = df["Duration_Mins"].apply(format_duration)

    df = df.sort_values(["Date", "Worker", "Login_DT"]).reset_index(drop=True)

    session_logs_df = df[
        [
            "Date",
            "Worker",
            "Login_Time",
            "Logout_Time",
            "Duration_Hours",
            "Duration",
            "Status",
        ]
    ].copy()

    day_grouped = (
        df.groupby(["Date", "Worker"], as_index=False)
        .agg(
            Total_Sessions=("Login_DT", "count"),
            First_Login_DT=("Login_DT", "min"),
            Last_Logout_DT=("Effective_Logout_DT", "max"),
            Any_Active=("Is_Active", "any"),
            Total_Mins=("Duration_Mins", "sum"),
        )
        .sort_values(["Date", "Worker"], ascending=[False, True])
        .reset_index(drop=True)
    )

    day_grouped["First_Login"] = day_grouped["First_Login_DT"].dt.strftime("%I:%M:%S %p")
    day_grouped["Last_Logout"] = day_grouped.apply(
        lambda r: "-" if r["Any_Active"] else r["Last_Logout_DT"].strftime("%I:%M:%S %p"),
        axis=1,
    )
    day_grouped["Total_Hours"] = (day_grouped["Total_Mins"] / 60.0).round(2)
    day_grouped["Total_Duration"] = day_grouped["Total_Mins"].apply(format_duration)

    day_wise_df = day_grouped[
        [
            "Date",
            "Worker",
            "Total_Sessions",
            "First_Login",
            "Last_Logout",
            "Total_Hours",
            "Total_Duration",
        ]
    ].copy()

    worker_grouped = (
        df.groupby("Worker", as_index=False)
        .agg(
            Days_Worked=("Date", "nunique"),
            Total_Sessions=("Login_DT", "count"),
            Last_Login_DT=("Login_DT", "max"),
            Last_Logout_DT=("Effective_Logout_DT", "max"),
            Any_Active=("Is_Active", "any"),
            Total_Mins=("Duration_Mins", "sum"),
        )
        .sort_values("Worker")
        .reset_index(drop=True)
    )

    worker_grouped["Last_Login"] = worker_grouped["Last_Login_DT"].dt.strftime(
        "%I:%M:%S %p"
    )
    worker_grouped["Last_Logout"] = worker_grouped.apply(
        lambda r: "-" if r["Any_Active"] else r["Last_Logout_DT"].strftime("%I:%M:%S %p"),
        axis=1,
    )
    worker_grouped["Total_Hours"] = (worker_grouped["Total_Mins"] / 60.0).round(2)
    worker_grouped["Total_Duration"] = worker_grouped["Total_Mins"].apply(format_duration)
    worker_grouped["Status"] = worker_grouped["Any_Active"].map(
        {True: "Active", False: "Offline"}
    )

    worker_rows_df = worker_grouped[
        [
            "Worker",
            "Days_Worked",
            "Total_Sessions",
            "Last_Login",
            "Last_Logout",
            "Total_Hours",
            "Total_Duration",
            "Status",
        ]
    ].copy()

    total_mins_all = float(df["Duration_Mins"].sum())
    overall_hours = round(total_mins_all / 60.0, 2)
    overall_duration_str = format_duration(total_mins_all)
    total_sessions = int(len(df))

    overall_row = pd.DataFrame(
        [
            {
                "Worker": "OVERALL TOTAL",
                "Days_Worked": "-",
                "Total_Sessions": total_sessions,
                "Last_Login": "-",
                "Last_Logout": "-",
                "Total_Hours": overall_hours,
                "Total_Duration": overall_duration_str,
                "Status": "-",
            }
        ]
    )

    worker_summary_df = pd.concat([worker_rows_df, overall_row], ignore_index=True)

    return (
        session_logs_df,
        day_wise_df,
        worker_summary_df,
        overall_hours,
        overall_duration_str,
        total_sessions,
    )


def create_excel_report(
    worker_summary_df: pd.DataFrame,
    day_wise_df: pd.DataFrame,
    session_logs_df: pd.DataFrame,
) -> bytes:
    """Generate multi-sheet Excel workbook bytes."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        worker_summary_df.to_excel(writer, sheet_name="Worker_Summary", index=False)
        day_wise_df.to_excel(writer, sheet_name="Day_Wise_Hours", index=False)
        session_logs_df.to_excel(writer, sheet_name="All_Session_Logs", index=False)
    output.seek(0)
    return output.getvalue()


def _attendance_records_to_raw_logs_df(attendance_data):
    if not attendance_data:
        return pd.DataFrame(columns=["Worker", "Login_DT", "Logout_DT"])
    raw_rows = []
    now_ist = get_ist_now()
    for row in attendance_data:
        w = str(row.get("Worker", "")).strip()
        if not w:
            continue
        d_clean = _normalize_date_ist(row.get("Date", "")) or now_ist.strftime("%Y-%m-%d")
        t_in = _parse_time_12_or_24(row.get("Login_Time", "-"))
        if not t_in:
            continue
        try:
            d_base = datetime.datetime.strptime(d_clean, "%Y-%m-%d").date()
        except Exception:
            d_base = now_ist.date()

        dt_in = datetime.datetime.combine(d_base, t_in)
        t_out = _parse_time_12_or_24(row.get("Logout_Time", "-"))
        dt_out = None
        if t_out:
            dt_out = datetime.datetime.combine(d_base, t_out)
            if dt_out < dt_in:
                dt_out = dt_out + datetime.timedelta(days=1)

        raw_rows.append({
            "Worker": w,
            "Login_DT": dt_in,
            "Logout_DT": dt_out,
        })
    return pd.DataFrame(raw_rows) if raw_rows else pd.DataFrame(columns=["Worker", "Login_DT", "Logout_DT"])


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
            dur_str = format_duration(dur_sec / 60.0) if (l_clean != "-" and (lo_clean != "-" or is_act)) else "0h 00m"
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
            dur_str = format_duration(dur_sec / 60.0) if (l_clean != "-" and (lo_clean != "-" or is_act)) else "0h 00m"
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
        "Duration": "0h 00m",
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
            "Duration": "0h 00m",
            "Status": "Logged Out",
        })
    save_attendance(attendance)


def build_attendance_summary_tables(attendance_data):
    raw_df = _attendance_records_to_raw_logs_df(attendance_data)
    _, day_wise_df, worker_summary_df, _, _, _ = build_attendance_reports(raw_df, now_dt=get_ist_now())
    if not day_wise_df.empty:
        day_wise_df.index = range(1, len(day_wise_df) + 1)
    if not worker_summary_df.empty:
        worker_summary_df.index = range(1, len(worker_summary_df) + 1)
    return day_wise_df, worker_summary_df


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


@st.cache_data(ttl=300, show_spinner=False)
def load_all_sheets_cached(file_mtime):
    remote_sheets = _fetch_sheets_from_webapp()
    all_sheets = None

    if remote_sheets is not None and len(remote_sheets) > 0:
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
            if not s_clean:
                continue
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


def get_sheet_names():
    all_sheets = load_all_sheets_cached(_get_file_mtime())
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


@st.cache_data(ttl=300, show_spinner=False)
def get_admin_gsheet_export_bytes(file_mtime):
    all_sheets = load_all_sheets_cached(file_mtime)
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        has_data = False
        for s_name, df_s in all_sheets.items():
            if str(s_name).strip() in SYSTEM_SHEETS:
                continue
            if df_s is not None:
                df_clean = reorder_dataframe_columns(df_s.copy())
                df_clean.to_excel(writer, sheet_name=str(s_name)[:31], index=False)
                has_data = True
        if not has_data:
            pd.DataFrame().to_excel(writer, sheet_name="Sheet1", index=False)
    return excel_buffer.getvalue()


@st.cache_data(ttl=300, show_spinner=False)
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
                    u_rows_s.to_excel(writer, sheet_name=str(s)[:31], index=False)
        if not has_user_data and current_sheet in all_sheets:
            df_cur = all_sheets[current_sheet]
            login_cs = find_login_user_column(df_cur)
            u_rows_cur = df_cur[
                df_cur[login_cs].astype(str).str.strip().str.lower()
                == current_user.lower()
            ]
            u_rows_cur = reorder_dataframe_columns(u_rows_cur)
            u_rows_cur.to_excel(writer, sheet_name=str(current_sheet)[:31], index=False)
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
    cached_sheets = load_all_sheets_cached(_get_file_mtime())
    if cached_sheets:
        for s_k, df_k in cached_sheets.items():
            sheets_data[str(s_k)] = df_k.copy()

    if os.path.exists(EXCEL_FILE):
        try:
            disk_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
            if disk_sheets:
                for s_d, df_d in disk_sheets.items():
                    if str(s_d) not in sheets_data:
                        sheets_data[str(s_d)] = df_d.copy()
        except Exception:
            pass

    target_sheet_key = str(sheet_name)
    if isinstance(sheet_name, int) and sheets_data:
        keys = list(sheets_data.keys())
        if 0 <= sheet_name < len(keys):
            target_sheet_key = str(keys[sheet_name])

    df_to_save = reorder_dataframe_columns(df_to_save)
    sheets_data[target_sheet_key] = df_to_save

    try:
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            for s, df_val in sheets_data.items():
                df_val = reorder_dataframe_columns(df_val)
                df_val.to_excel(writer, sheet_name=str(s)[:31], index=False)
    except Exception:
        pass

    _push_sheet_to_webapp(target_sheet_key, df_to_save, sheets_data)
    st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
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
        if str(s_name).strip() in SYSTEM_SHEETS:
            continue
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


def get_all_workers_list():
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    all_workers = set()
    for s_name, df_s in all_sheets.items():
        if str(s_name).strip() in SYSTEM_SHEETS:
            continue
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
    for extra_w in ["Kalyan", "Anil"]:
        if not any(w.lower() == extra_w.lower() for w in all_workers):
            all_workers.add(extra_w)
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
        close_case_popup_state()
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

default_s_init = st.session_state.get("active_sheet", "Sheet1")
if st.session_state.get("logged_in") or st.session_state.get("open_dialog"):
    sheets_init = get_sheet_names()
    default_s_init = get_default_active_sheet(sheets_init)
    if "active_sheet" not in st.session_state:
        st.session_state["active_sheet"] = default_s_init
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
    restrict_edit_to_user=None,
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
        ["Appointment_PDF", "__orig_idx__"]
        if allow_edit
        else [p_col, "Appointment_PDF", "__orig_idx__"]
    )
    display_cols = [c for c in page_df.columns if c not in cols_to_hide]

    if "Sheet_Name" in display_cols:
        display_cols = ["Sheet_Name"] + [
            c for c in display_cols if c != "Sheet_Name"
        ]

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

    q_logged = urllib.parse.quote(str(curr_logged))
    q_user = urllib.parse.quote(str(curr_user))
    q_role = urllib.parse.quote(str(curr_role))

    indices = page_df.index.tolist()
    orig_indices = (
        page_df["__orig_idx__"].tolist()
        if "__orig_idx__" in page_df.columns
        else indices
    )
    row_sheets = (
        page_df["Sheet_Name"].astype(str).tolist()
        if "Sheet_Name" in page_df.columns
        else [curr_sheet] * len(page_df)
    )
    login_c_tbl = find_login_user_column(page_df)
    row_workers = (
        page_df[login_c_tbl].astype(str).str.strip().tolist()
        if login_c_tbl in page_df.columns
        else [""] * len(page_df)
    )

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
        target_row_id = orig_indices[i]
        target_sheet_name = row_sheets[i] if row_sheets[i] else curr_sheet
        q_sheet_row = urllib.parse.quote(str(target_sheet_name))
        row_link = f"?logged_in={q_logged}&username={q_user}&role={q_role}&active_sheet={q_sheet_row}&edit_row={target_row_id}"

        row_can_edit = allow_edit
        if row_can_edit and restrict_edit_to_user:
            if row_workers[i].lower() != restrict_edit_to_user.strip().lower():
                row_can_edit = False

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
            if col == "Sheet_Name":
                cell_content = f'<span style="font-weight:700; color:#0f766e; background:#f0fdfa; padding:2px 8px; border-radius:6px; border:1px solid #99f6e4;">{val if val != "" else "-"}</span>'
            elif col == "Case_ID":
                cell_content = f'<span style="font-weight:800; color:#1e3a8a; background:#eff6ff; padding:2px 8px; border-radius:6px; border:1px solid #bfdbfe;">{val if val != "" else "-"}</span>'
            elif col == display_c:
                display_val = (
                    val.strip() if val.strip() != "" else "Not Assigned"
                )
                if row_can_edit:
                    cell_content = f'<a href="{row_link}" target="_self" class="username-hyperlink">{display_val}</a>'
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
        elif b_status.lower() == "issues":
            status_display = "Issues"
            status_color = "#7c3aed"
        else:
            status_display = "Active"
            status_color = "#1e293b"

        html_parts.append(
            f'<td style="padding: 9px 10px; border: 1px solid #e2e8f0; color: {status_color}; font-weight: 700; white-space: nowrap;">{status_display}</td>'
        )
        html_parts.append("</tr>")

    html_parts.append("</tbody></table></div>")
    st.markdown("".join(html_parts), unsafe_allow_html=True)


@st.dialog("📋 Case Details, Edit & Appointment PDF", width="large", on_dismiss=close_case_popup_state)
def view_and_edit_case_dialog(sheet_name, row_idx, role, current_user, can_update=True):
    enable_right_cursor_navigation_in_popup()
    df_full = load_data(sheet_name=sheet_name)
    if df_full is None or row_idx not in df_full.index:
        st.error("Case data not found.")
        if st.button("✖️ Close Popup", key="dlg_missing_close"):
            close_case_popup_state()
            st.rerun()
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

        issue_rem_view = str(row_data.get("Issue_Remark", "")).strip()
        if issue_rem_view:
            st.warning(f"⚠️ **Issue Remark:** {issue_rem_view}")

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
                                close_case_popup_state()
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
                                close_case_popup_state()
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
            f"### 🔄 Reschedule Appointment Case (Booked ID stays: `{case_id_val}` | New ID will be created)"
        )
        all_workers_res = get_all_workers_list()
        worker_opts_res = ["Not Assigned"] + [
            w for w in all_workers_res if w.lower() != "not assigned"
        ]
        curr_w_val = str(row_data.get(login_c, "")).strip()
        preferred_res_worker = (
            "Anil"
            if (curr_w_val.lower() == "kalyan" or current_user.lower() == "kalyan")
            and "Anil" in worker_opts_res
            else curr_w_val
        )
        def_w_idx = (
            worker_opts_res.index(preferred_res_worker)
            if preferred_res_worker in worker_opts_res
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
                        original_booked_worker = str(booked_snapshot_row.get(login_c, "")).strip()

                        if is_booked_sheet:
                            df_full.at[row_idx, "Case_ID"] = case_id_val
                            df_full.at[row_idx, login_c] = original_booked_worker
                            df_full.at[row_idx, "Booking_Status"] = "Booked"
                            df_full.at[row_idx, "Action_Type"] = "Booked"
                        else:
                            new_res_id = generate_next_case_id()
                            df_full.at[row_idx, "Case_ID"] = new_res_id
                            df_full.at[row_idx, "Preferred_Location"] = (
                                pref_loc_input.strip()
                            )
                            df_full.at[row_idx, "Preferred_Dates"] = (
                                pref_dates_input.strip()
                            )
                            df_full.at[row_idx, "Prior_Notice"] = (
                                prior_notice_input.strip()
                            )
                            df_full.at[row_idx, "Booking_Status"] = "Rescheduled"
                            df_full.at[row_idx, "Updated_By"] = current_user
                            df_full.at[row_idx, "Updated_At"] = ts_now
                            df_full.at[row_idx, "Action_Type"] = "Rescheduled Case"

                        save_data(df_full, sheet_name=sheet_name)

                        if is_booked_sheet:
                            append_rescheduled_to_active_sheet(
                                source_row=booked_snapshot_row,
                                pref_loc=pref_loc_input,
                                pref_dates=pref_dates_input,
                                prior_notice=prior_notice_input,
                                updated_by=current_user,
                                updated_at=ts_now,
                                assigned_worker=assign_worker_res,
                            )

                        close_case_popup_state()
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
                        close_case_popup_state()
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error: {e}")

            if cancelled_req:
                st.session_state["is_delete_request_mode"] = False
                st.rerun()

    else:
        st.markdown("### ✏️ Edit Case & Booking Status")
        curr_bs = str(row_data.get("Booking_Status", "Active")).strip()
        opts = ["Active", "Booked", "Rescheduled", "Issues"]
        b_idx = next(
            (i for i, o in enumerate(opts) if o.lower() == curr_bs.lower()),
            0,
        )

        booking_status_choice = st.selectbox(
            "📌 Select Status",
            options=opts,
            index=b_idx,
            key=f"popup_bs_{row_idx}",
        )

        with st.form(f"popup_edit_form_{row_idx}"):
            updated_values = {}

            all_workers_list = get_all_workers_list()
            if not all_workers_list:
                all_workers_list = ["Not Assigned"]
            elif "Not Assigned" not in all_workers_list:
                all_workers_list.insert(0, "Not Assigned")

            cols_to_show = (
                [c for c in df_full.columns if c != "Issue_Remark"]
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
                        "Issue_Remark",
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
                                parsed_d = get_ist_now().date()
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

            app_entry_val = ""
            uploaded_pdf_file = None
            popup_pref_loc = ""
            popup_pref_dates = ""
            popup_prior_notice = ""
            popup_issue_remark = ""

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
            elif booking_status_choice == "Issues":
                popup_issue_remark = st.text_area(
                    "📝 Issue Remark (Required)",
                    value=str(row_data.get("Issue_Remark", "")).strip(),
                    placeholder="Enter issue remark / reason before saving...",
                    key=f"popup_issue_remark_{row_idx}",
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
                    elif booking_status_choice == "Issues" and not popup_issue_remark.strip():
                        st.warning(
                            "⚠️ Please fill the Issue Remark text box before saving Issues status!"
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

                        active_case_id = case_id_val
                        if booking_status_choice == "Rescheduled" and not is_booked_sheet:
                            active_case_id = generate_next_case_id()

                        df_full.at[row_idx, "Case_ID"] = active_case_id

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
                                pdf_filename = f"Appointment_{active_case_id}_{safe_uname}_{int(current_time.timestamp())}.pdf"
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

                        elif booking_status_choice == "Issues":
                            df_full.at[row_idx, "Issue_Remark"] = (
                                popup_issue_remark.strip()
                            )
                            changed = True

                        if changed:
                            df_full.at[row_idx, "Updated_By"] = current_user
                            df_full.at[row_idx, "Updated_At"] = current_time_str
                            if booking_status_choice == "Rescheduled":
                                df_full.at[row_idx, "Action_Type"] = "Rescheduled Case"
                            elif booking_status_choice == "Issues":
                                df_full.at[row_idx, "Action_Type"] = "Issues Case"
                            else:
                                df_full.at[row_idx, "Action_Type"] = "Updated Case"

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

                        elif booking_status_choice == "Issues":
                            sheets_all = get_sheet_names()
                            issues_sheet_name = next(
                                (s for s in sheets_all if "issue" in s.lower()),
                                "Issues",
                            )
                            if issues_sheet_name != sheet_name:
                                df_target_issues = load_data(
                                    sheet_name=issues_sheet_name
                                )
                                row_to_transfer = df_full.loc[row_idx].copy()
                                if df_target_issues is None:
                                    df_target_issues = pd.DataFrame(
                                        columns=df_full.columns
                                    )
                                for col in df_target_issues.columns:
                                    if col not in row_to_transfer:
                                        row_to_transfer[col] = ""

                                df_target_issues = pd.concat(
                                    [
                                        df_target_issues,
                                        pd.DataFrame([row_to_transfer]),
                                    ],
                                    ignore_index=True,
                                )
                                save_data(
                                    df_target_issues, sheet_name=issues_sheet_name
                                )
                                df_full = df_full.drop(row_idx)

                        save_data(df_full, sheet_name=sheet_name)

                        if booking_status_choice == "Rescheduled" and is_booked_sheet:
                            active_worker_pick = str(
                                updated_values.get(login_c, df_full.at[row_idx, login_c])
                            ).strip()
                            orig_worker_booked = str(row_data.get(login_c, "")).strip()
                            df_full.at[row_idx, "Case_ID"] = case_id_val
                            df_full.at[row_idx, login_c] = orig_worker_booked
                            df_full.at[row_idx, "Booking_Status"] = "Booked"
                            df_full.at[row_idx, "Action_Type"] = "Booked"
                            save_data(df_full, sheet_name=sheet_name)

                            append_rescheduled_to_active_sheet(
                                source_row=df_full.loc[row_idx].copy(),
                                pref_loc=popup_pref_loc,
                                pref_dates=popup_pref_dates,
                                prior_notice=popup_prior_notice,
                                updated_by=current_user,
                                updated_at=current_time_str,
                                assigned_worker=active_worker_pick,
                            )

                        close_case_popup_state()
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ Error updating case: {e}")

            if cancelled:
                close_case_popup_state()
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


@st.dialog("📊 Case Worker Attendance & Hours Report", width="large")
def admin_attendance_dialog():
    st.info("💡 Case Worker-wise & Overall login/logout hours summary along with daily & session logs:")
    attendance_data = load_attendance()
    if not attendance_data:
        st.warning("No attendance records found.")
    else:
        raw_logs_df = _attendance_records_to_raw_logs_df(attendance_data)
        (
            session_logs_df,
            day_wise_df,
            worker_summary_df,
            overall_hours,
            overall_duration_str,
            total_sessions,
        ) = build_attendance_reports(raw_logs_df, now_dt=get_ist_now())

        if not day_wise_df.empty:
            day_wise_df.index = range(1, len(day_wise_df) + 1)
        if not worker_summary_df.empty:
            worker_summary_df.index = range(1, len(worker_summary_df) + 1)
        if not session_logs_df.empty:
            session_logs_df.index = range(1, len(session_logs_df) + 1)

        st.success(
            f"⏱️ Overall Total Logged Hours: **{overall_duration_str}** ({overall_hours} hrs) | Total Sessions: **{total_sessions}**"
        )

        tab_ov, tab_dy, tab_lg = st.tabs([
            "👤 Case Worker Wise & Overall Total",
            "📅 Day-wise Worker Hours",
            "🕒 All Session Logs",
        ])

        with tab_ov:
            st.markdown("#### Case Worker Wise & Overall Hours Summary")
            st.dataframe(worker_summary_df, use_container_width=True)

        with tab_dy:
            st.markdown("#### Day-wise Worker Hours (All Sessions Combined per Day)")
            st.dataframe(day_wise_df, use_container_width=True)

        with tab_lg:
            st.markdown("#### Detailed Login / Logout Session Logs")
            st.dataframe(session_logs_df, use_container_width=True)

        excel_bytes = create_excel_report(worker_summary_df, day_wise_df, session_logs_df)
        st.download_button(
            label="📥 Download Attendance & Hours Report (.xlsx)",
            data=excel_bytes,
            file_name=f"Case_Worker_Attendance_Report_{get_ist_now().strftime('%Y-%m-%d')}.xlsx",
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
        if str(s).strip() in SYSTEM_SHEETS:
            continue
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
        sheets = [s for s in all_sheets.keys() if str(s).strip() not in SYSTEM_SHEETS]
        if not sheets:
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
                    value=get_ist_now().date(),
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
        if str(s).strip() in SYSTEM_SHEETS:
            continue
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
    for s_name, df_s in all_sheets.items():
        if str(s_name).strip() in SYSTEM_SHEETS:
            continue
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
                        all_sheets_data = load_all_sheets_cached(_get_file_mtime())
                        for s, df_s in all_sheets_data.items():
                            if str(s).strip() in SYSTEM_SHEETS:
                                continue
                            df_copy = df_s.copy()
                            lc = find_login_user_column(df_copy)
                            pc = find_password_column(df_copy)
                            if lc in df_copy.columns and pc in df_copy.columns:
                                mask = (
                                    df_copy[lc]
                                    .astype(str)
                                    .str.strip()
                                    .str.lower()
                                    == u_name.lower()
                                )
                                if mask.any():
                                    df_copy.loc[mask, pc] = new_pwd_val.strip()
                            all_sheets_data[s] = reorder_dataframe_columns(
                                df_copy
                            )

                        with pd.ExcelWriter(
                            EXCEL_FILE, engine="openpyxl"
                        ) as writer:
                            for s, df_val in all_sheets_data.items():
                                df_val.to_excel(
                                    writer, sheet_name=str(s)[:31], index=False
                                )

                        if all_sheets_data:
                            first_key = list(all_sheets_data.keys())[0]
                            _push_sheet_to_webapp(
                                first_key, all_sheets_data[first_key], all_sheets_data
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
                        st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
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
        "💡 Upload your updated local Excel file here to instantly synchronize it with Google Sheets and Cloud."
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
                uploaded_sheets = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
                if uploaded_sheets:
                    first_k = list(uploaded_sheets.keys())[0]
                    _push_sheet_to_webapp(first_k, uploaded_sheets[first_k], uploaded_sheets)
                st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
                st.cache_data.clear()
                ensure_all_cases_have_unique_ids(force=True)
                st.success("✅ Cloud & Google Sheet data updated successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {e}")


@st.dialog("🔑 Change Password", width="medium")
def worker_password_dialog(worker_name):
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    worker_current_pass = ""
    for s_name, df_chk in all_sheets.items():
        if str(s_name).strip() in SYSTEM_SHEETS:
            continue
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
                    all_sheets_data = load_all_sheets_cached(_get_file_mtime())
                    for s, df_s in all_sheets_data.items():
                        if str(s).strip() in SYSTEM_SHEETS:
                            continue
                        df_copy = df_s.copy()
                        lc = find_login_user_column(df_copy)
                        pc = find_password_column(df_copy)
                        if lc in df_copy.columns and pc in df_copy.columns:
                            mask = (
                                df_copy[lc].astype(str).str.strip().str.lower()
                                == worker_name.lower()
                            )
                            if mask.any():
                                df_copy.loc[mask, pc] = new_wp.strip()
                        all_sheets_data[s] = reorder_dataframe_columns(df_copy)
                    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
                        for s, df_val in all_sheets_data.items():
                            df_val.to_excel(
                                writer, sheet_name=str(s)[:31], index=False
                            )
                    if all_sheets_data:
                        first_k = list(all_sheets_data.keys())[0]
                        _push_sheet_to_webapp(first_k, all_sheets_data[first_k], all_sheets_data)
                    st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
                    st.cache_data.clear()
                    st.success("✅ Password updated successfully!")
                except Exception as e:
                    st.error(f"❌ Error: {e}")


@st.dialog("➕ Add New Case", width="large")
def add_case_dialog(worker_name):
    sheets = get_sheet_names()
    default_idx = 0
    for idx_s, s_nm in enumerate(sheets):
        if "booked" not in s_nm.lower() and "issue" not in s_nm.lower():
            default_idx = idx_s
            break

    selected_target_sheet = st.selectbox(
        "Select Target Sheet for New Case",
        options=sheets,
        index=default_idx,
        key="dlg_target_sheet",
    )
    target_df = load_data(sheet_name=selected_target_sheet)
    if target_df is None:
        target_df = pd.DataFrame(columns=["Case_ID", "USERNAME", "Password", "Lead"])

    target_df = reorder_dataframe_columns(target_df)
    login_c = find_login_user_column(target_df)
    display_c = find_display_username_column(target_df)
    p_col = find_password_column(target_df)

    next_cid = generate_next_case_id()
    all_workers = get_all_workers_list()
    worker_options = [worker_name] if worker_name else []
    for w in all_workers:
        if w and w.lower() != str(worker_name).strip().lower() and w not in worker_options:
            worker_options.append(w)
    if "Not Assigned" not in worker_options:
        worker_options.append("Not Assigned")

    excluded_add_cols = {
        "Case_ID",
        "Updated_By",
        "Updated_At",
        "Action_Type",
        "Delete_Reason",
        "Booking_Status",
        "Issue_Remark",
        "Appointment_Entry",
        "Appointment_PDF",
        "Preferred_Location",
        "Preferred_Dates",
        "Prior_Notice",
    }
    editable_cols = [c for c in target_df.columns if c not in excluded_add_cols]

    with st.form(f"add_case_form_{selected_target_sheet}"):
        st.text_input("Auto Generated Case_ID", value=next_cid, disabled=True)
        new_row_vals = {col: "" for col in target_df.columns}
        new_row_vals["Case_ID"] = next_cid

        col_a, col_b = st.columns(2)
        for idx_c, col_nm in enumerate(editable_cols):
            tgt_box = col_a if idx_c % 2 == 0 else col_b
            with tgt_box:
                if col_nm == login_c:
                    picked_w = st.selectbox(
                        f"{col_nm} (Assigned Case Worker)",
                        options=worker_options,
                        index=0,
                        key=f"add_w_{selected_target_sheet}_{col_nm}",
                    )
                    new_row_vals[col_nm] = "" if picked_w == "Not Assigned" else picked_w
                elif col_nm == p_col:
                    new_row_vals[col_nm] = st.text_input(
                        f"{col_nm}",
                        value="",
                        key=f"add_p_{selected_target_sheet}_{col_nm}",
                    )
                else:
                    c_low = col_nm.lower()
                    if "date" in c_low and "prefer" not in c_low:
                        d_pick = st.date_input(
                            f"{col_nm}",
                            value=get_ist_now().date(),
                            key=f"add_d_{selected_target_sheet}_{col_nm}",
                        )
                        new_row_vals[col_nm] = d_pick.strftime("%Y-%m-%d")
                    else:
                        new_row_vals[col_nm] = st.text_input(
                            f"{col_nm}",
                            value="",
                            key=f"add_f_{selected_target_sheet}_{col_nm}",
                        )

        status_new = st.selectbox(
            "Initial Booking Status",
            options=["Active", "Booked", "Rescheduled", "Issues"],
            index=0,
            key=f"add_status_{selected_target_sheet}",
        )
        remark_or_entry = st.text_input(
            "Appointment Entry / Issue Remark / Notes (Optional)",
            value="",
            key=f"add_note_{selected_target_sheet}",
        )

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            sub_add = st.form_submit_button(
                "💾 Save New Case", type="primary", use_container_width=True
            )
        with col_s2:
            can_add = st.form_submit_button("↩️ Cancel", use_container_width=True)

        if can_add:
            st.session_state["open_dialog"] = None
            st.rerun()

        if sub_add:
            main_ident = str(new_row_vals.get(display_c, "")).strip()
            if not main_ident:
                st.warning(f"⚠️ Please enter '{display_c}' before saving the case.")
            else:
                try:
                    now_ist = get_ist_now()
                    ts_str = now_ist.strftime("%Y-%m-%d %I:%M:%S %p")
                    final_cid = generate_next_case_id(extra_reserved_ids=[next_cid])
                    new_row_vals["Case_ID"] = final_cid
                    new_row_vals["Booking_Status"] = status_new
                    new_row_vals["Updated_By"] = worker_name or "Admin"
                    new_row_vals["Updated_At"] = ts_str
                    new_row_vals["Action_Type"] = "New Case"
                    if status_new == "Booked" and remark_or_entry.strip():
                        new_row_vals["Appointment_Entry"] = remark_or_entry.strip()
                    elif status_new == "Issues" and remark_or_entry.strip():
                        new_row_vals["Issue_Remark"] = remark_or_entry.strip()

                    df_updated = pd.concat(
                        [target_df, pd.DataFrame([new_row_vals])], ignore_index=True
                    )
                    save_data(df_updated, sheet_name=selected_target_sheet)
                    st.session_state["open_dialog"] = None
                    st.success(f"✅ Case '{final_cid}' added to '{selected_target_sheet}'!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error adding case: {e}")


# ==============================================================================
# CROSS-CASE EDIT PERMISSION WORKFLOW (CASE WORKER <-> ADMIN APPROVAL)
# ==============================================================================
EDIT_REQ_COLS = [
    "Request_ID",
    "Case_ID",
    "Sheet_Name",
    "Row_Index",
    "Case_Username",
    "Case_Owner",
    "Requested_By",
    "Reason",
    "Status",
    "Requested_At",
    "Reviewed_By",
    "Reviewed_At",
]


def load_edit_requests():
    rows = _read_sheet_rows("App_Edit_Requests")
    cleaned = []
    for r in rows:
        cid = str(r.get("Case_ID", "")).strip()
        req_by = str(r.get("Requested_By", "")).strip()
        if cid or req_by:
            cleaned.append({
                "Request_ID": str(r.get("Request_ID", "")).strip(),
                "Case_ID": cid,
                "Sheet_Name": str(r.get("Sheet_Name", "")).strip(),
                "Row_Index": str(r.get("Row_Index", "")).strip(),
                "Case_Username": str(r.get("Case_Username", "")).strip(),
                "Case_Owner": str(r.get("Case_Owner", "")).strip(),
                "Requested_By": req_by,
                "Reason": str(r.get("Reason", "")).strip(),
                "Status": str(r.get("Status", "Pending")).strip() or "Pending",
                "Requested_At": str(r.get("Requested_At", "")).strip(),
                "Reviewed_By": str(r.get("Reviewed_By", "")).strip(),
                "Reviewed_At": str(r.get("Reviewed_At", "")).strip(),
            })
    return cleaned


def save_edit_requests(req_list):
    records = []
    for item in (req_list or []):
        records.append({col: str(item.get(col, "")).strip() for col in EDIT_REQ_COLS})
    _write_sheet_records("App_Edit_Requests", EDIT_REQ_COLS, records)


def has_approved_cross_case_edit(worker_name, sheet_name, case_id, row_idx):
    if not worker_name:
        return False
    w_low = str(worker_name).strip().lower()
    cid_str = str(case_id).strip()
    s_str = str(sheet_name).strip()
    r_str = str(row_idx).strip()
    for req in load_edit_requests():
        if (
            req.get("Requested_By", "").strip().lower() == w_low
            and req.get("Status", "").strip().lower() == "approved"
            and req.get("Sheet_Name", "").strip() == s_str
            and (
                (cid_str and cid_str != "N/A" and req.get("Case_ID", "").strip() == cid_str)
                or req.get("Row_Index", "").strip() == r_str
            )
        ):
            return True
    return False


def consume_approved_cross_case_edit(worker_name, sheet_name, case_id, row_idx):
    w_low = str(worker_name).strip().lower()
    cid_str = str(case_id).strip()
    s_str = str(sheet_name).strip()
    r_str = str(row_idx).strip()
    reqs = load_edit_requests()
    changed = False
    for req in reqs:
        if (
            req.get("Requested_By", "").strip().lower() == w_low
            and req.get("Status", "").strip().lower() == "approved"
            and req.get("Sheet_Name", "").strip() == s_str
            and (
                (cid_str and cid_str != "N/A" and req.get("Case_ID", "").strip() == cid_str)
                or req.get("Row_Index", "").strip() == r_str
            )
        ):
            req["Status"] = "Completed"
            req["Reviewed_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
            changed = True
    if changed:
        save_edit_requests(reqs)


@st.dialog("🔐 Request Cross-Case Edit Permission", width="medium")
def request_cross_case_edit_dialog(sheet_name, row_idx, case_id, case_user, case_owner, requester_name):
    st.warning(
        f"🔒 **Case Owner Mismatch:** This case belongs to **{case_owner or 'Not Assigned'}** (`{case_id}`). "
        f"You are logged in as **{requester_name}**. Admin approval is required to edit another worker's case."
    )

    existing_reqs = [
        r for r in load_edit_requests()
        if r.get("Requested_By", "").strip().lower() == str(requester_name).strip().lower()
        and r.get("Sheet_Name", "").strip() == str(sheet_name).strip()
        and str(r.get("Row_Index", "")).strip() == str(row_idx).strip()
    ]
    if existing_reqs:
        latest = existing_reqs[-1]
        st.info(
            f"📌 Latest Request Status: **{latest.get('Status')}** | Reason: *{latest.get('Reason')}*"
        )

    with st.form(f"cross_edit_req_form_{sheet_name}_{row_idx}"):
        reason_txt = st.text_area(
            "✍️ Why do you need permission to edit this case? (Required)",
            placeholder="Enter clear reason for Admin approval...",
        )
        c1, c2 = st.columns(2)
        with c1:
            sub_req = st.form_submit_button(
                "📤 Send Request to Admin", type="primary", use_container_width=True
            )
        with c2:
            can_req = st.form_submit_button("↩️ Cancel", use_container_width=True)

        if can_req:
            close_case_popup_state()
            st.rerun()

        if sub_req:
            if not reason_txt.strip():
                st.warning("⚠️ Please enter a reason before sending the edit request.")
            else:
                reqs = load_edit_requests()
                req_id = f"ER{1001 + len(reqs)}"
                reqs.append({
                    "Request_ID": req_id,
                    "Case_ID": str(case_id).strip(),
                    "Sheet_Name": str(sheet_name).strip(),
                    "Row_Index": str(row_idx).strip(),
                    "Case_Username": str(case_user).strip(),
                    "Case_Owner": str(case_owner).strip() or "Not Assigned",
                    "Requested_By": str(requester_name).strip(),
                    "Reason": reason_txt.strip(),
                    "Status": "Pending",
                    "Requested_At": get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p"),
                    "Reviewed_By": "",
                    "Reviewed_At": "",
                })
                save_edit_requests(reqs)
                close_case_popup_state()
                st.success(f"✅ Edit permission request ({req_id}) sent to Admin!")
                st.rerun()


@st.dialog("🔐 Pending Cross-Case Edit Requests", width="large")
def admin_edit_requests_dialog(admin_username="admin"):
    st.info("💡 Review Case Worker requests to edit cases assigned to other workers:")
    reqs = load_edit_requests()
    pending_reqs = [r for r in reqs if r.get("Status", "").strip().lower() == "pending"]
    history_reqs = [r for r in reqs if r.get("Status", "").strip().lower() != "pending"]

    tab_pend, tab_hist = st.tabs([
        f"⏳ Pending Requests ({len(pending_reqs)})",
        f"📜 History ({len(history_reqs)})",
    ])

    with tab_pend:
        if not pending_reqs:
            st.info("No pending cross-case edit requests.")
        else:
            for idx_p, r in enumerate(pending_reqs):
                with st.container(border=True):
                    st.markdown(
                        f"**🆔 {r.get('Request_ID')}** | Case ID: **`{r.get('Case_ID')}`** | "
                        f"Case User: **`{r.get('Case_Username')}`** | Sheet: **`{r.get('Sheet_Name')}`**"
                    )
                    st.markdown(
                        f"🙋 **Requested By:** `{r.get('Requested_By')}` &nbsp;→&nbsp; "
                        f"👤 **Case Owner:** `{r.get('Case_Owner')}`"
                    )
                    st.markdown(f"✍️ **Reason:** {r.get('Reason')}")
                    st.caption(f"Requested At: {r.get('Requested_At')}")

                    b_col1, b_col2 = st.columns(2)
                    with b_col1:
                        if st.button(
                            "✅ Approve Edit Permission",
                            key=f"approve_edit_{r.get('Request_ID')}_{idx_p}",
                            type="primary",
                            use_container_width=True,
                        ):
                            for item in reqs:
                                if item.get("Request_ID") == r.get("Request_ID"):
                                    item["Status"] = "Approved"
                                    item["Reviewed_By"] = admin_username
                                    item["Reviewed_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                            save_edit_requests(reqs)
                            st.success(f"✅ Approved edit permission for {r.get('Requested_By')} on {r.get('Case_ID')}!")
                            st.rerun()
                    with b_col2:
                        if st.button(
                            "❌ Reject Request",
                            key=f"reject_edit_{r.get('Request_ID')}_{idx_p}",
                            use_container_width=True,
                        ):
                            for item in reqs:
                                if item.get("Request_ID") == r.get("Request_ID"):
                                    item["Status"] = "Rejected"
                                    item["Reviewed_By"] = admin_username
                                    item["Reviewed_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                            save_edit_requests(reqs)
                            st.warning(f"❌ Rejected edit request {r.get('Request_ID')}.")
                            st.rerun()

    with tab_hist:
        if not history_reqs:
            st.info("No past edit request records.")
        else:
            df_hist = pd.DataFrame(history_reqs)
            st.dataframe(df_hist, use_container_width=True)


# ==============================================================================
# ADMIN & WORKER PORTALS + LOGIN FLOW
# ==============================================================================
def render_admin_portal():
    admin_u, _ = get_admin_creds()
    st.markdown(
        """
        <div class="workspace-header-bar">
            <div>🛡️ ADMIN COMMAND CENTER — HemanVisa Operations</div>
            <div class="workspace-sheet-pill">Role: Administrator</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    all_sheets = load_all_sheets_cached(_get_file_mtime())
    business_sheets = [s for s in all_sheets.keys() if str(s).strip() not in SYSTEM_SHEETS]
    if not business_sheets:
        business_sheets = ["Sheet1"]

    active_s = st.session_state.get("active_sheet", business_sheets[0])
    if active_s not in business_sheets:
        active_s = business_sheets[0]
        st.session_state["active_sheet"] = active_s

    pending_del_count = 0
    for s_nm, df_s in all_sheets.items():
        if str(s_nm).strip() not in SYSTEM_SHEETS and df_s is not None and not df_s.empty:
            if "Action_Type" in df_s.columns:
                pending_del_count += int(
                    (df_s["Action_Type"].astype(str).str.strip().str.lower() == "delete request").sum()
                )

    pending_edit_count = sum(
        1 for r in load_edit_requests() if r.get("Status", "").strip().lower() == "pending"
    )

    df_curr = load_data(sheet_name=active_s)
    total_cases_cnt = len(df_curr) if df_curr is not None else 0
    today_str = get_ist_now().strftime("%Y-%m-%d")
    today_upd_cnt = (
        int(df_curr.apply(lambda r: is_today_updated_row(r, today_str), axis=1).sum())
        if df_curr is not None and not df_curr.empty
        else 0
    )
    not_assigned_cnt = 0
    if df_curr is not None and not df_curr.empty:
        lc_curr = find_login_user_column(df_curr)
        not_assigned_cnt = int(df_curr[lc_curr].apply(is_not_assigned_value).sum())

    k1, k2, k3, k4, k5 = st.columns(5, gap="small")
    with k1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Cases ({active_s})</div><div class="kpi-value">{total_cases_cnt}</div></div>', unsafe_allow_html=True)
        if st.button("View All", key="kpi_btn_all", use_container_width=True):
            admin_kpi_cases_dialog(active_s, "all")
    with k2:
        st.markdown(f'<div class="kpi-card green"><div class="kpi-title">Updated Today</div><div class="kpi-value">{today_upd_cnt}</div></div>', unsafe_allow_html=True)
        if st.button("View Today", key="kpi_btn_today", use_container_width=True):
            admin_kpi_cases_dialog(active_s, "today_updated")
    with k3:
        st.markdown(f'<div class="kpi-card gold"><div class="kpi-title">Not Assigned</div><div class="kpi-value">{not_assigned_cnt}</div></div>', unsafe_allow_html=True)
        if st.button("View Unassigned", key="kpi_btn_unass", use_container_width=True):
            admin_kpi_cases_dialog(active_s, "not_assigned")
    with k4:
        st.markdown(f'<div class="kpi-card rose"><div class="kpi-title">Delete Requests</div><div class="kpi-value">{pending_del_count}</div></div>', unsafe_allow_html=True)
        if st.button("Review Deletes", key="kpi_btn_del", use_container_width=True):
            admin_delete_requests_dialog()
    with k5:
        st.markdown(f'<div class="kpi-card purple"><div class="kpi-title">Edit Requests</div><div class="kpi-value">{pending_edit_count}</div></div>', unsafe_allow_html=True)
        if st.button("Review Edits", key="kpi_btn_edit_req", use_container_width=True):
            admin_edit_requests_dialog(admin_username=admin_u)

    st.markdown('<div class="action-group-title">Admin Controls & Reports</div>', unsafe_allow_html=True)
    b_cols = st.columns(8, gap="small")
    with b_cols[0]:
        st.markdown('<div class="admin-btn-0"></div>', unsafe_allow_html=True)
        if st.button("👥 Users & Perms", key="adm_perms", use_container_width=True):
            permissions_dialog()
    with b_cols[1]:
        st.markdown('<div class="admin-btn-1"></div>', unsafe_allow_html=True)
        if st.button("🟢 Login Status", key="adm_logins", use_container_width=True):
            admin_login_status_dialog()
    with b_cols[2]:
        st.markdown('<div class="admin-btn-2"></div>', unsafe_allow_html=True)
        if st.button("⏱️ Attendance", key="adm_att", use_container_width=True):
            admin_attendance_dialog()
    with b_cols[3]:
        st.markdown('<div class="admin-btn-3"></div>', unsafe_allow_html=True)
        if st.button("📊 Worker Report", key="adm_wrep", use_container_width=True):
            admin_worker_wise_report_dialog()
    with b_cols[4]:
        st.markdown('<div class="admin-btn-4"></div>', unsafe_allow_html=True)
        if st.button("📥 PDF Center", key="adm_pdfs", use_container_width=True):
            admin_pdf_download_dialog()
    with b_cols[5]:
        st.markdown('<div class="admin-btn-5"></div>', unsafe_allow_html=True)
        if st.button("📢 Notice Board", key="adm_notice", use_container_width=True):
            admin_notice_dialog()
    with b_cols[6]:
        st.markdown('<div class="admin-btn-6"></div>', unsafe_allow_html=True)
        if st.button("📤 Sync Excel", key="adm_sync", use_container_width=True):
            sync_excel_dialog()
    with b_cols[7]:
        st.markdown('<div class="admin-btn-7"></div>', unsafe_allow_html=True)
        if st.button("🔑 Admin Pass", key="adm_pass", use_container_width=True):
            admin_password_dialog()

    r_cols = st.columns([2.2, 2.2, 1.4, 1.4, 1.2], gap="small")
    with r_cols[0]:
        picked_sheet = st.selectbox("Select Sheet", options=business_sheets, index=business_sheets.index(active_s), key="adm_sheet_sel")
        if picked_sheet != active_s:
            st.session_state["active_sheet"] = picked_sheet
            st.rerun()
    with r_cols[1]:
        search_q = st.text_input("Search Cases (Case ID, Username, Worker, DS-160)", placeholder="Type to search...", key="adm_search_q")
    with r_cols[2]:
        all_w = ["All"] + get_all_workers_list()
        worker_filt = st.selectbox("Filter Worker", options=all_w, key="adm_worker_filt")
    with r_cols[3]:
        excel_bytes = get_admin_gsheet_export_bytes(_get_file_mtime())
        st.download_button(
            "📥 Export All (.xlsx)",
            data=excel_bytes,
            file_name=f"HemanVisa_All_Sheets_{today_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="adm_export_all",
        )
    with r_cols[4]:
        st.markdown("<div style='height: 22px;'></div>", unsafe_allow_html=True)
        if st.button("🚪 Logout", key="adm_logout", use_container_width=True):
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            st.query_params.clear()
            st.rerun()

    if df_curr is not None and not df_curr.empty:
        view_df = reorder_dataframe_columns(df_curr.copy())
        if worker_filt != "All":
            lc_f = find_login_user_column(view_df)
            view_df = view_df[view_df[lc_f].astype(str).str.strip().str.lower() == worker_filt.lower()]
        if search_q.strip():
            mask = search_username_ds160_password_worker(view_df, search_q.strip())
            view_df = view_df[mask]
        render_clickable_table(
            view_df,
            find_display_username_column(view_df),
            find_password_column(view_df),
            table_key=f"adm_tbl_{active_s}",
            sheet_for_link=active_s,
            allow_edit=True,
        )
    else:
        st.info(f"No records found in '{active_s}'.")


def render_worker_portal(worker_name):
    notice_txt = get_admin_notice()
    if notice_txt:
        st.markdown(
            f'<div class="marquee-container"><marquee behavior="scroll" direction="left">📢 NOTICE: {notice_txt}</marquee></div>',
            unsafe_allow_html=True,
        )

    perms_map = load_permissions()
    w_perm = next((v for k, v in perms_map.items() if k.lower() == worker_name.strip().lower()), {})
    can_update = w_perm.get("can_update", w_perm.get("can_edit", True))
    can_add = w_perm.get("can_add", True)
    can_search = w_perm.get("can_search", True)
    can_download = w_perm.get("can_download", False)

    all_sheets = load_all_sheets_cached(_get_file_mtime())
    business_sheets = [s for s in all_sheets.keys() if str(s).strip() not in SYSTEM_SHEETS]
    if not business_sheets:
        business_sheets = ["Sheet1"]

    active_s = st.session_state.get("active_sheet", get_default_active_sheet(business_sheets))
    if active_s not in business_sheets:
        active_s = get_default_active_sheet(business_sheets)
        st.session_state["active_sheet"] = active_s

    df_s = load_data(sheet_name=active_s)
    my_cases_cnt = 0
    if df_s is not None and not df_s.empty:
        lc = find_login_user_column(df_s)
        my_cases_cnt = int((df_s[lc].astype(str).str.strip().str.lower() == worker_name.strip().lower()).sum())

    my_reqs = [
        r for r in load_edit_requests()
        if r.get("Requested_By", "").strip().lower() == worker_name.strip().lower()
    ]
    approved_cnt = sum(1 for r in my_reqs if r.get("Status", "").strip().lower() == "approved")
    pending_cnt = sum(1 for r in my_reqs if r.get("Status", "").strip().lower() == "pending")

    st.markdown(
        f"""
        <div class="worker-hero-banner">
            <div class="worker-hero-title">👷 CASE WORKER WORKSPACE — {worker_name}</div>
            <div class="worker-pill-group">
                <span class="worker-pill-gold">📂 Sheet: {active_s}</span>
                <span class="worker-pill-cyan">📋 My Cases: {my_cases_cnt}</span>
                <span class="worker-pill-gold">🔓 Approved Edits: {approved_cnt} | ⏳ Pending: {pending_cnt}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    t_cols = st.columns([2.2, 2.2, 1.3, 1.3, 1.3, 1.2], gap="small")
    with t_cols[0]:
        picked_ws = st.selectbox("Select Sheet", options=business_sheets, index=business_sheets.index(active_s), key="wrk_sheet_sel")
        if picked_ws != active_s:
            st.session_state["active_sheet"] = picked_ws
            st.rerun()
    with t_cols[1]:
        w_search = st.text_input(
            "Search Cases",
            placeholder="Search Case ID / Username / DS-160...",
            disabled=not can_search,
            key="wrk_search_inp",
        )
    with t_cols[2]:
        st.markdown("<div style='height: 22px;'></div>", unsafe_allow_html=True)
        if st.button("➕ Add Case", disabled=not can_add, use_container_width=True, key="wrk_add_btn"):
            st.session_state["open_dialog"] = "add_case_dlg"
            st.rerun()
    with t_cols[3]:
        st.markdown("<div style='height: 22px;'></div>", unsafe_allow_html=True)
        if can_download:
            w_bytes = get_worker_export_bytes(_get_file_mtime(), worker_name, active_s)
            st.download_button(
                "📥 Export My Cases",
                data=w_bytes,
                file_name=f"{worker_name}_Cases_{get_ist_now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="wrk_dl_btn",
            )
        else:
            st.button("🔒 Export Off", disabled=True, use_container_width=True, key="wrk_dl_dis")
    with t_cols[4]:
        st.markdown("<div style='height: 22px;'></div>", unsafe_allow_html=True)
        if st.button("🔑 Password", use_container_width=True, key="wrk_pwd_btn"):
            worker_password_dialog(worker_name)
    with t_cols[5]:
        st.markdown("<div style='height: 22px;'></div>", unsafe_allow_html=True)
        if st.button("🚪 Logout", use_container_width=True, key="wrk_logout_btn"):
            record_worker_logout(worker_name)
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            st.session_state["attendance_recorded"] = False
            st.query_params.clear()
            st.rerun()

    view_mode = st.radio(
        "Case Scope",
        options=["My Assigned Cases", "All Cases (Request Edit for Others)", "My Edit Requests"],
        horizontal=True,
        key="wrk_view_mode",
    )

    if view_mode == "My Edit Requests":
        if not my_reqs:
            st.info("You have not submitted any cross-case edit requests yet.")
        else:
            st.dataframe(pd.DataFrame(my_reqs), use_container_width=True)
        return

    if df_s is not None and not df_s.empty:
        df_prep = reorder_dataframe_columns(df_s.copy())
        lc = find_login_user_column(df_prep)
        disp_c = find_display_username_column(df_prep)
        p_c = find_password_column(df_prep)

        if view_mode == "My Assigned Cases":
            filtered_df = df_prep[df_prep[lc].astype(str).str.strip().str.lower() == worker_name.strip().lower()]
        else:
            filtered_df = df_prep

        if can_search and w_search.strip():
            mask = search_username_ds160_password_worker(filtered_df, w_search.strip())
            filtered_df = filtered_df[mask]

        if filtered_df.empty:
            st.info("No matching cases found.")
            return

        # Render table where own cases or approved cross-case edits are directly clickable,
        # and other workers' cases have a clean "Request Edit" selector below.
        render_clickable_table(
            filtered_df,
            disp_c,
            p_c,
            table_key=f"wrk_tbl_{active_s}_{view_mode}",
            sheet_for_link=active_s,
            allow_edit=can_update,
            restrict_edit_to_user=None,
        )

        # If a worker clicks a row that is NOT their own case and NOT yet approved, intercept & ask for Admin permission
        sel_r = st.session_state.get("selected_row_idx")
        if sel_r is not None and sel_r in df_prep.index and st.session_state.get("open_dialog") == "edit_case_popup":
            row_obj = df_prep.loc[sel_r]
            owner_val = str(row_obj.get(lc, "")).strip()
            cid_val = str(row_obj.get("Case_ID", "N/A")).strip()
            cuser_val = str(row_obj.get(disp_c, "N/A")).strip()
            is_own = owner_val.lower() == worker_name.strip().lower()
            is_approved_other = has_approved_cross_case_edit(worker_name, active_s, cid_val, sel_r)

            if not is_own and not is_approved_other:
                st.session_state["open_dialog"] = None
                request_cross_case_edit_dialog(
                    sheet_name=active_s,
                    row_idx=sel_r,
                    case_id=cid_val,
                    case_user=cuser_val,
                    case_owner=owner_val,
                    requester_name=worker_name,
                )


def render_login_screen():
    st.markdown(
        """
        <div class="hero-section">
            <div class="hero-title">✈️ HemanVisa Case & Appointment Portal</div>
            <div class="hero-subtitle">
                Streamlined case assignment, appointment booking status tracking, PDF management, and attendance reporting.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.session_state.get("force_logged_out_msg"):
        st.error("🔒 Your account has been blocked by Admin or session expired. Please contact Admin.")
        st.session_state["force_logged_out_msg"] = False

    _, mid_col, _ = st.columns([1, 1.4, 1])
    with mid_col:
        with st.container(border=True):
            login_role = st.selectbox("Login As", options=["Case Worker", "Admin"], key="login_role_sel")
            with st.form("portal_login_form"):
                u_inp = st.text_input("Username", placeholder="Enter username...")
                p_inp = st.text_input("Password", type="password", placeholder="Enter password...")
                sub_login = st.form_submit_button("🔐 Login", type="primary", use_container_width=True)

            if sub_login:
                u_clean = u_inp.strip()
                p_clean = p_inp.strip()
                if login_role == "Admin":
                    adm_u, adm_p = get_admin_creds()
                    if u_clean == adm_u and p_clean == adm_p:
                        st.session_state["logged_in"] = True
                        st.session_state["username"] = adm_u
                        st.session_state["role"] = "admin"
                        st.query_params["logged_in"] = "true"
                        st.query_params["username"] = adm_u
                        st.query_params["role"] = "admin"
                        st.rerun()
                    else:
                        st.error("❌ Invalid Admin credentials.")
                else:
                    if is_worker_blocked(u_clean):
                        st.error("🔒 Your login is currently blocked by Admin.")
                    else:
                        all_s = load_all_sheets_cached(_get_file_mtime())
                        matched_worker = None
                        for s_nm, df_chk in all_s.items():
                            if str(s_nm).strip() in SYSTEM_SHEETS or df_chk is None or df_chk.empty:
                                continue
                            lc = find_login_user_column(df_chk)
                            pc = find_password_column(df_chk)
                            if lc in df_chk.columns and pc in df_chk.columns:
                                hit = df_chk[
                                    (df_chk[lc].astype(str).str.strip().str.lower() == u_clean.lower())
                                    & (df_chk[pc].astype(str).str.strip() == p_clean)
                                ]
                                if not hit.empty:
                                    matched_worker = str(hit.iloc[0][lc]).strip()
                                    break
                        if matched_worker:
                            st.session_state["logged_in"] = True
                            st.session_state["username"] = matched_worker
                            st.session_state["role"] = "worker"
                            st.query_params["logged_in"] = "true"
                            st.query_params["username"] = matched_worker
                            st.query_params["role"] = "worker"
                            if not st.session_state.get("attendance_recorded"):
                                record_worker_login(matched_worker)
                                st.session_state["attendance_recorded"] = True
                            st.rerun()
                        else:
                            st.error("❌ Invalid Case Worker username or password.")


# ==============================================================================
# MAIN ENTRYPOINT & ACTIVE DIALOG DISPATCHER
# ==============================================================================
ensure_all_cases_have_unique_ids()

if not st.session_state.get("logged_in"):
    render_login_screen()
else:
    role_curr = st.session_state.get("role", "worker")
    user_curr = st.session_state.get("username", "")

    if role_curr == "admin":
        render_admin_portal()
    else:
        render_worker_portal(user_curr)

    active_dlg = st.session_state.get("open_dialog")
    if active_dlg == "add_case_dlg":
        add_case_dialog(user_curr)
    elif active_dlg == "edit_case_popup":
        sel_idx = st.session_state.get("selected_row_idx")
        sheet_target = st.session_state.get("active_sheet", get_default_active_sheet())
        perms_map = load_permissions()
        w_p = next((v for k, v in perms_map.items() if k.lower() == user_curr.strip().lower()), {})
        can_upd_flag = True if role_curr == "admin" else w_p.get("can_update", w_p.get("can_edit", True))

        if sel_idx is not None:
            df_chk_edit = load_data(sheet_name=sheet_target)
            if df_chk_edit is not None and sel_idx in df_chk_edit.index:
                if role_curr == "worker":
                    lc_e = find_login_user_column(df_chk_edit)
                    disp_e = find_display_username_column(df_chk_edit)
                    owner_e = str(df_chk_edit.loc[sel_idx].get(lc_e, "")).strip()
                    cid_e = str(df_chk_edit.loc[sel_idx].get("Case_ID", "N/A")).strip()
                    cuser_e = str(df_chk_edit.loc[sel_idx].get(disp_e, "N/A")).strip()
                    own_case = owner_e.lower() == user_curr.strip().lower()
                    approved_case = has_approved_cross_case_edit(user_curr, sheet_target, cid_e, sel_idx)

                    if not own_case and not approved_case:
                        st.session_state["open_dialog"] = None
                        request_cross_case_edit_dialog(
                            sheet_name=sheet_target,
                            row_idx=sel_idx,
                            case_id=cid_e,
                            case_user=cuser_e,
                            case_owner=owner_e,
                            requester_name=user_curr,
                        )
                    else:
                        view_and_edit_case_dialog(
                            sheet_name=sheet_target,
                            row_idx=sel_idx,
                            role=role_curr,
                            current_user=user_curr,
                            can_update=can_upd_flag,
                        )
                        if not own_case and approved_case:
                            consume_approved_cross_case_edit(user_curr, sheet_target, cid_e, sel_idx)
                else:
                    view_and_edit_case_dialog(
                        sheet_name=sheet_target,
                        row_idx=sel_idx,
                        role=role_curr,
                        current_user=user_curr,
                        can_update=True,
                    )
