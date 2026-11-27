import datetime
import io
import json
import math
import os
import re
import urllib.parse

import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components

try:
    from zoneinfo import ZoneInfo

    IST_TZ = ZoneInfo("Asia/Kolkata")
except Exception:
    IST_TZ = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

# --- Step 1: Script Location Paths, Config & Core Helpers ---
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except Exception:
    BASE_DIR = os.getcwd()


def _resolve_path(filename: str) -> str:
    if os.path.isabs(filename):
        return filename
    script_path = os.path.join(BASE_DIR, filename)
    if os.path.exists(script_path):
        return script_path
    cwd_path = os.path.abspath(filename)
    if os.path.exists(cwd_path):
        return cwd_path
    return script_path


EXCEL_FILE = _resolve_path("Heman_Report.xlsx")
ADMIN_FILE = _resolve_path("admin_config.txt")
PERMS_FILE = _resolve_path("worker_permissions.json")
NOTICE_FILE = _resolve_path("admin_notice.txt")
LOGINS_FILE = _resolve_path("worker_logins_status.json")
ATTENDANCE_FILE = _resolve_path("worker_attendance.json")
PDF_BASE_DIR = _resolve_path("uploaded_appointment_pdfs")
LOGO_FILES = [
    _resolve_path("pngTShirtPrintCricket.jpg"),
    _resolve_path("logo.jpg"),
    _resolve_path("logo.png"),
    _resolve_path("logo.jpeg"),
]

SYSTEM_SHEETS = {
    "App_Logins",
    "App_Attendance",
    "App_Permissions",
    "App_Config",
    "App_Edit_Requests",
}

EDIT_REQ_COLS = [
    "Request_ID",
    "Worker",
    "Sheet_Name",
    "Case_ID",
    "Case_Username",
    "Row_Index",
    "Reason",
    "Status",
    "Requested_At",
    "Reviewed_By",
    "Reviewed_At",
]

ATTENDANCE_COLS = ["Worker", "Date", "Login_Time", "Logout_Time", "Duration", "Status"]

os.makedirs(PDF_BASE_DIR, exist_ok=True)
st.set_page_config(page_title="HemanVisa Portal", page_icon="✈️", layout="wide")


def get_ist_now() -> datetime.datetime:
    try:
        now_aware = datetime.datetime.now(datetime.timezone.utc).astimezone(IST_TZ)
        return now_aware.replace(tzinfo=None)
    except Exception:
        return datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)


def _clean_cell_text(val) -> str:
    s = "" if val is None else str(val).strip()
    return "" if s.lower() in ["nan", "none", "null"] else s


def _get_case_field_ci(row_data, candidate_names) -> str:
    if row_data is None:
        return ""
    for name in candidate_names:
        if name in row_data.index:
            val = _clean_cell_text(row_data.get(name, ""))
            if val:
                return val
    lower_map = {str(k).strip().lower(): k for k in row_data.index}
    for name in candidate_names:
        matched_k = lower_map.get(str(name).strip().lower())
        if matched_k is not None:
            val = _clean_cell_text(row_data.get(matched_k, ""))
            if val:
                return val
    return ""


def build_cgi_formatted_message(row_data) -> str:
    u_name = _get_case_field_ci(row_data, ["USERNAME", "Username", "Case Username", "user"])
    p_word = _get_case_field_ci(
        row_data,
        [
            "Password.1",
            "PASSWORD.1",
            "password.1",
            "Password1",
            "PASSWORD1",
            "password1",
            "Case Password",
            "PASSWORD",
            "Password",
            "pass",
        ],
    )

    sq1 = _get_case_field_ci(row_data, ["SQ1", "Security Question 1", "SQ 1"])
    sqa1 = _get_case_field_ci(row_data, ["SQA", "SQA1", "Security Answer 1", "SQA 1"])

    sq2 = _get_case_field_ci(row_data, ["SQ2", "Security Question 2", "SQ 2"])
    sqa2 = _get_case_field_ci(row_data, ["SQA2", "Security Answer 2", "SQA 2"])

    sq3 = _get_case_field_ci(row_data, ["SQ3", "Security Question 3", "SQ 3"])
    sqa3 = _get_case_field_ci(row_data, ["SQA3", "Security Answer 3", "SQA 3"])

    sq_lines = []
    for q_val, a_val in [(sq1, sqa1), (sq2, sqa2), (sq3, sqa3)]:
        if q_val or a_val:
            if q_val and a_val:
                sq_lines.append(f"{q_val} - {a_val}")
            elif q_val:
                sq_lines.append(f"{q_val} -")
            else:
                sq_lines.append(f"- {a_val}")

    sq_section = "\n".join(sq_lines) if sq_lines else "No security questions found"

    return (
        f"CGI Portal Login Credentials*\n"
        f"Username: {u_name}\n"
        f"Password: {p_word}\n\n"
        f"Security questions and answers\n"
        f"{sq_section}"
    )


def render_copy_message_button(msg_text, btn_key="copy_cgi_msg_btn"):
    safe_json_text = json.dumps(msg_text)
    components.html(
        f"""
        <div style="margin: 6px 0 10px 0;">
            <button id="{btn_key}" type="button" style="
                background: linear-gradient(135deg, #059669 0%, #047857 100%);
                color: #ffffff;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
                font-size: 13px;
                font-weight: 700;
                cursor: pointer;
                box-shadow: 0 3px 8px rgba(4, 120, 87, 0.25);
            ">📋 Copy Message</button>
            <span id="{btn_key}_status" style="margin-left: 10px; font-size: 12.5px; font-weight: 700; color: #047857;"></span>
        </div>
        <script>
        (function() {{
            const btn = document.getElementById({json.dumps(btn_key)});
            const statusEl = document.getElementById({json.dumps(btn_key + "_status")});
            const textToCopy = {safe_json_text};
            if (!btn) return;
            btn.addEventListener("click", async function() {{
                try {{
                    if (navigator.clipboard && navigator.clipboard.writeText) {{
                        await navigator.clipboard.writeText(textToCopy);
                    }} else {{
                        const ta = document.createElement("textarea");
                        ta.value = textToCopy;
                        ta.style.position = "fixed";
                        ta.style.opacity = "0";
                        document.body.appendChild(ta);
                        ta.focus();
                        ta.select();
                        document.execCommand("copy");
                        document.body.removeChild(ta);
                    }}
                    if (statusEl) {{
                        statusEl.textContent = "✅ Copied!";
                        setTimeout(() => {{ statusEl.textContent = ""; }}, 2500);
                    }}
                }} catch (err) {{
                    const ta = document.createElement("textarea");
                    ta.value = textToCopy;
                    ta.style.position = "fixed";
                    ta.style.opacity = "0";
                    document.body.appendChild(ta);
                    ta.focus();
                    ta.select();
                    try {{ document.execCommand("copy"); }} catch (e) {{}}
                    document.body.removeChild(ta);
                    if (statusEl) {{
                        statusEl.textContent = "✅ Copied!";
                        setTimeout(() => {{ statusEl.textContent = ""; }}, 2500);
                    }}
                }}
            }});
        }})();
        </script>
        """,
        height=46,
    )


@st.cache_data(ttl=3600, show_spinner=False)
def _cached_webapp_url():
    try:
        url = st.secrets.get("GSHEET_WEBAPP_URL", "")
        return str(url).strip() if url else ""
    except Exception:
        return ""


def get_webapp_url():
    return _cached_webapp_url()


def _clear_sheet_runtime_caches():
    st.session_state.pop("_mem_sheets_cache", None)
    st.session_state.pop("_parsed_meta_cache", None)
    _fetch_and_process_all_sheets_cached.clear()
    get_admin_gsheet_export_bytes.clear()
    get_worker_export_bytes.clear()


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
        resp = requests.get(url, timeout=10, allow_redirects=True)
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


def _update_session_sheets_cache(updated_sheets_map):
    if not updated_sheets_map:
        return
    mem_cache = st.session_state.get("_mem_sheets_cache")
    if isinstance(mem_cache, dict):
        for s_k, df_v in updated_sheets_map.items():
            if df_v is not None:
                s_clean = str(s_k).strip()
                if s_clean in SYSTEM_SHEETS:
                    df_sys = df_v.copy()
                    df_sys.columns = df_sys.columns.str.strip()
                    mem_cache[s_clean] = df_sys.fillna("").astype(str)
                else:
                    mem_cache[s_clean] = _normalize_sheet_df(df_v.copy())
        st.session_state["_mem_sheets_cache"] = mem_cache
    st.session_state.pop("_parsed_meta_cache", None)


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

        resp = requests.post(url, json=body, timeout=10, allow_redirects=True)
        resp.raise_for_status()
        res_json = resp.json()
        if isinstance(res_json, dict) and res_json.get("status") != "success":
            st.error(f"Google Sheet Save Error ({sheet_name}): {res_json}")
        else:
            to_update = dict(all_sheets_dict) if all_sheets_dict else {str(sheet_name): df_clean}
            _update_session_sheets_cache(to_update)
    except Exception as e:
        st.error(f"Google Sheet Push Error ({sheet_name}): {e}")


def _push_multiple_sheets_batch(sheets_map):
    if not sheets_map:
        return
    url = get_webapp_url()
    if not url:
        st.error("GSHEET_WEBAPP_URL కనిపించడం లేదు!")
        return
    try:
        items = [(str(k), v.fillna("").astype(str)) for k, v in sheets_map.items() if v is not None]
        if not items:
            return
        first_name, first_df = items[0]
        rows_2d = [first_df.columns.tolist()] + first_df.values.tolist()
        all_payload = {}
        for s_k, df_k in items:
            all_payload[s_k] = [df_k.columns.tolist()] + df_k.values.tolist()
        body = {
            "action": "save_sheet",
            "sheet_name": first_name,
            "data": rows_2d,
            "sheets": all_payload,
        }
        resp = requests.post(url, json=body, timeout=8, allow_redirects=True)
        resp.raise_for_status()
        res_json = resp.json()
        if isinstance(res_json, dict) and res_json.get("status") != "success":
            st.error(f"Google Sheet Batch Save Error: {res_json}")
        else:
            _update_session_sheets_cache({k: v for k, v in items})
    except Exception as e:
        st.warning(f"Background sync slow ({e}) — session active.")


def _get_file_mtime():
    sync_ver = st.session_state.get("_gsheet_sync_version", 0)
    f_mtime = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
    return f"{f_mtime}_{sync_ver}"


def _get_runtime_meta():
    token = _get_file_mtime()
    cache = st.session_state.get("_parsed_meta_cache")
    if isinstance(cache, dict) and cache.get("_token") == token:
        return cache

    meta = {
        "_token": token,
        "perms": None,
        "notice": None,
        "admin_creds": None,
        "logins": None,
        "attendance": None,
        "edit_requests": None,
        "workers_list": None,
    }
    st.session_state["_parsed_meta_cache"] = meta
    return meta


def _read_sheet_rows(sheet_name):
    sheets = load_all_sheets_cached(_get_file_mtime())
    df = sheets.get(sheet_name) if isinstance(sheets, dict) else None
    if df is None or df.empty:
        return []
    return df.fillna("").astype(str).to_dict(orient="records")


def _build_records_df(columns, records):
    if records:
        df_new = pd.DataFrame(records)
        for col in columns:
            if col not in df_new.columns:
                df_new[col] = ""
        return df_new[columns].fillna("").astype(str)
    return pd.DataFrame(columns=columns, dtype=str)


def _write_sheet_records(sheet_name, columns, records):
    df_new = _build_records_df(columns, records)
    _push_sheet_to_webapp(sheet_name, df_new, all_sheets_dict=None)
    st.session_state["_gsheet_sync_version"] = (
        st.session_state.get("_gsheet_sync_version", 0) + 1
    )
    st.session_state.pop("_parsed_meta_cache", None)


# --- Step 2: Edit Requests, Admin Config, Permissions, Logins & Attendance ---

def load_edit_requests():
    meta = _get_runtime_meta()
    if meta["edit_requests"] is not None:
        return meta["edit_requests"]
    rows = _read_sheet_rows("App_Edit_Requests")
    cleaned = []
    for r in rows:
        w = str(r.get("Worker", "")).strip()
        cid = str(r.get("Case_ID", "")).strip()
        if w or cid:
            cleaned.append({col: str(r.get(col, "")).strip() for col in EDIT_REQ_COLS})
    meta["edit_requests"] = cleaned
    return cleaned


def save_edit_requests(req_list):
    records = []
    for item in (req_list or []):
        records.append({col: str(item.get(col, "")).strip() for col in EDIT_REQ_COLS})
    _write_sheet_records("App_Edit_Requests", EDIT_REQ_COLS, records)


def get_worker_case_edit_request(worker_name, sheet_name, case_id, case_username=""):
    if not worker_name:
        return None
    w_low = str(worker_name).strip().lower()
    s_low = str(sheet_name).strip().lower()
    cid_low = str(case_id).strip().lower()
    cu_low = str(case_username).strip().lower()

    reqs = load_edit_requests()
    matched = None
    for r in reqs:
        if r.get("Worker", "").strip().lower() != w_low:
            continue
        if r.get("Sheet_Name", "").strip().lower() != s_low:
            continue
        r_cid = r.get("Case_ID", "").strip().lower()
        r_cu = r.get("Case_Username", "").strip().lower()
        same_case = (cid_low and cid_low != "n/a" and r_cid == cid_low) or (
            cu_low and r_cu == cu_low and (not cid_low or cid_low == "n/a" or r_cid == cid_low)
        )
        if same_case:
            matched = r
    return matched


def create_or_reopen_edit_request(worker_name, sheet_name, case_id, case_username, row_idx, reason):
    now_str = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
    reqs = load_edit_requests()
    w_low = str(worker_name).strip().lower()
    s_low = str(sheet_name).strip().lower()
    cid_low = str(case_id).strip().lower()

    for r in reqs:
        if (
            r.get("Worker", "").strip().lower() == w_low
            and r.get("Sheet_Name", "").strip().lower() == s_low
            and r.get("Case_ID", "").strip().lower() == cid_low
        ):
            r["Row_Index"] = str(row_idx)
            r["Case_Username"] = str(case_username).strip()
            r["Reason"] = str(reason).strip()
            r["Status"] = "Pending"
            r["Requested_At"] = now_str
            r["Reviewed_By"] = ""
            r["Reviewed_At"] = ""
            save_edit_requests(reqs)
            return r

    new_req = {
        "Request_ID": f"ER{int(get_ist_now().timestamp())}{len(reqs) + 1}",
        "Worker": str(worker_name).strip(),
        "Sheet_Name": str(sheet_name).strip(),
        "Case_ID": str(case_id).strip(),
        "Case_Username": str(case_username).strip(),
        "Row_Index": str(row_idx),
        "Reason": str(reason).strip(),
        "Status": "Pending",
        "Requested_At": now_str,
        "Reviewed_By": "",
        "Reviewed_At": "",
    }
    reqs.append(new_req)
    save_edit_requests(reqs)
    return new_req


def get_admin_creds():
    meta = _get_runtime_meta()
    if meta["admin_creds"] is not None:
        return meta["admin_creds"]
    rows = _read_sheet_rows("App_Config")
    for r in rows:
        if r.get("Key", "").strip() == "admin_creds":
            u = r.get("Value1", "").strip() or "admin"
            p = r.get("Value2", "").strip() or "admin123"
            meta["admin_creds"] = (u, p)
            return u, p
    meta["admin_creds"] = ("admin", "admin123")
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
    meta = _get_runtime_meta()
    if meta["notice"] is not None:
        return meta["notice"]
    rows = _read_sheet_rows("App_Config")
    for r in rows:
        if r.get("Key", "").strip() == "admin_notice":
            val = r.get("Value1", "").strip()
            meta["notice"] = val
            return val
    meta["notice"] = ""
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
    meta = _get_runtime_meta()
    if meta["perms"] is not None:
        return meta["perms"]
    rows = _read_sheet_rows("App_Permissions")
    perms = {}
    for r in rows:
        w = r.get("Worker", "").strip()
        if w:
            perms[w] = {
                "can_edit": str(r.get("can_edit", "True")).strip().lower() in ["true", "1", "yes"],
                "can_delete_req": str(r.get("can_delete_req", "True")).strip().lower() in ["true", "1", "yes"],
                "can_reschedule": str(r.get("can_reschedule", "True")).strip().lower() in ["true", "1", "yes"],
                "can_add": str(r.get("can_add", "True")).strip().lower() in ["true", "1", "yes"],
                "can_update": str(r.get("can_update", r.get("can_edit", "True"))).strip().lower() in ["true", "1", "yes"],
                "can_search": str(r.get("can_search", "True")).strip().lower() in ["true", "1", "yes"],
                "can_download": str(r.get("can_download", "False")).strip().lower() in ["true", "1", "yes"],
                "can_sync": str(r.get("can_sync", "False")).strip().lower() in ["true", "1", "yes"],
                "block_login": str(r.get("block_login", "False")).strip().lower() in ["true", "1", "yes"],
            }
    meta["perms"] = perms
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
    meta = _get_runtime_meta()
    if meta["logins"] is not None:
        return meta["logins"]
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
    meta["logins"] = logins
    return logins


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
    meta = _get_runtime_meta()
    if meta["attendance"] is not None:
        return meta["attendance"]
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
    meta["attendance"] = cleaned
    return cleaned


def _build_login_tracking_df(logins_dict):
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
    return _build_records_df(cols, records)


def _build_attendance_df(att_list):
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
    return _build_records_df(ATTENDANCE_COLS, records)


def record_worker_login(worker_name):
    if not worker_name:
        return
    now_dt = get_ist_now()
    today_str = now_dt.strftime("%Y-%m-%d")
    time_str = now_dt.strftime("%I:%M:%S %p")
    full_timestamp = f"{today_str} {time_str}"

    session_guard = f"_login_sync_done_{worker_name.strip().lower()}_{today_str}"
    if st.session_state.get(session_guard):
        return
    st.session_state[session_guard] = True

    logins = dict(load_login_tracking())
    logins[worker_name.strip()] = {
        "status": "Logged In",
        "last_login": full_timestamp,
    }
    df_logins = _build_login_tracking_df(logins)

    attendance = list(load_attendance())
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
    df_att = _build_attendance_df(attendance)

    _update_session_sheets_cache({"App_Logins": df_logins, "App_Attendance": df_att})
    _push_multiple_sheets_batch({
        "App_Logins": df_logins,
        "App_Attendance": df_att,
    })
    st.session_state["_gsheet_sync_version"] = (
        st.session_state.get("_gsheet_sync_version", 0) + 1
    )
    st.session_state.pop("_parsed_meta_cache", None)


def record_worker_logout(worker_name):
    if not worker_name:
        return
    now_dt = get_ist_now()
    today_str = now_dt.strftime("%Y-%m-%d")
    time_str = now_dt.strftime("%I:%M:%S %p")

    logins = dict(load_login_tracking())
    for k in list(logins.keys()):
        if k.strip().lower() == worker_name.strip().lower():
            logins[k]["status"] = "Not Logged In"
    df_logins = _build_login_tracking_df(logins)

    attendance = list(load_attendance())
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
    df_att = _build_attendance_df(attendance)

    _update_session_sheets_cache({"App_Logins": df_logins, "App_Attendance": df_att})
    _push_multiple_sheets_batch({
        "App_Logins": df_logins,
        "App_Attendance": df_att,
    })
    st.session_state["_gsheet_sync_version"] = (
        st.session_state.get("_gsheet_sync_version", 0) + 1
    )
    st.session_state.pop("_parsed_meta_cache", None)


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
        st_val = str(row.get("Status", "")).strip() or ("Logged Out" if lo_t != "-" else "Active")
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
                "Current_Status": st_val,
            }
        wd = worker_date_map[wd_key]
        wd["Sessions"] += 1
        if l_t != "-" and wd["First_Login"] is None:
            wd["First_Login"] = l_t
        if lo_t != "-":
            wd["Last_Logout"] = lo_t
        wd["Total_Seconds"] += sec
        wd["Current_Status"] = st_val

        if w not in worker_overall_map:
            worker_overall_map[w] = {
                "Worker": w,
                "Active_Days": set(),
                "Total_Sessions": 0,
                "Last_Login": l_t if l_t != "-" else "-",
                "Last_Logout": lo_t if lo_t != "-" else "-",
                "Total_Seconds": 0,
                "Current_Status": st_val,
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
        wo["Current_Status"] = st_val

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
