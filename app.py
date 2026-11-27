# ================= STEP 1: IMPORTS, CONSTANTS, PAGE CONFIG & FORMATTING/CLIPBOARD HELPERS =================
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

SYSTEM_SHEETS = {
    "App_Logins",
    "App_Attendance",
    "App_Permissions",
    "App_Config",
    "App_Edit_Requests",
}

if not os.path.exists(PDF_BASE_DIR):
    os.makedirs(PDF_BASE_DIR)

st.set_page_config(page_title="HemanVisa Portal", page_icon="✈️", layout="wide")


def _clean_cell_text(val):
    s = "" if val is None else str(val).strip()
    return "" if s.lower() in ["nan", "none", "null"] else s


def _get_case_field_ci(row_data, candidate_names):
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


def build_cgi_formatted_message(row_data):
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
# ================= END OF STEP 1 =================
# ================= STEP 2: GOOGLE SHEETS WEBAPP SYNC & CACHE LAYER =================
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
# ================= END OF STEP 2 =================
# ================= STEP 3: IST TIME, RUNTIME METADATA, EDIT REQUESTS, ADMIN CONFIG & PERMISSIONS =================
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
# ================= END OF STEP 3 =================
# ================= STEP 4: LOGIN TRACKING & ATTENDANCE / HOURS ENGINE =================
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
# ================= END OF STEP 4 =================
# ================= STEP 5: CASE ID GENERATION, RESCHEDULE & DATA LOAD/SAVE =================
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


@st.cache_data(ttl=120, show_spinner=False)
def _fetch_and_process_all_sheets_cached(base_cache_key):
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


def load_all_sheets_cached(file_mtime):
    mem_cache = st.session_state.get("_mem_sheets_cache")
    if isinstance(mem_cache, dict) and len(mem_cache) > 0:
        return mem_cache
    base_key = os.path.getmtime(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
    processed = _fetch_and_process_all_sheets_cached(base_key)
    if isinstance(processed, dict) and len(processed) > 0:
        st.session_state["_mem_sheets_cache"] = {k: v.copy() for k, v in processed.items()}
        return st.session_state["_mem_sheets_cache"]
    return processed


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
            _clear_sheet_runtime_caches()
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

    target_sheet_key = str(sheet_name)
    if isinstance(sheet_name, int) and sheets_data:
        keys = list(sheets_data.keys())
        if 0 <= sheet_name < len(keys):
            target_sheet_key = str(keys[sheet_name])

    df_to_save = reorder_dataframe_columns(df_to_save)
    sheets_data[target_sheet_key] = df_to_save

    _update_session_sheets_cache({target_sheet_key: df_to_save})
    _push_sheet_to_webapp(target_sheet_key, df_to_save, sheets_data)
    st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
    st.session_state.pop("_parsed_meta_cache", None)
    st.session_state["_last_id_sync_mtime"] = _get_file_mtime()
# ================= END OF STEP 5 =================
# ================= STEP 6: WORKER DISCOVERY, COLUMN HELPERS, PDF STORAGE & STYLING =================
def get_all_known_workers():
    meta = _get_runtime_meta()
    if meta["workers_list"] is not None:
        return meta["workers_list"]

    workers = set()
    perms = load_permissions()
    for w in perms.keys():
        if str(w).strip():
            workers.add(str(w).strip())

    logins = load_login_tracking()
    for w in logins.keys():
        if str(w).strip():
            workers.add(str(w).strip())

    att = load_attendance()
    for r in att:
        w = str(r.get("Worker", "")).strip()
        if w:
            workers.add(w)

    all_sheets = load_all_sheets_cached(_get_file_mtime())
    for s_name, df in all_sheets.items():
        if str(s_name).strip() in SYSTEM_SHEETS:
            continue
        if df is not None and not df.empty:
            login_c = find_login_user_column(df)
            if login_c in df.columns:
                for val in df[login_c].astype(str):
                    v = val.strip()
                    if _is_valid_case_username(v):
                        workers.add(v)

    result = sorted(list(workers), key=lambda x: x.lower())
    meta["workers_list"] = result
    return result


def get_worker_permissions(worker_name):
    perms = load_permissions()
    matched_key = None
    for k in perms.keys():
        if k.strip().lower() == str(worker_name).strip().lower():
            matched_key = k
            break
    if matched_key is None:
        return {
            "can_edit": True,
            "can_delete_req": True,
            "can_reschedule": True,
            "can_add": True,
            "can_update": True,
            "can_search": True,
            "can_download": False,
            "can_sync": False,
            "block_login": False,
        }
    p = perms[matched_key]
    return {
        "can_edit": bool(p.get("can_edit", True)),
        "can_delete_req": bool(p.get("can_delete_req", True)),
        "can_reschedule": bool(p.get("can_reschedule", True)),
        "can_add": bool(p.get("can_add", True)),
        "can_update": bool(p.get("can_update", p.get("can_edit", True))),
        "can_search": bool(p.get("can_search", True)),
        "can_download": bool(p.get("can_download", False)),
        "can_sync": bool(p.get("can_sync", False)),
        "block_login": bool(p.get("block_login", False)),
    }


def find_login_user_column(df):
    for col in df.columns:
        if col.strip().lower() in [
            "username",
            "case worker",
            "cases_worker",
            "user",
        ]:
            if col.strip() != "USERNAME":
                return col
    for col in df.columns:
        c_low = col.strip().lower().replace("_", " ")
        if "case" in c_low and "user" in c_low:
            return col
    return "Username" if "Username" in df.columns else df.columns[0]


def find_lead_column(df):
    for col in df.columns:
        if col.strip().lower() in ["lead", "team lead", "lead name"]:
            return col
    return None


def find_no_of_app_column(df):
    for col in df.columns:
        c_low = col.strip().lower().replace("_", " ")
        if c_low in ["no of app", "no_of_app", "no of applications", "apps"]:
            return col
    return None


def get_row_display_title(row, df_cols):
    parts = []
    if "Case_ID" in df_cols and str(row.get("Case_ID", "")).strip():
        parts.append(f"Case ID: {str(row.get('Case_ID', '')).strip()}")

    for c in df_cols:
        c_low = c.strip().lower().replace("_", " ")
        if (
            c_low in ["case", "case name", "applicant", "applicant name", "client"]
            or "case" in c_low
        ) and "user" not in c_low and "id" not in c_low:
            val = str(row.get(c, "")).strip()
            if val and val.lower() != "nan":
                parts.append(val)
                break

    disp_u = _detect_case_username_col(pd.DataFrame(columns=df_cols))
    if disp_u and str(row.get(disp_u, "")).strip():
        u_val = str(row.get(disp_u, "")).strip()
        if u_val.lower() != "nan":
            parts.append(f"User: {u_val}")

    if not parts:
        for c in df_cols:
            if c not in [
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
            ]:
                val = str(row.get(c, "")).strip()
                if val and val.lower() != "nan":
                    parts.append(f"{c}: {val}")
                    break

    return " | ".join(parts) if parts else "Case Record"


def sanitize_filename_part(text):
    clean = re.sub(r"[^A-Za-z0-9._-]+", "_", str(text).strip())
    return clean.strip("_") or "case"


def get_pdf_storage_path(sheet_name, case_id, username):
    sheet_folder = os.path.join(PDF_BASE_DIR, sanitize_filename_part(sheet_name))
    if not os.path.exists(sheet_folder):
        os.makedirs(sheet_folder)
    cid_part = sanitize_filename_part(case_id if case_id else "NOID")
    user_part = sanitize_filename_part(username if username else "NOUSER")
    filename = f"{cid_part}_{user_part}_Appointment.pdf"
    return os.path.join(sheet_folder, filename)


def load_logo_base64():
    for logo_path in LOGO_FILES:
        if os.path.exists(logo_path):
            try:
                with open(logo_path, "rb") as f:
                    data = base64.b64encode(f.read()).decode("utf-8")
                ext = os.path.splitext(logo_path)[1].lower().replace(".", "")
                mime = "image/jpeg" if ext in ["jpg", "jpeg"] else f"image/{ext}"
                return f"data:{mime};base64,{data}"
            except Exception:
                continue
    return None


def inject_custom_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        }

        .stApp {
            background: linear-gradient(135deg, #f0f4f8 0%, #e6eef8 50%, #f4f7fb 100%);
        }

        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        .block-container {
            padding-top: 1.2rem !important;
            padding-bottom: 2rem !important;
            max-width: 98% !important;
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
            border-right: 1px solid rgba(255, 255, 255, 0.08);
        }

        [data-testid="stSidebar"] * {
            color: #f8fafc !important;
        }

        [data-testid="stSidebar"] .stButton > button {
            background: rgba(255, 255, 255, 0.08);
            color: #ffffff !important;
            border: 1px solid rgba(255, 255, 255, 0.16);
            border-radius: 10px;
            font-weight: 600;
            transition: all 0.2s ease;
        }

        [data-testid="stSidebar"] .stButton > button:hover {
            background: rgba(59, 130, 246, 0.35);
            border-color: #60a5fa;
            transform: translateY(-1px);
        }

        [data-testid="stSidebar"] .stDownloadButton > button {
            background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 10px;
            font-weight: 700;
        }

        .top-header-card {
            background: linear-gradient(120deg, #0f172a 0%, #1d4ed8 55%, #0284c7 100%);
            border-radius: 18px;
            padding: 20px 26px;
            color: #ffffff;
            box-shadow: 0 12px 28px rgba(15, 23, 42, 0.16);
            margin-bottom: 16px;
            border: 1px solid rgba(255, 255, 255, 0.15);
        }

        .top-header-title {
            font-size: 26px;
            font-weight: 800;
            letter-spacing: -0.02em;
            margin: 0;
            line-height: 1.2;
            color: #ffffff !important;
        }

        .top-header-sub {
            font-size: 13.5px;
            color: #dbeafe !important;
            margin-top: 5px;
            font-weight: 500;
        }

        .header-badge {
            display: inline-block;
            background: rgba(255, 255, 255, 0.16);
            backdrop-filter: blur(6px);
            border: 1px solid rgba(255, 255, 255, 0.25);
            padding: 6px 14px;
            border-radius: 999px;
            font-size: 12.5px;
            font-weight: 700;
            color: #ffffff !important;
        }

        .notice-marquee-box {
            background: linear-gradient(90deg, #fffbeb 0%, #fef3c7 50%, #fffbeb 100%);
            border: 1px solid #f59e0b;
            border-left: 6px solid #d97706;
            border-radius: 12px;
            padding: 10px 16px;
            margin-bottom: 16px;
            box-shadow: 0 4px 12px rgba(217, 119, 6, 0.08);
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .notice-tag {
            background: #d97706;
            color: #ffffff;
            font-size: 11px;
            font-weight: 800;
            padding: 3px 9px;
            border-radius: 999px;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            white-space: nowrap;
        }

        .kpi-card {
            background: #ffffff;
            border-radius: 16px;
            padding: 16px 18px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
            height: 100%;
        }

        .kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.09);
        }

        .kpi-title {
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #64748b;
            margin-bottom: 6px;
        }

        .kpi-value {
            font-size: 26px;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.15;
        }

        .kpi-sub {
            font-size: 12px;
            color: #475569;
            margin-top: 6px;
            font-weight: 500;
        }

        .kpi-accent-blue { border-top: 4px solid #2563eb; }
        .kpi-accent-green { border-top: 4px solid #059669; }
        .kpi-accent-amber { border-top: 4px solid #d97706; }
        .kpi-accent-purple { border-top: 4px solid #7c3aed; }
        .kpi-accent-red { border-top: 4px solid #dc2626; }

        .section-card {
            background: #ffffff;
            border-radius: 16px;
            padding: 18px 20px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
            margin-bottom: 16px;
        }

        .section-heading {
            font-size: 17px;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 4px;
        }

        .section-caption {
            font-size: 12.5px;
            color: #64748b;
            margin-bottom: 12px;
        }

        .login-hero-box {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 60%, #0369a1 100%);
            border-radius: 20px;
            padding: 28px;
            color: #ffffff;
            box-shadow: 0 16px 36px rgba(15, 23, 42, 0.18);
            margin-bottom: 18px;
        }

        .stButton > button {
            border-radius: 10px;
            font-weight: 700;
            transition: all 0.18s ease;
        }

        .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
            border: none;
            color: #ffffff !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.28);
        }

        .stButton > button[kind="primary"]:hover {
            background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%);
            transform: translateY(-1px);
        }

        [data-testid="stDataFrame"] {
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid #e2e8f0;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #f1f5f9;
            padding: 6px;
            border-radius: 12px;
        }

        .stTabs [data-baseweb="tab"] {
            border-radius: 8px;
            font-weight: 700;
            font-size: 13px;
            padding: 8px 14px;
        }

        .stTabs [aria-selected="true"] {
            background-color: #ffffff !important;
            color: #1d4ed8 !important;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.08);
        }

        .clickable-case-card {
            background: #ffffff;
            border: 1px solid #dbeafe;
            border-left: 5px solid #2563eb;
            border-radius: 12px;
            padding: 12px 14px;
            margin-bottom: 10px;
            box-shadow: 0 3px 10px rgba(15, 23, 42, 0.04);
        }

        .clickable-case-card:hover {
            border-color: #2563eb;
            box-shadow: 0 6px 16px rgba(37, 99, 235, 0.12);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_top_header(role_label, user_label, active_sheet_label):
    notice_text = get_admin_notice()
    logo_b64 = load_logo_base64()
    logo_html = (
        f'<img src="{logo_b64}" alt="HemanVisa Logo" style="width:52px;height:52px;object-fit:cover;border-radius:12px;border:2px solid rgba(255,255,255,0.35);background:#ffffff;padding:2px;flex-shrink:0;" />'
        if logo_b64
        else '<div style="width:52px;height:52px;border-radius:12px;background:rgba(255,255,255,0.18);display:flex;align-items:center;justify-content:center;font-size:26px;border:1px solid rgba(255,255,255,0.3);flex-shrink:0;">✈️</div>'
    )
    st.markdown(
        f"""
        <div class="top-header-card">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
                <div style="display:flex; align-items:center; gap:14px;">
                    {logo_html}
                    <div>
                        <div class="top-header-title">HemanVisa Case & Appointment Portal</div>
                        <div class="top-header-sub">Live Case Tracking • Appointment Booking & Rescheduling • Google Sheets Sync</div>
                    </div>
                </div>
                <div style="display:flex; gap:8px; flex-wrap:wrap; align-items:center;">
                    <span class="header-badge">👤 {user_label}</span>
                    <span class="header-badge">🛡️ {role_label}</span>
                    <span class="header-badge">📄 {active_sheet_label}</span>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if notice_text:
        st.markdown(
            f"""
            <div class="notice-marquee-box">
                <span class="notice-tag">📢 Admin Notice</span>
                <marquee behavior="scroll" direction="left" scrollamount="5" style="font-weight:600; color:#78350f; font-size:13.5px;">
                    {notice_text}
                </marquee>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_kpi_row(cards):
    cols = st.columns(len(cards))
    for col, card in zip(cols, cards):
        accent = card.get("accent", "kpi-accent-blue")
        title = card.get("title", "")
        value = card.get("value", "0")
        sub = card.get("sub", "")
        with col:
            st.markdown(
                f"""
                <div class="kpi-card {accent}">
                    <div class="kpi-title">{title}</div>
                    <div class="kpi-value">{value}</div>
                    <div class="kpi-sub">{sub}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
# ================= END OF STEP 6 =================
# ================= STEP 7: SESSION DEFAULTS & LOGIN SCREEN (ADMIN & CASE WORKER) =================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_role" not in st.session_state:
    st.session_state.user_role = None
if "current_user" not in st.session_state:
    st.session_state.current_user = None
if "active_sheet" not in st.session_state:
    st.session_state.active_sheet = None
if "is_editing_popup" not in st.session_state:
    st.session_state.is_editing_popup = False
if "is_delete_request_mode" not in st.session_state:
    st.session_state.is_delete_request_mode = False
if "is_reschedule_mode" not in st.session_state:
    st.session_state.is_reschedule_mode = False
if "selected_row_idx" not in st.session_state:
    st.session_state.selected_row_idx = None
if "open_dialog" not in st.session_state:
    st.session_state.open_dialog = None
if "bulk_action_type" not in st.session_state:
    st.session_state.bulk_action_type = None
if "bulk_action_rows" not in st.session_state:
    st.session_state.bulk_action_rows = []
if "bulk_action_sheet" not in st.session_state:
    st.session_state.bulk_action_sheet = None
if "show_bulk_confirm" not in st.session_state:
    st.session_state.show_bulk_confirm = False

inject_custom_css()


def show_login_page():
    logo_b64 = load_logo_base64()
    hero_logo_html = (
        f'<img src="{logo_b64}" alt="HemanVisa Logo" style="width:68px;height:68px;object-fit:cover;border-radius:16px;border:2px solid rgba(255,255,255,0.4);background:#ffffff;padding:3px;flex-shrink:0;" />'
        if logo_b64
        else '<div style="width:68px;height:68px;border-radius:16px;background:rgba(255,255,255,0.18);display:flex;align-items:center;justify-content:center;font-size:34px;border:1px solid rgba(255,255,255,0.3);flex-shrink:0;">✈️</div>'
    )
    st.markdown(
        f"""
        <div class="login-hero-box">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
                <div style="display:flex; align-items:center; gap:16px;">
                    {hero_logo_html}
                    <div>
                        <div style="font-size:28px; font-weight:800; letter-spacing:-0.02em; color:#ffffff;">HemanVisa Case & Appointment Portal</div>
                        <div style="font-size:14px; color:#dbeafe; margin-top:6px;">
                            Sign in with your assigned Case Worker Username or Admin credentials to manage cases, appointments & reschedules.
                        </div>
                    </div>
                </div>
                <div style="background:rgba(255,255,255,0.14); padding:8px 16px; border-radius:999px; font-size:12.5px; font-weight:700; color:#ffffff;">
                    🔒 Secure Portal Access
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _, center_col, _ = st.columns([1, 1.3, 1])
    with center_col:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        login_tab_worker, login_tab_admin = st.tabs(
            ["👷 Case Worker Login", "🛡️ Admin Login"]
        )

        with login_tab_worker:
            st.markdown(
                '<div class="section-heading">Case Worker Sign In</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                '<div class="section-caption">Enter your Case Worker Username to view and update your assigned cases.</div>',
                unsafe_allow_html=True,
            )
            with st.form("worker_login_form", clear_on_submit=False):
                worker_username_input = st.text_input(
                    "Username",
                    placeholder="e.g., Keerthana, Sai, Kalyan, Anil",
                    key="worker_login_user_input",
                )
                worker_submit = st.form_submit_button(
                    "🚀 Login as Case Worker",
                    type="primary",
                    use_container_width=True,
                )

            if worker_submit:
                entered_user = worker_username_input.strip()
                if not entered_user:
                    st.warning("Please enter your Username.")
                else:
                    perms_map = load_permissions()
                    blocked_users = [
                        k.lower()
                        for k, v in perms_map.items()
                        if bool(v.get("block_login", False))
                    ]
                    if entered_user.lower() in blocked_users:
                        st.error(
                            "🚫 Your login access has been blocked by Admin. Please contact Admin."
                        )
                    else:
                        ensure_all_cases_have_unique_ids(force=False)
                        all_sheets = load_all_sheets_cached(_get_file_mtime())
                        valid_users = set()
                        for s_name, df_s in all_sheets.items():
                            if str(s_name).strip() in SYSTEM_SHEETS:
                                continue
                            if df_s is not None and not df_s.empty:
                                login_col = find_login_user_column(df_s)
                                if login_col in df_s.columns:
                                    for val in df_s[login_col].astype(str):
                                        v = val.strip()
                                        if v and v.lower() != "nan":
                                            valid_users.add(v.lower())
                        for k in perms_map.keys():
                            if k.strip():
                                valid_users.add(k.strip().lower())

                        if entered_user.lower() in valid_users:
                            st.session_state.logged_in = True
                            st.session_state.user_role = "worker"
                            st.session_state.current_user = entered_user
                            sheets_avail = get_sheet_names()
                            st.session_state.active_sheet = get_default_active_sheet(
                                sheets_avail
                            )
                            record_worker_login(entered_user)
                            st.rerun()
                        else:
                            st.error(
                                "❌ Invalid Case Worker Username! Please check your spelling or contact Admin."
                            )

        with login_tab_admin:
            st.markdown(
                '<div class="section-heading">Administrator Sign In</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                '<div class="section-caption">Full access to all sheets, team leads, worker permissions, attendance & edit approvals.</div>',
                unsafe_allow_html=True,
            )
            admin_u, admin_p = get_admin_creds()
            with st.form("admin_login_form", clear_on_submit=False):
                admin_user_input = st.text_input(
                    "Admin Username",
                    placeholder="Enter admin username",
                    key="admin_login_user_input",
                )
                admin_pass_input = st.text_input(
                    "Admin Password",
                    type="password",
                    placeholder="Enter admin password",
                    key="admin_login_pass_input",
                )
                admin_submit = st.form_submit_button(
                    "🔐 Login as Admin",
                    type="primary",
                    use_container_width=True,
                )

            if admin_submit:
                if (
                    admin_user_input.strip() == admin_u
                    and admin_pass_input.strip() == admin_p
                ):
                    st.session_state.logged_in = True
                    st.session_state.user_role = "admin"
                    st.session_state.current_user = admin_u
                    sheets_avail = get_sheet_names()
                    st.session_state.active_sheet = get_default_active_sheet(
                        sheets_avail
                    )
                    st.rerun()
                else:
                    st.error("❌ Invalid Admin Username or Password!")

        st.markdown("</div>", unsafe_allow_html=True)


if not st.session_state.logged_in:
    show_login_page()
    st.stop()
# ================= END OF STEP 7 =================
# ================= STEP 8: COLUMN DETECTION, SEARCH & FILTERING HELPERS =================
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
    meta = _get_runtime_meta()
    if meta["workers_list"] is not None:
        return meta["workers_list"]
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
    res = sorted(list(all_workers))
    meta["workers_list"] = res
    return res


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
    current_perms = dict(load_permissions())
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
# ================= END OF STEP 8 =================
# ================= STEP 9: STICKY TOP BANNER, CLICKABLE TABLE & CASE MODAL =================
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


def render_sticky_top_banner(qp):
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
    qp,
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
    row_case_ids = (
        page_df["Case_ID"].astype(str).str.strip().tolist()
        if "Case_ID" in page_df.columns
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

    all_edit_reqs = (
        load_edit_requests()
        if (curr_role == "worker" and allow_edit)
        else []
    )
    worker_edit_req_lookup = {}
    if all_edit_reqs and curr_role == "worker":
        u_low_k = curr_user.strip().lower()
        for er in all_edit_reqs:
            if er.get("Worker", "").strip().lower() == u_low_k:
                s_k = er.get("Sheet_Name", "").strip().lower()
                cid_k = er.get("Case_ID", "").strip().lower()
                cu_k = er.get("Case_Username", "").strip().lower()
                worker_edit_req_lookup[(s_k, cid_k, cu_k)] = er

    for i, idx in enumerate(indices):
        serial_no = start_idx + i + 1
        act_t = act_col_vals[i]
        b_status = booking_status_vals[i]
        target_row_id = orig_indices[i]
        target_sheet_name = row_sheets[i] if row_sheets[i] else curr_sheet
        q_sheet_row = urllib.parse.quote(str(target_sheet_name))
        row_link = f"?logged_in={q_logged}&username={q_user}&role={q_role}&active_sheet={q_sheet_row}&edit_row={target_row_id}"

        row_can_edit = allow_edit
        is_other_worker_case = False
        req_badge = ""

        if row_can_edit and curr_role == "worker":
            r_worker = row_workers[i].strip().lower()
            r_cid = row_case_ids[i].strip().lower()
            r_cu = col_data[display_c][i].strip().lower() if display_c in col_data else ""
            if r_worker != curr_user.strip().lower():
                is_other_worker_case = True
                matched_req = worker_edit_req_lookup.get(
                    (target_sheet_name.strip().lower(), r_cid, r_cu)
                )
                if matched_req is None:
                    for (s_k, cid_k, cu_k), er_val in worker_edit_req_lookup.items():
                        if s_k == target_sheet_name.strip().lower() and (
                            (r_cid and cid_k == r_cid) or (not r_cid and r_cu and cu_k == r_cu)
                        ):
                            matched_req = er_val
                            break
                if matched_req and matched_req.get("Status", "").strip().lower() == "approved":
                    req_badge = ' <span style="font-size:10.5px; color:#047857; background:#d1fae5; padding:1px 6px; border-radius:5px; font-weight:700;">Approved Edit</span>'
                elif matched_req and matched_req.get("Status", "").strip().lower() == "pending":
                    req_badge = ' <span style="font-size:10.5px; color:#b45309; background:#fef3c7; padding:1px 6px; border-radius:5px; font-weight:700;">Edit Pending</span>'
                elif matched_req and matched_req.get("Status", "").strip().lower() == "rejected":
                    req_badge = ' <span style="font-size:10.5px; color:#b91c1c; background:#fee2e2; padding:1px 6px; border-radius:5px; font-weight:700;">Edit Rejected</span>'
                else:
                    req_badge = ' <span style="font-size:10.5px; color:#4338ca; background:#e0e7ff; padding:1px 6px; border-radius:5px; font-weight:700;">Request Edit</span>'

        if row_can_edit and restrict_edit_to_user and not is_other_worker_case:
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
                    cell_content = f'<a href="{row_link}" target="_self" class="username-hyperlink">{display_val}</a>{req_badge}'
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
    case_user_val = str(row_data.get(display_c, "")).strip()
    case_owner_val = str(row_data.get(login_c, "")).strip()

    is_foreign_case_for_worker = (
        role == "worker"
        and case_owner_val.lower() != str(current_user).strip().lower()
    )
    edit_req_rec = (
        get_worker_case_edit_request(current_user, sheet_name, case_id_val, case_user_val)
        if is_foreign_case_for_worker
        else None
    )
    edit_req_status = (
        str(edit_req_rec.get("Status", "")).strip().lower()
        if edit_req_rec
        else ""
    )
    foreign_case_approved = is_foreign_case_for_worker and (edit_req_status == "approved")

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

        st.markdown("#### ✉️ Formatted Message (Copy or edit before sending):")
        formatted_cgi_msg = build_cgi_formatted_message(row_data)
        st.text_area(
            "Formatted Message",
            value=formatted_cgi_msg,
            height=195,
            label_visibility="collapsed",
            key=f"cgi_formatted_msg_{sheet_name}_{row_idx}",
        )
        render_copy_message_button(
            formatted_cgi_msg,
            btn_key=f"copy_cgi_btn_{sheet_name}_{row_idx}",
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
        elif is_foreign_case_for_worker and not foreign_case_approved:
            st.warning(
                f"🔒 This case belongs to **{case_owner_val if case_owner_val else 'another / unassigned worker'}**. To edit this case, send an Edit Request to Admin first; you can edit only after Admin accepts it."
            )
            if edit_req_status == "pending":
                st.info(
                    f"⏳ Your Edit Request for Case **{case_id_val}** is currently **Pending** Admin approval (Requested at: `{edit_req_rec.get('Requested_At', '')}`)."
                )
            elif edit_req_status == "rejected":
                st.error(
                    f"❌ Your previous Edit Request for Case **{case_id_val}** was **Rejected** by Admin. You can submit a fresh request below."
                )

            with st.form(f"worker_edit_req_form_{sheet_name}_{row_idx}"):
                edit_reason_in = st.text_area(
                    "✍️ Why do you need to edit this case? (Required)",
                    value=str(edit_req_rec.get("Reason", "")).strip() if edit_req_rec else "",
                    placeholder="Enter reason for editing this case...",
                )
                sub_edit_req = st.form_submit_button(
                    "📤 Send Edit Request to Admin",
                    type="primary",
                    use_container_width=True,
                )
                if sub_edit_req:
                    if not edit_reason_in.strip():
                        st.warning("⚠️ Please enter a reason before sending the edit request.")
                    else:
                        create_or_reopen_edit_request(
                            worker_name=current_user,
                            sheet_name=sheet_name,
                            case_id=case_id_val,
                            case_username=case_user_val,
                            row_idx=row_idx,
                            reason=edit_reason_in,
                        )
                        st.success("✅ Edit Request sent to Admin! Wait for Admin approval to edit.")
                        st.rerun()
        else:
            if foreign_case_approved:
                st.success("✅ Admin has **Approved** your Edit Request for this case. You can now edit and save it.")
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
                    if role == "worker" and is_foreign_case_for_worker and not foreign_case_approved:
                        st.error("❌ Edit not allowed! Admin has not approved your Edit Request for this case.")
                    elif booking_status_choice == "Rescheduled" and (
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
# ================= END OF STEP 9 =================
# ================= STEP 10: ADMIN & WORKER DIALOGS =================
@st.dialog("📢 Manage Notice Board", width="medium")
def admin_notice_dialog():
    curr_notice = get_admin_notice()
    with st.form("admin_notice_modal_form"):
        new_notice = st.text_area(
            "Scrolling Notice / Announcement Text (Leave empty to clear)",
            value=curr_notice,
            height=90,
        )
        c1, c2 = st.columns(2)
        with c1:
            if st.form_submit_button("💾 Save Notice", type="primary", use_container_width=True):
                save_admin_notice(new_notice)
                st.session_state["open_dialog"] = None
                st.rerun()
        with c2:
            if st.form_submit_button("❌ Close", use_container_width=True):
                st.session_state["open_dialog"] = None
                st.rerun()


@st.dialog("🔐 Change Admin Credentials", width="medium")
def admin_creds_dialog():
    admin_u, admin_p = get_admin_creds()
    with st.form("admin_creds_modal_form"):
        new_u = st.text_input("New Admin Username", value=admin_u)
        new_p = st.text_input("New Admin Password", value=admin_p, type="password")
        c1, c2 = st.columns(2)
        with c1:
            if st.form_submit_button("💾 Update Credentials", type="primary", use_container_width=True):
                if new_u.strip() and new_p.strip():
                    save_admin_creds(new_u.strip(), new_p.strip())
                    st.session_state["open_dialog"] = None
                    st.rerun()
                else:
                    st.warning("Username and Password cannot be empty.")
        with c2:
            if st.form_submit_button("❌ Close", use_container_width=True):
                st.session_state["open_dialog"] = None
                st.rerun()


@st.dialog("📥 Bulk Import Cases (Excel / CSV)", width="large")
def admin_bulk_import_dialog(target_sheet):
    st.markdown(f"Target Sheet: **{target_sheet}**")
    up_file = st.file_uploader(
        "Select Excel (.xlsx, .xls) or CSV (.csv) file",
        type=["xlsx", "xls", "csv"],
        key="modal_bulk_import_file",
    )
    if up_file is not None:
        try:
            df_imp = (
                pd.read_csv(up_file, dtype=str).fillna("")
                if up_file.name.lower().endswith(".csv")
                else pd.read_excel(up_file, dtype=str).fillna("")
            )
            st.dataframe(df_imp.head(10), use_container_width=True)
            if st.button("🚀 Import All Rows into Sheet", type="primary", use_container_width=True):
                df_curr = load_data(sheet_name=target_sheet)
                if df_curr is None:
                    df_curr = pd.DataFrame()
                existing_ids = list(df_curr["Case_ID"]) if "Case_ID" in df_curr.columns else []
                reserved_ids = []
                for _, r_imp in df_imp.iterrows():
                    cid_val = str(r_imp.get("Case_ID", "")).strip()
                    if not cid_val or cid_val.lower() == "nan":
                        cid_val = generate_next_case_id(extra_reserved_ids=existing_ids + reserved_ids)
                        reserved_ids.append(cid_val)
                    new_row = {col: "" for col in df_curr.columns}
                    for c in df_imp.columns:
                        if c in new_row:
                            new_row[c] = str(r_imp[c]).strip()
                    new_row["Case_ID"] = cid_val
                    if "Booking_Status" in new_row and not new_row["Booking_Status"]:
                        new_row["Booking_Status"] = "Active"
                    if "Updated_By" in new_row:
                        new_row["Updated_By"] = st.session_state.get("username", "admin")
                    if "Updated_At" in new_row:
                        new_row["Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    if "Action_Type" in new_row:
                        new_row["Action_Type"] = "Bulk Imported"
                    df_curr = pd.concat([df_curr, pd.DataFrame([new_row])], ignore_index=True)
                save_data(df_curr, sheet_name=target_sheet)
                st.session_state["open_dialog"] = None
                st.rerun()
        except Exception as e:
            st.error(f"Error reading file: {e}")

    if st.button("❌ Close", key="close_bulk_import_modal"):
        st.session_state["open_dialog"] = None
        st.rerun()


@st.dialog("➕ Add New Case Record", width="large")
def admin_add_case_dialog(target_sheet):
    df_curr = load_data(sheet_name=target_sheet)
    if df_curr is None or df_curr.empty:
        st.error("Sheet is empty or not found.")
        return
    display_c = find_display_username_column(df_curr)
    login_c = find_login_user_column(df_curr)
    next_id = generate_next_case_id()
    workers_list = ["Not Assigned"] + [w for w in get_all_workers_list() if w.lower() != "not assigned"]

    with st.form("admin_add_case_modal_form"):
        st.text_input("Case_ID (Auto Generated)", value=next_id, disabled=True)
        c1, c2 = st.columns(2)
        input_vals = {}
        skip_cols = {"Case_ID", "Updated_By", "Updated_At", "Action_Type", "Delete_Reason", "Booking_Status", "Appointment_PDF"}
        editable_cols = [c for c in df_curr.columns if c not in skip_cols]
        for idx_c, col_n in enumerate(editable_cols):
            col_box = c1 if idx_c % 2 == 0 else c2
            with col_box:
                if col_n == login_c:
                    input_vals[col_n] = st.selectbox(f"{col_n} (Case Worker)", options=workers_list)
                else:
                    input_vals[col_n] = st.text_input(f"{col_n}", value="")

        status_new = st.selectbox("Booking Status", ["Active", "Booked", "Rescheduled", "Issues"])
        b1, b2 = st.columns(2)
        with b1:
            if st.form_submit_button("💾 Save New Case", type="primary", use_container_width=True):
                new_r = {c: "" for c in df_curr.columns}
                for k, v in input_vals.items():
                    new_r[k] = str(v).strip()
                new_r["Case_ID"] = next_id
                if "Booking_Status" in new_r:
                    new_r["Booking_Status"] = status_new
                if "Updated_By" in new_r:
                    new_r["Updated_By"] = st.session_state.get("username", "admin")
                if "Updated_At" in new_r:
                    new_r["Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                if "Action_Type" in new_r:
                    new_r["Action_Type"] = "Added New Case"
                df_updated = pd.concat([df_curr, pd.DataFrame([new_r])], ignore_index=True)
                save_data(df_updated, sheet_name=target_sheet)
                st.session_state["open_dialog"] = None
                st.rerun()
        with b2:
            if st.form_submit_button("❌ Cancel", use_container_width=True):
                st.session_state["open_dialog"] = None
                st.rerun()


@st.dialog("🔐 Worker Permissions & Login Controls", width="large")
def admin_permissions_dialog():
    workers = get_all_workers_list()
    perms = dict(load_permissions())
    if not workers:
        st.info("No workers found.")
        return

    with st.form("admin_perms_modal_form"):
        updated_perms = {}
        for w in workers:
            p = perms.get(w, {})
            st.markdown(f"**👷 {w}**")
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                can_upd = st.checkbox("Update", value=bool(p.get("can_update", True)), key=f"perm_upd_{w}")
            with c2:
                can_add = st.checkbox("Add Case", value=bool(p.get("can_add", True)), key=f"perm_add_{w}")
            with c3:
                can_dl = st.checkbox("Download", value=bool(p.get("can_download", False)), key=f"perm_dl_{w}")
            with c4:
                can_sync = st.checkbox("Sync", value=bool(p.get("can_sync", False)), key=f"perm_sync_{w}")
            with c5:
                blk = st.checkbox("Block Login", value=bool(p.get("block_login", False)), key=f"perm_blk_{w}")

            updated_perms[w] = {
                "can_edit": can_upd,
                "can_delete_req": True,
                "can_reschedule": True,
                "can_add": can_add,
                "can_update": can_upd,
                "can_search": True,
                "can_download": can_dl,
                "can_sync": can_sync,
                "block_login": blk,
            }

        b1, b2 = st.columns(2)
        with b1:
            if st.form_submit_button("💾 Save Permissions", type="primary", use_container_width=True):
                save_permissions(updated_perms)
                st.session_state["open_dialog"] = None
                st.rerun()
        with b2:
            if st.form_submit_button("❌ Close", use_container_width=True):
                st.session_state["open_dialog"] = None
                st.rerun()


@st.dialog("🕒 Worker Login Status & Attendance Report", width="large")
def admin_attendance_dialog():
    logins = load_login_tracking()
    att_rows = load_attendance()
    if logins:
        st.markdown("#### 🟢 Current Worker Login Status")
        l_df = pd.DataFrame([
            {"Worker": w, "Status": info.get("status", "-"), "Last Login (IST)": info.get("last_login", "-")}
            for w, info in logins.items()
        ])
        st.dataframe(l_df, use_container_width=True, hide_index=True)

    df_day, df_overall = build_attendance_summary_tables(att_rows)
    if not df_overall.empty:
        st.markdown("#### 📊 Overall Worker Hours Summary")
        st.dataframe(df_overall, use_container_width=True)
    if not df_day.empty:
        st.markdown("#### 📅 Day-Wise Attendance Log")
        st.dataframe(df_day, use_container_width=True)

    if st.button("❌ Close", key="close_att_modal"):
        st.session_state["open_dialog"] = None
        st.rerun()


@st.dialog("✏️ Worker Case Edit Requests Approval", width="large")
def admin_edit_requests_dialog():
    reqs = load_edit_requests()
    if not reqs:
        st.info("No edit requests found.")
        if st.button("❌ Close", key="close_er_empty"):
            st.session_state["open_dialog"] = None
            st.rerun()
        return

    for idx_r, r in enumerate(reqs):
        st.markdown(
            f"**{r.get('Request_ID')}** | Worker: **{r.get('Worker')}** | Sheet: `{r.get('Sheet_Name')}` | Case: `{r.get('Case_ID')}` ({r.get('Case_Username')}) | Status: **{r.get('Status')}**"
        )
        st.caption(f"Reason: {r.get('Reason')} | Requested: {r.get('Requested_At')}")
        if str(r.get("Status", "")).strip().lower() == "pending":
            c1, c2, _ = st.columns([1.2, 1.2, 3])
            with c1:
                if st.button("✅ Approve", key=f"app_er_{idx_r}", use_container_width=True):
                    reqs[idx_r]["Status"] = "Approved"
                    reqs[idx_r]["Reviewed_By"] = st.session_state.get("username", "admin")
                    reqs[idx_r]["Reviewed_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    save_edit_requests(reqs)
                    st.rerun()
            with c2:
                if st.button("❌ Reject", key=f"rej_er_{idx_r}", use_container_width=True):
                    reqs[idx_r]["Status"] = "Rejected"
                    reqs[idx_r]["Reviewed_By"] = st.session_state.get("username", "admin")
                    reqs[idx_r]["Reviewed_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    save_edit_requests(reqs)
                    st.rerun()
        st.markdown("---")

    if st.button("❌ Close", key="close_er_modal"):
        st.session_state["open_dialog"] = None
        st.rerun()


@st.dialog("🗑️ Case Delete Requests Approval", width="large")
def admin_delete_requests_dialog():
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    found_any = False
    for s_name, df_s in all_sheets.items():
        if str(s_name).strip() in SYSTEM_SHEETS or df_s is None or df_s.empty:
            continue
        if "Action_Type" not in df_s.columns:
            continue
        del_rows = df_s[df_s["Action_Type"].astype(str).str.strip().str.lower() == "delete request"]
        for r_idx, r_data in del_rows.iterrows():
            found_any = True
            st.markdown(
                f"Sheet: **{s_name}** | Case_ID: **{r_data.get('Case_ID', '')}** | Requested By: **{r_data.get('Updated_By', '')}** | Reason: *{r_data.get('Delete_Reason', '')}*"
            )
            c1, c2, _ = st.columns([1.2, 1.2, 3])
            with c1:
                if st.button("🗑️ Confirm Delete", key=f"conf_del_{s_name}_{r_idx}", use_container_width=True):
                    df_mod = load_data(sheet_name=s_name).drop(r_idx)
                    save_data(df_mod, sheet_name=s_name)
                    st.rerun()
            with c2:
                if st.button("↩️ Reject Delete", key=f"rej_del_{s_name}_{r_idx}", use_container_width=True):
                    df_mod = load_data(sheet_name=s_name)
                    df_mod.at[r_idx, "Action_Type"] = "Delete Rejected"
                    df_mod.at[r_idx, "Updated_By"] = st.session_state.get("username", "admin")
                    df_mod.at[r_idx, "Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    save_data(df_mod, sheet_name=s_name)
                    st.rerun()
            st.markdown("---")

    if not found_any:
        st.info("No pending delete requests.")
    if st.button("❌ Close", key="close_del_req_modal"):
        st.session_state["open_dialog"] = None
        st.rerun()


@st.dialog("➕ Add New Case Record", width="large")
def worker_add_case_dialog(target_sheet, current_user):
    df_curr = load_data(sheet_name=target_sheet)
    if df_curr is None or df_curr.empty:
        st.error("Target sheet not available.")
        return
    login_c = find_login_user_column(df_curr)
    next_id = generate_next_case_id()
    skip_cols = {
        "Case_ID", "Updated_By", "Updated_At", "Action_Type",
        "Delete_Reason", "Booking_Status", "Appointment_PDF",
        "Appointment_Entry", "Preferred_Location", "Preferred_Dates",
        "Prior_Notice", "Issue_Remark",
    }
    fields = [c for c in df_curr.columns if c not in skip_cols]

    with st.form("worker_add_case_modal_form"):
        st.text_input("Case_ID (Auto Generated)", value=next_id, disabled=True)
        vals = {}
        c1, c2 = st.columns(2)
        for i, col_n in enumerate(fields):
            box = c1 if i % 2 == 0 else c2
            with box:
                if col_n == login_c:
                    vals[col_n] = st.text_input(f"{col_n} (Assigned Worker)", value=current_user, disabled=True)
                else:
                    vals[col_n] = st.text_input(f"{col_n}", value="")
        b1, b2 = st.columns(2)
        with b1:
            if st.form_submit_button("💾 Save Case", type="primary", use_container_width=True):
                new_r = {c: "" for c in df_curr.columns}
                for k, v in vals.items():
                    new_r[k] = str(v).strip()
                new_r[login_c] = current_user
                new_r["Case_ID"] = next_id
                if "Booking_Status" in new_r:
                    new_r["Booking_Status"] = "Active"
                if "Updated_By" in new_r:
                    new_r["Updated_By"] = current_user
                if "Updated_At" in new_r:
                    new_r["Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                if "Action_Type" in new_r:
                    new_r["Action_Type"] = "Added New Case"
                df_upd = pd.concat([df_curr, pd.DataFrame([new_r])], ignore_index=True)
                save_data(df_upd, sheet_name=target_sheet)
                st.session_state["open_dialog"] = None
                st.rerun()
        with b2:
            if st.form_submit_button("❌ Cancel", use_container_width=True):
                st.session_state["open_dialog"] = None
                st.rerun()
# ================= END OF STEP 10 =================
# ================= STEP 11: QUERY PARAMS, SHEET INIT, STICKY CONTROL BAR & ADMIN DASHBOARD =================
qp = st.query_params
if qp.get("logged_in") == "true" and qp.get("username"):
    q_user_val = qp.get("username")
    q_role_val = qp.get("role", "worker")
    q_sheet_val = qp.get("active_sheet", "")

    if q_role_val == "worker" and is_worker_blocked(q_user_val):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.role = ""
        st.query_params.clear()
        st.rerun()

    st.session_state.logged_in = True
    st.session_state.username = q_user_val
    st.session_state.role = q_role_val
    if q_sheet_val:
        st.session_state.active_sheet = q_sheet_val
    if q_role_val == "worker":
        record_worker_login(q_user_val)

if "edit_row" in qp and st.session_state.logged_in:
    try:
        r_idx = int(qp.get("edit_row"))
        q_s = qp.get("active_sheet") or st.session_state.get("active_sheet")
        st.session_state["selected_row_idx"] = r_idx
        if q_s:
            st.session_state["active_sheet"] = q_s
        st.session_state["is_editing_popup"] = False
        st.session_state["is_delete_request_mode"] = False
        st.session_state["is_reschedule_mode"] = False
        st.session_state["open_dialog"] = "view_edit_case"
    except Exception:
        pass

render_sticky_top_banner(qp)

role = st.session_state.role
username = st.session_state.username
current_user = username

ensure_all_cases_have_unique_ids(force=False)
sheets = get_sheet_names()
default_active_sheet = get_default_active_sheet(sheets)

if "active_sheet" not in st.session_state or st.session_state.active_sheet not in sheets:
    st.session_state.active_sheet = default_active_sheet

if st.session_state.get("open_dialog") == "view_edit_case" and st.session_state.get("selected_row_idx") is not None:
    dlg_sheet = st.session_state.get("active_sheet") or default_active_sheet
    dlg_row = st.session_state.get("selected_row_idx")
    if role == "admin":
        view_and_edit_case_dialog(dlg_sheet, dlg_row, "admin", username, can_update=True)
    else:
        w_perms = get_worker_permissions(username)
        view_and_edit_case_dialog(
            dlg_sheet,
            dlg_row,
            "worker",
            username,
            can_update=w_perms.get("can_update", True),
        )

active_dialog = st.session_state.get("open_dialog")
if role == "admin":
    if active_dialog == "admin_notice":
        admin_notice_dialog()
    elif active_dialog == "admin_creds":
        admin_creds_dialog()
    elif active_dialog == "admin_bulk_import":
        admin_bulk_import_dialog(st.session_state.active_sheet)
    elif active_dialog == "admin_add_case":
        admin_add_case_dialog(st.session_state.active_sheet)
    elif active_dialog == "admin_perms":
        admin_permissions_dialog()
    elif active_dialog == "admin_attendance":
        admin_attendance_dialog()
    elif active_dialog == "admin_edit_requests":
        admin_edit_requests_dialog()
    elif active_dialog == "admin_delete_requests":
        admin_delete_requests_dialog()
elif role == "worker":
    if active_dialog == "worker_add_case":
        worker_add_case_dialog(st.session_state.active_sheet, username)


def render_sticky_control_bar(
    role_label,
    user_label,
    sheet_label,
    is_admin=False,
    can_sync=True,
    can_download=True,
    export_bytes=None,
    export_filename="Heman_Report.xlsx",
    logout_key="sticky_logout_btn",
    sync_key="sticky_sync_btn",
    dl_key="sticky_dl_btn",
    home_key="sticky_home_btn",
):
    st.markdown('<div class="sticky-control-panel">', unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns([4.6, 1.35, 1.35, 1.45, 1.35], gap="small")
    with c1:
        st.markdown(
            f"""
            <div class="workspace-header-bar">
                <span>👤 {role_label} : {user_label}</span>
                <span class="workspace-sheet-pill">📄 {sheet_label}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        if st.button("🏠 Home", key=home_key, use_container_width=True):
            close_case_popup_state()
            st.session_state["open_dialog"] = None
            st.query_params.clear()
            st.query_params["logged_in"] = "true"
            st.query_params["username"] = username
            st.query_params["role"] = role
            st.query_params["active_sheet"] = st.session_state.get("active_sheet", default_active_sheet)
            if is_admin:
                st.session_state["active_view"] = "Home"
                st.session_state["admin_page"] = "Home"
            else:
                st.session_state["worker_tab_radio"] = "🏠 My Dashboard"
            st.rerun()
    with c3:
        if can_sync:
            if st.button("🔄 Sync", key=sync_key, use_container_width=True):
                _clear_sheet_runtime_caches()
                st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
                st.session_state["_last_id_sync_mtime"] = None
                ensure_all_cases_have_unique_ids(force=True)
                st.rerun()
        else:
            st.button("🔄 Sync", key=sync_key, disabled=True, use_container_width=True)
    with c4:
        if can_download and export_bytes is not None:
            st.download_button(
                label="📥 Download",
                data=export_bytes,
                file_name=export_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=dl_key,
                use_container_width=True,
            )
        else:
            st.button("📥 Download", key=dl_key, disabled=True, use_container_width=True)
    with c5:
        if st.button("🚪 Logout", key=logout_key, use_container_width=True):
            close_case_popup_state()
            st.session_state["open_dialog"] = None
            if not is_admin:
                record_worker_logout(username)
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.role = ""
            st.session_state.pop("_mem_sheets_cache", None)
            st.session_state.pop("_parsed_meta_cache", None)
            st.query_params.clear()
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)


def get_sheet_status_counts(df_sheet):
    total_cnt = len(df_sheet) if df_sheet is not None else 0
    if df_sheet is None or df_sheet.empty or "Booking_Status" not in df_sheet.columns:
        return {
            "total": total_cnt,
            "booked": 0,
            "rescheduled": 0,
            "issues": 0,
            "active": total_cnt,
        }
    st_series = df_sheet["Booking_Status"].astype(str).str.strip().str.lower()
    booked_cnt = int((st_series == "booked").sum())
    resched_cnt = int((st_series == "rescheduled").sum())
    issues_cnt = int((st_series == "issues").sum())
    active_cnt = int((st_series.isin(["", "active"])).sum())
    return {
        "total": total_cnt,
        "booked": booked_cnt,
        "rescheduled": resched_cnt,
        "issues": issues_cnt,
        "active": active_cnt,
    }


# ================= ADMIN DASHBOARD =================
if role == "admin":
    if "active_view" not in st.session_state:
        st.session_state["active_view"] = "Home"
    if "admin_page" not in st.session_state:
        st.session_state["admin_page"] = "Home"

    all_admin_sheets = load_all_sheets_cached(_get_file_mtime())
    all_workers_names = get_all_workers_list()
    all_perms = load_permissions()
    all_logins = load_login_tracking()
    all_attendance = load_attendance()
    all_edit_reqs = load_edit_requests()

    booked_sheet_name = next((s for s in sheets if "booked" in s.lower()), None)
    issues_sheet_name = next((s for s in sheets if "issue" in s.lower()), None)

    df_booked_sheet = (
        all_admin_sheets.get(booked_sheet_name, pd.DataFrame())
        if booked_sheet_name
        else pd.DataFrame()
    )
    df_issues_sheet = (
        all_admin_sheets.get(issues_sheet_name, pd.DataFrame())
        if issues_sheet_name
        else pd.DataFrame()
    )

    active_sheets_list = [
        s
        for s in sheets
        if s != booked_sheet_name and s != issues_sheet_name
    ]
    total_active_cases_all = 0
    for s_act in active_sheets_list:
        df_a = all_admin_sheets.get(s_act)
        if df_a is not None and not df_a.empty:
            total_active_cases_all += len(df_a)

    total_booked_all = len(df_booked_sheet) if df_booked_sheet is not None else 0
    total_issues_all = len(df_issues_sheet) if df_issues_sheet is not None else 0
    total_resched_all = 0
    for s_k, df_k in all_admin_sheets.items():
        if str(s_k).strip() in SYSTEM_SHEETS or df_k is None or df_k.empty:
            continue
        if "Booking_Status" in df_k.columns:
            total_resched_all += int(
                (df_k["Booking_Status"].astype(str).str.strip().str.lower() == "rescheduled").sum()
            )

    pending_edit_reqs_cnt = sum(
        1
        for er in all_edit_reqs
        if str(er.get("Status", "")).strip().lower() == "pending"
    )
    pending_del_reqs_cnt = 0
    for s_k, df_k in all_admin_sheets.items():
        if str(s_k).strip() in SYSTEM_SHEETS or df_k is None or df_k.empty:
            continue
        if "Action_Type" in df_k.columns:
            pending_del_reqs_cnt += int(
                (df_k["Action_Type"].astype(str).str.strip().str.lower() == "delete request").sum()
            )

    admin_export_bytes = get_admin_gsheet_export_bytes(_get_file_mtime())

    render_sticky_control_bar(
        role_label="Admin",
        user_label=username,
        sheet_label=st.session_state.active_sheet,
        is_admin=True,
        can_sync=True,
        can_download=True,
        export_bytes=admin_export_bytes,
        export_filename="HemanVisa_All_Sheets_Export.xlsx",
        logout_key="admin_sticky_logout",
        sync_key="admin_sticky_sync",
        dl_key="admin_sticky_dl",
        home_key="admin_sticky_home",
    )

    notice_text = get_admin_notice()
    if notice_text:
        st.markdown(
            f'<div class="marquee-container"><marquee behavior="scroll" direction="left" scrollamount="6">📢 {notice_text}</marquee></div>',
            unsafe_allow_html=True,
        )

    admin_nav_choice = st.radio(
        "Admin Navigation",
        [
            "Home",
            "Sheets & Cases",
            "Worker & Attendance",
            "Requests & Approvals",
        ],
        index=[
            "Home",
            "Sheets & Cases",
            "Worker & Attendance",
            "Requests & Approvals",
        ].index(
            st.session_state["active_view"]
            if st.session_state["active_view"]
            in ["Home", "Sheets & Cases", "Worker & Attendance", "Requests & Approvals"]
            else "Home"
        ),
        horizontal=True,
        label_visibility="collapsed",
        key="admin_main_nav_radio",
    )
    if admin_nav_choice != st.session_state["active_view"]:
        st.session_state["active_view"] = admin_nav_choice
        st.rerun()

    st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)

    active_view = st.session_state["active_view"]

    if active_view == "Home":
        st.markdown(
            """
            <div class="hero-section" style="padding: 24px 26px; margin-bottom: 14px;">
                <div class="hero-title" style="font-size: 26px; margin-bottom: 6px;">🌍 HemanVisa Admin Control Center</div>
                <div class="hero-subtitle" style="font-size: 14px;">
                    Live overview of Active Cases, Booked Appointments, Reschedules, Issues, Case Workers, Attendance & Pending Requests.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        k1, k2, k3, k4, k5 = st.columns(5, gap="small")
        with k1:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-title">Active Cases</div><div class="kpi-value">{total_active_cases_all}</div></div>',
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                f'<div class="kpi-card green"><div class="kpi-title">Booked Cases</div><div class="kpi-value">{total_booked_all}</div></div>',
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                f'<div class="kpi-card gold"><div class="kpi-title">Rescheduled</div><div class="kpi-value">{total_resched_all}</div></div>',
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                f'<div class="kpi-card rose"><div class="kpi-title">Issues</div><div class="kpi-value">{total_issues_all}</div></div>',
                unsafe_allow_html=True,
            )
        with k5:
            st.markdown(
                f'<div class="kpi-card purple"><div class="kpi-title">Pending Requests</div><div class="kpi-value">{pending_edit_reqs_cnt + pending_del_reqs_cnt}</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="section-banner">⚡ Quick Admin Actions</div>', unsafe_allow_html=True)
        qa1, qa2, qa3, qa4, qa5, qa6 = st.columns(6, gap="small")
        with qa1:
            st.markdown('<span class="admin-btn-0"></span>', unsafe_allow_html=True)
            if st.button("➕ Add Case", key="qa_add_case", use_container_width=True):
                st.session_state["open_dialog"] = "admin_add_case"
                st.rerun()
        with qa2:
            st.markdown('<span class="admin-btn-1"></span>', unsafe_allow_html=True)
            if st.button("📥 Bulk Import", key="qa_bulk_imp", use_container_width=True):
                st.session_state["open_dialog"] = "admin_bulk_import"
                st.rerun()
        with qa3:
            st.markdown('<span class="admin-btn-2"></span>', unsafe_allow_html=True)
            if st.button("📢 Notice Board", key="qa_notice", use_container_width=True):
                st.session_state["open_dialog"] = "admin_notice"
                st.rerun()
        with qa4:
            st.markdown('<span class="admin-btn-3"></span>', unsafe_allow_html=True)
            if st.button("🔐 Permissions", key="qa_perms", use_container_width=True):
                st.session_state["open_dialog"] = "admin_perms"
                st.rerun()
        with qa5:
            st.markdown('<span class="admin-btn-4"></span>', unsafe_allow_html=True)
            if st.button("🕒 Attendance", key="qa_att", use_container_width=True):
                st.session_state["open_dialog"] = "admin_attendance"
                st.rerun()
        with qa6:
            st.markdown('<span class="admin-btn-5"></span>', unsafe_allow_html=True)
            if st.button("🔑 Admin Creds", key="qa_creds", use_container_width=True):
                st.session_state["open_dialog"] = "admin_creds"
                st.rerun()

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="section-banner">📂 All Sheets Summary</div>', unsafe_allow_html=True)
        sheet_sum_rows = []
        for s_item in sheets:
            df_s_item = all_admin_sheets.get(s_item, pd.DataFrame())
            cnts = get_sheet_status_counts(df_s_item)
            sheet_sum_rows.append({
                "Sheet Name": s_item,
                "Total Rows": cnts["total"],
                "Active": cnts["active"],
                "Booked": cnts["booked"],
                "Rescheduled": cnts["rescheduled"],
                "Issues": cnts["issues"],
            })
        if sheet_sum_rows:
            st.dataframe(pd.DataFrame(sheet_sum_rows), use_container_width=True, hide_index=True)

    elif active_view == "Sheets & Cases":
        st.markdown('<div class="section-banner">📂 Select Sheet & Search / Manage Cases</div>', unsafe_allow_html=True)

        s_cols = st.columns(min(len(sheets), 6) if sheets else 1, gap="small")
        for idx_s, s_name in enumerate(sheets):
            col_b = s_cols[idx_s % len(s_cols)]
            with col_b:
                st.markdown(f'<span class="visa-color-{idx_s % 8}"></span>', unsafe_allow_html=True)
                label_s = f"{'✅ ' if st.session_state.active_sheet == s_name else ''}{s_name}"
                if st.button(label_s, key=f"admin_sel_sheet_{s_name}", use_container_width=True):
                    st.session_state.active_sheet = s_name
                    st.session_state[f"page_num_admin_{s_name}"] = 1
                    st.rerun()

        curr_s_name = st.session_state.active_sheet
        df_curr_sheet = load_data(sheet_name=curr_s_name)
        if df_curr_sheet is None:
            df_curr_sheet = pd.DataFrame()
        df_curr_sheet = reorder_dataframe_columns(df_curr_sheet)

        f_c1, f_c2, f_c3, f_c4 = st.columns([2.4, 1.5, 1.5, 1.4], gap="small")
        with f_c1:
            search_q = st.text_input(
                "🔍 Search Case ID, Username, Password, DS-160 or Worker",
                value="",
                placeholder="Type to search across current or all sheets...",
                key="admin_case_search_input",
            )
        with f_c2:
            worker_filter_opts = ["All Workers"] + [w for w in all_workers_names]
            sel_w_filter = st.selectbox("👷 Filter Worker", worker_filter_opts, key="admin_w_filter")
        with f_c3:
            status_filter_opts = ["All Status", "Active", "Booked", "Rescheduled", "Issues"]
            sel_st_filter = st.selectbox("📌 Filter Status", status_filter_opts, key="admin_st_filter")
        with f_c4:
            search_all_sheets_chk = st.checkbox("🌐 Search All Sheets", value=False, key="admin_search_all_chk")

        if not df_curr_sheet.empty:
            login_c_curr = find_login_user_column(df_curr_sheet)
            disp_c_curr = find_display_username_column(df_curr_sheet)
            pass_c_curr = find_password_column(df_curr_sheet)

            if search_q.strip() and search_all_sheets_chk:
                df_filtered = search_across_all_sheets_by_fields(search_q.strip())
            else:
                df_work = df_curr_sheet.copy()
                df_work["__orig_idx__"] = df_work.index
                if sel_w_filter != "All Workers" and login_c_curr in df_work.columns:
                    df_work = df_work[
                        df_work[login_c_curr].astype(str).str.strip().str.lower()
                        == sel_w_filter.strip().lower()
                    ]
                if sel_st_filter != "All Status" and "Booking_Status" in df_work.columns:
                    df_work = df_work[
                        df_work["Booking_Status"].astype(str).str.strip().str.lower()
                        == sel_st_filter.strip().lower()
                    ]
                if search_q.strip():
                    mask_s = search_username_ds160_password_worker(df_work, search_q.strip())
                    df_work = df_work[mask_s]
                df_filtered = df_work

            st.caption(f"Showing **{len(df_filtered)}** matching record(s) in **{curr_s_name}** (Click any Username link to view, edit, reschedule, or download PDF).")
            render_clickable_table(
                df=df_filtered,
                display_c=disp_c_curr,
                p_col=pass_c_curr,
                qp=qp,
                table_key=f"admin_{curr_s_name}",
                sheet_for_link=curr_s_name,
                page_size=100,
                allow_edit=True,
                restrict_edit_to_user=None,
            )
        else:
            st.info(f"No rows found in sheet '{curr_s_name}'.")

    elif active_view == "Worker & Attendance":
        st.markdown('<div class="section-banner">👷 Case Workers, Permissions & Login Attendance</div>', unsafe_allow_html=True)
        btn_a1, btn_a2, btn_a3 = st.columns(3, gap="small")
        with btn_a1:
            if st.button("🔐 Configure Worker Permissions", type="primary", use_container_width=True, key="wa_btn_perms"):
                st.session_state["open_dialog"] = "admin_perms"
                st.rerun()
        with btn_a2:
            if st.button("🕒 View Full Attendance & Hours", use_container_width=True, key="wa_btn_att"):
                st.session_state["open_dialog"] = "admin_attendance"
                st.rerun()
        with btn_a3:
            all_blk = are_all_workers_blocked()
            blk_label = "🔓 Unblock All Workers" if all_blk else "🚫 Block All Workers"
            if st.button(blk_label, use_container_width=True, key="wa_btn_blk_all"):
                toggle_all_workers_block(not all_blk)
                st.rerun()

        df_day_sum, df_overall_sum = build_attendance_summary_tables(all_attendance)
        if not df_overall_sum.empty:
            st.markdown("#### 📊 Overall Worker Hours Summary")
            st.dataframe(df_overall_sum, use_container_width=True)
        if not df_day_sum.empty:
            st.markdown("#### 📅 Day-Wise Attendance Log")
            st.dataframe(df_day_sum, use_container_width=True)

    elif active_view == "Requests & Approvals":
        st.markdown('<div class="section-banner">📬 Edit Requests & Delete Requests</div>', unsafe_allow_html=True)
        r_c1, r_c2 = st.columns(2, gap="small")
        with r_c1:
            if st.button(f"✏️ Manage Edit Requests ({pending_edit_reqs_cnt} Pending)", type="primary", use_container_width=True, key="req_btn_edit"):
                st.session_state["open_dialog"] = "admin_edit_requests"
                st.rerun()
        with r_c2:
            if st.button(f"🗑️ Manage Delete Requests ({pending_del_reqs_cnt} Pending)", use_container_width=True, key="req_btn_del"):
                st.session_state["open_dialog"] = "admin_delete_requests"
                st.rerun()

        if all_edit_reqs:
            st.markdown("#### ✏️ Recent Edit Requests")
            st.dataframe(pd.DataFrame(all_edit_reqs), use_container_width=True, hide_index=True)
        else:
            st.info("No worker edit requests logged.")
# ================= END OF STEP 11 =================
# ================= STEP 12: CASE WORKER DASHBOARD (PART 1) =================
elif role == "worker":
    w_perms = get_worker_permissions(username)
    can_update_w = bool(w_perms.get("can_update", True))
    can_add_w = bool(w_perms.get("can_add", True))
    can_search_w = bool(w_perms.get("can_search", True))
    can_download_w = bool(w_perms.get("can_download", False))
    can_sync_w = bool(w_perms.get("can_sync", False))

    worker_export_bytes = (
        get_worker_export_bytes(_get_file_mtime(), username, st.session_state.active_sheet)
        if can_download_w
        else None
    )

    render_sticky_control_bar(
        role_label="Case Worker",
        user_label=username,
        sheet_label=st.session_state.active_sheet,
        is_admin=False,
        can_sync=can_sync_w,
        can_download=can_download_w,
        export_bytes=worker_export_bytes,
        export_filename=f"HemanVisa_{username}_Cases.xlsx",
        logout_key="worker_sticky_logout",
        sync_key="worker_sticky_sync",
        dl_key="worker_sticky_dl",
        home_key="worker_sticky_home",
    )

    notice_text = get_admin_notice()
    if notice_text:
        st.markdown(
            f'<div class="marquee-container"><marquee behavior="scroll" direction="left" scrollamount="6">📢 {notice_text}</marquee></div>',
            unsafe_allow_html=True,
        )

    today_ist_str = get_ist_now().strftime("%Y-%m-%d")
    all_w_sheets = load_all_sheets_cached(_get_file_mtime())

    my_active_cnt = 0
    my_booked_cnt = 0
    my_resched_cnt = 0
    my_issues_cnt = 0
    my_today_upd_cnt = 0

    for s_k, df_k in all_w_sheets.items():
        if str(s_k).strip() in SYSTEM_SHEETS or df_k is None or df_k.empty:
            continue
        lc_k = find_login_user_column(df_k)
        if lc_k in df_k.columns:
            sub_k = df_k[
                df_k[lc_k].astype(str).str.strip().str.lower()
                == str(username).strip().lower()
            ]
            if not sub_k.empty:
                s_low_k = str(s_k).strip().lower()
                if "booked" in s_low_k:
                    my_booked_cnt += len(sub_k)
                elif "issue" in s_low_k:
                    my_issues_cnt += len(sub_k)
                else:
                    my_active_cnt += len(sub_k)
                    if "Booking_Status" in sub_k.columns:
                        st_s = sub_k["Booking_Status"].astype(str).str.strip().str.lower()
                        my_resched_cnt += int((st_s == "rescheduled").sum())
                for _, r_row in sub_k.iterrows():
                    if is_today_updated_row(r_row, today_ist_str):
                        my_today_upd_cnt += 1

    st.markdown(
        f"""
        <div class="worker-hero-banner">
            <div class="worker-hero-title">👷 Welcome back, {username}</div>
            <div class="worker-pill-group">
                <span class="worker-pill-gold">📅 IST Date: {today_ist_str}</span>
                <span class="worker-pill-cyan">⚡ Today Updated: {my_today_upd_cnt}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    wk1, wk2, wk3, wk4, wk5 = st.columns(5, gap="small")
    with wk1:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">My Active Cases</div><div class="kpi-value">{my_active_cnt}</div></div>',
            unsafe_allow_html=True,
        )
    with wk2:
        st.markdown(
            f'<div class="kpi-card green"><div class="kpi-title">My Booked</div><div class="kpi-value">{my_booked_cnt}</div></div>',
            unsafe_allow_html=True,
        )
    with wk3:
        st.markdown(
            f'<div class="kpi-card gold"><div class="kpi-title">Rescheduled</div><div class="kpi-value">{my_resched_cnt}</div></div>',
            unsafe_allow_html=True,
        )
    with wk4:
        st.markdown(
            f'<div class="kpi-card rose"><div class="kpi-title">Issues</div><div class="kpi-value">{my_issues_cnt}</div></div>',
            unsafe_allow_html=True,
        )
    with wk5:
        st.markdown(
            f'<div class="kpi-card purple"><div class="kpi-title">Updated Today</div><div class="kpi-value">{my_today_upd_cnt}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)

    worker_tab_choice = st.radio(
        "Worker Workspace Tabs",
        ["🏠 My Dashboard", "📋 Assigned Cases", "🔎 Global Search"],
        index=["🏠 My Dashboard", "📋 Assigned Cases", "🔎 Global Search"].index(
            st.session_state.get("worker_tab_radio", "🏠 My Dashboard")
            if st.session_state.get("worker_tab_radio", "🏠 My Dashboard")
            in ["🏠 My Dashboard", "📋 Assigned Cases", "🔎 Global Search"]
            else "🏠 My Dashboard"
        ),
        horizontal=True,
        label_visibility="collapsed",
        key="worker_tab_radio_ctrl",
    )
    st.session_state["worker_tab_radio"] = worker_tab_choice

    st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)

    if worker_tab_choice == "🏠 My Dashboard":
        st.markdown('<div class="section-banner">📌 Quick Actions & Sheet Overview</div>', unsafe_allow_html=True)
        w_act1, w_act2, w_act3 = st.columns([1.4, 1.4, 3.2], gap="small")
        with w_act1:
            if can_add_w:
                if st.button("➕ Add New Case", type="primary", use_container_width=True, key="w_btn_add_case"):
                    st.session_state["open_dialog"] = "worker_add_case"
                    st.rerun()
            else:
                st.button("➕ Add New Case", disabled=True, use_container_width=True, key="w_btn_add_case_dis")
        with w_act2:
            if st.button("📋 Open Assigned Cases", use_container_width=True, key="w_btn_go_cases"):
                st.session_state["worker_tab_radio"] = "📋 Assigned Cases"
                st.rerun()
        with w_act3:
            st.info(f"Permissions: Update (`{can_update_w}`) • Add (`{can_add_w}`) • Download (`{can_download_w}`) • Sync (`{can_sync_w}`)")
# ================= END OF STEP 12 =================
# ================= STEP 13: CASE WORKER DASHBOARD (PART 2 & FINAL CLOSE) =================
    elif worker_tab_choice == "📋 Assigned Cases":
        st.markdown('<div class="section-banner">📋 My Assigned Cases & Sheet Switcher</div>', unsafe_allow_html=True)

        ws_cols = st.columns(min(len(sheets), 6) if sheets else 1, gap="small")
        for idx_ws, ws_name in enumerate(sheets):
            col_wb = ws_cols[idx_ws % len(ws_cols)]
            with col_wb:
                st.markdown(f'<span class="visa-color-{idx_ws % 8}"></span>', unsafe_allow_html=True)
                w_lbl = f"{'✅ ' if st.session_state.active_sheet == ws_name else ''}{ws_name}"
                if st.button(w_lbl, key=f"worker_sel_sheet_{ws_name}", use_container_width=True):
                    st.session_state.active_sheet = ws_name
                    st.session_state[f"page_num_worker_{ws_name}"] = 1
                    st.rerun()

        curr_w_sheet = st.session_state.active_sheet
        df_w_sheet = load_data(sheet_name=curr_w_sheet)
        if df_w_sheet is None:
            df_w_sheet = pd.DataFrame()
        df_w_sheet = reorder_dataframe_columns(df_w_sheet)

        wf1, wf2, wf3 = st.columns([2.6, 1.5, 1.4], gap="small")
        with wf1:
            w_search_q = st.text_input(
                "🔍 Search Case ID, Username, DS-160 or Password",
                value="",
                placeholder="Search in current sheet...",
                key="worker_sheet_search_input",
                disabled=not can_search_w,
            )
        with wf2:
            w_st_filter = st.selectbox(
                "📌 Status Filter",
                ["All Status", "Active", "Booked", "Rescheduled", "Issues"],
                key="worker_st_filter",
            )
        with wf3:
            only_mine_chk = st.checkbox(
                "👤 Only My Cases",
                value=True,
                key="worker_only_mine_chk",
            )

        if not df_w_sheet.empty:
            lc_w = find_login_user_column(df_w_sheet)
            dc_w = find_display_username_column(df_w_sheet)
            pc_w = find_password_column(df_w_sheet)

            df_w_work = df_w_sheet.copy()
            df_w_work["__orig_idx__"] = df_w_work.index

            if only_mine_chk and lc_w in df_w_work.columns:
                df_w_work = df_w_work[
                    df_w_work[lc_w].astype(str).str.strip().str.lower()
                    == str(username).strip().lower()
                ]

            if w_st_filter != "All Status" and "Booking_Status" in df_w_work.columns:
                df_w_work = df_w_work[
                    df_w_work["Booking_Status"].astype(str).str.strip().str.lower()
                    == w_st_filter.strip().lower()
                ]

            if can_search_w and w_search_q.strip():
                w_mask = search_username_ds160_password_worker(df_w_work, w_search_q.strip())
                df_w_work = df_w_work[w_mask]

            st.caption(f"Showing **{len(df_w_work)}** record(s) in **{curr_w_sheet}**.")
            render_clickable_table(
                df=df_w_work,
                display_c=dc_w,
                p_col=pc_w,
                qp=qp,
                table_key=f"worker_{curr_w_sheet}",
                sheet_for_link=curr_w_sheet,
                page_size=100,
                allow_edit=True,
                restrict_edit_to_user=None,
            )
        else:
            st.info(f"No cases available in '{curr_w_sheet}'.")

    elif worker_tab_choice == "🔎 Global Search":
        st.markdown('<div class="section-banner">🔎 Search Across All Sheets</div>', unsafe_allow_html=True)
        g_q = st.text_input(
            "Enter Case ID, Username, DS-160, Password or Case Worker",
            value="",
            placeholder="Search across all sheets...",
            key="worker_global_search_q",
            disabled=not can_search_w,
        )
        if can_search_w and g_q.strip():
            g_res = search_across_all_sheets_by_fields(g_q.strip())
            if not g_res.empty:
                g_disp_c = find_display_username_column(g_res)
                g_pass_c = find_password_column(g_res)
                st.caption(f"Found **{len(g_res)}** matching case(s) across sheets.")
                render_clickable_table(
                    df=g_res,
                    display_c=g_disp_c,
                    p_col=g_pass_c,
                    qp=qp,
                    table_key="worker_global_search_tbl",
                    sheet_for_link="",
                    page_size=100,
                    allow_edit=True,
                    restrict_edit_to_user=None,
                )
            else:
                st.info("No matching cases found across any sheet.")
        elif not can_search_w:
            st.warning("⚠️ Search permission is disabled for your account.")
        else:
            st.info("Type a keyword above to search across all sheets.")
# ================= END OF STEP 13 (COMPLETE FILE) =================
