# ================= STEP 1: IMPORTS, CONSTANTS & PAGE SETUP =================
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
# ================= END OF STEP 1 =================
# ================= STEP 2: FORMATTING & CLIPBOARD HELPERS =================
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
# ================= END OF STEP 2 =================
# ================= STEP 3: GOOGLE SHEETS WEBAPP SYNC & CACHE LAYER =================
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
# ================= END OF STEP 3 =================
# ================= STEP 4: IST TIME & RUNTIME METADATA HELPERS =================
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
# ================= END OF STEP 4 =================
# ================= STEP 5: EDIT REQUESTS, ADMIN CONFIG & PERMISSIONS =================
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
# ================= END OF STEP 5 =================
# ================= STEP 6: LOGIN TRACKING & ATTENDANCE / HOURS ENGINE =================
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
# ================= END OF STEP 6 =================
# ================= STEP 7: CASE ID GENERATION, RESCHEDULE & DATA LOAD/SAVE =================
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
# ================= STEP 9: UI STYLING, KPI CARDS & TABLE RENDERER =================
def inject_custom_css():
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
            max-width: 100% !important;
            padding-left: 1.5rem !important;
            padding-right: 1.5rem !important;
        }
        [data-testid="stAppViewContainer"] .main .block-container {
            max-width: 100% !important;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            border-bottom: 2px solid #e2e8f0;
            padding-bottom: 4px;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            scrollbar-width: thin;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: #f8fafc;
            border: 1px solid #cbd5e1;
            border-bottom: none;
            border-radius: 8px 8px 0 0;
            padding: 8px 14px;
            font-weight: 700;
            font-size: 13px;
            color: #334155;
            white-space: nowrap !important;
            flex-shrink: 0 !important;
            transition: all 0.15s ease;
        }
        .stTabs [data-baseweb="tab"]:hover {
            background-color: #eef2ff;
            color: #1e3a8a;
        }
        .stTabs [aria-selected="true"] {
            background: linear-gradient(180deg, #1e3a8a 0%, #1d4ed8 100%) !important;
            color: #ffffff !important;
            border-color: #1e3a8a !important;
            box-shadow: 0 -2px 6px rgba(29, 78, 216, 0.2);
        }

        .kpi-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 14px 16px;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.05);
            position: relative;
            overflow: hidden;
            height: 100%;
        }
        .kpi-card::before {
            content: "";
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 4px;
            background: var(--kpi-accent, #2563eb);
        }
        .kpi-title {
            font-size: 11.5px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #64748b;
            margin-bottom: 6px;
        }
        .kpi-value {
            font-size: 24px;
            font-weight: 800;
            color: #0f172a;
            line-height: 1.15;
        }
        .kpi-sub {
            font-size: 11px;
            color: #64748b;
            margin-top: 4px;
        }

        .notice-ticker-wrap {
            background: linear-gradient(90deg, #fffbeb 0%, #fef3c7 50%, #fffbeb 100%);
            border: 1px solid #f59e0b;
            border-radius: 10px;
            padding: 8px 14px;
            margin: 6px 0 12px 0;
            display: flex;
            align-items: center;
            gap: 10px;
            overflow: hidden;
            box-shadow: 0 2px 6px rgba(245, 158, 11, 0.12);
        }
        .notice-badge {
            background: #b45309;
            color: #ffffff;
            font-size: 11px;
            font-weight: 800;
            padding: 3px 9px;
            border-radius: 999px;
            white-space: nowrap;
            letter-spacing: 0.4px;
        }
        .notice-marquee {
            overflow: hidden;
            white-space: nowrap;
            flex: 1;
        }
        .notice-marquee-text {
            display: inline-block;
            padding-left: 100%;
            color: #78350f;
            font-size: 13px;
            font-weight: 700;
            animation: noticeScroll 18s linear infinite;
        }
        .notice-marquee-text:hover {
            animation-play-state: paused;
        }
        @keyframes noticeScroll {
            0% { transform: translateX(0%); }
            100% { transform: translateX(-200%); }
        }

        .html-table-wrap {
            width: 100%;
            overflow-x: auto;
            overflow-y: auto;
            max-height: 460px;
            border: 1px solid #cbd5e1;
            border-radius: 10px;
            background: #ffffff;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
            margin-bottom: 12px;
        }
        .styled-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12.5px;
            text-align: left;
            color: #1e293b;
        }
        .styled-table thead th {
            position: sticky;
            top: 0;
            background: #f1f5f9;
            color: #0f172a;
            font-weight: 700;
            padding: 7px 10px;
            border-bottom: 2px solid #cbd5e1;
            border-right: 1px solid #e2e8f0;
            white-space: nowrap;
            z-index: 2;
        }
        .styled-table tbody td {
            padding: 6px 10px;
            border-bottom: 1px solid #e2e8f0;
            border-right: 1px solid #f1f5f9;
            white-space: nowrap;
            max-width: 260px;
            overflow: hidden;
            text-overflow: ellipsis;
            vertical-align: middle;
        }
        .styled-table tbody tr:nth-child(even) {
            background-color: #f8fafc;
        }
        .styled-table tbody tr:hover {
            background-color: #eff6ff !important;
        }
        .styled-table tbody tr.clickable-row {
            cursor: pointer;
        }
        .styled-table tbody tr.clickable-row:hover {
            background-color: #dbeafe !important;
        }
        .styled-table tbody tr.today-updated-row {
            background-color: #dcfce7 !important;
        }
        .styled-table tbody tr.today-updated-row:hover {
            background-color: #bbf7d0 !important;
        }
        .row-link {
            color: inherit;
            text-decoration: none;
            display: block;
            width: 100%;
            height: 100%;
        }

        [data-testid="stSidebar"] .stButton > button,
        [data-testid="stSidebar"] .stDownloadButton > button {
            padding: 0.34rem 0.65rem !important;
            min-height: 2.05rem !important;
            font-size: 0.82rem !important;
            line-height: 1.2 !important;
            border-radius: 7px !important;
        }
        [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
            gap: 0.38rem !important;
        }
        [data-testid="stSidebar"] hr {
            margin: 0.45rem 0 !important;
        }
        [data-testid="stSidebar"] .stAlert {
            padding: 0.4rem 0.65rem !important;
            font-size: 0.8rem !important;
            margin-top: 0.25rem !important;
            margin-bottom: 0.25rem !important;
        }
        .compact-divider {
            border: 0;
            border-top: 1px solid #e2e8f0;
            margin: 8px 0 10px 0;
        }

        .edit-modal-card {
            background: linear-gradient(135deg, #f8fafc 0%, #eff6ff 100%);
            border: 1.5px solid #93c5fd;
            border-radius: 12px;
            padding: 16px 18px;
            margin: 10px 0 14px 0;
            box-shadow: 0 6px 18px rgba(30, 58, 138, 0.12);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_card(title, value, accent="#2563eb", subtitle=""):
    sub_html = f'<div class="kpi-sub">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f"""
        <div class="kpi-card" style="--kpi-accent: {accent};">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
            {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_html_table(
    df_display,
    max_height=460,
    clickable_rows=False,
    row_indices=None,
    row_key_prefix="row",
    highlight_today_rows=False,
    today_str="",
):
    if df_display is None or df_display.empty:
        st.info("చూపించడానికి రికార్డులు ఏవీ లేవు.")
        return

    df_show = df_display.fillna("").astype(str)
    headers_html = "".join(f"<th>{col}</th>" for col in df_show.columns)

    rows_html_list = []
    for pos, (_, r) in enumerate(df_show.iterrows()):
        orig_idx = row_indices[pos] if (row_indices is not None and pos < len(row_indices)) else pos
        is_today = False
        if highlight_today_rows and today_str:
            upd_at = str(r.get("Updated_At", "")).strip()
            upd_by = str(r.get("Updated_By", "")).strip()
            if today_str in upd_at and upd_by != "":
                is_today = True

        row_classes = []
        if clickable_rows:
            row_classes.append("clickable-row")
        if is_today:
            row_classes.append("today-updated-row")
        tr_class_attr = f' class="{" ".join(row_classes)}"' if row_classes else ""

        cells_html_list = []
        for col_name in df_show.columns:
            val_str = str(r[col_name])
            safe_val = (
                val_str.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )
            if clickable_rows:
                href_val = f"?edit_row={urllib.parse.quote(str(orig_idx))}"
                cells_html_list.append(
                    f'<td><a class="row-link" href="{href_val}" target="_self" title="{safe_val}">{safe_val}</a></td>'
                )
            else:
                cells_html_list.append(f'<td title="{safe_val}">{safe_val}</td>')

        rows_html_list.append(f"<tr{tr_class_attr}>{''.join(cells_html_list)}</tr>")

    table_html = (
        f'<div class="html-table-wrap" style="max-height: {max_height}px;">'
        f'<table class="styled-table">'
        f'<thead><tr>{headers_html}</tr></thead>'
        f'<tbody>{"".join(rows_html_list)}</tbody>'
        f'</table>'
        f'</div>'
    )
    st.markdown(table_html, unsafe_allow_html=True)
# ================= END OF STEP 9 =================
# ================= STEP 10: CASE ACTION / EDIT MODAL DIALOG =================
@st.dialog("✏️ Case Details, Reschedule & CGI Credentials", width="large")
def show_case_action_modal(
    target_row_idx,
    modal_df,
    modal_sheet_name,
    is_admin_view=False,
    current_worker_name="",
    can_edit_allowed=True,
    can_delete_req_allowed=True,
    can_reschedule_allowed=True,
):
    if modal_df is None or target_row_idx not in modal_df.index:
        st.warning("ఈ case row అందుబాటులో లేదు.")
        if st.button("Close", key="dlg_missing_close"):
            close_case_popup_state()
            st.rerun()
        return

    row_data = modal_df.loc[target_row_idx]
    disp_u_col = find_display_username_column(modal_df)
    p_col = find_password_column(modal_df)
    login_u_col = find_login_user_column(modal_df)
    v_col = find_visa_type_column(modal_df)
    ds_cols = find_ds160_columns(modal_df)

    case_id_val = str(row_data.get("Case_ID", "N/A")).strip() if "Case_ID" in modal_df.columns else "N/A"
    case_user_val = str(row_data.get(disp_u_col, "")).strip() if disp_u_col in modal_df.columns else ""
    assigned_worker_val = str(row_data.get(login_u_col, "")).strip() if login_u_col in modal_df.columns else ""

    st.caption(
        f"**Sheet:** `{modal_sheet_name}` | **Case ID:** `{case_id_val}` | "
        f"**Case Username:** `{case_user_val or '-'}` | **Assigned Worker:** `{assigned_worker_val or 'Not Assigned'}`"
    )

    tab_edit, tab_resched, tab_cgi = st.tabs([
        "✏️ Edit / Update Case",
        "🔄 Reschedule Case",
        "📋 CGI Credentials Message",
    ])

    with tab_edit:
        if not is_admin_view and not can_edit_allowed:
            st.error("మీకు case edit చేసే permission లేదు.")
        else:
            with st.form(f"dlg_edit_form_{modal_sheet_name}_{target_row_idx}"):
                c1, c2, c3 = st.columns(3)
                with c1:
                    new_case_user = st.text_input(
                        "Case Username",
                        value=case_user_val,
                    )
                    new_password = st.text_input(
                        "Password",
                        value=str(row_data.get(p_col, "")).strip() if p_col in modal_df.columns else "",
                    )
                with c2:
                    worker_options = [""] + get_all_workers_list()
                    cur_w_idx = 0
                    for idx_w, w_opt in enumerate(worker_options):
                        if w_opt.lower() == assigned_worker_val.lower():
                            cur_w_idx = idx_w
                            break
                    new_worker = st.selectbox(
                        "Assigned Case Worker",
                        options=worker_options,
                        index=cur_w_idx,
                    )
                    status_opts = ["Active", "Booked", "Rescheduled", "Issue", "Completed", "On Hold"]
                    cur_st = str(row_data.get("Booking_Status", "Active")).strip() or "Active"
                    st_idx = status_opts.index(cur_st) if cur_st in status_opts else 0
                    new_booking_status = st.selectbox("Booking Status", options=status_opts, index=st_idx)
                with c3:
                    new_visa = ""
                    if v_col and v_col in modal_df.columns:
                        new_visa = st.text_input(
                            f"{v_col}",
                            value=str(row_data.get(v_col, "")).strip(),
                        )
                    new_appt_entry = st.text_input(
                        "Appointment Entry / Slot",
                        value=str(row_data.get("Appointment_Entry", "")).strip(),
                    )

                ds_inputs = {}
                if ds_cols:
                    ds_cols_ui = st.columns(min(len(ds_cols), 3))
                    for i_ds, d_col in enumerate(ds_cols):
                        with ds_cols_ui[i_ds % len(ds_cols_ui)]:
                            ds_inputs[d_col] = st.text_input(
                                f"{d_col}",
                                value=str(row_data.get(d_col, "")).strip(),
                                key=f"dlg_ds_{target_row_idx}_{i_ds}",
                            )

                c4, c5 = st.columns(2)
                with c4:
                    new_issue_remark = st.text_area(
                        "Issue / Remark",
                        value=str(row_data.get("Issue_Remark", "")).strip(),
                        height=75,
                    )
                with c5:
                    uploaded_pdf = st.file_uploader(
                        "Upload Appointment PDF (Optional)",
                        type=["pdf"],
                        key=f"dlg_pdf_{target_row_idx}",
                    )
                    existing_pdf = str(row_data.get("Appointment_PDF", "")).strip()
                    if existing_pdf:
                        st.caption(f"Current PDF: `{existing_pdf}`")

                delete_reason = ""
                request_delete_flag = False
                if not is_admin_view and can_delete_req_allowed:
                    st.markdown("---")
                    request_delete_flag = st.checkbox("🗑️ Request Admin to Delete this Case")
                    delete_reason = st.text_input("Delete Request Reason", placeholder="Reason for deletion...")

                save_btn = st.form_submit_button("💾 Save Case Changes", type="primary")

            if save_btn:
                if not str(new_case_user).strip() and not str(new_password).strip():
                    st.error("కనీసం Case Username లేదా Password enter చేయండి.")
                elif request_delete_flag and not str(delete_reason).strip():
                    st.error("Delete request పంపడానికి Reason తప్పనిసరి.")
                else:
                    actor_name = "Admin" if is_admin_view else (current_worker_name or "Worker")
                    now_str = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")

                    updated_df = load_data(sheet_name=modal_sheet_name)
                    if updated_df is None:
                        updated_df = modal_df.copy()

                    if target_row_idx in updated_df.index:
                        if disp_u_col in updated_df.columns:
                            updated_df.at[target_row_idx, disp_u_col] = str(new_case_user).strip()
                        if p_col in updated_df.columns:
                            updated_df.at[target_row_idx, p_col] = str(new_password).strip()
                        if login_u_col in updated_df.columns:
                            updated_df.at[target_row_idx, login_u_col] = str(new_worker).strip()
                        if v_col and v_col in updated_df.columns:
                            updated_df.at[target_row_idx, v_col] = str(new_visa).strip()
                        for d_col, d_val in ds_inputs.items():
                            if d_col in updated_df.columns:
                                updated_df.at[target_row_idx, d_col] = str(d_val).strip()

                        if "Booking_Status" in updated_df.columns:
                            updated_df.at[target_row_idx, "Booking_Status"] = str(new_booking_status).strip()
                        if "Appointment_Entry" in updated_df.columns:
                            updated_df.at[target_row_idx, "Appointment_Entry"] = str(new_appt_entry).strip()
                        if "Issue_Remark" in updated_df.columns:
                            updated_df.at[target_row_idx, "Issue_Remark"] = str(new_issue_remark).strip()

                        if uploaded_pdf is not None:
                            safe_cid = re.sub(r"[^A-Za-z0-9_-]", "_", case_id_val or "case")
                            pdf_dir = os.path.join(PDF_BASE_DIR, safe_cid)
                            os.makedirs(pdf_dir, exist_ok=True)
                            pdf_fname = f"{safe_cid}_{int(get_ist_now().timestamp())}.pdf"
                            pdf_path = os.path.join(pdf_dir, pdf_fname)
                            with open(pdf_path, "wb") as pf:
                                pf.write(uploaded_pdf.getbuffer())
                            if "Appointment_PDF" in updated_df.columns:
                                updated_df.at[target_row_idx, "Appointment_PDF"] = pdf_path

                        if request_delete_flag and not is_admin_view:
                            if "Delete_Reason" in updated_df.columns:
                                updated_df.at[target_row_idx, "Delete_Reason"] = str(delete_reason).strip()
                            if "Action_Type" in updated_df.columns:
                                updated_df.at[target_row_idx, "Action_Type"] = "Delete Requested"
                        elif "Action_Type" in updated_df.columns:
                            updated_df.at[target_row_idx, "Action_Type"] = "Updated"

                        if "Updated_By" in updated_df.columns:
                            updated_df.at[target_row_idx, "Updated_By"] = actor_name
                        if "Updated_At" in updated_df.columns:
                            updated_df.at[target_row_idx, "Updated_At"] = now_str

                        save_data(updated_df, sheet_name=modal_sheet_name)
                        st.success("✅ Case details update అయ్యాయి!")
                        close_case_popup_state()
                        st.rerun()

    with tab_resched:
        if not is_admin_view and not can_reschedule_allowed:
            st.error("మీకు case reschedule చేసే permission లేదు.")
        else:
            with st.form(f"dlg_resched_form_{modal_sheet_name}_{target_row_idx}"):
                r1, r2 = st.columns(2)
                with r1:
                    pref_loc = st.text_input(
                        "Preferred Location / Consular Post",
                        value=str(row_data.get("Preferred_Location", "")).strip(),
                        placeholder="e.g., Hyderabad, Chennai, Delhi",
                    )
                    pref_dates = st.text_input(
                        "Preferred Dates / Month",
                        value=str(row_data.get("Preferred_Dates", "")).strip(),
                        placeholder="e.g., Dec 2026 / Jan 2027",
                    )
                with r2:
                    prior_notice = st.text_input(
                        "Prior Notice / Notes",
                        value=str(row_data.get("Prior_Notice", "")).strip(),
                        placeholder="e.g., Need 2 days notice",
                    )
                    resched_worker_opts = ["Keep Same"] + get_all_workers_list()
                    resched_w_choice = st.selectbox(
                        "Assign Worker for New Rescheduled Case",
                        options=resched_worker_opts,
                    )

                resched_submit = st.form_submit_button("🔄 Create Rescheduled Case", type="primary")

            if resched_submit:
                actor_name = "Admin" if is_admin_view else (current_worker_name or "Worker")
                now_str = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                chosen_w = assigned_worker_val if resched_w_choice == "Keep Same" else resched_w_choice

                new_id, active_s = append_rescheduled_to_active_sheet(
                    source_row=row_data,
                    pref_loc=pref_loc,
                    pref_dates=pref_dates,
                    prior_notice=prior_notice,
                    updated_by=actor_name,
                    updated_at=now_str,
                    assigned_worker=chosen_w,
                )
                if new_id:
                    st.success(f"✅ Rescheduled case created with ID `{new_id}` in sheet `{active_s}`!")
                    close_case_popup_state()
                    st.rerun()
                else:
                    st.error("Rescheduled case create చేయడంలో error వచ్చింది.")

    with tab_cgi:
        cgi_msg = build_cgi_formatted_message(row_data)
        render_copy_message_button(cgi_msg, btn_key=f"dlg_copy_cgi_{target_row_idx}")
        st.code(cgi_msg, language="text")

    if st.button("❌ Close Modal", key=f"dlg_close_btn_{target_row_idx}"):
        close_case_popup_state()
        st.rerun()
# ================= END OF STEP 10 =================
# ================= STEP 11: SIDEBAR LOGIN, QUICK CONTROLS & NAVIGATION =================
def init_session_defaults():
    defaults = {
        "is_admin": False,
        "current_user": None,
        "selected_sheet": None,
        "is_editing_popup": False,
        "is_delete_request_mode": False,
        "is_reschedule_mode": False,
        "selected_row_idx": None,
        "open_dialog": None,
        "lead_filter": "All",
        "search_query": "",
        "active_worker_tab": "My Cases",
        "_gsheet_sync_version": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def render_sidebar():
    init_session_defaults()
    logo_path = next((p for p in LOGO_FILES if os.path.exists(p)), None)

    with st.sidebar:
        if logo_path:
            try:
                b64_logo = get_base64_image(logo_path)
                st.markdown(
                    f"""
                    <div style="text-align:center; padding: 6px 0 10px 0;">
                        <img src="data:image/jpeg;base64,{b64_logo}"
                             style="max-height:78px; border-radius:10px; box-shadow:0 2px 6px rgba(0,0,0,0.12);" />
                        <div style="font-size:15px; font-weight:800; color:#0f172a; margin-top:6px;">
                            HemanVisa Portal
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            except Exception:
                st.markdown("### ✈️ HemanVisa Portal")
        else:
            st.markdown("### ✈️ HemanVisa Portal")

        st.markdown('<hr class="compact-divider">', unsafe_allow_html=True)

        is_admin = st.session_state.get("is_admin", False)
        current_user = st.session_state.get("current_user")

        if not is_admin and not current_user:
            st.markdown("#### 🔐 Portal Login")
            login_role = st.radio(
                "Login As",
                options=["Case Worker", "Admin"],
                horizontal=True,
                key="sidebar_login_role",
            )

            if login_role == "Admin":
                with st.form("admin_login_form"):
                    adm_u = st.text_input("Admin Username", key="adm_u_inp")
                    adm_p = st.text_input("Admin Password", type="password", key="adm_p_inp")
                    adm_submit = st.form_submit_button("Login as Admin", type="primary")
                if adm_submit:
                    cfg_u, cfg_p = get_admin_creds()
                    if adm_u.strip() == cfg_u and adm_p.strip() == cfg_p:
                        st.session_state["is_admin"] = True
                        st.session_state["current_user"] = None
                        close_case_popup_state()
                        st.rerun()
                    else:
                        st.error("తప్పుడు Admin Username లేదా Password.")
            else:
                workers_all = get_all_workers_list()
                if are_all_workers_blocked():
                    st.error("⛔ అన్ని Case Worker logins తాత్కాలికంగా నిలిపివేయబడ్డాయి (Blocked by Admin).")
                else:
                    with st.form("worker_login_form"):
                        w_choice = st.selectbox(
                            "Select Case Worker",
                            options=[""] + workers_all,
                            key="w_login_select",
                        )
                        w_submit = st.form_submit_button("Login to Workspace", type="primary")
                    if w_submit:
                        w_clean = str(w_choice).strip()
                        if not w_clean:
                            st.warning("దయచేసి మీ పేరును ఎంచుకోండి.")
                        elif is_worker_blocked(w_clean):
                            st.error(f"⛔ `{w_clean}` login block చేయబడింది. Admin ని సంప్రదించండి.")
                        else:
                            st.session_state["current_user"] = w_clean
                            st.session_state["is_admin"] = False
                            close_case_popup_state()
                            record_worker_login(w_clean)
                            st.rerun()
        else:
            if is_admin:
                st.success("🛡️ Logged in as: **Admin**")
            else:
                st.success(f"👤 Logged in as: **{current_user}**")

            sheets_list = get_sheet_names()
            if not sheets_list:
                sheets_list = ["Sheet1"]
            default_sheet = get_default_active_sheet(sheets_list)
            if (
                st.session_state.get("selected_sheet") not in sheets_list
            ):
                st.session_state["selected_sheet"] = default_sheet

            chosen_sheet = st.selectbox(
                "📄 Active Sheet",
                options=sheets_list,
                index=sheets_list.index(st.session_state["selected_sheet"]),
                key="sidebar_sheet_selector",
            )
            if chosen_sheet != st.session_state["selected_sheet"]:
                st.session_state["selected_sheet"] = chosen_sheet
                close_case_popup_state()
                st.rerun()

            can_sync_now = True
            if not is_admin:
                perms_all = load_permissions()
                u_perm = next(
                    (v for k, v in perms_all.items() if k.lower() == str(current_user).lower()),
                    {},
                )
                can_sync_now = bool(u_perm.get("can_sync", False))

            if can_sync_now:
                if st.button("🔄 Sync Google Sheets", key="sidebar_sync_btn", use_container_width=True):
                    _clear_sheet_runtime_caches()
                    st.session_state["_gsheet_sync_version"] = (
                        st.session_state.get("_gsheet_sync_version", 0) + 1
                    )
                    st.session_state.pop("_last_id_sync_mtime", None)
                    st.rerun()

            if is_admin:
                try:
                    exp_bytes = get_admin_gsheet_export_bytes(_get_file_mtime())
                    st.download_button(
                        label="📥 Download All Sheets (.xlsx)",
                        data=exp_bytes,
                        file_name=f"HemanVisa_All_Sheets_{get_ist_now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="sidebar_admin_dl_btn",
                        use_container_width=True,
                    )
                except Exception:
                    pass
            else:
                perms_all = load_permissions()
                u_perm = next(
                    (v for k, v in perms_all.items() if k.lower() == str(current_user).lower()),
                    {},
                )
                if u_perm.get("can_download", False):
                    try:
                        w_bytes = get_worker_export_bytes(
                            _get_file_mtime(),
                            str(current_user),
                            st.session_state["selected_sheet"],
                        )
                        st.download_button(
                            label="📥 Download My Cases (.xlsx)",
                            data=w_bytes,
                            file_name=f"{current_user}_Cases_{get_ist_now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="sidebar_worker_dl_btn",
                            use_container_width=True,
                        )
                    except Exception:
                        pass

            st.markdown('<hr class="compact-divider">', unsafe_allow_html=True)

            if st.button("🚪 Logout", key="sidebar_logout_btn", type="secondary", use_container_width=True):
                if not is_admin and current_user:
                    record_worker_logout(str(current_user))
                    w_low = str(current_user).strip().lower()
                    for k_sess in list(st.session_state.keys()):
                        if str(k_sess).startswith(f"_login_sync_done_{w_low}_"):
                            st.session_state.pop(k_sess, None)
                st.session_state["is_admin"] = False
                st.session_state["current_user"] = None
                close_case_popup_state()
                st.rerun()

        if not get_webapp_url():
            st.caption("⚠️ `GSHEET_WEBAPP_URL` secret set చేయబడలేదు.")
# ================= END OF STEP 11 =================
