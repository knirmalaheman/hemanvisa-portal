# ================= PART 1: IMPORTS, CONSTANTS, CGI FORMATTER & WEBAPP BASE =================
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
    # ================= PART 2: EDIT REQUESTS, ADMIN CREDS, NOTICE & PERMISSIONS =================
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
```[cite: 2]

---

**Part 2** paste chesi, tarvata **"next"** ani cheppandi — **Part 3** (IST Date/Time Parsing, Login Tracking, Attendance & Summary Tables) istaanu[cite: 2].
    # ================= PART 3: IST PARSING, LOGIN TRACKING, ATTENDANCE & SUMMARY =================
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
    # ================= PART 4: COLUMN REORDERING, CASE IDS, SHEET NORMALIZATION & EXPORT =================
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
    # ================= PART 5: DATA LOAD/SAVE, SEARCH, WORKER BLOCKING & POPUP NAV =================
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
    # ================= PART 6: SESSION STATE, CSS, TOP BANNER & CLICKABLE TABLE =================
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
    if st.session_state.pop("_just_logged_in", False):
        pass
    elif is_worker_blocked(curr_w):
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
    # ================= PART 7: VIEW/EDIT CASE DIALOG & ADMIN DIALOGS =================
@st.dialog("Case Details & Actions", width="large")
def view_and_edit_case_dialog():
    enable_right_cursor_navigation_in_popup()
    row_idx = st.session_state.get("selected_row_idx")
    if row_idx is None:
        st.session_state["open_dialog"] = None
        close_case_popup_state()
        st.rerun()
        return

    all_sheets_map = load_all_sheets_cached(_get_file_mtime())
    all_sheets_map, _ = _assign_case_username_based_ids(all_sheets_map)
    visible_sheets = [
        s for s in all_sheets_map.keys() if str(s).strip() not in SYSTEM_SHEETS
    ]
    if not visible_sheets:
        visible_sheets = list(all_sheets_map.keys())

    target_sheet = st.session_state.get(
        "active_sheet", get_default_active_sheet(visible_sheets)
    )
    if target_sheet not in all_sheets_map:
        target_sheet = get_default_active_sheet(visible_sheets)

    df = all_sheets_map.get(target_sheet)
    if df is None or df.empty or row_idx not in df.index:
        found_sheet = None
        for s_k, df_k in all_sheets_map.items():
            if (
                str(s_k).strip() not in SYSTEM_SHEETS
                and df_k is not None
                and not df_k.empty
                and row_idx in df_k.index
            ):
                found_sheet = s_k
                df = df_k
                break
        if found_sheet is not None:
            target_sheet = found_sheet
            st.session_state["active_sheet"] = target_sheet
        else:
            st.warning("Case record not found or already updated.")
            if st.button("Close", key="close_missing_case_btn"):
                st.session_state["open_dialog"] = None
                close_case_popup_state()
                st.rerun()
            return

    df = reorder_dataframe_columns(df.copy())
    row_data = df.loc[row_idx].copy()
    display_c = find_display_username_column(df)
    login_c = find_login_user_column(df)
    p_col = find_password_column(df)
    all_workers = get_all_workers_list()

    curr_role = st.session_state.get("role", "")
    curr_user = st.session_state.get("username", "")
    perms = load_permissions()
    u_perms = {}
    for k, v in perms.items():
        if k.lower() == curr_user.strip().lower():
            u_perms = v
            break

    can_edit_perm = (
        True
        if curr_role == "admin"
        else u_perms.get("can_edit", u_perms.get("can_update", True))
    )
    can_delete_perm = (
        True if curr_role == "admin" else u_perms.get("can_delete_req", True)
    )
    can_reschedule_perm = (
        True if curr_role == "admin" else u_perms.get("can_reschedule", True)
    )

    row_assigned_worker = (
        str(row_data.get(login_c, "")).strip() if login_c in row_data.index else ""
    )
    case_id_val = (
        str(row_data.get("Case_ID", "")).strip() if "Case_ID" in row_data.index else ""
    )
    case_user_val = (
        str(row_data.get(display_c, "")).strip() if display_c in row_data.index else ""
    )

    is_other_worker_case = (
        curr_role == "worker"
        and row_assigned_worker.lower() != curr_user.strip().lower()
    )
    existing_edit_req = (
        get_worker_case_edit_request(
            curr_user, target_sheet, case_id_val, case_user_val
        )
        if is_other_worker_case
        else None
    )
    edit_req_status = (
        str(existing_edit_req.get("Status", "")).strip().lower()
        if existing_edit_req
        else ""
    )

    worker_can_edit_this_case = True
    if is_other_worker_case:
        worker_can_edit_this_case = edit_req_status == "approved"

    case_title = (
        str(row_data.get(display_c, "")).strip()
        if display_c in row_data.index and str(row_data.get(display_c, "")).strip()
        else (case_id_val if case_id_val else f"Row #{row_idx}")
    )

    st.markdown(
        f"""
        <div style="background: linear-gradient(135deg, #0f1c3f 0%, #1e3a8a 100%); color: #ffffff; padding: 12px 16px; border-radius: 10px; border-left: 5px solid #d4af37; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px;">
            <div>
                <div style="font-size: 16px; font-weight: 800;">📄 Case: {case_title}</div>
                <div style="font-size: 12px; color: #fde047; font-weight: 600;">Sheet: {target_sheet} &nbsp;|&nbsp; Case ID: {case_id_val or 'N/A'} &nbsp;|&nbsp; Assigned Worker: {row_assigned_worker or 'Not Assigned'}</div>
            </div>
            <div style="background: rgba(255,255,255,0.14); padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 700;">
                Status: {str(row_data.get('Booking_Status', 'Active')) or 'Active'}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    cgi_msg_text = build_cgi_formatted_message(row_data)
    render_copy_message_button(cgi_msg_text, btn_key=f"copy_cgi_btn_{target_sheet}_{row_idx}")

    with st.expander("📩 View Formatted CGI Credentials Message", expanded=False):
        st.code(cgi_msg_text, language="text")

    if not st.session_state.get("is_editing_popup", False) and not st.session_state.get(
        "is_delete_request_mode", False
    ) and not st.session_state.get("is_reschedule_mode", False):
        col_a, col_b, col_c, col_d = st.columns(4, gap="small")
        with col_a:
            edit_btn_label = "✏️ Edit Case"
            if is_other_worker_case and edit_req_status != "approved":
                edit_btn_label = "🔐 Request Edit"
            if st.button(
                edit_btn_label,
                key="popup_btn_edit",
                use_container_width=True,
                disabled=(not can_edit_perm),
            ):
                if is_other_worker_case and edit_req_status != "approved":
                    st.session_state["show_edit_req_form"] = True
                else:
                    st.session_state["is_editing_popup"] = True
                st.rerun()
        with col_b:
            if st.button(
                "📅 Reschedule Case",
                key="popup_btn_reschedule",
                use_container_width=True,
                disabled=(not can_reschedule_perm),
            ):
                st.session_state["is_reschedule_mode"] = True
                st.rerun()
        with col_c:
            del_label = (
                "🗑️ Delete Case"
                if curr_role == "admin"
                else "🗑️ Request Delete"
            )
            if st.button(
                del_label,
                key="popup_btn_delete",
                use_container_width=True,
                disabled=(not can_delete_perm),
            ):
                st.session_state["is_delete_request_mode"] = True
                st.rerun()
        with col_d:
            if st.button("❌ Close", key="popup_btn_close", use_container_width=True):
                st.session_state["open_dialog"] = None
                close_case_popup_state()
                st.rerun()

        if is_other_worker_case and edit_req_status != "approved":
            if edit_req_status == "pending":
                st.info(
                    f"⏳ Edit request is **Pending** admin approval (Reason: {existing_edit_req.get('Reason', '-')})."
                )
            elif edit_req_status == "rejected":
                st.warning(
                    f"❌ Previous edit request was **Rejected** (Reason: {existing_edit_req.get('Reason', '-')}). You can submit a new request below."
                )

            if st.session_state.get("show_edit_req_form", False) or edit_req_status in [
                "",
                "rejected",
            ]:
                with st.form(f"worker_edit_req_form_{target_sheet}_{row_idx}"):
                    st.markdown(
                        f"**Request Admin Permission to Edit Another Worker's Case ({row_assigned_worker or 'Unassigned'})**"
                    )
                    req_reason = st.text_input(
                        "Why do you need to edit this case? *",
                        value="",
                        placeholder="Enter clear reason for editing this case...",
                    )
                    sub_req_col1, sub_req_col2 = st.columns(2)
                    with sub_req_col1:
                        submit_er = st.form_submit_button(
                            "📤 Send Edit Request", use_container_width=True
                        )
                    with sub_req_col2:
                        cancel_er = st.form_submit_button(
                            "Cancel", use_container_width=True
                        )
                    if submit_er:
                        if not req_reason.strip():
                            st.error("Please enter a reason for the edit request.")
                        else:
                            create_or_reopen_edit_request(
                                worker_name=curr_user,
                                sheet_name=target_sheet,
                                case_id=case_id_val,
                                case_username=case_user_val,
                                row_idx=row_idx,
                                reason=req_reason.strip(),
                            )
                            st.session_state["show_edit_req_form"] = False
                            st.success("✅ Edit request sent to Admin!")
                            st.rerun()
                    if cancel_er:
                        st.session_state["show_edit_req_form"] = False
                        st.rerun()

        pdf_rel = (
            str(row_data.get("Appointment_PDF", "")).strip()
            if "Appointment_PDF" in row_data.index
            else ""
        )
        if pdf_rel and os.path.exists(pdf_rel):
            with open(pdf_rel, "rb") as pf:
                st.download_button(
                    label=f"📎 Download Appointment PDF ({os.path.basename(pdf_rel)})",
                    data=pf.read(),
                    file_name=os.path.basename(pdf_rel),
                    mime="application/pdf",
                    key=f"dl_pdf_{target_sheet}_{row_idx}",
                )

        detail_records = []
        for col in df.columns:
            if col == "Appointment_PDF":
                continue
            val_str = str(row_data.get(col, "")).strip()
            detail_records.append(
                {"Field": col, "Value": val_str if val_str else "-"}
            )
        detail_df = pd.DataFrame(detail_records)
        st.dataframe(detail_df, use_container_width=True, hide_index=True, height=340)

    elif st.session_state.get("is_editing_popup", False):
        if not can_edit_perm or not worker_can_edit_this_case:
            st.error("You do not have permission to edit this case.")
            if st.button("Back", key="back_no_edit_perm"):
                st.session_state["is_editing_popup"] = False
                st.rerun()
            return

        st.markdown("### ✏️ Edit Case Fields")
        with st.form(f"edit_case_form_{target_sheet}_{row_idx}"):
            edited_values = {}
            editable_cols = [c for c in df.columns if c != "Appointment_PDF"]
            cols_per_row = 2
            for start_i in range(0, len(editable_cols), cols_per_row):
                row_cols = st.columns(cols_per_row, gap="medium")
                for offset_i in range(cols_per_row):
                    if start_i + offset_i >= len(editable_cols):
                        break
                    c_name = editable_cols[start_i + offset_i]
                    cur_val = str(row_data.get(c_name, "")).strip()
                    with row_cols[offset_i]:
                        if c_name == "Case_ID":
                            st.text_input(
                                f"{c_name} (Auto ID)",
                                value=cur_val,
                                disabled=True,
                                key=f"ed_cid_{target_sheet}_{row_idx}_{c_name}",
                            )
                            edited_values[c_name] = cur_val
                        elif c_name == login_c and curr_role == "admin":
                            w_opts = ["Not Assigned"] + [
                                w for w in all_workers if w.lower() != "not assigned"
                            ]
                            if cur_val and cur_val not in w_opts:
                                w_opts.append(cur_val)
                            sel_w_idx = (
                                w_opts.index(cur_val) if cur_val in w_opts else 0
                            )
                            edited_values[c_name] = st.selectbox(
                                f"{c_name} (Assigned Worker)",
                                options=w_opts,
                                index=sel_w_idx,
                                key=f"ed_w_{target_sheet}_{row_idx}_{c_name}",
                            )
                        elif c_name == "Booking_Status":
                            b_opts = ["Active", "Booked", "Rescheduled", "Issues"]
                            b_idx = 0
                            for idx_b, b_opt in enumerate(b_opts):
                                if b_opt.lower() == cur_val.lower():
                                    b_idx = idx_b
                                    break
                            edited_values[c_name] = st.selectbox(
                                c_name,
                                options=b_opts,
                                index=b_idx,
                                key=f"ed_bs_{target_sheet}_{row_idx}_{c_name}",
                            )
                        else:
                            edited_values[c_name] = st.text_input(
                                c_name,
                                value=cur_val,
                                key=f"ed_txt_{target_sheet}_{row_idx}_{c_name}",
                            )

            uploaded_pdf = st.file_uploader(
                "Upload Appointment PDF (Optional)",
                type=["pdf"],
                key=f"ed_pdf_{target_sheet}_{row_idx}",
            )

            save_c1, save_c2 = st.columns(2, gap="medium")
            with save_c1:
                save_clicked = st.form_submit_button(
                    "💾 Save Changes", use_container_width=True
                )
            with save_c2:
                cancel_clicked = st.form_submit_button(
                    "❌ Cancel Edit", use_container_width=True
                )

            if save_clicked:
                now_ist_str = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                actor_name = curr_user if curr_user else "Admin"
                for c_name, new_v in edited_values.items():
                    if c_name == login_c and str(new_v).strip().lower() == "not assigned":
                        df.at[row_idx, c_name] = ""
                    else:
                        df.at[row_idx, c_name] = str(new_v).strip()

                if uploaded_pdf is not None:
                    safe_cid = re.sub(r"[^A-Za-z0-9_-]", "_", case_id_val or f"row_{row_idx}")
                    pdf_folder = os.path.join(PDF_BASE_DIR, safe_cid)
                    os.makedirs(pdf_folder, exist_ok=True)
                    pdf_save_path = os.path.join(pdf_folder, uploaded_pdf.name)
                    with open(pdf_save_path, "wb") as f_out:
                        f_out.write(uploaded_pdf.getbuffer())
                    df.at[row_idx, "Appointment_PDF"] = pdf_save_path

                df.at[row_idx, "Updated_By"] = actor_name
                df.at[row_idx, "Updated_At"] = now_ist_str
                df.at[row_idx, "Action_Type"] = "Updated"

                save_data(df, sheet_name=target_sheet)
                st.session_state["is_editing_popup"] = False
                st.session_state["open_dialog"] = None
                close_case_popup_state()
                st.success("✅ Case updated successfully!")
                st.rerun()

            if cancel_clicked:
                st.session_state["is_editing_popup"] = False
                st.rerun()

    elif st.session_state.get("is_reschedule_mode", False):
        if not can_reschedule_perm:
            st.error("You do not have permission to reschedule cases.")
            if st.button("Back", key="back_no_resched_perm"):
                st.session_state["is_reschedule_mode"] = False
                st.rerun()
            return

        st.markdown("### 📅 Reschedule Case")
        with st.form(f"reschedule_form_{target_sheet}_{row_idx}"):
            pref_loc = st.text_input(
                "Preferred Location *",
                value=str(row_data.get("Preferred_Location", "")).strip(),
            )
            pref_dates = st.text_input(
                "Preferred Dates *",
                value=str(row_data.get("Preferred_Dates", "")).strip(),
            )
            prior_notice = st.text_input(
                "Prior Notice / Remarks",
                value=str(row_data.get("Prior_Notice", "")).strip(),
            )
            assign_worker_opt = None
            if curr_role == "admin":
                w_opts_r = ["Keep Current"] + [
                    w for w in all_workers if w.lower() != "not assigned"
                ]
                assign_worker_opt = st.selectbox(
                    "Assign Worker for Rescheduled Case",
                    options=w_opts_r,
                    index=0,
                )

            rs_c1, rs_c2 = st.columns(2, gap="medium")
            with rs_c1:
                confirm_rs = st.form_submit_button(
                    "✅ Confirm Reschedule", use_container_width=True
                )
            with rs_c2:
                cancel_rs = st.form_submit_button(
                    "❌ Cancel", use_container_width=True
                )

            if confirm_rs:
                if not pref_loc.strip() or not pref_dates.strip():
                    st.error("Preferred Location and Preferred Dates are required.")
                else:
                    now_ist_str = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    actor_name = curr_user if curr_user else "Admin"
                    df.at[row_idx, "Booking_Status"] = "Rescheduled"
                    df.at[row_idx, "Preferred_Location"] = pref_loc.strip()
                    df.at[row_idx, "Preferred_Dates"] = pref_dates.strip()
                    df.at[row_idx, "Prior_Notice"] = prior_notice.strip()
                    df.at[row_idx, "Updated_By"] = actor_name
                    df.at[row_idx, "Updated_At"] = now_ist_str
                    df.at[row_idx, "Action_Type"] = "Rescheduled"
                    save_data(df, sheet_name=target_sheet)

                    w_for_new = (
                        row_assigned_worker
                        if (curr_role != "admin" or assign_worker_opt == "Keep Current")
                        else assign_worker_opt
                    )
                    new_cid, act_s = append_rescheduled_to_active_sheet(
                        source_row=df.loc[row_idx],
                        pref_loc=pref_loc,
                        pref_dates=pref_dates,
                        prior_notice=prior_notice,
                        updated_by=actor_name,
                        updated_at=now_ist_str,
                        assigned_worker=w_for_new,
                    )
                    st.session_state["is_reschedule_mode"] = False
                    st.session_state["open_dialog"] = None
                    close_case_popup_state()
                    st.success(
                        f"✅ Rescheduled! New Case ID {new_cid or ''} added to {act_s or 'Active Sheet'}."
                    )
                    st.rerun()

            if cancel_rs:
                st.session_state["is_reschedule_mode"] = False
                st.rerun()

    elif st.session_state.get("is_delete_request_mode", False):
        if not can_delete_perm:
            st.error("You do not have permission to delete cases.")
            if st.button("Back", key="back_no_del_perm"):
                st.session_state["is_delete_request_mode"] = False
                st.rerun()
            return

        st.markdown(
            "### 🗑️ Delete Case"
            if curr_role == "admin"
            else "### 🗑️ Request Case Deletion"
        )
        with st.form(f"delete_case_form_{target_sheet}_{row_idx}"):
            del_reason = st.text_area(
                "Reason for Deletion *",
                placeholder="Enter reason for deleting this case...",
            )
            dl_c1, dl_c2 = st.columns(2, gap="medium")
            with dl_c1:
                confirm_dl = st.form_submit_button(
                    "🗑️ Delete Permanently"
                    if curr_role == "admin"
                    else "📤 Send Delete Request",
                    use_container_width=True,
                )
            with dl_c2:
                cancel_dl = st.form_submit_button(
                    "❌ Cancel", use_container_width=True
                )

            if confirm_dl:
                if not del_reason.strip():
                    st.error("Please provide a deletion reason.")
                else:
                    now_ist_str = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    actor_name = curr_user if curr_user else "Admin"
                    if curr_role == "admin":
                        df_after_del = df.drop(index=[row_idx]).reset_index(drop=True)
                        save_data(df_after_del, sheet_name=target_sheet)
                        st.session_state["is_delete_request_mode"] = False
                        st.session_state["open_dialog"] = None
                        close_case_popup_state()
                        st.success("✅ Case deleted successfully!")
                        st.rerun()
                    else:
                        df.at[row_idx, "Action_Type"] = "Delete Requested"
                        df.at[row_idx, "Delete_Reason"] = del_reason.strip()
                        df.at[row_idx, "Updated_By"] = actor_name
                        df.at[row_idx, "Updated_At"] = now_ist_str
                        save_data(df, sheet_name=target_sheet)
                        st.session_state["is_delete_request_mode"] = False
                        st.session_state["open_dialog"] = None
                        close_case_popup_state()
                        st.success("✅ Deletion request submitted to Admin!")
                        st.rerun()

            if cancel_dl:
                st.session_state["is_delete_request_mode"] = False
                st.rerun()


@st.dialog("Admin Control Center", width="large")
def open_admin_modal():
    dialog_mode = st.session_state.get("open_dialog")
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    visible_sheets = [
        s for s in all_sheets.keys() if str(s).strip() not in SYSTEM_SHEETS
    ]
    if not visible_sheets:
        visible_sheets = list(all_sheets.keys()) or ["Sheet1"]
    active_s = st.session_state.get(
        "active_sheet", get_default_active_sheet(visible_sheets)
    )
    if active_s not in all_sheets:
        active_s = visible_sheets[0]
    df_s = all_sheets.get(active_s, pd.DataFrame())
    all_workers = get_all_workers_list()

    if dialog_mode == "add_case":
        st.markdown(f"### ➕ Add New Case to **{active_s}**")
        next_cid = generate_next_case_id()
        login_c = find_login_user_column(df_s) if not df_s.empty else "Case Worker"
        disp_c = find_display_username_column(df_s) if not df_s.empty else "USERNAME"
        p_c = find_password_column(df_s) if not df_s.empty else "Password"

        with st.form("admin_add_case_form"):
            st.text_input("Auto-Generated Case ID", value=next_cid, disabled=True)
            c1, c2 = st.columns(2)
            with c1:
                w_opts = ["Not Assigned"] + [
                    w for w in all_workers if w.lower() != "not assigned"
                ]
                sel_worker = st.selectbox("Assign Case Worker", options=w_opts, index=0)
                u_val = st.text_input(f"{disp_c} *", placeholder="CGI / Portal Username")
            with c2:
                pw_val = st.text_input(f"{p_c}", placeholder="Password")
                lead_val = st.text_input("Lead", placeholder="Kiran / Shanthi / Lead Name")

            extra_cols = [
                c
                for c in df_s.columns
                if c
                not in [
                    "Case_ID",
                    login_c,
                    disp_c,
                    p_c,
                    "Appointment_PDF",
                    "Updated_By",
                    "Updated_At",
                    "Action_Type",
                    "Delete_Reason",
                ]
                and "lead" not in c.strip().lower()
            ]
            extra_inputs = {}
            if extra_cols:
                with st.expander("Additional Fields", expanded=False):
                    for ec in extra_cols[:12]:
                        extra_inputs[ec] = st.text_input(ec, key=f"add_ec_{ec}")

            if st.form_submit_button("✅ Create Case", use_container_width=True):
                if not u_val.strip():
                    st.error("Username is required.")
                else:
                    now_str = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    new_row = {col: "" for col in df_s.columns}
                    new_row["Case_ID"] = next_cid
                    new_row[disp_c] = u_val.strip()
                    if p_c in new_row:
                        new_row[p_c] = pw_val.strip()
                    if login_c in new_row:
                        new_row[login_c] = (
                            "" if sel_worker == "Not Assigned" else sel_worker
                        )
                    for col in df_s.columns:
                        if "lead" in col.strip().lower():
                            new_row[col] = lead_val.strip()
                    for ec, ev in extra_inputs.items():
                        new_row[ec] = ev.strip()
                    new_row["Booking_Status"] = "Active"
                    new_row["Updated_By"] = st.session_state.get("username", "admin")
                    new_row["Updated_At"] = now_str
                    new_row["Action_Type"] = "Added"

                    df_updated = pd.concat(
                        [df_s, pd.DataFrame([new_row])], ignore_index=True
                    )
                    save_data(df_updated, sheet_name=active_s)
                    st.session_state["open_dialog"] = None
                    st.success(f"✅ Case {next_cid} added to {active_s}!")
                    st.rerun()

    elif dialog_mode == "worker_permissions":
        st.markdown("### 🔐 Worker Permissions & Login Control")
        perms = dict(load_permissions())
        selected_w = st.selectbox(
            "Select Worker",
            options=all_workers if all_workers else ["Keerthana", "Sai", "Kalyan", "Anil"],
            key="perm_worker_select",
        )
        cur_p = perms.get(
            selected_w,
            {
                "can_edit": True,
                "can_delete_req": True,
                "can_reschedule": True,
                "can_add": True,
                "can_update": True,
                "can_search": True,
                "can_download": False,
                "can_sync": False,
                "block_login": False,
            },
        )
        with st.form("worker_perms_form"):
            p1, p2 = st.columns(2)
            with p1:
                c_edit = st.checkbox("Allow Edit / Update Case", value=bool(cur_p.get("can_edit", True)))
                c_del = st.checkbox("Allow Delete Request", value=bool(cur_p.get("can_delete_req", True)))
                c_res = st.checkbox("Allow Reschedule Case", value=bool(cur_p.get("can_reschedule", True)))
                c_add = st.checkbox("Allow Add Case", value=bool(cur_p.get("can_add", True)))
            with p2:
                c_srch = st.checkbox("Allow Search", value=bool(cur_p.get("can_search", True)))
                c_dl = st.checkbox("Allow Excel Download", value=bool(cur_p.get("can_download", False)))
                c_sync = st.checkbox("Allow Sheet Sync", value=bool(cur_p.get("can_sync", False)))
                c_blk = st.checkbox("🚫 Block Worker Login", value=bool(cur_p.get("block_login", False)))

            if st.form_submit_button("💾 Save Permissions", use_container_width=True):
                perms[selected_w] = {
                    "can_edit": c_edit,
                    "can_delete_req": c_del,
                    "can_reschedule": c_res,
                    "can_add": c_add,
                    "can_update": c_edit,
                    "can_search": c_srch,
                    "can_download": c_dl,
                    "can_sync": c_sync,
                    "block_login": c_blk,
                }
                save_permissions(perms)
                st.session_state["open_dialog"] = None
                st.success(f"✅ Permissions saved for {selected_w}!")
                st.rerun()

    elif dialog_mode == "edit_requests":
        st.markdown("### 📩 Worker Edit Requests Approval")
        reqs = load_edit_requests()
        if not reqs:
            st.info("No edit requests found.")
        else:
            for idx_r, req_item in enumerate(reqs):
                st.markdown(
                    f"**#{idx_r + 1} | Worker:** `{req_item.get('Worker')}` | **Sheet:** `{req_item.get('Sheet_Name')}` | **Case ID:** `{req_item.get('Case_ID')}` | **Status:** `{req_item.get('Status')}`"
                )
                st.caption(f"Reason: {req_item.get('Reason', '-')} | Requested: {req_item.get('Requested_At', '-')}")
                b1, b2, _ = st.columns([1.2, 1.2, 3.6])
                with b1:
                    if st.button("✅ Approve", key=f"apv_er_{idx_r}"):
                        reqs[idx_r]["Status"] = "Approved"
                        reqs[idx_r]["Reviewed_By"] = st.session_state.get("username", "admin")
                        reqs[idx_r]["Reviewed_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        save_edit_requests(reqs)
                        st.rerun()
                with b2:
                    if st.button("❌ Reject", key=f"rej_er_{idx_r}"):
                        reqs[idx_r]["Status"] = "Rejected"
                        reqs[idx_r]["Reviewed_By"] = st.session_state.get("username", "admin")
                        reqs[idx_r]["Reviewed_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        save_edit_requests(reqs)
                        st.rerun()
                st.divider()

    elif dialog_mode == "delete_requests":
        st.markdown("### 🗑️ Pending Delete Requests")
        found_del = False
        for s_name, df_item in all_sheets.items():
            if str(s_name).strip() in SYSTEM_SHEETS or df_item is None or df_item.empty:
                continue
            if "Action_Type" not in df_item.columns:
                continue
            pending_rows = df_item[
                df_item["Action_Type"].astype(str).str.strip().str.lower() == "delete requested"
            ]
            for r_i, r_row in pending_rows.iterrows():
                found_del = True
                cid_d = str(r_row.get("Case_ID", "")).strip()
                by_d = str(r_row.get("Updated_By", "")).strip()
                rsn_d = str(r_row.get("Delete_Reason", "")).strip()
                st.markdown(f"**Sheet:** `{s_name}` | **Case ID:** `{cid_d}` | **By:** `{by_d}`")
                st.caption(f"Reason: {rsn_d or '-'}")
                d1, d2, _ = st.columns([1.3, 1.3, 3.4])
                with d1:
                    if st.button("🗑️ Approve Delete", key=f"app_del_{s_name}_{r_i}"):
                        df_new_d = df_item.drop(index=[r_i]).reset_index(drop=True)
                        save_data(df_new_d, sheet_name=s_name)
                        st.rerun()
                with d2:
                    if st.button("↩️ Reject Delete", key=f"rej_del_{s_name}_{r_i}"):
                        df_item.at[r_i, "Action_Type"] = "Delete Rejected"
                        df_item.at[r_i, "Updated_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        save_data(df_item, sheet_name=s_name)
                        st.rerun()
                st.divider()
        if not found_del:
            st.info("No pending delete requests.")

    elif dialog_mode == "login_attendance":
        st.markdown("### 🕒 Worker Login Status & Attendance Summary")
        logins_map = load_login_tracking()
        att_list = load_attendance()

        if logins_map:
            st.markdown("#### Current Login Status")
            st.dataframe(
                _build_login_tracking_df(logins_map),
                use_container_width=True,
                hide_index=True,
            )

        df_day, df_overall = build_attendance_summary_tables(att_list)
        if not df_overall.empty:
            st.markdown("#### Overall Worker Attendance Summary")
            st.dataframe(df_overall, use_container_width=True)
        if not df_day.empty:
            st.markdown("#### Daily Attendance Sessions")
            st.dataframe(df_day, use_container_width=True)
        if not logins_map and df_day.empty and df_overall.empty:
            st.info("No login or attendance records found yet.")

    elif dialog_mode == "notice_admin_creds":
        st.markdown("### 📢 Notice Board & Admin Credentials")
        cur_notice = get_admin_notice()
        cur_u, cur_p = get_admin_creds()

        with st.form("admin_notice_creds_form"):
            new_notice = st.text_area(
                "Scrolling Notice / Announcement",
                value=cur_notice,
                placeholder="Enter announcement for workers...",
            )
            c1, c2 = st.columns(2)
            with c1:
                new_u = st.text_input("Admin Username", value=cur_u)
            with c2:
                new_p = st.text_input("Admin Password", value=cur_p, type="password")

            if st.form_submit_button("💾 Save Notice & Credentials", use_container_width=True):
                save_admin_notice(new_notice)
                if new_u.strip() and new_p.strip():
                    save_admin_creds(new_u.strip(), new_p.strip())
                st.session_state["open_dialog"] = None
                st.success("✅ Notice & Admin credentials saved!")
                st.rerun()

    if st.button("Close Window", key="close_admin_dialog_btn"):
        st.session_state["open_dialog"] = None
        st.rerun()
    # ================= PART 8: HOME PAGE, LOGIN PAGE & WORKER PORTAL =================
def show_home_page():
    st.markdown(
        """
        <div class="hero-section">
            <div class="hero-title">✈️ HemanVisa — Global Immigration & Appointment Portal</div>
            <div class="hero-subtitle">
                Streamlined case tracking, CGI appointment management, and real-time team coordination
                for student, work, and visitor visa workflows.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3, gap="medium")
    with c1:
        st.markdown(
            """
            <div class="corporate-card">
                <h3>📋 Case & Lead Tracking</h3>
                <p>Track assigned cases, lead categories, and appointment statuses in real time with full audit logs.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            """
            <div class="corporate-card">
                <h3>📅 Booking & Rescheduling</h3>
                <p>Manage booked appointments, upload confirmation PDFs, and seamlessly move rescheduled cases into active queues.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            """
            <div class="corporate-card">
                <h3>🔐 Role-Based Access</h3>
                <p>Dedicated Admin Control Center and Case Worker Workspace with permissions, login tracking, and instant reports.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )


def show_login_page():
    if st.session_state.get("force_logged_out_msg"):
        st.error("🚫 Your worker login is currently blocked by Admin or all worker logins are paused.")
        st.session_state["force_logged_out_msg"] = False

    if are_all_workers_blocked():
        st.warning("⚠️ All case worker logins are currently paused by Admin.")

    st.markdown("### 🔐 Portal Login")
    login_tab = st.radio(
        "Select Portal",
        options=["Case Worker Login", "Admin Login"],
        horizontal=True,
        key="login_portal_radio",
    )

    all_workers = get_all_workers_list()
    worker_options = all_workers if all_workers else ["Keerthana", "Sai", "Kalyan", "Anil"]

    if login_tab == "Case Worker Login":
        with st.form("worker_login_form"):
            w_name = st.selectbox("Select Case Worker", options=worker_options, key="login_worker_select")
            w_pass = st.text_input("Access Code / Password", type="password", key="login_worker_pass")
            submitted_w = st.form_submit_button("🚀 Login to Worker Workspace", use_container_width=True)

            if submitted_w:
                if is_worker_blocked(w_name):
                    st.error(f"🚫 Login blocked for {w_name}. Please contact Admin.")
                else:
                    clean_pass = w_pass.strip()
                    valid_worker_login = bool(clean_pass) and (
                        clean_pass.lower() == w_name.strip().lower()
                        or clean_pass.lower() in ["123", "1234", "worker", "heman", "heman123", "pass"]
                    )
                    if not valid_worker_login:
                        st.error("Invalid worker access code. Enter your worker name or valid access code.")
                    else:
                        st.session_state["logged_in"] = True
                        st.session_state["username"] = w_name.strip()
                        st.session_state["role"] = "worker"
                        st.session_state["_just_logged_in"] = True
                        record_worker_login(w_name.strip())
                        st.session_state["attendance_recorded"] = True
                        st.query_params["logged_in"] = "true"
                        st.query_params["username"] = w_name.strip()
                        st.query_params["role"] = "worker"
                        st.rerun()
    else:
        with st.form("admin_login_form"):
            a_user = st.text_input("Admin Username", key="login_admin_user")
            a_pass = st.text_input("Admin Password", type="password", key="login_admin_pass")
            submitted_a = st.form_submit_button("🛡️ Login to Admin Center", use_container_width=True)

            if submitted_a:
                expected_u, expected_p = get_admin_creds()
                if a_user.strip() == expected_u and a_pass.strip() == expected_p:
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = expected_u
                    st.session_state["role"] = "admin"
                    st.query_params["logged_in"] = "true"
                    st.query_params["username"] = expected_u
                    st.query_params["role"] = "admin"
                    st.rerun()
                else:
                    st.error("Invalid Admin Username or Password.")


def show_worker_portal():
    curr_user = st.session_state.get("username", "Worker")
    perms = load_permissions()
    u_perms = {}
    for k, v in perms.items():
        if k.lower() == curr_user.strip().lower():
            u_perms = v
            break

    can_edit_perm = u_perms.get("can_edit", u_perms.get("can_update", True))
    can_add_perm = u_perms.get("can_add", True)
    can_search_perm = u_perms.get("can_search", True)
    can_download_perm = u_perms.get("can_download", False)
    can_sync_perm = u_perms.get("can_sync", False)

    all_sheets = load_all_sheets_cached(_get_file_mtime())
    visible_sheets = [s for s in all_sheets.keys() if str(s).strip() not in SYSTEM_SHEETS]
    if not visible_sheets:
        visible_sheets = list(all_sheets.keys()) or ["Sheet1"]

    active_sheet = st.session_state.get("active_sheet", get_default_active_sheet(visible_sheets))
    if active_sheet not in visible_sheets:
        active_sheet = get_default_active_sheet(visible_sheets)
        st.session_state["active_sheet"] = active_sheet

    notice_txt = get_admin_notice()
    if notice_txt:
        st.markdown(
            f'<div class="marquee-container"><marquee behavior="scroll" direction="left">📢 {notice_txt}</marquee></div>',
            unsafe_allow_html=True,
        )

    st.markdown(
        f"""
        <div class="worker-hero-banner">
            <div class="worker-hero-title">👩‍💼 Case Worker Workspace — {curr_user}</div>
            <div class="worker-pill-group">
                <span class="worker-pill-gold">Sheet: {active_sheet}</span>
                <span class="worker-pill-cyan">IST: {get_ist_now().strftime("%I:%M %p")}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    h1, h2, h3, h4, h5 = st.columns([2.2, 1.4, 1.3, 1.3, 1.2], gap="small")
    with h1:
        selected_s = st.selectbox(
            "Select Sheet",
            options=visible_sheets,
            index=visible_sheets.index(active_sheet),
            label_visibility="collapsed",
            key="worker_sheet_selector",
        )
        if selected_s != active_sheet:
            st.session_state["active_sheet"] = selected_s
            st.query_params["active_sheet"] = selected_s
            st.rerun()
    with h2:
        if st.button("🔄 Sync Sheet", key="worker_sync_btn", disabled=(not can_sync_perm), use_container_width=True):
            _clear_sheet_runtime_caches()
            st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
            st.rerun()
    with h3:
        if can_download_perm:
            exp_bytes = get_worker_export_bytes(_get_file_mtime(), curr_user, active_sheet)
            st.download_button(
                label="📥 Export Excel",
                data=exp_bytes,
                file_name=f"{curr_user}_{active_sheet}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="worker_export_btn",
                use_container_width=True,
            )
        else:
            st.button("📥 Export Excel", key="worker_export_disabled", disabled=True, use_container_width=True)
    with h4:
        if st.button("➕ Add Case", key="worker_add_case_btn", disabled=(not can_add_perm), use_container_width=True):
            st.session_state["open_dialog"] = "worker_add_case"
            st.rerun()
    with h5:
        if st.button("🚪 Logout", key="worker_logout_btn", use_container_width=True):
            record_worker_logout(curr_user)
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            st.session_state["attendance_recorded"] = False
            close_case_popup_state()
            st.query_params.clear()
            st.rerun()

    if st.session_state.get("open_dialog") == "worker_add_case":
        @st.dialog("Add New Case", width="large")
        def worker_add_case_modal():
            df_cur = all_sheets.get(active_sheet, pd.DataFrame())
            next_cid = generate_next_case_id()
            login_c = find_login_user_column(df_cur) if not df_cur.empty else "Case Worker"
            disp_c = find_display_username_column(df_cur) if not df_cur.empty else "USERNAME"
            p_c = find_password_column(df_cur) if not df_cur.empty else "Password"

            with st.form("worker_add_case_form"):
                st.text_input("Auto Case ID", value=next_cid, disabled=True)
                st.text_input("Assigned Worker", value=curr_user, disabled=True)
                u_val = st.text_input(f"{disp_c} *", placeholder="CGI Username")
                pw_val = st.text_input(f"{p_c}", placeholder="Password")
                if st.form_submit_button("✅ Save New Case", use_container_width=True):
                    if not u_val.strip():
                        st.error("Username is required.")
                    else:
                        now_str = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                        new_row = {col: "" for col in df_cur.columns}
                        new_row["Case_ID"] = next_cid
                        new_row[disp_c] = u_val.strip()
                        if p_c in new_row:
                            new_row[p_c] = pw_val.strip()
                        if login_c in new_row:
                            new_row[login_c] = curr_user
                        new_row["Booking_Status"] = "Active"
                        new_row["Updated_By"] = curr_user
                        new_row["Updated_At"] = now_str
                        new_row["Action_Type"] = "Added"
                        df_new = pd.concat([df_cur, pd.DataFrame([new_row])], ignore_index=True)
                        save_data(df_new, sheet_name=active_sheet)
                        st.session_state["open_dialog"] = None
                        st.success(f"✅ Added case {next_cid}!")
                        st.rerun()
            if st.button("Close", key="close_worker_add_modal"):
                st.session_state["open_dialog"] = None
                st.rerun()

        worker_add_case_modal()

    df_sheet = all_sheets.get(active_sheet, pd.DataFrame()).copy()
    if df_sheet.empty:
        st.info(f"No records found in sheet '{active_sheet}'.")
        return

    df_sheet = reorder_dataframe_columns(df_sheet)
    login_c = find_login_user_column(df_sheet)
    display_c = find_display_username_column(df_sheet)
    p_col = find_password_column(df_sheet)

    my_mask = df_sheet[login_c].astype(str).str.strip().str.lower() == curr_user.strip().lower()
    df_my = df_sheet[my_mask].copy()
    df_my["__orig_idx__"] = df_my.index

    total_my = len(df_my)
    booked_my = int((df_my["Booking_Status"].astype(str).str.strip().str.lower() == "booked").sum()) if "Booking_Status" in df_my.columns else 0
    resched_my = int((df_my["Booking_Status"].astype(str).str.strip().str.lower() == "rescheduled").sum()) if "Booking_Status" in df_my.columns else 0
    issues_my = int((df_my["Booking_Status"].astype(str).str.strip().str.lower() == "issues").sum()) if "Booking_Status" in df_my.columns else 0

    k1, k2, k3, k4 = st.columns(4, gap="small")
    with k1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">My Cases</div><div class="kpi-value">{total_my}</div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="kpi-card green"><div class="kpi-title">Booked</div><div class="kpi-value">{booked_my}</div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="kpi-card gold"><div class="kpi-title">Rescheduled</div><div class="kpi-value">{resched_my}</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="kpi-card purple"><div class="kpi-title">Issues</div><div class="kpi-value">{issues_my}</div></div>', unsafe_allow_html=True)

    view_scope = st.radio(
        "Case View",
        options=["My Assigned Cases", "All Cases in Sheet", "Global Search (All Sheets)"],
        horizontal=True,
        key="worker_view_scope_radio",
    )

    search_q = ""
    if can_search_perm:
        search_q = st.text_input(
            "🔍 Search Case ID, Username, DS-160, Password, or Worker",
            value="",
            placeholder="Type search term...",
            key="worker_search_input",
        )

    if view_scope == "Global Search (All Sheets)":
        df_show = search_across_all_sheets_by_fields(search_q) if search_q.strip() else pd.DataFrame()
        if not search_q.strip():
            st.info("Enter a search term above to search across all sheets.")
            return
    elif view_scope == "All Cases in Sheet":
        df_show = df_sheet.copy()
        df_show["__orig_idx__"] = df_show.index
        if can_search_perm and search_q.strip():
            df_show = df_show[search_username_ds160_password_worker(df_show, search_q)]
    else:
        df_show = df_my.copy()
        if can_search_perm and search_q.strip():
            df_show = df_show[search_username_ds160_password_worker(df_show, search_q)]

    render_clickable_table(
        df=df_show,
        display_c=display_c,
        p_col=p_col,
        table_key=f"worker_{active_sheet}_{view_scope}",
        sheet_for_link=active_sheet,
        page_size=100,
        allow_edit=can_edit_perm,
        restrict_edit_to_user=None,
    )
    # ================= PART 9: ADMIN PORTAL, DIALOG TRIGGER & ENTRYPOINT =================
def show_admin_portal():
    all_sheets = load_all_sheets_cached(_get_file_mtime())
    visible_sheets = [s for s in all_sheets.keys() if str(s).strip() not in SYSTEM_SHEETS]
    if not visible_sheets:
        visible_sheets = list(all_sheets.keys()) or ["Sheet1"]

    active_sheet = st.session_state.get("active_sheet", get_default_active_sheet(visible_sheets))
    if active_sheet not in visible_sheets:
        active_sheet = get_default_active_sheet(visible_sheets)
        st.session_state["active_sheet"] = active_sheet

    notice_txt = get_admin_notice()
    if notice_txt:
        st.markdown(
            f'<div class="marquee-container"><marquee behavior="scroll" direction="left">📢 {notice_txt}</marquee></div>',
            unsafe_allow_html=True,
        )

    st.markdown(
        f"""
        <div class="workspace-header-bar">
            <div>🛡️ Admin Control Center — HemanVisa</div>
            <div class="workspace-sheet-pill">Active Sheet: {active_sheet}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    top_c1, top_c2, top_c3, top_c4, top_c5 = st.columns([2.2, 1.3, 1.4, 1.3, 1.2], gap="small")
    with top_c1:
        picked_sheet = st.selectbox(
            "Select Sheet",
            options=visible_sheets,
            index=visible_sheets.index(active_sheet),
            label_visibility="collapsed",
            key="admin_sheet_select",
        )
        if picked_sheet != active_sheet:
            st.session_state["active_sheet"] = picked_sheet
            st.query_params["active_sheet"] = picked_sheet
            st.rerun()
    with top_c2:
        if st.button("🔄 Sync All", key="admin_sync_all_btn", use_container_width=True):
            _clear_sheet_runtime_caches()
            st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
            st.rerun()
    with top_c3:
        admin_export_bytes = get_admin_gsheet_export_bytes(_get_file_mtime())
        st.download_button(
            label="📥 Export Workbook",
            data=admin_export_bytes,
            file_name="Heman_Report_Export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="admin_export_btn",
            use_container_width=True,
        )
    with top_c4:
        if st.button("➕ Add Case", key="admin_add_case_btn", use_container_width=True):
            st.session_state["open_dialog"] = "add_case"
            st.rerun()
    with top_c5:
        if st.button("🚪 Logout", key="admin_logout_btn", use_container_width=True):
            st.session_state["logged_in"] = False
            st.session_state["username"] = ""
            st.session_state["role"] = ""
            close_case_popup_state()
            st.query_params.clear()
            st.rerun()

    st.markdown('<div class="action-group-title">Admin Management Actions</div>', unsafe_allow_html=True)
    a_cols = st.columns(6, gap="small")
    admin_actions = [
        ("🔐 Permissions", "worker_permissions", "admin-btn-0"),
        ("📩 Edit Reqs", "edit_requests", "admin-btn-1"),
        ("🗑️ Del Reqs", "delete_requests", "admin-btn-2"),
        ("🕒 Attendance", "login_attendance", "admin-btn-3"),
        ("📢 Notice & Auth", "notice_admin_creds", "admin-btn-4"),
    ]
    for idx_a, (lbl, d_mode, css_cls) in enumerate(admin_actions):
        with a_cols[idx_a]:
            st.markdown(f'<span class="{css_cls}"></span>', unsafe_allow_html=True)
            if st.button(lbl, key=f"adm_act_{d_mode}", use_container_width=True):
                st.session_state["open_dialog"] = d_mode
                st.rerun()

    with a_cols[5]:
        st.markdown('<span class="admin-btn-5"></span>', unsafe_allow_html=True)
        all_blk = are_all_workers_blocked()
        blk_lbl = "🔓 Unblock All" if all_blk else "🚫 Block All"
        if st.button(blk_lbl, key="adm_toggle_all_block", use_container_width=True):
            toggle_all_workers_block(not all_blk)
            st.rerun()

    df_sheet = all_sheets.get(active_sheet, pd.DataFrame()).copy()
    if df_sheet.empty:
        st.info(f"No data found in sheet '{active_sheet}'.")
        return

    df_sheet = reorder_dataframe_columns(df_sheet)
    login_c = find_login_user_column(df_sheet)
    display_c = find_display_username_column(df_sheet)
    p_col = find_password_column(df_sheet)

    total_cases = len(df_sheet)
    assigned_cases = int(
        df_sheet[login_c]
        .astype(str)
        .apply(lambda x: 0 if is_not_assigned_value(x) else 1)
        .sum()
    )
    unassigned_cases = total_cases - assigned_cases
    booked_cases = (
        int((df_sheet["Booking_Status"].astype(str).str.strip().str.lower() == "booked").sum())
        if "Booking_Status" in df_sheet.columns
        else 0
    )
    resched_cases = (
        int((df_sheet["Booking_Status"].astype(str).str.strip().str.lower() == "rescheduled").sum())
        if "Booking_Status" in df_sheet.columns
        else 0
    )
    issues_cases = (
        int((df_sheet["Booking_Status"].astype(str).str.strip().str.lower() == "issues").sum())
        if "Booking_Status" in df_sheet.columns
        else 0
    )

    k1, k2, k3, k4, k5, k6 = st.columns(6, gap="small")
    with k1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Cases</div><div class="kpi-value">{total_cases}</div></div>', unsafe_allow_html=True)
    with k2:
        st.markdown(f'<div class="kpi-card green"><div class="kpi-title">Assigned</div><div class="kpi-value">{assigned_cases}</div></div>', unsafe_allow_html=True)
    with k3:
        st.markdown(f'<div class="kpi-card gold"><div class="kpi-title">Unassigned</div><div class="kpi-value">{unassigned_cases}</div></div>', unsafe_allow_html=True)
    with k4:
        st.markdown(f'<div class="kpi-card teal"><div class="kpi-title">Booked</div><div class="kpi-value">{booked_cases}</div></div>', unsafe_allow_html=True)
    with k5:
        st.markdown(f'<div class="kpi-card purple"><div class="kpi-title">Rescheduled</div><div class="kpi-value">{resched_cases}</div></div>', unsafe_allow_html=True)
    with k6:
        st.markdown(f'<div class="kpi-card rose"><div class="kpi-title">Issues</div><div class="kpi-value">{issues_cases}</div></div>', unsafe_allow_html=True)

    f1, f2, f3 = st.columns([1.4, 1.4, 2.4], gap="small")
    with f1:
        workers_list = ["All Workers"] + get_all_workers_list()
        sel_worker_f = st.selectbox("Filter by Worker", options=workers_list, key="admin_worker_filter")
    with f2:
        lead_col = next((c for c in df_sheet.columns if "lead" in c.strip().lower()), None)
        lead_options = ["All Leads"]
        if lead_col:
            lead_vals = sorted(
                {str(v).strip() for v in df_sheet[lead_col].dropna() if str(v).strip() and str(v).strip().lower() != "nan"}
            )
            lead_options.extend(lead_vals)
        sel_lead_f = st.selectbox("Filter by Lead", options=lead_options, key="admin_lead_filter")
    with f3:
        admin_search_q = st.text_input(
            "🔍 Search Cases",
            value="",
            placeholder="Case ID, Username, DS-160, Worker, Password...",
            key="admin_search_input",
        )

    df_filtered = df_sheet.copy()
    df_filtered["__orig_idx__"] = df_filtered.index

    if sel_worker_f != "All Workers":
        df_filtered = df_filtered[
            df_filtered[login_c].astype(str).str.strip().str.lower() == sel_worker_f.strip().lower()
        ]
    if lead_col and sel_lead_f != "All Leads":
        df_filtered = df_filtered[
            df_filtered[lead_col].astype(str).str.strip().str.lower() == sel_lead_f.strip().lower()
        ]
    if admin_search_q.strip():
        df_filtered = df_filtered[search_username_ds160_password_worker(df_filtered, admin_search_q)]

    render_clickable_table(
        df=df_filtered,
        display_c=display_c,
        p_col=p_col,
        table_key=f"admin_{active_sheet}_{sel_worker_f}_{sel_lead_f}",
        sheet_for_link=active_sheet,
        page_size=100,
        allow_edit=True,
        restrict_edit_to_user=None,
    )


if st.session_state.get("logged_in"):
    if st.session_state.get("role") == "admin":
        show_admin_portal()
    else:
        show_worker_portal()

    active_dialog = st.session_state.get("open_dialog")
    if active_dialog == "edit_case_popup" and st.session_state.get("selected_row_idx") is not None:
        view_and_edit_case_dialog()
    elif active_dialog in [
        "add_case",
        "worker_permissions",
        "edit_requests",
        "delete_requests",
        "login_attendance",
        "notice_admin_creds",
    ]:
        open_admin_modal()
else:
    nav_choice = st.radio(
        "Navigation",
        options=["🏠 Home", "🔐 Login"],
        horizontal=True,
        key="public_nav_radio",
    )
    if nav_choice == "🏠 Home":
        show_home_page()
    else:
        show_login_page()
    
