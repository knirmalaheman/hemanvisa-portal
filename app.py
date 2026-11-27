# ================= STEP 1: IMPORTS, CONFIG, CGI FORMATTER & SHEET SYNC BASE =================
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
# ================= END OF STEP 1 =================
# ================= STEP 2: EDIT REQUESTS, ADMIN CREDS, NOTICE & PERMISSIONS =================
EDIT_REQ_COLS = [
    "Request_ID",
    "Case_ID",
    "Case_Username",
    "Worker",
    "Field_Name",
    "Old_Value",
    "New_Value",
    "Reason",
    "Requested_At",
    "Status",
    "Reviewed_By",
    "Reviewed_At",
]


def load_edit_requests():
    meta = _get_runtime_meta()
    if meta["edit_requests"] is not None:
        return list(meta["edit_requests"])
    rows = _read_sheet_rows("App_Edit_Requests")
    cleaned = []
    for r in rows:
        item = {col: str(r.get(col, "")).strip() for col in EDIT_REQ_COLS}
        if item["Request_ID"] or item["Case_ID"]:
            if not item["Status"]:
                item["Status"] = "Pending"
            cleaned.append(item)
    meta["edit_requests"] = cleaned
    return list(cleaned)


def save_edit_requests(req_list):
    cleaned = []
    for r in req_list or []:
        item = {col: str(r.get(col, "")).strip() for col in EDIT_REQ_COLS}
        if item["Request_ID"] or item["Case_ID"]:
            if not item["Status"]:
                item["Status"] = "Pending"
            cleaned.append(item)
    _write_sheet_records("App_Edit_Requests", EDIT_REQ_COLS, cleaned)
    meta = _get_runtime_meta()
    meta["edit_requests"] = list(cleaned)


def get_admin_creds():
    meta = _get_runtime_meta()
    if meta["admin_creds"] is not None:
        return dict(meta["admin_creds"])
    rows = _read_sheet_rows("App_Config")
    for r in rows:
        if str(r.get("Key", "")).strip() == "admin_creds":
            raw_val = str(r.get("Value", "")).strip()
            if "|" in raw_val:
                u, p = raw_val.split("|", 1)
                creds = {"user": u.strip(), "pass": p.strip()}
                meta["admin_creds"] = creds
                return dict(creds)
    creds = {"user": "admin", "pass": "heman123"}
    meta["admin_creds"] = creds
    return dict(creds)


def save_admin_creds(user, pwd):
    clean_user = str(user).strip()
    clean_pwd = str(pwd).strip()
    existing = _read_sheet_rows("App_Config")
    found = False
    records = []
    for r in existing:
        k = str(r.get("Key", "")).strip()
        v = str(r.get("Value", "")).strip()
        if not k:
            continue
        if k == "admin_creds":
            records.append({"Key": "admin_creds", "Value": f"{clean_user}|{clean_pwd}"})
            found = True
        else:
            records.append({"Key": k, "Value": v})
    if not found:
        records.append({"Key": "admin_creds", "Value": f"{clean_user}|{clean_pwd}"})
    _write_sheet_records("App_Config", ["Key", "Value"], records)
    meta = _get_runtime_meta()
    meta["admin_creds"] = {"user": clean_user, "pass": clean_pwd}


def get_notice():
    meta = _get_runtime_meta()
    if meta["notice"] is not None:
        return meta["notice"]
    rows = _read_sheet_rows("App_Config")
    for r in rows:
        if str(r.get("Key", "")).strip() == "admin_notice":
            val = str(r.get("Value", "")).strip()
            meta["notice"] = val
            return val
    meta["notice"] = ""
    return ""


def save_notice(text):
    clean_text = str(text).strip()
    existing = _read_sheet_rows("App_Config")
    found = False
    records = []
    for r in existing:
        k = str(r.get("Key", "")).strip()
        v = str(r.get("Value", "")).strip()
        if not k:
            continue
        if k == "admin_notice":
            records.append({"Key": "admin_notice", "Value": clean_text})
            found = True
        else:
            records.append({"Key": k, "Value": v})
    if not found:
        records.append({"Key": "admin_notice", "Value": clean_text})
    _write_sheet_records("App_Config", ["Key", "Value"], records)
    meta = _get_runtime_meta()
    meta["notice"] = clean_text


def load_permissions():
    meta = _get_runtime_meta()
    if meta["perms"] is not None:
        return {k: dict(v) for k, v in meta["perms"].items()}
    rows = _read_sheet_rows("App_Permissions")
    perms = {}
    for r in rows:
        w = str(r.get("Worker", "")).strip()
        if not w:
            continue
        perms[w] = {
            "can_edit": str(r.get("can_edit", "True")).strip().lower() == "true",
            "can_delete": str(r.get("can_delete", "False")).strip().lower() == "true",
            "can_add_case": str(r.get("can_add_case", "True")).strip().lower() == "true",
            "can_view_all": str(r.get("can_view_all", "False")).strip().lower() == "true",
            "can_export": str(r.get("can_export", "True")).strip().lower() == "true",
        }
    meta["perms"] = perms
    return {k: dict(v) for k, v in perms.items()}


def save_permissions(perms_dict):
    records = []
    for w, p in (perms_dict or {}).items():
        w_clean = str(w).strip()
        if not w_clean:
            continue
        records.append(
            {
                "Worker": w_clean,
                "can_edit": str(bool(p.get("can_edit", True))),
                "can_delete": str(bool(p.get("can_delete", False))),
                "can_add_case": str(bool(p.get("can_add_case", True))),
                "can_view_all": str(bool(p.get("can_view_all", False))),
                "can_export": str(bool(p.get("can_export", True))),
            }
        )
    _write_sheet_records(
        "App_Permissions",
        ["Worker", "can_edit", "can_delete", "can_add_case", "can_view_all", "can_export"],
        records,
    )
    meta = _get_runtime_meta()
    meta["perms"] = {
        r["Worker"]: {
            "can_edit": r["can_edit"] == "True",
            "can_delete": r["can_delete"] == "True",
            "can_add_case": r["can_add_case"] == "True",
            "can_view_all": r["can_view_all"] == "True",
            "can_export": r["can_export"] == "True",
        }
        for r in records
    }


def get_worker_perms(worker_name):
    all_perms = load_permissions()
    default_p = {
        "can_edit": True,
        "can_delete": False,
        "can_add_case": True,
        "can_view_all": False,
        "can_export": True,
    }
    return all_perms.get(worker_name, default_p)
# ================= END OF STEP 2 =================
# ================= STEP 3: WORKER LOGINS, ATTENDANCE, CASE ID & PDF HELPERS =================
def load_worker_logins():
    meta = _get_runtime_meta()
    if meta["logins"] is not None:
        return {k: dict(v) for k, v in meta["logins"].items()}
    rows = _read_sheet_rows("App_Logins")
    logins = {}
    for r in rows:
        w = str(r.get("Worker", "")).strip()
        if not w:
            continue
        logins[w] = {
            "password": str(r.get("Password", "")).strip(),
            "last_login": str(r.get("Last_Login", "")).strip(),
            "last_logout": str(r.get("Last_Logout", "")).strip(),
            "status": str(r.get("Status", "Offline")).strip() or "Offline",
            "last_seen": str(r.get("Last_Seen", "")).strip(),
        }
    meta["logins"] = logins
    return {k: dict(v) for k, v in logins.items()}


def save_worker_logins(logins_dict):
    records = []
    for w, info in (logins_dict or {}).items():
        w_clean = str(w).strip()
        if not w_clean:
            continue
        records.append(
            {
                "Worker": w_clean,
                "Password": str(info.get("password", "")).strip(),
                "Last_Login": str(info.get("last_login", "")).strip(),
                "Last_Logout": str(info.get("last_logout", "")).strip(),
                "Status": str(info.get("status", "Offline")).strip() or "Offline",
                "Last_Seen": str(info.get("last_seen", "")).strip(),
            }
        )
    _write_sheet_records(
        "App_Logins",
        ["Worker", "Password", "Last_Login", "Last_Logout", "Status", "Last_Seen"],
        records,
    )
    meta = _get_runtime_meta()
    meta["logins"] = {
        r["Worker"]: {
            "password": r["Password"],
            "last_login": r["Last_Login"],
            "last_logout": r["Last_Logout"],
            "status": r["Status"],
            "last_seen": r["Last_Seen"],
        }
        for r in records
    }


def load_attendance():
    meta = _get_runtime_meta()
    if meta["attendance"] is not None:
        return list(meta["attendance"])
    rows = _read_sheet_rows("App_Attendance")
    cleaned = []
    for r in rows:
        item = {
            "Date": str(r.get("Date", "")).strip(),
            "Worker": str(r.get("Worker", "")).strip(),
            "Login_Time": str(r.get("Login_Time", "")).strip(),
            "Logout_Time": str(r.get("Logout_Time", "")).strip(),
            "Total_Hours": str(r.get("Total_Hours", "")).strip(),
            "Status": str(r.get("Status", "")).strip(),
        }
        if item["Date"] or item["Worker"]:
            cleaned.append(item)
    meta["attendance"] = cleaned
    return list(cleaned)


def save_attendance(att_list):
    cols = ["Date", "Worker", "Login_Time", "Logout_Time", "Total_Hours", "Status"]
    cleaned = []
    for r in att_list or []:
        item = {c: str(r.get(c, "")).strip() for c in cols}
        if item["Date"] or item["Worker"]:
            cleaned.append(item)
    _write_sheet_records("App_Attendance", cols, cleaned)
    meta = _get_runtime_meta()
    meta["attendance"] = list(cleaned)


def record_worker_login_time(worker_name):
    now_dt = get_ist_now()
    date_str = now_dt.strftime("%Y-%m-%d")
    time_str = now_dt.strftime("%I:%M:%S %p")
    full_str = now_dt.strftime("%Y-%m-%d %I:%M:%S %p")

    logins = load_worker_logins()
    if worker_name not in logins:
        logins[worker_name] = {
            "password": "",
            "last_login": full_str,
            "last_logout": "",
            "status": "Online",
            "last_seen": full_str,
        }
    else:
        logins[worker_name]["last_login"] = full_str
        logins[worker_name]["status"] = "Online"
        logins[worker_name]["last_seen"] = full_str
    save_worker_logins(logins)

    att = load_attendance()
    found_idx = None
    for idx, rec in enumerate(att):
        if rec.get("Date") == date_str and rec.get("Worker") == worker_name:
            found_idx = idx
            break
    if found_idx is not None:
        if not att[found_idx].get("Login_Time"):
            att[found_idx]["Login_Time"] = time_str
        att[found_idx]["Status"] = "Present"
    else:
        att.append(
            {
                "Date": date_str,
                "Worker": worker_name,
                "Login_Time": time_str,
                "Logout_Time": "",
                "Total_Hours": "",
                "Status": "Present",
            }
        )
    save_attendance(att)


def record_worker_logout_time(worker_name):
    now_dt = get_ist_now()
    date_str = now_dt.strftime("%Y-%m-%d")
    time_str = now_dt.strftime("%I:%M:%S %p")
    full_str = now_dt.strftime("%Y-%m-%d %I:%M:%S %p")

    logins = load_worker_logins()
    if worker_name in logins:
        logins[worker_name]["last_logout"] = full_str
        logins[worker_name]["status"] = "Offline"
        logins[worker_name]["last_seen"] = full_str
        save_worker_logins(logins)

    att = load_attendance()
    for rec in att:
        if rec.get("Date") == date_str and rec.get("Worker") == worker_name:
            rec["Logout_Time"] = time_str
            login_str = rec.get("Login_Time", "")
            if login_str:
                try:
                    t1 = datetime.datetime.strptime(login_str, "%I:%M:%S %p")
                    t2 = datetime.datetime.strptime(time_str, "%I:%M:%S %p")
                    diff = t2 - t1
                    if diff.total_seconds() >= 0:
                        mins = int(diff.total_seconds() // 60)
                        hrs = mins // 60
                        rem_m = mins % 60
                        rec["Total_Hours"] = f"{hrs}h {rem_m}m"
                except Exception:
                    pass
            break
    save_attendance(att)


def get_next_case_id(all_sheets):
    used_nums = set()
    if isinstance(all_sheets, dict):
        for s_name, df_s in all_sheets.items():
            if s_name in SYSTEM_SHEETS or df_s is None or df_s.empty:
                continue
            for col in df_s.columns:
                if "case_id" in str(col).strip().lower() or str(col).strip().lower() == "case id":
                    for val in df_s[col].dropna().astype(str):
                        m = re.search(r"HM(\d+)", val.strip(), re.IGNORECASE)
                        if m:
                            used_nums.add(int(m.group(1)))
    next_num = max(used_nums) + 1 if used_nums else 1111
    return f"HM{next_num}"


def _sanitize_case_folder_name(case_id):
    cleaned = re.sub(r"[^A-Za-z0-9_-]", "", str(case_id or "").strip())
    return cleaned if cleaned else "UNKNOWN_CASE"


def get_case_pdf_folder(case_id, create=False):
    folder = os.path.join(PDF_BASE_DIR, _sanitize_case_folder_name(case_id))
    if create and not os.path.exists(folder):
        os.makedirs(folder, exist_ok=True)
    return folder


def list_case_pdfs(case_id):
    folder = get_case_pdf_folder(case_id, create=False)
    if not os.path.exists(folder):
        return []
    files = []
    for fname in sorted(os.listdir(folder)):
        fpath = os.path.join(folder, fname)
        if os.path.isfile(fpath) and fname.lower().endswith(".pdf"):
            files.append(fname)
    return files


def save_case_pdf_file(case_id, uploaded_file):
    if uploaded_file is None:
        return None
    folder = get_case_pdf_folder(case_id, create=True)
    raw_name = os.path.basename(str(uploaded_file.name or "appointment.pdf"))
    safe_name = re.sub(r"[^A-Za-z0-9._-]", "_", raw_name)
    if not safe_name.lower().endswith(".pdf"):
        safe_name += ".pdf"
    dest_path = os.path.join(folder, safe_name)
    if os.path.exists(dest_path):
        ts = get_ist_now().strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(safe_name)
        safe_name = f"{base}_{ts}{ext}"
        dest_path = os.path.join(folder, safe_name)
    file_bytes = uploaded_file.getvalue()
    with open(dest_path, "wb") as f:
        f.write(file_bytes)
    return safe_name


def delete_case_pdf_file(case_id, filename):
    folder = get_case_pdf_folder(case_id, create=False)
    safe_name = os.path.basename(str(filename or ""))
    fpath = os.path.join(folder, safe_name)
    if os.path.exists(fpath) and os.path.isfile(fpath):
        os.remove(fpath)
        return True
    return False
# ================= END OF STEP 3 =================
# ================= STEP 4: SHEET NORMALIZATION, LOADING, SAVING & WORKERS LIST =================
def _normalize_sheet_df(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=["Case_ID", "Username", "USERNAME", "Password", "Lead"], dtype=str)
    df = df.copy()
    df.columns = df.columns.str.strip()
    df = df.fillna("").astype(str)

    if "Case_ID" not in df.columns:
        df.insert(0, "Case_ID", "")
    if "Lead" not in df.columns:
        lead_idx = df.columns.get_loc("Case_ID") + 1
        df.insert(lead_idx, "Lead", "")

    used_ids = set()
    for col in df.columns:
        if "case_id" in str(col).strip().lower() or str(col).strip().lower() == "case id":
            for v in df[col]:
                m = re.search(r"HM(\d+)", str(v).strip(), re.IGNORECASE)
                if m:
                    used_ids.add(int(m.group(1)))
    next_num = max(used_ids) + 1 if used_ids else 1111

    has_content_mask = pd.Series(False, index=df.index)
    for col in df.columns:
        if col in ["Case_ID", "Lead"]:
            continue
        col_vals = df[col].astype(str).str.strip()
        has_content_mask = has_content_mask | (
            (col_vals != "") & (~col_vals.str.lower().isin(["nan", "none", "null"]))
        )

    for idx in df.index:
        cid_val = str(df.at[idx, "Case_ID"]).strip()
        if not cid_val or cid_val.lower() in ["nan", "none", "null"]:
            if bool(has_content_mask.loc[idx]):
                df.at[idx, "Case_ID"] = f"HM{next_num}"
                next_num += 1
            else:
                df.at[idx, "Case_ID"] = ""

    df = df[has_content_mask | (df["Case_ID"].astype(str).str.strip() != "")].copy()
    df = df.reset_index(drop=True)
    return df


@st.cache_data(ttl=120, show_spinner=False)
def _fetch_and_process_all_sheets_cached(file_token):
    webapp_sheets = _fetch_sheets_from_webapp()
    if isinstance(webapp_sheets, dict) and len(webapp_sheets) > 0:
        processed = {}
        for s_name, df_raw in webapp_sheets.items():
            s_clean = str(s_name).strip()
            if not s_clean:
                continue
            if s_clean in SYSTEM_SHEETS:
                df_sys = df_raw.copy()
                df_sys.columns = df_sys.columns.str.strip()
                processed[s_clean] = df_sys.fillna("").astype(str)
            else:
                processed[s_clean] = _normalize_sheet_df(df_raw)
        if not processed:
            processed["Active Cases"] = _normalize_sheet_df(pd.DataFrame())
        return processed

    if os.path.exists(EXCEL_FILE):
        try:
            excel_data = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
            processed = {}
            for s_name, df_raw in excel_data.items():
                s_clean = str(s_name).strip()
                if not s_clean:
                    continue
                if s_clean in SYSTEM_SHEETS:
                    df_sys = df_raw.copy()
                    df_sys.columns = df_sys.columns.str.strip()
                    processed[s_clean] = df_sys.fillna("").astype(str)
                else:
                    processed[s_clean] = _normalize_sheet_df(df_raw)
            if processed:
                return processed
        except Exception:
            pass

    return {"Active Cases": _normalize_sheet_df(pd.DataFrame())}


def load_all_sheets_cached(file_token=None):
    token = file_token if file_token is not None else _get_file_mtime()
    mem_cache = st.session_state.get("_mem_sheets_cache")
    if isinstance(mem_cache, dict) and mem_cache.get("_token") == token:
        return {k: v.copy() for k, v in mem_cache.items() if k != "_token"}

    sheets = _fetch_and_process_all_sheets_cached(token)
    copy_map = {k: v.copy() for k, v in sheets.items()}
    st.session_state["_mem_sheets_cache"] = {"_token": token, **copy_map}
    return {k: v.copy() for k, v in copy_map.items()}


def save_all_sheets(sheets_dict, push_to_cloud=True):
    cleaned_sheets = {}
    for s_name, df_s in (sheets_dict or {}).items():
        s_clean = str(s_name).strip()
        if not s_clean:
            continue
        if s_clean in SYSTEM_SHEETS:
            df_sys = df_s.copy()
            df_sys.columns = df_sys.columns.str.strip()
            cleaned_sheets[s_clean] = df_sys.fillna("").astype(str)
        else:
            cleaned_sheets[s_clean] = _normalize_sheet_df(df_s)

    if not cleaned_sheets:
        cleaned_sheets["Active Cases"] = _normalize_sheet_df(pd.DataFrame())

    st.session_state["_gsheet_sync_version"] = (
        st.session_state.get("_gsheet_sync_version", 0) + 1
    )
    new_token = _get_file_mtime()
    st.session_state["_mem_sheets_cache"] = {
        "_token": new_token,
        **{k: v.copy() for k, v in cleaned_sheets.items()},
    }
    st.session_state.pop("_parsed_meta_cache", None)
    get_admin_gsheet_export_bytes.clear()
    get_worker_export_bytes.clear()

    if push_to_cloud and get_webapp_url():
        first_sheet = next(iter(cleaned_sheets.keys()))
        _push_sheet_to_webapp(
            first_sheet, cleaned_sheets[first_sheet], all_sheets_dict=cleaned_sheets
        )


def get_workers_list(all_sheets=None):
    meta = _get_runtime_meta()
    if all_sheets is None and meta["workers_list"] is not None:
        return list(meta["workers_list"])

    workers = []
    seen = set()

    def _add_w(val):
        s = str(val or "").strip()
        if not s or s.lower() in ["nan", "none", "null", "admin", "all"]:
            return
        if s not in seen:
            seen.add(s)
            workers.append(s)

    logins = load_worker_logins()
    for w_name in logins.keys():
        _add_w(w_name)

    perms = load_permissions()
    for w_name in perms.keys():
        _add_w(w_name)

    sheets_map = all_sheets if isinstance(all_sheets, dict) else load_all_sheets_cached()
    for s_name, df_s in sheets_map.items():
        if s_name in SYSTEM_SHEETS or df_s is None or df_s.empty:
            continue
        for col in df_s.columns:
            col_l = str(col).strip().lower()
            if col_l in ["lead", "worker", "case worker", "assigned to", "handler"]:
                for v in df_s[col].dropna().astype(str):
                    _add_w(v)

    for default_w in ["Kiran", "Shanthi", "Keerthana", "Sai", "Kalyan", "Anil"]:
        _add_w(default_w)

    if all_sheets is None:
        meta["workers_list"] = list(workers)
    return list(workers)
# ================= END OF STEP 4 =================
# ================= STEP 5: EXPORT HELPERS & GLOBAL CSS THEME =================
@st.cache_data(ttl=60, show_spinner=False)
def get_admin_gsheet_export_bytes(file_token):
    sheets = load_all_sheets_cached(file_token)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wrote_any = False
        for s_name, df_s in sheets.items():
            if s_name in SYSTEM_SHEETS:
                continue
            safe_name = re.sub(r"[\\/*?:\[\]]", "", str(s_name))[:31] or "Sheet"
            df_out = df_s.fillna("").astype(str) if df_s is not None else pd.DataFrame()
            df_out.to_excel(writer, sheet_name=safe_name, index=False)
            wrote_any = True
        if not wrote_any:
            pd.DataFrame(columns=["Case_ID", "Lead"]).to_excel(
                writer, sheet_name="Cases", index=False
            )
    output.seek(0)
    return output.getvalue()


@st.cache_data(ttl=60, show_spinner=False)
def get_worker_export_bytes(worker_name, file_token):
    sheets = load_all_sheets_cached(file_token)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wrote_any = False
        w_low = str(worker_name or "").strip().lower()
        for s_name, df_s in sheets.items():
            if s_name in SYSTEM_SHEETS or df_s is None or df_s.empty:
                continue
            lead_col = None
            for col in df_s.columns:
                if str(col).strip().lower() in ["lead", "worker", "case worker", "assigned to"]:
                    lead_col = col
                    break
            if lead_col is not None:
                sub_df = df_s[
                    df_s[lead_col].astype(str).str.strip().str.lower() == w_low
                ].copy()
            else:
                sub_df = df_s.iloc[0:0].copy()
            safe_name = re.sub(r"[\\/*?:\[\]]", "", str(s_name))[:31] or "MyCases"
            sub_df.fillna("").astype(str).to_excel(writer, sheet_name=safe_name, index=False)
            wrote_any = True
        if not wrote_any:
            pd.DataFrame(columns=["Case_ID", "Lead"]).to_excel(
                writer, sheet_name="MyCases", index=False
            )
    output.seek(0)
    return output.getvalue()


def _escape_html(val):
    s = "" if val is None else str(val)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def render_image_report_card(title, subtitle, kpi_items, breakdown_rows, filename_prefix="HemanVisa_Report"):
    kpi_html_parts = []
    for label, value, color in kpi_items:
        kpi_html_parts.append(
            f"""
            <div style="flex:1; min-width:140px; background:#f8fafc; border:1px solid #e2e8f0;
                        border-top:4px solid {color}; border-radius:10px; padding:12px 14px;">
                <div style="font-size:11.5px; font-weight:700; color:#64748b; text-transform:uppercase;">
                    {_escape_html(label)}
                </div>
                <div style="font-size:22px; font-weight:800; color:#0f172a; margin-top:4px;">
                    {_escape_html(value)}
                </div>
            </div>
            """
        )
    kpi_block = "".join(kpi_html_parts)

    row_html_parts = []
    for idx, (col1, col2, col3) in enumerate(breakdown_rows):
        bg = "#ffffff" if idx % 2 == 0 else "#f8fafc"
        row_html_parts.append(
            f"""
            <tr style="background:{bg};">
                <td style="padding:8px 12px; border-bottom:1px solid #e2e8f0; font-weight:600; color:#1e293b;">{_escape_html(col1)}</td>
                <td style="padding:8px 12px; border-bottom:1px solid #e2e8f0; text-align:center; font-weight:700; color:#0f172a;">{_escape_html(col2)}</td>
                <td style="padding:8px 12px; border-bottom:1px solid #e2e8f0; text-align:right; color:#334155;">{_escape_html(col3)}</td>
            </tr>
            """
        )
    rows_block = "".join(row_html_parts) if row_html_parts else (
        '<tr><td colspan="3" style="padding:12px; text-align:center; color:#64748b;">No records in selected range</td></tr>'
    )

    safe_prefix = json.dumps(str(filename_prefix))
    components.html(
        f"""
        <div id="hv_img_report_box" style="
            font-family: 'Segoe UI', Roboto, Arial, sans-serif;
            background: #ffffff;
            border: 1px solid #cbd5e1;
            border-radius: 14px;
            padding: 18px 20px;
            margin: 8px 0 14px 0;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.06);
        ">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;
                        border-bottom:2px solid #1e3a8a; padding-bottom:10px; margin-bottom:14px;">
                <div>
                    <div style="font-size:18px; font-weight:800; color:#0f172a; letter-spacing:0.3px;">
                        {_escape_html(title)}
                    </div>
                    <div style="font-size:12px; color:#475569; margin-top:2px;">
                        {_escape_html(subtitle)}
                    </div>
                </div>
                <button id="hv_dl_img_btn" type="button" style="
                    background: linear-gradient(135deg, #1d4ed8 0%, #1e3a8a 100%);
                    color: #ffffff;
                    border: none;
                    border-radius: 8px;
                    padding: 8px 14px;
                    font-size: 12.5px;
                    font-weight: 700;
                    cursor: pointer;
                ">🖼️ Download Image</button>
            </div>
            <div style="display:flex; gap:12px; flex-wrap:wrap; margin-bottom:14px;">
                {kpi_block}
            </div>
            <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <thead>
                    <tr style="background:#eff6ff; color:#1e3a8a;">
                        <th style="padding:8px 12px; text-align:left; border-bottom:2px solid #bfdbfe;">Name / Category</th>
                        <th style="padding:8px 12px; text-align:center; border-bottom:2px solid #bfdbfe;">Count</th>
                        <th style="padding:8px 12px; text-align:right; border-bottom:2px solid #bfdbfe;">Details</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_block}
                </tbody>
            </table>
        </div>
        <script>
        (function() {{
            const btn = document.getElementById("hv_dl_img_btn");
            const box = document.getElementById("hv_img_report_box");
            if (!btn || !box) return;
            btn.addEventListener("click", function() {{
                const rect = box.getBoundingClientRect();
                const w = Math.max(600, Math.ceil(rect.width || 800));
                const h = Math.max(320, Math.ceil(rect.height || 450));
                const svgXml = `
                    <svg xmlns="http://www.w3.org/2000/svg" width="${{w}}" height="${{h}}">
                        <foreignObject width="100%" height="100%">
                            <div xmlns="http://www.w3.org/1999/xhtml">${{box.outerHTML}}</div>
                        </foreignObject>
                    </svg>
                `;
                const svgBlob = new Blob([svgXml], {{type: "image/svg+xml;charset=utf-8"}});
                const url = URL.createObjectURL(svgBlob);
                const img = new Image();
                img.onload = function() {{
                    const canvas = document.createElement("canvas");
                    canvas.width = w * 2;
                    canvas.height = h * 2;
                    const ctx = canvas.getContext("2d");
                    ctx.scale(2, 2);
                    ctx.fillStyle = "#ffffff";
                    ctx.fillRect(0, 0, w, h);
                    ctx.drawImage(img, 0, 0);
                    URL.revokeObjectURL(url);
                    canvas.toBlob(function(pngBlob) {{
                        if (!pngBlob) return;
                        const dlUrl = URL.createObjectURL(pngBlob);
                        const a = document.createElement("a");
                        a.href = dlUrl;
                        a.download = {safe_prefix} + ".png";
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        setTimeout(() => URL.revokeObjectURL(dlUrl), 1500);
                    }}, "image/png");
                }};
                img.src = url;
            }});
        }})();
        </script>
        """,
        height=390,
    )


def inject_global_css():
    st.markdown(
        """
        <style>
        .stApp {
            background: linear-gradient(180deg, #f1f5f9 0%, #e2e8f0 100%);
        }
        .hv-header-banner {
            background: linear-gradient(120deg, #0f172a 0%, #1e3a8a 60%, #1d4ed8 100%);
            color: #ffffff;
            padding: 16px 22px;
            border-radius: 14px;
            margin-bottom: 14px;
            box-shadow: 0 8px 20px rgba(15, 23, 42, 0.14);
            display: flex;
            align-items: center;
            justify-content: space-between;
            flex-wrap: wrap;
            gap: 12px;
        }
        .hv-header-title {
            font-size: 22px;
            font-weight: 800;
            letter-spacing: 0.4px;
            margin: 0;
        }
        .hv-header-sub {
            font-size: 12.5px;
            color: #cbd5e1;
            margin-top: 3px;
        }
        .hv-notice-marquee {
            background: #fef3c7;
            border: 1px solid #f59e0b;
            color: #78350f;
            border-radius: 10px;
            padding: 8px 14px;
            margin-bottom: 14px;
            font-weight: 600;
            font-size: 13.5px;
        }
        .hv-kpi-card {
            background: #ffffff;
            border-radius: 12px;
            padding: 14px 16px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
        }
        .hv-kpi-label {
            font-size: 12px;
            font-weight: 700;
            color: #64748b;
            text-transform: uppercase;
        }
        .hv-kpi-value {
            font-size: 24px;
            font-weight: 800;
            color: #0f172a;
            margin-top: 4px;
        }
        .stButton > button {
            border-radius: 8px;
            font-weight: 600;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
# ================= END OF STEP 5 =================
# ================= STEP 6: LOGIN SCREEN, LOGO & TOP HEADER / NOTICE BAR =================
def _get_logo_base64():
    for fname in LOGO_FILES:
        if os.path.exists(fname) and os.path.isfile(fname):
            try:
                with open(fname, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode("utf-8")
                ext = os.path.splitext(fname)[1].lower().lstrip(".")
                mime = "image/jpeg" if ext in ["jpg", "jpeg"] else f"image/{ext or 'png'}"
                return f"data:{mime};base64,{b64}"
            except Exception:
                pass
    return ""


def render_top_header(role_label="", user_label=""):
    logo_uri = _get_logo_base64()
    logo_html = (
        f'<img src="{logo_uri}" alt="Logo" style="height:44px;width:44px;object-fit:cover;border-radius:10px;border:2px solid rgba(255,255,255,0.35);" />'
        if logo_uri
        else '<div style="height:44px;width:44px;border-radius:10px;background:rgba(255,255,255,0.16);display:flex;align-items:center;justify-content:center;font-size:22px;">✈️</div>'
    )
    badge_text = " • ".join([x for x in [role_label, user_label] if x])
    badge_html = (
        f'<span style="background:rgba(255,255,255,0.16);border:1px solid rgba(255,255,255,0.28);padding:5px 12px;border-radius:999px;font-size:12.5px;font-weight:700;">{badge_text}</span>'
        if badge_text
        else ""
    )
    st.markdown(
        f"""
        <div class="hv-header-banner">
            <div style="display:flex;align-items:center;gap:14px;">
                {logo_html}
                <div>
                    <div class="hv-header-title">HemanVisa Case & Appointment Portal</div>
                    <div class="hv-header-sub">Case Tracking • Lead Distribution • Appointment & Reschedule Management</div>
                </div>
            </div>
            <div>{badge_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    notice_txt = get_notice()
    if notice_txt:
        st.markdown(
            f"""
            <div class="hv-notice-marquee">
                <marquee behavior="scroll" direction="left" scrollamount="5">📢 {_escape_html(notice_txt)}</marquee>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_login_screen():
    render_top_header("Portal Login", "")
    all_sheets = load_all_sheets_cached()
    workers = get_workers_list(all_sheets)
    logins = load_worker_logins()
    admin_creds = get_admin_creds()

    col_l, col_c, col_r = st.columns([1, 1.4, 1])
    with col_c:
        st.markdown("### 🔐 Sign In to HemanVisa Portal")
        login_mode = st.radio(
            "Select Login Role",
            ["Case Worker", "Admin"],
            horizontal=True,
            key="login_role_radio",
        )

        if login_mode == "Admin":
            with st.form("admin_login_form", clear_on_submit=False):
                adm_user = st.text_input("Admin Username", key="adm_user_input")
                adm_pass = st.text_input("Admin Password", type="password", key="adm_pass_input")
                submitted = st.form_submit_button("Login as Admin", use_container_width=True)
                if submitted:
                    if (
                        adm_user.strip() == admin_creds.get("user", "admin")
                        and adm_pass.strip() == admin_creds.get("pass", "heman123")
                    ):
                        st.session_state["logged_in"] = True
                        st.session_state["role"] = "admin"
                        st.session_state["username"] = adm_user.strip()
                        st.rerun()
                    else:
                        st.error("Invalid Admin credentials.")
        else:
            with st.form("worker_login_form", clear_on_submit=False):
                selected_worker = st.selectbox(
                    "Select Case Worker",
                    workers if workers else ["Kiran", "Shanthi", "Keerthana", "Sai", "Kalyan", "Anil"],
                    key="worker_select_login",
                )
                w_pass = st.text_input("Worker Password", type="password", key="worker_pass_input")
                submitted = st.form_submit_button("Login as Case Worker", use_container_width=True)
                if submitted:
                    w_info = logins.get(selected_worker, {})
                    saved_pw = str(w_info.get("password", "")).strip()
                    # Allow login if password matches saved password, or if no custom password is set yet
                    if not saved_pw or w_pass.strip() == saved_pw:
                        st.session_state["logged_in"] = True
                        st.session_state["role"] = "worker"
                        st.session_state["username"] = selected_worker
                        record_worker_login_time(selected_worker)
                        st.rerun()
                    else:
                        st.error("Invalid Case Worker password.")
# ================= END OF STEP 6 =================
# ================= STEP 7: CASE SEARCH, FILTER, EDIT/RESCHEDULE/PDF DIALOG & ADD CASE =================
def _find_col_ci(df, candidates):
    if df is None or df.empty:
        return None
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = str(cand).strip().lower()
        if key in lower_map:
            return lower_map[key]
    for c in df.columns:
        c_low = str(c).strip().lower()
        for cand in candidates:
            if str(cand).strip().lower() in c_low:
                return c
    return None


def filter_cases_dataframe(
    df,
    search_text="",
    lead_filter="All",
    status_filter="All",
    date_from=None,
    date_to=None,
    worker_only=None,
):
    if df is None or df.empty:
        return pd.DataFrame(columns=["Case_ID", "Username", "USERNAME", "Password", "Lead"], dtype=str)

    out = df.copy()
    out["__orig_row_idx__"] = out.index

    lead_col = _find_col_ci(out, ["Lead", "Worker", "Case Worker", "Assigned To", "Handler"])
    if worker_only and lead_col:
        w_low = str(worker_only).strip().lower()
        out = out[out[lead_col].astype(str).str.strip().str.lower() == w_low]

    if lead_filter and str(lead_filter).strip().lower() != "all" and lead_col:
        lf_low = str(lead_filter).strip().lower()
        out = out[out[lead_col].astype(str).str.strip().str.lower() == lf_low]

    status_col = _find_col_ci(out, ["Status", "Booking_Status", "Booking Status", "Case Status", "State"])
    if status_filter and str(status_filter).strip().lower() != "all" and status_col:
        sf_low = str(status_filter).strip().lower()
        out = out[out[status_col].astype(str).str.strip().str.lower() == sf_low]

    if date_from is not None or date_to is not None:
        date_col = _find_col_ci(
            out,
            ["Appointment_Date", "Appointment Date", "Appt_Date", "Date", "Booking_Date", "Created_At", "Updated_At"],
        )
        if date_col:
            parsed_dates = pd.to_datetime(out[date_col], errors="coerce").dt.date
            mask = pd.Series(True, index=out.index)
            if date_from is not None:
                mask = mask & (parsed_dates >= date_from)
            if date_to is not None:
                mask = mask & (parsed_dates <= date_to)
            out = out[mask.fillna(False)]

    q = str(search_text or "").strip().lower()
    if q:
        q_mask = pd.Series(False, index=out.index)
        for col in out.columns:
            if col == "__orig_row_idx__":
                continue
            col_s = out[col].astype(str).str.lower()
            q_mask = q_mask | col_s.str.contains(re.escape(q), na=False)
        out = out[q_mask]

    return out


def _get_case_identifier(row_series, fallback_idx=0):
    if row_series is None:
        return f"HM{1111 + int(fallback_idx)}"
    for k in row_series.index:
        k_low = str(k).strip().lower()
        if k_low in ["case_id", "case id", "caseid", "id"]:
            val = _clean_cell_text(row_series.get(k, ""))
            if val:
                return val
    return f"HM{1111 + int(fallback_idx)}"


@st.dialog("📁 Case Details, Edit, Reschedule & Appointment PDFs", width="large")
def open_case_action_dialog(sheet_name, row_idx, actor_role, actor_name, can_edit=True, can_delete=False):
    sheets_map = load_all_sheets_cached()
    if sheet_name not in sheets_map or sheets_map[sheet_name] is None or sheets_map[sheet_name].empty:
        st.error("Sheet data not found.")
        return

    df_sheet = sheets_map[sheet_name].copy()
    if row_idx not in df_sheet.index:
        st.error("Selected case row was not found.")
        return

    row_series = df_sheet.loc[row_idx]
    case_id = _get_case_identifier(row_series, row_idx)

    st.markdown(f"### 🆔 Case ID: `{case_id}` — Sheet: `{sheet_name}`")

    edit_reqs = load_edit_requests()
    approved_for_case = any(
        str(r.get("Case_ID", "")).strip().lower() == case_id.strip().lower()
        and str(r.get("Status", "")).strip().lower() == "approved"
        for r in edit_reqs
    )
    effective_can_edit = bool(can_edit) or (actor_role == "worker" and approved_for_case)

    tab_view_edit, tab_reschedule, tab_pdf, tab_msg = st.tabs(
        ["✏️ View & Edit Case", "🗓️ Reschedule Case", "📄 Appointment PDFs", "✉️ CGI Message"]
    )

    with tab_view_edit:
        if not effective_can_edit:
            st.warning("Editing is currently restricted for this case.")
            with st.form(f"req_edit_form_{sheet_name}_{row_idx}"):
                target_field = st.text_input("Field Name to Edit", key=f"req_fld_{sheet_name}_{row_idx}")
                proposed_val = st.text_input("Proposed New Value", key=f"req_val_{sheet_name}_{row_idx}")
                reason_txt = st.text_area("Reason for Edit Request", key=f"req_rsn_{sheet_name}_{row_idx}")
                req_sub = st.form_submit_button("📩 Submit Edit Request to Admin", use_container_width=True)
                if req_sub:
                    if not reason_txt.strip():
                        st.error("Please provide a reason for the edit request.")
                    else:
                        u_col = _find_col_ci(df_sheet, ["USERNAME", "Username", "Case Username"])
                        c_user = _clean_cell_text(row_series.get(u_col, "")) if u_col else ""
                        req_id = f"REQ_{get_ist_now().strftime('%Y%m%d%H%M%S')}_{row_idx}"
                        edit_reqs.append(
                            {
                                "Request_ID": req_id,
                                "Case_ID": case_id,
                                "Case_Username": c_user,
                                "Worker": actor_name,
                                "Field_Name": target_field.strip(),
                                "Old_Value": _clean_cell_text(row_series.get(target_field.strip(), "")) if target_field.strip() in df_sheet.columns else "",
                                "New_Value": proposed_val.strip(),
                                "Reason": reason_txt.strip(),
                                "Requested_At": get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p"),
                                "Status": "Pending",
                                "Reviewed_By": "",
                                "Reviewed_At": "",
                            }
                        )
                        save_edit_requests(edit_reqs)
                        st.success("Edit request submitted to Admin.")
                        st.rerun()
        else:
            editable_cols = [c for c in df_sheet.columns if c != "__orig_row_idx__"]
            with st.form(f"edit_case_form_{sheet_name}_{row_idx}"):
                updated_vals = {}
                cols_per_row = 2
                for i in range(0, len(editable_cols), cols_per_row):
                    f_cols = st.columns(cols_per_row)
                    for j in range(cols_per_row):
                        if i + j < len(editable_cols):
                            col_name = editable_cols[i + j]
                            curr_val = _clean_cell_text(row_series.get(col_name, ""))
                            with f_cols[j]:
                                updated_vals[col_name] = st.text_input(
                                    col_name,
                                    value=curr_val,
                                    key=f"edit_inp_{sheet_name}_{row_idx}_{col_name}",
                                )

                save_btn = st.form_submit_button("💾 Save Case Changes", type="primary", use_container_width=True)
                if save_btn:
                    for col_name, val in updated_vals.items():
                        df_sheet.at[row_idx, col_name] = str(val).strip()
                    upd_col = _find_col_ci(df_sheet, ["Updated_At", "Updated At", "Last_Updated"])
                    by_col = _find_col_ci(df_sheet, ["Updated_By", "Updated By", "Modified_By"])
                    if upd_col:
                        df_sheet.at[row_idx, upd_col] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    if by_col:
                        df_sheet.at[row_idx, by_col] = actor_name
                    sheets_map[sheet_name] = df_sheet
                    save_all_sheets(sheets_map, push_to_cloud=True)
                    st.session_state["active_dialog_case"] = None
                    st.success("Case updated and synced.")
                    st.rerun()

        if can_delete:
            st.markdown("---")
            if st.button("🗑️ Delete This Case", key=f"del_case_btn_{sheet_name}_{row_idx}", use_container_width=True):
                df_sheet = df_sheet.drop(index=row_idx).reset_index(drop=True)
                sheets_map[sheet_name] = df_sheet
                save_all_sheets(sheets_map, push_to_cloud=True)
                st.session_state["active_dialog_case"] = None
                st.success("Case deleted.")
                st.rerun()

    with tab_reschedule:
        st.markdown("#### 🗓️ Reschedule Appointment & Assign Case Worker")
        workers_list = get_workers_list(sheets_map)
        curr_lead = _clean_cell_text(
            row_series.get(_find_col_ci(df_sheet, ["Lead", "Worker", "Case Worker"]) or "", "")
        )
        default_worker_idx = workers_list.index(curr_lead) if curr_lead in workers_list else 0

        appt_date_col = _find_col_ci(df_sheet, ["Appointment_Date", "Appointment Date", "Appt_Date", "Date"])
        appt_time_col = _find_col_ci(df_sheet, ["Appointment_Time", "Appointment Time", "Appt_Time", "Time"])
        status_col = _find_col_ci(df_sheet, ["Status", "Booking_Status", "Booking Status", "Case Status"])
        notes_col = _find_col_ci(df_sheet, ["Notes", "Remarks", "Remark", "Reschedule_Reason", "Comment"])

        curr_date_val = _clean_cell_text(row_series.get(appt_date_col, "")) if appt_date_col else ""
        parsed_d = pd.to_datetime(curr_date_val, errors="coerce")
        init_date = parsed_d.date() if pd.notna(parsed_d) else get_ist_now().date()

        with st.form(f"resched_form_{sheet_name}_{row_idx}"):
            new_worker = st.selectbox(
                "Assigned Case Worker / Lead",
                workers_list if workers_list else ["Kiran", "Shanthi"],
                index=default_worker_idx,
                key=f"resched_w_{sheet_name}_{row_idx}",
            )
            new_date = st.date_input("New Appointment Date", value=init_date, key=f"resched_d_{sheet_name}_{row_idx}")
            new_time = st.text_input(
                "New Appointment Time",
                value=_clean_cell_text(row_series.get(appt_time_col, "")) if appt_time_col else "",
                placeholder="e.g., 10:30 AM IST",
                key=f"resched_t_{sheet_name}_{row_idx}",
            )
            new_status = st.selectbox(
                "Case Status",
                ["Rescheduled", "Booked", "Pending", "Completed", "Cancelled", "On Hold"],
                index=0,
                key=f"resched_s_{sheet_name}_{row_idx}",
            )
            resched_note = st.text_area(
                "Reschedule Note / Reason",
                value=_clean_cell_text(row_series.get(notes_col, "")) if notes_col else "",
                key=f"resched_n_{sheet_name}_{row_idx}",
            )
            sub_resched = st.form_submit_button("✅ Confirm Reschedule", type="primary", use_container_width=True)
            if sub_resched:
                lead_c = _find_col_ci(df_sheet, ["Lead", "Worker", "Case Worker"])
                if lead_c:
                    df_sheet.at[row_idx, lead_c] = new_worker
                if appt_date_col:
                    df_sheet.at[row_idx, appt_date_col] = new_date.strftime("%Y-%m-%d")
                if appt_time_col:
                    df_sheet.at[row_idx, appt_time_col] = new_time.strip()
                if status_col:
                    df_sheet.at[row_idx, status_col] = new_status
                if notes_col and resched_note.strip():
                    df_sheet.at[row_idx, notes_col] = resched_note.strip()
                upd_c = _find_col_ci(df_sheet, ["Updated_At", "Updated At"])
                by_c = _find_col_ci(df_sheet, ["Updated_By", "Updated By"])
                if upd_c:
                    df_sheet.at[row_idx, upd_c] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                if by_c:
                    df_sheet.at[row_idx, by_c] = actor_name

                sheets_map[sheet_name] = df_sheet
                save_all_sheets(sheets_map, push_to_cloud=True)
                st.session_state["active_dialog_case"] = None
                st.success(f"Case {case_id} rescheduled to {new_date.strftime('%Y-%m-%d')} ({new_worker}).")
                st.rerun()

    with tab_pdf:
        st.markdown("#### 📄 Appointment Confirmation PDFs")
        up_pdf = st.file_uploader(
            "Upload Appointment PDF",
            type=["pdf"],
            key=f"up_pdf_{sheet_name}_{row_idx}",
        )
        if up_pdf is not None:
            saved_fname = save_case_pdf_file(case_id, up_pdf)
            if saved_fname:
                st.success(f"Uploaded `{saved_fname}` for `{case_id}`.")
                st.rerun()

        existing_pdfs = list_case_pdfs(case_id)
        if not existing_pdfs:
            st.info("No PDF files attached for this case yet.")
        else:
            folder_path = get_case_pdf_folder(case_id, create=False)
            for fname in existing_pdfs:
                fpath = os.path.join(folder_path, fname)
                c1, c2, c3 = st.columns([3, 1.2, 1.2])
                with c1:
                    st.write(f"📎 `{fname}`")
                with c2:
                    try:
                        with open(fpath, "rb") as pf:
                            st.download_button(
                                "⬇️ Download",
                                data=pf.read(),
                                file_name=fname,
                                mime="application/pdf",
                                key=f"dl_pdf_{case_id}_{fname}_{row_idx}",
                                use_container_width=True,
                            )
                    except Exception:
                        st.write("Unavailable")
                with c3:
                    if can_delete or actor_role == "admin":
                        if st.button("🗑️ Remove", key=f"rm_pdf_{case_id}_{fname}_{row_idx}", use_container_width=True):
                            delete_case_pdf_file(case_id, fname)
                            st.rerun()

    with tab_msg:
        st.markdown("#### ✉️ Formatted CGI Portal Credentials Message")
        msg_txt = build_cgi_formatted_message(row_series)
        st.text_area("Message Preview", value=msg_txt, height=190, key=f"cgi_ta_{sheet_name}_{row_idx}")
        render_copy_message_button(msg_txt, btn_key=f"copy_btn_{sheet_name}_{row_idx}")


def render_add_case_form(target_sheet, actor_role, actor_name, key_prefix="add_case"):
    sheets_map = load_all_sheets_cached()
    if target_sheet not in sheets_map or sheets_map[target_sheet] is None:
        st.error("Target sheet not available.")
        return

    df_sheet = sheets_map[target_sheet].copy()
    next_id = get_next_case_id(sheets_map)
    workers_list = get_workers_list(sheets_map)
    default_lead = actor_name if (actor_role == "worker" and actor_name in workers_list) else (workers_list[0] if workers_list else "Kiran")

    st.markdown(f"#### ➕ Add New Case to `{target_sheet}` (Auto ID: `{next_id}`)")
    base_cols = [c for c in df_sheet.columns if c != "__orig_row_idx__"]
    if not base_cols:
        base_cols = ["Case_ID", "Username", "USERNAME", "Password", "Lead", "Status", "Appointment_Date", "Appointment_Time", "Notes"]

    with st.form(f"{key_prefix}_form_{target_sheet}", clear_on_submit=True):
        field_inputs = {}
        for i in range(0, len(base_cols), 2):
            f_cols = st.columns(2)
            for j in range(2):
                if i + j < len(base_cols):
                    col_name = base_cols[i + j]
                    col_low = str(col_name).strip().lower()
                    with f_cols[j]:
                        if col_low in ["case_id", "case id", "caseid"]:
                            field_inputs[col_name] = st.text_input(
                                col_name,
                                value=next_id,
                                disabled=True,
                                key=f"{key_prefix}_{col_name}_{target_sheet}",
                            )
                            field_inputs[col_name] = next_id
                        elif col_low in ["lead", "worker", "case worker", "assigned to"]:
                            idx_w = workers_list.index(default_lead) if default_lead in workers_list else 0
                            field_inputs[col_name] = st.selectbox(
                                col_name,
                                workers_list if workers_list else [default_lead],
                                index=idx_w,
                                key=f"{key_prefix}_{col_name}_{target_sheet}",
                            )
                        elif col_low in ["status", "booking_status", "booking status", "case status"]:
                            field_inputs[col_name] = st.selectbox(
                                col_name,
                                ["Active", "Booked", "Rescheduled", "Pending", "Completed"],
                                index=0,
                                key=f"{key_prefix}_{col_name}_{target_sheet}",
                            )
                        else:
                            field_inputs[col_name] = st.text_input(
                                col_name,
                                value="",
                                key=f"{key_prefix}_{col_name}_{target_sheet}",
                            )

        submitted = st.form_submit_button("➕ Create Case Record", type="primary", use_container_width=True)
        if submitted:
            new_row = {c: "" for c in df_sheet.columns if c != "__orig_row_idx__"}
            for c, val in field_inputs.items():
                if c in new_row:
                    new_row[c] = str(val).strip()
            id_col = _find_col_ci(df_sheet, ["Case_ID", "Case ID", "caseid"])
            if id_col:
                new_row[id_col] = get_next_case_id(sheets_map)
            created_col = _find_col_ci(df_sheet, ["Created_At", "Created At", "Date_Created"])
            if created_col and not new_row.get(created_col):
                new_row[created_col] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
            by_col = _find_col_ci(df_sheet, ["Created_By", "Updated_By", "Updated By"])
            if by_col:
                new_row[by_col] = actor_name

            df_new = pd.concat([df_sheet, pd.DataFrame([new_row])], ignore_index=True)
            sheets_map[target_sheet] = df_new
            save_all_sheets(sheets_map, push_to_cloud=True)
            st.success(f"Added new case `{new_row.get(id_col, next_id)}` to `{target_sheet}`.")
            st.rerun()
# ================= END OF STEP 7 =================
# ================= STEP 8: CASES TABLE, KPI CARDS & LEAD / DATE SLICER WIDGETS =================
def render_kpi_row(df, all_sheets=None):
    total_cnt = len(df) if df is not None else 0
    booked_cnt = 0
    resched_cnt = 0
    today_cnt = 0
    today_str = get_ist_now().strftime("%Y-%m-%d")

    if df is not None and not df.empty:
        status_col = _find_col_ci(df, ["Status", "Booking_Status", "Booking Status", "Case Status", "State"])
        if status_col:
            s_vals = df[status_col].astype(str).str.strip().str.lower()
            booked_cnt = int(s_vals.str.contains("book", na=False).sum())
            resched_cnt = int(s_vals.str.contains("resched", na=False).sum())

        date_col = _find_col_ci(
            df,
            ["Appointment_Date", "Appointment Date", "Appt_Date", "Date", "Booking_Date", "Updated_At", "Created_At"],
        )
        if date_col:
            d_vals = pd.to_datetime(df[date_col], errors="coerce").dt.strftime("%Y-%m-%d")
            today_cnt = int((d_vals == today_str).sum())

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f'<div class="hv-kpi-card"><div class="hv-kpi-label">Total Cases</div><div class="hv-kpi-value">{total_cnt}</div></div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="hv-kpi-card"><div class="hv-kpi-label">Booked</div><div class="hv-kpi-value">{booked_cnt}</div></div>',
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f'<div class="hv-kpi-card"><div class="hv-kpi-label">Rescheduled</div><div class="hv-kpi-value">{resched_cnt}</div></div>',
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f'<div class="hv-kpi-card"><div class="hv-kpi-label">Today ({today_str})</div><div class="hv-kpi-value">{today_cnt}</div></div>',
            unsafe_allow_html=True,
        )


def render_filter_bar(df, key_prefix="flt", allow_worker_select=True, default_lead="All"):
    all_sheets = load_all_sheets_cached()
    workers_list = get_workers_list(all_sheets)

    lead_col = _find_col_ci(df, ["Lead", "Worker", "Case Worker", "Assigned To", "Handler"])
    df_leads = []
    if df is not None and not df.empty and lead_col:
        df_leads = [
            x for x in sorted(df[lead_col].dropna().astype(str).str.strip().unique())
            if x and x.lower() not in ["nan", "none", "null"]
        ]
    lead_choices = ["All"] + sorted(list(set(df_leads + workers_list)))

    status_col = _find_col_ci(df, ["Status", "Booking_Status", "Booking Status", "Case Status", "State"])
    df_statuses = []
    if df is not None and not df.empty and status_col:
        df_statuses = [
            x for x in sorted(df[status_col].dropna().astype(str).str.strip().unique())
            if x and x.lower() not in ["nan", "none", "null"]
        ]
    base_statuses = ["Active", "Booked", "Rescheduled", "Pending", "Completed", "Cancelled", "On Hold"]
    status_choices = ["All"] + sorted(list(set(base_statuses + df_statuses)))

    r1_1, r1_2, r1_3 = st.columns([2.2, 1.3, 1.3])
    with r1_1:
        search_val = st.text_input(
            "Search Cases (Case ID, Username, Lead, DS-160, Status, Notes)",
            value="",
            placeholder="Type to search...",
            key=f"{key_prefix}_search",
        )
    with r1_2:
        if allow_worker_select:
            def_idx = lead_choices.index(default_lead) if default_lead in lead_choices else 0
            lead_val = st.selectbox("Lead / Case Worker", lead_choices, index=def_idx, key=f"{key_prefix}_lead")
        else:
            lead_val = default_lead
            st.text_input("Lead / Case Worker", value=str(default_lead), disabled=True, key=f"{key_prefix}_lead_ro")
    with r1_3:
        status_val = st.selectbox("Case Status", status_choices, index=0, key=f"{key_prefix}_status")

    r2_1, r2_2, r2_3 = st.columns([1.2, 1.2, 2.4])
    with r2_1:
        use_from = st.checkbox("From Date", value=False, key=f"{key_prefix}_use_from")
        date_from = st.date_input("Start Date", value=get_ist_now().date(), disabled=not use_from, key=f"{key_prefix}_d_from")
    with r2_2:
        use_to = st.checkbox("To Date", value=False, key=f"{key_prefix}_use_to")
        date_to = st.date_input("End Date", value=get_ist_now().date(), disabled=not use_to, key=f"{key_prefix}_d_to")
    with r2_3:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        if st.button("🧹 Reset Filters", key=f"{key_prefix}_reset_btn"):
            for k in [
                f"{key_prefix}_search",
                f"{key_prefix}_lead",
                f"{key_prefix}_status",
                f"{key_prefix}_use_from",
                f"{key_prefix}_use_to",
            ]:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()

    return {
        "search": search_val,
        "lead": lead_val,
        "status": status_val,
        "date_from": date_from if use_from else None,
        "date_to": date_to if use_to else None,
    }


def render_cases_interactive_table(
    df_filtered,
    sheet_name,
    actor_role,
    actor_name,
    can_edit=True,
    can_delete=False,
    key_prefix="tbl",
):
    if df_filtered is None or df_filtered.empty:
        st.info("No matching case records found.")
        return

    visible_cols = [c for c in df_filtered.columns if c != "__orig_row_idx__"]
    st.dataframe(df_filtered[visible_cols], use_container_width=True, hide_index=True)

    st.markdown("##### 🎯 Select a Case Row to Open Action Popup (Edit / Reschedule / PDF / Copy)")
    row_options = []
    idx_lookup = {}
    for disp_i, (df_i, row_s) in enumerate(df_filtered.iterrows()):
        orig_i = row_s.get("__orig_row_idx__", df_i)
        cid = _get_case_identifier(row_s, orig_i)
        uname = _clean_cell_text(
            row_s.get(_find_col_ci(df_filtered, ["USERNAME", "Username", "Case Username"]) or "", "")
        )
        lead_n = _clean_cell_text(
            row_s.get(_find_col_ci(df_filtered, ["Lead", "Worker", "Case Worker"]) or "", "")
        )
        stat_n = _clean_cell_text(
            row_s.get(_find_col_ci(df_filtered, ["Status", "Booking_Status", "Booking Status"]) or "", "")
        )
        label = f"#{disp_i + 1} | {cid}"
        if uname:
            label += f" | {uname}"
        if lead_n:
            label += f" | Lead: {lead_n}"
        if stat_n:
            label += f" | {stat_n}"
        row_options.append(label)
        idx_lookup[label] = orig_i

    sel_col, btn_col = st.columns([3.5, 1.2])
    with sel_col:
        chosen_label = st.selectbox(
            "Choose Case",
            row_options,
            label_visibility="collapsed",
            key=f"{key_prefix}_case_picker_{sheet_name}",
        )
    with btn_col:
        if st.button("📂 Open Case", type="primary", use_container_width=True, key=f"{key_prefix}_open_btn_{sheet_name}"):
            target_orig_idx = idx_lookup.get(chosen_label)
            if target_orig_idx is not None:
                st.session_state["active_dialog_case"] = {
                    "sheet": sheet_name,
                    "row_idx": target_orig_idx,
                    "role": actor_role,
                    "user": actor_name,
                    "can_edit": can_edit,
                    "can_delete": can_delete,
                }
                st.rerun()
# ================= END OF STEP 8 =================
# ================= STEP 9: ADMIN DASHBOARD & ADMIN MANAGEMENT DIALOGS =================
@st.dialog("👥 Manage Case Worker Permissions & Passwords", width="large")
def open_admin_permissions_dialog(all_sheets):
    workers = get_workers_list(all_sheets)
    perms = load_permissions()
    logins = load_worker_logins()

    for w_name in workers:
        cur_p = perms.get(
            w_name,
            {
                "can_edit": True,
                "can_delete": False,
                "can_add_case": True,
                "can_view_all": False,
                "can_export": True,
            },
        )
        cur_pw = str(logins.get(w_name, {}).get("password", "")).strip()
        with st.expander(f"👷 {w_name}", expanded=False):
            p_c1, p_c2, p_c3, p_c4, p_c5 = st.columns(5)
            with p_c1:
                e_val = st.checkbox("Can Edit", value=bool(cur_p.get("can_edit", True)), key=f"prm_e_{w_name}")
            with p_c2:
                d_val = st.checkbox("Can Delete", value=bool(cur_p.get("can_delete", False)), key=f"prm_d_{w_name}")
            with p_c3:
                a_val = st.checkbox("Can Add Case", value=bool(cur_p.get("can_add_case", True)), key=f"prm_a_{w_name}")
            with p_c4:
                v_val = st.checkbox("View All Cases", value=bool(cur_p.get("can_view_all", False)), key=f"prm_v_{w_name}")
            with p_c5:
                x_val = st.checkbox("Can Export", value=bool(cur_p.get("can_export", True)), key=f"prm_x_{w_name}")

            new_pw = st.text_input(
                f"Worker Password ({w_name}) — leave blank to keep current",
                value=cur_pw,
                type="password",
                key=f"prm_pw_{w_name}",
            )

            if st.button(f"💾 Save {w_name} Settings", key=f"prm_save_{w_name}", use_container_width=True):
                perms[w_name] = {
                    "can_edit": e_val,
                    "can_delete": d_val,
                    "can_add_case": a_val,
                    "can_view_all": v_val,
                    "can_export": x_val,
                }
                save_permissions(perms)

                if w_name not in logins:
                    logins[w_name] = {
                        "password": new_pw.strip(),
                        "last_login": "",
                        "last_logout": "",
                        "status": "Offline",
                        "last_seen": "",
                    }
                else:
                    logins[w_name]["password"] = new_pw.strip()
                save_worker_logins(logins)
                st.success(f"Saved permissions and password for {w_name}.")
                st.rerun()


@st.dialog("📝 Edit Requests Approval Queue", width="large")
def open_edit_requests_dialog(all_sheets):
    reqs = load_edit_requests()
    admin_user = st.session_state.get("username", "admin")
    if not reqs:
        st.info("No edit requests found.")
        return

    df_reqs = pd.DataFrame(reqs)
    st.dataframe(df_reqs, use_container_width=True, hide_index=True)

    pending_indices = [i for i, r in enumerate(reqs) if str(r.get("Status", "")).strip().lower() == "pending"]
    if not pending_indices:
        st.success("All edit requests have been reviewed.")
        return

    for idx in pending_indices:
        r = reqs[idx]
        with st.container(border=True):
            st.markdown(
                f"**{r.get('Request_ID', '')}** • Case `{r.get('Case_ID', '')}` • Worker: **{r.get('Worker', '')}**"
            )
            st.write(f"Field: `{r.get('Field_Name', '')}` | Old: `{r.get('Old_Value', '')}` → New: `{r.get('New_Value', '')}`")
            st.write(f"Reason: {r.get('Reason', '')}")
            b1, b2 = st.columns(2)
            with b1:
                if st.button("✅ Approve", key=f"app_req_{idx}", type="primary", use_container_width=True):
                    reqs[idx]["Status"] = "Approved"
                    reqs[idx]["Reviewed_By"] = admin_user
                    reqs[idx]["Reviewed_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    fld = str(r.get("Field_Name", "")).strip()
                    cid = str(r.get("Case_ID", "")).strip()
                    nval = str(r.get("New_Value", "")).strip()
                    if fld and cid:
                        for s_name, df_s in all_sheets.items():
                            if s_name in SYSTEM_SHEETS or df_s is None or df_s.empty:
                                continue
                            id_col = _find_col_ci(df_s, ["Case_ID", "Case ID", "caseid"])
                            if id_col and fld in df_s.columns:
                                match_mask = df_s[id_col].astype(str).str.strip().str.lower() == cid.lower()
                                if match_mask.any():
                                    df_s.loc[match_mask, fld] = nval
                                    all_sheets[s_name] = df_s
                                    save_all_sheets(all_sheets, push_to_cloud=True)
                                    break
                    save_edit_requests(reqs)
                    st.rerun()
            with b2:
                if st.button("❌ Reject", key=f"rej_req_{idx}", use_container_width=True):
                    reqs[idx]["Status"] = "Rejected"
                    reqs[idx]["Reviewed_By"] = admin_user
                    reqs[idx]["Reviewed_At"] = get_ist_now().strftime("%Y-%m-%d %I:%M:%S %p")
                    save_edit_requests(reqs)
                    st.rerun()


@st.dialog("🕒 Worker Logins & Attendance Logs", width="large")
def open_attendance_dialog(all_sheets):
    workers = get_workers_list(all_sheets)
    logins = load_worker_logins()
    att = load_attendance()

    st.markdown("#### 🟢 Current Worker Login Status")
    login_rows = []
    for w in workers:
        info = logins.get(w, {})
        login_rows.append(
            {
                "Worker": w,
                "Status": info.get("status", "Offline"),
                "Last_Login": info.get("last_login", ""),
                "Last_Logout": info.get("last_logout", ""),
                "Last_Seen": info.get("last_seen", ""),
            }
        )
    if login_rows:
        st.dataframe(pd.DataFrame(login_rows), use_container_width=True, hide_index=True)

    st.markdown("#### 📅 Daily Attendance & Working Hours")
    if att:
        st.dataframe(pd.DataFrame(att), use_container_width=True, hide_index=True)
    else:
        st.info("No attendance records logged yet.")


@st.dialog("📢 Scrolling Notice & Admin Password Settings", width="medium")
def open_admin_settings_dialog():
    curr_notice = get_notice()
    admin_creds = get_admin_creds()

    with st.form("admin_notice_settings_form"):
        new_notice = st.text_area("Scrolling Marquee Notice (Visible to All Users)", value=curr_notice, height=90)
        st.markdown("---")
        new_adm_user = st.text_input("Admin Username", value=admin_creds.get("user", "admin"))
        new_adm_pass = st.text_input("Admin Password", value=admin_creds.get("pass", "heman123"), type="password")
        saved = st.form_submit_button("💾 Save Settings", type="primary", use_container_width=True)
        if saved:
            save_notice(new_notice)
            if new_adm_user.strip() and new_adm_pass.strip():
                save_admin_creds(new_adm_user.strip(), new_adm_pass.strip())
            st.success("Admin settings and scrolling notice saved.")
            st.rerun()


@st.dialog("📂 All Case Appointment PDFs Repository", width="large")
def open_all_pdfs_dialog(all_sheets):
    search_pdf_q = st.text_input("Filter PDFs by Case ID or Filename", value="", key="adm_pdf_filter_q").strip().lower()
    if not os.path.exists(PDF_BASE_DIR):
        st.info("No appointment PDFs uploaded yet.")
        return

    subfolders = sorted([d for d in os.listdir(PDF_BASE_DIR) if os.path.isdir(os.path.join(PDF_BASE_DIR, d))])
    found_any = False
    for folder_name in subfolders:
        if search_pdf_q and search_pdf_q not in folder_name.lower():
            continue
        pdfs = list_case_pdfs(folder_name)
        matching_pdfs = [p for p in pdfs if not search_pdf_q or search_pdf_q in p.lower() or search_pdf_q in folder_name.lower()]
        if matching_pdfs:
            found_any = True
            with st.expander(f"📁 Case: {folder_name} ({len(matching_pdfs)} PDFs)", expanded=True):
                f_path = get_case_pdf_folder(folder_name, create=False)
                for p_name in matching_pdfs:
                    full_p = os.path.join(f_path, p_name)
                    c1, c2 = st.columns([3.5, 1.2])
                    with c1:
                        st.write(f"📄 `{p_name}`")
                    with c2:
                        try:
                            with open(full_p, "rb") as pf:
                                st.download_button(
                                    "⬇️ Download",
                                    data=pf.read(),
                                    file_name=p_name,
                                    mime="application/pdf",
                                    key=f"adm_dl_pdf_{folder_name}_{p_name}",
                                    use_container_width=True,
                                )
                        except Exception:
                            st.write("Error")
    if not found_any:
        st.info("No matching PDFs found.")


def render_admin_dashboard(all_sheets):
    admin_user = st.session_state.get("username", "admin")
    render_top_header("Admin Portal", admin_user)

    case_sheet_names = [s for s in all_sheets.keys() if s not in SYSTEM_SHEETS]
    if not case_sheet_names:
        case_sheet_names = ["Active Cases"]
        all_sheets["Active Cases"] = _normalize_sheet_df(pd.DataFrame())

    active_sheet = st.session_state.get("admin_active_sheet", case_sheet_names[0])
    if active_sheet not in case_sheet_names:
        active_sheet = case_sheet_names[0]
        st.session_state["admin_active_sheet"] = active_sheet

    # Top Admin Action Bar
    a_c1, a_c2, a_c3, a_c4, a_c5, a_c6 = st.columns(6)
    with a_c1:
        if st.button("👥 Permissions", key="adm_btn_perms", use_container_width=True):
            st.session_state["admin_modal"] = "permissions"
            st.rerun()
    with a_c2:
        pending_cnt = sum(1 for r in load_edit_requests() if str(r.get("Status", "")).strip().lower() == "pending")
        if st.button(f"📝 Edit Reqs ({pending_cnt})", key="adm_btn_reqs", use_container_width=True):
            st.session_state["admin_modal"] = "edit_requests"
            st.rerun()
    with a_c3:
        if st.button("🕒 Attendance", key="adm_btn_att", use_container_width=True):
            st.session_state["admin_modal"] = "attendance"
            st.rerun()
    with a_c4:
        if st.button("📄 All PDFs", key="adm_btn_pdfs", use_container_width=True):
            st.session_state["admin_modal"] = "pdfs"
            st.rerun()
    with a_c5:
        if st.button("📢 Notice & Auth", key="adm_btn_cfg", use_container_width=True):
            st.session_state["admin_modal"] = "settings"
            st.rerun()
    with a_c6:
        excel_bytes = get_admin_gsheet_export_bytes(_get_file_mtime())
        st.download_button(
            "📊 Export Excel",
            data=excel_bytes,
            file_name=f"HemanVisa_Admin_{get_ist_now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="adm_btn_export",
            use_container_width=True,
        )

    s_bar1, s_bar2, s_bar3 = st.columns([2.5, 1, 1])
    with s_bar1:
        chosen_sheet = st.selectbox("Select Case Sheet", case_sheet_names, index=case_sheet_names.index(active_sheet), key="adm_sheet_picker")
        if chosen_sheet != active_sheet:
            st.session_state["admin_active_sheet"] = chosen_sheet
            st.rerun()
    with s_bar2:
        if st.button("🔄 Sync Sheets", key="adm_sync_btn", use_container_width=True):
            _clear_sheet_runtime_caches()
            st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
            st.rerun()
    with s_bar3:
        if st.button("🚪 Logout", key="adm_logout_btn", use_container_width=True):
            st.session_state["logged_in"] = False
            st.session_state["role"] = ""
            st.session_state["username"] = ""
            st.rerun()

    df_current = all_sheets.get(active_sheet, pd.DataFrame())
    render_kpi_row(df_current, all_sheets)

    admin_tab = st.radio(
        "Admin Workspace View",
        ["📋 Cases & Search", "➕ Add New Case", "📈 Leads & Booked Report"],
        horizontal=True,
        key="adm_workspace_radio",
    )

    if admin_tab == "📋 Cases & Search":
        flt = render_filter_bar(df_current, key_prefix="adm_flt", allow_worker_select=True, default_lead="All")
        df_filtered = filter_cases_dataframe(
            df_current,
            search_text=flt["search"],
            lead_filter=flt["lead"],
            status_filter=flt["status"],
            date_from=flt["date_from"],
            date_to=flt["date_to"],
            worker_only=None,
        )
        render_cases_interactive_table(
            df_filtered,
            sheet_name=active_sheet,
            actor_role="admin",
            actor_name=admin_user,
            can_edit=True,
            can_delete=True,
            key_prefix="adm_tbl",
        )

    elif admin_tab == "➕ Add New Case":
        render_add_case_form(active_sheet, actor_role="admin", actor_name=admin_user, key_prefix="adm_add")

    elif admin_tab == "📈 Leads & Booked Report":
        st.markdown("### 📈 Leads & Booked Cases Summary Report")
        flt_r = render_filter_bar(df_current, key_prefix="adm_rep_flt", allow_worker_select=True, default_lead="All")
        df_rep = filter_cases_dataframe(
            df_current,
            search_text=flt_r["search"],
            lead_filter=flt_r["lead"],
            status_filter=flt_r["status"],
            date_from=flt_r["date_from"],
            date_to=flt_r["date_to"],
            worker_only=None,
        )

        lead_col = _find_col_ci(df_rep, ["Lead", "Worker", "Case Worker", "Assigned To"])
        status_col = _find_col_ci(df_rep, ["Status", "Booking_Status", "Booking Status", "Case Status"])

        total_r = len(df_rep)
        booked_r = 0
        resched_r = 0
        if not df_rep.empty and status_col:
            s_l = df_rep[status_col].astype(str).str.strip().str.lower()
            booked_r = int(s_l.str.contains("book", na=False).sum())
            resched_r = int(s_l.str.contains("resched", na=False).sum())

        breakdown_rows = []
        if not df_rep.empty and lead_col:
            for l_val, grp in df_rep.groupby(df_rep[lead_col].astype(str).str.strip()):
                l_clean = l_val if l_val else "Unassigned"
                g_cnt = len(grp)
                g_bk = 0
                if status_col:
                    g_bk = int(grp[status_col].astype(str).str.strip().str.lower().str.contains("book", na=False).sum())
                breakdown_rows.append((l_clean, str(g_cnt), f"Booked: {g_bk}"))

        kpi_items = [
            ("Filtered Cases", total_r, "#1d4ed8"),
            ("Booked Cases", booked_r, "#059669"),
            ("Rescheduled", resched_r, "#d97706"),
            ("Sheet", active_sheet, "#4f46e5"),
        ]
        render_image_report_card(
            title="LEADS & BOOKED CASES REPORT",
            subtitle=f"Sheet: {active_sheet} • Lead: {flt_r['lead']} • Generated: {get_ist_now().strftime('%Y-%m-%d %I:%M %p IST')}",
            kpi_items=kpi_items,
            breakdown_rows=breakdown_rows,
            filename_prefix=f"LEADS_BOOKED_CASES_{active_sheet}_{get_ist_now().strftime('%Y%m%d')}",
        )

    active_modal = st.session_state.get("admin_modal")
    if active_modal == "permissions":
        st.session_state["admin_modal"] = None
        open_admin_permissions_dialog(all_sheets)
    elif active_modal == "edit_requests":
        st.session_state["admin_modal"] = None
        open_edit_requests_dialog(all_sheets)
    elif active_modal == "attendance":
        st.session_state["admin_modal"] = None
        open_attendance_dialog(all_sheets)
    elif active_modal == "pdfs":
        st.session_state["admin_modal"] = None
        open_all_pdfs_dialog(all_sheets)
    elif active_modal == "settings":
        st.session_state["admin_modal"] = None
        open_admin_settings_dialog()
# ================= END OF STEP 9 =================
# ================= STEP 10: CASE WORKER DASHBOARD & MAIN ENTRYPOINT =================
@st.dialog("🔑 Change My Password", width="small")
def open_worker_password_dialog(worker_name):
    logins = load_worker_logins()
    curr_info = logins.get(
        worker_name,
        {
            "password": "",
            "last_login": "",
            "last_logout": "",
            "status": "Online",
            "last_seen": "",
        },
    )
    with st.form("worker_change_pw_form"):
        new_pw = st.text_input("New Password", value=str(curr_info.get("password", "")), type="password")
        conf_pw = st.text_input("Confirm Password", value="", type="password")
        sub_pw = st.form_submit_button("💾 Update Password", type="primary", use_container_width=True)
        if sub_pw:
            if not new_pw.strip():
                st.error("Password cannot be empty.")
            elif new_pw != conf_pw:
                st.error("New password and confirmation do not match.")
            else:
                curr_info["password"] = new_pw.strip()
                logins[worker_name] = curr_info
                save_worker_logins(logins)
                st.success("Password updated successfully.")
                st.rerun()


def render_worker_dashboard(all_sheets):
    worker_name = st.session_state.get("username", "")
    render_top_header("Case Worker Portal", worker_name)

    w_perms = get_worker_perms(worker_name)
    can_edit = bool(w_perms.get("can_edit", True))
    can_delete = bool(w_perms.get("can_delete", False))
    can_add_case = bool(w_perms.get("can_add_case", True))
    can_view_all = bool(w_perms.get("can_view_all", False))
    can_export = bool(w_perms.get("can_export", True))

    case_sheet_names = [s for s in all_sheets.keys() if s not in SYSTEM_SHEETS]
    if not case_sheet_names:
        case_sheet_names = ["Active Cases"]
        all_sheets["Active Cases"] = _normalize_sheet_df(pd.DataFrame())

    active_sheet = st.session_state.get("worker_active_sheet", case_sheet_names[0])
    if active_sheet not in case_sheet_names:
        active_sheet = case_sheet_names[0]
        st.session_state["worker_active_sheet"] = active_sheet

    w_c1, w_c2, w_c3, w_c4, w_c5 = st.columns(5)
    with w_c1:
        chosen_s = st.selectbox(
            "Case Sheet",
            case_sheet_names,
            index=case_sheet_names.index(active_sheet),
            key="wrk_sheet_picker",
        )
        if chosen_s != active_sheet:
            st.session_state["worker_active_sheet"] = chosen_s
            st.rerun()
    with w_c2:
        if st.button("🔄 Refresh Data", key="wrk_refresh_btn", use_container_width=True):
            _clear_sheet_runtime_caches()
            st.session_state["_gsheet_sync_version"] = st.session_state.get("_gsheet_sync_version", 0) + 1
            st.rerun()
    with w_c3:
        if st.button("🔑 Change Password", key="wrk_pw_btn", use_container_width=True):
            st.session_state["worker_modal"] = "password"
            st.rerun()
    with w_c4:
        if can_export:
            w_excel = get_worker_export_bytes(worker_name, _get_file_mtime())
            st.download_button(
                "📊 My Excel",
                data=w_excel,
                file_name=f"HemanVisa_{worker_name}_{get_ist_now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="wrk_export_btn",
                use_container_width=True,
            )
        else:
            st.button("📊 Export Locked", disabled=True, key="wrk_export_locked", use_container_width=True)
    with w_c5:
        if st.button("🚪 Logout", key="wrk_logout_btn", use_container_width=True):
            record_worker_logout_time(worker_name)
            st.session_state["logged_in"] = False
            st.session_state["role"] = ""
            st.session_state["username"] = ""
            st.rerun()

    df_current = all_sheets.get(active_sheet, pd.DataFrame())
    worker_only_filter = None if can_view_all else worker_name

    base_worker_df = filter_cases_dataframe(
        df_current,
        search_text="",
        lead_filter="All",
        status_filter="All",
        date_from=None,
        date_to=None,
        worker_only=worker_only_filter,
    )
    render_kpi_row(base_worker_df, all_sheets)

    w_tab_labels = ["📋 My Assigned Cases"]
    if can_add_case:
        w_tab_labels.append("➕ Add New Case")
    w_tab_labels.append("📈 My Summary Report")

    w_tab = st.radio("Worker Workspace View", w_tab_labels, horizontal=True, key="wrk_workspace_radio")

    if w_tab == "📋 My Assigned Cases":
        flt = render_filter_bar(
            df_current,
            key_prefix="wrk_flt",
            allow_worker_select=can_view_all,
            default_lead="All" if can_view_all else worker_name,
        )
        df_filtered = filter_cases_dataframe(
            df_current,
            search_text=flt["search"],
            lead_filter=flt["lead"] if can_view_all else "All",
            status_filter=flt["status"],
            date_from=flt["date_from"],
            date_to=flt["date_to"],
            worker_only=worker_only_filter,
        )
        render_cases_interactive_table(
            df_filtered,
            sheet_name=active_sheet,
            actor_role="worker",
            actor_name=worker_name,
            can_edit=can_edit,
            can_delete=can_delete,
            key_prefix="wrk_tbl",
        )

    elif w_tab == "➕ Add New Case" and can_add_case:
        render_add_case_form(active_sheet, actor_role="worker", actor_name=worker_name, key_prefix="wrk_add")

    elif w_tab == "📈 My Summary Report":
        status_col = _find_col_ci(base_worker_df, ["Status", "Booking_Status", "Booking Status", "Case Status"])
        total_w = len(base_worker_df)
        booked_w = 0
        resched_w = 0
        if not base_worker_df.empty and status_col:
            s_l = base_worker_df[status_col].astype(str).str.strip().str.lower()
            booked_w = int(s_l.str.contains("book", na=False).sum())
            resched_w = int(s_l.str.contains("resched", na=False).sum())

        breakdown_rows = []
        if not base_worker_df.empty and status_col:
            for st_val, grp in base_worker_df.groupby(base_worker_df[status_col].astype(str).str.strip()):
                breakdown_rows.append((st_val or "Unspecified", str(len(grp)), f"Sheet: {active_sheet}"))

        render_image_report_card(
            title=f"{worker_name.upper()} — CASE SUMMARY REPORT",
            subtitle=f"Sheet: {active_sheet} • Generated: {get_ist_now().strftime('%Y-%m-%d %I:%M %p IST')}",
            kpi_items=[
                ("My Total Cases", total_w, "#1d4ed8"),
                ("Booked Cases", booked_w, "#059669"),
                ("Rescheduled", resched_w, "#d97706"),
                ("Sheet", active_sheet, "#4f46e5"),
            ],
            breakdown_rows=breakdown_rows,
            filename_prefix=f"CASE_SUMMARY_{worker_name}_{active_sheet}_{get_ist_now().strftime('%Y%m%d')}",
        )

    if st.session_state.get("worker_modal") == "password":
        st.session_state["worker_modal"] = None
        open_worker_password_dialog(worker_name)


def main():
    inject_global_css()

    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    if "role" not in st.session_state:
        st.session_state["role"] = ""
    if "username" not in st.session_state:
        st.session_state["username"] = ""

    if not st.session_state.get("logged_in"):
        render_login_screen()
        return

    all_sheets = load_all_sheets_cached()
    role = str(st.session_state.get("role", "")).strip().lower()

    if role == "admin":
        render_admin_dashboard(all_sheets)
    else:
        render_worker_dashboard(all_sheets)

    dlg_case = st.session_state.get("active_dialog_case")
    if isinstance(dlg_case, dict) and dlg_case.get("sheet") is not None and dlg_case.get("row_idx") is not None:
        open_case_action_dialog(
            sheet_name=dlg_case["sheet"],
            row_idx=dlg_case["row_idx"],
            actor_role=dlg_case.get("role", role),
            actor_name=dlg_case.get("user", st.session_state.get("username", "")),
            can_edit=bool(dlg_case.get("can_edit", True)),
            can_delete=bool(dlg_case.get("can_delete", False)),
        )


if __name__ == "__main__":
    main()
# ================= END OF STEP 10 (COMPLETE FILE) =================
