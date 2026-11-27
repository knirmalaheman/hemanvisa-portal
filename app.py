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
