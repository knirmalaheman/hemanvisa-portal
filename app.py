import datetime
import io
import json
import os
import re
from zoneinfo import ZoneInfo

import pandas as pd
import requests
import streamlit as st

try:
    IST_TZ = ZoneInfo("Asia/Kolkata")
except Exception:
    IST_TZ = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

EXCEL_FILE = "Heman_Report.xlsx"
PDF_BASE_DIR = "uploaded_appointment_pdfs"
LOGO_FILES = ["pngTShirtPrintCricket.jpg", "logo.jpg", "logo.png", "logo.jpeg"]

SYSTEM_SHEETS = {
    "App_Logins",
    "App_Attendance",
    "App_Permissions",
    "App_Config",
    "App_Edit_Requests",
}


def get_ist_now() -> datetime.datetime:
    try:
        return datetime.datetime.now(datetime.timezone.utc).astimezone(IST_TZ).replace(tzinfo=None)
    except Exception:
        return datetime.datetime.utcnow() + datetime.timedelta(hours=5, minutes=30)


@st.cache_data(ttl=3600, show_spinner=False)
def get_webapp_url() -> str:
    try:
        url = st.secrets.get("GSHEET_WEBAPP_URL", "")
        return str(url).strip() if url else ""
    except Exception:
        return ""


def _rows_to_df(raw_rows: list) -> pd.DataFrame:
    if not raw_rows or not isinstance(raw_rows, list):
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


def fetch_sheets_from_webapp() -> dict | None:
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
                clean_name = str(s_name).strip()
                if clean_name:
                    sheets_dict[clean_name] = _rows_to_df(s_rows)
            return sheets_dict or None

        if isinstance(payload, list):
            return {"Active Cases": _rows_to_df(payload)}
        return None
    except Exception as exc:
        st.session_state["_last_webapp_fetch_error"] = str(exc)
        return None


def reorder_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    cols_to_drop = [
        c for c in df.columns if c.strip().lower() in ["ds-160 case worker", "ds160 case worker"]
    ]
    if cols_to_drop:
        df = df.drop(columns=cols_to_drop, errors="ignore")

    cols = list(df.columns)
    priority = []
    for key in ["Sheet_Name", "Case_ID"]:
        if key in cols:
            priority.append(key)

    for prefix in ["case", "username", "password", "no of app", "lead"]:
        for c in cols:
            c_low = c.strip().lower().replace("_", " ")
            if (c_low == prefix or prefix in c_low) and c not in priority:
                priority.append(c)

    remaining = [c for c in cols if c not in priority]
    return df[priority + remaining]


def _sanitize_sheet_title(title: str, used_titles: set) -> str:
    # Excel sheet names: max 31 chars, cannot contain \ / ? * : [ ]
    cleaned = re.sub(r"[\\/?*:\[\]]", "_", str(title).strip()) or "Sheet"
    base = cleaned[:31]
    candidate = base
    counter = 1
    while candidate.lower() in used_titles:
        suffix = f"_{counter}"
        candidate = f"{base[: 31 - len(suffix)]}{suffix}"
        counter += 1
    used_titles.add(candidate.lower())
    return candidate


@st.cache_data(ttl=120, show_spinner=False)
def get_gsheet_as_excel_bytes(cache_bust: int = 0) -> bytes:
    """
    Google Sheet lo unna data ni (SYSTEM_SHEETS kakunda) exact .xlsx file bytes ga marustundi.
    """
    sheets_map = fetch_sheets_from_webapp()
    if not sheets_map and os.path.exists(EXCEL_FILE):
        try:
            sheets_map = pd.read_excel(EXCEL_FILE, sheet_name=None, dtype=str)
        except Exception:
            sheets_map = {}

    buffer = io.BytesIO()
    used_titles = set()
    wrote_any_sheet = False

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if sheets_map:
            for s_name, df_s in sheets_map.items():
                if str(s_name).strip() in SYSTEM_SHEETS:
                    continue
                df_clean = reorder_dataframe_columns(df_s.copy()) if df_s is not None else pd.DataFrame()
                if df_clean is None:
                    df_clean = pd.DataFrame()
                df_clean = df_clean.fillna("").astype(str)
                sheet_tab = _sanitize_sheet_title(s_name, used_titles)
                df_clean.to_excel(writer, sheet_name=sheet_tab, index=False)
                wrote_any_sheet = True

        if not wrote_any_sheet:
            pd.DataFrame().to_excel(writer, sheet_name="Sheet1", index=False)

    return buffer.getvalue()


def render_gsheet_excel_download_button(
    label: str = "📥 Download Google Sheet (.xlsx)",
    file_name: str = EXCEL_FILE,
    key: str = "download_gsheet_excel_btn",
    use_container_width: bool = True,
) -> None:
    """
    Website lo direct ga Google Sheet Excel file download chese button render chestundi.
    """
    sync_ver = int(st.session_state.get("_gsheet_sync_version", 0))
    excel_bytes = get_gsheet_as_excel_bytes(sync_ver)
    st.download_button(
        label=label,
        data=excel_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
        use_container_width=use_container_width,
    )
