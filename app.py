from datetime import datetime
from pathlib import Path
import re
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Case & Appointment Portal v6",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
PDF_ROOT_DIR = BASE_DIR / "appointments_pdf"
EXCEL_FILE = DATA_DIR / "cases_v6.xlsx"

DATA_DIR.mkdir(parents=True, exist_ok=True)
PDF_ROOT_DIR.mkdir(parents=True, exist_ok=True)

USERS = {
    "admin": {"role": "Admin", "name": "Admin"},
    "kiran": {"role": "Case Worker", "name": "Kiran"},
    "shanthi": {"role": "Case Worker", "name": "Shanthi"},
    "keerthana": {"role": "Case Worker", "name": "Keerthana"},
    "sai": {"role": "Case Worker", "name": "Sai"},
}

LEADS = ["All", "Kiran", "Shanthi"]
WORKERS = ["Keerthana", "Sai", "Kiran", "Shanthi"]
STATUSES = ["Active", "Booked", "Rescheduled", "Closed"]

COLUMNS = [
    "HV_ID",
    "Client_Name",
    "Phone",
    "Lead",
    "Assigned_To",
    "Status",
    "Appointment_Date",
    "Appointment_Time_12h",
    "PDF_Path",
    "Notes",
    "Updated_By",
    "Updated_At",
]


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")


def format_12h(val: str) -> str:
    txt = (val or "").strip()
    if not txt:
        return ""
    for fmt in ("%I:%M %p", "%I:%M%p", "%H:%M", "%H:%M:%S"):
        try:
            dt = datetime.strptime(txt.upper(), fmt)
            return dt.strftime("%I:%M %p")
        except ValueError:
            continue
    return txt


def generate_next_hv_id(df: pd.DataFrame) -> str:
    max_num = 0
    if not df.empty and "HV_ID" in df.columns:
        for val in df["HV_ID"].dropna().astype(str):
            m = re.search(r"HV[-_]?(\d+)", val.strip(), re.IGNORECASE)
            if m:
                max_num = max(max_num, int(m.group(1)))
    return f"HV-{max_num + 1:04d}"


def load_cases_df() -> pd.DataFrame:
    if EXCEL_FILE.exists():
        df = pd.read_excel(EXCEL_FILE, dtype=str).fillna("")
        for col in COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df[COLUMNS]

    today_str = datetime.now().strftime("%Y-%m-%d")
    sample_df = pd.DataFrame(
        [
            {
                "HV_ID": "HV-0001",
                "Client_Name": "Ramesh Kumar",
                "Phone": "9848000001",
                "Lead": "Kiran",
                "Assigned_To": "Keerthana",
                "Status": "Active",
                "Appointment_Date": today_str,
                "Appointment_Time_12h": "10:30 AM",
                "PDF_Path": "",
                "Notes": "Initial intake case",
                "Updated_By": "Admin",
                "Updated_At": now_str(),
            },
            {
                "HV_ID": "HV-0002",
                "Client_Name": "Sita Devi",
                "Phone": "9848000002",
                "Lead": "Shanthi",
                "Assigned_To": "Sai",
                "Status": "Booked",
                "Appointment_Date": today_str,
                "Appointment_Time_12h": "02:15 PM",
                "PDF_Path": "",
                "Notes": "Confirmed slot",
                "Updated_By": "Admin",
                "Updated_At": now_str(),
            },
            {
                "HV_ID": "HV-0003",
                "Client_Name": "Anil Reddy",
                "Phone": "9848000003",
                "Lead": "Kiran",
                "Assigned_To": "Kiran",
                "Status": "Rescheduled",
                "Appointment_Date": today_str,
                "Appointment_Time_12h": "04:00 PM",
                "PDF_Path": "",
                "Notes": "Moved to evening slot",
                "Updated_By": "Kiran",
                "Updated_At": now_str(),
            },
        ],
        columns=COLUMNS,
    )
    save_cases_df(sample_df)
    return sample_df


def save_cases_df(df: pd.DataFrame) -> None:
    clean_df = df[COLUMNS].copy()
    clean_df.to_excel(EXCEL_FILE, index=False)


def init_session_state():
    if "logged_in_user" not in st.session_state:
        st.session_state.logged_in_user = None
    if "cases_df" not in st.session_state:
        st.session_state.cases_df = load_cases_df()
    if "attendance_log" not in st.session_state:
        st.session_state.attendance_log = []
    if "change_log" not in st.session_state:
        st.session_state.change_log = []
    if "delete_queue" not in st.session_state:
        st.session_state.delete_queue = []
    if "edit_hv_id" not in st.session_state:
        st.session_state.edit_hv_id = None


def log_change(user_name: str, hv_id: str, action: str, details: str):
    st.session_state.change_log.insert(
        0,
        {
            "Time": now_str(),
            "User": user_name,
            "HV_ID": hv_id,
            "Action": action,
            "Details": details,
        },
    )


def save_pdf_to_date_folder(uploaded_file, appt_date_str: str, hv_id: str) -> str:
    safe_date = re.sub(r"[^0-9-]", "", appt_date_str or "") or datetime.now().strftime(
        "%Y-%m-%d"
    )
    date_folder = PDF_ROOT_DIR / safe_date
    date_folder.mkdir(parents=True, exist_ok=True)
    clean_name = re.sub(r"[^A-Za-z0-9._-]", "_", uploaded_file.name)
    target_path = date_folder / f"{hv_id}_{clean_name}"
    with open(target_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return str(target_path)


def render_kpi_cards(full_df: pd.DataFrame):
    active_cnt = int((full_df["Status"] == "Active").sum())
    booked_cnt = int((full_df["Status"] == "Booked").sum())
    resched_cnt = int((full_df["Status"] == "Rescheduled").sum())
    total_cnt = int(len(full_df))

    cards = [
        ("Active Cases", active_cnt, "#16a34a", "#f0fdf4"),
        ("Booked Cases", booked_cnt, "#2563eb", "#eff6ff"),
        ("Rescheduled", resched_cnt, "#d97706", "#fffbeb"),
        ("Total Cases", total_cnt, "#475569", "#f8fafc"),
    ]
    cols = st.columns(4)
    for col, (label, count, border_color, bg_color) in zip(cols, cards):
        col.markdown(
            f"""
            <div style="
                background:{bg_color};
                border-left:5px solid {border_color};
                border-radius:8px;
                padding:12px 16px;
                margin-bottom:8px;
            ">
                <div style="font-size:13px;color:#334155;font-weight:600;">{label}</div>
                <div style="font-size:26px;color:#0f172a;font-weight:700;">{count}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def enable_tab_navigation():
    components.html(
        """
        <script>
        const doc = window.parent.document;
        doc.addEventListener("keydown", function(e) {
            if (e.key === "Tab") {
                const focusable = Array.from(doc.querySelectorAll(
                    'input, select, textarea, button, [tabindex]:not([tabindex="-1"])'
                )).filter(el => el.offsetParent !== null && !el.disabled);
                if (focusable.length > 0) {
                    const idx = focusable.indexOf(doc.activeElement);
                    if (idx !== -1) {
                        e.preventDefault();
                        const nextIdx = e.shiftKey
                            ? (idx - 1 + focusable.length) % focusable.length
                            : (idx + 1) % focusable.length;
                        focusable[nextIdx].focus();
                    }
                }
            }
        }, true);
        </script>
        """,
        height=0,
    )


@st.dialog("Edit Case Details")
def edit_case_dialog(hv_id: str, actor_name: str):
    df = st.session_state.cases_df.copy()
    matches = df.index[df["HV_ID"] == hv_id].tolist()
    if not matches:
        st.warning("Case not found.")
        return

    idx = matches[0]
    row = df.loc[idx]

    st.markdown(f"**HV ID (Non-Editable):** `{row['HV_ID']}`")
    with st.form(f"edit_form_{hv_id}"):
        c1, c2 = st.columns(2)
        client_name = c1.text_input("Client Name", value=row["Client_Name"])
        phone = c2.text_input("Phone", value=row["Phone"])

        c3, c4, c5 = st.columns(3)
        lead_options = ["Kiran", "Shanthi"]
        lead_val = row["Lead"] if row["Lead"] in lead_options else "Kiran"
        lead = c3.selectbox("Lead", lead_options, index=lead_options.index(lead_val))

        assigned_val = row["Assigned_To"] if row["Assigned_To"] in WORKERS else "Keerthana"
        assigned_to = c4.selectbox("Assigned To", WORKERS, index=WORKERS.index(assigned_val))

        status_val = row["Status"] if row["Status"] in STATUSES else "Active"
        status = c5.selectbox("Status", STATUSES, index=STATUSES.index(status_val))

        c6, c7 = st.columns(2)
        appt_date = c6.text_input("Appointment Date (YYYY-MM-DD)", value=row["Appointment_Date"])
        appt_time_12h = c7.text_input(
            "Appointment Time (12-hr, e.g., 11:30 AM)", value=row["Appointment_Time_12h"]
        )

        notes = st.text_area("Notes", value=row["Notes"], height=80)
        pdf_upload = st.file_uploader("Upload / Replace Appointment PDF", type=["pdf"])

        b1, b2 = st.columns(2)
        submitted = b1.form_submit_button("Save Changes", type="primary", width="stretch")
        cancel = b2.form_submit_button("Cancel", width="stretch")

        if cancel:
            st.session_state.edit_hv_id = None
            st.rerun()

        if submitted:
            pdf_path = row["PDF_Path"]
            if pdf_upload is not None:
                pdf_path = save_pdf_to_date_folder(pdf_upload, appt_date, hv_id)

            formatted_time = format_12h(appt_time_12h)
            df.loc[idx, "Client_Name"] = client_name.strip()
            df.loc[idx, "Phone"] = phone.strip()
            df.loc[idx, "Lead"] = lead
            df.loc[idx, "Assigned_To"] = assigned_to
            df.loc[idx, "Status"] = status
            df.loc[idx, "Appointment_Date"] = appt_date.strip()
            df.loc[idx, "Appointment_Time_12h"] = formatted_time
            df.loc[idx, "PDF_Path"] = pdf_path
            df.loc[idx, "Notes"] = notes.strip()
            df.loc[idx, "Updated_By"] = actor_name
            df.loc[idx, "Updated_At"] = now_str()

            st.session_state.cases_df = df
            save_cases_df(df)
            log_change(
                actor_name,
                hv_id,
                "Updated Case",
                f"Status={status}, Lead={lead}, Worker={assigned_to}, Date={appt_date}, Time={formatted_time}",
            )
            st.session_state.edit_hv_id = None
            st.success(f"{hv_id} updated and synced to local Excel.")
            st.rerun()


def render_login_screen():
    st.title("Case & Appointment Management Portal (v6)")
    st.caption("Local Excel Sync • Role-Based Access • 12-Hour Schedule Format")

    with st.form("login_form"):
        username = st.selectbox(
            "Select User",
            options=list(USERS.keys()),
            format_func=lambda u: f"{USERS[u]['name']} ({USERS[u]['role']})",
        )
        submitted = st.form_submit_button("Login", type="primary")
        if submitted:
            user_info = USERS[username]
            st.session_state.logged_in_user = {
                "username": username,
                "name": user_info["name"],
                "role": user_info["role"],
                "login_time": now_str(),
            }
            st.session_state.attendance_log.insert(
                0,
                {
                    "User": user_info["name"],
                    "Role": user_info["role"],
                    "Login_Time": now_str(),
                },
            )
            st.rerun()


def main():
    init_session_state()
    enable_tab_navigation()

    if st.session_state.logged_in_user is None:
        render_login_screen()
        return

    current_user = st.session_state.logged_in_user
    actor_name = current_user["name"]
    is_admin = current_user["role"] == "Admin"

    with st.sidebar:
        st.markdown(f"### {actor_name}")
        st.caption(f"Role: {current_user['role']}")
        st.caption(f"Excel File: `{EXCEL_FILE.name}`")
        if st.button("Reload From Local Excel", width="stretch"):
            st.session_state.cases_df = load_cases_df()
            st.rerun()
        if st.button("Logout", width="stretch"):
            st.session_state.logged_in_user = None
            st.rerun()

    st.markdown("### Case & Appointment Dashboard")

    f1, f2, f3 = st.columns([2, 2, 1.5])
    lead_filter = f1.radio("Lead Filter", LEADS, horizontal=True)
    status_filter = f2.multiselect(
        "Status Filter", STATUSES, default=["Active", "Booked", "Rescheduled"]
    )
    search_q = f3.text_input("Search HV / Client / Phone", placeholder="HV-0001 or name")

    full_df = st.session_state.cases_df.copy()
    render_kpi_cards(full_df)

    df = full_df.copy()
    if lead_filter != "All":
        df = df[df["Lead"] == lead_filter]
    if status_filter:
        df = df[df["Status"].isin(status_filter)]
    if not is_admin:
        df = df[(df["Assigned_To"] == actor_name) | (df["Lead"] == actor_name)]
    if search_q.strip():
        q = search_q.strip().lower()
        mask = (
            df["HV_ID"].str.lower().str.contains(q, na=False)
            | df["Client_Name"].str.lower().str.contains(q, na=False)
            | df["Phone"].str.lower().str.contains(q, na=False)
        )
        df = df[mask]

    st.markdown("#### Cases Table")
    st.dataframe(df, width="stretch", hide_index=True)

    st.markdown("#### Quick Case Actions")
    hv_options = df["HV_ID"].tolist() if not df.empty else full_df["HV_ID"].tolist()
    if hv_options:
        a1, a2, a3, a4 = st.columns([2, 1.2, 1.3, 1.5])
        selected_hv = a1.selectbox("Select HV Case", hv_options, key="selected_hv_action")

        if a2.button("Edit Case", width="stretch"):
            st.session_state.edit_hv_id = selected_hv
            st.rerun()

        if a3.button("Mark Rescheduled", width="stretch"):
            idx_list = full_df.index[full_df["HV_ID"] == selected_hv].tolist()
            if idx_list:
                idx = idx_list[0]
                full_df.loc[idx, "Status"] = "Rescheduled"
                full_df.loc[idx, "Updated_By"] = actor_name
                full_df.loc[idx, "Updated_At"] = now_str()
                st.session_state.cases_df = full_df
                save_cases_df(full_df)
                log_change(actor_name, selected_hv, "Rescheduled", "Status marked Rescheduled")
                st.success(f"{selected_hv} marked as Rescheduled.")
                st.rerun()

        selected_row = full_df[full_df["HV_ID"] == selected_hv].iloc[0]
        pdf_file_path = Path(selected_row["PDF_Path"]) if selected_row["PDF_Path"] else None
        if pdf_file_path and pdf_file_path.exists():
            with open(pdf_file_path, "rb") as pf:
                a4.download_button(
                    label="Download PDF",
                    data=pf.read(),
                    file_name=pdf_file_path.name,
                    mime="application/pdf",
                    width="stretch",
                )
        else:
            a4.button("No PDF Attached", disabled=True, width="stretch")

        if not is_admin:
            if st.button("Request Deletion (Admin Approval Queue)"):
                already_queued = any(
                    item["HV_ID"] == selected_hv and item["Status"] == "Pending"
                    for item in st.session_state.delete_queue
                )
                if not already_queued:
                    st.session_state.delete_queue.append(
                        {
                            "HV_ID": selected_hv,
                            "Client_Name": selected_row["Client_Name"],
                            "Requested_By": actor_name,
                            "Requested_At": now_str(),
                            "Status": "Pending",
                        }
                    )
                    log_change(actor_name, selected_hv, "Delete Requested", "Sent to approval queue")
                    st.info(f"Deletion request for {selected_hv} added to approval queue.")
                    st.rerun()

    if st.session_state.edit_hv_id:
        edit_case_dialog(st.session_state.edit_hv_id, actor_name)

    st.divider()

    next_hv = generate_next_hv_id(full_df)
    st.markdown(f"#### Add New Case (Auto HV ID: `{next_hv}`)")
    with st.form("add_case_form", clear_on_submit=True):
        c0, c1, c2, c3 = st.columns([1.2, 2, 1.5, 1.3])
        c0.text_input("HV ID (Auto)", value=next_hv, disabled=True)
        client_name = c1.text_input("Client Name")
        phone = c2.text_input("Phone")
        lead = c3.selectbox("Lead", ["Kiran", "Shanthi"])

        c4, c5, c6, c7 = st.columns(4)
        assigned_to = c4.selectbox("Assigned To", WORKERS)
        status = c5.selectbox("Status", STATUSES, index=0)
        appt_date = c6.text_input(
            "Appointment Date (YYYY-MM-DD)", value=datetime.now().strftime("%Y-%m-%d")
        )
        appt_time_12h = c7.text_input("Appointment Time (12-hr)", value="11:00 AM")

        notes = st.text_input("Notes")
        pdf_file = st.file_uploader("Upload Appointment PDF (Optional)", type=["pdf"])

        create_btn = st.form_submit_button("Add Case & Sync Excel", type="primary")
        if create_btn:
            if not client_name.strip():
                st.error("Client Name is required.")
            else:
                fresh_hv = generate_next_hv_id(st.session_state.cases_df)
                pdf_saved_path = ""
                if pdf_file is not None:
                    pdf_saved_path = save_pdf_to_date_folder(pdf_file, appt_date, fresh_hv)

                formatted_time = format_12h(appt_time_12h)
                new_row = {
                    "HV_ID": fresh_hv,
                    "Client_Name": client_name.strip(),
                    "Phone": phone.strip(),
                    "Lead": lead,
                    "Assigned_To": assigned_to,
                    "Status": status,
                    "Appointment_Date": appt_date.strip(),
                    "Appointment_Time_12h": formatted_time,
                    "PDF_Path": pdf_saved_path,
                    "Notes": notes.strip(),
                    "Updated_By": actor_name,
                    "Updated_At": now_str(),
                }
                updated_df = pd.concat(
                    [st.session_state.cases_df, pd.DataFrame([new_row], columns=COLUMNS)],
                    ignore_index=True,
                )
                st.session_state.cases_df = updated_df
                save_cases_df(updated_df)
                log_change(
                    actor_name,
                    fresh_hv,
                    "Created Case",
                    f"Client={client_name.strip()}, Lead={lead}, Time={formatted_time}",
                )
                st.success(f"Case {fresh_hv} added and saved to {EXCEL_FILE.name}.")
                st.rerun()

    st.divider()

    tab1, tab2, tab3 = st.tabs(["Deletion Approval Queue", "Change Tracking Log", "Login Attendance"])

    with tab1:
        if not st.session_state.delete_queue:
            st.info("No deletion requests in queue.")
        else:
            st.dataframe(
                pd.DataFrame(st.session_state.delete_queue),
                width="stretch",
                hide_index=True,
            )
            if is_admin:
                pending_hvs = [
                    item["HV_ID"]
                    for item in st.session_state.delete_queue
                    if item["Status"] == "Pending"
                ]
                if pending_hvs:
                    d1, d2, d3 = st.columns([2, 1, 1])
                    target_del_hv = d1.selectbox("Pending Request HV ID", pending_hvs)
                    if d2.button("Approve & Delete", type="primary", width="stretch"):
                        curr_df = st.session_state.cases_df.copy()
                        curr_df = curr_df[curr_df["HV_ID"] != target_del_hv].reset_index(drop=True)
                        st.session_state.cases_df = curr_df
                        save_cases_df(curr_df)
                        for item in st.session_state.delete_queue:
                            if item["HV_ID"] == target_del_hv and item["Status"] == "Pending":
                                item["Status"] = f"Approved by {actor_name}"
                        log_change(actor_name, target_del_hv, "Deleted Case", "Approved deletion")
                        st.success(f"Deleted {target_del_hv} from Excel.")
                        st.rerun()
                    if d3.button("Reject Request", width="stretch"):
                        for item in st.session_state.delete_queue:
                            if item["HV_ID"] == target_del_hv and item["Status"] == "Pending":
                                item["Status"] = f"Rejected by {actor_name}"
                        log_change(actor_name, target_del_hv, "Delete Rejected", "Kept case active")
                        st.rerun()

    with tab2:
        if st.session_state.change_log:
            st.dataframe(
                pd.DataFrame(st.session_state.change_log),
                width="stretch",
                hide_index=True,
            )
        else:
            st.info("No changes recorded yet in this session.")

    with tab3:
        if st.session_state.attendance_log:
            st.dataframe(
                pd.DataFrame(st.session_state.attendance_log),
                width="stretch",
                hide_index=True,
            )
        else:
            st.info("No login attendance records yet.")


if __name__ == "__main__":
    main()
