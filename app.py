import os
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Case Workflow Manager (Kalyan & Anil)", layout="wide")

EXCEL_FILE = "cases_data.xlsx"
COLUMNS = [
    "HV_ID",
    "Client_Name",
    "Case_Worker",
    "Bucket",
    "Status",
    "Scheduled_Date",
    "Scheduled_Time",
    "Previous_ID",
    "Notes",
    "Updated_At",
]


def now_12hr() -> str:
    return datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")


def generate_new_hv_id(df_all: pd.DataFrame, prefix: str = "HV") -> str:
    """Generates the next unique HV ID like HV-1001, HV-1002..."""
    max_num = 1000
    if not df_all.empty and "HV_ID" in df_all.columns:
        for cid in df_all["HV_ID"].dropna().astype(str):
            cleaned = cid.strip().upper()
            if "-" in cleaned:
                parts = cleaned.split("-", 1)
                if parts[1].isdigit():
                    max_num = max(max_num, int(parts[1]))
            else:
                digits = "".join(ch for ch in cleaned if ch.isdigit())
                if digits.isdigit():
                    max_num = max(max_num, int(digits))
    return f"{prefix}-{max_num + 1}"


def init_workbook(path: str):
    """Creates initial Excel workbook with sheets if it does not exist."""
    if os.path.exists(path):
        return

    sample_rows = [
        {
            "HV_ID": "HV-1001",
            "Client_Name": "Ramesh Rao",
            "Case_Worker": "Kalyan",
            "Bucket": "Booked",
            "Status": "Booked",
            "Scheduled_Date": "2026-11-20",
            "Scheduled_Time": "10:30 AM",
            "Previous_ID": "",
            "Notes": "Initial booked case under Kalyan",
            "Updated_At": now_12hr(),
        },
        {
            "HV_ID": "HV-1002",
            "Client_Name": "Sita Devi",
            "Case_Worker": "Kalyan",
            "Bucket": "Active",
            "Status": "Active",
            "Scheduled_Date": "2026-11-21",
            "Scheduled_Time": "02:00 PM",
            "Previous_ID": "",
            "Notes": "Active case under Kalyan",
            "Updated_At": now_12hr(),
        },
        {
            "HV_ID": "HV-1003",
            "Client_Name": "Karthik Reddy",
            "Case_Worker": "Anil",
            "Bucket": "Active",
            "Status": "Active",
            "Scheduled_Date": "2026-11-22",
            "Scheduled_Time": "11:15 AM",
            "Previous_ID": "",
            "Notes": "Existing active case under Anil",
            "Updated_At": now_12hr(),
        },
    ]
    df_init = pd.DataFrame(sample_rows, columns=COLUMNS)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df_init.to_excel(writer, sheet_name="Master_Cases", index=False)
        df_init[df_init["Case_Worker"] == "Kalyan"].to_excel(
            writer, sheet_name="Kalyan_Cases", index=False
        )
        df_init[(df_init["Case_Worker"] == "Anil") & (df_init["Bucket"] == "Active")].to_excel(
            writer, sheet_name="Anil_Active", index=False
        )
        pd.DataFrame(columns=COLUMNS + ["Action", "Details"]).to_excel(
            writer, sheet_name="Audit_Log", index=False
        )


def load_data(path: str) -> dict[str, pd.DataFrame]:
    init_workbook(path)
    xls = pd.ExcelFile(path, engine="openpyxl")
    data = {}

    for sheet in ["Master_Cases", "Kalyan_Cases", "Anil_Active"]:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str).fillna("")
            for col in COLUMNS:
                if col not in df.columns:
                    df[col] = ""
            data[sheet] = df[COLUMNS]
        else:
            data[sheet] = pd.DataFrame(columns=COLUMNS)

    log_cols = COLUMNS + ["Action", "Details"]
    if "Audit_Log" in xls.sheet_names:
        ldf = pd.read_excel(xls, sheet_name="Audit_Log", dtype=str).fillna("")
        for col in log_cols:
            if col not in ldf.columns:
                ldf[col] = ""
        data["Audit_Log"] = ldf[log_cols]
    else:
        data["Audit_Log"] = pd.DataFrame(columns=log_cols)

    return data


def save_data(path: str, data: dict[str, pd.DataFrame]):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name in ["Master_Cases", "Kalyan_Cases", "Anil_Active", "Audit_Log"]:
            df = data.get(sheet_name, pd.DataFrame(columns=COLUMNS))
            df.to_excel(writer, sheet_name=sheet_name, index=False)


def get_all_ids(data: dict[str, pd.DataFrame]) -> pd.DataFrame:
    frames = []
    for name, df in data.items():
        if name != "Audit_Log" and not df.empty and "HV_ID" in df.columns:
            frames.append(df[["HV_ID"]])
    if not frames:
        return pd.DataFrame(columns=["HV_ID"])
    return pd.concat(frames, ignore_index=True)


def sync_sheets(data: dict[str, pd.DataFrame]):
    """Keeps Kalyan_Cases and Anil_Active synchronized with Master_Cases."""
    master = data["Master_Cases"].copy()
    data["Kalyan_Cases"] = (
        master[master["Case_Worker"] == "Kalyan"].copy().reset_index(drop=True)
    )
    data["Anil_Active"] = (
        master[(master["Case_Worker"] == "Anil") & (master["Bucket"] == "Active")]
        .copy()
        .reset_index(drop=True)
    )


def append_audit_log(
    data: dict[str, pd.DataFrame],
    row_dict: dict,
    action: str,
    details: str,
):
    entry = {col: str(row_dict.get(col, "")) for col in COLUMNS}
    entry["Action"] = action
    entry["Details"] = details
    data["Audit_Log"] = pd.concat(
        [data["Audit_Log"], pd.DataFrame([entry])], ignore_index=True
    )


# --- App Initialization ---
data = load_data(EXCEL_FILE)
sync_sheets(data)
master_df = data["Master_Cases"].copy()

st.title("Case Workflow Automation — Kalyan & Anil")
st.caption(
    "Rules active: (1) Reschedule Kalyan Booked -> Keep Old HV ID | "
    "(2) Reschedule TO Kalyan -> Generate New HV ID | "
    "(3) Kalyan Assigns to Anil -> Anil Active Sheet + Generate New HV ID"
)

# --- KPI Row ---
k1, k2, k3, k4 = st.columns(4)
kalyan_booked_df = master_df[
    (master_df["Case_Worker"] == "Kalyan") & (master_df["Bucket"] == "Booked")
]
kalyan_all_df = master_df[master_df["Case_Worker"] == "Kalyan"]
anil_active_df = master_df[
    (master_df["Case_Worker"] == "Anil") & (master_df["Bucket"] == "Active")
]

k1.metric("Kalyan Booked Cases", len(kalyan_booked_df))
k2.metric("Kalyan Total Cases", len(kalyan_all_df))
k3.metric("Anil Active Sheet Cases", len(anil_active_df))
k4.metric("Master Cases", len(master_df))

st.divider()

# --- SECTION 1: Reschedule within Kalyan Booked -> KEEP OLD ID ---
st.subheader("1. Reschedule Kalyan's Booked Case (Keep Old HV ID)")
if kalyan_booked_df.empty:
    st.info("No Booked cases currently under Kalyan.")
else:
    s1_id = st.selectbox(
        "Select Kalyan Booked Case",
        options=kalyan_booked_df["HV_ID"].tolist(),
        key="s1_id",
    )
    row1 = kalyan_booked_df[kalyan_booked_df["HV_ID"] == s1_id].iloc[0]

    c1, c2, c3 = st.columns(3)
    d1 = c1.text_input("New Date", value=row1["Scheduled_Date"], key="d1")
    t1 = c2.text_input("New Time (12-hr)", value=row1["Scheduled_Time"], key="t1")
    n1 = c3.text_input("Notes", value=row1["Notes"], key="n1")

    if st.button("Save Reschedule in Kalyan (Keep Old ID)", key="btn_s1"):
        idx = master_df.index[master_df["HV_ID"] == s1_id][0]
        ts = now_12hr()
        # Keep HV_ID unchanged (old ID preserved)
        master_df.at[idx, "Case_Worker"] = "Kalyan"
        master_df.at[idx, "Bucket"] = "Booked"
        master_df.at[idx, "Status"] = "Rescheduled (Booked)"
        master_df.at[idx, "Scheduled_Date"] = d1.strip()
        master_df.at[idx, "Scheduled_Time"] = t1.strip()
        master_df.at[idx, "Notes"] = n1.strip()
        master_df.at[idx, "Updated_At"] = ts

        data["Master_Cases"] = master_df
        sync_sheets(data)
        append_audit_log(
            data,
            master_df.loc[idx].to_dict(),
            "RESCHEDULE_KALYAN_BOOKED_KEEP_ID",
            f"Rescheduled within Kalyan Booked keeping Old ID {s1_id}",
        )
        save_data(EXCEL_FILE, data)
        st.success(f"Updated {s1_id} under Kalyan Booked with Old HV ID preserved.")
        st.rerun()

st.divider()

# --- SECTION 2: Reschedule TO Kalyan -> GENERATE NEW ID ---
st.subheader("2. Reschedule Case TO Kalyan (Generate New HV ID)")
eligible_to_kalyan = master_df[
    ~(
        (master_df["Case_Worker"] == "Kalyan")
        & (master_df["Bucket"] == "Booked")
        & (~master_df["Status"].str.startswith("Rescheduled to Kalyan"))
    )
]

if eligible_to_kalyan.empty:
    st.info("No eligible cases available to reschedule to Kalyan.")
else:
    s2_id = st.selectbox(
        "Select Source Case to Reschedule TO Kalyan",
        options=eligible_to_kalyan["HV_ID"].tolist(),
        key="s2_id",
    )
    row2 = eligible_to_kalyan[eligible_to_kalyan["HV_ID"] == s2_id].iloc[0]

    c1, c2, c3 = st.columns(3)
    b2 = c1.selectbox("Kalyan Target Bucket", ["Booked", "Active"], key="b2")
    d2 = c2.text_input("New Date", value=row2["Scheduled_Date"], key="d2")
    t2 = c3.text_input("New Time (12-hr)", value=row2["Scheduled_Time"], key="t2")

    if st.button("Reschedule TO Kalyan (Create New ID)", key="btn_s2"):
        new_hv_id = generate_new_hv_id(get_all_ids(data), prefix="HV")
        ts = now_12hr()

        src_idx = master_df.index[master_df["HV_ID"] == s2_id][0]
        master_df.at[src_idx, "Status"] = f"Rescheduled to Kalyan ({new_hv_id})"
        master_df.at[src_idx, "Updated_At"] = ts

        new_kalyan_row = {
            "HV_ID": new_hv_id,
            "Client_Name": row2["Client_Name"],
            "Case_Worker": "Kalyan",
            "Bucket": b2,
            "Status": "Rescheduled",
            "Scheduled_Date": d2.strip(),
            "Scheduled_Time": t2.strip(),
            "Previous_ID": s2_id,
            "Notes": f"Rescheduled to Kalyan from {s2_id}. {row2['Notes']}".strip(),
            "Updated_At": ts,
        }

        master_df = pd.concat(
            [master_df, pd.DataFrame([new_kalyan_row], columns=COLUMNS)],
            ignore_index=True,
        )
        data["Master_Cases"] = master_df
        sync_sheets(data)
        append_audit_log(
            data,
            new_kalyan_row,
            "RESCHEDULE_TO_KALYAN_NEW_ID",
            f"Created New ID {new_hv_id} for Kalyan from source {s2_id}",
        )
        save_data(EXCEL_FILE, data)
        st.success(f"Rescheduled to Kalyan with New HV ID: {new_hv_id} (Old ID: {s2_id})")
        st.rerun()

st.divider()

# --- SECTION 3: Kalyan Assigns Case to Anil -> Anil Active Sheet + NEW ID ---
st.subheader("3. Kalyan Assigns Case to Anil (Go to Anil Active Sheet + New HV ID)")
kalyan_assignable = master_df[
    (master_df["Case_Worker"] == "Kalyan")
    & (~master_df["Status"].str.startswith("Assigned to Anil"))
]

if kalyan_assignable.empty:
    st.info("No Kalyan cases available to assign to Anil.")
else:
    s3_id = st.selectbox(
        "Select Kalyan Case to Assign to Anil",
        options=kalyan_assignable["HV_ID"].tolist(),
        key="s3_id",
    )
    row3 = kalyan_assignable[kalyan_assignable["HV_ID"] == s3_id].iloc[0]

    c1, c2, c3 = st.columns(3)
    d3 = c1.text_input("Anil Active Date", value=row3["Scheduled_Date"], key="d3")
    t3 = c2.text_input("Anil Active Time (12-hr)", value=row3["Scheduled_Time"], key="t3")
    n3 = c3.text_input("Assignment Notes", value=row3["Notes"], key="n3")

    if st.button("Assign to Anil Active Sheet (Create New ID)", key="btn_s3"):
        new_anil_id = generate_new_hv_id(get_all_ids(data), prefix="HV")
        ts = now_12hr()

        src_idx = master_df.index[master_df["HV_ID"] == s3_id][0]
        master_df.at[src_idx, "Status"] = f"Assigned to Anil ({new_anil_id})"
        master_df.at[src_idx, "Updated_At"] = ts

        anil_active_row = {
            "HV_ID": new_anil_id,
            "Client_Name": row3["Client_Name"],
            "Case_Worker": "Anil",
            "Bucket": "Active",
            "Status": "Active",
            "Scheduled_Date": d3.strip(),
            "Scheduled_Time": t3.strip(),
            "Previous_ID": s3_id,
            "Notes": f"Assigned by Kalyan from {s3_id}. {n3.strip()}".strip(),
            "Updated_At": ts,
        }

        master_df = pd.concat(
            [master_df, pd.DataFrame([anil_active_row], columns=COLUMNS)],
            ignore_index=True,
        )
        data["Master_Cases"] = master_df
        sync_sheets(data)
        append_audit_log(
            data,
            anil_active_row,
            "ASSIGN_KALYAN_TO_ANIL_ACTIVE_NEW_ID",
            f"Assigned to Anil_Active with New ID {new_anil_id} from Kalyan ID {s3_id}",
        )
        save_data(EXCEL_FILE, data)
        st.success(
            f"Assigned to Anil Active sheet with New HV ID: {new_anil_id} (from Kalyan ID: {s3_id})"
        )
        st.rerun()

st.divider()

# --- SECTION 4: Add New Case ---
with st.expander("4. Add New Case"):
    nc1, nc2, nc3, nc4 = st.columns(4)
    nc_client = nc1.text_input("Client Name", key="nc_client")
    nc_worker = nc2.selectbox("Case Worker", ["Kalyan", "Anil"], key="nc_worker")
    nc_bucket = nc3.selectbox("Bucket", ["Booked", "Active"], key="nc_bucket")
    nc_date = nc4.text_input("Date (YYYY-MM-DD)", value="2026-11-23", key="nc_date")

    nc5, nc6 = st.columns(2)
    nc_time = nc5.text_input("Time (12-hr)", value="10:00 AM", key="nc_time")
    nc_notes = nc6.text_input("Notes", key="nc_notes")

    if st.button("Add Case", key="btn_nc"):
        if nc_client.strip():
            new_id = generate_new_hv_id(get_all_ids(data), prefix="HV")
            ts = now_12hr()
            new_rec = {
                "HV_ID": new_id,
                "Client_Name": nc_client.strip(),
                "Case_Worker": nc_worker,
                "Bucket": nc_bucket,
                "Status": nc_bucket,
                "Scheduled_Date": nc_date.strip(),
                "Scheduled_Time": nc_time.strip(),
                "Previous_ID": "",
                "Notes": nc_notes.strip(),
                "Updated_At": ts,
            }
            master_df = pd.concat(
                [master_df, pd.DataFrame([new_rec], columns=COLUMNS)],
                ignore_index=True,
            )
            data["Master_Cases"] = master_df
            sync_sheets(data)
            append_audit_log(
                data, new_rec, "ADD_NEW_CASE", f"Created new case {new_id} for {nc_worker}"
            )
            save_data(EXCEL_FILE, data)
            st.success(f"Added new case {new_id} for {nc_worker} ({nc_bucket}).")
            st.rerun()
        else:
            st.warning("Please enter a Client Name.")

st.divider()

# --- Sheet Views ---
tab1, tab2, tab3, tab4 = st.tabs(
    ["Kalyan Cases", "Anil Active Sheet", "Master Cases", "Audit Log"]
)
with tab1:
    st.dataframe(data["Kalyan_Cases"], use_container_width=True)
with tab2:
    st.dataframe(data["Anil_Active"], use_container_width=True)
with tab3:
    st.dataframe(data["Master_Cases"], use_container_width=True)
with tab4:
    st.dataframe(data["Audit_Log"], use_container_width=True)
