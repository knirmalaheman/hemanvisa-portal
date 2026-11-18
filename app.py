from pathlib import Path
import re, py_compile

src_path = Path("/mnt/data/app (4).py")
code = src_path.read_text(encoding="utf-8")

old_nav = '''    admin_dash_tab = st.radio(
        "Admin Navigation",
        ["📊 Dashboard & Quick Actions", "📂 Case Records & Search", "ℹ️ System Guidelines"],
        horizontal=True,
        label_visibility="collapsed",
        key="admin_dash_tab_radio"
    )'''

new_nav = '''    admin_dash_tab = st.radio(
        "Admin Navigation",
        ["📊 Dashboard & Quick Actions", "📈 Leads Dashboard", "📂 Case Records & Search", "ℹ️ System Guidelines"],
        horizontal=True,
        label_visibility="collapsed",
        key="admin_dash_tab_radio"
    )'''

if old_nav not in code:
    raise RuntimeError("Could not find the Admin Navigation block in the uploaded app.py.")

code = code.replace(old_nav, new_nav, 1)

anchor = '''    elif admin_dash_tab == "📂 Case Records & Search":
        st.markdown("<div class='section-banner'>📂 Case Records & Visa Type Management Hub</div>", unsafe_allow_html=True)'''

leads_branch = '''    elif admin_dash_tab == "📈 Leads Dashboard":
        st.markdown("<div class='section-banner'>📈 Leads Dashboard & Records</div>", unsafe_allow_html=True)

        # Detect the Lead column from the existing Excel sheet headers.
        lead_cols = [c for c in df.columns if "lead" in str(c).strip().lower()]
        lead_search = st.text_input(
            "Search Leads:",
            placeholder="Search lead name, case ID, username, or any available detail...",
            key="admin_leads_search"
        ).strip()

        if not lead_cols:
            st.warning(
                "⚠️ ఈ Excel sheet లో 'Lead' అనే column కనిపించలేదు. "
                "Excel header లో Lead / Leads column add చేసి, app refresh చేయండి."
            )
        else:
            # Keep rows that contain at least one value in a Lead-related column.
            lead_mask = pd.Series(False, index=df.index)
            for lead_column in lead_cols:
                lead_mask = lead_mask | (
                    df[lead_column].notna()
                    & df[lead_column].astype(str).str.strip().ne("")
                    & df[lead_column].astype(str).str.strip().str.lower().ne("nan")
                )

            leads_df = df.loc[lead_mask].copy()

            if lead_search:
                search_mask = leads_df.apply(
                    lambda row: row.astype(str).str.contains(
                        lead_search, case=False, na=False
                    ).any(),
                    axis=1
                )
                leads_display_df = leads_df.loc[search_mask]
            else:
                leads_display_df = leads_df

            lead_count = len(leads_df)
            shown_lead_count = len(leads_display_df)

            lead_k1, lead_k2, lead_k3 = st.columns(3, gap="medium")
            with lead_k1:
                st.markdown(
                    f"<div class='kpi-card'><div class='kpi-title'>📈 Total Lead Records</div>"
                    f"<div class='kpi-value'>{lead_count}</div></div>",
                    unsafe_allow_html=True
                )
            with lead_k2:
                st.markdown(
                    f"<div class='kpi-card green'><div class='kpi-title'>🔎 Matching Records</div>"
                    f"<div class='kpi-value'>{shown_lead_count}</div></div>",
                    unsafe_allow_html=True
                )
            with lead_k3:
                st.markdown(
                    f"<div class='kpi-card purple'><div class='kpi-title'>🧾 Lead Columns</div>"
                    f"<div class='kpi-value'>{len(lead_cols)}</div></div>",
                    unsafe_allow_html=True
                )

            st.markdown("<div class='action-group-title'>📋 Lead Records</div>", unsafe_allow_html=True)
            st.caption(
                "Lead records are identified from rows where at least one Lead-related Excel column has a value."
            )
            st.dataframe(
                leads_display_df,
                use_container_width=True,
                hide_index=True
            )

            if leads_df.empty:
                st.info("ℹ️ ప్రస్తుతం ఈ sheet లో Lead data లేదు.")
            else:
                csv_bytes = leads_df.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    label="📥 Download Leads CSV",
                    data=csv_bytes,
                    file_name="leads_records.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="admin_download_leads_csv"
                )

''' + anchor

if anchor not in code:
    raise RuntimeError("Could not find the Case Records branch insertion point.")

code = code.replace(anchor, leads_branch, 1)

out_path = Path("/mnt/data/app.py")
out_path.write_text(code, encoding="utf-8")
py_compile.compile(str(out_path), doraise=True)

print(f"Updated full application saved: {out_path}")
print(f"Original lines: {len(src_path.read_text(encoding='utf-8').splitlines())}")
print(f"Updated lines: {len(code.splitlines())}")
print("Python syntax check: PASSED")
