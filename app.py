            
            if st.session_state["selected_visa_filter"] != "All":
                if st.button(f"🔄 Active Filter: [{st.session_state['selected_visa_filter']}] — Click to Reset (Show All My Cases)", use_container_width=True, key="reset_worker_visa_filter"):
                    st.session_state["selected_visa_filter"] = "All"
                    st.rerun()

        st.markdown("<div class='action-group-title'>⚡ Workspace Tools</div>", unsafe_allow_html=True)
        has_sync_perm = worker_perm.get("can_sync", False)
        w_cols = st.columns(5 if has_sync_perm else 4, gap="medium")
        
        with w_cols[0]:
            if worker_perm.get("can_add", True):
                st.markdown('<span class="admin-btn-1" style="display:none;"></span>', unsafe_allow_html=True)
                if st.button("➕ Add New Case", use_container_width=True, key="worker_add_btn"):
                    st.session_state["open_dialog"] = "add_case"
                    st.rerun()
            else:
                st.markdown('<span class="admin-btn-7" style="display:none;"></span>', unsafe_allow_html=True)
                st.button("🚫 No Add Access", disabled=True, use_container_width=True, key="worker_add_disabled")
        with w_cols[1]:
            if worker_perm.get("can_search", True):
                st.markdown('<span class="admin-btn-2" style="display:none;"></span>', unsafe_allow_html=True)
                if st.button("🔍 Search Records", use_container_width=True, key="worker_search_btn"):
                    st.session_state["open_dialog"] = "search_data"
                    st.rerun()
            else:
                st.markdown('<span class="admin-btn-7" style="display:none;"></span>', unsafe_allow_html=True)
                st.button("🚫 No Search Access", disabled=True, use_container_width=True, key="worker_search_disabled")
        with w_cols[2]:
            if worker_perm.get("can_download", False):
                st.markdown('<span class="admin-btn-3" style="display:none;"></span>', unsafe_allow_html=True)
                worker_excel_bytes = get_worker_export_bytes(_get_file_mtime(), current_user, current_sheet)
                st.download_button(
                    label="📥 Export My Excel", 
                    data=worker_excel_bytes, 
                    file_name=f"{current_user}_all_sheets_cases.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    use_container_width=True, 
                    key="worker_dl_btn"
                )
            else:
                st.markdown('<span class="admin-btn-7" style="display:none;"></span>', unsafe_allow_html=True)
                st.button("🚫 No Download", disabled=True, use_container_width=True, key="worker_dl_disabled")
        with w_cols[3]:
            st.markdown('<span class="admin-btn-4" style="display:none;"></span>', unsafe_allow_html=True)
            if st.button("🔑 Change Password", use_container_width=True, key="worker_pw_btn"):
                st.session_state["open_dialog"] = "worker_password"
                st.rerun()

        if has_sync_perm:
            with w_cols[4]:
                st.markdown('<span class="admin-btn-5" style="display:none;"></span>', unsafe_allow_html=True)
                if st.button("🔄 Sync Excel", use_container_width=True, key="worker_sync_btn"):
                    st.session_state["open_dialog"] = "sync_excel"
                    st.rerun()
                
        st.divider()

        if sheets:
            st.markdown("<div class='action-group-title'>📑 Switch Sheet:</div>", unsafe_allow_html=True)
            sw_cols_w = st.columns(len(sheets), gap="small")
            for idx_w, s_name_w in enumerate(sheets):
                with sw_cols_w[idx_w]:
                    is_active_w = (s_name_w == current_sheet)
                    color_cls_w = "admin-btn-1" if is_active_w else "admin-api-6" if False else "admin-btn-6"
                    btn_txt_w = f"✅ {s_name_w}" if is_active_w else f"📄 {s_name_w}"
                    st.markdown(f'<span class="{color_cls_w}" style="display:none;"></span>', unsafe_allow_html=True)
                    if st.button(btn_txt_w, key=f"worker_switch_sheet_{s_name_w}_{idx_w}", use_container_width=True):
                        st.session_state["active_sheet"] = s_name_w
                        st.query_params["active_sheet"] = s_name_w
                        st.rerun()
                
        st.divider()
        
        display_user_rows = user_rows
        if st.session_state["selected_visa_filter"] != "All" and visa_col:
            target_visa = st.session_state["selected_visa_filter"]
            display_user_rows = display_user_rows[display_user_rows[visa_col].astype(str).str.strip().str.upper() == target_visa]

        st.markdown(f"<div class='section-banner'>✏️ Your Assigned Cases ({current_sheet} — Total: {len(display_user_rows)})</div>", unsafe_allow_html=True)

        if display_user_rows.empty:
            st.info("No cases found matching your criteria in this sheet.")
        else:
            st.caption("💡 Click USERNAME inside the table to view, edit, or request deletion.")
            render_clickable_table(display_user_rows, display_c, p_col, table_key=f"worker_{current_user}_{current_sheet}")

    elif worker_dash_tab == "ℹ️ Case Worker Guide":
        st.markdown("<div class='section-banner'>ℹ️ Case Worker Quick Guide & Instructions</div>", unsafe_allow_html=True)
        w_g1, w_g2 = st.columns(2, gap="medium")
        with w_g1:
            st.markdown("""
                <div class="corporate-card">
                    <h3>✏️ Editing & Updating Cases</h3>
                    <p>Click on any USERNAME inside your assigned cases table to open the pop-up modal. You can review all details and update information instantly.</p>
                </div>
                <div class="corporate-card">
                    <h3>➕ Adding New Cases</h3>
                    <p>Use the 'Add New Case' button in the toolbar to create new records directly in your designated sheet with automatic timestamping.</p>
                </div>
            """, unsafe_allow_html=True)
        with w_g2:
            st.markdown("""
                <div class="corporate-card">
                    <h3>🗑️ Requesting Case Deletions</h3>
                    <p>If a case needs removal, click 'Request Delete' and provide a mandatory reason. Admin reviews and approves/rejects your request.</p>
                </div>
                <div class="corporate-card">
                    <h3>🔒 Security & Password</h3>
                    <p>Keep your account secure by changing your password periodically using the 'Change Password' button in the toolbar.</p>
                </div>
            """, unsafe_allow_html=True)
