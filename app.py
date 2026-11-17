from pathlib import Path
import py_compile
p=Path("/mnt/data/app.py")
code=p.read_text(encoding="utf-8")
# Apply only to the worker dashboard visa button, not the similarly named admin button.
old_worker = '''                        btn_label = f"✈️ {v_name}: {v_count}"
                        if st.button(btn_label, key=f"worker_visa_btn_{v_name}_{idx}", use_container_width=True):'''
new_worker = '''                        visa_icon = "🪪" if "H1B" in v_name else ("⌛" if "PENDING" in v_name else ("🌐" if "B1" in v_name or "B2" in v_name else "✈️"))
                        btn_label = f"{visa_icon} {v_name}: {v_count}"
                        if st.button(btn_label, key=f"worker_visa_btn_{v_name}_{idx}", use_container_width=True):'''
if old_worker in code:
    code=code.replace(old_worker,new_worker,1)
# Insert CSS into the original file if not already present.
needle = "    .username-hyperlink {\n"
css = r"""
    /* Dashboard buttons styled to match the supplied reference image */
    [class*="st-key-worker_visa_btn"] button {
        min-height: 48px !important; border-radius: 28px !important;
        font-size: 17px !important; font-weight: 600 !important;
        color: #243b53 !important; background: #dbeafe !important;
        border: 1px solid #8db5dc !important;
        box-shadow: 0 2px 5px rgba(15, 23, 42, 0.08) !important;
    }
    [class*="st-key-worker_visa_btn_PENDING"] button {
        color: #8a6415 !important; background: #ffefc7 !important; border-color: #d9b85c !important;
    }
    [class*="st-key-worker_visa_btn_B1"] button,
    [class*="st-key-worker_visa_btn_B2"] button {
        color: #155e52 !important; background: #c9f0e8 !important; border-color: #65b7a8 !important;
    }
    [class*="st-key-worker_add_btn"] button {
        background: #0860df !important; border: 1px solid #0860df !important;
        color: #fff !important; min-height: 58px !important; border-radius: 13px !important;
        font-size: 17px !important; font-weight: 600 !important;
    }
    [class*="st-key-worker_search_btn"] button {
        background: #fff !important; border: 2px solid #2365ad !important;
        color: #245b98 !important; min-height: 58px !important; border-radius: 13px !important;
        font-size: 17px !important; font-weight: 600 !important; box-shadow: none !important;
    }
    [class*="st-key-worker_dl_disabled"] button {
        background: #e5e7eb !important; border: 1px solid #e5e7eb !important;
        color: #9298a2 !important; min-height: 58px !important; border-radius: 13px !important;
        font-size: 17px !important; font-weight: 600 !important; box-shadow: none !important; opacity: 1 !important;
    }
    [class*="st-key-worker_pw_btn"] button {
        background: #e2e6ed !important; border: 1px solid #e2e6ed !important;
        color: #3d4658 !important; min-height: 58px !important; border-radius: 13px !important;
        font-size: 17px !important; font-weight: 600 !important; box-shadow: none !important;
    }
    [class*="st-key-worker_sync_btn"] button { min-height: 58px !important; border-radius: 13px !important; }

"""
if "Dashboard buttons styled to match the supplied reference image" not in code:
    css = "    /* Dashboard buttons styled to match the supplied reference image */\n" + css
    code=code.replace(needle,css+needle,1)
out=Path("/mnt/data/app_styled.py")
out.write_text(code,encoding="utf-8")
try:
    py_compile.compile(str(out),doraise=True)
    print("Updated app_styled.py created. Python syntax check passed.")
except Exception as e:
    print("Syntax check failed:",e)
