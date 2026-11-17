from pathlib import Path
import streamlit as st

# Page configuration
st.set_page_config(
    page_title="Visa Portal Dashboard", page_icon="✈️", layout="wide"
)

# Custom CSS styling for dashboard buttons & layout
st.markdown(
    """
    <style>
    /* Dashboard buttons styled to match the reference design */
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
    .username-hyperlink { font-weight: 600; color: #0860df; }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("✈️ Visa Portal Dashboard")
st.write(
    "Welcome to the portal. Worker dashboard and visa count buttons are loaded successfully."
)

# Example implementation of worker visa buttons with icons and dynamic styling
visa_categories = [
    ("H1B Visa", 12),
    ("PENDING Cases", 5),
    ("B1/B2 Visa", 8),
    ("General Visa", 3),
]

st.subheader("Worker Visa Metrics")
cols = st.columns(len(visa_categories))

for idx, (v_name, v_count) in enumerate(visa_categories):
    with cols[idx]:
        # Determine icon based on category name
        visa_icon = (
            "🪪"
            if "H1B" in v_name
            else (
                "⌛"
                if "PENDING" in v_name
                else ("🌐" if "B1" in v_name or "B2" in v_name else "✈️")
            )
        )
        btn_label = f"{visa_icon} {v_name}: {v_count}"

        # Unique key assignment for worker dashboard buttons
        if st.button(
            btn_label, key=f"worker_visa_btn_{v_name.replace(' ', '_')}_{idx}"
        ):
            st.success(f"You clicked on {v_name}!")
