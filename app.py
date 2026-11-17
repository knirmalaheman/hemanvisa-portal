    /* 🎨 Base Button Style */
    .stButton > button, [data-testid="stDownloadButton"] > button {
        width: 100% !important;
        min-height: 46px !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        padding: 10px 16px !important;
        margin-top: 4px !important;
        margin-bottom: 8px !important;
        transition: all 0.22s ease;
        white-space: nowrap !important;
    }

    /* 💊 1. Visa Category Distribution Buttons (Pill Shape & Soft Pastel Colors) */
    button[aria-label*="H1B"] {
        background: #e0ecff !important;
        color: #1e293b !important;
        border: 1.5px solid #93c5fd !important;
        border-radius: 50px !important;
        box-shadow: none !important;
    }
    button[aria-label*="PENDING"] {
        background: #fef3c7 !important;
        color: #1e293b !important;
        border: 1.5px solid #fcd34d !important;
        border-radius: 50px !important;
        box-shadow: none !important;
    }
    button[aria-label*="B1/B2"] {
        background: #d1fae5 !important;
        color: #1e293b !important;
        border: 1.5px solid #6ee7b7 !important;
        border-radius: 50px !important;
        box-shadow: none !important;
    }
    /* ఇతర Visa కేటగిరీలకు డిఫాల్ట్ పిల్ స్టైల్ */
    button[aria-label*="✈️"] {
        border-radius: 50px !important;
    }

    /* ⚡ 2. Workspace Tools Buttons (Exact Image Colors) */
    /* 🔵 Add New Case - Solid Royal Blue */
    button[aria-label*="Add New Case"] {
        background: #1d4ed8 !important;
        color: #ffffff !important;
        border: none !important;
        box-shadow: 0 4px 10px rgba(29, 78, 216, 0.25) !important;
    }

    /* ⚪ Search Records - White Fill with Blue Outline */
    button[aria-label*="Search Records"] {
        background: #ffffff !important;
        color: #1d4ed8 !important;
        border: 2px solid #2563eb !important;
        box-shadow: none !important;
    }

    /* 🚫 No Download - Light Gray Disabled Style */
    button[aria-label*="No Download"], button:disabled {
        background: #e2e8f0 !important;
        color: #94a3b8 !important;
        border: 1px solid #cbd5e1 !important;
        box-shadow: none !important;
        opacity: 1 !important;
    }

    /* 🔑 Change Password - Soft Slate Gray */
    button[aria-label*="Change Password"] {
        background: #e2e8f0 !important;
        color: #334155 !important;
        border: 1px solid #cbd5e1 !important;
        box-shadow: none !important;
    }

    /* 📥 Export My Excel (Download అందుబాటులో ఉన్నప్పుడు) */
    [data-testid="stDownloadButton"] > button {
        background: #059669 !important;
        color: #ffffff !important;
        border: none !important;
    }

    /* ✨ Hover Effects */
    .stButton > button:not(:disabled):hover, [data-testid="stDownloadButton"] > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(30, 58, 138, 0.18) !important;
        opacity: 0.95;
    }
