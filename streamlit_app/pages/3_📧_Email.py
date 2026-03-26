"""
📧 Email — DrawingAI Pro
Email management: test connections, browse folders, download attachments
"""
import streamlit as st
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from dotenv import load_dotenv
load_dotenv(PROJECT_ROOT / ".env")

from streamlit_app.backend.config_manager import load_config
from streamlit_app.backend.email_helpers import (
    parse_mailboxes_text,
    test_mailbox_connection,
    load_folders_for_mailbox,
    format_folder_label,
)
from streamlit_app.brand import BRAND_CSS, brand_header, sidebar_logo

st.set_page_config(page_title="📧 Email — DrawingAI Pro", page_icon="🌿", layout="wide")

st.markdown(BRAND_CSS, unsafe_allow_html=True)
sidebar_logo()
st.html(brand_header("אימייל — DrawingAI Pro"))

cfg = load_config()

# Session state for this page
if "email_folders" not in st.session_state:
    st.session_state.email_folders = []

# Mailbox input
mailbox = st.text_input(
    "תיבה משותפת:",
    value=cfg.get("shared_mailbox", ""),
    key="email_mailbox",
)

col1, col2 = st.columns(2)
with col1:
    if st.button("🔌 בדוק חיבור", width='stretch'):
        if mailbox:
            with st.spinner("בודק חיבור..."):
                success, msg = test_mailbox_connection(mailbox)
            if success:
                st.success(msg)
            else:
                st.error(msg)
        else:
            st.warning("הזן כתובת תיבה")

with col2:
    if st.button("📂 טען תיקיות", width='stretch'):
        if mailbox:
            with st.spinner("טוען תיקיות..."):
                folders, err = load_folders_for_mailbox(mailbox)
            if err:
                st.error(err)
            else:
                st.session_state.email_folders = folders
                st.success(f"נטענו {len(folders)} תיקיות")
        else:
            st.warning("הזן כתובת תיבה")

# Show folders
if st.session_state.email_folders:
    st.subheader("📁 תיקיות")
    folder_data = []
    for f in st.session_state.email_folders:
        folder_data.append({
            "נתיב": f["path"],
            "פריטים": f.get("totalItemCount", "—"),
        })
    st.dataframe(folder_data, width='stretch')
