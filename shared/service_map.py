"""Shared sidebar helper: Service ID Mapping file uploader.

Every page calls load_service_map_sidebar() at the top before any other logic.
The uploaded bytes are stored in st.session_state["service_map_bytes"] so they
persist as the user navigates between pages.
"""

import streamlit as st


def load_service_map_sidebar() -> bytes | None:
    """Render the Service ID Mapping sidebar section and return the file bytes.

    Returns the bytes of the uploaded serviceIDs.csv if loaded, or None if not.
    The calling page should call st.stop() when None is returned.
    """
    with st.sidebar:
        st.markdown("## ⚙️ Service ID Mapping")
        st.markdown("---")
        st.markdown(
            "Upload the **service ID mapping file** for this project (any filename). "
            "This file is required to use any tool.",
            help="Required columns: Service Name, Service Type, QKDBID, QKServiceID, Xplor Service ID",
        )

        uploaded = st.file_uploader(
            "Service ID mapping file (required)",
            type=["csv"],
            key="service_map_upload",
            label_visibility="collapsed",
        )

        if uploaded is not None:
            file_bytes = uploaded.read()
            st.session_state["service_map_bytes"] = file_bytes
            st.session_state["service_map_name"] = uploaded.name
            st.success(f"✅ Using: **{uploaded.name}**")
        elif "service_map_bytes" in st.session_state:
            name = st.session_state.get("service_map_name", "serviceIDs.csv")
            st.success(f"✅ Using: **{name}**")
        else:
            st.error("⚠️ Please upload a service ID mapping file to continue.")

        st.markdown("---")

    return st.session_state.get("service_map_bytes")
