"""Shared CSS styles injected into every page."""

import streamlit as st


CUSTOM_CSS = """
<style>
    #MainMenu { visibility: hidden; }
    footer    { visibility: hidden; }

    .tool-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .tool-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
    .tool-header p  { margin: 0.4rem 0 0; opacity: 0.75; font-size: 0.95rem; }

    .status-banner {
        padding: 0.9rem 1.5rem;
        border-radius: 8px;
        font-weight: 600;
        margin-bottom: 1.5rem;
        font-size: 1rem;
    }
    .status-error   { background: #fef2f2; color: #991b1b; border-left: 4px solid #dc2626; }
    .status-warning { background: #fffbeb; color: #92400e; border-left: 4px solid #d97706; }
    .status-ok      { background: #f0fdf4; color: #14532d; border-left: 4px solid #16a34a; }

    .how-to-step {
        display: flex;
        align-items: flex-start;
        gap: 0.75rem;
        margin-bottom: 0.9rem;
    }
    .step-num {
        background: #0f3460;
        color: white;
        border-radius: 50%;
        width: 24px;
        height: 24px;
        min-width: 24px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.75rem;
        font-weight: 700;
        margin-top: 2px;
    }
    .step-text { font-size: 0.88rem; line-height: 1.5; color: #1e293b; }
    .step-text strong { color: #0f3460; }

    .tool-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 1.2rem 1.5rem;
        margin-bottom: 0.75rem;
        border-left: 4px solid #0f3460;
    }
    .tool-card h3 { margin: 0 0 0.3rem; font-size: 1rem; color: #0f3460; }
    .tool-card p  { margin: 0; font-size: 0.875rem; color: #64748b; }

    .stDataFrame { font-size: 0.875rem; }
</style>
"""


def apply_styles() -> None:
    """Inject the shared CSS into the current Streamlit page."""
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
