"""
app.py — Streamlit frontend for the AI-Powered PowerPoint Generator.

This is the entry point. Run with:
    streamlit run app.py
"""

from __future__ import annotations

import io
import traceback
import uuid
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import streamlit as st

# ---------------------------------------------------------------------------
# Page configuration (must be first Streamlit call)
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="AI PowerPoint Generator",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Bootstrap path & logging
# ---------------------------------------------------------------------------
import sys
PROJECT_ROOT = Path(__file__).parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from utils.logging_utils import configure_logging, get_logger
configure_logging()
logger = get_logger(__name__)

# ---------------------------------------------------------------------------
# Imports (after path setup)
# ---------------------------------------------------------------------------
import config
from core.content_analyzer import analyze_content, analyze_content_semantic, ContentAnalysisError
from core.document_parser import parse_document, parse_document_with_structure, DocumentParseError
from core.document_structure import render_structure_text, count_sections
from core.layout_manager import LayoutManager, LayoutManagerError
from core.presentation_planner import plan_presentation, PresentationPlanningError
from core.pptx_builder import build_presentation, PptxBuilderError
from core.validator import validate_plan, validate_pptx_bytes
from core.presentation_intelligence import ContentSource, classify_presentation
from core.governance import (
    validate_plan_governance,
    validate_pptx_governance,
    write_audit_record,
)
from llm.groq_client import GroqClient, GroqAuthError, GroqRateLimitError, GroqAPIError
from llm.watsonx_client import WatsonxAuthError, WatsonxRateLimitError, WatsonxAPIError
from llm.key_manager import KeyManager, KeyManagerError
from utils.file_utils import validate_upload, get_output_path, cleanup_old_outputs
from utils.text_utils import sanitize_filename, truncate_text


# ---------------------------------------------------------------------------
# Custom CSS — Executive White & Corporate Navy Blue Theme
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    :root {
        --bg-page: #F8FAFC;
        --bg-card: #FFFFFF;
        --navy-dark: #081E39;
        --navy-primary: #0B2545;
        --navy-medium: #134074;
        --navy-light: #1E40AF;
        --blue-subtle: #EEF4FA;
        --blue-border: #E0EDF8;
        --gold-cta: #FFB800;
        --gold-hover: #E5A600;
        --gold-active: #CCA000;
        --text-navy: #0B2545;
        --text-primary: #0F172A;
        --text-secondary: #334155;
        --text-muted: #64748B;
        --text-white: #FFFFFF;
        --border-color: #E2E8F0;
        --border-medium: #CBD5E1;
        --card-shadow: 0 4px 20px -2px rgba(11, 37, 69, 0.06), 0 2px 6px -1px rgba(11, 37, 69, 0.04);
        --card-shadow-hover: 0 12px 28px -4px rgba(11, 37, 69, 0.12);
    }

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        color: var(--text-primary);
    }

    .stApp {
        background-color: var(--bg-page) !important;
        color: var(--text-primary) !important;
    }

    /* Top Brand Navigation Bar */
    .top-nav-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #FFFFFF;
        border-bottom: 1px solid var(--border-color);
        padding: 0.85rem 1.5rem;
        margin: -4rem -3rem 1.5rem -3rem;
        box-shadow: 0 1px 3px rgba(11, 37, 69, 0.04);
    }

    .nav-brand {
        display: flex;
        align-items: center;
        gap: 10px;
    }

    .nav-brand-icon {
        font-size: 1.5rem;
    }

    .nav-brand-text {
        font-size: 1.15rem;
        font-weight: 800;
        color: var(--navy-primary);
        letter-spacing: -0.01em;
    }

    .nav-items {
        display: flex;
        align-items: center;
        gap: 1.2rem;
    }

    .nav-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: var(--blue-subtle);
        color: var(--navy-primary);
        padding: 4px 12px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
        border: 1px solid var(--border-medium);
    }

    .status-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background-color: #10B981;
        display: inline-block;
    }

    .nav-item {
        color: var(--text-muted);
        font-size: 0.85rem;
        font-weight: 500;
    }

    /* Sidebar Styling */
    [data-testid="stSidebar"] {
        background: #FFFFFF !important;
        border-right: 1px solid var(--border-color) !important;
        box-shadow: 2px 0 12px rgba(11, 37, 69, 0.03) !important;
    }

    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] strong {
        color: var(--navy-primary) !important;
    }

    /* Cards */
    .ppt-card {
        background: var(--bg-card);
        border: 1px solid var(--border-color);
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: var(--card-shadow);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }

    .ppt-card:hover {
        box-shadow: var(--card-shadow-hover);
    }

    /* Executive Hero Header (Deep Navy with Warm Gold Accent) */
    .hero-header {
        text-align: left;
        padding: 2.8rem 2.5rem 2.6rem;
        background: linear-gradient(135deg, #081E39 0%, #0B2545 60%, #134074 100%);
        border-radius: 16px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        margin-bottom: 2rem;
        position: relative;
        overflow: hidden;
        box-shadow: 0 14px 32px -6px rgba(11, 37, 69, 0.28);
    }

    .hero-header::after {
        content: "";
        position: absolute;
        top: -60%;
        right: -15%;
        width: 480px;
        height: 480px;
        background: radial-gradient(circle, rgba(255,184,0,0.14) 0%, rgba(30,64,175,0.1) 45%, transparent 70%);
        border-radius: 50%;
        pointer-events: none;
    }

    .hero-title {
        font-size: 2.6rem;
        font-weight: 800;
        color: #FFFFFF !important;
        margin: 0;
        line-height: 1.2;
        letter-spacing: -0.02em;
    }

    .hero-accent-line {
        width: 54px;
        height: 4px;
        background: var(--gold-cta);
        border-radius: 2px;
        margin: 1.1rem 0 1rem 0;
    }

    .hero-subtitle {
        color: #E2E8F0 !important;
        font-size: 1.08rem;
        margin: 0;
        font-weight: 400;
        line-height: 1.6;
        max-width: 820px;
    }

    .hero-badges {
        display: flex;
        flex-wrap: wrap;
        gap: 0.6rem;
        margin-top: 1.3rem;
    }

    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 5px 14px;
        border-radius: 20px;
        background: rgba(255, 255, 255, 0.12);
        border: 1px solid rgba(255, 255, 255, 0.22);
        color: #FFFFFF;
        font-size: 0.82rem;
        font-weight: 500;
        backdrop-filter: blur(6px);
    }

    /* Step badge */
    .step-badge {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        background: var(--navy-primary);
        color: white;
        font-size: 0.8rem;
        font-weight: 700;
        margin-right: 8px;
        flex-shrink: 0;
    }

    /* Slide preview card */
    .slide-preview {
        background: #FFFFFF;
        border: 1px solid var(--border-color);
        border-radius: 8px;
        padding: 0.85rem 1.15rem;
        margin-bottom: 0.55rem;
        display: flex;
        align-items: center;
        gap: 14px;
        box-shadow: 0 2px 6px rgba(11, 37, 69, 0.03);
        transition: transform 0.15s ease, border-color 0.15s ease;
    }

    .slide-preview:hover {
        border-color: var(--navy-light);
        transform: translateY(-1px);
    }

    .slide-number {
        font-size: 0.8rem;
        font-weight: 700;
        color: var(--navy-primary);
        background: var(--blue-subtle);
        border-radius: 6px;
        padding: 4px 8px;
        min-width: 32px;
        text-align: center;
    }

    .layout-badge {
        font-size: 0.7rem;
        font-weight: 600;
        padding: 3px 10px;
        border-radius: 20px;
        background: var(--blue-subtle);
        color: var(--navy-primary);
        border: 1px solid var(--border-medium);
        white-space: nowrap;
        margin-left: auto;
    }

    /* Metric cards */
    .metric-row {
        display: flex;
        gap: 1rem;
        margin-bottom: 1rem;
    }

    .metric-card {
        flex: 1;
        background: #FFFFFF;
        border: 1px solid var(--border-color);
        border-radius: 12px;
        padding: 1.25rem 1rem;
        text-align: center;
        box-shadow: var(--card-shadow);
        border-top: 4px solid var(--navy-primary);
    }

    .metric-value {
        font-size: 2.2rem;
        font-weight: 800;
        color: var(--navy-primary);
        line-height: 1.1;
    }

    .metric-label {
        font-size: 0.82rem;
        font-weight: 600;
        color: var(--text-muted);
        margin-top: 6px;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }

    /* Status indicators */
    .status-success { color: #059669; }
    .status-warning { color: #D97706; }
    .status-error { color: #DC2626; }

    /* Primary Action Buttons (Warm Gold / Amber CTA from Reference Image) */
    div.stButton > button:first-child[kind="primary"],
    div.stButton > button[data-testid="baseButton-primary"] {
        background: var(--gold-cta) !important;
        color: var(--navy-dark) !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.75rem 2rem !important;
        box-shadow: 0 4px 14px rgba(255, 184, 0, 0.38) !important;
        transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }

    div.stButton > button:first-child[kind="primary"]:hover,
    div.stButton > button[data-testid="baseButton-primary"]:hover {
        background: var(--gold-hover) !important;
        color: var(--navy-dark) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 22px rgba(255, 184, 0, 0.48) !important;
    }

    /* Secondary Buttons */
    div.stButton > button:first-child[kind="secondary"],
    div.stButton > button[data-testid="baseButton-secondary"] {
        background: #FFFFFF !important;
        color: var(--navy-primary) !important;
        border: 1px solid var(--border-medium) !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        transition: all 0.2s ease !important;
    }

    div.stButton > button:first-child[kind="secondary"]:hover,
    div.stButton > button[data-testid="baseButton-secondary"]:hover {
        background: var(--blue-subtle) !important;
        border-color: var(--navy-medium) !important;
        color: var(--navy-dark) !important;
    }

    /* Sleek Lower Corporate Banner for Presentation Download */
    .download-banner {
        background: linear-gradient(135deg, #081E39 0%, #0B2545 100%);
        border-radius: 12px;
        padding: 1.8rem 2.2rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 2rem;
        box-shadow: 0 8px 24px -3px rgba(11, 37, 69, 0.22);
        margin-top: 1.5rem;
    }

    .download-banner-left {
        flex: 1;
    }

    .download-banner-title {
        color: #FFFFFF !important;
        font-size: 1.35rem;
        font-weight: 700;
        margin: 0 0 0.35rem 0;
        letter-spacing: -0.01em;
    }

    .download-banner-subtitle {
        color: #CBD5E1 !important;
        font-size: 0.92rem;
        margin: 0;
        line-height: 1.5;
    }

    .download-gold-cta {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: var(--gold-cta);
        color: var(--navy-dark) !important;
        font-weight: 700;
        font-size: 1rem;
        padding: 0.8rem 1.8rem;
        border-radius: 8px;
        text-decoration: none !important;
        box-shadow: 0 4px 14px rgba(255, 184, 0, 0.38);
        transition: all 0.2s ease;
        white-space: nowrap;
    }

    .download-gold-cta:hover {
        background: var(--gold-hover);
        transform: translateY(-2px);
        box-shadow: 0 8px 22px rgba(255, 184, 0, 0.48);
        color: var(--navy-dark) !important;
    }

    /* Streamlit Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        background-color: transparent;
        border-bottom: 2px solid var(--border-color);
        padding-bottom: 0px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 6px 6px 0 0;
        padding: 10px 22px;
        background-color: transparent;
        color: var(--text-secondary);
        font-weight: 600;
        border: none;
    }

    .stTabs [aria-selected="true"] {
        background-color: transparent !important;
        color: var(--navy-primary) !important;
        border-bottom: 3px solid var(--navy-primary) !important;
        font-weight: 700 !important;
    }

    /* Expander styling */
    .streamlit-expanderHeader {
        background: #FFFFFF !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        color: var(--navy-primary) !important;
    }

    /* Form Controls */
    .stTextInput input, .stTextArea textarea, .stSelectbox [data-baseweb="select"] {
        background: #FFFFFF !important;
        color: var(--text-primary) !important;
        border: 1px solid var(--border-medium) !important;
        border-radius: 8px !important;
    }

    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: var(--navy-primary) !important;
        box-shadow: 0 0 0 3px rgba(11, 37, 69, 0.12) !important;
    }

    /* Divider */
    hr {
        border-color: var(--border-color) !important;
        margin: 1.8rem 0 !important;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Session state helpers
# ---------------------------------------------------------------------------

def _init_session_state() -> None:
    defaults = {
        "pptx_bytes": None,
        "pptx_filename": None,
        "presentation_plan": None,
        "generation_complete": False,
        "error_message": None,
        # Semantic pipeline extras
        "chunk_analyses": None,     # List[ChunkAnalysis] for debug view
        "doc_structure": None,      # DocumentSection root
        "analysis_cache": None,     # AnalysisCache
        "semantic_chunks": None,    # List[SemanticChunk]
        # Document upload state (persisted)
        "uploaded_file_bytes": None,
        "uploaded_filename": None,
        "extracted_text": None,
        "doc_root": None,
        "doc_detection_method": None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


_init_session_state()


# ---------------------------------------------------------------------------
# Sidebar — configuration
# ---------------------------------------------------------------------------

def render_sidebar() -> dict:
    """Render the sidebar and return configuration values."""
    with st.sidebar:
        st.markdown("""
        <div style="text-align:center; padding: 1rem 0 0.75rem;">
            <span style="font-size:2.2rem;">📊</span>
            <h2 style="color:#0B2545; margin:0.4rem 0 0; font-size:1.25rem; font-weight:800; letter-spacing:-0.01em;">AI PPT Studio</h2>
            <p style="color:#64748B; font-size:0.82rem; margin:0.25rem 0 0; font-weight:500;">Powered by IBM watsonx.ai</p>
        </div>
        """, unsafe_allow_html=True)

        st.divider()

        is_watsonx = config.LLM_PROVIDER == "watsonx"
        provider_label = "watsonx.ai" if is_watsonx else "Groq"
        provider_keys = config.WATSONX_API_KEYS if is_watsonx else config.GROQ_API_KEYS
        provider_env_var = "WATSONX_API_KEY" if is_watsonx else "GROQ_API_KEYS"

        # ── API Keys — multi-key, never displayed back ───────────────────
        st.markdown(f"**🔑 {provider_label} API Keys**")

        # Show keys from .env as pre-loaded (count only, no display)
        env_key_count = len(provider_keys)
        if env_key_count > 0:
            st.success(f"✅ {env_key_count} key(s) loaded from configuration")

        # Allow adding more keys via UI (password type, never echoed)
        extra_keys_raw = st.text_area(
            "Additional API Keys",
            height=80,
            placeholder="Paste extra API keys here (one per line)",
            help=f"Optional. Add more {provider_label} API keys to distribute load across sections. Keys entered here are never shown.",
            label_visibility="visible",
        )

        # Combine env keys + UI keys
        ui_keys = [k.strip() for k in (extra_keys_raw or "").splitlines() if k.strip()]
        all_keys = list(dict.fromkeys(provider_keys + ui_keys))  # deduplicate

        if all_keys:
            st.caption(f"🔀 {len(all_keys)} key(s) active — round-robin across sections")
        else:
            st.warning(f"⚠️ No API keys configured. Add keys above or set {provider_env_var} in .env")

        st.divider()

        # Model selection
        st.markdown("**🤖 Model**")
        if is_watsonx:
            model_options = [
                "openai/gpt-oss-120b",
                "meta-llama/llama-3-3-70b-instruct",
                "meta-llama/llama-3-1-8b",
                "mistralai/mistral-small-3-1-24b-instruct-2503",
                "mistralai/mistral-medium-2505",
            ]
            default_model = config.WATSONX_MODEL_ID
        else:
            model_options = [
                "openai/gpt-oss-120b",
                "openai/gpt-oss-20b",
                "qwen/qwen3.6-27b",
                "qwen/qwen3.8-27b",
                "groq/compound",
                "groq/compound-mini",
            ]
            default_model = config.GROQ_MODEL
        default_idx = model_options.index(default_model) if default_model in model_options else 0
        model = st.selectbox(
            "Model",
            options=model_options,
            index=default_idx,
            label_visibility="collapsed",
        )

        st.divider()
        st.caption(f"v{config.APP_VERSION} • Corporate White & Navy Blue Edition")

    return {"api_keys": all_keys, "model": model}


# ---------------------------------------------------------------------------
# Main content
# ---------------------------------------------------------------------------

def render_header() -> None:
    # st.markdown("""
    # <div class="top-nav-bar">
    #     <div class="nav-brand">
    #         <span class="nav-brand-icon">📊</span>
    #         <span class="nav-brand-text">AI POC SOLUTION</span>
    #     </div>
    #     <div class="nav-items">
    #         <span class="nav-badge"><span class="status-dot"></span> watsonx.ai Active</span>
    #         <span class="nav-item">Corporate Templates</span>
    #         <span class="nav-item">Smart Semantic Chunker</span>
    #     </div>
    # </div>

    # <div class="hero-header">
    #     <h1 class="hero-title">AI-Powered Branding Agent</h1>
    #     <div class="hero-accent-line"></div>
    #     <p class="hero-subtitle">
    #         Business Proposal | From raw business content to executive-ready, editable presentations
    #     </p>
    #     <div class="hero-badges">
    #         <span class="hero-badge">⚡ Ultra-Fast Inference</span>
    #         <span class="hero-badge">🏢 Corporate Templates</span>
    #         <span class="hero-badge">📊 Dynamic Visual Archetypes</span>
    #         <span class="hero-badge">🔒 100% Source-Grounded</span>
    #     </div>
    # </div>
    # """, unsafe_allow_html=True)
    st.markdown("""
    <div class="top-nav-bar">
        <div class="nav-brand">
            <span class="nav-brand-text">AI POC SOLUTION</span>
        </div>
        <div class="nav-items">
            <span class="nav-item">Corporate Templates</span>
        </div>
    </div>

    <div class="hero-header">
        <h1 class="hero-title">AI-Powered Branding Agent</h1>
        <div class="hero-accent-line"></div>
        <p class="hero-subtitle">
            Business Proposal | From raw business content to executive-ready, editable presentations
        </p>
    </div>
    """, unsafe_allow_html=True)


def render_input_section() -> tuple:
    """
    Render the dual-mode input section.
    Returns (source_text, input_mode, file_bytes, filename, doc_root, detection_method)
    """
    st.markdown("### 📥 Input")

    tab_doc, tab_prompt = st.tabs(["📄 Upload Document", "✍️ Enter Prompt / Topic"])
    source_text: Optional[str] = None
    input_mode: Optional[str] = None
    file_bytes_out = None
    filename_out = None
    doc_root_out = None
    detection_method_out = None

    with tab_doc:
        st.markdown("Upload a document — headings and structure are preserved for semantic analysis.")
        uploaded_file = st.file_uploader(
            "Choose a file",
            type=config.SUPPORTED_FILE_TYPES,
            help=f"Supported formats: {', '.join(f'.{t}' for t in config.SUPPORTED_FILE_TYPES)}. "
                 f"Maximum size: {config.MAX_UPLOAD_SIZE_MB} MB",
        )

        if uploaded_file is not None:
            file_bytes = uploaded_file.read()
            error = validate_upload(
                filename=uploaded_file.name,
                file_bytes=file_bytes,
                max_size_bytes=config.MAX_UPLOAD_SIZE_BYTES,
            )
            if error:
                st.error(f"❌ {error}")
            else:
                st.success(
                    f"✅ **{uploaded_file.name}** loaded "
                    f"({len(file_bytes) / 1024:.1f} KB)"
                )
                # Only re-parse if file changed
                if (st.session_state.get("uploaded_filename") != uploaded_file.name or
                        st.session_state.get("extracted_text") is None):
                    with st.spinner("Extracting text and detecting document structure..."):
                        try:
                            plain_text, doc_root, detection_method = parse_document_with_structure(
                                file_bytes, uploaded_file.name
                            )
                            st.session_state.extracted_text = plain_text
                            st.session_state.doc_root = doc_root
                            st.session_state.doc_detection_method = detection_method
                            st.session_state.uploaded_filename = uploaded_file.name
                        except DocumentParseError as e:
                            st.error(f"❌ Document parsing failed: {e}")
                        except Exception as e:
                            st.error(f"❌ Unexpected error reading file: {e}")

                if st.session_state.get("extracted_text"):
                    plain_text = st.session_state.extracted_text
                    doc_root = st.session_state.doc_root
                    detection_method = st.session_state.doc_detection_method

                    source_text = plain_text
                    input_mode = "document"
                    file_bytes_out = file_bytes
                    filename_out = uploaded_file.name
                    doc_root_out = doc_root
                    detection_method_out = detection_method

                    char_count = len(plain_text)
                    st.info(
                        f"📝 Extracted **{char_count:,}** characters "
                        f"({'within limit' if char_count <= config.MAX_SOURCE_TEXT_CHARS else 'large — semantic chunking will split it'})"
                    )

                    # Show detected structure tree
                    h1, h2plus = count_sections(doc_root)
                    if h1 > 0:
                        with st.expander(
                            f"🏗️ Detected Document Structure — {h1} heading(s), {h2plus} sub-heading(s)",
                            expanded=True
                        ):
                            tree_text = render_structure_text(doc_root)
                            st.markdown(f"```\n{tree_text}\n```")
                            st.caption(f"Detection method: `{detection_method}`")
                    else:
                        st.info(
                            "ℹ️ No explicit headings detected. "
                            "The AI will automatically infer document structure before analysis."
                        )

                    with st.expander("Preview extracted text (first 500 chars)"):
                        st.text(plain_text[:500] + ("..." if len(plain_text) > 500 else ""))

    with tab_prompt:
        st.markdown(
            "Describe your presentation topic or paste your content directly. "
            "The AI will use this as the authoritative source."
        )
        prompt_text = st.text_area(
            "Your prompt or content",
            height=250,
            placeholder=(
                "Example:\n\n"
                "Create a presentation about our company's AI-powered logistics platform.\n"
                "Cover: platform overview, AI architecture, business benefits, "
                "operational challenges, performance metrics (40% faster processing, "
                "99.9% uptime), and our 2025-2027 expansion roadmap."
            ),
            label_visibility="collapsed",
        )

        if prompt_text and prompt_text.strip():
            source_text = prompt_text.strip()
            input_mode = "prompt"
            st.caption(f"✅ {len(source_text):,} characters entered")

    return source_text, input_mode, file_bytes_out, filename_out, doc_root_out, detection_method_out


def render_configuration() -> dict:
    """Render presentation configuration options. Returns config dict."""
    st.markdown("### ⚙️ Presentation Configuration")

    with st.expander("Configure your presentation", expanded=True):

        # ── Template Selection ─────────────────────────────────────────────
        st.markdown("**Presentation Template**")
        from core.template_registry import list_templates
        templates  = list_templates()
        tmpl_labels = [f"{t['icon']} {t['name']} — {t['description']}" for t in templates]
        tmpl_ids    = [t["id"] for t in templates]
        tmpl_idx = st.selectbox(
            "Template",
            options=range(len(tmpl_labels)),
            format_func=lambda i: tmpl_labels[i],
            index=0,
            label_visibility="collapsed",
            key="template_selector",
        )
        selected_template_id = tmpl_ids[tmpl_idx]
        st.caption(
            f"Selected: **{templates[tmpl_idx]['name']}** — "
            f"{templates[tmpl_idx]['description']}"
        )

        st.markdown("---")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Number of Slides**")
            auto_slides = st.checkbox("Automatic (LLM decides)", value=False)

            if auto_slides:
                slide_count = None
                st.caption("The LLM will determine the optimal number of slides.")
            else:
                slide_count = st.slider(
                    "Slides",
                    min_value=config.MIN_SLIDES,
                    max_value=config.MAX_SLIDES,
                    value=config.DEFAULT_SLIDES,
                    step=1,
                    label_visibility="collapsed",
                )
                st.caption(f"📊 {slide_count} content slides (+ 4 mandatory: Title, Agenda, Executive Summary, Conclusion)")

            st.markdown("**Presentation Title**")
            pres_title = st.text_input(
                "Title",
                placeholder="Leave blank to auto-generate from content",
                label_visibility="collapsed",
            )

            # st.markdown("**Target Audience**")
            # audience = st.selectbox(
            #     "Audience",
            #     options=config.AUDIENCE_OPTIONS,
            #     index=0,
            #     label_visibility="collapsed",
            # )

        with col2:
            st.markdown("**Presentation Style**")
            style = st.selectbox(
                "Style",
                options=config.STYLE_OPTIONS,
                index=0,
                label_visibility="collapsed",
            )

            st.markdown("**Language**")
            language = st.selectbox(
                "Language",
                options=config.LANGUAGE_OPTIONS,
                index=0,
                label_visibility="collapsed",
            )

            # st.markdown("**Additional Instructions**")
            # additional_instructions = st.text_area(
            #     "Instructions",
            #     height=100,
            #     placeholder="E.g.: Focus on technical details. Include ROI analysis. Keep slides concise.",
            #     label_visibility="collapsed",
            # )
            st.markdown("**Target Audience**")
            audience = st.selectbox(
                "Audience",
                options=config.AUDIENCE_OPTIONS,
                index=0,
                label_visibility="collapsed",
            )

    return {
        "slide_count": slide_count,
        "presentation_title": pres_title.strip() if pres_title else "",
        "audience": audience,
        "style": style,
        "language": language,
        "additional_instructions": "",
        # "additional_instructions": additional_instructions.strip() if additional_instructions else "",
        "template_id": selected_template_id,
    }



def render_preview(plan) -> None:
    """Render the presentation preview after generation."""
    st.markdown("---")
    st.markdown("### 📋 Presentation Preview")

    # Metrics row
    layout_counts: dict = {}
    for slide in plan.slides:
        lt = getattr(slide, "layout_pattern", getattr(slide, "layout_type", "CONTENT")).upper()
        layout_counts[lt] = layout_counts.get(lt, 0) + 1

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{len(plan.slides)}</div>
            <div class="metric-label">Total Slides</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{len(layout_counts)}</div>
            <div class="metric-label">Layout Types Used</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        words = sum(
            len(str(getattr(s, "composition", getattr(s, "content", ""))).split())
            for s in plan.slides
        )
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{words}</div>
            <div class="metric-label">Words Generated</div>
        </div>
        """, unsafe_allow_html=True)

    # Slide list
    st.markdown(f"**Title:** {plan.title}")
    if getattr(plan, "subtitle", None):
        st.markdown(f"**Subtitle:** {plan.subtitle}")

    # Structural badge map — visual role indicators for mandatory slides
    _STRUCTURE_BADGES = {
        "TITLE":             ("🎯", "#0B2545", "#FFFFFF", "COVER"),
        "EXECUTIVE_SUMMARY": ("⭐", "#1E40AF", "#FFFFFF", "EXEC SUMMARY"),
        "AGENDA":            ("📋", "#134074", "#FFFFFF", "AGENDA"),
        "SECTION_HEADER":    ("📌", "#EEF4FA", "#0B2545", "SECTION"),
        "CONCLUSION":        ("✅", "#059669", "#FFFFFF", "CONCLUSION"),
    }

    st.markdown("**Slides:**")
    for slide in plan.slides:
        layout_type = getattr(slide, "layout_pattern", getattr(slide, "layout_type", "CONTENT")).upper()
        content = getattr(slide, "composition", getattr(slide, "content", {}))
        if hasattr(content, "model_dump"):
            content = content.model_dump()
        elif not isinstance(content, dict):
            content = {}

        # Rich content preview per layout type
        preview_text = ""
        if getattr(slide, "executive_takeaway", None):
            preview_text = slide.executive_takeaway
        elif layout_type == "EXECUTIVE_SUMMARY":
            highlights = content.get("highlights", [])
            if highlights:
                preview_text = " • ".join(str(h)[:45] for h in highlights[:2])
        elif layout_type == "AGENDA":
            items = content.get("agenda_items", [])
            if items:
                preview_text = "  ".join(f"{i+1}. {str(a)[:30]}" for i, a in enumerate(items[:3]))
        elif layout_type == "CONCLUSION":
            recs = content.get("recommendations", [])
            summ = content.get("summary_points", [])
            preview_text = " • ".join(str(r)[:35] for r in (recs or summ)[:2])
        elif layout_type == "SECTION_HEADER":
            preview_text = content.get("section_title", "") or content.get("description", "")
        else:
            bullets = content.get("bullets", [])
            if bullets:
                preview_text = " • ".join(str(b)[:40] for b in bullets[:2])
            elif content:
                first_val = next(iter(content.values()), "")
                if isinstance(first_val, str):
                    preview_text = first_val[:80]

        # Structural badge (if this is a mandatory structural slide)
        struct = _STRUCTURE_BADGES.get(layout_type)
        struct_badge_html = ""
        if struct:
            icon, bg, fg, label = struct
            struct_badge_html = (
                f'<span style="display:inline-flex;align-items:center;gap:4px;'
                f'background:{bg};color:{fg};padding:2px 8px;border-radius:12px;'
                f'font-size:0.68rem;font-weight:700;margin-right:6px;">'
                f'{icon} {label}</span>'
            )

        st.markdown(f"""
        <div class="slide-preview">
            <span class="slide-number">#{slide.slide_number:02d}</span>
            <div style="flex:1; min-width:0;">
                <div style="color:var(--text-primary); font-weight:600; font-size:0.92rem;
                            white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                    {slide.title}
                </div>
                {f'<div style="color:var(--text-muted); font-size:0.8rem; margin-top:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{preview_text}</div>' if preview_text else ""}
            </div>
            {struct_badge_html}
            <span style="font-size:0.7rem;font-weight:600;padding:3px 10px;border-radius:20px;
                         background:var(--blue-subtle);color:var(--navy-primary);
                         border:1px solid var(--border-medium);white-space:nowrap;">{layout_type}</span>
        </div>
        """, unsafe_allow_html=True)


def render_download(pptx_bytes: bytes, filename: str) -> None:
    """
    Render a corporate download banner for the generated PPTX matching the reference theme.
    """
    import base64

    # Defensive type checks
    if isinstance(pptx_bytes, (bytearray, memoryview)):
        pptx_bytes = bytes(pptx_bytes)
    elif not isinstance(pptx_bytes, bytes):
        st.error("❌ Internal error: presentation bytes are corrupted. Please regenerate.")
        return

    # Verify it's a valid ZIP/PPTX (PK magic bytes)
    if len(pptx_bytes) < 4 or pptx_bytes[:4] != b'PK\x03\x04':
        st.error("❌ The generated file does not appear to be a valid PPTX. Please regenerate.")
        return

    # Encode to base64 for inline data URI
    b64 = base64.b64encode(pptx_bytes).decode("utf-8")
    mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    href = f"data:{mime};base64,{b64}"

    # Executive deep navy corporate banner with gold CTA button
    download_html = f"""
    <div class="download-banner">
        <div class="download-banner-left">
            <div class="download-banner-title">Presentation Generated Successfully</div>
            <div class="download-banner-subtitle">
                File: <strong>{filename}</strong> ({len(pptx_bytes):,} bytes) &bull; 100% editable slides ready for presentation in PowerPoint, Google Slides, or LibreOffice.
            </div>
        </div>
        <div>
            <a href="{href}" download="{filename}" class="download-gold-cta">
                ⬇️ Download Presentation (.pptx)
            </a>
        </div>
    </div>
    """
    st.markdown(download_html, unsafe_allow_html=True)





# ---------------------------------------------------------------------------
# Generation pipeline
# ---------------------------------------------------------------------------

def run_generation_pipeline(
    source_text: str,
    input_mode: str,
    api_config: dict,
    pres_config: dict,
    filename: Optional[str] = None,
    template_id: str = "dark_navy",
) -> None:
    """
    Full generation pipeline.
    - Document mode  → semantic chunking pipeline
    - Prompt mode    → legacy single-pass analysis
    Results stored in st.session_state.
    """
    st.session_state.generation_complete = False
    st.session_state.pptx_bytes = None
    st.session_state.pptx_filename = None
    st.session_state.error_message = None
    st.session_state.chunk_analyses = None
    st.session_state.semantic_chunks = None
    st.session_state.analysis_cache = None
    st.session_state.governance_report = None

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
    requested_slides = pres_config.get("slide_count")  # None = auto
    use_semantic = (input_mode == "document")

    with st.status("🚀 Generating your presentation...", expanded=True) as status:

        # ─ Step 1: Key Manager ────────────────────────────────────────
        st.write("🔑 Initializing API keys...")
        try:
            if config.LLM_PROVIDER == "watsonx":
                key_manager = KeyManager(
                    keys=api_config["api_keys"],
                    model=api_config["model"],
                    temperature=config.WATSONX_TEMPERATURE,
                    max_tokens=config.WATSONX_MAX_TOKENS_PLAN,
                    max_retries=config.LLM_MAX_RETRIES,
                    provider="watsonx",
                    project_id=config.WATSONX_PROJECT_ID,
                    url=config.WATSONX_URL,
                    fallback_model=config.WATSONX_FALLBACK_MODEL_ID,
                )
            else:
                key_manager = KeyManager(
                    keys=api_config["api_keys"],
                    model=api_config["model"],
                    temperature=config.GROQ_TEMPERATURE,
                    max_tokens=config.GROQ_MAX_TOKENS_PLAN,
                    max_retries=config.LLM_MAX_RETRIES,
                    provider="groq",
                )
            st.caption(f"✅ {key_manager.key_count} API key(s) active — round-robin per section")
        except KeyManagerError as e:
            status.update(label="❌ No API keys", state="error")
            st.error(f"❌ {e}")
            return
        except Exception as e:
            status.update(label="❌ Key init failed", state="error")
            st.error(f"❌ {e}")
            return

        # ─ Step 2: Layout Manager ─────────────────────────────────────
        st.write("📐 Loading layout system...")
        try:
            layout_manager = LayoutManager(
                template_path=config.TEMPLATE_FILE,
                metadata_path=config.LAYOUT_METADATA_FILE,
            )
        except LayoutManagerError as e:
            status.update(label="❌ Layout system error", state="error")
            st.error(f"❌ {e}")
            return

        # ─ Step 3: Content Analysis ───────────────────────────────────
        analysis = None
        cache = None
        chunks = None

        if use_semantic:
            # Semantic pipeline (document upload)
            st.write("🏗️ Running semantic section-by-section analysis...")
            doc_root = st.session_state.get("doc_root")
            doc_filename = filename or st.session_state.get("uploaded_filename") or "document.txt"

            progress_area = st.empty()
            section_log: List[str] = []

            def _progress(heading: str, current: int, total: int) -> None:
                section_log.append(f"✓ {heading[:55]}")
                progress_area.markdown(
                    f"**Analyzing sections ({current}/{total}):**\n\n```\n" +
                    "\n".join(section_log[-8:]) + "\n```"
                )

            try:
                analysis, cache, chunks = analyze_content_semantic(
                    root_section=doc_root,
                    plain_text=source_text,
                    plain_text_filename=doc_filename,
                    key_manager=key_manager,
                    logs_dir=config.LOGS_DIR,
                    run_id=run_id,
                    audience=pres_config.get("audience", "General"),
                    style=pres_config.get("style", "Professional"),
                    slide_count=requested_slides,
                    language=pres_config.get("language", "English"),
                    additional_instructions=pres_config.get("additional_instructions", ""),
                    max_chunk_chars=config.LLM_MAX_CHUNK_CHARS,
                    min_section_chars=config.LLM_MIN_SECTION_CHARS,
                    progress_callback=_progress,
                )
                progress_area.empty()
                n_chunks = len(chunks) if chunks else 0
                st.write(f"✅ Analyzed **{n_chunks}** semantic sections")
                st.session_state.analysis_cache = cache
                st.session_state.semantic_chunks = chunks
            except Exception as e:
                status.update(label="❌ Semantic analysis failed", state="error")
                st.error(f"❌ {e}")
                logger.error("Semantic analysis failed: %s", traceback.format_exc())
                return

        else:
            # Legacy single-pass (prompt mode)
            st.write("🔍 Analyzing content with AI...")
            single_client = key_manager.get_client(chunk_index=0)
            truncated_source = truncate_text(source_text, config.MAX_SOURCE_TEXT_CHARS)
            try:
                analysis = analyze_content(
                    source_text=truncated_source,
                    client=single_client,
                    audience=pres_config.get("audience", "General"),
                    style=pres_config.get("style", "Professional"),
                    slide_count=requested_slides,
                    language=pres_config.get("language", "English"),
                    additional_instructions=pres_config.get("additional_instructions", ""),
                    max_source_chars=config.MAX_SOURCE_TEXT_CHARS,
                )
            except ContentAnalysisError as e:
                status.update(label="❌ Content analysis failed", state="error")
                st.error(f"❌ {e}")
                return
            except (GroqRateLimitError, GroqAPIError, WatsonxRateLimitError, WatsonxAPIError) as e:
                status.update(label="❌ API error", state="error")
                st.error(f"❌ LLM API error: {e}")
                return

        if requested_slides is None:
            requested_slides = analysis.suggested_slide_count
            st.caption(f"🤖 AI suggests {requested_slides} content slides (+ 4 mandatory structural slides)")

        # Create a normalized source record before planning so every downstream
        # decision can be tied back to the originating upload or prompt.
        source = ContentSource(
            source_id=run_id,
            source_type="document" if use_semantic else "prompt",
            filename=filename or "prompt",
            title=getattr(analysis, "main_topic", ""),
            text=source_text,
            provenance=[{"source_id": run_id, "filename": filename or "prompt"}],
            extraction_warnings=([] if source_text.strip() else ["Source text is empty"]),
        )
        brief = classify_presentation(
            content_analysis=analysis,
            audience=pres_config.get("audience", "General"),
            template_id=template_id,
        )
        st.caption(
            f"🧭 Classified as **{brief.presentation_type}** for **{brief.audience}** audience"
        )
        intelligence_instructions = (
            f"Presentation intelligence classification: {brief.presentation_type}. "
            f"Recommended storyline: {'; '.join(brief.storyline[:8])}. "
            f"Executive message: {brief.executive_message or 'Derive from source content'}."
        )
        planner_instructions = " ".join(
            filter(None, [pres_config.get("additional_instructions", ""), intelligence_instructions])
        )

        # ─ Step 4: Presentation Planning ────────────────────────────────
        st.write(f"🗂️ Planning presentation ({requested_slides} content slides + 4 mandatory structural slides)...")
        plan_client = key_manager.get_client()  # next key in rotation
        try:
            if template_id == "techm_v3":
                from core.presentation_planner_v3 import plan_dynamic_presentation
                plan = plan_dynamic_presentation(
                    client=plan_client,
                    content_analysis=analysis,
                    source_text=truncate_text(source_text, 8000),
                    presentation_title=pres_config.get("presentation_title", ""),
                    audience=pres_config.get("audience", "Executive Leadership"),
                    style=pres_config.get("style", "Corporate Strategic"),
                    slide_count=requested_slides,
                    language=pres_config.get("language", "English"),
                    additional_instructions=planner_instructions,
                    max_source_chars=8000,
                )
            elif template_id == "template1":
                from core.presentation_planner_template1 import plan_template1_presentation
                plan = plan_template1_presentation(
                    client=plan_client,
                    content_analysis=analysis,
                    source_text=truncate_text(source_text, 12000),
                    presentation_title=pres_config.get("presentation_title", ""),
                    audience=pres_config.get("audience", "Executive Leadership"),
                    style=pres_config.get("style", "Corporate Strategic"),
                    slide_count=requested_slides,
                    language=pres_config.get("language", "English"),
                    additional_instructions=planner_instructions,
                    max_source_chars=12000,
                )
            elif template_id == "hld_qbr":
                from core.content_model_extractor import extract_content_model
                from core.presentation_planner_hld_qbr import plan_hld_qbr_presentation
                hld_content_model = None
                # For small decks (e.g. <=4 slides), bypass extra content extraction call to conserve token rate limit
                if requested_slides and requested_slides > 4:
                    st.write("🧩 Extracting content model from source...")
                    hld_content_model = extract_content_model(plan_client, truncate_text(source_text, 8000))
                    st.session_state.hld_content_model = hld_content_model
                    if hld_content_model.content_items:
                        st.caption(f"📚 Extracted {len(hld_content_model.content_items)} traceable content item(s) from source")

                source_char_limit = 6000 if (requested_slides and requested_slides <= 4) else 14000
                plan = plan_hld_qbr_presentation(
                    client=plan_client,
                    content_analysis=analysis,
                    source_text=truncate_text(source_text, source_char_limit),
                    presentation_title=pres_config.get("presentation_title", ""),
                    facility_name=pres_config.get("facility_name", ""),
                    audience=pres_config.get("audience", "Executive Leadership"),
                    style=pres_config.get("style", "Corporate Strategic"),
                    language=pres_config.get("language", "English"),
                    additional_instructions=planner_instructions,
                    slide_count=requested_slides,
                    content_model=hld_content_model,
                    max_source_chars=source_char_limit,
                )
            else:
                plan = plan_presentation(
                    content_analysis=analysis,
                    source_text=truncate_text(source_text, 8000),
                    client=plan_client,
                    presentation_title=pres_config.get("presentation_title", ""),
                    audience=pres_config.get("audience", "General"),
                    style=pres_config.get("style", "Professional"),
                    slide_count=requested_slides,
                    language=pres_config.get("language", "English"),
                    additional_instructions=planner_instructions,
                    max_source_chars=8000,
                )
        except Exception as e:
            status.update(label="❌ Planning failed", state="error")
            st.error(f"❌ {e}")
            logger.error("Planning failed: %s", e, exc_info=True)
            return

        # Save plan log
        try:
            import json as _json
            plan_path = config.LOGS_DIR / run_id / "presentation_plan.json"
            plan_path.parent.mkdir(parents=True, exist_ok=True)
            plan_path.write_text(plan.model_dump_json(indent=2), encoding="utf-8")
        except Exception:
            pass

        # ─ Step 5: Validate Plan ─────────────────────────────────────────
        st.write("✅ Validating presentation structure...")
        plan_governance = validate_plan_governance(
            plan, template_id, source.source_id and [source.source_id],
            content_model=st.session_state.get("hld_content_model") if template_id == "hld_qbr" else None,
        )
        st.session_state.governance_report = plan_governance
        if plan_governance.errors:
            for error in plan_governance.errors:
                st.error(f"❌ Governance: {error}")
        if plan_governance.warnings:
            with st.expander(f"🛡️ {len(plan_governance.warnings)} governance notice(s)"):
                for warning in plan_governance.warnings:
                    st.caption(f"• {warning}")
        if template_id not in ("techm_v3", "template1", "hld_qbr"):
            validation = validate_plan(plan, layout_manager, requested_slides)
            if validation.warnings:
                with st.expander(f"ℹ️ {len(validation.warnings)} validation notice(s)"):
                    for warn in validation.warnings:
                        st.caption(f"• {warn}")
        elif template_id == "hld_qbr":
            st.caption("✨ HLD QBR archetype-based layout active (sections included only where source content supports them)")
        else:
            tmpl_name = "TechM V3" if template_id == "techm_v3" else "Template-1 (Weekly Update)"
            st.caption(f"✨ {tmpl_name} Dynamic Layout Active ({len(plan.slides)} content slides composed)")

        # ─ Step 6: Build PowerPoint ───────────────────────────────────────
        st.write("🎨 Building PowerPoint presentation...")
        try:
            import importlib
            import core.builders.hld_qbr_builder
            importlib.reload(core.builders.hld_qbr_builder)
            import core.template_registry
            importlib.reload(core.template_registry)
            from core.template_registry import get_builder
            engine = get_builder(template_id, layout_manager)
            pptx_bytes = engine.build(plan)
        except Exception as e:
            status.update(label="❌ PPTX build failed", state="error")
            st.error(f"❌ {e}")
            logger.error("PPTX build failed: %s", e, exc_info=True)
            return

        # ─ Step 7: Final validation ──────────────────────────────────────
        st.write("🔎 Validating output file...")
        file_validation = validate_pptx_bytes(pptx_bytes)
        if file_validation.errors:
            status.update(label="❌ File validation failed", state="error")
            for err in file_validation.errors:
                st.error(f"❌ {err}")
            return

        output_governance = validate_pptx_governance(pptx_bytes, template_id)
        if output_governance.warnings:
            with st.expander(f"🔎 {len(output_governance.warnings)} output governance notice(s)"):
                for warning in output_governance.warnings:
                    st.caption(f"• {warning}")
        if output_governance.errors:
            status.update(label="❌ Governance validation failed", state="error")
            for error in output_governance.errors:
                st.error(f"❌ Governance: {error}")
            return
        st.session_state.governance_report = output_governance
        st.caption(
            f"🛡️ Governance score: **{output_governance.quality_score}/100** | "
            f"Brand: {'pass' if output_governance.brand_compliance else 'review required'} | "
            f"Accessibility: {'pass' if output_governance.accessibility else 'review required'}"
        )

        try:
            write_audit_record(
                config.LOGS_DIR / run_id / "governance_audit.json",
                source=source,
                brief=brief,
                plan=plan,
                plan_report=plan_governance,
                output_report=output_governance,
            )
        except Exception as audit_error:
            logger.warning("Could not write governance audit record: %s", audit_error)

        # HLDQBRPresentationPlan uses `presentation_title` and has no `slides` list
        plan_title = getattr(plan, "title", None) or getattr(plan, "presentation_title", "presentation")
        pptx_filename = sanitize_filename(plan_title)
        if hasattr(plan, "slides"):
            slide_count = len(plan.slides)
        else:
            from pptx import Presentation
            slide_count = len(Presentation(io.BytesIO(pptx_bytes)).slides)

        # --- Persist to disk (backup copy) so file is available even after Streamlit re-run ---
        try:
            out_path = config.OUTPUT_DIR / pptx_filename
            out_path.write_bytes(pptx_bytes)
            logger.info("PPTX saved to disk: %s (%d bytes)", out_path, len(pptx_bytes))
        except Exception as disk_err:
            logger.warning("Could not save PPTX to disk: %s", disk_err)

        st.write(f"✅ **{pptx_filename}** is ready!")
        st.session_state.pptx_bytes = bytes(pptx_bytes)   # ensure raw bytes in session state
        st.session_state.pptx_filename = pptx_filename
        st.session_state.presentation_plan = plan
        st.session_state.generation_complete = True
        status.update(label=f"✅ Done: {slide_count} slides", state="complete")
        cleanup_old_outputs(config.OUTPUT_DIR)



# ---------------------------------------------------------------------------
# Debug view — semantic analysis results
# ---------------------------------------------------------------------------

def render_semantic_debug(cache, chunks) -> None:
    """Show per-chunk analysis JSON and consolidation in expandable sections."""
    if not cache or not chunks:
        return

    st.markdown("---")
    st.markdown("### 🔬 Semantic Analysis Debug")

    # Per-chunk analyses
    analyses = cache.get_all()
    if analyses:
        st.markdown(f"**{len(analyses)} section(s) analyzed:**")
        for analysis in analyses:
            path_label = " → ".join(analysis.section_path) if analysis.section_path else analysis.heading
            importance_icon = {"high": "🔴", "medium": "🟡", "low": "⚪"}.get(analysis.importance, "🟡")
            with st.expander(
                f"{importance_icon} {path_label} [{analysis.chunk_id}]",
                expanded=False,
            ):
                col_a, col_b = st.columns(2)
                with col_a:
                    st.caption(f"**Level:** {analysis.level} | **Importance:** {analysis.importance}")
                    st.caption(f"**Summary:** {analysis.summary}")
                    if analysis.key_points:
                        st.caption("**Key Points:**")
                        for p in analysis.key_points[:5]:
                            st.caption(f"  • {p}")
                with col_b:
                    if analysis.statistics:
                        st.caption(f"📊 **Statistics:** {', '.join(analysis.statistics[:3])}")
                    if analysis.processes:
                        st.caption(f"⚙️ **Processes:** {', '.join(analysis.processes[:2])}")
                    if analysis.potential_visuals:
                        st.caption(f"🖼️ **Visuals:** {', '.join(analysis.potential_visuals[:2])}")
                    if analysis.content_types:
                        st.caption(f"🏷️ **Types:** {', '.join(analysis.content_types)}")
                st.json(analysis.model_dump())


# ---------------------------------------------------------------------------
# Main app layout
# ---------------------------------------------------------------------------

def main() -> None:
    # Sidebar (temporarily commented out)
    # api_config = render_sidebar()
    if config.LLM_PROVIDER == "watsonx":
        api_config = {
            "api_keys": config.WATSONX_API_KEYS,
            "model": config.WATSONX_MODEL_ID,
        }
    else:
        api_config = {
            "api_keys": getattr(config, "get_groq_api_keys", lambda: config.GROQ_API_KEYS)(),
            "model": config.GROQ_MODEL,
        }

    # Hero header
    render_header()

    tab_generate, tab_qa = st.tabs(["🎨 Generate Presentation", "💬 Ask watsonx.ai"])

    with tab_generate:
        render_generator_tab(api_config)

    with tab_qa:
        render_qa_tab(api_config)


def render_generator_tab(api_config: dict) -> None:
    # Input section — returns extended tuple now
    source_text, input_mode, file_bytes, filename, doc_root, detection_method = render_input_section()

    st.markdown("---")

    # Configuration
    pres_config = render_configuration()

    st.markdown("---")

    # Generate button
    col_btn, col_info = st.columns([2, 3])
    with col_btn:
        has_keys = bool(api_config.get("api_keys"))
        can_generate = bool(source_text) and has_keys
        generate_clicked = st.button(
            "🚀 Generate Presentation",
            type="primary",
            disabled=not can_generate,
            use_container_width=True,
        )

    with col_info:
        if not has_keys:
            env_var = "WATSONX_API_KEY" if config.LLM_PROVIDER == "watsonx" else "GROQ_API_KEYS"
            st.info(f"👈 No API keys configured. Add keys in the sidebar or set {env_var} in .env")
        elif not source_text:
            st.info("📄 Upload a document or enter a prompt above to get started.")
        else:
            slide_display = pres_config.get("slide_count") or "Auto"
            mode_label = "Document (Semantic)" if input_mode == "document" else "Prompt (Single-pass)"
            st.success(
                f"✅ Ready to generate! "
                f"Mode: **{mode_label}** | "
                f"Slides: **{slide_display}** | "
                f"Audience: **{pres_config.get('audience', 'General')}**"
            )

    # Run pipeline
    if generate_clicked and can_generate:
        run_generation_pipeline(
            source_text=source_text,
            input_mode=input_mode or "prompt",
            api_config=api_config,
            pres_config=pres_config,
            filename=filename,
            template_id=pres_config.get("template_id", "dark_navy"),
        )

    # Presentation Preview (temporarily commented out)
    # if st.session_state.generation_complete and st.session_state.presentation_plan:
    #     render_preview(st.session_state.presentation_plan)

    # Download option
    st.markdown("---")
    st.markdown("### 📥 Download Presentation")
    if st.session_state.pptx_bytes and st.session_state.pptx_filename:
        render_download(st.session_state.pptx_bytes, st.session_state.pptx_filename)
    else:
        st.info("ℹ️ The download option will become available once the presentation is generated.")
        st.download_button(
            label="⬇️ Download Presentation (.pptx)",
            data=b"",
            file_name="presentation.pptx",
            disabled=True,
            help="Generate a presentation first to enable download.",
        )

    # Debug view (semantic pipeline only) (temporarily commented out)
    # if st.session_state.analysis_cache:
    #     render_semantic_debug(
    #         cache=st.session_state.analysis_cache,
    #         chunks=st.session_state.semantic_chunks,
    #     )


# ---------------------------------------------------------------------------
# Simple Q&A tab — direct chat against the configured LLM provider
# ---------------------------------------------------------------------------

def _build_qa_client(api_config: dict):
    """Build a single LLM client (Groq or watsonx, per config.LLM_PROVIDER) for Q&A."""
    if config.LLM_PROVIDER == "watsonx":
        key_manager = KeyManager(
            keys=api_config["api_keys"],
            model=api_config["model"],
            temperature=config.WATSONX_TEMPERATURE,
            max_tokens=config.WATSONX_MAX_TOKENS_SLIDE,
            max_retries=config.LLM_MAX_RETRIES,
            provider="watsonx",
            project_id=config.WATSONX_PROJECT_ID,
            url=config.WATSONX_URL,
            fallback_model=config.WATSONX_FALLBACK_MODEL_ID,
        )
    else:
        key_manager = KeyManager(
            keys=api_config["api_keys"],
            model=api_config["model"],
            temperature=config.GROQ_TEMPERATURE,
            max_tokens=config.GROQ_MAX_TOKENS_SLIDE,
            max_retries=config.LLM_MAX_RETRIES,
            provider="groq",
        )
    return key_manager.get_client()


def render_qa_tab(api_config: dict) -> None:
    """Simple chat-style Q&A that sends questions directly to the configured LLM provider."""
    st.markdown("### 💬 Ask watsonx.ai")
    st.caption("Ask a general question and get a direct answer — no document upload required.")

    if "qa_history" not in st.session_state:
        st.session_state.qa_history = []  # list of (question, answer) tuples

    for question, answer in st.session_state.qa_history:
        with st.chat_message("user"):
            st.markdown(question)
        with st.chat_message("assistant"):
            st.markdown(answer)

    question = st.chat_input("Type your question...")
    if not question:
        return

    if not api_config.get("api_keys"):
        st.error("❌ No API keys configured. Set the appropriate keys in .env.")
        return

    with st.chat_message("user"):
        st.markdown(question)

    with st.chat_message("assistant"):
        with st.spinner("Thinking..."):
            try:
                client = _build_qa_client(api_config)
                answer = client.chat_complete(
                    messages=[
                        {"role": "system", "content": "You are a helpful, concise assistant."},
                        {"role": "user", "content": question},
                    ],
                )
                if not answer:
                    answer = "⚠️ The model returned an empty response. Please try rephrasing your question."
            except (GroqAuthError, WatsonxAuthError) as e:
                answer = f"❌ Authentication error: {e}"
            except (GroqRateLimitError, WatsonxRateLimitError) as e:
                answer = f"⏳ Rate limit/quota exceeded: {e}"
            except (GroqAPIError, WatsonxAPIError, KeyManagerError) as e:
                answer = f"❌ API error: {e}"
            except Exception as e:
                answer = f"❌ Unexpected error: {e}"
            st.markdown(answer)

    st.session_state.qa_history.append((question, answer))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    main()
