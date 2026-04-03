"""
Governance Assistant — Web interface with document upload + chat.
"""

import os
import shutil
from pathlib import Path
from typing import Generator

import streamlit as st
from dotenv import load_dotenv
import anthropic

from indexer import GovernanceIndexer
from exporter import to_docx, to_xlsx, to_pptx, to_pdf

load_dotenv()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
_BASE = os.getenv("RAILWAY_VOLUME_MOUNT_PATH", os.path.dirname(__file__))
UPLOADS_DIR = os.path.join(_BASE, "uploads")
DB_DIR = os.path.join(_BASE, "chroma_db")

SUPPORTED_TYPES = {
    "pdf":  "📄",
    "docx": "📝",
    "xlsx": "📊",
    "xls":  "📊",
    "xlsm": "📊",
    "pptx": "📑",
    "ppt":  "📑",
}

SYSTEM_PROMPT = """You are a governance expert assistant helping users navigate organizational governance documents — including Delegation of Authority (DOA), policies, procedures, operating models, committee charters, organization charts, job descriptions, and governance frameworks.

Your responsibilities:
- Answer questions accurately using ONLY the provided document excerpts
- Always cite the source document(s) your answer is drawn from
- Highlight approval thresholds, authority limits, responsibilities, or required processes where relevant
- If the excerpts don't fully answer the question, say so and suggest what document type might contain the answer
- If a question involves a decision that exceeds documented authority, flag this explicitly
- Keep answers structured and professional"""


# ------------------------------------------------------------------ #
# Page config + CSS
# ------------------------------------------------------------------ #

st.set_page_config(
    page_title="Governance Assistant",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.main .block-container {
    padding-top: 1.5rem;
    padding-bottom: 2rem;
    max-width: 960px;
}

/* ── Header ── */
.gov-header {
    background: linear-gradient(135deg, #1B3A6B 0%, #2D5FA3 100%);
    border-radius: 12px;
    padding: 1.4rem 1.8rem;
    margin-bottom: 1.2rem;
    display: flex;
    align-items: center;
    gap: 14px;
}
.gov-header-icon { font-size: 2rem; }
.gov-header-title {
    color: white; font-size: 1.5rem; font-weight: 700;
    margin: 0; line-height: 1.2;
}
.gov-header-sub {
    color: rgba(255,255,255,0.75); font-size: 0.82rem; margin: 2px 0 0 0;
}

/* ── Upload zone ── */
[data-testid="stFileUploader"] {
    background: white;
    border: 2px dashed #CBD5E1;
    border-radius: 10px;
    padding: 0.5rem;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover { border-color: #2D5FA3; }

/* ── Document card ── */
.doc-card {
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 0.6rem 0.8rem;
    margin-bottom: 0.4rem;
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 0.85rem;
    color: #1E293B;
}
.doc-card:hover { box-shadow: 0 2px 8px rgba(0,0,0,0.07); }
.doc-name {
    flex: 1; overflow: hidden;
    text-overflow: ellipsis; white-space: nowrap; font-weight: 500;
}

/* ── Stats pill ── */
.stat-pill {
    display: inline-flex; align-items: center;
    background: #EFF6FF; color: #1B3A6B;
    border-radius: 999px; padding: 4px 12px;
    font-size: 0.78rem; font-weight: 600;
    gap: 4px; margin-right: 6px;
}

/* ── Chat messages ── */
[data-testid="stChatMessage"] {
    border-radius: 12px; margin-bottom: 0.6rem;
}
div[class*="stChatMessage"]:has([data-testid="chatAvatarIcon-user"]) {
    background: #EFF6FF;
}
div[class*="stChatMessage"]:has([data-testid="chatAvatarIcon-assistant"]) {
    background: white;
    border: 1px solid #E2E8F0;
    box-shadow: 0 1px 4px rgba(0,0,0,0.04);
}

/* ── Source badges ── */
.sources-row {
    display: flex; flex-wrap: wrap; gap: 6px;
    margin-top: 10px; padding-top: 10px;
    border-top: 1px solid #E2E8F0;
}
.source-badge {
    background: #F1F5F9; border: 1px solid #CBD5E1;
    border-radius: 6px; padding: 2px 8px;
    font-size: 0.72rem; color: #475569; font-weight: 500;
}

/* ── Chat input ── */
[data-testid="stChatInput"] {
    border-radius: 12px; border: 1.5px solid #CBD5E1;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
[data-testid="stChatInput"]:focus-within { border-color: #2D5FA3; }

/* ── Welcome card ── */
.welcome-card {
    background: white; border: 1px solid #E2E8F0;
    border-radius: 12px; padding: 2rem;
    text-align: center; color: #475569;
}
.welcome-card h3 { color: #1B3A6B; margin-bottom: 0.4rem; }

/* ── Download cards ── */
.dl-card {
    background: white; border: 1px solid #E2E8F0;
    border-radius: 12px; padding: 1.2rem 1.4rem;
    margin-bottom: 0.8rem; display: flex;
    align-items: center; gap: 1rem;
}
.dl-icon { font-size: 2rem; flex-shrink: 0; }
.dl-info h4 { margin: 0 0 2px 0; color: #1B3A6B; font-size: 0.95rem; }
.dl-info p  { margin: 0; color: #64748B; font-size: 0.8rem; }

/* ── Tabs ── */
[data-testid="stTabs"] button { font-weight: 600; font-size: 0.92rem; }

hr { border-color: #E2E8F0; margin: 0.8rem 0; }
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stSidebar"] { display: none; }
</style>
""", unsafe_allow_html=True)


# ------------------------------------------------------------------ #
# System init (cached)
# ------------------------------------------------------------------ #

def _start_watcher(indexer: GovernanceIndexer):
    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler

        class Handler(FileSystemEventHandler):
            def on_created(self, event):
                if not event.is_directory:
                    indexer.index_file(event.src_path)
            def on_modified(self, event):
                if not event.is_directory:
                    indexer.index_file(event.src_path)
            def on_deleted(self, event):
                if not event.is_directory:
                    indexer.remove_file(event.src_path)

        observer = Observer()
        observer.schedule(Handler(), indexer.folder_path, recursive=True)
        observer.daemon = True
        observer.start()
    except Exception as e:
        print(f"[Watcher] {e}")


@st.cache_resource
def get_indexer() -> GovernanceIndexer:
    os.makedirs(UPLOADS_DIR, exist_ok=True)
    indexer = GovernanceIndexer(UPLOADS_DIR, db_path=DB_DIR)
    indexer.index_all()
    _start_watcher(indexer)
    return indexer


# ------------------------------------------------------------------ #
# Answer generation
# ------------------------------------------------------------------ #

def get_sources_and_stream(indexer, query, history):
    chunks = indexer.search(query, n_results=6)
    if not chunks:
        def _empty():
            yield "I couldn't find relevant content in the uploaded documents. Please upload your governance documents first."
        return [], _empty()

    sources = list(dict.fromkeys(c["source"] for c in chunks))
    context = "\n\n---\n\n".join(
        f"[Source: {c['source']}]\n{c['text']}" for c in chunks
    )
    messages = [{"role": m["role"], "content": m["content"]} for m in history[-12:]]
    messages.append({
        "role": "user",
        "content": (
            f"Question: {query}\n\n"
            f"Relevant governance document excerpts:\n\n{context}\n\n"
            "Please answer based on the excerpts above and cite your sources."
        ),
    })

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    def _stream():
        with client.messages.stream(
            model="claude-sonnet-4-6",
            max_tokens=1500,
            system=SYSTEM_PROMPT,
            messages=messages,
        ) as stream:
            for text in stream.text_stream:
                yield text

    return sources, _stream()


# ------------------------------------------------------------------ #
# UI helpers
# ------------------------------------------------------------------ #

def file_icon(filename: str) -> str:
    ext = Path(filename).suffix.lstrip(".").lower()
    return SUPPORTED_TYPES.get(ext, "📎")


def render_doc_card(filename: str):
    st.markdown(
        f'<div class="doc-card">'
        f'<span style="font-size:1.1rem">{file_icon(filename)}</span>'
        f'<span class="doc-name" title="{filename}">{filename}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )


def render_sources(sources: list[str]):
    if not sources:
        return
    badges = "".join(
        f'<span class="source-badge">{file_icon(s)} {s}</span>' for s in sources
    )
    st.markdown(f'<div class="sources-row">{badges}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------ #
# App
# ------------------------------------------------------------------ #

indexer = get_indexer()

# Header
st.markdown("""
<div class="gov-header">
    <div class="gov-header-icon">📋</div>
    <div>
        <p class="gov-header-title">Governance Assistant</p>
        <p class="gov-header-sub">Upload your governance documents and ask anything — DOA, policies, procedures, charters, org charts, JDs</p>
    </div>
</div>
""", unsafe_allow_html=True)

if "messages" not in st.session_state:
    st.session_state.messages = []

# ---- 3 Tabs ----
tab_upload, tab_chat, tab_download = st.tabs(["📂  Documents", "💬  Chat", "📥  Download"])


# ================================================================== #
# TAB 1 — UPLOAD & MANAGE DOCUMENTS
# ================================================================== #
with tab_upload:
    st.markdown("#### Upload Documents")
    st.caption("Supported: PDF, Word (.docx), PowerPoint (.pptx), Excel (.xlsx / .xls)")

    uploaded_files = st.file_uploader(
        "Drag and drop files here, or click to browse",
        type=["pdf", "docx", "pptx", "xlsx", "xls", "xlsm"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded_files:
        new_count = 0
        for uf in uploaded_files:
            dest = os.path.join(UPLOADS_DIR, uf.name)
            with open(dest, "wb") as f:
                f.write(uf.getbuffer())
            if indexer.index_file(dest):
                new_count += 1
        if new_count:
            st.success(f"{new_count} document(s) indexed successfully. Go to the Chat tab to ask questions.")
            st.rerun()
        else:
            st.info("These documents are already indexed (no changes detected).")

    st.markdown("<hr>", unsafe_allow_html=True)

    # Document list
    indexed_files = indexer.get_indexed_files()
    chunk_count = indexer.get_chunk_count()

    st.markdown("#### Indexed Documents")

    if indexed_files:
        st.markdown(
            f'<span class="stat-pill">📚 {len(indexed_files)} documents</span>'
            f'<span class="stat-pill">🔍 {chunk_count} searchable chunks</span>',
            unsafe_allow_html=True,
        )
        st.markdown("")

        for fname in indexed_files:
            col_name, col_del = st.columns([6, 1])
            with col_name:
                render_doc_card(fname)
            with col_del:
                fpath = os.path.join(UPLOADS_DIR, fname)
                if st.button("Remove", key=f"del_{fname}", use_container_width=True):
                    if os.path.exists(fpath):
                        os.remove(fpath)
                    indexer.remove_file(fpath)
                    st.rerun()

        st.markdown("<hr>", unsafe_allow_html=True)
        if st.button("🗑️ Remove All Documents", use_container_width=False):
            shutil.rmtree(UPLOADS_DIR, ignore_errors=True)
            os.makedirs(UPLOADS_DIR, exist_ok=True)
            get_indexer.clear()
            st.session_state.messages = []
            st.rerun()
    else:
        st.markdown("""
        <div class="welcome-card">
            <div style="font-size:2.5rem;margin-bottom:0.6rem">📂</div>
            <h3>No documents uploaded yet</h3>
            <p style="font-size:0.85rem;margin-top:0.3rem">
                Use the upload area above to add your governance documents.
            </p>
        </div>
        """, unsafe_allow_html=True)


# ================================================================== #
# TAB 2 — CHAT
# ================================================================== #
with tab_chat:
    indexed_files = indexer.get_indexed_files()

    if not indexed_files:
        st.markdown("""
        <div class="welcome-card">
            <div style="font-size:2.5rem;margin-bottom:0.6rem">📂</div>
            <h3>Upload documents first</h3>
            <p style="font-size:0.85rem;margin-top:0.3rem">
                Go to the <strong>Documents</strong> tab to upload your governance files.
            </p>
        </div>
        """, unsafe_allow_html=True)
    else:
        # Scrollable message area
        chat_area = st.container(height=480)
        with chat_area:
            if not st.session_state.messages:
                st.markdown("""
                <div class="welcome-card">
                    <div style="font-size:2rem;margin-bottom:0.5rem">💬</div>
                    <h3>Ready to answer your questions</h3>
                    <p style="font-size:0.85rem;margin-top:0.3rem">Type below or try a suggestion:</p>
                </div>
                """, unsafe_allow_html=True)
            else:
                for msg in st.session_state.messages:
                    with st.chat_message(msg["role"]):
                        st.markdown(msg["content"])
                        if msg["role"] == "assistant" and msg.get("sources"):
                            render_sources(msg["sources"])

        # Suggestion buttons (only when no messages yet)
        if not st.session_state.messages:
            cols = st.columns(2)
            suggestions = [
                "What are the approval limits in the DOA?",
                "What is the procurement process for high-value contracts?",
                "Who has authority to approve variation orders?",
                "What are the vendor registration requirements?",
            ]
            for i, s in enumerate(suggestions):
                if cols[i % 2].button(s, use_container_width=True, key=f"sug_{i}"):
                    st.session_state.messages.append({"role": "user", "content": s, "sources": []})
                    st.rerun()
        else:
            if st.button("🗨️ Clear Chat", key="clear_chat"):
                st.session_state.messages = []
                st.rerun()

        # Input row — always visible at bottom of tab
        col_input, col_send = st.columns([6, 1])
        with col_input:
            prompt = st.text_input(
                "Ask a question",
                placeholder="Ask about governance, DOA, policies, procedures…",
                label_visibility="collapsed",
                key="chat_input",
            )
        with col_send:
            send = st.button("Send", use_container_width=True, type="primary")

        if send and prompt:
            st.session_state.messages.append({"role": "user", "content": prompt, "sources": []})
            sources, stream = get_sources_and_stream(
                indexer, prompt, st.session_state.messages[:-1]
            )
            response = ""
            with chat_area:
                with st.chat_message("user"):
                    st.markdown(prompt)
                with st.chat_message("assistant"):
                    response = st.write_stream(stream)
                    render_sources(sources)

            st.session_state.messages.append({
                "role": "assistant",
                "content": response,
                "sources": sources,
            })
            st.rerun()


# ================================================================== #
# TAB 3 — DOWNLOAD
# ================================================================== #
with tab_download:
    msgs = st.session_state.messages
    has_chat = any(m["role"] == "assistant" for m in msgs)

    if not has_chat:
        st.markdown("""
        <div class="welcome-card">
            <div style="font-size:2rem;margin-bottom:0.5rem">📥</div>
            <h3>No conversation to export yet</h3>
            <p style="font-size:0.85rem;margin-top:0.3rem">
                Go to the <strong>Chat</strong> tab, ask some questions, then come back here to download.
            </p>
        </div>
        """, unsafe_allow_html=True)
    else:
        from datetime import datetime
        n_qa = sum(1 for m in msgs if m["role"] == "assistant")
        ts = datetime.now().strftime("%Y%m%d_%H%M")

        st.markdown(
            f'<p style="color:#64748B;font-size:0.88rem;margin-bottom:1.2rem">'
            f'Your conversation has <strong>{n_qa} Q&A pair{"s" if n_qa > 1 else ""}</strong>. '
            f'Choose a format to download.</p>',
            unsafe_allow_html=True,
        )

        formats = [
            ("📝", "Word Document (.docx)",
             "Formatted report with questions, answers and sources. Best for sharing and printing.",
             "word", f"governance_export_{ts}.docx",
             "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
             to_docx),
            ("📊", "Excel Spreadsheet (.xlsx)",
             "Tabular layout with one row per Q&A. Best for analysis and record-keeping.",
             "excel", f"governance_export_{ts}.xlsx",
             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
             to_xlsx),
            ("📑", "PowerPoint Presentation (.pptx)",
             "One slide per Q&A with branded design. Best for presenting findings.",
             "ppt", f"governance_export_{ts}.pptx",
             "application/vnd.openxmlformats-officedocument.presentationml.presentation",
             to_pptx),
            ("📄", "PDF Document (.pdf)",
             "Print-ready document with all Q&A and sources. Best for archiving.",
             "pdf", f"governance_export_{ts}.pdf",
             "application/pdf",
             to_pdf),
        ]

        for icon, title, desc, key, fname, mime, fn in formats:
            col_info, col_btn = st.columns([5, 1])
            with col_info:
                st.markdown(
                    f'<div class="dl-card">'
                    f'<div class="dl-icon">{icon}</div>'
                    f'<div class="dl-info"><h4>{title}</h4><p>{desc}</p></div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
            with col_btn:
                st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
                try:
                    data = fn(msgs)
                    st.download_button(
                        label="Download",
                        data=data,
                        file_name=fname,
                        mime=mime,
                        key=f"dl_{key}",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Error: {e}")
