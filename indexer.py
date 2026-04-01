"""
Document indexer for governance chatbot.
Handles PDF, DOCX, PPTX, XLSX — tracks file changes to avoid redundant re-indexing.
"""

import os
import hashlib
import json
from pathlib import Path
from typing import Optional
import chromadb
from fastembed import TextEmbedding

class _FastEmbedFunction:
    """Lightweight ONNX-based embedding function — no PyTorch required."""
    def __init__(self, model_name: str = "BAAI/bge-small-en-v1.5"):
        self._model = TextEmbedding(model_name)

    def __call__(self, input: list[str]) -> list[list[float]]:
        return [e.tolist() for e in self._model.embed(input)]

import pypdf
from docx import Document as DocxDocument
from pptx import Presentation
import pandas as pd

SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".pptx", ".xlsx", ".xls"}

# Each chunk is ~400 words with 80-word overlap to preserve context across boundaries
CHUNK_WORDS = 400
CHUNK_OVERLAP_WORDS = 80


class GovernanceIndexer:
    def __init__(self, folder_path: str, db_path: str = "./chroma_db"):
        self.folder_path = str(folder_path)

        ef = _FastEmbedFunction()
        self.client = chromadb.PersistentClient(path=db_path)
        self.collection = self.client.get_or_create_collection(
            name="governance_docs",
            embedding_function=ef,
            metadata={"hnsw:space": "cosine"},
        )

        self._hash_store_path = os.path.join(db_path, "file_hashes.json")
        self._file_hashes = self._load_hashes()

    # ------------------------------------------------------------------ #
    # Hash tracking (skip unchanged files on restart)
    # ------------------------------------------------------------------ #

    def _load_hashes(self) -> dict:
        if os.path.exists(self._hash_store_path):
            with open(self._hash_store_path) as f:
                return json.load(f)
        return {}

    def _save_hashes(self):
        os.makedirs(os.path.dirname(self._hash_store_path), exist_ok=True)
        with open(self._hash_store_path, "w") as f:
            json.dump(self._file_hashes, f)

    def _file_hash(self, path: str) -> str:
        with open(path, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()

    # ------------------------------------------------------------------ #
    # Document loaders
    # ------------------------------------------------------------------ #

    def _load_pdf(self, path: str) -> str:
        reader = pypdf.PdfReader(path)
        return "\n".join(page.extract_text() or "" for page in reader.pages)

    def _load_docx(self, path: str) -> str:
        doc = DocxDocument(path)
        return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

    def _load_pptx(self, path: str) -> str:
        prs = Presentation(path)
        parts = []
        for i, slide in enumerate(prs.slides, 1):
            texts = [
                shape.text.strip()
                for shape in slide.shapes
                if hasattr(shape, "text") and shape.text.strip()
            ]
            if texts:
                parts.append(f"[Slide {i}]\n" + "\n".join(texts))
        return "\n\n".join(parts)

    def _load_xlsx(self, path: str) -> str:
        engine = "xlrd" if path.endswith(".xls") else "openpyxl"
        xl = pd.ExcelFile(path, engine=engine)
        parts = []
        for sheet in xl.sheet_names:
            df = xl.parse(sheet)
            parts.append(f"[Sheet: {sheet}]\n{df.to_string(index=False, max_rows=300)}")
        return "\n\n".join(parts)

    def _load_document(self, path: str) -> Optional[str]:
        ext = Path(path).suffix.lower()
        try:
            if ext == ".pdf":
                return self._load_pdf(path)
            elif ext == ".docx":
                return self._load_docx(path)
            elif ext in (".pptx", ".ppt"):
                return self._load_pptx(path)
            elif ext in (".xlsx", ".xls"):
                return self._load_xlsx(path)
        except Exception as e:
            print(f"[Indexer] Error loading {Path(path).name}: {e}")
        return None

    # ------------------------------------------------------------------ #
    # Chunking
    # ------------------------------------------------------------------ #

    def _chunk_text(self, text: str) -> list[str]:
        words = text.split()
        step = CHUNK_WORDS - CHUNK_OVERLAP_WORDS
        chunks = []
        for i in range(0, len(words), step):
            chunk = " ".join(words[i : i + CHUNK_WORDS])
            if len(chunk.strip()) > 50:
                chunks.append(chunk)
        return chunks

    # ------------------------------------------------------------------ #
    # Index operations
    # ------------------------------------------------------------------ #

    def _remove_file_chunks(self, path: str):
        try:
            results = self.collection.get(where={"path": path})
            if results and results["ids"]:
                self.collection.delete(ids=results["ids"])
        except Exception as e:
            print(f"[Indexer] Error removing chunks for {Path(path).name}: {e}")

    def index_file(self, path: str) -> bool:
        """Index a single file. Returns True if newly indexed/updated."""
        if Path(path).suffix.lower() not in SUPPORTED_EXTENSIONS:
            return False
        if not os.path.exists(path):
            return False

        current_hash = self._file_hash(path)
        if self._file_hashes.get(path) == current_hash:
            return False  # file unchanged

        self._remove_file_chunks(path)

        text = self._load_document(path)
        if not text or not text.strip():
            print(f"[Indexer] No text extracted from {Path(path).name}")
            return False

        chunks = self._chunk_text(text)
        if not chunks:
            return False

        filename = Path(path).name
        ids = [f"{path}::chunk_{i}" for i in range(len(chunks))]
        metadatas = [
            {"source": filename, "path": path, "chunk": i}
            for i in range(len(chunks))
        ]

        self.collection.add(documents=chunks, ids=ids, metadatas=metadatas)

        self._file_hashes[path] = current_hash
        self._save_hashes()
        print(f"[Indexer] Indexed '{filename}' → {len(chunks)} chunks")
        return True

    def remove_file(self, path: str):
        """Remove a deleted file from the index."""
        self._remove_file_chunks(path)
        self._file_hashes.pop(path, None)
        self._save_hashes()
        print(f"[Indexer] Removed '{Path(path).name}' from index")

    def index_all(self):
        """Scan the governance folder recursively and index all new/changed files."""
        folder = Path(self.folder_path)
        if not folder.exists():
            print(f"[Indexer] Folder not found: {self.folder_path}")
            return

        files = [f for f in folder.rglob("*") if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS]
        print(f"[Indexer] Scanning {len(files)} document(s)…")

        updated = sum(self.index_file(str(f)) for f in files)

        # Clean stale hashes for deleted files
        current_paths = {str(f) for f in files}
        stale = [p for p in list(self._file_hashes) if p not in current_paths]
        for p in stale:
            self._remove_file_chunks(p)
            del self._file_hashes[p]
        if stale:
            self._save_hashes()

        print(f"[Indexer] Done — {updated} new/updated, {len(stale)} removed.")

    def force_reindex_all(self):
        """Clear hash cache and re-index everything from scratch."""
        self._file_hashes.clear()
        self._save_hashes()
        self.index_all()

    # ------------------------------------------------------------------ #
    # Search
    # ------------------------------------------------------------------ #

    def search(self, query: str, n_results: int = 6) -> list[dict]:
        """Return top-N relevant chunks for a query."""
        count = self.collection.count()
        if count == 0:
            return []

        results = self.collection.query(
            query_texts=[query],
            n_results=min(n_results, count),
        )

        return [
            {"text": doc, "source": meta["source"], "path": meta["path"]}
            for doc, meta in zip(
                results["documents"][0], results["metadatas"][0]
            )
        ]

    # ------------------------------------------------------------------ #
    # Status helpers
    # ------------------------------------------------------------------ #

    def get_indexed_files(self) -> list[str]:
        return sorted(Path(p).name for p in self._file_hashes)

    def get_chunk_count(self) -> int:
        return self.collection.count()
