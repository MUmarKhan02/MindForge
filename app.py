import os
import json
import tempfile
import requests
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from flask import Flask, request, jsonify, render_template, send_file

OLLAMA_URL = "http://localhost:11434"
OLLAMA_MODEL = "llama3.2"

def ollama_chat(messages: list, system: str = "") -> str:
    payload = {
        "model": OLLAMA_MODEL,
        "messages": ([{"role": "system", "content": system}] if system else []) + messages,
        "stream": False,
        "options": {"temperature": 0.1, "num_ctx": 8192}
    }
    try:
        res = requests.post(f"{OLLAMA_URL}/api/chat", json=payload, timeout=300)
        res.raise_for_status()
        return res.json()["message"]["content"]
    except requests.exceptions.ConnectionError:
        raise RuntimeError("Cannot connect to Ollama. Make sure it is running — run: ollama serve")
    except Exception as e:
        raise RuntimeError(f"Ollama error: {e}")

def check_ollama() -> dict:
    try:
        res = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5)
        models = [m["name"].split(":")[0] for m in res.json().get("models", [])]
        return {"running": True, "model_ready": OLLAMA_MODEL in models, "models": models}
    except Exception:
        return {"running": False, "model_ready": False, "models": []}

def safe_json(raw: str) -> dict:
    """Parse JSON from LLM output, fixing invalid LaTeX backslash escapes."""
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[1].rsplit("```", 1)[0].strip()
    start = raw.find("{")
    if start == -1:
        raise ValueError("No JSON object found in response")
    end = raw.rfind("}") + 1
    raw = raw[start:end] if end > start else raw[start:]

    # Walk the raw string and fix invalid JSON escape sequences.
    # LLMs writing LaTeX produce things like \sum, \frac, \left which are
    # not valid JSON escapes and cause json.loads to fail.
    VALID = set('"\\/bfnrtu')
    out = []
    i, n = 0, len(raw)
    in_str = False
    while i < n:
        c = raw[i]
        if in_str:
            if c == "\\":
                nxt = raw[i+1] if i+1 < n else ""
                if nxt == "u" and i+5 < n:
                    out.append(raw[i:i+6]); i += 6
                elif nxt in VALID:
                    out.append(c); out.append(nxt); i += 2
                else:
                    out.append("\\\\"); i += 1  # double the backslash
            elif c == '"':
                rest = raw[i+1:].lstrip()
                if rest and rest[0] in ":,]}":
                    out.append(c); in_str = False
                else:
                    out.append('\\"')
                i += 1
            elif c == "\n": out.append("\\n"); i += 1
            elif c == "\r": i += 1
            elif c == "\t": out.append("\\t"); i += 1
            elif ord(c) < 0x20: i += 1
            else: out.append(c); i += 1
        else:
            if c == '"': in_str = True
            out.append(c); i += 1

    fixed = "".join(out)
    if in_str: fixed += '"'

    try:
        return json.loads(fixed)
    except json.JSONDecodeError:
        import re as _re
        s = fixed.rstrip().rstrip(",")

        # Walk to find if we ended mid-string and count unclosed braces
        in_s2, esc2, depth2 = False, False, 0
        for ch in s:
            if esc2: esc2 = False; continue
            if ch == "\\" and in_s2: esc2 = True; continue
            if ch == '"': in_s2 = not in_s2; continue
            if not in_s2:
                if ch == '{': depth2 += 1
                elif ch == '}': depth2 -= 1

        if in_s2:
            s += '"'  # close open string

        # Append any missing required fields before closing
        for field, default in [("summary", '""'), ("keyPoints", "[]"), ("details", "[]")]:
            if ('"' + field + '"') not in s:
                s += ', "' + field + '": ' + default

        s += "}" * max(0, depth2)  # close open braces

        try:
            return json.loads(s)
        except Exception:
            # Last resort: regex-extract whatever fields exist
            result = {}
            for field in ("title", "summary"):
                m = _re.search(r'"' + field + r'"\s*:\s*"([^"]*)', s)
                if m:
                    result[field] = m.group(1)
            result.setdefault("title", "Unknown")
            result.setdefault("summary", "")
            result.setdefault("keyPoints", [])
            result.setdefault("details", [])
            return result



def is_clean_detail(text):
    import re
    if not text or len(text.strip()) < 4:
        return False
    t = text.strip()
    bad = [
        'Extra \\left', 'Extra \\right', 'Missing \\left', 'Missing \\right',
        'math input error', 'undefined control sequence',
        '\\left or', 'or missing \\right', 'or missing \\left',
    ]
    for b in bad:
        if b.lower() in t.lower():
            return False
    if re.match(r'^\\[a-zA-Z]+[\s(){}\[\]]*$', t):
        return False
    if re.match(r'^[\[\]{}()|\\]+$', t):
        return False
    if re.match(r'^[0-9]+\.[0-9]+$', t) or re.match(r'^\([A-Za-z]\.[0-9]+\)$', t):
        return False
    if not re.search(r'[a-zA-Z]', t):
        return False
    return True

def clean_page_text(text):
    import re
    lines = text.splitlines()
    cleaned = []
    for line in lines:
        s = line.strip()
        if not s:
            continue
        if any(x in s for x in ['Extra \\left', 'Extra \\right', 'Missing \\left', 'Missing \\right']):
            continue
        if re.match(r'^\\[a-zA-Z]+[\s(){}\[\]]*$', s):
            continue
        cleaned.append(s)
    result = '\n'.join(cleaned)
    result = re.sub(r'\n{3,}', '\n\n', result)
    return result.strip()

# ── Document parsers ──
def extract_pdf(path: str) -> list:
    import PyPDF2
    pages = []
    with open(path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for i, page in enumerate(reader.pages, 1):
            text = page.extract_text() or ""
            if text.strip():
                pages.append({"num": i, "label": f"Page {i}", "text": text.strip()})
    return pages

def extract_pptx(path: str) -> list:
    from pptx import Presentation
    prs = Presentation(path)
    slides = []
    for i, slide in enumerate(prs.slides, 1):
        texts = [shape.text.strip() for shape in slide.shapes if hasattr(shape, "text") and shape.text.strip()]
        if texts:
            slides.append({"num": i, "label": f"Slide {i}", "text": "\n".join(texts)})
    return slides

def extract_ppt(path: str) -> list:
    """Old .ppt format — try python-pptx (works if file is actually pptx),
    otherwise extract any readable text from the binary."""
    try:
        return extract_pptx(path)
    except Exception:
        pass
    # Fall back: pull ASCII text runs from the binary (crude but better than nothing)
    import re
    with open(path, "rb") as f:
        raw = f.read()
    # Extract printable ASCII strings of length >= 20
    strings = re.findall(rb'[ -~]{20,}', raw)
    text = "\n".join(s.decode("ascii", errors="ignore").strip() for s in strings if s.strip())
    if not text:
        return [{"num": 1, "label": "Slide 1", "text": "Could not extract text from this .ppt file. Try saving it as .pptx in PowerPoint."}]
    # Split into rough chunks per ~500 chars
    chunks, size = [], 500
    for i in range(0, len(text), size):
        chunk = text[i:i+size].strip()
        if chunk:
            chunks.append({"num": len(chunks)+1, "label": f"Slide {len(chunks)+1}", "text": chunk})
    return chunks or [{"num": 1, "label": "Slide 1", "text": text[:1000]}]

def extract_docx(path: str) -> list:
    from docx import Document
    doc = Document(path)
    sections, current_title, current_text = [], "Introduction", []
    for para in doc.paragraphs:
        if para.style.name.startswith("Heading") and para.text.strip():
            if current_text:
                sections.append({"num": len(sections)+1, "label": current_title, "text": "\n".join(current_text)})
            current_title = para.text.strip()
            current_text = []
        elif para.text.strip():
            current_text.append(para.text.strip())
    if current_text:
        sections.append({"num": len(sections)+1, "label": current_title, "text": "\n".join(current_text)})
    if not sections:
        all_text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        sections = [{"num": 1, "label": "Full document", "text": all_text}]
    return sections

def extract_txt(path: str) -> list:
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        text = f.read()
    words = text.split()
    chunks = []
    size = 500
    for i in range(0, len(words), size):
        chunk = " ".join(words[i:i+size])
        chunks.append({"num": len(chunks)+1, "label": f"Section {len(chunks)+1}", "text": chunk})
    return chunks or [{"num": 1, "label": "Full document", "text": text}]

EXTRACTORS = {
    ".pdf":  extract_pdf,
    ".docx": extract_docx,
    ".pptx": extract_pptx,
    ".ppt":  extract_ppt,
    ".txt":  extract_txt,
    ".md":   extract_txt,
}

MIME_TYPES = {
    "pdf":  "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "ppt":  "application/vnd.ms-powerpoint",
    "txt":  "text/plain",
    "md":   "text/plain",
}

app = Flask(__name__)
app.secret_key = os.urandom(24)

UPLOAD_FOLDER = Path(tempfile.gettempdir()) / "study_assistant_uploads"
UPLOAD_FOLDER.mkdir(exist_ok=True)

DOC_STORE: dict = {}

@app.route("/")
def home():
    return render_template("home.html")

@app.route("/study")
def index():
    return render_template("index.html")



@app.route("/api/status")
def status():
    return jsonify(check_ollama())

@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    file = request.files["file"]
    ext = Path(file.filename).suffix.lower()
    if ext not in EXTRACTORS:
        return jsonify({"error": f"Unsupported file type: {ext}"}), 400

    raw_bytes = file.read()
    save_path = UPLOAD_FOLDER / file.filename
    with open(str(save_path), "wb") as f:
        f.write(raw_bytes)

    try:
        pages = EXTRACTORS[ext](str(save_path))
    except Exception as e:
        import traceback; traceback.print_exc()
        save_path.unlink(missing_ok=True)
        return jsonify({"error": f"Could not parse file: {e}"}), 500

    if not pages:
        save_path.unlink(missing_ok=True)
        return jsonify({"error": "No text could be extracted from this file."}), 400

    full_text = "\n\n".join(f"[{p['label']}]\n{p['text']}" for p in pages)
    doc_id = str(abs(hash(file.filename + full_text[:100])))

    DOC_STORE[doc_id] = {
        "name": file.filename,
        "ext": ext.lstrip("."),
        "pages": pages,
        "text": full_text,
        "file_path": str(save_path),
        "word_count": len(full_text.split()),
        "summary": None,
        "chat_history": [],
    }
    return jsonify({
        "doc_id": doc_id,
        "name": file.filename,
        "ext": ext.lstrip("."),
        "word_count": len(full_text.split()),
        "page_count": len(pages),
    })

# How many pages to summarize in parallel.
# Higher = faster but uses more RAM and may overwhelm slower machines.
# Tune this: 4 is safe for most laptops, 8 is good for 16GB+ machines.
PARALLEL_WORKERS = 4

def summarize_page(page: dict) -> dict:
    """Summarize a single page — designed to run in a thread pool."""
    label = page["label"]
    text = clean_page_text(page["text"])[:3000]
    prompt = f"""Summarize this single page/slide from a study document. Be specific and detailed.

{label}:
\"\"\"
{text}
\"\"\"

Return ONLY valid JSON, no markdown, no extra text:
{{
  "title": "topic of this {label} in 5 words or less",
  "summary": "2-3 sentences covering everything on this page",
  "keyPoints": ["specific point 1", "specific point 2", "specific point 3"],
  "details": ["term, formula, or fact 1", "term, formula, or fact 2"]
}}

Rules:
- title must reflect the actual content, not just say "{label}"
- summary must mention specific concepts, names, formulas, or numbers
- keyPoints are the most important things to know from this page
- details are specific terms, definitions, formulas — only meaningful human-readable items
- NEVER include raw LaTeX errors or fragments like "Extra \\left"
- Write math using LaTeX: \\( ... \\) for inline, \\[ ... \\] for display
- Output ONLY the JSON"""
    try:
        raw = ollama_chat([{"role": "user", "content": prompt}])
        sec = safe_json(raw)
        sec["num"] = page["num"]
        sec["label"] = label
        sec.setdefault("title", label)
        sec.setdefault("summary", "")
        sec.setdefault("keyPoints", [])
        sec.setdefault("details", [])
        sec["details"] = [d for d in sec["details"] if is_clean_detail(d)]
        print(f"  ✓ {label}")
        return sec
    except Exception as e:
        print(f"  ✗ {label}: {e}")
        return {
            "num": page["num"], "label": label, "title": label,
            "summary": text[:200] + "...", "keyPoints": [], "details": []
        }

def glossary_chunk(args: tuple) -> list:
    """Extract terms from a single chunk — designed to run in a thread pool."""
    ci, total, chunk = args
    prompt = f"""Extract every keyword, term, concept, or formula that is defined or explained in this text.

TEXT:
\"\"\"
{chunk}
\"\"\"

Return ONLY valid JSON:
{{
  "terms": [
    {{"term": "exact term", "definition": "clear 1-2 sentence definition from the text"}}
  ]
}}

Rules:
- Only include terms actually defined or explained in the text
- Write math using LaTeX e.g. $A^{{-1}}$
- NEVER include LaTeX errors or garbage
- If nothing defined return {{"terms": []}}
- Output ONLY the JSON"""
    try:
        raw = ollama_chat([{"role": "user", "content": prompt}])
        parsed = safe_json(raw)
        terms = [
            t for t in parsed.get("terms", [])
            if t.get("term","").strip() and t.get("definition","").strip()
            and is_clean_detail(t.get("term",""))
        ]
        print(f"  Glossary chunk {ci+1}/{total}: {len(terms)} terms")
        return terms
    except Exception as e:
        print(f"  Glossary chunk {ci+1} failed: {e}")
        return []

@app.route("/api/analyze", methods=["POST"])
def analyze():
    data = request.json or {}
    doc_id = data.get("doc_id")
    if not doc_id or doc_id not in DOC_STORE:
        return jsonify({"error": "Document not found"}), 404

    doc = DOC_STORE[doc_id]
    pages = doc["pages"]
    t_start = time.time()

    # ── Step 1: Summarize all pages IN PARALLEL ──
    # ThreadPoolExecutor runs PARALLEL_WORKERS page summaries simultaneously.
    # Ollama handles concurrent requests to the same model fine.
    print(f"  Analyzing {len(pages)} pages with {PARALLEL_WORKERS} workers…")
    results = [None] * len(pages)
    with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as pool:
        future_to_idx = {pool.submit(summarize_page, page): i for i, page in enumerate(pages)}
        for future in as_completed(future_to_idx):
            i = future_to_idx[future]
            results[i] = future.result()
    # results is now in original page order
    section_summaries = results

    # ── Step 2: Overall summary (fast, single call) ──
    combined = "\n".join(f"- {s['label']}: {s['summary']}" for s in section_summaries[:30])
    overview_prompt = f"""Based on these page summaries from "{doc['name']}", write an overall summary.

Page summaries:
{combined}

Return ONLY valid JSON:
{{
  "title": "document title",
  "overview": "3-4 sentence overview of the entire document",
  "keyThemes": ["theme1", "theme2", "theme3", "theme4"]
}}"""
    try:
        raw = ollama_chat([{"role": "user", "content": overview_prompt}])
        overview = safe_json(raw)
    except Exception as e:
        print(f"  Overview failed: {e}")
        overview = {"title": doc["name"], "overview": f"Document with {len(pages)} pages.", "keyThemes": []}

    # ── Step 3: Glossary extraction IN PARALLEL ──
    full_text = doc["text"]
    chunks = [full_text[i:i+4000] for i in range(0, min(len(full_text), 40000), 4000)]
    chunk_args = [(i, len(chunks), chunk) for i, chunk in enumerate(chunks)]

    all_terms = []
    with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as pool:
        for terms in pool.map(glossary_chunk, chunk_args):
            all_terms.extend(terms)

    seen = set()
    unique_terms = []
    for t in all_terms:
        key = t["term"].strip().lower()
        if key not in seen:
            seen.add(key)
            unique_terms.append(t)
    unique_terms.sort(key=lambda t: t["term"].lower())

    elapsed = time.time() - t_start
    print(f"  ✓ Analysis complete in {elapsed:.1f}s ({len(pages)} pages, {len(unique_terms)} terms)")

    summary = {
        "title": overview.get("title", doc["name"]),
        "overview": overview.get("overview", ""),
        "keyThemes": overview.get("keyThemes", []),
        "sections": section_summaries,
        "glossary": unique_terms,
        "elapsed_seconds": round(elapsed, 1)
    }

    doc["summary"] = summary
    doc["chat_history"] = []
    return jsonify({"summary": summary})

@app.route("/api/chat", methods=["POST"])
def chat():
    data = request.json or {}
    doc_id = data.get("doc_id")
    question = data.get("question", "").strip()
    if not doc_id or doc_id not in DOC_STORE:
        return jsonify({"error": "Document not found"}), 404
    if not question:
        return jsonify({"error": "Question required"}), 400

    doc = DOC_STORE[doc_id]
    system = f"""You are a precise study assistant. Answer questions based only on the document below.
Cite page or slide numbers when referencing content. Be thorough but concise.

DOCUMENT:
\"\"\"
{doc['text'][:20_000]}
\"\"\""""

    history = doc["chat_history"][-10:]
    if history and history[0]["role"] != "user":
        history = history[1:]
    history = history + [{"role": "user", "content": question}]

    try:
        reply = ollama_chat(history, system=system)
        doc["chat_history"].append({"role": "user", "content": question})
        doc["chat_history"].append({"role": "assistant", "content": reply})
        return jsonify({"reply": reply})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/raw/<doc_id>")
def raw_text(doc_id):
    if doc_id not in DOC_STORE:
        return jsonify({"error": "Not found"}), 404
    doc = DOC_STORE[doc_id]
    return jsonify({"text": doc["text"], "name": doc["name"]})

@app.route("/api/file/<doc_id>")
def serve_file(doc_id):
    # Try DOC_STORE first
    if doc_id in DOC_STORE:
        doc = DOC_STORE[doc_id]
        file_path = doc.get("file_path")
        if file_path and Path(file_path).exists():
            mime = MIME_TYPES.get(doc["ext"], "application/octet-stream")
            return send_file(file_path, mimetype=mime, download_name=doc["name"])

    # Fallback: scan UPLOAD_FOLDER for any file whose hash matches doc_id
    # (handles server restart where DOC_STORE was cleared)
    for fp in UPLOAD_FOLDER.iterdir():
        try:
            ext = fp.suffix.lower()
            if ext not in EXTRACTORS:
                continue
            pages = EXTRACTORS[ext](str(fp))
            full_text = "\n\n".join(f"[{p['label']}]\n{p['text']}" for p in pages)
            candidate_id = str(abs(hash(fp.name + full_text[:100])))
            if candidate_id == doc_id:
                mime = MIME_TYPES.get(ext.lstrip("."), "application/octet-stream")
                return send_file(str(fp), mimetype=mime, download_name=fp.name)
        except Exception:
            continue

    return jsonify({"error": "Not found"}), 404

@app.route("/api/delete/<doc_id>", methods=["DELETE"])
def delete_doc(doc_id):
    doc = DOC_STORE.pop(doc_id, None)
    if doc:
        fp = doc.get("file_path")
        if fp and Path(fp).exists():
            Path(fp).unlink(missing_ok=True)
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════
#  ResearchAI — RAG-based multi-document research engine
# ════════════════════════════════════════════════════════════
import re as _re

RESEARCH_STORE: dict = {}   # session_id → research session
CHUNK_SIZE     = 400        # words per RAG chunk
TOP_K_CHUNKS   = 6          # chunks to retrieve per document for synthesis

def chunk_text(text: str, doc_name: str, doc_id: str) -> list[dict]:
    """Split text into overlapping word-level chunks for RAG retrieval."""
    words = text.split()
    chunks = []
    step = CHUNK_SIZE - 50  # 50-word overlap
    for i in range(0, len(words), step):
        chunk_words = words[i:i + CHUNK_SIZE]
        if len(chunk_words) < 30:
            break
        chunks.append({
            "doc_id":   doc_id,
            "doc_name": doc_name,
            "chunk_id": len(chunks),
            "text":     " ".join(chunk_words),
            "start_word": i,
        })
    return chunks

def simple_score(query: str, chunk_text: str) -> float:
    """
    Lightweight TF-IDF-style relevance score without any dependencies.
    Scores based on query term frequency in the chunk.
    """
    q_terms = set(_re.sub(r'[^a-z0-9 ]', ' ', query.lower()).split())
    c_lower = chunk_text.lower()
    if not q_terms:
        return 0.0
    hits = sum(c_lower.count(t) for t in q_terms if len(t) > 2)
    # Normalise by chunk length so short chunks don't get unfairly penalised
    return hits / (len(chunk_text.split()) ** 0.5 + 1)

def retrieve_chunks(query: str, chunks: list[dict], top_k: int) -> list[dict]:
    """Return the top_k most relevant chunks for a query."""
    scored = [(simple_score(query, c["text"]), c) for c in chunks]
    scored.sort(key=lambda x: x[0], reverse=True)
    return [c for _, c in scored[:top_k]]

def ollama_stream(messages: list, system: str = ""):
    """Generator that yields text tokens from Ollama streaming API."""
    payload = {
        "model": OLLAMA_MODEL,
        "messages": ([{"role": "system", "content": system}] if system else []) + messages,
        "stream": True,
        "options": {"temperature": 0.2, "num_ctx": 8192}
    }
    with requests.post(f"{OLLAMA_URL}/api/chat", json=payload, stream=True, timeout=300) as r:
        for line in r.iter_lines():
            if line:
                try:
                    data = json.loads(line)
                    token = data.get("message", {}).get("content", "")
                    if token:
                        yield token
                    if data.get("done"):
                        break
                except Exception:
                    continue


def ddg_search(query: str, max_results: int = 6) -> list:
    """
    Web search using DuckDuckGo POST — sends form data like a real browser.
    No API key required. Falls back to scraping if POST fails.
    """
    import urllib.request, urllib.parse, html as _html, re as _re
    results = []

    # Method 1: POST to DDG HTML (mimics browser form submission, bypasses blocks)
    try:
        post_data = urllib.parse.urlencode({
            "q": query, "b": "", "kl": "us-en", "df": ""
        }).encode("utf-8")
        req = urllib.request.Request(
            "https://html.duckduckgo.com/html/",
            data=post_data,
            headers={
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Content-Type": "application/x-www-form-urlencoded",
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.5",
                "Referer": "https://duckduckgo.com/",
                "Origin": "https://duckduckgo.com",
            }
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = resp.read().decode("utf-8", errors="ignore")

        print(f"  DDG POST: got {len(body)} chars")

        # Parse result blocks — each result is in a div.result
        # Extract title+url pairs
        title_re = _re.compile(
            r'<a[^>]+class="result__a"[^>]*href="([^"]*)"[^>]*>(.*?)</a>',
            _re.DOTALL | _re.IGNORECASE
        )
        # Extract snippets
        snip_re = _re.compile(
            r'<a[^>]+class="result__snippet"[^>]*>(.*?)</a>',
            _re.DOTALL | _re.IGNORECASE
        )

        titles   = title_re.findall(body)
        snippets = snip_re.findall(body)

        print(f"  Found {len(titles)} titles, {len(snippets)} snippets")

        for i, (href, title) in enumerate(titles[:max_results]):
            title_clean = _html.unescape(_re.sub(r"<[^>]+>", "", title).strip())
            snip_clean  = _html.unescape(_re.sub(r"<[^>]+>", "", snippets[i] if i < len(snippets) else "").strip())
            if title_clean and snip_clean and href.startswith("http"):
                results.append({"title": title_clean, "snippet": snip_clean[:400], "url": href})

    except Exception as e:
        print(f"  ⚠ DDG POST failed: {e}")

    # Method 2: DuckDuckGo Lite GET (simpler page, easier to parse)
    if not results:
        try:
            params = urllib.parse.urlencode({"q": query, "kl": "us-en", "o": "json"})
            req = urllib.request.Request(
                f"https://lite.duckduckgo.com/lite/?{params}",
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                    "Accept": "text/html",
                    "Accept-Language": "en-US,en;q=0.5",
                }
            )
            with urllib.request.urlopen(req, timeout=15) as resp:
                body = resp.read().decode("utf-8", errors="ignore")

            print(f"  DDG Lite: got {len(body)} chars")

            # DDG Lite structure: alternating rows of link + snippet
            # Find all table rows with actual content
            rows = _re.findall(r'<tr[^>]*>(.*?)</tr>', body, _re.DOTALL)
            i = 0
            while i < len(rows) and len(results) < max_results:
                row = rows[i]
                # Link row has class="result-link"
                link_m = _re.search(r'href="(https?://[^"]+)"[^>]*>([^<]+)<', row)
                if link_m:
                    href  = link_m.group(1)
                    title = _html.unescape(link_m.group(2).strip())
                    # Next row or same row has snippet
                    snip = ""
                    snip_m = _re.search(r'<td[^>]*>([^<]{30,})<', row)
                    if not snip_m and i + 1 < len(rows):
                        snip_m = _re.search(r'<td[^>]*>([^<]{30,})<', rows[i+1])
                    if snip_m:
                        snip = _html.unescape(snip_m.group(1).strip())
                    if title and snip:
                        results.append({"title": title, "snippet": snip[:400], "url": href})
                i += 1

        except Exception as e:
            print(f"  ⚠ DDG Lite failed: {e}")

    print(f"  🌐 Web: {len(results)} results for '{query[:50]}'")
    return results


@app.route("/api/research/file/<doc_id>")
def research_file(doc_id):
    """Serve an uploaded research document file."""
    for session in RESEARCH_STORE.values():
        doc = session["docs"].get(doc_id)
        if doc and doc.get("file_path"):
            fp = Path(doc["file_path"])
            if fp.exists():
                mime = MIME_TYPES.get(doc["ext"], "application/octet-stream")
                return send_file(str(fp), mimetype=mime, download_name=doc["name"])
    # Fallback: scan upload folder, recompute hash same way as research_upload
    for fp in UPLOAD_FOLDER.iterdir():
        try:
            ext = fp.suffix.lower()
            if ext not in EXTRACTORS:
                continue
            pages = EXTRACTORS[ext](str(fp))
            full_text = "\n\n".join(p["text"] for p in pages)
            candidate_id = str(abs(hash(fp.name + full_text[:80])))
            if candidate_id == doc_id:
                mime = MIME_TYPES.get(ext.lstrip("."), "application/octet-stream")
                return send_file(str(fp), mimetype=mime, download_name=fp.name)
        except Exception:
            continue
    return jsonify({"error": "Not found"}), 404

# ── Research routes ──

@app.route("/research")
def research_page():
    return render_template("research.html")

@app.route("/api/research/upload", methods=["POST"])
def research_upload():
    """Upload a document into a research session."""
    session_id = request.form.get("session_id", "default")
    if "file" not in request.files:
        return jsonify({"error": "No file"}), 400

    file = request.files["file"]
    ext  = Path(file.filename).suffix.lower()
    if ext not in EXTRACTORS:
        return jsonify({"error": f"Unsupported: {ext}"}), 400

    raw  = file.read()
    path = UPLOAD_FOLDER / file.filename
    path.write_bytes(raw)

    try:
        pages = EXTRACTORS[ext](str(path))
    except Exception as e:
        import traceback; traceback.print_exc()
        path.unlink(missing_ok=True)
        return jsonify({"error": str(e)}), 500
    # Keep file on disk so View file tab can serve it

    full_text = "\n\n".join(p["text"] for p in pages)
    doc_id    = str(abs(hash(file.filename + full_text[:80])))

    if session_id not in RESEARCH_STORE:
        RESEARCH_STORE[session_id] = {"docs": {}, "query": "", "report": None}

    session = RESEARCH_STORE[session_id]
    chunks  = chunk_text(full_text, file.filename, doc_id)

    session["docs"][doc_id] = {
        "doc_id":     doc_id,
        "name":       file.filename,
        "ext":        ext.lstrip("."),
        "file_path":  str(path),
        "pages":      pages,
        "text":       full_text,
        "word_count": len(full_text.split()),
        "page_count": len(pages),
        "chunks":     chunks,
        "summary":    None,
    }

    return jsonify({
        "doc_id":     doc_id,
        "name":       file.filename,
        "ext":        ext.lstrip("."),
        "word_count": len(full_text.split()),
        "page_count": len(pages),
        "chunk_count": len(chunks),
    })

@app.route("/api/research/remove", methods=["POST"])
def research_remove():
    data       = request.json or {}
    session_id = data.get("session_id", "default")
    doc_id     = data.get("doc_id")
    if session_id in RESEARCH_STORE:
        RESEARCH_STORE[session_id]["docs"].pop(doc_id, None)
    return jsonify({"ok": True})

@app.route("/api/research/analyze", methods=["POST"])
def research_analyze():
    """
    Full RAG research pipeline:
    1. Retrieve top chunks per doc for the query
    2. Generate per-doc query-aware summaries (parallel)
    3. Cross-doc comparison
    4. Structured report (background, findings, consensus, contradictions, gaps)
    5. Knowledge gaps
    All streamed back as server-sent events.
    """
    from flask import Response, stream_with_context

    data              = request.json or {}
    session_id        = data.get("session_id", "default")
    query             = data.get("query", "").strip()
    docs_only         = data.get("docs_only", True)
    selected_doc_ids  = data.get("selected_doc_ids", [])

    if not query:
        return jsonify({"error": "Query required"}), 400
    if session_id not in RESEARCH_STORE:
        RESEARCH_STORE[session_id] = {"docs": {}, "query": "", "report": None}

    session = RESEARCH_STORE[session_id]
    all_docs = list(session["docs"].values())

    # Filter to only the docs the user selected in the frontend
    if selected_doc_ids:
        docs = [d for d in all_docs if d["doc_id"] in selected_doc_ids]
    else:
        docs = all_docs

    if not docs and docs_only:
        return jsonify({"error": "NO_DOCS_SERVER_RESET"}), 400

    session["query"] = query

    def generate():
        def emit(event: str, data_obj):
            yield f"data: {json.dumps({'event': event, **data_obj})}\n\n"

        yield from emit("start", {"total_docs": len(docs), "query": query})

        # ── Step 1: RAG retrieval per document ──
        all_chunks = []
        for doc in docs:
            all_chunks.extend(doc["chunks"])

        doc_chunks = {}   # doc_id → list of top chunks
        for doc in docs:
            top = retrieve_chunks(query, doc["chunks"], TOP_K_CHUNKS)
            doc_chunks[doc["doc_id"]] = top

        if docs:
            yield from emit("status", {"msg": f"Retrieved passages from {len(docs)} documents"})

        # ── Web search (when docs_only is False) ──
        web_results = []
        if not docs_only:
            yield from emit("status", {"msg": "Searching the web…"})
            web_results = ddg_search(query, max_results=6)
            if web_results:
                yield from emit("web_results", {"results": web_results})
                msg = f"Found {len(web_results)} web sources"
                if docs: msg += f" + {len(docs)} document(s)"
                yield from emit("status", {"msg": msg})
            else:
                yield from emit("status", {"msg": "No web results — continuing with documents"})

        # ── Step 2: Per-doc query-aware summaries (parallel) ──
        def summarize_doc_for_query(doc):
            chunks = doc_chunks[doc["doc_id"]]
            ctx    = "\n\n---\n\n".join(c["text"] for c in chunks)
            prompt = f"""You are a research assistant. Based ONLY on these excerpts from "{doc['name']}", answer the research question.

Research question: {query}

Excerpts from {doc['name']}:
\"\"\"{ctx}\"\"\"

Return ONLY valid JSON:
{{
  "doc_name": "{doc['name']}",
  "relevance": "high/medium/low — how relevant is this document to the query?",
  "key_finding": "The single most important finding from this document relevant to the query (1-2 sentences)",
  "summary": "3-4 sentence summary of what this document says about the research question",
  "methods": "How does this document approach/study the topic? (1-2 sentences, or 'Not specified')",
  "evidence": ["specific quote or fact 1 from the text", "specific quote or fact 2", "specific quote or fact 3"],
  "conclusion": "What conclusion does this document reach on the topic? (1-2 sentences)"
}}
Output ONLY the JSON."""
            try:
                raw = ollama_chat([{"role": "user", "content": prompt}])
                result = safe_json(raw)
                result["doc_id"]   = doc["doc_id"]
                result["doc_name"] = doc["name"]
                result.setdefault("relevance",   "medium")
                result.setdefault("key_finding", "")
                result.setdefault("summary",     "")
                result.setdefault("methods",     "Not specified")
                result.setdefault("evidence",    [])
                result.setdefault("conclusion",  "")
                print(f"  ✓ Summarized {doc['name']} for query")
                return result
            except Exception as e:
                print(f"  ✗ {doc['name']}: {e}")
                return {
                    "doc_id": doc["doc_id"],
                    "doc_name": doc["name"],
                    "relevance": "unknown",
                    "key_finding": "Could not analyze",
                    "summary": "Analysis failed for this document.",
                    "methods": "N/A",
                    "evidence": [],
                    "conclusion": "N/A"
                }

        doc_summaries = [None] * len(docs)
        with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as pool:
            futures = {pool.submit(summarize_doc_for_query, doc): i for i, doc in enumerate(docs)}
            for future in as_completed(futures):
                i = futures[future]
                doc_summaries[i] = future.result()
                yield from emit("doc_done", {
                    "doc_id":  doc_summaries[i]["doc_id"],
                    "doc_name": doc_summaries[i]["doc_name"],
                    "summary": doc_summaries[i]
                })

        # ── Step 3: Cross-doc comparison (skip if no doc summaries) ──
        comparison = {"consensus": [], "contradictions": [], "unique_contributions": [], "overall_verdict": ""}
        if doc_summaries:
            yield from emit("status", {"msg": "Generating cross-document comparison…"})
            summaries_text = "\n\n".join(
                f"Document: {s['doc_name']}\n"
                f"Key finding: {s.get('key_finding','')}\n"
                f"Methods: {s.get('methods','')}\n"
                f"Conclusion: {s.get('conclusion','')}"
                for s in doc_summaries
            )
            compare_prompt = f"""You are a research analyst. Compare these {len(doc_summaries)} documents.

Research question: {query}

Document summaries:
{summaries_text}

Return ONLY valid JSON:
{{
  "consensus": ["point all or most documents agree on"],
  "contradictions": ["point where documents disagree"],
  "unique_contributions": [
    {{"doc_name": "document name", "contribution": "unique insight from this doc"}}
  ],
  "overall_verdict": "2-3 sentence synthesis of what the documents collectively say"
}}
Output ONLY the JSON."""
            try:
                raw = ollama_chat([{"role": "user", "content": compare_prompt}])
                comparison = safe_json(raw)
            except Exception as e:
                print(f"  Comparison failed: {e}")

        yield from emit("comparison", {"comparison": comparison})
        yield from emit("status", {"msg": "Writing research report…"})

        # ── Step 4: Structured report (streamed) ──
        all_evidence = "\n\n".join(
            f"[{s['doc_name']}]: {s.get('summary','')}"
            for s in doc_summaries
        )
        web_section = ""
        if web_results:
            web_lines = "\n".join(f"- [{r['title']}] ({r['url']}): {r['snippet']}" for r in web_results)
            web_section = f"\n\nWeb search results:\n{web_lines}"
        report_prompt = f"""You are writing a structured research report based on the documents provided.

Research question: {query}

Evidence from uploaded documents:
{all_evidence}{web_section}

Cross-document analysis:
- Consensus: {comparison.get('overall_verdict','')}
- Agreements: {'; '.join(comparison.get('consensus',[]))}
- Contradictions: {'; '.join(comparison.get('contradictions',[]))}

Write a structured research report with these exact sections using markdown:

## Background
(2-3 sentences framing the research question and why it matters)

## Key Findings
(bullet points — one per document, cite the document name in brackets)

## Points of Consensus
(what the documents agree on)

## Contradictions & Debates
(where documents disagree, or if none say so)

## Synthesis
(3-4 sentence overall answer to the research question based on all documents)

## Knowledge Gaps
(what questions these documents do NOT answer — what would require more research)

Be specific. Cite document names. Do not fabricate information not in the documents."""

        report_text = ""
        for token in ollama_stream([{"role": "user", "content": report_prompt}]):
            report_text += token
            yield from emit("report_token", {"token": token})

        yield from emit("status", {"msg": "Analysis complete"})

        # Store result
        session["report"] = {
            "query":        query,
            "doc_summaries": doc_summaries,
            "comparison":   comparison,
            "report_text":  report_text,
        }

        yield from emit("done", {
            "doc_summaries": doc_summaries,
            "comparison":   comparison,
            "report_text":  report_text,
        })

    return Response(stream_with_context(generate()), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

@app.route("/api/research/chat", methods=["POST"])
def research_chat():
    """Multi-doc grounded chat — retrieves from ALL documents."""
    data       = request.json or {}
    session_id = data.get("session_id", "default")
    question   = data.get("question", "").strip()

    if not question:
        return jsonify({"error": "Question required"}), 400
    if session_id not in RESEARCH_STORE:
        return jsonify({"error": "Session not found"}), 404

    session = RESEARCH_STORE[session_id]
    docs    = list(session["docs"].values())
    history = data.get("history", [])

    # Retrieve top chunks from ALL docs combined
    all_chunks = []
    for doc in docs:
        all_chunks.extend(doc["chunks"])
    top = retrieve_chunks(question, all_chunks, top_k=12)

    context = "\n\n---\n\n".join(
        f"[{c['doc_name']}]:\n{c['text']}" for c in top
    )
    system = f"""You are a research assistant with access to {len(docs)} documents.
Answer questions based ONLY on the provided document excerpts.
Always cite which document(s) your answer comes from using [Document Name] format.
If the answer isn't in the documents, say so clearly.

DOCUMENT EXCERPTS:
\"\"\"{context}\"\"\""""

    msgs = history[-10:] + [{"role": "user", "content": question}]
    if msgs and msgs[0]["role"] != "user":
        msgs = msgs[1:]

    try:
        reply = ollama_chat(msgs, system=system)
        return jsonify({"reply": reply})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/research/evidence", methods=["POST"])
def research_evidence():
    """Find evidence for a specific claim across documents and optionally the web."""
    data       = request.json or {}
    session_id = data.get("session_id", "default")
    claim      = data.get("claim", "").strip()
    search_web = data.get("search_web", False)

    if not claim:
        return jsonify({"error": "Invalid request"}), 400

    results = []

    # Search uploaded documents via RAG (only if session exists)
    if session_id in RESEARCH_STORE:
        docs = list(RESEARCH_STORE[session_id]["docs"].values())
        for doc in docs:
            top = retrieve_chunks(claim, doc["chunks"], top_k=3)
            if top and simple_score(claim, top[0]["text"]) > 0.1:
                results.append({
                    "type":     "document",
                    "doc_id":   doc["doc_id"],
                    "doc_name": doc["name"],
                    "passages": [c["text"][:400] for c in top],
                })

    # Search the web if requested
    if search_web:
        web = ddg_search(claim, max_results=5)
        for r in web:
            results.append({
                "type":    "web",
                "title":   r["title"],
                "url":     r["url"],
                "snippet": r["snippet"],
            })

    return jsonify({"results": results, "claim": claim})

if __name__ == "__main__":
    print("\n🧠  MindForge — StudyAI + ResearchAI (powered by Ollama)")
    info = check_ollama()
    if not info["running"]:
        print("\u26a0  Ollama not running. Run: ollama serve")
    elif not info["model_ready"]:
        print(f"\u26a0  Model not found. Run: ollama pull {OLLAMA_MODEL}")
    else:
        print(f"\u2713  Ollama running \u2014 model: {OLLAMA_MODEL}")
    print("\n   Open http://localhost:5000\n")
    app.run(debug=True, port=5000, use_reloader=False)
