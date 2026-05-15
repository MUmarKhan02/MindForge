# MindForge

An AI-powered study and research assistant built with Python (Flask) and the Anthropic API.
Upload any document and get a detailed, section-by-section breakdown with key points,
summaries, themes, and an interactive Q&A chat.

## Supported formats
- **PDF** — parsed page by page
- **DOCX** — Word documents
- **PPTX** — PowerPoint (slide by slide)
- **TXT / MD** — plain text and Markdown

---

## Setup

### Prerequisites
- Python 3.x
- [Ollama](https://ollama.com) installed and running (used internally by the app)

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the app
```bash
python app.py
```

### 3. Open in your browser
```
http://localhost:5000
```

### 4. Add your API key
Paste your Anthropic API key into the field in the top-right corner.
Get one at https://console.anthropic.com

---

## Windows quick launch (recommended)

A `MindForge.bat` file is included in the root of the repo for Windows users.
Double-clicking it will:
1. Start Ollama in the background (skips if already running)
2. Launch the Flask app
3. Open your browser to `http://localhost:5000` automatically

> **Note:** Make sure Ollama is installed before using the launcher.
> The `.bat` file should stay in the root folder (next to the `study_assistant/` directory) — don't move it inside.

---

## Features

| Feature | Description |
|---|---|
| Document upload | Drag & drop or click to upload |
| AI analysis | Full section-by-section summary with key points |
| Key themes | High-level themes extracted across the document |
| Notable details | Small but important terms, figures, dates |
| Raw text view | See exactly what was extracted from your file |
| Ask AI (chat) | Ask follow-up questions grounded in your document |
| Multi-doc | Upload and switch between multiple documents |

---

## Project structure
```
MindForge/
├── MindForge.bat           ← Windows launcher (starts Ollama + Flask + browser)
└── study_assistant/
    ├── app.py              ← Flask backend + document parsers + Anthropic API calls
    ├── requirements.txt    ← Python dependencies
    ├── static/
    │   └── MindForge_Logo.png
    ├── templates/
    │   ├── home.html       ← Landing page
    │   ├── index.html      ← Study assistant UI
    │   └── research.html   ← Research assistant UI
    └── README.md
```
