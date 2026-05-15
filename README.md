# StudyAI — Python Study Assistant

An AI-powered study assistant built with Python (Flask) and the Anthropic API.
Upload any document and get a detailed, section-by-section breakdown with key points,
summaries, themes, and an interactive Q&A chat.

## Supported formats
- **PDF** — parsed page by page
- **DOCX** — Word documents
- **PPTX** — PowerPoint (slide by slide)
- **TXT / MD** — plain text and Markdown

---

## Setup

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
study_assistant/
├── app.py              ← Flask backend + document parsers + Anthropic API calls
├── requirements.txt    ← Python dependencies
├── templates/
│   └── index.html      ← Full UI (single-page app)
└── README.md
```
