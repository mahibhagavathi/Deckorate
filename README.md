# 🎨 Deckorate AI

An AI-powered Streamlit app that transforms messy PowerPoint decks into consulting-grade presentations — powered by Google Gemini (free tier available).

## Features

- **Upload** any `.pptx` deck
- **Pick a colour palette** — 6 consulting presets (McKinsey, BCG, Bain, etc.) or build your own
- **Choose font pairing** — 7 curated pairings with live preview
- **Cleanup toggles** — fix typos, standardise font sizes, align text, clean bullets
- **Per-slide preview & editing** — Gemini AI suggests the best layout per slide, you can edit text, swap layouts, override fonts and colours per slide
- **3–4 layout templates per slide** based on slide type
- **Download** a fully editable `.pptx`

## Quick start

```bash
pip install -r requirements.txt
streamlit run app.py
```

Get a **free** Gemini API key at [ai.google.dev](https://ai.google.dev) — no credit card needed. Paste it into the sidebar when the app loads.

Or set it as an environment variable:

```bash
export GEMINI_API_KEY=your_key_here
streamlit run app.py
```

## Resume talking points

- Built a full-stack AI application using **Streamlit + Python**
- Integrated **Google Gemini API** for per-slide content analysis and layout recommendation
- Used **python-pptx** to programmatically parse and rewrite PowerPoint XML
- Designed a multi-step UX with persistent session state, real-time previews, and per-slide override logic
- Implemented a colour palette system with 6 consulting presets and a custom colour picker
