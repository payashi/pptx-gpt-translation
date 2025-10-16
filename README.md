# GPT PPTX Translator

Translate PowerPoint (`.pptx`) decks using **GPT** with **python-pptx**. 
Designed for high-quality, context-aware translation (e.g., EN→JA) with optional glossary enforcement and notes/masters support.

## Features
- Slide-aware context: include slide titles, neighbors, and optional speaker notes as **context**.
- Batching by slide to keep topics coherent and control token usage.
- Optional **glossary** (CSV) post-pass to enforce terminology.
- Supports translating **speaker notes** and **slide masters/layouts**.
- Dry-run preview.
- Simple, dependency-light CLI.

## Install
```bash
python -m venv .venv && source .venv/bin/activate   # on Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Set your API key:
```bash
export OPENAI_API_KEY=sk-...   # Windows (PowerShell): $env:OPENAI_API_KEY="sk-..."
```

## Usage
```bash
python gpt_pptx_translate.py input.pptx output_ja.pptx   --target JA --source EN   --notes --masters   --glossary glossary_example.csv   --model gpt-4o-mini   --strategy neighbor   --dry_run
```

Key flags:
- `--source` / `--target`: language codes (free-form; used to guide GPT).
- `--notes`: include speaker notes for translation and as extra context.
- `--masters`: include slide masters/layouts (titles, footers) too.
- `--glossary glossary.csv`: CSV with `source,target` (no header) to enforce terms *after* translation.
- `--model`: any chat-capable GPT model (e.g., `gpt-4o`, `gpt-4o-mini`).
- `--strategy`:
  - `neighbor` (default): send current slide plus neighbors +/-1 as context.
  - `title-only`: use deck title & current slide title as context.
  - `deck`: send a truncated deck summary (first 2 slides + section titles) as context.
- `--dry_run`: preview what would change without saving.

## Notes & Limits
- Text inside images is not translated (OCR not included).
- Complex inline formatting (multiple runs inside a textbox) is simplified (text replaced as a whole).
- Some chart elements may not expose text frames via python-pptx.
- Keep an eye on **token usage**; this tool avoids sending the entire deck every time to reduce cost/latency.

## Cost tip
If cost matters, try `--model gpt-4o-mini` first. For best nuance/tone, try `gpt-4o`.

## Glossary example
See `glossary_example.csv` for format. The glossary is applied as an exact string replacement after GPT translation.

## License
MIT