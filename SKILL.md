---
name: markdown-docx-exporter
description: |
  Converts any Markdown file (.md) into a premium-styled, export-ready
  Microsoft Word document (.docx) with color-coded headings, styled tables,
  syntax-highlighted code blocks, and professional headers/footers.

  Use this skill if any of the following are true:
    1. User asks to "export to Word", "convert to docx", "download as Word file",
       or "make a Word document" from a Markdown file or artifact.
    2. User asks to "export this artifact" and the source is a .md file.
    3. User provides a Markdown file path and wants a printable/shareable document.
license: Apache-2.0
metadata:
  version: v1.0.0
  publisher: bluejaypro
---

# Markdown → Word (.docx) Exporter

> Premium-styled document conversion. Takes any `.md` file and produces an
> export-ready `.docx` with zero intermediate steps.

## When to Invoke

> [!WARNING]
> **Strict Invocation Constraint:** Do NOT automatically invoke this skill, pipe data to it, or run it during other workflows (e.g., `a2a-content-pipeline`, `bjp-cro-intelligence`, `url-cluster-analyzer`, etc.) unless specifically asked. This skill operates in strict manual isolation.

Trigger this skill **ONLY** when the user explicitly requests it, such as:
- "Export to Word / docx"
- "Download as Word file"
- "Convert this markdown to a document"
- "Make a printable version of this artifact"
- "Run the docx exporter"
- Any explicit request pairing a `.md` file with `.docx` output intent

## Architecture

```
Input (.md file)
    ↓
[Python script: scripts/convert.py]
    ↓
Output (.docx file — export-ready)
```

No splitting, no scaffolding. The script reads the full markdown, converts it
to styled Word XML in a single pass, and writes the final `.docx`.

## Execution Workflow

### Step 1 — Resolve Inputs

Determine the **input markdown path** and **output docx path**:

- **Local Input**: The `.md` file the user referenced.
- **Web Input (Random Forest Test)**: If the user provides a live URL, you MUST first run `dynamic_scraper.py` to extract the JavaScript-rendered content into a markdown file.
  ```powershell
  C:\Users\pczoon\.gemini\antigravity-ide\scratch\docx-exporter\.venv\Scripts\python.exe ^
    C:\Users\pczoon\.gemini\antigravity-ide\scratch\docx-exporter\dynamic_scraper.py ^
    "<URL>" ^
    "C:\Users\pczoon\.gemini\antigravity-ide\scratch\docx-exporter\web_scrape.md"
  ```
  Then use `web_scrape.md` as your input markdown path.
- **Output**: Default to the same directory as the input file, with the `.md` extension replaced by `.docx`. If the user wants it in the workspace, use `C:\Users\pczoon\.gemini\antigravity-ide\scratch\docx-exporter\`.

### Step 2 — Ensure Virtual Environment Exists

The script requires `python-docx`, `markdown`, and `beautifulsoup4`. The
skill maintains a dedicated venv at:

```
C:\Users\pczoon\.gemini\antigravity-ide\scratch\docx-exporter\.venv
```

**Check if it exists first.** If missing, create it and install dependencies:

```powershell
python -m venv C:\Users\pczoon\.gemini\antigravity-ide\scratch\docx-exporter\.venv
C:\Users\pczoon\.gemini\antigravity-ide\scratch\docx-exporter\.venv\Scripts\pip.exe install -r C:\Users\pczoon\.gemini\antigravity-ide\scratch\docx-exporter\requirements.txt
```

### Step 3 — Run the Converter

Execute the script using the skill's venv python:

```powershell
C:\Users\pczoon\.gemini\antigravity-ide\scratch\docx-exporter\.venv\Scripts\python.exe ^
  C:\Users\pczoon\.gemini\config\skills\markdown-docx-exporter\scripts\convert.py ^
  "<INPUT_MD_PATH>" ^
  "<OUTPUT_DOCX_PATH>" ^
  --title "<DOCUMENT_TITLE>"
```

**Arguments:**

| Arg | Required | Description |
|---|---|---|
| `<INPUT_MD_PATH>` | Yes | Absolute path to the source `.md` file |
| `<OUTPUT_DOCX_PATH>` | Yes | Absolute path for the output `.docx` file |
| `--title "<TITLE>"` | No | Custom header title. Defaults to auto-detected H1 or filename |

### Step 4 — Report Results

After successful conversion:
1. Confirm the output path with a clickable file link
2. Report the file size

## Design Tokens

The script uses a hardcoded, single-pass styling process based on the **Bluejaypro Deployment RAG** design:
- **Headings**: Deep Blue (#1F4E78) → Medium Blue (#2E75B6) → Steel Blue (#41719C)
- **Table headers**: #1F4E78 with white text
- **Alternating rows**: #F2F5F8
- **Code blocks**: #F4F4F5 background, Consolas font
- **Inline code**: Deep red (#C7254E)
- **Links**: #1F4E78 underlined

## Features

- **Premium typography**: Segoe UI body, Consolas code, hierarchical heading sizes
- **Color-coded headings**: 3-tier gradient matching the selected theme
- **Styled tables**: Dark header row with white text, alternating row shading
- **Code blocks**: Shaded background, monospace font, proper indentation
- **Blockquotes**: Left-indented with subtle background shading
- **Inline formatting**: Bold, italic, code, links all preserved
- **Headers & footers**: Auto-generated document title + page numbers
- **Horizontal rules**: Rendered as styled bottom borders
- **Lists**: Bullet and numbered lists with proper nesting depth
- **Single-pass**: Full document processed in one run — no splitting required
