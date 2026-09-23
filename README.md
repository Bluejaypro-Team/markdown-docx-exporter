# Markdown → Word (.docx) Exporter

> **Premium-styled document conversion pipeline.** Takes any static `.md` file or dynamic JavaScript-rendered live URL and produces an export-ready `.docx` Word document with zero intermediate steps.

Designed specifically for the **Bluejaypro Team**, enforcing strict Midnight Gold / Bluejaypro design tokens automatically.

## Features

- **The "Random Forest" Dynamic Scraper:** Integrated `playwright` Chromium automation completely bypasses "Live Content" shells on JS-rendered sites (like Google Support), extracting the true HTML payload before conversion.
- **Premium Typography:** Segoe UI body, Consolas code, and hierarchical heading sizes.
- **Color-Coded Headings:** 3-tier gradient matching the Bluejaypro brand aesthetic (Deep Blue → Medium Blue → Steel Blue).
- **Styled Tables:** Dark header rows with white text and alternating `#F2F5F8` row shading.
- **Single-Pass Engine:** The full document is processed in one run—no splitting required. Inline formatting (bold, italic, links, code) is perfectly preserved.

## Architecture

This tool operates as an isolated Agentic Skill within the Antigravity IDE:

```text
Input (.md file or Live Web URL)
    ↓
[Python Pipeline: dynamic_scraper.py -> convert.py]
    ↓
Output (.docx file — export-ready)
```

## Setup & Installation

The script requires a virtual environment with specific dependencies. 

```bash
# Clone the repository
git clone https://github.com/Bluejaypro-Team/markdown-docx-exporter.git
cd markdown-docx-exporter

# Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate  # Or .venv\Scripts\activate on Windows

# Install dependencies
pip install -r requirements.txt

# Install Playwright browser binaries for the dynamic scraper
playwright install chromium
```

## Usage

### 1. Static Markdown Conversion
Convert a local markdown file directly into a styled `.docx`:
```bash
python scripts/convert.py "input.md" "output.docx" --title "My Document Title"
```

### 2. Live Web Source Extraction (Dynamic Sites)
Scrape a live JavaScript-rendered URL and export it to a markdown file:
```bash
python scripts/dynamic_scraper.py "https://example.com" "scraped_output.md"
```
*(You can then pipe `scraped_output.md` directly into `convert.py`)*

## Strict Invocation Constraint (Agentic Workflow)
If used as an AI Agent Skill: Do **NOT** automatically invoke this skill or pipe data to it during other automated workflows (e.g., `a2a-content-pipeline` or `url-cluster-analyzer`) unless specifically requested by the user. It operates in strict manual isolation.
