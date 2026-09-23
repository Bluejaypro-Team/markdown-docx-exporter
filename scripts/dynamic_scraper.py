import sys
import argparse
from playwright.sync_api import sync_playwright
from readability import Document
from markdownify import markdownify as md

def scrape_dynamic_page(url, output_path):
    print(f"[1/4] Launching Playwright browser...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        print(f"[2/4] Navigating to {url} and waiting for network to idle...")
        page.goto(url, wait_until="networkidle")
        
        # Give it an extra second just in case there are late-firing rendering scripts
        page.wait_for_timeout(2000)
        
        html_content = page.content()
        browser.close()

    print("[3/4] Extracting main content using Readability...")
    doc = Document(html_content)
    # Get the extracted HTML of the main content
    article_html = doc.summary()
    title = doc.title()

    print("[4/4] Converting to clean Markdown...")
    # Convert HTML to markdown, keeping it clean
    markdown_content = f"# {title}\n\n" + md(article_html, heading_style="ATX")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(markdown_content)
        
    print(f"Success! Dynamic content saved to {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scrape dynamic JS-rendered websites and convert to Markdown.")
    parser.add_argument("url", help="The URL to scrape.")
    parser.add_argument("output", help="The output Markdown file path.")
    
    args = parser.parse_args()
    scrape_dynamic_page(args.url, args.output)
