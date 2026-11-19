# Minimize Pages

A client-side web tool that reduces odd-page `.docx` files to an even page count for double-sided printing on A4 paper. No server, no uploads — everything runs in the browser.

## What it does

When a Word document has an odd number of pages (3, 5, 7…), printing double-sided wastes one blank sheet. This tool adjusts spacing to fit the content into one fewer page — or, if that's not possible, suggests expanding to the next even page count instead.

### Adjustments applied (by level)

| Level | Line spacing | Paragraph gaps | Margins | Images |
|-------|-------------|----------------|---------|--------|
| 1–2   | Slightly tighter | Reduced | Unchanged | Slightly smaller |
| 3–5   | Tight | Significantly reduced | Narrow (1.27cm) | Smaller |
| 6–7   | Very tight | Minimal | Narrow (1.27cm) | Much smaller |

### Rules

- **Content is never removed** — all text, vocabulary, logos, headers, and dates stay intact.
- **Font size is never reduced** — minimum 12pt guaranteed.
- **Target is always even** — if reducing fails, the tool offers to expand to the next even page count.
- Margins that are already smaller than Narrow are left as-is.

## Usage

1. Open `index.html` in a browser.
2. Drop a `.docx` file or click to browse.
3. Enter the current page count (as shown in Word).
4. Choose a reduction level or leave on **Auto** (tries levels progressively until it fits).
5. Download the result and verify in Word before printing.

## Deploy

Static files only — upload the folder to any hosting:

```
index.html
style.css
app.js
```

JSZip is loaded from cdnjs CDN. No build step, no dependencies to install.

### Cache busting

CSS and JS are loaded with `?v=` query strings. Bump the version number in `index.html` after each update.

## Tech

- **JSZip** — unzips and re-zips the `.docx` (which is a ZIP of XML files)
- **DOMParser / XMLSerializer** — parses and modifies `word/document.xml` directly
- Adjusts `w:spacing` (line height, before/after), `w:pgMar` (margins), and `wp:extent` / `a:ext` (image dimensions)
