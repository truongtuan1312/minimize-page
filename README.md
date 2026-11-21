# Worksheet Minimizer

A browser tool that reshapes a `.docx` worksheet to fit a page count you choose. Built for printable classroom handouts — vocabulary tables, picture grids, reminder sheets — where the goal is usually an even page count for double-sided printing.

Nothing is uploaded. The file is unzipped, edited and re-zipped entirely in your browser.

## Using it

1. Open `index.html` in a browser (or visit the hosted copy).
2. Drop a `.docx` file on the left panel, or click to browse.
3. **Read the scan.** The left panel reports what it measured: how many pictures, the tallest one, picture rows versus text rows, the largest font, blank lines, and the top margin. Values that break your limits turn amber.
4. **Check the detected page count.** The app reads it from the file — you don't type it in. See [Page detection](#page-detection) for when it can be wrong.
5. **Pick your target** with the `− / +` stepper. The app works out what to do:
   - **Target lower** → minimizes, using the rules below
   - **Target higher** → expands, opening up line spacing and row heights
   - **Target equal** → nothing to do, the button disables
6. **Adjust the rules** if the estimate says you won't make it.
7. Download, then **open it in Word and check it before printing.**

### The estimate meter

Shows roughly how much vertical height the enabled rules will reclaim, measured against how much is actually needed to drop the pages you asked for. It computes the usable page height from the document's own page size and margins.

It is an estimate, not a guarantee. Only Word knows how text truly reflows, so the final page count can differ. Always verify.

## Rules

Each rule is a toggle; three carry an editable value. Defaults in bold.

| Rule | Value | What it does |
|------|-------|--------------|
| Cap font size | **18 pt** | Finds the largest font, scales it to the cap, and scales every other font by that same ratio so relative sizes are preserved |
| Picture cell height | **0.75 in** | Caps the height of table rows holding pictures, and shrinks pictures taller than that to fit (aspect ratio kept) |
| Text cell height | **0.2 in** | Minimum height for table rows with no picture |
| Tighten line spacing | auto | Halves space-before, trims space-after, pulls line spacing toward single |
| Shrink page margins | auto | Sets the top margin to zero when a header exists, so the header art provides the visual gap |
| Drop blank lines | auto | Removes empty paragraphs outside tables |
| Scale loose pictures | auto | Shrinks pictures that sit outside tables, toward the page target |

**Reset** restores all defaults.

### Why pictures in tables are treated separately

Pictures sitting in table cells are the usual case in a vocabulary worksheet, and they are the thing most likely to break. Scaling them by a flat percentage makes the cell text overlap the artwork. Instead, picture rows get their own height cap and pictures are fitted to it — so a 2 in picture becomes 0.75 in and the row closes up, with the text staying clear.

Pictures in the same row sit side by side, so only the tallest one in a row sets that row's height. The estimate accounts for this.

## Page detection

Word writes its own page count into `docProps/app.xml` each time it saves. That number is exact and is what the app uses when present — the readout says *read from the file*.

It can be wrong in one case: **the file was last saved by something other than Word** (Google Docs export, LibreOffice, a script, or this tool's own output). Then the count is stale or missing. When it's missing the app estimates from content height and says *estimated*, with a correction box underneath. When it's present but you know it's wrong, click **wrong?** next to the count to override it.

## Known limits

- **The font cap can push small text below 12 pt.** Because every font scales by the same ratio, a document whose largest font is 28 pt gets a ratio of `18/28 ≈ 0.64`, which turns 12 pt body text into about 7.7 pt. If you need a floor, raise the cap or switch the rule off.
- **Page counts are not simulated.** The app edits the document's layout instructions; Word does the actual layout. Hitting your exact target is not guaranteed.
- **Expanding is coarse.** Growing a document stretches spacing proportionally, which is less precise than shrinking.
- Only `.docx` is supported — not `.doc`, not Google Docs links.

## Deploy

Static files, no build step:

```
index.html
style.css
app.js
```

Upload the folder to any static host. JSZip loads from cdnjs.

### Cache busting

`app.js` is loaded with a `Date.now()` query string, so code changes take effect on reload with no action needed. `style.css` still uses a manual `?v=N` in `index.html` — **bump it after editing the stylesheet**, or returning visitors keep the old CSS.

## How it works

A `.docx` is a ZIP of XML. The app opens it with **JSZip**, parses the XML with **DOMParser**, edits nodes, and serialises back with **XMLSerializer**.

Parts it reads and writes:

| Part | Used for |
|------|----------|
| `word/document.xml` | body content, tables, pictures, section properties |
| `word/header*.xml`, `word/footer*.xml` | fonts and spacing in headers and footers |
| `docProps/app.xml` | cached page count (read only) |

Elements it changes:

- `w:sz` / `w:szCs` — font size, in half-points
- `w:trHeight` — table row height, in twips (1440 per inch)
- `w:tcMar` — table cell padding
- `w:spacing` — line spacing and paragraph gaps
- `w:pgMar` — page margins
- `wp:extent` and `a:ext` — picture dimensions, in EMU (914400 per inch)
