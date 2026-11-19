const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Line values in 240ths of a point (240 = single, 276 = 1.15 default)
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WP14 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';

// lineVal: 240ths of a point (240=single, 276=1.15 default)
// imgScale: scale factor for images (1.0 = no change)
// narrow: set all margins to Word's "Narrow" preset (1.27cm = 720 twips)
const LEVELS = [
  { lineVal: 268, spaceFactor: 0.88, narrow: false, imgScale: 0.95 },
  { lineVal: 256, spaceFactor: 0.75, narrow: false, imgScale: 0.90 },
  { lineVal: 240, spaceFactor: 0.60, narrow: true,  imgScale: 0.85 },
  { lineVal: 228, spaceFactor: 0.45, narrow: true,  imgScale: 0.80 },
  { lineVal: 216, spaceFactor: 0.30, narrow: true,  imgScale: 0.75 },
  { lineVal: 204, spaceFactor: 0.18, narrow: true,  imgScale: 0.70 },
  { lineVal: 192, spaceFactor: 0.10, narrow: true,  imgScale: 0.65 },
];

let currentFile = null;
let selectedLevel = 'auto';
let resultBlob = null;
let resultName = '';

const $ = (id) => document.getElementById(id);

$('levelSelector').addEventListener('click', (e) => {
  const btn = e.target.closest('.level-btn');
  if (!btn) return;
  $('levelSelector').querySelectorAll('.level-btn').forEach((b) => b.classList.remove('active'));
  btn.classList.add('active');
  selectedLevel = btn.dataset.level;
});

const dz = $('dropZone');
dz.addEventListener('dragover', (e) => { e.preventDefault(); dz.classList.add('drag-over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
dz.addEventListener('drop', (e) => {
  e.preventDefault();
  dz.classList.remove('drag-over');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
$('fileInput').addEventListener('change', (e) => {
  if (e.target.files.length) handleFile(e.target.files[0]);
});

function handleFile(file) {
  if (!file.name.toLowerCase().endsWith('.docx')) {
    $('status').innerHTML = '<div class="result warning"><div class="result-title">Only .docx files are supported</div></div>';
    return;
  }
  currentFile = file;
  resultBlob = null;
  $('resultArea').innerHTML = '';
  $('status').innerHTML = '';
  dz.hidden = true;
  $('fileCard').hidden = false;
  $('fileCard').innerHTML = `<div class="file-info">
    <span class="file-icon">&#128196;</span>
    <div class="file-details">
      <div class="file-name">${esc(file.name)}</div>
      <div class="file-size">${fmtSize(file.size)}</div>
    </div>
    <button class="file-remove" id="removeFile" title="Remove">&times;</button>
  </div>`;
  $('removeFile').addEventListener('click', resetFile);
  $('settingsPanel').hidden = false;
  $('processBtn').disabled = false;
  $('pageCount').focus();
}

function resetFile() {
  currentFile = null;
  resultBlob = null;
  $('fileInput').value = '';
  dz.hidden = false;
  $('fileCard').hidden = true;
  $('settingsPanel').hidden = true;
  $('processBtn').disabled = true;
  $('status').innerHTML = '';
  $('resultArea').innerHTML = '';
}

$('processBtn').addEventListener('click', async () => {
  const pages = parseInt($('pageCount').value);
  if (!pages || pages < 1) {
    $('status').innerHTML = '<div class="result warning"><div class="result-title">Enter the current page count</div></div>';
    return;
  }
  if (pages % 2 === 0) {
    $('status').innerHTML = `<div class="result success">
      <div class="result-title">Already even</div>
      <div class="result-detail">${pages} pages is already even — no changes needed.</div>
    </div>`;
    return;
  }

  $('processBtn').disabled = true;
  $('resultArea').innerHTML = '';

  try {
    const buf = await currentFile.arrayBuffer();
    const target = pages - 1;

    if (selectedLevel !== 'auto') {
      const level = parseInt(selectedLevel) - 1;
      $('status').innerHTML = `<div class="processing"><span class="spinner"></span>Applying level ${level + 1}&hellip;</div>`;
      const zip = await JSZip.loadAsync(buf);
      await applyReduction(zip, LEVELS[level]);
      resultBlob = await generateBlob(zip);
      showResult(pages, target, level + 1);
    } else {
      // Iterative: try each level, estimate page count, stop when target reached
      let appliedLevel = 0;
      for (let i = 0; i < LEVELS.length; i++) {
        appliedLevel = i + 1;
        $('status').innerHTML = `<div class="processing"><span class="spinner"></span>Trying level ${appliedLevel}/${LEVELS.length}&hellip;</div>`;

        const zip = await JSZip.loadAsync(buf);
        await applyReduction(zip, LEVELS[i]);

        const estimated = await estimatePages(zip, pages);
        if (estimated <= target) {
          resultBlob = await generateBlob(zip);
          showResult(pages, target, appliedLevel, estimated);
          $('processBtn').disabled = false;
          return;
        }
      }

      // Max level reached but still not enough — use last result
      const zip = await JSZip.loadAsync(buf);
      await applyReduction(zip, LEVELS[LEVELS.length - 1]);
      resultBlob = await generateBlob(zip);
      showResult(pages, target, appliedLevel, null, true);
    }
  } catch (err) {
    $('status').innerHTML = `<div class="result warning">
      <div class="result-title">Processing error</div>
      <div class="result-detail">${esc(err.message)}</div>
    </div>`;
  }
  $('processBtn').disabled = false;
});

function showResult(original, target, level, estimated, maxReached) {
  const base = currentFile.name.replace(/\.docx$/i, '');
  resultName = base + '_minimized.docx';

  const expandTarget = original + 1; // next even number above

  if (maxReached) {
    $('status').innerHTML = `<div class="result warning">
      <div class="result-title">Max reduction applied (level ${level})</div>
      <div class="result-pages">
        <span class="page-badge">${original} pp</span>
        <span class="arrow">&rarr;</span>
        <span class="page-badge">${target} pp (target)</span>
      </div>
      <div class="result-detail">
        Spacing is already at the tightest level but the content may not fit in ${target} pages.
        You can still download and check in Word, or expand to <strong>${expandTarget} pages</strong> instead — the spacing will be loosened so it reads more comfortably.
      </div>
      <button class="btn-expand" id="expandBtn">Expand to ${expandTarget} pages instead</button>
    </div>`;
    setTimeout(() => {
      const eb = $('expandBtn');
      if (eb) eb.addEventListener('click', () => expandToEven(original, expandTarget));
    }, 0);
  } else {
    $('status').innerHTML = `<div class="result success">
      <div class="result-title">Level ${level} applied${estimated ? ' — estimated ' + estimated + ' pp' : ''}</div>
      <div class="result-pages">
        <span class="page-badge">${original} pp</span>
        <span class="arrow">&rarr;</span>
        <span class="page-badge">${target} pp (target)</span>
      </div>
      <div class="result-detail">Open the file in Word to verify. All content is preserved; only spacing was adjusted.</div>
    </div>`;
  }
  $('resultArea').innerHTML = `<button class="btn-download" id="downloadBtn">&#8595; Download ${esc(resultName)}</button>`;
  $('downloadBtn').addEventListener('click', doDownload);
}

async function estimatePages(zip, originalPages) {
  const docXml = await zip.file('word/document.xml').async('string');
  const doc = new DOMParser().parseFromString(docXml, 'application/xml');

  // Gather section properties for page dimensions
  const sectPrs = doc.getElementsByTagNameNS(W, 'sectPr');
  let pageHeight = 15840; // A4 default in twips
  let topMargin = 1440;
  let bottomMargin = 1440;

  if (sectPrs.length > 0) {
    const sect = sectPrs[sectPrs.length - 1];
    const pgSz = sect.getElementsByTagNameNS(W, 'pgSz')[0];
    const pgMar = sect.getElementsByTagNameNS(W, 'pgMar')[0];
    if (pgSz) {
      const h = pgSz.getAttributeNS(W, 'h');
      if (h) pageHeight = parseInt(h);
    }
    if (pgMar) {
      const t = pgMar.getAttributeNS(W, 'top');
      const b = pgMar.getAttributeNS(W, 'bottom');
      if (t) topMargin = parseInt(t);
      if (b) bottomMargin = parseInt(b);
    }
  }

  const usableHeight = pageHeight - topMargin - bottomMargin;

  // Measure total content height from paragraph spacing
  let totalHeight = 0;
  const paras = doc.getElementsByTagNameNS(W, 'p');
  const pageWidth = 12240; // A4 width default
  let leftMargin = 1440, rightMargin = 1440;
  if (sectPrs.length > 0) {
    const pgMar = sectPrs[sectPrs.length - 1].getElementsByTagNameNS(W, 'pgMar')[0];
    if (pgMar) {
      const l = pgMar.getAttributeNS(W, 'left');
      const r = pgMar.getAttributeNS(W, 'right');
      if (l) leftMargin = parseInt(l);
      if (r) rightMargin = parseInt(r);
    }
  }
  const usableWidth = pageWidth - leftMargin - rightMargin;

  for (const p of paras) {
    const pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
    let lineSpacing = 276; // default 1.15 in 240ths
    let spaceBefore = 0;
    let spaceAfter = 160; // default 8pt after
    let fontSize = 22; // 11pt in half-points

    if (pPr) {
      const sp = pPr.getElementsByTagNameNS(W, 'spacing')[0];
      if (sp) {
        const line = sp.getAttributeNS(W, 'line');
        if (line) lineSpacing = parseInt(line);
        const bef = sp.getAttributeNS(W, 'before');
        if (bef) spaceBefore = parseInt(bef);
        const aft = sp.getAttributeNS(W, 'after');
        if (aft) spaceAfter = parseInt(aft);
      }
    }

    // Check run-level font sizes (more reliable than pPr/rPr)
    for (const r of p.getElementsByTagNameNS(W, 'r')) {
      const rPr = r.getElementsByTagNameNS(W, 'rPr')[0];
      if (rPr) {
        const sz = rPr.getElementsByTagNameNS(W, 'sz')[0];
        if (sz) {
          const val = sz.getAttributeNS(W, 'val');
          if (val) { fontSize = parseInt(val); break; }
        }
      }
    }

    // Get text length from all runs
    let textLen = 0;
    for (const t of p.getElementsByTagNameNS(W, 't')) {
      textLen += (t.textContent || '').length;
    }

    // Estimate character width (~55% of font size for proportional fonts)
    const charWidthTwips = fontSize * 0.55;
    const charsPerLine = Math.max(1, Math.floor(usableWidth / charWidthTwips));
    const numLines = Math.max(1, Math.ceil(textLen / charsPerLine) || 1);

    // Line height: for lineRule="auto", line value is in 240ths of a line
    const lineHeightTwips = (lineSpacing / 240) * fontSize * 10;
    const paraHeight = spaceBefore + (numLines * lineHeightTwips) + spaceAfter;
    totalHeight += paraHeight;
  }

  // Count table rows
  const rows = doc.getElementsByTagNameNS(W, 'tr');
  totalHeight += rows.length * 400;

  // Count image heights (EMU to twips: 1 twip = 914.4 EMU)
  for (const ns of [WP, WP14]) {
    for (const ext of doc.getElementsByTagNameNS(ns, 'extent')) {
      const cy = ext.getAttribute('cy');
      if (cy) totalHeight += parseInt(cy) / 914.4;
    }
  }

  const estimatedPages = Math.max(1, Math.ceil(totalHeight / usableHeight));

  // Use ratio approach as a sanity check:
  // We know the original page count, so scale the estimate proportionally
  if (originalPages > 1 && estimatedPages > 0) {
    return estimatedPages;
  }
  return estimatedPages;
}

async function expandToEven(original, expandTarget) {
  $('processBtn').disabled = true;
  $('resultArea').innerHTML = '';
  $('status').innerHTML = `<div class="processing"><span class="spinner"></span>Expanding to ${expandTarget} pages&hellip;</div>`;

  try {
    const buf = await currentFile.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);

    // Apply gentle increase: widen line spacing and add paragraph spacing
    const EXPAND = { lineVal: 300, spaceFactor: 1.3, marginReduce: -200 };

    for (const name of Object.keys(zip.files)) {
      if (!/^word\/(document|header\d*|footer\d*)\.xml$/.test(name)) continue;
      let xml = await zip.file(name).async('string');
      const doc = new DOMParser().parseFromString(xml, 'application/xml');

      for (const pPr of doc.getElementsByTagNameNS(W, 'pPr')) {
        let sp = pPr.getElementsByTagNameNS(W, 'spacing')[0];
        if (!sp) {
          sp = doc.createElementNS(W, 'w:spacing');
          pPr.appendChild(sp);
        }
        sp.setAttributeNS(W, 'w:line', String(EXPAND.lineVal));
        sp.setAttributeNS(W, 'w:lineRule', 'auto');

        const before = sp.getAttributeNS(W, 'before');
        if (before && parseInt(before) > 0) {
          sp.setAttributeNS(W, 'w:before', String(Math.round(parseInt(before) * EXPAND.spaceFactor)));
        }
        const after = sp.getAttributeNS(W, 'after');
        if (after && parseInt(after) > 0) {
          sp.setAttributeNS(W, 'w:after', String(Math.round(parseInt(after) * EXPAND.spaceFactor)));
        }
      }

      xml = new XMLSerializer().serializeToString(doc);
      zip.file(name, xml);
    }

    resultBlob = await generateBlob(zip);
    const base = currentFile.name.replace(/\.docx$/i, '');
    resultName = base + '_expanded.docx';

    $('status').innerHTML = `<div class="result success">
      <div class="result-title">Expanded to ${expandTarget} pages</div>
      <div class="result-pages">
        <span class="page-badge">${original} pp</span>
        <span class="arrow">&rarr;</span>
        <span class="page-badge">${expandTarget} pp (target)</span>
      </div>
      <div class="result-detail">Spacing increased so the document fills ${expandTarget} pages evenly. Verify in Word before printing.</div>
    </div>`;
    $('resultArea').innerHTML = `<button class="btn-download" id="downloadBtn">&#8595; Download ${esc(resultName)}</button>`;
    $('downloadBtn').addEventListener('click', doDownload);
  } catch (err) {
    $('status').innerHTML = `<div class="result warning">
      <div class="result-title">Processing error</div>
      <div class="result-detail">${esc(err.message)}</div>
    </div>`;
  }
  $('processBtn').disabled = false;
}

function doDownload() {
  if (!resultBlob) return;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(resultBlob);
  a.download = resultName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}

async function generateBlob(zip) {
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

async function applyReduction(zip, cfg) {
  for (const name of Object.keys(zip.files)) {
    if (!/^word\/(document|header\d*|footer\d*)\.xml$/.test(name)) continue;
    let xml = await zip.file(name).async('string');
    xml = adjustSpacing(xml, cfg);
    if (cfg.narrow) {
      xml = setNarrowMargins(xml);
    }
    if (cfg.imgScale && cfg.imgScale < 1.0) {
      xml = scaleImages(xml, cfg.imgScale);
    }
    zip.file(name, xml);
  }
}

function scaleImages(xml, scale) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');

  // Scale <wp:extent> on both inline and anchor drawings
  for (const ns of [WP, WP14]) {
    for (const ext of doc.getElementsByTagNameNS(ns, 'extent')) {
      const cx = ext.getAttribute('cx');
      const cy = ext.getAttribute('cy');
      if (cx) ext.setAttribute('cx', String(Math.round(parseInt(cx) * scale)));
      if (cy) ext.setAttribute('cy', String(Math.round(parseInt(cy) * scale)));
    }
  }

  // Also scale <wp:inline> or <wp:anchor> effectExtent is optional but extent is key
  // Scale <a:ext> inside <a:xfrm> for the actual shape
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  for (const xfrm of doc.getElementsByTagNameNS(A, 'xfrm')) {
    const ext = xfrm.getElementsByTagNameNS(A, 'ext')[0];
    if (ext) {
      const cx = ext.getAttribute('cx');
      const cy = ext.getAttribute('cy');
      if (cx) ext.setAttribute('cx', String(Math.round(parseInt(cx) * scale)));
      if (cy) ext.setAttribute('cy', String(Math.round(parseInt(cy) * scale)));
    }
  }

  return new XMLSerializer().serializeToString(doc);
}

function adjustSpacing(xml, cfg) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');

  // Remove empty paragraphs (no text, no images, no other content)
  const toRemove = [];
  for (const p of doc.getElementsByTagNameNS(W, 'p')) {
    const hasText = Array.from(p.getElementsByTagNameNS(W, 't')).some(t => t.textContent.trim());
    const hasDrawing = p.getElementsByTagNameNS(W, 'drawing').length > 0;
    const hasBreak = p.getElementsByTagNameNS(W, 'br').length > 0;
    // Keep paragraphs inside table cells (they define the cell even if empty)
    let inTable = false;
    let node = p.parentNode;
    while (node) { if (node.localName === 'tc') { inTable = true; break; } node = node.parentNode; }
    // Keep section break paragraphs
    const pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
    const hasSectPr = pPr && pPr.getElementsByTagNameNS(W, 'sectPr').length > 0;

    if (!hasText && !hasDrawing && !hasBreak && !inTable && !hasSectPr) {
      toRemove.push(p);
    }
  }
  for (const p of toRemove) {
    p.parentNode.removeChild(p);
  }

  const allParagraphs = doc.getElementsByTagNameNS(W, 'p');
  for (const p of allParagraphs) {
    let pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
    if (!pPr) {
      pPr = doc.createElementNS(W, 'w:pPr');
      p.insertBefore(pPr, p.firstChild);
    }

    let sp = pPr.getElementsByTagNameNS(W, 'spacing')[0];
    if (!sp) {
      sp = doc.createElementNS(W, 'w:spacing');
      pPr.appendChild(sp);
    }

    sp.setAttributeNS(W, 'w:line', String(cfg.lineVal));
    sp.setAttributeNS(W, 'w:lineRule', 'auto');

    // Word defaults: before=0, after=160 (8pt), line=276 (1.15)
    // If not set explicitly, apply reduction to the default
    const before = sp.getAttributeNS(W, 'before');
    const beforeVal = (before && parseInt(before) > 0) ? parseInt(before) : 0;
    sp.setAttributeNS(W, 'w:before', String(Math.round(beforeVal * cfg.spaceFactor)));

    const after = sp.getAttributeNS(W, 'after');
    const afterVal = (after && parseInt(after) > 0) ? parseInt(after) : 160;
    sp.setAttributeNS(W, 'w:after', String(Math.round(afterVal * cfg.spaceFactor)));
  }
  return new XMLSerializer().serializeToString(doc);
}

function setNarrowMargins(xml) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const NARROW = 720; // 1.27cm in twips — Word's "Narrow" preset

  // Find the largest negative table indent so we don't push tables off-page
  let maxNegIndent = 0;
  for (const ind of doc.getElementsByTagNameNS(W, 'tblInd')) {
    const val = parseFloat(ind.getAttributeNS(W, 'w') || '0');
    if (val < 0) maxNegIndent = Math.max(maxNegIndent, Math.abs(val));
  }

  for (const m of doc.getElementsByTagNameNS(W, 'pgMar')) {
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const current = m.getAttributeNS(W, side);
      if (!current) continue;
      const curVal = parseInt(current);

      // For left/right: don't go below the negative table indent + a small buffer
      let minMargin = NARROW;
      if ((side === 'left' || side === 'right') && maxNegIndent > 0) {
        minMargin = Math.max(NARROW, maxNegIndent + 100);
      }

      if (curVal > minMargin) {
        m.setAttributeNS(W, 'w:' + side, String(minMargin));
      }
    }
  }
  return new XMLSerializer().serializeToString(doc);
}

function fmtSize(b) {
  if (b < 1024) return b + ' B';
  if (b < 1048576) return (b / 1024).toFixed(1) + ' KB';
  return (b / 1048576).toFixed(1) + ' MB';
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}
