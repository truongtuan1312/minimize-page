const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WP14 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

let currentFile = null;
let resultBlob = null;
let resultName = '';

const $ = (id) => document.getElementById(id);

const dz = $('dropZone');
dz.addEventListener('click', () => $('fileInput').click());
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
    <div class="file-icon-wrap">
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/>
        <polyline points="14 2 14 8 20 8"/>
      </svg>
    </div>
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
  const sparseLastPage = $('sparseLastPage').checked;

  if (pages % 2 === 0 && !sparseLastPage) {
    $('status').innerHTML = `<div class="result success">
      <div class="result-title">Already even</div>
      <div class="result-detail">${pages} pages is already even — no changes needed. If the last page is mostly empty, enable "Last page is mostly empty" to try removing it.</div>
    </div>`;
    return;
  }

  $('processBtn').disabled = true;
  $('resultArea').innerHTML = '';

  try {
    const buf = await currentFile.arrayBuffer();
    const target = pages - 1;
    const zip = await JSZip.loadAsync(buf);

    $('status').innerHTML = `<div class="processing"><span class="spinner"></span>Analyzing and minimizing&hellip;</div>`;

    const applied = await smartMinimize(zip, pages, target);

    resultBlob = await generateBlob(zip);
    showSmartResult(pages, target, applied);

  } catch (err) {
    $('status').innerHTML = `<div class="result warning">
      <div class="result-title">Processing error</div>
      <div class="result-detail">${esc(err.message)}</div>
    </div>`;
  }
  $('processBtn').disabled = false;
});

async function smartMinimize(zip, currentPages, targetPages) {
  const applied = [];

  // --- 1. Remove empty paragraphs ---
  await transformDocXml(zip, (doc) => {
    const toRemove = [];
    for (const p of doc.getElementsByTagNameNS(W, 'p')) {
      const hasText = Array.from(p.getElementsByTagNameNS(W, 't')).some(t => t.textContent.trim());
      const hasDrawing = p.getElementsByTagNameNS(W, 'drawing').length > 0;
      const hasBreak = p.getElementsByTagNameNS(W, 'br').length > 0;
      let inTable = false;
      let node = p.parentNode;
      while (node) { if (node.localName === 'tc') { inTable = true; break; } node = node.parentNode; }
      const pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
      const hasSectPr = pPr && pPr.getElementsByTagNameNS(W, 'sectPr').length > 0;
      if (!hasText && !hasDrawing && !hasBreak && !inTable && !hasSectPr) toRemove.push(p);
    }
    if (toRemove.length > 0) {
      for (const p of toRemove) p.parentNode.removeChild(p);
      applied.push('Remove empty paragraphs');
    }
  });

  // --- 2. Compact table rows: remove atLeast heights, set to auto-fit ---
  await transformDocXml(zip, (doc) => {
    let changed = false;
    for (const trPr of doc.getElementsByTagNameNS(W, 'trPr')) {
      const trH = trPr.getElementsByTagNameNS(W, 'trHeight')[0];
      if (trH) {
        const rule = trH.getAttributeNS(W, 'hRule');
        const val = parseInt(trH.getAttributeNS(W, 'val') || '0');
        if (rule === 'atLeast' || val > 200) {
          trH.setAttributeNS(W, 'w:val', '0');
          trH.removeAttributeNS(W, 'hRule');
          changed = true;
        }
      }
    }

    // Add compact cell margins (top/bottom = 60 twips)
    for (const tc of doc.getElementsByTagNameNS(W, 'tc')) {
      let tcPr = tc.getElementsByTagNameNS(W, 'tcPr')[0];
      if (!tcPr) continue;
      let tcMar = tcPr.getElementsByTagNameNS(W, 'tcMar')[0];
      if (!tcMar) {
        tcMar = doc.createElementNS(W, 'w:tcMar');
        tcPr.appendChild(tcMar);
      }
      for (const side of ['top', 'bottom']) {
        let el = tcMar.getElementsByTagNameNS(W, side)[0];
        if (!el) {
          el = doc.createElementNS(W, 'w:' + side);
          tcMar.appendChild(el);
        }
        el.setAttributeNS(W, 'w:w', '60');
        el.setAttributeNS(W, 'w:type', 'dxa');
      }
      changed = true;
    }
    if (changed) applied.push('Compact table rows');
  });

  // --- 3. Tighten paragraph spacing ---
  await transformAllXml(zip, (doc) => {
    for (const p of doc.getElementsByTagNameNS(W, 'p')) {
      let pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
      if (!pPr) continue;

      let sp = pPr.getElementsByTagNameNS(W, 'spacing')[0];
      if (!sp) continue;

      // Reduce before-spacing
      const before = sp.getAttributeNS(W, 'before');
      if (before) {
        const val = parseInt(before);
        if (val > 0) {
          sp.setAttributeNS(W, 'w:before', String(Math.round(val * 0.5)));
        }
      }

      // Reduce after-spacing (only if explicitly set and > 0)
      const after = sp.getAttributeNS(W, 'after');
      if (after) {
        const val = parseInt(after);
        if (val > 100) {
          sp.setAttributeNS(W, 'w:after', String(Math.round(val * 0.7)));
        }
      }

      // Reduce line spacing only if above single (240)
      const line = sp.getAttributeNS(W, 'line');
      if (line) {
        const val = parseInt(line);
        if (val > 260) {
          sp.setAttributeNS(W, 'w:line', String(Math.max(240, val - 36)));
          sp.setAttributeNS(W, 'w:lineRule', 'auto');
        }
      }
    }
  });
  applied.push('Tighten paragraph spacing');

  // --- 4. Optimize margins: top → 0 when header exists (biggest win) ---
  await transformDocXml(zip, (doc) => {
    const hasHeader = doc.getElementsByTagNameNS(W, 'headerReference').length > 0;
    const hasFooter = doc.getElementsByTagNameNS(W, 'footerReference').length > 0;
    let changed = false;

    for (const m of doc.getElementsByTagNameNS(W, 'pgMar')) {
      // When there's a header, set top margin to 0 and header distance to 720
      // The header image/content provides visual spacing
      if (hasHeader) {
        const top = parseInt(m.getAttributeNS(W, 'top') || '0');
        if (top > 0) {
          m.setAttributeNS(W, 'w:top', '0');
          m.setAttributeNS(W, 'w:header', '720');
          changed = true;
        }
      }

      // Bottom margin: keep reasonable but can tighten
      if (!hasFooter) {
        const bot = parseInt(m.getAttributeNS(W, 'bottom') || '0');
        if (bot > 450) {
          m.setAttributeNS(W, 'w:bottom', '450');
          changed = true;
        }
      }
    }
    if (changed) applied.push('Optimize margins');
  });

  // --- 5. Scale images based on page reduction ratio ---
  {
    const pageRatio = targetPages / currentPages;
    const imgScale = Math.max(0.5, Math.min(0.95, pageRatio));

    const docXml = await zip.file('word/document.xml').async('string');
    const doc = new DOMParser().parseFromString(docXml, 'application/xml');
    let hasImages = false;

    for (const ns of [WP, WP14]) {
      for (const ext of doc.getElementsByTagNameNS(ns, 'extent')) {
        const cx = ext.getAttribute('cx');
        const cy = ext.getAttribute('cy');
        if (cx) { ext.setAttribute('cx', String(Math.round(parseInt(cx) * imgScale))); hasImages = true; }
        if (cy) ext.setAttribute('cy', String(Math.round(parseInt(cy) * imgScale)));
      }
    }

    for (const xfrm of doc.getElementsByTagNameNS(A, 'xfrm')) {
      const ext = xfrm.getElementsByTagNameNS(A, 'ext')[0];
      if (ext) {
        const cx = ext.getAttribute('cx');
        const cy = ext.getAttribute('cy');
        if (cx) ext.setAttribute('cx', String(Math.round(parseInt(cx) * imgScale)));
        if (cy) ext.setAttribute('cy', String(Math.round(parseInt(cy) * imgScale)));
      }
    }

    zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
    if (hasImages) applied.push(`Scale images (${Math.round((1 - imgScale) * 100)}%)`);
  }

  return applied;
}

// --- Helpers ---

async function transformDocXml(zip, fn) {
  const xml = await zip.file('word/document.xml').async('string');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  fn(doc);
  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
}

async function transformAllXml(zip, fn) {
  for (const name of Object.keys(zip.files)) {
    if (!/^word\/(document|header\d*|footer\d*)\.xml$/.test(name)) continue;
    const xml = await zip.file(name).async('string');
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    fn(doc);
    zip.file(name, new XMLSerializer().serializeToString(doc));
  }
}

async function generateBlob(zip) {
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

function showSmartResult(original, target, applied) {
  const base = currentFile.name.replace(/\.docx$/i, '');
  resultName = base + '_minimized.docx';

  const expandTarget = original % 2 === 0 ? original : original + 1;
  const techList = applied.map(t => `<span class="tech-badge">${esc(t)}</span>`).join(' ');

  $('status').innerHTML = `<div class="result success">
    <div class="result-title">Minimized — ${applied.length} technique${applied.length > 1 ? 's' : ''} applied</div>
    <div class="result-pages">
      <span class="page-badge">${original} pp</span>
      <span class="arrow">&rarr;</span>
      <span class="page-badge">${target} pp (target)</span>
    </div>
    <div class="result-detail">Open the file in Word to verify — all content is preserved.</div>
    <div class="tech-list">${techList}</div>
  </div>`;
  $('resultArea').innerHTML = `<button class="btn-download" id="downloadBtn">&#8595; Download ${esc(resultName)}</button>`;
  $('downloadBtn').addEventListener('click', doDownload);
}

async function expandToEven(original, expandTarget) {
  $('processBtn').disabled = true;
  $('resultArea').innerHTML = '';
  $('status').innerHTML = `<div class="processing"><span class="spinner"></span>Expanding to ${expandTarget} pages&hellip;</div>`;

  try {
    const buf = await currentFile.arrayBuffer();
    const zip = await JSZip.loadAsync(buf);

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
        sp.setAttributeNS(W, 'w:line', '300');
        sp.setAttributeNS(W, 'w:lineRule', 'auto');

        const before = sp.getAttributeNS(W, 'before');
        if (before && parseInt(before) > 0) {
          sp.setAttributeNS(W, 'w:before', String(Math.round(parseInt(before) * 1.3)));
        }
        const after = sp.getAttributeNS(W, 'after');
        if (after && parseInt(after) > 0) {
          sp.setAttributeNS(W, 'w:after', String(Math.round(parseInt(after) * 1.3)));
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
