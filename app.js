const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WP14 = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const EMU_PER_IN = 914400;
const TWIP_PER_IN = 1440;

const DEFAULTS = { maxFont: 18, imgRowH: 0.75, textRowH: 0.2 };

let currentFile = null;
let scan = null;
let detected = null;   // { pages, source: 'word' | 'estimate' }
let target = null;     // pages the user wants
let resultBlob = null;
let resultName = '';

const $ = (id) => document.getElementById(id);

// --- File intake ---

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

async function handleFile(file) {
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
  $('processBtn').disabled = false;

  try {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    scan = await analyzeDocx(zip);
    detected = await detectPages(zip, scan);
    target = defaultTarget(detected.pages);
    $('pageCount').value = detected.pages;
    renderScan(scan);
  } catch (err) {
    scan = null;
    detected = null;
    target = null;
    $('scanPanel').hidden = true;
    $('status').innerHTML = `<div class="result warning">
      <div class="result-title">Couldn't read this file</div>
      <div class="result-detail">${esc(err.message)}</div>
    </div>`;
  }
  refreshPages();
  refreshEstimate();
}

// Word stores the page count from its last layout in docProps/app.xml.
// That is exact when present; otherwise fall back to measuring content height.
async function detectPages(zip, s) {
  const part = zip.file('docProps/app.xml');
  if (part) {
    const xml = await part.async('string');
    const m = xml.match(/<Pages>(\d+)<\/Pages>/);
    if (m && parseInt(m[1]) > 0) return { pages: parseInt(m[1]), source: 'word' };
  }
  return { pages: estimatePages(s), source: 'estimate' };
}

function estimatePages(s) {
  const usable = Math.max(1, s.pageHeightIn - s.topMarginIn - s.bottomMarginIn);
  let content = 0;
  for (const h of s.rowTallestIn) content += Math.max(h, 0.2);
  content += s.textRows * 0.26;
  content += s.bodyParas * (s.maxFontPt || 12) / 72 * 1.45;
  return Math.max(1, Math.ceil(content / usable));
}

function defaultTarget(pages) {
  return Math.max(1, pages - 1);
}

function resetFile() {
  currentFile = null;
  scan = null;
  resultBlob = null;
  $('fileInput').value = '';
  dz.hidden = false;
  $('fileCard').hidden = true;
  $('scanPanel').hidden = true;
  $('processBtn').disabled = true;
  $('status').innerHTML = '';
  $('resultArea').innerHTML = '';
  refreshEstimate();
}

// --- Document analysis ---

async function analyzeDocx(zip) {
  const doc = await readXml(zip, 'word/document.xml');
  const s = {
    images: 0, imgRows: 0, textRows: 0, tables: 0,
    emptyParas: 0, maxFontPt: 0, tallestImgIn: 0, looseImages: 0,
    imgHeightsIn: [],
    rowTallestIn: [],
    hasHeader: doc.getElementsByTagNameNS(W, 'headerReference').length > 0,
    topMarginIn: 0, bottomMarginIn: 0, pageHeightIn: 11, bodyParas: 0,
  };

  s.tables = doc.getElementsByTagNameNS(W, 'tbl').length;

  for (const tr of doc.getElementsByTagNameNS(W, 'tr')) {
    if (tr.getElementsByTagNameNS(W, 'drawing').length === 0) { s.textRows++; continue; }
    s.imgRows++;
    // Pictures in one row sit side by side, so only the tallest sets row height.
    let tallest = 0;
    for (const ext of allExtents(tr)) {
      const h = parseInt(ext.getAttribute('cy') || '0') / EMU_PER_IN;
      if (h > tallest) tallest = h;
    }
    s.rowTallestIn.push(tallest);
  }

  for (const ext of allExtents(doc)) {
    s.images++;
    const inches = parseInt(ext.getAttribute('cy') || '0') / EMU_PER_IN;
    s.imgHeightsIn.push(inches);
    if (inches > s.tallestImgIn) s.tallestImgIn = inches;
    if (!insideTable(ext)) s.looseImages++;
  }

  for (const p of doc.getElementsByTagNameNS(W, 'p')) {
    if (isEmptyPara(p)) s.emptyParas++;
    else if (!insideTable(p)) s.bodyParas++;
  }

  const pgMar = doc.getElementsByTagNameNS(W, 'pgMar')[0];
  if (pgMar) {
    s.topMarginIn = parseInt(pgMar.getAttributeNS(W, 'top') || '0') / TWIP_PER_IN;
    s.bottomMarginIn = parseInt(pgMar.getAttributeNS(W, 'bottom') || '0') / TWIP_PER_IN;
  }

  const pgSz = doc.getElementsByTagNameNS(W, 'pgSz')[0];
  if (pgSz) {
    const h = parseInt(pgSz.getAttributeNS(W, 'h') || '0');
    if (h > 0) s.pageHeightIn = h / TWIP_PER_IN;
  }

  // Largest font across body, headers and footers
  for (const name of partNames(zip)) {
    const part = await readXml(zip, name);
    for (const sz of fontEls(part)) {
      const hp = parseInt(sz.getAttributeNS(W, 'val') || '0');
      if (hp / 2 > s.maxFontPt) s.maxFontPt = hp / 2;
    }
  }

  return s;
}

function renderScan(s) {
  const r = getRules();
  const over = s.tallestImgIn > r.imgRowH;
  const bigFont = s.maxFontPt > r.maxFont;

  const rows = [
    ['Pictures', `${s.images}`, s.images ? '' : 'ok'],
    ['Tallest picture', s.tallestImgIn ? `${s.tallestImgIn.toFixed(2)} in` : '—', over ? 'flag' : 'ok'],
    ['Picture rows', `${s.imgRows}`, ''],
    ['Text rows', `${s.textRows}`, ''],
    ['Largest font', s.maxFontPt ? `${trimNum(s.maxFontPt)} pt` : '—', bigFont ? 'flag' : 'ok'],
    ['Blank lines', `${s.emptyParas}`, s.emptyParas ? 'flag' : 'ok'],
    ['Top margin', `${s.topMarginIn.toFixed(2)} in`, ''],
  ];

  $('scanGrid').innerHTML = rows
    .map(([k, v, cls]) => `<div><dt>${k}</dt><dd class="${cls}">${v}</dd></div>`)
    .join('');

  const tooTall = s.imgHeightsIn.filter(h => h > r.imgRowH).length;
  const notes = [];
  if (over) notes.push(`${tooTall} of ${s.images} picture${s.images > 1 ? 's' : ''} exceed ${trimNum(r.imgRowH)} in and will be shrunk to fit.`);
  if (bigFont) notes.push(`Fonts scale down by ${Math.round((1 - r.maxFont / s.maxFontPt) * 100)}%.`);
  if (!over && !bigFont) notes.push('Nothing exceeds your limits — gains will come from spacing and margins.');
  $('scanNote').textContent = notes.join(' ');
  $('scanPanel').hidden = false;
}

// --- Rules ---

function getRules() {
  return {
    fonts: $('ruleFonts').checked,
    imgRow: $('ruleImgRow').checked,
    textRow: $('ruleTextRow').checked,
    spacing: $('ruleSpacing').checked,
    margins: $('ruleMargins').checked,
    empties: $('ruleEmpties').checked,
    freeImg: $('ruleFreeImg').checked,
    maxFont: clampNum($('maxFont').value, 8, 72, DEFAULTS.maxFont),
    imgRowH: clampNum($('imgRowH').value, 0.2, 4, DEFAULTS.imgRowH),
    textRowH: clampNum($('textRowH').value, 0.1, 2, DEFAULTS.textRowH),
  };
}

for (const id of ['ruleFonts', 'ruleImgRow', 'ruleTextRow', 'ruleSpacing', 'ruleMargins', 'ruleEmpties', 'ruleFreeImg']) {
  $(id).addEventListener('change', onRulesChanged);
}
for (const id of ['maxFont', 'imgRowH', 'textRowH']) {
  $(id).addEventListener('input', onRulesChanged);
}
$('tgtDown').addEventListener('click', () => nudgeTarget(-1));
$('tgtUp').addEventListener('click', () => nudgeTarget(1));

$('pageCount').addEventListener('input', () => {
  const v = parseInt($('pageCount').value);
  if (v > 0) {
    detected = { pages: v, source: 'manual' };
    target = Math.min(target ?? v, 99);
    refreshPages();
    refreshEstimate();
  }
});

function nudgeTarget(delta) {
  if (target == null) return;
  target = Math.min(99, Math.max(1, target + delta));
  refreshPages();
  refreshEstimate();
}

function direction() {
  if (!detected || target == null) return 'none';
  if (target < detected.pages) return 'shrink';
  if (target > detected.pages) return 'grow';
  return 'none';
}

$('resetRules').addEventListener('click', () => {
  for (const id of ['ruleFonts', 'ruleImgRow', 'ruleTextRow', 'ruleSpacing', 'ruleMargins', 'ruleEmpties', 'ruleFreeImg']) {
    $(id).checked = true;
  }
  $('maxFont').value = DEFAULTS.maxFont;
  $('imgRowH').value = DEFAULTS.imgRowH;
  $('textRowH').value = DEFAULTS.textRowH;
  onRulesChanged();
});

function onRulesChanged() {
  for (const el of document.querySelectorAll('.rule')) {
    const box = el.querySelector('input[type=checkbox]');
    el.classList.toggle('off', !box.checked);
  }
  if (scan) renderScan(scan);
  refreshEstimate();
}

function refreshPages() {
  const det = $('pageDetected');
  const detHint = $('detectedHint');
  const tgt = $('pageTarget');
  const tgtHint = $('targetHint');
  const mode = $('modeLine');
  const label = $('btnLabel');

  if (!detected) {
    det.textContent = '—';
    det.classList.add('none');
    detHint.textContent = 'Load a file';
    tgt.textContent = '—';
    tgt.classList.add('none');
    tgtHint.textContent = 'pages';
    mode.hidden = true;
    $('manualRow').hidden = true;
    $('tgtUp').disabled = $('tgtDown').disabled = true;
    label.textContent = 'Minimize';
    $('processBtn').disabled = true;
    return;
  }

  det.classList.remove('none');
  tgt.classList.remove('none');
  det.textContent = detected.pages;
  tgt.textContent = target;
  tgtHint.textContent = target % 2 === 0 ? 'pages · good for duplex' : target === 1 ? 'page' : 'pages · odd';
  tgtHint.classList.toggle('odd', target % 2 !== 0 && target !== 1);
  $('tgtDown').disabled = target <= 1;
  $('tgtUp').disabled = target >= 99;
  if (detected.source === 'word') {
    detHint.innerHTML = 'read from the file &middot; <button type="button" class="link-btn tiny" id="fixCount">wrong?</button>';
    $('fixCount').addEventListener('click', () => {
      $('manualRow').hidden = false;
      $('pageCount').focus();
      $('pageCount').select();
    });
  } else {
    detHint.textContent = detected.source === 'manual' ? 'set by you' : 'estimated — correct it below';
    $('manualRow').hidden = false;
  }

  const dir = direction();
  const gap = Math.abs(target - detected.pages);
  mode.hidden = false;
  mode.classList.toggle('warn', dir === 'none');

  if (dir === 'shrink') {
    mode.innerHTML = `Squeezing <b>${gap} page${gap > 1 ? 's' : ''}</b> out — rules below do the work.`;
    label.textContent = `Minimize to ${target} page${target > 1 ? 's' : ''}`;
    $('processBtn').disabled = false;
  } else if (dir === 'grow') {
    mode.innerHTML = `Stretching to fill <b>${gap} more page${gap > 1 ? 's' : ''}</b> by opening up spacing.`;
    label.textContent = `Expand to ${target} pages`;
    $('processBtn').disabled = false;
  } else {
    mode.innerHTML = `Already ${target} page${target > 1 ? 's' : ''} — pick a different target.`;
    label.textContent = 'Nothing to change';
    $('processBtn').disabled = true;
  }
}

// Rough estimate of reclaimed vertical space, from the rules that are on
// and what the scan actually measured. Not a page count — Word does the layout.
function refreshEstimate() {
  const fill = $('effortFill');
  const pct = $('effortPct');
  const note = $('effortNote');

  if (!scan) {
    fill.style.width = '0%';
    pct.textContent = '—';
    note.textContent = 'Load a file to see an estimate.';
    return;
  }

  const dir = direction();
  $('rulesWrap').classList.toggle('off', dir === 'grow');

  if (dir !== 'shrink') {
    fill.style.width = '0%';
    pct.textContent = '—';
    note.textContent = dir === 'grow'
      ? 'Expanding instead — the shrink rules above are not used.'
      : 'Pick a smaller target to see what can be reclaimed.';
    return;
  }

  const r = getRules();
  let saved = 0;
  const reasons = [];

  if (r.imgRow && scan.imgRows) {
    const cut = scan.rowTallestIn
      .filter(h => h > r.imgRowH)
      .reduce((sum, h) => sum + (h - r.imgRowH), 0);
    if (cut > 0) { saved += cut; reasons.push('picture rows'); }
  }
  if (r.textRow && scan.textRows) {
    saved += scan.textRows * 0.05;
    reasons.push('text rows');
  }
  if (r.fonts && scan.maxFontPt > r.maxFont) {
    saved += (scan.maxFontPt - r.maxFont) / 72 * (scan.textRows + 4);
    reasons.push('fonts');
  }
  if (r.empties && scan.emptyParas) {
    saved += scan.emptyParas * 0.2;
    reasons.push('blank lines');
  }
  if (r.spacing) { saved += 0.35; reasons.push('spacing'); }
  if (r.margins && scan.hasHeader && scan.topMarginIn > 0) {
    saved += scan.topMarginIn;
    reasons.push('margins');
  }
  if (r.freeImg && scan.looseImages) { saved += 0.3; reasons.push('loose pictures'); }

  // How much height one page actually holds, so progress is measured
  // against the real goal rather than a guessed page size.
  const usable = Math.max(1, scan.pageHeightIn - scan.topMarginIn - scan.bottomMarginIn);
  const needed = (detected.pages - target) * usable;
  const share = Math.min(1, saved / needed);

  fill.style.width = `${Math.round(share * 100)}%`;
  pct.textContent = `${Math.round(share * 100)}%`;

  if (!reasons.length) {
    note.textContent = 'Every rule is off — the file would come back unchanged.';
  } else if (share >= 1) {
    note.textContent = `About ${saved.toFixed(1)} in reclaimed from ${listify(reasons)} — enough for ${needed.toFixed(1)} in needed. Verify in Word.`;
  } else {
    note.textContent = `About ${saved.toFixed(1)} in from ${listify(reasons)}, but ${needed.toFixed(1)} in is needed. Tighten the limits or accept a smaller cut.`;
  }
}

// --- Run ---

$('processBtn').addEventListener('click', async () => {
  const dir = direction();
  if (dir === 'none') return;

  const pages = detected.pages;
  $('processBtn').disabled = true;
  $('resultArea').innerHTML = '';
  $('status').innerHTML = `<div class="processing"><span class="spinner"></span>${dir === 'shrink' ? 'Minimizing' : 'Expanding'}&hellip;</div>`;

  try {
    const zip = await JSZip.loadAsync(await currentFile.arrayBuffer());
    const applied = dir === 'shrink'
      ? await smartMinimize(zip, pages, target, getRules())
      : await expandDoc(zip, pages, target);
    resultBlob = await generateBlob(zip);
    showResult(pages, target, applied, dir);
  } catch (err) {
    $('status').innerHTML = `<div class="result warning">
      <div class="result-title">Processing error</div>
      <div class="result-detail">${esc(err.message)}</div>
    </div>`;
  }
  $('processBtn').disabled = false;
});

// Opening up spacing so the content fills more pages.
async function expandDoc(zip, currentPages, targetPages) {
  const applied = [];
  const stretch = Math.min(2.2, Math.max(1.1, targetPages / currentPages));

  await transformAllXml(zip, (doc) => {
    for (const p of doc.getElementsByTagNameNS(W, 'p')) {
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
      sp.setAttributeNS(W, 'w:line', String(Math.round(240 * stretch)));
      sp.setAttributeNS(W, 'w:lineRule', 'auto');

      const before = parseInt(sp.getAttributeNS(W, 'before') || '0');
      if (before > 0) sp.setAttributeNS(W, 'w:before', String(Math.round(before * stretch)));
      const after = parseInt(sp.getAttributeNS(W, 'after') || '0');
      if (after > 0) sp.setAttributeNS(W, 'w:after', String(Math.round(after * stretch)));
    }
  });
  applied.push(`Line spacing opened to ${stretch.toFixed(2)}×`);

  await transformDocXml(zip, (doc) => {
    let n = 0;
    for (const tr of doc.getElementsByTagNameNS(W, 'tr')) {
      const trPr = tr.getElementsByTagNameNS(W, 'trPr')[0];
      if (!trPr) continue;
      const trH = trPr.getElementsByTagNameNS(W, 'trHeight')[0];
      if (!trH) continue;
      const val = parseInt(trH.getAttributeNS(W, 'val') || '0');
      if (val > 0) {
        trH.setAttributeNS(W, 'w:val', String(Math.round(val * stretch)));
        n++;
      }
    }
    if (n) applied.push(`Grew ${n} table row${n > 1 ? 's' : ''}`);
  });

  return applied;
}

async function smartMinimize(zip, currentPages, targetPages, r) {
  const applied = [];

  if (r.empties) {
    await transformDocXml(zip, (doc) => {
      const toRemove = [];
      for (const p of doc.getElementsByTagNameNS(W, 'p')) {
        if (isEmptyPara(p)) toRemove.push(p);
      }
      for (const p of toRemove) p.parentNode.removeChild(p);
      if (toRemove.length) applied.push(`Dropped ${toRemove.length} blank line${toRemove.length > 1 ? 's' : ''}`);
    });
  }

  if (r.imgRow || r.textRow) {
    await transformDocXml(zip, (doc) => {
      const imgRowTwips = Math.round(r.imgRowH * TWIP_PER_IN);
      const textRowTwips = Math.round(r.textRowH * TWIP_PER_IN);
      const imgMaxEmu = Math.round(r.imgRowH * EMU_PER_IN);
      let shrunk = 0;

      for (const tr of doc.getElementsByTagNameNS(W, 'tr')) {
        const hasImg = tr.getElementsByTagNameNS(W, 'drawing').length > 0;
        if (hasImg && !r.imgRow) continue;
        if (!hasImg && !r.textRow) continue;

        setRowHeight(doc, tr, hasImg ? imgRowTwips : textRowTwips);

        if (hasImg) {
          for (const ext of allExtents(tr)) {
            if (fitHeight(ext, imgMaxEmu)) shrunk++;
          }
        }
      }

      for (const tc of doc.getElementsByTagNameNS(W, 'tc')) {
        setCellPadding(doc, tc, 0);
      }

      applied.push('Compacted table rows');
      if (shrunk) applied.push(`Fitted ${shrunk} picture${shrunk > 1 ? 's' : ''} to ${trimNum(r.imgRowH)} in`);
    });
  }

  if (r.fonts) {
    const maxHp = r.maxFont * 2;
    let biggest = 0;
    for (const name of partNames(zip)) {
      const doc = await readXml(zip, name);
      for (const sz of fontEls(doc)) {
        const v = parseInt(sz.getAttributeNS(W, 'val') || '0');
        if (v > biggest) biggest = v;
      }
    }
    if (biggest > maxHp) {
      const ratio = maxHp / biggest;
      await transformAllXml(zip, (doc) => {
        for (const sz of fontEls(doc)) {
          const v = parseInt(sz.getAttributeNS(W, 'val') || '0');
          if (v > 0) sz.setAttributeNS(W, 'w:val', String(Math.max(2, Math.round(v * ratio))));
        }
      });
      applied.push(`Capped fonts at ${r.maxFont}pt (was ${trimNum(biggest / 2)}pt)`);
    }
  }

  if (r.spacing) {
    await transformAllXml(zip, (doc) => {
      for (const p of doc.getElementsByTagNameNS(W, 'p')) {
        const pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
        if (!pPr) continue;
        const sp = pPr.getElementsByTagNameNS(W, 'spacing')[0];
        if (!sp) continue;

        const before = parseInt(sp.getAttributeNS(W, 'before') || '0');
        if (before > 0) sp.setAttributeNS(W, 'w:before', String(Math.round(before * 0.5)));

        const after = parseInt(sp.getAttributeNS(W, 'after') || '0');
        if (after > 100) sp.setAttributeNS(W, 'w:after', String(Math.round(after * 0.7)));

        const line = parseInt(sp.getAttributeNS(W, 'line') || '0');
        if (line > 260) {
          sp.setAttributeNS(W, 'w:line', String(Math.max(240, line - 36)));
          sp.setAttributeNS(W, 'w:lineRule', 'auto');
        }
      }
    });
    applied.push('Tightened line spacing');
  }

  if (r.margins) {
    await transformDocXml(zip, (doc) => {
      const hasHeader = doc.getElementsByTagNameNS(W, 'headerReference').length > 0;
      const hasFooter = doc.getElementsByTagNameNS(W, 'footerReference').length > 0;
      let changed = false;
      for (const m of doc.getElementsByTagNameNS(W, 'pgMar')) {
        if (hasHeader && parseInt(m.getAttributeNS(W, 'top') || '0') > 0) {
          m.setAttributeNS(W, 'w:top', '0');
          m.setAttributeNS(W, 'w:header', '720');
          changed = true;
        }
        if (!hasFooter && parseInt(m.getAttributeNS(W, 'bottom') || '0') > 450) {
          m.setAttributeNS(W, 'w:bottom', '450');
          changed = true;
        }
      }
      if (changed) applied.push('Shrank page margins');
    });
  }

  if (r.freeImg) {
    await transformDocXml(zip, (doc) => {
      const ratio = Math.max(0.5, Math.min(0.95, targetPages / currentPages));
      let n = 0;
      for (const ext of allExtents(doc)) {
        if (insideTable(ext)) continue;
        scaleExtent(ext, ratio);
        n++;
      }
      if (n) applied.push(`Scaled ${n} loose picture${n > 1 ? 's' : ''} to ${Math.round(ratio * 100)}%`);
    });
  }

  return applied;
}

// --- OOXML helpers ---

function isEmptyPara(p) {
  if (Array.from(p.getElementsByTagNameNS(W, 't')).some(t => t.textContent.trim())) return false;
  if (p.getElementsByTagNameNS(W, 'drawing').length) return false;
  if (p.getElementsByTagNameNS(W, 'br').length) return false;
  if (insideTable(p)) return false;
  const pPr = p.getElementsByTagNameNS(W, 'pPr')[0];
  if (pPr && pPr.getElementsByTagNameNS(W, 'sectPr').length) return false;
  return true;
}

function insideTable(node) {
  for (let n = node.parentNode; n; n = n.parentNode) {
    if (n.localName === 'tc') return true;
  }
  return false;
}

// Every image size box: the wp:extent that Word lays out with, plus the
// a:ext inside a:xfrm that shapes and floating pictures carry.
function allExtents(root) {
  const out = [];
  for (const ns of [WP, WP14]) {
    out.push(...root.getElementsByTagNameNS(ns, 'extent'));
  }
  for (const xfrm of root.getElementsByTagNameNS(A, 'xfrm')) {
    const ext = xfrm.getElementsByTagNameNS(A, 'ext')[0];
    if (ext) out.push(ext);
  }
  return out;
}

function fitHeight(ext, maxEmu) {
  const cy = parseInt(ext.getAttribute('cy') || '0');
  if (!cy || cy <= maxEmu) return false;
  scaleExtent(ext, maxEmu / cy);
  return true;
}

function scaleExtent(ext, ratio) {
  const cx = parseInt(ext.getAttribute('cx') || '0');
  const cy = parseInt(ext.getAttribute('cy') || '0');
  if (cx) ext.setAttribute('cx', String(Math.round(cx * ratio)));
  if (cy) ext.setAttribute('cy', String(Math.round(cy * ratio)));
}

function setRowHeight(doc, tr, twips) {
  let trPr = tr.getElementsByTagNameNS(W, 'trPr')[0];
  if (!trPr) {
    trPr = doc.createElementNS(W, 'w:trPr');
    tr.insertBefore(trPr, tr.firstChild);
  }
  let trH = trPr.getElementsByTagNameNS(W, 'trHeight')[0];
  if (!trH) {
    trH = doc.createElementNS(W, 'w:trHeight');
    trPr.appendChild(trH);
  }
  trH.setAttributeNS(W, 'w:val', String(twips));
  trH.setAttributeNS(W, 'w:hRule', 'atLeast');
}

function setCellPadding(doc, tc, twips) {
  const tcPr = tc.getElementsByTagNameNS(W, 'tcPr')[0];
  if (!tcPr) return;
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
    el.setAttributeNS(W, 'w:w', String(twips));
    el.setAttributeNS(W, 'w:type', 'dxa');
  }
}

function fontEls(doc) {
  return [...doc.getElementsByTagNameNS(W, 'sz'), ...doc.getElementsByTagNameNS(W, 'szCs')];
}

function partNames(zip) {
  return Object.keys(zip.files).filter(n => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n));
}

async function readXml(zip, name) {
  const xml = await zip.file(name).async('string');
  return new DOMParser().parseFromString(xml, 'application/xml');
}

async function transformDocXml(zip, fn) {
  const doc = await readXml(zip, 'word/document.xml');
  fn(doc);
  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
}

async function transformAllXml(zip, fn) {
  for (const name of partNames(zip)) {
    const doc = await readXml(zip, name);
    fn(doc);
    zip.file(name, new XMLSerializer().serializeToString(doc));
  }
}

function generateBlob(zip) {
  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

// --- Result ---

function showResult(original, want, applied, dir) {
  const suffix = dir === 'shrink' ? '_minimized' : '_expanded';
  resultName = currentFile.name.replace(/\.docx$/i, '') + suffix + '.docx';
  const list = applied.map(t => `<span class="tech-badge">${esc(t)}</span>`).join(' ');

  $('status').innerHTML = `<div class="result success">
    <div class="result-title">Done — ${applied.length} change${applied.length === 1 ? '' : 's'} applied</div>
    <div class="result-pages">
      <span class="page-badge">${original} pp</span>
      <span class="arrow">&rarr;</span>
      <span class="page-badge">${want} pp target</span>
    </div>
    <div class="result-detail">Open it in Word to confirm the page count — all content is preserved.</div>
    <div class="tech-list">${list}</div>
  </div>`;
  $('resultArea').innerHTML = `<button class="btn-download" id="downloadBtn">&#8595; Download ${esc(resultName)}</button>`;
  $('downloadBtn').addEventListener('click', doDownload);
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

// --- Small utils ---

function clampNum(raw, lo, hi, fallback) {
  const v = parseFloat(raw);
  if (!isFinite(v)) return fallback;
  return Math.min(hi, Math.max(lo, v));
}

function trimNum(n) {
  return String(Math.round(n * 100) / 100);
}

function listify(items) {
  if (items.length === 1) return items[0];
  return items.slice(0, -1).join(', ') + ' and ' + items[items.length - 1];
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

refreshPages();
refreshEstimate();
