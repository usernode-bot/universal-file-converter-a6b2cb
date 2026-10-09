// The converter screen: choose a file, see it, pick a format, convert,
// see the result and download it. The conversions are in convert.js.
import {
  FORMATS, TARGETS, MAX_PAGES, IMAGE_KINDS, OCR_LANGUAGES, ConvertError, Cancelled,
  detectKind, formatBytes, baseName, convert, lib,
  openPdf, closePdf, renderPdfPage, imageToCanvas, openImage, docxToHtml,
  readWorkbook, sheetGrid, readPresentation, drawSlide,
} from '/convert.js';

const $ = (id) => document.getElementById(id);

const TARGET_NOTES = {
  docx: 'The text of each page becomes editable paragraphs.',
  xlsx: 'Each page becomes a sheet, with text lined up into rows and columns.',
  pptx: 'Each page becomes a slide. The page text goes in the speaker notes.',
  jpg: 'Each page becomes a JPG image. Several pages come as one ZIP file.',
  png: 'Each page becomes a PNG image. Several pages come as one ZIP file.',
  pdf: '',
};
const SOURCE_NOTES = {
  docx: 'Text, headings, lists, tables and pictures are kept. Page layout follows A4.',
  xlsx: 'Each sheet becomes a table, on as many pages as it needs.',
  pptx: 'Each slide becomes a page. Text, pictures and simple shapes are kept; charts and effects are not.',
};
// For JPG, PNG and TIFF, by target.
const IMAGE_TARGET_NOTES = {
  pdf: 'The image becomes a single-page PDF.',
  docx: 'The text is read from the image (OCR) and becomes editable paragraphs.',
  xlsx: 'The text is read from the image (OCR). Tables are split into rows and columns.',
};
// What a PDF's Word or Excel note adds: how scanned pages are handled, or,
// when this PDF looks scanned, that its text will be read with OCR.
const PDF_OCR_NOTES = {
  docx: 'Pages with no text, like scans, are read with OCR.',
  xlsx: 'Scanned pages are read with OCR.',
};
const SCANNED_NOTE = 'This PDF looks scanned, so its text will be read with OCR.';
const LANGUAGE_KEY = 'ocr-language';

const state = {
  file: null,
  kind: null,
  target: null,
  result: null,
  signal: null,
  urls: [],
  readable: true,
  // For a PDF: whether its first pages have no text of their own
  // ('all', 'some' or 'none'). For a TIFF: how many pages it has.
  scanned: 'none',
  pages: 1,
};

// ── Object URLs, released when the screen resets ───────────────────────────

function track(url) {
  state.urls.push(url);
  return url;
}
function releaseUrls() {
  state.urls.forEach((u) => URL.revokeObjectURL(u));
  state.urls = [];
}

// ── Previews ────────────────────────────────────────────────────────────────

const PREVIEW_PAGES = 6;

function skeleton(el) {
  el.innerHTML = `<div class="grid grid-cols-2 gap-3 sm:grid-cols-3" aria-label="Loading preview">
    <div class="skeleton h-40"></div><div class="skeleton h-40"></div><div class="skeleton hidden h-40 sm:block"></div></div>`;
}

function figure(node, caption) {
  const fig = document.createElement('figure');
  fig.className = 'flex flex-col gap-1';
  fig.appendChild(node);
  const cap = document.createElement('figcaption');
  cap.className = 'text-center text-small text-muted';
  cap.textContent = caption;
  fig.appendChild(cap);
  return fig;
}

function pageGrid(el, total, noun) {
  el.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'grid grid-cols-2 gap-3 sm:grid-cols-3';
  el.appendChild(grid);
  if (total > PREVIEW_PAGES) {
    const more = document.createElement('p');
    more.className = 'mt-3 text-center text-small text-muted';
    const rest = total - PREVIEW_PAGES;
    more.textContent = `and ${rest} more ${rest === 1 ? noun : noun + 's'}`;
    el.appendChild(more);
  }
  return grid;
}

function thumb(canvas) {
  canvas.className = 'page-thumb';
  canvas.style.width = '100%';
  canvas.style.height = 'auto';
  return canvas;
}

async function previewPdf(el, blob) {
  const pdf = await openPdf(await blob.arrayBuffer());
  const total = pdf.numPages;
  const grid = pageGrid(el, total, 'page');
  for (let i = 1; i <= Math.min(total, PREVIEW_PAGES); i++) {
    const page = await pdf.getPage(i);
    const vp = page.getViewport({ scale: 1 });
    const canvas = await renderPdfPage(page, 360 / vp.width);
    grid.appendChild(figure(thumb(canvas), `Page ${i}`));
    page.cleanup();
  }
  // Whether it looks scanned: the first few pages with no text of their own.
  const sampled = Math.min(total, 5);
  let textless = 0;
  for (let i = 1; i <= sampled; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    if (!content.items.some((it) => it.str && it.str.trim())) textless++;
    page.cleanup();
  }
  await closePdf(pdf);
  return { pages: total, scanned: textless === sampled ? 'all' : textless ? 'some' : 'none' };
}

async function previewImages(el, images) {
  const grid = pageGrid(el, images.length, 'image');
  images.slice(0, PREVIEW_PAGES).forEach((img, i) => {
    const node = document.createElement('img');
    node.className = 'page-thumb';
    node.alt = images.length > 1 ? `Page ${i + 1}` : 'Converted image';
    node.src = track(URL.createObjectURL(img.blob));
    grid.appendChild(figure(node, images.length > 1 ? img.name : `${formatBytes(img.blob.size)}`));
  });
}

async function previewImage(el, blob) {
  const canvas = await imageToCanvas(blob, 6000);
  const img = document.createElement('img');
  img.className = 'mx-auto block h-auto max-h-80 max-w-full rounded border border-line';
  img.alt = 'Preview of the image';
  img.src = track(URL.createObjectURL(blob));
  el.innerHTML = '';
  el.appendChild(img);
  return { dims: `${canvas.width} × ${canvas.height} px` };
}

// Mammoth builds its HTML from the document's text, so it carries no
// scripts; this still drops anything active before it is shown.
function cleanHtml(html) {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  doc.querySelectorAll('script,style,iframe,object,embed,link,meta').forEach((n) => n.remove());
  doc.querySelectorAll('*').forEach((n) => {
    for (const a of Array.from(n.attributes)) {
      const v = a.value.trim().toLowerCase();
      if (a.name.startsWith('on') || ((a.name === 'href' || a.name === 'src') && !/^(https?:|mailto:|data:image\/|#)/.test(v))) n.removeAttribute(a.name);
    }
    if (n.tagName === 'A') { n.target = '_blank'; n.rel = 'noopener noreferrer'; }
  });
  return doc.body.innerHTML;
}

async function previewDocx(el, blob) {
  const html = await docxToHtml(await blob.arrayBuffer());
  const box = document.createElement('div');
  box.className = 'doc-preview';
  box.innerHTML = cleanHtml(html) || '<p>This document has no text.</p>';
  el.innerHTML = '';
  el.appendChild(box);
}

async function previewXlsx(el, blob) {
  const wb = await readWorkbook(await blob.arrayBuffer());
  const sheets = wb.worksheets.filter((ws) => ws.state !== 'hidden' && ws.state !== 'veryHidden');
  el.innerHTML = '';
  const shown = sheets.slice(0, 3);
  for (const ws of shown) {
    const grid = sheetGrid(ws, 50);
    const label = document.createElement('p');
    label.className = 'mb-2 text-small font-medium text-muted';
    label.textContent = ws.name;
    el.appendChild(label);
    if (!grid.cols) {
      const empty = document.createElement('p');
      empty.className = 'mb-4 text-small text-muted';
      empty.textContent = 'This sheet is empty.';
      el.appendChild(empty);
      continue;
    }
    const wrap = document.createElement('div');
    wrap.className = 'mb-4 overflow-x-auto';
    const table = document.createElement('table');
    table.className = 'sheet-preview';
    const cols = Math.min(grid.cols, 12);
    grid.rows.slice(0, 30).forEach((row, r) => {
      const tr = document.createElement('tr');
      const head = document.createElement('td');
      head.className = 'row-head';
      head.textContent = r + 1;
      tr.appendChild(head);
      for (let c = 0; c < cols; c++) {
        const cell = row[c] || { text: '' };
        const td = document.createElement('td');
        if (cell.number) td.className = 'num';
        if (cell.bold) td.classList.add('font-semibold');
        td.textContent = cell.text;
        tr.appendChild(td);
      }
      table.appendChild(tr);
    });
    wrap.appendChild(table);
    el.appendChild(wrap);
    if (grid.rows.length > 30 || grid.cols > 12 || grid.truncated) {
      const more = document.createElement('p');
      more.className = '-mt-2 mb-4 text-small text-muted';
      more.textContent = 'Showing the top-left corner of this sheet.';
      el.appendChild(more);
    }
  }
  if (sheets.length > shown.length) {
    const more = document.createElement('p');
    more.className = 'text-small text-muted';
    more.textContent = `and ${sheets.length - shown.length} more sheets`;
    el.appendChild(more);
  }
  return { sheets: sheets.length };
}

async function previewPptx(el, blob) {
  const pres = await readPresentation(await blob.arrayBuffer());
  const total = pres.slides.length;
  const grid = pageGrid(el, total, 'slide');
  for (let i = 0; i < Math.min(total, PREVIEW_PAGES); i++) {
    const canvas = await drawSlide(pres, pres.slides[i], 360 / pres.width);
    grid.appendChild(figure(thumb(canvas), `Slide ${i + 1}`));
  }
  return { slides: total };
}

// A TIFF: browsers can't show one in an <img>, so its pages are drawn.
async function previewTiff(el, blob) {
  const img = await openImage(blob, 'tiff');
  if (img.count === 1) {
    const canvas = await img.page(0);
    const dims = `${canvas.width} × ${canvas.height} px`;
    canvas.className = 'mx-auto block h-auto max-h-80 max-w-full rounded border border-line';
    canvas.setAttribute('role', 'img');
    canvas.setAttribute('aria-label', 'Preview of the image');
    el.innerHTML = '';
    el.appendChild(canvas);
    return { dims };
  }
  const grid = pageGrid(el, img.count, 'page');
  for (let n = 0; n < Math.min(img.count, PREVIEW_PAGES); n++) {
    const full = await img.page(n);
    const canvas = document.createElement('canvas');
    canvas.width = 360;
    canvas.height = Math.max(1, Math.round(360 * (full.height / full.width)));
    canvas.getContext('2d').drawImage(full, 0, 0, canvas.width, canvas.height);
    grid.appendChild(figure(thumb(canvas), `Page ${n + 1}`));
  }
  return { pages: img.count };
}

function previewFor(kind) {
  return { pdf: previewPdf, docx: previewDocx, xlsx: previewXlsx, pptx: previewPptx, jpg: previewImage, png: previewImage, tiff: previewTiff }[kind];
}

function showPreviewError(el, message, retry) {
  el.innerHTML = '';
  const box = document.createElement('div');
  box.className = 'state-error';
  box.setAttribute('role', 'alert');
  const title = document.createElement('p');
  title.className = 'text-heading';
  title.textContent = "Couldn't open this file";
  const text = document.createElement('p');
  text.className = 'text-body text-muted';
  text.textContent = message;
  box.append(title, text);
  if (retry) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn-secondary';
    btn.textContent = 'Choose another file';
    btn.addEventListener('click', retry);
    box.appendChild(btn);
  }
  el.appendChild(box);
}

function friendly(err) {
  if (err instanceof ConvertError) return err.message;
  console.error(err);
  return "Something went wrong while reading the file. If it opens in other apps, it may use a feature this converter doesn't handle.";
}

// ── Screen states ───────────────────────────────────────────────────────────

function showPickError(message) {
  $('pick-error-text').textContent = message;
  $('pick-error').hidden = false;
}

// Stops a conversion that is running: at the next page, or at once while a
// scan is being read.
function cancelRunning() {
  if (!state.signal) return;
  state.signal.cancelled = true;
  if (state.signal.onCancel) state.signal.onCancel();
}

function resetToPick() {
  cancelRunning();
  releaseUrls();
  Object.assign(state, { file: null, kind: null, target: null, result: null, signal: null, readable: true, scanned: 'none', pages: 1 });
  $('work').hidden = true;
  $('pick').hidden = false;
  $('file-input').value = '';
  $('source-preview').innerHTML = '';
  $('result-preview').innerHTML = '';
  $('result-note').hidden = true;
  $('ocr-lang').hidden = true;
}

function setBusy(busy) {
  $('progress').hidden = !busy;
  $('convert').hidden = busy || !!state.result;
  $('change').disabled = busy;
  document.querySelectorAll('#targets input').forEach((i) => { i.disabled = busy; });
  $('ocr-lang-select').disabled = busy;
}

// Whether converting to the chosen format will read text from a scan.
function willReadScan() {
  if (state.target !== 'docx' && state.target !== 'xlsx') return false;
  return IMAGE_KINDS.includes(state.kind) || (state.kind === 'pdf' && state.scanned !== 'none');
}

function renderTargets() {
  const box = $('targets');
  box.innerHTML = '';
  for (const t of TARGETS[state.kind]) {
    const label = document.createElement('label');
    label.className = 'choice';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = 'target';
    input.value = t;
    input.className = 'sr-only';
    input.checked = t === state.target;
    input.addEventListener('change', () => selectTarget(t));
    label.append(input, document.createTextNode(FORMATS[t].name));
    box.appendChild(label);
  }
}

function selectTarget(t) {
  state.target = t;
  state.result = null;
  $('result').hidden = true;
  $('convert-error').hidden = true;
  $('convert').hidden = false;
  $('convert').textContent = `Convert to ${FORMATS[t].label}`;
  $('convert').disabled = !state.readable;
  let note;
  if (state.kind === 'pdf') {
    note = TARGET_NOTES[t];
    if (willReadScan()) note += ` ${SCANNED_NOTE}`;
    else if (PDF_OCR_NOTES[t]) note += ` ${PDF_OCR_NOTES[t]}`;
  } else if (IMAGE_KINDS.includes(state.kind)) {
    note = t === 'pdf' && state.pages > 1 ? 'Each page becomes a page of the PDF.' : IMAGE_TARGET_NOTES[t];
  } else {
    note = SOURCE_NOTES[state.kind];
  }
  $('target-note').textContent = note;
  $('ocr-lang').hidden = !willReadScan();
}

// `target` preselects a format, when the file can become it.
async function loadFile(file, target) {
  $('pick-error').hidden = true;
  let kind;
  try {
    kind = detectKind(file);
  } catch (e) {
    if (!$('work').hidden) resetToPick();
    showPickError(friendly(e));
    return;
  }
  releaseUrls();
  cancelRunning();
  Object.assign(state, { file, kind, result: null, signal: null, readable: true, scanned: 'none', pages: 1 });
  $('pick').hidden = true;
  $('work').hidden = false;
  $('result').hidden = true;
  $('convert-error').hidden = true;
  $('result-note').hidden = true;
  setBusy(false);
  $('source-tile').textContent = FORMATS[kind].ext.toUpperCase();
  $('source-name').textContent = file.name;
  $('source-meta').textContent = `${FORMATS[kind].name} · ${formatBytes(file.size)}`;
  state.target = TARGETS[kind].includes(target) ? target : TARGETS[kind][0];
  renderTargets();
  selectTarget(state.target);

  const el = $('source-preview');
  skeleton(el);
  try {
    const info = (await previewFor(kind)(el, file)) || {};
    if (state.file !== file) return;
    const extra = info.pages ? `${info.pages} ${info.pages === 1 ? 'page' : 'pages'}`
      : info.slides ? `${info.slides} ${info.slides === 1 ? 'slide' : 'slides'}`
      : info.sheets ? `${info.sheets} ${info.sheets === 1 ? 'sheet' : 'sheets'}`
      : info.dims || '';
    if (extra) $('source-meta').textContent += ` · ${extra}`;
    if (info.scanned || info.pages) {
      // Now that the file's pages are known, the note and the language
      // field can say whether its text will be read from a scan.
      state.scanned = info.scanned || 'none';
      state.pages = info.pages || 1;
      if (!state.signal && !state.result) selectTarget(state.target);
    }
    if (info.pages > MAX_PAGES) {
      state.readable = false;
      $('convert').disabled = true;
      $('target-note').textContent = `This PDF has ${info.pages} pages. PDFs can have up to ${MAX_PAGES} pages.`;
    }
  } catch (e) {
    if (state.file !== file) return;
    state.readable = false;
    $('convert').disabled = true;
    showPreviewError(el, friendly(e), () => $('file-input').click());
  }
}

function setProgress(done, total, label) {
  const pct = total ? Math.round((done / total) * 100) : 0;
  $('progress-label').textContent = label;
  $('progress-pct').textContent = `${pct}%`;
  $('progress-bar').style.width = `${pct}%`;
  $('progress-track').setAttribute('aria-valuenow', String(pct));
}

async function runConvert() {
  const { file, kind, target } = state;
  if (!file || !target) return;
  const signal = { cancelled: false };
  state.signal = signal;
  state.result = null;
  $('result').hidden = true;
  $('convert-error').hidden = true;
  setProgress(0, 1, 'Getting the converter ready');
  setBusy(true);
  try {
    const result = await convert(file, kind, target, {
      signal,
      language: $('ocr-lang-select').value,
      onProgress: (d, t, label) => { if (!signal.cancelled) setProgress(d, t, label); },
    });
    if (signal.cancelled || state.file !== file) return;
    setProgress(1, 1, 'Making the preview');
    state.result = result;
    await showResult(result);
  } catch (e) {
    if (e instanceof Cancelled || signal.cancelled || state.file !== file) return;
    $('convert-error-text').textContent = friendly(e);
    $('convert-error').hidden = false;
  } finally {
    if (state.signal === signal) {
      state.signal = null;
      setBusy(false);
    }
  }
}

async function showResult(result) {
  const fmt = FORMATS[result.kind];
  $('result-tile').textContent = fmt.ext.toUpperCase();
  $('result-name').textContent = result.name;
  const count = result.images && result.images.length > 1 ? ` · ${result.images.length} images` : '';
  $('result-meta').textContent = `${fmt.name} · ${formatBytes(result.blob.size)}${count}`;
  const link = $('download');
  link.href = track(URL.createObjectURL(result.blob));
  link.download = result.name;
  link.textContent = `Download ${result.kind === 'zip' ? 'ZIP file' : fmt.label + ' file'}`;
  if (result.ocr) $('result-meta').textContent += ' · text read with OCR';
  const note = result.ocr ? ocrNote(result.ocr, result.kind) : '';
  $('result-note').textContent = note;
  $('result-note').hidden = !note;
  $('result').hidden = false;
  const el = $('result-preview');
  skeleton(el);
  try {
    if (result.images) await previewImages(el, result.images);
    else await previewFor(result.kind)(el, result.blob);
  } catch (e) {
    showPreviewError(el, `The preview couldn't be shown, but the file is ready to download. ${friendly(e)}`);
  }
}

// "pages 2 and 5", "page 4", "pages 1, 3 and 6"
function pageList(nums) {
  if (nums.length === 1) return `page ${nums[0]}`;
  return `pages ${nums.slice(0, -1).join(', ')} and ${nums[nums.length - 1]}`;
}

// The note under a result read from a scan, naming the pages that were hard
// to read or had nothing that could be read. Empty when every page read well.
function ocrNote(ocr, kind) {
  const parts = [];
  if (ocr.hardToRead.length) {
    parts.push(`Some text on ${pageList(ocr.hardToRead)} was hard to read. Check it before you rely on it. A sharper, straighter scan gives better results.`);
  }
  if (ocr.unreadable.length) {
    const many = ocr.unreadable.length > 1;
    const kept = kind === 'docx'
      ? `so ${many ? 'they were' : 'it was'} kept as ${many ? 'pictures' : 'a picture'}`
      : `so ${many ? 'their sheets are' : 'its sheet is'} empty`;
    parts.push(`No text could be read on ${pageList(ocr.unreadable)}, ${kept}.`);
  }
  return parts.join(' ');
}

// ── Language of the text ────────────────────────────────────────────────────

// Fills the language field and picks its starting value: the person's last
// choice here, else their Homeroom language, else their device's, else English.
async function setUpLanguages() {
  const select = $('ocr-lang-select');
  for (const l of OCR_LANGUAGES) {
    const opt = document.createElement('option');
    opt.value = l.code;
    opt.textContent = l.name;
    select.appendChild(opt);
  }
  select.addEventListener('change', () => {
    try { localStorage.setItem(LANGUAGE_KEY, select.value); } catch { /* storage blocked */ }
  });
  let saved = null;
  try { saved = localStorage.getItem(LANGUAGE_KEY); } catch { /* storage blocked */ }
  if (OCR_LANGUAGES.some((l) => l.code === saved)) {
    select.value = saved;
    return;
  }
  let tag = null;
  try {
    if (window.usernode && window.usernode.getUserLocale) tag = (await window.usernode.getUserLocale()).locale;
  } catch { /* no platform shell */ }
  const match = (t) => t && OCR_LANGUAGES.find((l) => l.tag === String(t).toLowerCase().split('-')[0]);
  const lang = match(tag) || match(navigator.language) || OCR_LANGUAGES[0];
  select.value = lang.code;
}

// ── Sample file ─────────────────────────────────────────────────────────────

async function sampleFile() {
  const pdfMake = await lib.pdfmake();
  const def = {
    pageSize: 'A4',
    pageMargins: [56, 56, 56, 56],
    info: { title: 'Sample monthly report' },
    content: [
      { text: 'Sample monthly report', fontSize: 22, bold: true, margin: [0, 0, 0, 12] },
      { text: 'This PDF was made by the converter so you can try it without a file of your own. It has a heading, paragraphs, a table and a second page.', margin: [0, 0, 0, 10] },
      { text: 'Visits to the community garden went up every month this spring. The table below lists visits and new volunteers for each month.', margin: [0, 0, 0, 14] },
      {
        table: {
          headerRows: 1,
          widths: ['*', 'auto', 'auto'],
          body: [
            [{ text: 'Month', bold: true }, { text: 'Visits', bold: true }, { text: 'New volunteers', bold: true }],
            ['March', '1,240', '18'],
            ['April', '1,615', '25'],
            ['May', '2,090', '31'],
            ['June', '2,480', '27'],
          ],
        },
        layout: 'lightHorizontalLines',
        margin: [0, 0, 0, 14],
      },
      { text: 'Next steps', fontSize: 16, bold: true, margin: [0, 10, 0, 8], pageBreak: 'before' },
      { ul: ['Add two more raised beds by the gate.', 'Start a Saturday watering rota.', 'Share the harvest list each Friday.'] },
    ],
  };
  const blob = await new Promise((resolve) => pdfMake.createPdf(def).getBlob(resolve));
  return new File([blob], 'sample-report.pdf', { type: 'application/pdf' });
}

// A made-up "scan": the sample report's table drawn as a slightly crooked
// photo of a printed page, so OCR can be tried without a file.
async function sampleScan() {
  const canvas = document.createElement('canvas');
  canvas.width = 1700;
  canvas.height = 1100;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.translate(canvas.width / 2, canvas.height / 2);
  ctx.rotate((1 * Math.PI) / 180);
  ctx.translate(-canvas.width / 2, -canvas.height / 2);
  ctx.fillStyle = '#1f2937';
  ctx.textBaseline = 'alphabetic';
  const font = 'Arial, Helvetica, sans-serif';
  ctx.font = `bold 64px ${font}`;
  ctx.fillText('Sample monthly report', 140, 190);
  ctx.font = `36px ${font}`;
  ctx.fillText('Visits and new volunteers at the community garden.', 140, 270);
  const rows = [
    ['Month', 'Visits', 'New volunteers'],
    ['March', '1,240', '18'],
    ['April', '1,615', '25'],
    ['May', '2,090', '31'],
    ['June', '2,480', '27'],
  ];
  const cols = [140, 640, 1040];
  rows.forEach((row, r) => {
    const y = 400 + r * 110;
    ctx.font = `${r ? '' : 'bold '}44px ${font}`;
    row.forEach((text, c) => ctx.fillText(text, cols[c], y));
  });
  const blob = await new Promise((resolve, reject) => canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('sample scan'))), 'image/png'));
  return new File([blob], 'sample-scan.png', { type: 'image/png' });
}

// ── Wiring ──────────────────────────────────────────────────────────────────

function wire() {
  const input = $('file-input');
  const zone = $('dropzone');
  $('choose').addEventListener('click', () => input.click());
  zone.addEventListener('click', (e) => { if (e.target === zone) input.click(); });
  input.addEventListener('change', () => { if (input.files[0]) loadFile(input.files[0]); });

  let depth = 0;
  zone.addEventListener('dragenter', (e) => { e.preventDefault(); depth++; zone.dataset.dragging = ''; });
  zone.addEventListener('dragleave', () => { depth = Math.max(0, depth - 1); if (!depth) delete zone.dataset.dragging; });
  zone.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; });
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    depth = 0;
    delete zone.dataset.dragging;
    const files = e.dataTransfer.files;
    if (files.length > 1) { showPickError('Drop one file at a time.'); return; }
    if (files[0]) loadFile(files[0]);
  });
  // A file dropped beside the zone would otherwise replace the app.
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());

  $('sample').addEventListener('click', async () => {
    const btn = $('sample');
    btn.disabled = true;
    try {
      await loadFile(await sampleFile());
    } catch (e) {
      showPickError(friendly(e));
    } finally {
      btn.disabled = false;
    }
  });
  $('sample-scan').addEventListener('click', async () => {
    const btn = $('sample-scan');
    btn.disabled = true;
    try {
      await loadFile(await sampleScan(), 'xlsx');
    } catch (e) {
      showPickError(friendly(e));
    } finally {
      btn.disabled = false;
    }
  });
  $('change').addEventListener('click', () => input.click());
  $('convert').addEventListener('click', runConvert);
  $('retry').addEventListener('click', runConvert);
  $('cancel').addEventListener('click', () => {
    cancelRunning();
    state.signal = null;
    setBusy(false);
  });
  $('again').addEventListener('click', resetToPick);
  $('download').addEventListener('click', () => {
    if (window.unNative && state.result) window.unNative.toast(`Downloading ${state.result.name}`);
  });
}

wire();
setUpLanguages();
// PDF previews need pdf.js straight away, so start loading it now. The other
// libraries load the first time a conversion or preview needs them.
lib.pdfjs().then(() => {
  document.body.dataset.converter = 'ready';
}).catch((e) => {
  console.warn('pdf.js failed to load', e);
  document.body.dataset.converter = 'degraded';
});
