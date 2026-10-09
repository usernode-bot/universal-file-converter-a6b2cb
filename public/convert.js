// The conversion engine. Everything here runs in the browser: the file is read
// from the visitor's device, converted in memory and handed back as a Blob.
// Libraries are loaded on first use from /vendor (see server.js).

export const MAX_BYTES = 25 * 1024 * 1024;
export const MAX_PAGES = 200;

export const FORMATS = {
  pdf: { label: 'PDF', name: 'PDF', ext: 'pdf', mime: 'application/pdf' },
  docx: { label: 'Word', name: 'Word (DOCX)', ext: 'docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  xlsx: { label: 'Excel', name: 'Excel (XLSX)', ext: 'xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
  pptx: { label: 'PowerPoint', name: 'PowerPoint (PPTX)', ext: 'pptx', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
  jpg: { label: 'JPG', name: 'JPG image', ext: 'jpg', mime: 'image/jpeg' },
  png: { label: 'PNG', name: 'PNG image', ext: 'png', mime: 'image/png' },
  tiff: { label: 'TIFF', name: 'TIFF image', ext: 'tiff', mime: 'image/tiff' },
  zip: { label: 'ZIP', name: 'ZIP archive', ext: 'zip', mime: 'application/zip' },
};

// What each input format can become.
export const TARGETS = {
  pdf: ['docx', 'xlsx', 'pptx', 'jpg', 'png'],
  docx: ['pdf'],
  xlsx: ['pdf'],
  pptx: ['pdf'],
  jpg: ['pdf', 'docx', 'xlsx'],
  png: ['pdf', 'docx', 'xlsx'],
  tiff: ['pdf', 'docx', 'xlsx'],
};

export const IMAGE_KINDS = ['jpg', 'png', 'tiff'];

const EXT_KIND = { pdf: 'pdf', docx: 'docx', xlsx: 'xlsx', pptx: 'pptx', jpg: 'jpg', jpeg: 'jpg', png: 'png', tif: 'tiff', tiff: 'tiff' };
const OLD_OFFICE = { doc: 'docx', xls: 'xlsx', ppt: 'pptx' };

export class ConvertError extends Error {}
export class Cancelled extends Error {}

// Which supported format a file is, or a ConvertError saying why it isn't.
export function detectKind(file) {
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (OLD_OFFICE[ext]) {
    throw new ConvertError(`Older .${ext} files aren't supported. Save it as .${OLD_OFFICE[ext]} and try again.`);
  }
  let kind = EXT_KIND[ext];
  if (!kind) {
    if (file.type === 'application/pdf') kind = 'pdf';
    else if (file.type === 'image/jpeg') kind = 'jpg';
    else if (file.type === 'image/png') kind = 'png';
    else if (file.type === 'image/tiff') kind = 'tiff';
  }
  if (!kind) {
    throw new ConvertError(`${file.name} isn't a file this app can convert. Use a PDF, Word, Excel, PowerPoint, JPG, PNG or TIFF file.`);
  }
  if (file.size > MAX_BYTES) {
    throw new ConvertError(`${file.name} is ${formatBytes(file.size)}. Files can be up to ${formatBytes(MAX_BYTES)}.`);
  }
  if (file.size === 0) throw new ConvertError(`${file.name} is empty.`);
  return kind;
}

export function formatBytes(n) {
  if (n < 1024) return n + ' B';
  if (n < 1024 * 1024) return Math.round(n / 1024) + ' KB';
  return (n / (1024 * 1024)).toFixed(1).replace(/\.0$/, '') + ' MB';
}

export function baseName(name) {
  const i = name.lastIndexOf('.');
  return (i > 0 ? name.slice(0, i) : name) || 'file';
}

// ── Library loading ─────────────────────────────────────────────────────────

const scriptCache = {};
function loadScript(src) {
  if (!scriptCache[src]) {
    scriptCache[src] = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => {
        delete scriptCache[src];
        reject(new ConvertError('A part of the converter failed to load. Check your connection and try again.'));
      };
      document.head.appendChild(s);
    });
  }
  return scriptCache[src];
}

export const lib = {
  async pdfjs() {
    if (!lib._pdfjs) {
      lib._pdfjs = import('/vendor/pdfjs/legacy/build/pdf.min.mjs').then((m) => {
        m.GlobalWorkerOptions.workerSrc = '/vendor/pdfjs/legacy/build/pdf.worker.min.mjs';
        return m;
      }).catch((e) => {
        lib._pdfjs = null;
        throw e;
      });
    }
    return lib._pdfjs;
  },
  async jszip() { await loadScript('/vendor/jszip/jszip.min.js'); return window.JSZip; },
  async docx() { await loadScript('/vendor/docx/index.iife.js'); return window.docx; },
  async exceljs() { await loadScript('/vendor/exceljs/exceljs.min.js'); return window.ExcelJS; },
  async pptxgen() { await loadScript('/vendor/pptxgenjs/pptxgen.bundle.js'); return window.PptxGenJS; },
  async mammoth() { await loadScript('/vendor/mammoth/mammoth.browser.min.js'); return window.mammoth; },
  async pdfmake() {
    await loadScript('/vendor/pdfmake/pdfmake.min.js');
    await loadScript('/vendor/pdfmake/vfs_fonts.js');
    return window.pdfMake;
  },
  async tesseract() { await loadScript('/vendor/tesseract/tesseract.min.js'); return window.Tesseract; },
  async utif() {
    // UTIF picks up pako (for deflate-compressed TIFFs) when it loads.
    await loadScript('/vendor/pako/pako_inflate.min.js');
    await loadScript('/vendor/utif/UTIF.js');
    return window.UTIF;
  },
};

// ── Shared helpers ──────────────────────────────────────────────────────────

function check(signal) {
  if (signal && signal.cancelled) throw new Cancelled('Cancelled');
}

function canvasToBlob(canvas, mime, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new ConvertError('The image could not be created. The page may be too large.'))), mime, quality);
  });
}

function blobToDataURL(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
}

function pdfMakeBlob(pdfMake, def) {
  return new Promise((resolve, reject) => {
    try {
      pdfMake.createPdf(def).getBlob(resolve);
    } catch (e) {
      reject(e);
    }
  });
}

// Opens a PDF with pdf.js, turning its errors into ones a person can act on.
export async function openPdf(data) {
  const pdfjs = await lib.pdfjs();
  try {
    return await pdfjs.getDocument({
      data,
      cMapUrl: '/vendor/pdfjs/cmaps/',
      cMapPacked: true,
      standardFontDataUrl: '/vendor/pdfjs/standard_fonts/',
      wasmUrl: '/vendor/pdfjs/wasm/',
    }).promise;
  } catch (e) {
    if (e && e.name === 'PasswordException') {
      throw new ConvertError('This PDF is password-protected. Remove the password and try again.');
    }
    throw new ConvertError("This PDF couldn't be read. It may be damaged.");
  }
}

// Releases a PDF opened with openPdf, and its worker.
export async function closePdf(pdf) {
  try {
    await (pdf.loadingTask ? pdf.loadingTask.destroy() : pdf.destroy());
  } catch { /* already closed */ }
}

// Opens a PDF for converting, refusing ones too long to convert here.
async function openPdfToConvert(file) {
  const pdf = await openPdf(await file.arrayBuffer());
  if (pdf.numPages > MAX_PAGES) {
    await closePdf(pdf);
    throw new ConvertError(`This PDF has ${pdf.numPages} pages. PDFs can have up to ${MAX_PAGES} pages.`);
  }
  return pdf;
}

// Renders one PDF page to a canvas, `scale` times its size in points.
export async function renderPdfPage(page, scale) {
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, canvas, viewport }).promise;
  return canvas;
}

// Output pixels for a page: about 150 dpi, and never more than 4000 px a side.
function exportScale(page) {
  const vp = page.getViewport({ scale: 1 });
  return Math.min(150 / 72, 4000 / Math.max(vp.width, vp.height));
}

// The text of one PDF page as lines of positioned runs, top to bottom.
async function pdfPageLines(page) {
  const vp = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();
  const items = [];
  for (const it of content.items) {
    if (!it.str || !it.str.trim()) continue;
    const [a, b, , , e, f] = it.transform;
    const size = Math.hypot(a, b) || it.height || 10;
    items.push({ str: it.str, x: e, y: vp.height - f, w: it.width, size });
  }
  return groupLines(items);
}

// Groups positioned runs ({ str, x, y, w, size }) into lines, top to bottom,
// each with its runs left to right.
function groupLines(items) {
  items.sort((p, q) => p.y - q.y || p.x - q.x);
  const lines = [];
  for (const it of items) {
    const line = lines.find((l) => Math.abs(l.y - it.y) < Math.min(l.size, it.size) * 0.5);
    if (line) {
      line.items.push(it);
      line.size = Math.max(line.size, it.size);
    } else {
      lines.push({ y: it.y, size: it.size, items: [it] });
    }
  }
  lines.sort((p, q) => p.y - q.y);
  for (const l of lines) l.items.sort((p, q) => p.x - q.x);
  return lines;
}

// Joins a line's runs into cells: runs closer than about a character apart
// belong to the same cell, a wider gap starts a new one.
function lineCells(line, gapFactor) {
  const cells = [];
  for (const it of line.items) {
    const last = cells[cells.length - 1];
    const gap = last ? it.x - (last.x + last.w) : Infinity;
    if (last && gap < it.size * gapFactor) {
      const space = gap > it.size * 0.15 && !/\s$/.test(last.text) && !/^\s/.test(it.str) ? ' ' : '';
      last.text += space + it.str;
      last.w = it.x + it.w - last.x;
    } else {
      cells.push({ x: it.x, w: it.w, text: it.str });
    }
  }
  for (const c of cells) c.text = c.text.replace(/\s+/g, ' ').trim();
  return cells.filter((c) => c.text);
}

function median(nums) {
  if (!nums.length) return 0;
  const s = nums.slice().sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

function zipName(base, i, total, ext) {
  const digits = String(total).length;
  return `${base}-page-${String(i).padStart(Math.max(2, digits), '0')}.${ext}`;
}

// ── PDF to images ───────────────────────────────────────────────────────────

async function pdfToImages(file, target, { onProgress, signal }) {
  const fmt = FORMATS[target];
  const pdf = await openPdfToConvert(file);
  const total = pdf.numPages;
  const base = baseName(file.name);
  const pages = [];
  for (let i = 1; i <= total; i++) {
    check(signal);
    onProgress(i - 1, total, `Rendering page ${i} of ${total}`);
    const page = await pdf.getPage(i);
    const canvas = await renderPdfPage(page, exportScale(page));
    const blob = await canvasToBlob(canvas, fmt.mime, 0.92);
    pages.push({ name: total === 1 ? `${base}.${fmt.ext}` : zipName(base, i, total, fmt.ext), blob });
    page.cleanup();
  }
  await closePdf(pdf);
  if (pages.length === 1) {
    return { blob: pages[0].blob, name: pages[0].name, kind: target, images: pages };
  }
  check(signal);
  onProgress(total, total, 'Packing the images into a ZIP file');
  const JSZip = await lib.jszip();
  const zip = new JSZip();
  for (const p of pages) zip.file(p.name, p.blob);
  const blob = await zip.generateAsync({ type: 'blob', mimeType: FORMATS.zip.mime });
  return { blob, name: `${base}-${fmt.ext}.zip`, kind: 'zip', images: pages };
}

// ── Reading text from scans (OCR) ───────────────────────────────────────────
//
// Pages with no text of their own (scans, photos) are read with Tesseract,
// in a worker on the visitor's device. Its words come back as the same lines
// of positioned runs that pdfPageLines gives, so Word and Excel are built the
// same way whichever way the text was found.

export const OCR_LANGUAGES = [
  { code: 'eng', tag: 'en', name: 'English' },
  { code: 'spa', tag: 'es', name: 'Spanish' },
  { code: 'fra', tag: 'fr', name: 'French' },
  { code: 'deu', tag: 'de', name: 'German' },
  { code: 'por', tag: 'pt', name: 'Portuguese' },
  { code: 'ita', tag: 'it', name: 'Italian' },
];
export const MAX_OCR_PAGES = 50;

const NO_TEXT = 'No text could be read from this scan. It may be blurry, too small, sideways or handwritten, or in a different language from the one chosen. Try a sharper scan, or change the language.';

// The worker and its files are fetched from inside a worker, where a
// relative path has nothing to resolve against.
const absolute = (path) => new URL(path, location.href).href;

// Starts the text reader for one language. read() takes a page's canvas and
// resolves { lines, confidence }; close() frees the worker. Cancelling the
// conversion (signal.onCancel) stops it straight away.
async function createOcr(language, total, { onProgress, signal }) {
  const lang = OCR_LANGUAGES.find((l) => l.code === language) || OCR_LANGUAGES[0];
  onProgress(0, total, 'Getting the text reader ready');
  const Tesseract = await lib.tesseract();
  check(signal);
  let stopped = false;
  let worker = null;
  let page = null;
  let stop;
  const stopped$ = new Promise((_, reject) => { stop = reject; });
  stopped$.catch(() => {});
  // Waits for a call to the worker unless the conversion is cancelled first.
  // A call left behind by a cancel fails on its own later; that is expected.
  const unlessStopped = (p) => {
    p.catch(() => {});
    return Promise.race([p, stopped$]);
  };
  signal.onCancel = () => {
    if (stopped) return;
    stopped = true;
    if (worker) worker.terminate().catch(() => {});
    stop(new Cancelled('Cancelled'));
  };
  const logger = (m) => {
    if (stopped || signal.cancelled) return;
    if (m.status === 'recognizing text' && page) {
      onProgress(page.i - 1 + Math.min(1, m.progress || 0), page.total, `Reading the scanned text on page ${page.i} of ${page.total}`);
    } else if (m.status === 'loading language traineddata' && !page && m.progress < 1) {
      onProgress(0, total, `Downloading the ${lang.name} reading data`);
    }
  };
  const pending = Tesseract.createWorker(lang.code, 1 /* LSTM only */, {
    workerPath: absolute('/vendor/tesseract/worker.min.js'),
    corePath: absolute('/vendor/tesseract-core/'),
    langPath: absolute(`/vendor/tessdata/${lang.code}`),
    gzip: true,
    logger,
    // Failures reach us as rejected promises; without a handler the
    // library also throws them uncaught.
    errorHandler: () => {},
  });
  pending.then((w) => { if (stopped) w.terminate().catch(() => {}); }, () => {});
  try {
    worker = await unlessStopped(pending);
    await unlessStopped(worker.setParameters({ preserve_interword_spaces: '1' }));
  } catch (e) {
    if (e instanceof Cancelled) throw e;
    console.warn('OCR failed to start', e);
    throw new ConvertError(`The ${lang.name} reading data couldn't be downloaded. Check your connection and try again.`);
  }
  return {
    language: lang.code,
    async read(canvas, ptWidth, i) {
      page = { i, total };
      onProgress(i - 1, total, `Reading the scanned text on page ${i} of ${total}`);
      // Handed over as PNG bytes, which the library posts to its worker at
      // once: given a canvas it encodes it first, and a cancel in that gap
      // leaves it posting to a worker that has gone.
      const png = new Uint8Array(await (await canvasToBlob(canvas, 'image/png')).arrayBuffer());
      check(signal);
      let data;
      try {
        ({ data } = await unlessStopped(worker.recognize(png, {}, { text: false, blocks: true })));
      } catch (e) {
        if (e instanceof Cancelled) throw e;
        console.warn('OCR failed', e);
        throw new ConvertError(`The text on page ${i} couldn't be read. The page may be too large, or the device ran out of memory.`);
      }
      return ocrLines(data, ptWidth / canvas.width);
    },
    async close() {
      stopped = true;
      signal.onCancel = null;
      if (worker) await worker.terminate().catch(() => {});
    },
  };
}

// Tesseract's words as lines of runs, in points (k points per pixel). The
// page's skew, measured from the baselines, is taken out first so a slightly
// crooked scan still lines up into rows. Words read with very low confidence
// are dropped; `confidence` is the mean of the rest, weighted by length.
function ocrLines(data, k) {
  const words = [];
  const slopes = [];
  for (const block of data.blocks || []) {
    for (const para of block.paragraphs || []) {
      for (const line of para.lines || []) {
        const bl = line.baseline;
        if (bl && bl.x1 - bl.x0 > 100) slopes.push((bl.y1 - bl.y0) / (bl.x1 - bl.x0));
        words.push(...(line.words || []));
      }
    }
  }
  const slope = median(slopes);
  const items = [];
  let sum = 0;
  let chars = 0;
  for (const word of words) {
    const str = (word.text || '').trim();
    // Table rules and borders come back as runs of dashes or bars.
    if (!str || word.confidence < 30 || /^([-—–_=~]{2,}|\|+)$/.test(str)) continue;
    const { x0, y0, x1, y1 } = word.bbox;
    const y = (y0 + y1) / 2 - slope * ((x0 + x1) / 2);
    // A word's box is about 0.7 of the type's size.
    items.push({ str, x: x0 * k, y: y * k, w: (x1 - x0) * k, size: ((y1 - y0) * k) / 0.7 });
    sum += word.confidence * str.length;
    chars += str.length;
  }
  const lines = groupLines(items);
  // Every run in a line takes the line's typical size, so a word with no
  // capitals or descenders doesn't read as smaller type.
  for (const l of lines) {
    l.size = median(l.items.map((it) => it.size));
    for (const it of l.items) it.size = l.size;
  }
  return { lines, confidence: chars ? sum / chars : 0 };
}

function greyscale(canvas) {
  const ctx = canvas.getContext('2d');
  const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const d = img.data;
  for (let p = 0; p < d.length; p += 4) {
    const v = Math.round(d[p] * 0.299 + d[p + 1] * 0.587 + d[p + 2] * 0.114);
    d[p] = v;
    d[p + 1] = v;
    d[p + 2] = v;
  }
  ctx.putImageData(img, 0, 0);
  return canvas;
}

// A PDF page drawn for reading: about 300 dpi, at most 4000 px a side, grey.
async function ocrPdfCanvas(page) {
  const vp = page.getViewport({ scale: 1 });
  return greyscale(await renderPdfPage(page, Math.min(300 / 72, 4000 / Math.max(vp.width, vp.height))));
}

// A grey copy of an image for reading: small images are doubled, and none is
// more than 4000 px a side.
function ocrImageCanvas(source) {
  const long = Math.max(source.width, source.height);
  const s = Math.min(long < 1500 ? 2 : 1, 4000 / long);
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(source.width * s));
  canvas.height = Math.max(1, Math.round(source.height * s));
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(source, 0, 0, canvas.width, canvas.height);
  return greyscale(canvas);
}

// What happened to each page read with OCR, for the note under the result.
function ocrReport(language) {
  return { language, pagesRead: [], hardToRead: [], unreadable: [] };
}
function notePage(report, i, read) {
  report.pagesRead.push(i);
  if (!read.lines.length) report.unreadable.push(i);
  else if (read.confidence < 70) report.hardToRead.push(i);
}

// Numbers misread as letters, in a cell that is otherwise a number: O for 0,
// l or I for 1, and a stray full stop or comma at the end.
function ocrNumber(text) {
  if (!/^[-+]?[\dOolI.,]*\d[\dOolI.,]*$/.test(text)) return text;
  return text.replace(/[Oo]/g, '0').replace(/[lI]/g, '1').replace(/[.,]$/, '');
}

// The page numbers of a PDF with no text of their own.
async function textlessPages(pdf, { onProgress, signal }) {
  const out = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    check(signal);
    onProgress(0, pdf.numPages, 'Checking the pages for text');
    const page = await pdf.getPage(i);
    if (!(await pdfPageLines(page)).length) out.push(i);
    page.cleanup();
  }
  if (out.length > MAX_OCR_PAGES) {
    throw new ConvertError(`This PDF has ${out.length} scanned pages. Up to ${MAX_OCR_PAGES} scanned pages can be read at a time. Split it and try again.`);
  }
  return out;
}

// ── PDF to Word ─────────────────────────────────────────────────────────────

// Lines of text as Word paragraphs: bigger type becomes headings, lines close
// together join into one paragraph, and a row of a table keeps its cells
// apart with tabs.
function linesToParagraphs(docx, lines) {
  const out = [];
  const body = median(lines.map((l) => l.size));
  let para = null;
  let prev = null;
  const flush = () => {
    if (!para) return;
    const big = para.size > body * 1.6 ? docx.HeadingLevel.HEADING_1 : para.size > body * 1.25 ? docx.HeadingLevel.HEADING_2 : null;
    const runs = para.cells
      ? para.cells.map((t, k) => (k ? new docx.TextRun({ children: [new docx.Tab(), t] }) : new docx.TextRun(t)))
      : [new docx.TextRun(para.text)];
    out.push(new docx.Paragraph({ heading: big || undefined, children: runs, spacing: { after: 120 } }));
    para = null;
  };
  for (const line of lines) {
    const cells = lineCells(line, 0.5).map((c) => c.text);
    if (!cells.length) continue;
    const text = cells.join(' ');
    const gap = prev ? line.y - prev.y : 0;
    const sameStyle = para && !para.cells && Math.abs(line.size - para.size) < body * 0.15;
    if (cells.length === 1 && sameStyle && gap < line.size * 1.8) {
      para.text = /-$/.test(para.text) ? para.text.slice(0, -1) + text : para.text + ' ' + text;
    } else {
      flush();
      para = { text, size: line.size, cells: cells.length > 1 ? cells : null };
    }
    prev = line;
  }
  flush();
  return out;
}

// A page kept as a picture, as wide as the page allows. width and height are
// the page's size in points.
async function pictureParagraph(docx, canvas, width, height) {
  const data = new Uint8Array(await (await canvasToBlob(canvas, 'image/jpeg', 0.9)).arrayBuffer());
  const w = Math.min(624, width * (96 / 72)); // 6.5 in at 96 px/in
  return new docx.Paragraph({
    children: [new docx.ImageRun({ type: 'jpg', data, transformation: { width: w, height: w * (height / width) } })],
  });
}

async function pdfToDocx(file, o) {
  const { onProgress, signal } = o;
  const docx = await lib.docx();
  const pdf = await openPdfToConvert(file);
  const total = pdf.numPages;
  const children = [];
  let ocr = null;
  let report = null;
  let anyText = false;
  try {
    const scanned = await textlessPages(pdf, o);
    if (scanned.length) {
      ocr = await createOcr(o.language, total, o);
      report = ocrReport(ocr.language);
    }
    for (let i = 1; i <= total; i++) {
      check(signal);
      const page = await pdf.getPage(i);
      let lines;
      if (scanned.includes(i)) {
        const read = await ocr.read(await ocrPdfCanvas(page), page.getViewport({ scale: 1 }).width, i);
        notePage(report, i, read);
        lines = read.lines;
      } else {
        onProgress(i - 1, total, `Reading the text on page ${i} of ${total}`);
        lines = await pdfPageLines(page);
      }
      if (i > 1) children.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
      if (!lines.length) {
        // No text could be read: keep the page as a picture.
        const vp = page.getViewport({ scale: 1 });
        children.push(await pictureParagraph(docx, await renderPdfPage(page, exportScale(page)), vp.width, vp.height));
      } else {
        children.push(...linesToParagraphs(docx, lines));
        anyText = true;
      }
      page.cleanup();
    }
  } finally {
    if (ocr) await ocr.close();
    await closePdf(pdf);
  }
  if (!anyText && report) throw new ConvertError(NO_TEXT);
  check(signal);
  onProgress(total, total, 'Writing the Word document');
  const doc = new docx.Document({ sections: [{ children: children }] });
  const blob = await docx.Packer.toBlob(doc);
  return { blob, name: `${baseName(file.name)}.docx`, kind: 'docx', ocr: report };
}

// ── PDF to Excel ────────────────────────────────────────────────────────────

const NUMBER = /^[-+]?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?$/;

// Puts lines of cells into a sheet: cells whose left edges line up go in one
// column, and numbers become numbers. `ocr` tidies numbers read from a scan.
// Says whether anything was written.
function fillSheet(ws, cellLines, ocr) {
  let anyText = false;
  // Column starts: every cell's left edge, merged when within 8 points.
  const xs = cellLines.flat().map((c) => c.x).sort((a, b) => a - b);
  const cols = [];
  for (const x of xs) {
    if (!cols.length || x - cols[cols.length - 1].max > 8) cols.push({ min: x, max: x });
    else cols[cols.length - 1].max = x;
  }
  const widths = cols.map(() => 8);
  cellLines.forEach((cells, r) => {
    for (const c of cells) {
      let col = 0;
      for (let k = 0; k < cols.length; k++) if (c.x >= cols[k].min - 8) col = k;
      const cell = ws.getCell(r + 1, col + 1);
      const prior = cell.value == null ? '' : String(cell.value) + ' ';
      const text = prior + (ocr ? ocrNumber(c.text) : c.text);
      const numeric = !prior && text.length < 16 && NUMBER.test(text) && /\d/.test(text);
      cell.value = numeric ? Number(text.replace(/,/g, '')) : text;
      if (numeric && text.includes(',')) cell.numFmt = text.includes('.') ? '#,##0.00' : '#,##0';
      widths[col] = Math.min(60, Math.max(widths[col], text.length + 2));
      anyText = true;
    }
  });
  widths.forEach((w, k) => { ws.getColumn(k + 1).width = w; });
  return anyText;
}

async function pdfToXlsx(file, o) {
  const { onProgress, signal } = o;
  const ExcelJS = await lib.exceljs();
  const pdf = await openPdfToConvert(file);
  const total = pdf.numPages;
  const wb = new ExcelJS.Workbook();
  let anyText = false;
  let ocr = null;
  let report = null;
  try {
    const scanned = await textlessPages(pdf, o);
    if (scanned.length) {
      ocr = await createOcr(o.language, total, o);
      report = ocrReport(ocr.language);
    }
    for (let i = 1; i <= total; i++) {
      check(signal);
      const page = await pdf.getPage(i);
      let lines;
      if (scanned.includes(i)) {
        const read = await ocr.read(await ocrPdfCanvas(page), page.getViewport({ scale: 1 }).width, i);
        notePage(report, i, read);
        lines = read.lines;
      } else {
        onProgress(i - 1, total, `Finding the rows and columns on page ${i} of ${total}`);
        lines = await pdfPageLines(page);
      }
      page.cleanup();
      const ws = wb.addWorksheet(`Page ${i}`);
      if (fillSheet(ws, lines.map((l) => lineCells(l, 0.5)), scanned.includes(i))) anyText = true;
    }
  } finally {
    if (ocr) await ocr.close();
    await closePdf(pdf);
  }
  if (!anyText) throw new ConvertError(NO_TEXT);
  check(signal);
  onProgress(total, total, 'Writing the Excel workbook');
  const buf = await wb.xlsx.writeBuffer();
  return { blob: new Blob([buf], { type: FORMATS.xlsx.mime }), name: `${baseName(file.name)}.xlsx`, kind: 'xlsx', ocr: report };
}

// ── PDF to PowerPoint ───────────────────────────────────────────────────────

async function pdfToPptx(file, { onProgress, signal }) {
  const PptxGenJS = await lib.pptxgen();
  const pdf = await openPdfToConvert(file);
  const total = pdf.numPages;
  const pptx = new PptxGenJS();
  const first = (await pdf.getPage(1)).getViewport({ scale: 1 });
  const W = first.width / 72;
  const H = first.height / 72;
  pptx.defineLayout({ name: 'PDF', width: W, height: H });
  pptx.layout = 'PDF';
  for (let i = 1; i <= total; i++) {
    check(signal);
    onProgress(i - 1, total, `Making slide ${i} of ${total}`);
    const page = await pdf.getPage(i);
    const vp = page.getViewport({ scale: 1 });
    const canvas = await renderPdfPage(page, exportScale(page));
    const data = await blobToDataURL(await canvasToBlob(canvas, 'image/jpeg', 0.9));
    const lines = await pdfPageLines(page);
    page.cleanup();
    // Fit the page inside the slide, centred, keeping its shape.
    const s = Math.min(W / (vp.width / 72), H / (vp.height / 72));
    const w = (vp.width / 72) * s;
    const h = (vp.height / 72) * s;
    const slide = pptx.addSlide();
    slide.addImage({ data, x: (W - w) / 2, y: (H - h) / 2, w, h });
    const notes = lines.map((l) => lineCells(l, 1.5).map((c) => c.text).join(' ')).join('\n').trim();
    if (notes) slide.addNotes(notes);
  }
  await closePdf(pdf);
  check(signal);
  onProgress(total, total, 'Writing the presentation');
  const blob = await pptx.write({ outputType: 'blob' });
  return { blob, name: `${baseName(file.name)}.pptx`, kind: 'pptx' };
}

// ── Images to PDF ───────────────────────────────────────────────────────────

// Decodes an image (applying its EXIF rotation) onto a canvas.
export async function imageToCanvas(blob, maxSide = 6000) {
  let source;
  try {
    source = await createImageBitmap(blob, { imageOrientation: 'from-image' });
  } catch {
    source = await new Promise((resolve, reject) => {
      const img = new Image();
      const url = URL.createObjectURL(blob);
      img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new ConvertError("This image couldn't be read. It may be damaged.")); };
      img.src = url;
    });
  }
  const scale = Math.min(1, maxSide / Math.max(source.width, source.height));
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(source.width * scale));
  canvas.height = Math.max(1, Math.round(source.height * scale));
  canvas.getContext('2d').drawImage(source, 0, 0, canvas.width, canvas.height);
  if (source.close) source.close();
  return canvas;
}

// The pages of an image: one for a JPG or PNG, one per page of a TIFF.
// page(n) decodes page n (from 0) onto a canvas only when it is asked for,
// so a long TIFF never holds every page in memory at once.
export async function openImage(file, kind) {
  if (kind !== 'tiff') return { count: 1, page: () => imageToCanvas(file) };
  const UTIF = await lib.utif();
  const buf = await file.arrayBuffer();
  const unreadable = () => new ConvertError("This TIFF couldn't be read. Save it as PNG or PDF and try again.");
  let ifds;
  try {
    // Reduced-size copies (thumbnails) are left out.
    ifds = UTIF.decode(buf).filter((ifd) => ifd.t256 && ifd.t257 && !(ifd.t254 && ifd.t254[0] & 1));
  } catch {
    throw unreadable();
  }
  if (!ifds.length) throw unreadable();
  if (ifds.length > MAX_PAGES) throw new ConvertError(`This TIFF has ${ifds.length} pages. Images can have up to ${MAX_PAGES} pages.`);
  return {
    count: ifds.length,
    async page(n) {
      const ifd = ifds[n];
      let rgba;
      try {
        UTIF.decodeImage(buf, ifd);
        rgba = UTIF.toRGBA8(ifd);
      } catch {
        throw unreadable();
      }
      if (!ifd.width || !ifd.height || rgba.length < ifd.width * ifd.height * 4) throw unreadable();
      const full = document.createElement('canvas');
      full.width = ifd.width;
      full.height = ifd.height;
      full.getContext('2d').putImageData(new ImageData(new Uint8ClampedArray(rgba.buffer, rgba.byteOffset, ifd.width * ifd.height * 4), ifd.width, ifd.height), 0, 0);
      delete ifd.data;
      const scale = Math.min(1, 6000 / Math.max(full.width, full.height));
      if (scale === 1) return full;
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(full.width * scale);
      canvas.height = Math.round(full.height * scale);
      canvas.getContext('2d').drawImage(full, 0, 0, canvas.width, canvas.height);
      return canvas;
    },
  };
}

async function imageToPdf(file, kind, { onProgress, signal }) {
  onProgress(0, 2, 'Reading the image');
  const pdfMake = await lib.pdfmake();
  const img = await openImage(file, kind);
  const total = img.count;
  const content = [];
  let width = 0;
  let height = 0;
  for (let n = 0; n < total; n++) {
    check(signal);
    if (total > 1) onProgress(n, total + 1, `Reading page ${n + 1} of ${total}`);
    const canvas = await img.page(n);
    const data = canvas.toDataURL(kind === 'jpg' ? 'image/jpeg' : 'image/png', 0.92);
    if (!n) {
      // The page takes the (first) image's shape, with its long side as long as A4's.
      const k = 842 / Math.max(canvas.width, canvas.height);
      width = canvas.width * k;
      height = canvas.height * k;
      content.push({ image: data, width, height });
    } else {
      // Later pages of a TIFF fit inside the same page size, centred.
      content.push({ image: data, fit: [width, height], alignment: 'center', pageBreak: 'before' });
    }
  }
  check(signal);
  onProgress(total, total + 1, 'Writing the PDF');
  const blob = await pdfMakeBlob(pdfMake, {
    pageSize: { width, height },
    pageMargins: 0,
    content,
    info: { title: baseName(file.name) },
  });
  return { blob, name: `${baseName(file.name)}.pdf`, kind: 'pdf' };
}

// ── Images to Word and Excel (OCR) ──────────────────────────────────────────

// Images have no page size of their own: their text is measured as if each
// page were 612 points (US Letter) wide.
const IMAGE_PAGE_WIDTH = 612;

async function imageToOffice(file, kind, target, o) {
  const { onProgress, signal } = o;
  onProgress(0, 1, 'Reading the image');
  const img = await openImage(file, kind);
  const total = img.count;
  if (total > MAX_OCR_PAGES) {
    throw new ConvertError(`This TIFF has ${total} pages. Up to ${MAX_OCR_PAGES} scanned pages can be read at a time. Split it and try again.`);
  }
  const docx = target === 'docx' ? await lib.docx() : null;
  const ExcelJS = target === 'xlsx' ? await lib.exceljs() : null;
  const wb = ExcelJS ? new ExcelJS.Workbook() : null;
  const children = [];
  const ocr = await createOcr(o.language, total, o);
  const report = ocrReport(ocr.language);
  let anyText = false;
  try {
    for (let i = 1; i <= total; i++) {
      check(signal);
      const canvas = await img.page(i - 1);
      const read = await ocr.read(ocrImageCanvas(canvas), IMAGE_PAGE_WIDTH, i);
      notePage(report, i, read);
      if (read.lines.length) anyText = true;
      if (docx) {
        if (i > 1) children.push(new docx.Paragraph({ children: [new docx.PageBreak()] }));
        if (read.lines.length) children.push(...linesToParagraphs(docx, read.lines));
        else children.push(await pictureParagraph(docx, canvas, IMAGE_PAGE_WIDTH, IMAGE_PAGE_WIDTH * (canvas.height / canvas.width)));
      } else {
        fillSheet(wb.addWorksheet(`Page ${i}`), read.lines.map((l) => lineCells(l, 0.5)), true);
      }
    }
  } finally {
    await ocr.close();
  }
  if (!anyText) throw new ConvertError(NO_TEXT);
  check(signal);
  const base = baseName(file.name);
  if (docx) {
    onProgress(total, total, 'Writing the Word document');
    const blob = await docx.Packer.toBlob(new docx.Document({ sections: [{ children }] }));
    return { blob, name: `${base}.docx`, kind: 'docx', ocr: report };
  }
  onProgress(total, total, 'Writing the Excel workbook');
  const buf = await wb.xlsx.writeBuffer();
  return { blob: new Blob([buf], { type: FORMATS.xlsx.mime }), name: `${base}.xlsx`, kind: 'xlsx', ocr: report };
}

// ── Word to PDF ─────────────────────────────────────────────────────────────

const HEADING_SIZES = { H1: 22, H2: 18, H3: 15, H4: 13, H5: 12, H6: 11 };
const PDF_IMAGE = /^data:image\/(png|jpe?g);base64,/i;

// Turns mammoth's HTML into pdfmake content: headings, paragraphs, lists,
// tables, links and inline images.
function htmlToPdfmake(root, maxWidth) {
  function inlines(node, style) {
    const out = [];
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        const text = n.nodeValue.replace(/\s+/g, ' ');
        if (text) out.push({ text, ...style });
      } else if (n.nodeType === 1) {
        const tag = n.tagName;
        if (tag === 'BR') out.push({ text: '\n' });
        else if (tag === 'IMG') continue;
        else {
          const s = { ...style };
          if (tag === 'STRONG' || tag === 'B') s.bold = true;
          if (tag === 'EM' || tag === 'I') s.italics = true;
          if (tag === 'U') s.decoration = 'underline';
          if (tag === 'S' || tag === 'DEL') s.decoration = 'lineThrough';
          if (tag === 'SUP') s.sup = true;
          if (tag === 'SUB') s.sub = true;
          if (tag === 'A' && /^(https?:|mailto:)/i.test(n.getAttribute('href') || '')) {
            s.link = n.getAttribute('href');
            s.color = '#1d4ed8';
            s.decoration = 'underline';
          }
          out.push(...inlines(n, s));
        }
      }
    }
    return out;
  }
  function images(node) {
    return Array.from(node.querySelectorAll('img'))
      .map((img) => img.getAttribute('src') || '')
      .filter((src) => PDF_IMAGE.test(src))
      .map((src) => ({ image: src, fit: [maxWidth, 600], margin: [0, 4, 0, 8] }));
  }
  function textBlock(node, extra) {
    const parts = inlines(node, {});
    const out = [];
    if (parts.some((p) => p.text.trim())) out.push({ text: parts, ...extra });
    out.push(...images(node));
    return out;
  }
  function list(node) {
    const items = [];
    for (const li of node.children) {
      if (li.tagName !== 'LI') continue;
      const stack = [];
      const inner = li.cloneNode(true);
      const nested = Array.from(inner.querySelectorAll(':scope > ul, :scope > ol'));
      nested.forEach((n) => n.remove());
      const parts = inlines(inner, {});
      stack.push({ text: parts.length ? parts : ' ' });
      for (const n of li.querySelectorAll(':scope > ul, :scope > ol')) stack.push(list(n));
      items.push(stack.length === 1 ? stack[0] : { stack });
    }
    return { [node.tagName === 'OL' ? 'ol' : 'ul']: items.length ? items : [' '], margin: [0, 0, 0, 8] };
  }
  function table(node) {
    const rows = Array.from(node.querySelectorAll('tr')).filter((tr) => tr.closest('table') === node);
    const body = rows.map((tr) => Array.from(tr.children).map((td) => {
      const content = blocks(td);
      return content.length ? { stack: content } : { text: '' };
    }));
    const cols = Math.max(1, ...body.map((r) => r.length));
    body.forEach((r) => { while (r.length < cols) r.push({ text: '' }); });
    if (!body.length) return null;
    return { table: { widths: Array(cols).fill('*'), body }, fontSize: 10, margin: [0, 4, 0, 10] };
  }
  function blocks(node) {
    const out = [];
    let loose = [];
    const flushLoose = () => {
      if (loose.some((p) => p.text.trim())) out.push({ text: loose, margin: [0, 0, 0, 8] });
      loose = [];
    };
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        if (n.nodeValue.trim()) loose.push({ text: n.nodeValue });
        continue;
      }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName;
      if (HEADING_SIZES[tag]) {
        flushLoose();
        out.push(...textBlock(n, { fontSize: HEADING_SIZES[tag], bold: true, margin: [0, 10, 0, 6] }));
      } else if (tag === 'P' || tag === 'DIV' || tag === 'BLOCKQUOTE') {
        flushLoose();
        out.push(...textBlock(n, { margin: tag === 'BLOCKQUOTE' ? [20, 0, 0, 8] : [0, 0, 0, 8] }));
      } else if (tag === 'UL' || tag === 'OL') {
        flushLoose();
        out.push(list(n));
      } else if (tag === 'TABLE') {
        flushLoose();
        const t = table(n);
        if (t) out.push(t);
      } else if (tag === 'IMG') {
        flushLoose();
        const src = n.getAttribute('src') || '';
        if (PDF_IMAGE.test(src)) out.push({ image: src, fit: [maxWidth, 600], margin: [0, 4, 0, 8] });
      } else {
        loose.push(...inlines(n, tag === 'STRONG' ? { bold: true } : tag === 'EM' ? { italics: true } : {}));
      }
    }
    flushLoose();
    return out;
  }
  return blocks(root);
}

export async function docxToHtml(data) {
  const mammoth = await lib.mammoth();
  try {
    const result = await mammoth.convertToHtml({ arrayBuffer: data });
    return result.value;
  } catch {
    throw new ConvertError("This Word file couldn't be read. It may be damaged or password-protected.");
  }
}

async function docxToPdf(file, { onProgress, signal }) {
  onProgress(0, 2, 'Reading the Word document');
  const html = await docxToHtml(await file.arrayBuffer());
  check(signal);
  onProgress(1, 2, 'Laying out the PDF pages');
  const pdfMake = await lib.pdfmake();
  const root = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html').body;
  let content = htmlToPdfmake(root, 451);
  if (!content.length) content = [{ text: ' ' }];
  const blob = await pdfMakeBlob(pdfMake, {
    pageSize: 'A4',
    pageMargins: [72, 72, 72, 72],
    content,
    defaultStyle: { fontSize: 11, lineHeight: 1.2 },
    info: { title: baseName(file.name) },
  });
  return { blob, name: `${baseName(file.name)}.pdf`, kind: 'pdf' };
}

// ── Excel to PDF ────────────────────────────────────────────────────────────

const MAX_ROWS = 3000;

export function cellText(cell) {
  let v = cell.value;
  if (v == null) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') {
    if ('result' in v) v = v.result;
    else if (v.richText) return v.richText.map((r) => r.text).join('');
    else if ('text' in v) return String(v.text && v.text.richText ? v.text.richText.map((r) => r.text).join('') : v.text);
    else if (v.error) return String(v.error);
  }
  if (v == null) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'number') {
    try { return cell.text && cell.text !== '[object Object]' ? cell.text : String(v); } catch { return String(v); }
  }
  return typeof v === 'object' ? '' : String(v);
}

export async function readWorkbook(data) {
  const ExcelJS = await lib.exceljs();
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(data);
  } catch {
    throw new ConvertError("This Excel file couldn't be read. It may be damaged or password-protected.");
  }
  return wb;
}

// The used part of a sheet as rows of { text, bold, number } cells, with
// trailing empty rows and columns trimmed.
export function sheetGrid(ws, maxRows) {
  const rows = [];
  let cols = 0;
  const last = Math.min(ws.rowCount, maxRows);
  for (let r = 1; r <= last; r++) {
    const row = ws.getRow(r);
    const cells = [];
    for (let c = 1; c <= ws.columnCount; c++) {
      const cell = row.getCell(c);
      const text = cellText(cell);
      cells.push({ text, bold: !!(cell.font && cell.font.bold), number: typeof cell.value === 'number' || (cell.value && typeof cell.value.result === 'number') });
      if (text) cols = Math.max(cols, c);
    }
    rows.push(cells);
  }
  while (rows.length && rows[rows.length - 1].every((c) => !c.text)) rows.pop();
  return { rows: rows.map((r) => r.slice(0, cols)), cols, truncated: ws.rowCount > maxRows };
}

async function xlsxToPdf(file, { onProgress, signal }) {
  onProgress(0, 2, 'Reading the workbook');
  const wb = await readWorkbook(await file.arrayBuffer());
  const pdfMake = await lib.pdfmake();
  const sheets = wb.worksheets.filter((ws) => ws.state !== 'hidden' && ws.state !== 'veryHidden');
  const content = [];
  let widest = 0;
  sheets.forEach((ws, i) => {
    check(signal);
    onProgress(1, 2, `Laying out sheet ${i + 1} of ${sheets.length}`);
    const grid = sheetGrid(ws, MAX_ROWS);
    if (!grid.cols) return;
    widest = Math.max(widest, grid.cols);
    if (content.length) content.push({ text: '', pageBreak: 'after' });
    content.push({ text: ws.name, fontSize: 14, bold: true, margin: [0, 0, 0, 8] });
    const fontSize = grid.cols > 16 ? 6 : grid.cols > 10 ? 7 : 9;
    content.push({
      table: {
        headerRows: grid.rows.length > 1 ? 1 : 0,
        widths: Array(grid.cols).fill(grid.cols > 8 ? '*' : 'auto'),
        body: grid.rows.map((r) => r.map((c) => ({ text: c.text, bold: c.bold, alignment: c.number ? 'right' : 'left' }))),
      },
      fontSize,
      layout: {
        hLineWidth: () => 0.5,
        vLineWidth: () => 0.5,
        hLineColor: () => '#c7ccd4',
        vLineColor: () => '#c7ccd4',
        paddingLeft: () => 3,
        paddingRight: () => 3,
        paddingTop: () => 2,
        paddingBottom: () => 2,
      },
    });
    if (grid.truncated) {
      content.push({ text: `Only the first ${MAX_ROWS} rows of this sheet are included.`, italics: true, fontSize: 9, margin: [0, 6, 0, 0] });
    }
  });
  if (!content.length) throw new ConvertError('This workbook has no cells with anything in them.');
  check(signal);
  onProgress(1, 2, 'Writing the PDF');
  const blob = await pdfMakeBlob(pdfMake, {
    pageSize: 'A4',
    pageOrientation: widest > 6 ? 'landscape' : 'portrait',
    pageMargins: [28, 28, 28, 28],
    content,
    info: { title: baseName(file.name) },
  });
  return { blob, name: `${baseName(file.name)}.pdf`, kind: 'pdf' };
}

// ── PowerPoint to PDF ───────────────────────────────────────────────────────

const EMU_PER_PT = 12700;
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function kids(el, name) {
  return el ? Array.from(el.children).filter((c) => c.localName === name) : [];
}
function kid(el, name) {
  return kids(el, name)[0] || null;
}
function find(el, name) {
  if (!el) return null;
  return Array.from(el.getElementsByTagName('*')).find((c) => c.localName === name) || null;
}
function findAll(el, name) {
  return el ? Array.from(el.getElementsByTagName('*')).filter((c) => c.localName === name) : [];
}

function resolvePath(dir, target) {
  const parts = (target.startsWith('/') ? target.slice(1) : dir + '/' + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p && p !== '.') out.push(p);
  }
  return out.join('/');
}

// Reads a .pptx into slides of positioned shapes, pictures and tables, in
// points. Placeholders take their position from the slide layout or master.
export async function readPresentation(data) {
  const JSZip = await lib.jszip();
  let zip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    throw new ConvertError("This PowerPoint file couldn't be read. It may be damaged or password-protected.");
  }
  const parser = new DOMParser();
  const xml = async (p) => {
    const f = zip.file(p);
    return f ? parser.parseFromString(await f.async('string'), 'application/xml') : null;
  };
  const rels = async (p) => {
    const dir = p.slice(0, p.lastIndexOf('/'));
    const doc = await xml(`${dir}/_rels/${p.slice(dir.length + 1)}.rels`);
    const map = {};
    if (doc) {
      for (const r of findAll(doc.documentElement, 'Relationship')) {
        map[r.getAttribute('Id')] = { type: r.getAttribute('Type') || '', path: resolvePath(dir, r.getAttribute('Target') || '') };
      }
    }
    return map;
  };

  const pres = await xml('ppt/presentation.xml');
  if (!pres) throw new ConvertError("This PowerPoint file couldn't be read. It may be damaged or password-protected.");
  const size = find(pres.documentElement, 'sldSz');
  const width = (Number(size && size.getAttribute('cx')) || 9144000) / EMU_PER_PT;
  const height = (Number(size && size.getAttribute('cy')) || 6858000) / EMU_PER_PT;
  const presRels = await rels('ppt/presentation.xml');
  const slidePaths = findAll(pres.documentElement, 'sldId')
    .map((s) => presRels[s.getAttributeNS(NS_R, 'id') || s.getAttribute('r:id')])
    .filter(Boolean)
    .map((r) => r.path);

  const imageCache = {};
  const loadImage = async (p) => {
    if (!(p in imageCache)) {
      const f = zip.file(p);
      const ext = p.split('.').pop().toLowerCase();
      const type = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' }[ext];
      imageCache[p] = f && type ? new Blob([await f.async('uint8array')], { type }) : null;
    }
    return imageCache[p];
  };

  const placeholderOf = (sp) => {
    const ph = find(sp, 'ph');
    return ph ? { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') } : null;
  };
  const xfrmOf = (el) => {
    const x = find(el, 'xfrm');
    const off = kid(x, 'off');
    const ext = kid(x, 'ext');
    if (!off || !ext) return null;
    return {
      x: Number(off.getAttribute('x')) / EMU_PER_PT,
      y: Number(off.getAttribute('y')) / EMU_PER_PT,
      w: Number(ext.getAttribute('cx')) / EMU_PER_PT,
      h: Number(ext.getAttribute('cy')) / EMU_PER_PT,
    };
  };
  const shapesOf = (doc) => (doc ? findAll(doc.documentElement, 'sp') : []);
  const matchPlaceholder = (doc, ph) => {
    if (!doc || !ph) return null;
    const family = (t) => (t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t);
    const all = shapesOf(doc).map((sp) => ({ sp, ph: placeholderOf(sp) })).filter((s) => s.ph);
    return (ph.idx != null && all.find((s) => s.ph.idx === ph.idx))
      || all.find((s) => s.ph.type === ph.type)
      || all.find((s) => family(s.ph.type) === family(ph.type))
      || null;
  };
  const color = (el) => {
    const fill = el && kid(el, 'solidFill');
    const c = fill && kid(fill, 'srgbClr');
    return c ? '#' + c.getAttribute('val') : null;
  };

  const slides = [];
  for (const slidePath of slidePaths) {
    const doc = await xml(slidePath);
    if (!doc) continue;
    const slideRels = await rels(slidePath);
    const layoutRel = Object.values(slideRels).find((r) => r.type.endsWith('/slideLayout'));
    const layout = layoutRel ? await xml(layoutRel.path) : null;
    const masterRel = layoutRel ? Object.values(await rels(layoutRel.path)).find((r) => r.type.endsWith('/slideMaster')) : null;
    const master = masterRel ? await xml(masterRel.path) : null;

    const bgPr = find(doc.documentElement, 'bgPr');
    const slide = { background: color(bgPr) || '#ffffff', items: [] };

    const walk = async (tree, map) => {
      for (const el of tree.children) {
        const name = el.localName;
        if (name === 'grpSp') {
          const g = find(el, 'grpSpPr');
          const gx = kid(g, 'xfrm');
          const off = kid(gx, 'off');
          const ext = kid(gx, 'ext');
          const chOff = kid(gx, 'chOff');
          const chExt = kid(gx, 'chExt');
          let inner = map;
          if (off && ext && chOff && chExt) {
            const sx = (Number(ext.getAttribute('cx')) / Number(chExt.getAttribute('cx'))) || 1;
            const sy = (Number(ext.getAttribute('cy')) / Number(chExt.getAttribute('cy'))) || 1;
            const ox = Number(off.getAttribute('x')) / EMU_PER_PT;
            const oy = Number(off.getAttribute('y')) / EMU_PER_PT;
            const cx = Number(chOff.getAttribute('x')) / EMU_PER_PT;
            const cy = Number(chOff.getAttribute('y')) / EMU_PER_PT;
            inner = (b) => map({ x: ox + (b.x - cx) * sx, y: oy + (b.y - cy) * sy, w: b.w * sx, h: b.h * sy });
          }
          await walk(el, inner);
        } else if (name === 'sp') {
          const ph = placeholderOf(el);
          const spPr = kid(el, 'spPr');
          let box = xfrmOf(spPr);
          if (!box && ph) {
            const fromLayout = matchPlaceholder(layout, ph);
            box = fromLayout && xfrmOf(kid(fromLayout.sp, 'spPr'));
            if (!box) {
              const fromMaster = matchPlaceholder(master, ph);
              box = fromMaster && xfrmOf(kid(fromMaster.sp, 'spPr'));
            }
          }
          if (!box) continue;
          box = map(box);
          const fill = color(spPr);
          if (fill) slide.items.push({ type: 'rect', ...box, color: fill });
          const body = kid(el, 'txBody');
          if (!body) continue;
          const bodyPr = kid(body, 'bodyPr');
          const autofit = find(bodyPr, 'normAutofit');
          const scale = autofit && autofit.getAttribute('fontScale') ? Number(autofit.getAttribute('fontScale')) / 100000 : 1;
          const isTitle = ph && (ph.type === 'title' || ph.type === 'ctrTitle');
          const bulleted = ph && !isTitle && ph.type !== 'subTitle' && ['body', 'obj'].includes(ph.type);
          const defaultSize = ph ? (ph.type === 'ctrTitle' ? 40 : isTitle ? 32 : ph.type === 'subTitle' ? 20 : 18) : 18;
          const paras = [];
          for (const p of kids(body, 'p')) {
            const pPr = kid(p, 'pPr');
            const runs = [];
            for (const r of Array.from(p.children)) {
              if (r.localName === 'br') { runs.push({ text: '\n', size: defaultSize * scale }); continue; }
              if (r.localName !== 'r' && r.localName !== 'fld') continue;
              const rPr = kid(r, 'rPr');
              const t = kid(r, 't');
              if (!t) continue;
              const sz = rPr && rPr.getAttribute('sz');
              runs.push({
                text: t.textContent,
                size: (sz ? Number(sz) / 100 : defaultSize) * scale,
                bold: rPr ? rPr.getAttribute('b') === '1' : false,
                italic: rPr ? rPr.getAttribute('i') === '1' : false,
                color: color(rPr) || '#1f2937',
              });
            }
            const level = Number(pPr && pPr.getAttribute('lvl')) || 0;
            const buChar = find(pPr, 'buChar');
            const bullet = find(pPr, 'buNone') ? '' : buChar ? buChar.getAttribute('char') : bulleted ? '•' : '';
            paras.push({ runs, align: (pPr && pPr.getAttribute('algn')) || (ph && ph.type === 'ctrTitle' ? 'ctr' : 'l'), bullet: runs.length ? bullet : '', level });
          }
          if (paras.some((p) => p.runs.length)) {
            slide.items.push({ type: 'text', ...box, paras, anchor: (bodyPr && bodyPr.getAttribute('anchor')) || (isTitle ? 'ctr' : 't') });
          }
        } else if (name === 'pic') {
          const blip = find(el, 'blip');
          const rid = blip && (blip.getAttributeNS(NS_R, 'embed') || blip.getAttribute('r:embed'));
          const rel = rid && slideRels[rid];
          const box = xfrmOf(kid(el, 'spPr'));
          const blob = rel && box ? await loadImage(rel.path) : null;
          if (blob) slide.items.push({ type: 'image', ...map(box), blob });
        } else if (name === 'graphicFrame') {
          const tbl = find(el, 'tbl');
          const box = xfrmOf(el);
          if (!tbl || !box) continue;
          const rows = kids(tbl, 'tr').map((tr) => kids(tr, 'tc').map((tc) => findAll(tc, 't').map((t) => t.textContent).join(' ')));
          slide.items.push({ type: 'table', ...map(box), rows });
        }
      }
    };
    const tree = find(doc.documentElement, 'spTree');
    if (tree) await walk(tree, (b) => b);
    slides.push(slide);
  }
  if (!slides.length) throw new ConvertError('This presentation has no slides.');
  return { width, height, slides };
}

function wrapWords(ctx, text, maxWidth) {
  const out = [];
  for (const para of text.split('\n')) {
    let line = '';
    for (const word of para.split(/(\s+)/)) {
      const next = line + word;
      if (line && ctx.measureText(next).width > maxWidth && word.trim()) {
        out.push(line.trimEnd());
        line = word.trimStart();
      } else {
        line = next;
      }
    }
    out.push(line);
  }
  return out;
}

const FONT = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif';

// Draws one slide onto a canvas `scale` px per point.
export async function drawSlide(pres, slide, scale) {
  const canvas = document.createElement('canvas');
  canvas.width = Math.round(pres.width * scale);
  canvas.height = Math.round(pres.height * scale);
  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);
  ctx.fillStyle = slide.background;
  ctx.fillRect(0, 0, pres.width, pres.height);
  for (const it of slide.items) {
    if (it.type === 'rect') {
      ctx.fillStyle = it.color;
      ctx.fillRect(it.x, it.y, it.w, it.h);
    } else if (it.type === 'image') {
      try {
        const bmp = await imageToCanvas(it.blob, 3000);
        ctx.drawImage(bmp, it.x, it.y, it.w, it.h);
      } catch { /* an image the browser can't decode is left out */ }
    } else if (it.type === 'table') {
      const rows = it.rows.length || 1;
      const cols = Math.max(1, ...it.rows.map((r) => r.length));
      const rh = it.h / rows;
      const cw = it.w / cols;
      const size = Math.max(6, Math.min(14, rh * 0.45));
      ctx.strokeStyle = '#9ca3af';
      ctx.lineWidth = 0.75;
      ctx.font = `${size}px ${FONT}`;
      ctx.fillStyle = '#1f2937';
      ctx.textBaseline = 'middle';
      it.rows.forEach((r, ri) => r.forEach((text, ci) => {
        ctx.strokeRect(it.x + ci * cw, it.y + ri * rh, cw, rh);
        const shown = wrapWords(ctx, text, cw - 8)[0] || '';
        ctx.fillText(shown, it.x + ci * cw + 4, it.y + ri * rh + rh / 2);
      }));
      ctx.textBaseline = 'alphabetic';
    } else if (it.type === 'text') {
      const inset = 7.2;
      const lines = [];
      for (const p of it.paras) {
        const size = p.runs.length ? Math.max(...p.runs.map((r) => r.size)) : 18;
        const first = p.runs[0] || {};
        ctx.font = `${first.italic ? 'italic ' : ''}${first.bold ? 'bold ' : ''}${size}px ${FONT}`;
        const indent = p.level * size * 1.2 + (p.bullet ? size * 1.1 : 0);
        const text = p.runs.map((r) => r.text).join('');
        const wrapped = wrapWords(ctx, text, Math.max(10, it.w - inset * 2 - indent));
        wrapped.forEach((t, k) => lines.push({ t, size, font: ctx.font, color: first.color || '#1f2937', align: p.align, indent, bullet: k === 0 ? p.bullet : '', bulletX: p.level * size * 1.2 }));
      }
      const totalH = lines.reduce((s, l) => s + l.size * 1.2, 0);
      let y = it.y + 3.6;
      if (it.anchor === 'ctr') y = it.y + Math.max(3.6, (it.h - totalH) / 2);
      else if (it.anchor === 'b') y = it.y + Math.max(3.6, it.h - totalH - 3.6);
      ctx.save();
      for (const l of lines) {
        y += l.size;
        ctx.font = l.font;
        ctx.fillStyle = l.color;
        const left = it.x + inset + l.indent;
        const room = it.w - inset * 2 - l.indent;
        const w = ctx.measureText(l.t).width;
        const x = l.align === 'ctr' ? left + (room - w) / 2 : l.align === 'r' ? left + room - w : left;
        if (l.bullet) ctx.fillText(l.bullet, it.x + inset + l.bulletX, y);
        ctx.fillText(l.t, x, y);
        y += l.size * 0.2;
      }
      ctx.restore();
    }
  }
  return canvas;
}

async function pptxToPdf(file, { onProgress, signal }) {
  onProgress(0, 1, 'Reading the presentation');
  const pres = await readPresentation(await file.arrayBuffer());
  const pdfMake = await lib.pdfmake();
  const total = pres.slides.length;
  const scale = Math.min(2, 3000 / Math.max(pres.width, pres.height));
  const content = [];
  for (let i = 0; i < total; i++) {
    check(signal);
    onProgress(i, total, `Drawing slide ${i + 1} of ${total}`);
    const canvas = await drawSlide(pres, pres.slides[i], scale);
    content.push({
      image: canvas.toDataURL('image/jpeg', 0.9),
      width: pres.width,
      height: pres.height,
      absolutePosition: { x: 0, y: 0 },
    });
    if (i < total - 1) content.push({ text: '', pageBreak: 'after' });
  }
  check(signal);
  onProgress(total, total, 'Writing the PDF');
  const blob = await pdfMakeBlob(pdfMake, {
    pageSize: { width: pres.width, height: pres.height },
    pageMargins: 0,
    content,
    info: { title: baseName(file.name) },
  });
  return { blob, name: `${baseName(file.name)}.pdf`, kind: 'pdf' };
}

// ── Entry point ─────────────────────────────────────────────────────────────

// Converts `file` (of `kind`) to `target`. Resolves { blob, name, kind,
// images?, ocr? }, where kind is the result's own format (zip when a
// many-page PDF became images) and ocr says which pages were read from a
// scan. onProgress(done, total, label) reports each step; setting
// signal.cancelled stops at the next page, and calling signal.onCancel(),
// when set, stops reading a scan straight away. opts.language is the OCR
// language code (OCR_LANGUAGES).
export async function convert(file, kind, target, opts = {}) {
  const o = { onProgress: opts.onProgress || (() => {}), signal: opts.signal || {}, language: opts.language || 'eng' };
  if (!(TARGETS[kind] || []).includes(target)) throw new ConvertError(`${FORMATS[kind].label} files can't be converted to ${FORMATS[target].label}.`);
  if (kind === 'pdf') {
    if (target === 'jpg' || target === 'png') return pdfToImages(file, target, o);
    if (target === 'docx') return pdfToDocx(file, o);
    if (target === 'xlsx') return pdfToXlsx(file, o);
    if (target === 'pptx') return pdfToPptx(file, o);
  }
  if (IMAGE_KINDS.includes(kind)) {
    return target === 'pdf' ? imageToPdf(file, kind, o) : imageToOffice(file, kind, target, o);
  }
  if (kind === 'docx') return docxToPdf(file, o);
  if (kind === 'xlsx') return xlsxToPdf(file, o);
  if (kind === 'pptx') return pptxToPdf(file, o);
  throw new ConvertError('This conversion is not supported.');
}
