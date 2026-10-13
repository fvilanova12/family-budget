// Receipt reading, entirely in the browser (free): Tesseract.js for photos,
// pdf.js for PDF invoices. Libraries load only when first needed.
(function () {
  'use strict';

  const LIBS = {
    tesseract: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js',
    pdf: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js',
    pdfWorker: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js',
  };

  const loaded = {};
  function loadScript(src) {
    if (!loaded[src]) {
      loaded[src] = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = src;
        s.onload = resolve;
        s.onerror = () => { delete loaded[src]; reject(new Error('Could not load ' + src)); };
        document.head.appendChild(s);
      });
    }
    return loaded[src];
  }

  // ---------- Images ----------
  async function decode(file) {
    try {
      return await createImageBitmap(file, { imageOrientation: 'from-image' });
    } catch (e) {
      const url = URL.createObjectURL(file);
      try {
        const img = new Image();
        img.src = url;
        await img.decode();
        return img;
      } finally {
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      }
    }
  }

  function draw(src, maxSide) {
    const w = src.width, h = src.height;
    const k = Math.min(1, maxSide / Math.max(w, h));
    const c = document.createElement('canvas');
    c.width = Math.round(w * k);
    c.height = Math.round(h * k);
    c.getContext('2d').drawImage(src, 0, 0, c.width, c.height);
    return c;
  }

  // Greyscale + contrast stretch: helps Tesseract on thermal receipts.
  function forOcr(src) {
    const c = draw(src, 2000);
    const ctx = c.getContext('2d');
    const img = ctx.getImageData(0, 0, c.width, c.height);
    const d = img.data;
    let min = 255, max = 0;
    for (let i = 0; i < d.length; i += 4) {
      const g = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
      d[i] = g;
      if (g < min) min = g;
      if (g > max) max = g;
    }
    const range = Math.max(1, max - min);
    for (let i = 0; i < d.length; i += 4) {
      const v = ((d[i] - min) / range) * 255;
      d[i] = d[i + 1] = d[i + 2] = v;
    }
    ctx.putImageData(img, 0, 0);
    return c;
  }

  function toJpeg(canvas, quality) {
    return new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', quality));
  }

  let workerPromise = null;
  let progressCb = null;
  async function getWorker() {
    await loadScript(LIBS.tesseract);
    if (!workerPromise) {
      workerPromise = window.Tesseract.createWorker(['spa', 'eng'], 1, {
        logger: (m) => {
          if (!progressCb) return;
          if (m.status === 'recognizing text') progressCb(0.3 + m.progress * 0.7);
          else if (typeof m.progress === 'number') progressCb(Math.min(0.3, m.progress * 0.3));
        },
      }).catch((e) => { workerPromise = null; throw e; });
    }
    return workerPromise;
  }

  async function ocrCanvas(canvas, onProgress) {
    progressCb = onProgress || null;
    try {
      const worker = await getWorker();
      const { data } = await worker.recognize(canvas);
      return data.text || '';
    } finally {
      progressCb = null;
    }
  }

  // ---------- PDFs ----------
  async function readPdf(file, onProgress) {
    await loadScript(LIBS.pdf);
    const pdfjs = window.pdfjsLib;
    pdfjs.GlobalWorkerOptions.workerSrc = LIBS.pdfWorker;
    const doc = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
    let text = '';
    for (let p = 1; p <= Math.min(doc.numPages, 3); p++) {
      const page = await doc.getPage(p);
      const content = await page.getTextContent();
      let lastY = null;
      for (const item of content.items) {
        const y = item.transform ? item.transform[5] : null;
        if (lastY !== null && y !== null && Math.abs(y - lastY) > 2) text += '\n';
        else if (text && !text.endsWith('\n')) text += ' ';
        text += item.str;
        lastY = y;
      }
      text += '\n';
    }
    const page1 = await doc.getPage(1);
    const vp1 = page1.getViewport({ scale: 1 });
    const viewport = page1.getViewport({ scale: Math.min(3, 1400 / vp1.width) });
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(viewport.width);
    canvas.height = Math.round(viewport.height);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page1.render({ canvasContext: ctx, viewport }).promise;
    if (text.replace(/\s/g, '').length < 20) {
      // Scanned PDF without a text layer: fall back to OCR.
      text = await ocrCanvas(forOcr(canvas), onProgress);
    }
    return { text, canvas };
  }

  // Returns { text, photo (JPEG blob ~200 KB), previewUrl }
  async function read(file, onProgress) {
    let text, photoCanvas;
    if (file.type === 'application/pdf' || /\.pdf$/i.test(file.name || '')) {
      const r = await readPdf(file, onProgress);
      text = r.text;
      photoCanvas = draw(r.canvas, 1600);
    } else {
      const img = await decode(file);
      photoCanvas = draw(img, 1600);
      text = await ocrCanvas(forOcr(img), onProgress);
    }
    const photo = await toJpeg(photoCanvas, 0.72);
    return { text, photo, previewUrl: URL.createObjectURL(photo) };
  }

  // ---------- Parsing ----------
  const norm = (s) => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();

  const AMOUNT_RE = /(?:^|[^\d.,])(\d{1,3}(?:[.\s]\d{3})+|\d{1,6})\s?[,.]\s?(\d{2})(?![\d])/g;
  function amountsIn(line) {
    const out = [];
    let m;
    AMOUNT_RE.lastIndex = 0;
    while ((m = AMOUNT_RE.exec(line))) {
      const v = parseFloat(m[1].replace(/[.\s]/g, '') + '.' + m[2]);
      if (v > 0 && v < 50000) out.push(v);
    }
    return out;
  }

  const SKIP = /subtotal|base imponible|\bbase\b|\biva\b|cuota|cambio|entregado|descuento|ahorro|dto\b|puntos|saldo|propina/;
  function scoreLine(n) {
    if (SKIP.test(n)) return 0;
    if (/total a pagar|importe total|total compra|total factura|total eur|total \(?€|a pagar/.test(n)) return 6;
    if (/\btotal\b/.test(n)) return 5;
    if (/\bimporte\b/.test(n)) return 3;
    if (/tarjeta|visa|mastercard|efectivo/.test(n)) return 2;
    return 0;
  }

  function findAmount(lines) {
    let best = null;
    lines.forEach((line, i) => {
      const n = norm(line);
      const score = scoreLine(n);
      if (!score) return;
      let vals = amountsIn(line);
      if (!vals.length && lines[i + 1]) vals = amountsIn(lines[i + 1]);
      if (!vals.length) return;
      const v = Math.max(...vals);
      if (!best || score > best.score || (score === best.score && v > best.v)) best = { score, v };
    });
    if (best) return { value: best.v, sure: best.score >= 5 };
    const all = [];
    lines.forEach((line) => { if (!SKIP.test(norm(line))) all.push(...amountsIn(line)); });
    if (all.length) return { value: Math.max(...all), sure: false };
    return null;
  }

  const MONTHS = { ene: 1, feb: 2, mar: 3, abr: 4, may: 5, jun: 6, jul: 7, ago: 8, sep: 9, set: 9, oct: 10, nov: 11, dic: 12,
    jan: 1, apr: 4, aug: 8, dec: 12 };
  function iso(y, m, d) { return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0'); }
  function plausible(y, m, d, today) {
    if (m < 1 || m > 12 || d < 1 || d > 31) return false;
    const dt = new Date(y, m - 1, d);
    if (dt.getMonth() !== m - 1) return false;
    const t = new Date(today + 'T12:00:00');
    const diff = (t - dt) / 86400000;
    return diff > -2 && diff < 400;
  }
  function findDate(text, today) {
    const n = norm(text);
    let m;
    const re = /\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/g;
    while ((m = re.exec(n))) {
      let y = +m[3];
      if (y < 100) y += 2000;
      if (plausible(y, +m[2], +m[1], today)) return iso(y, +m[2], +m[1]);
    }
    const re2 = /\b(\d{1,2})\s*(?:de\s+)?([a-z]{3})[a-z]*\.?\s*(?:de\s+)?(\d{4})\b/g;
    while ((m = re2.exec(n))) {
      const mo = MONTHS[m[2]];
      if (mo && plausible(+m[3], mo, +m[1], today)) return iso(+m[3], mo, +m[1]);
    }
    const re3 = /\b(20\d{2})-(\d{2})-(\d{2})\b/g;
    while ((m = re3.exec(n))) {
      if (plausible(+m[1], +m[2], +m[3], today)) return iso(+m[1], +m[2], +m[3]);
    }
    return null;
  }

  const JUNK = /factura|ticket|simplificad|\bnif\b|\bcif\b|\btel\b|telf|c\/|calle|avda|avenida|www|http|fecha|hora|\bcp\b|gracias|bienvenid|\bcliente\s*:|n\.?\s*cliente/;
  function titleCase(s) {
    return s.toLowerCase().replace(/(^|[\s\-/.])([a-zà-ÿ])/g, (m, a, b) => a + b.toUpperCase());
  }
  function findMerchant(lines, normText, rules) {
    const sorted = rules.slice().sort((a, b) => b.pattern.length - a.pattern.length);
    for (const r of sorted) {
      if (r.pattern && normText.includes(r.pattern)) return { name: r.display_name, rule: r };
    }
    for (const line of lines.slice(0, 8)) {
      if (amountsIn(line).length) continue;
      const clean = line.replace(/[^A-Za-zÀ-ÿ0-9&'.\- ]/g, ' ').replace(/\s+/g, ' ').trim()
        .replace(/[,\s]+(s\.?\s?l\.?\s?u?\.?|s\.?\s?a\.?\s?u?\.?)(\s+[A-Z]-?\d{7,8})?$/i, '').trim();
      const letters = clean.replace(/[^A-Za-zÀ-ÿ]/g, '');
      if (letters.length < 3 || JUNK.test(norm(clean))) continue;
      if (/\d{5,}/.test(clean)) continue;
      return { name: titleCase(clean).slice(0, 40), rule: null };
    }
    return null;
  }

  function findPayment(n) {
    if (/tarjeta|visa|mastercard|contactless|maestro|\btpv\b|credito|debito/.test(n)) return 'card';
    if (/efectivo|entregado|cambio/.test(n)) return 'cash';
    if (/domiciliac|adeudo|recibo domiciliado/.test(n)) return 'direct_debit';
    if (/transferencia/.test(n)) return 'transfer';
    return null;
  }

  // rules: [{pattern, display_name, category_id}] (category_id null = ask)
  function parse(text, rules, today) {
    const lines = String(text || '').split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    const n = norm(text);
    const amount = findAmount(lines);
    const date = findDate(text, today);
    const merchant = findMerchant(lines, n, rules || []);
    return {
      amount: amount ? amount.value : null,
      amountSure: !!(amount && amount.sure),
      date,
      merchant: merchant ? merchant.name : '',
      categoryId: merchant && merchant.rule ? merchant.rule.category_id : null,
      fromKnownShop: !!(merchant && merchant.rule),
      payment: findPayment(n),
    };
  }

  // Find the rule that matches a typed shop name (longest pattern wins).
  function matchRule(name, rules) {
    const n = norm(name).trim();
    if (!n) return null;
    let best = null;
    for (const r of rules || []) {
      if (r.pattern && (n.includes(r.pattern) || r.pattern === n) && (!best || r.pattern.length > best.pattern.length)) best = r;
    }
    return best;
  }

  window.Receipt = { read, parse, matchRule, norm, loadScript };
})();
