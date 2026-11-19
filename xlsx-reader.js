/* Minimal .xlsx / .csv reader for the browser, no libraries.
 * Reads the first worksheet and returns an array of row objects keyed by the header row.
 * An .xlsx file is a zip of XML files; the browser's DecompressionStream unzips it. */
(function (root) {
  'use strict';

  async function inflateRaw(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function readZipIndex(buf) {
    const dv = new DataView(buf), u8 = new Uint8Array(buf);
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('This file isn’t a valid .xlsx (no zip directory found).');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = {};
    const dec = new TextDecoder();
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Unreadable .xlsx (bad zip entry).');
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      files[name] = { method: method, csize: csize, off: off };
      p += 46 + nlen + xlen + clen;
    }
    return { u8: u8, dv: dv, files: files };
  }

  async function readEntry(z, name) {
    const f = z.files[name];
    if (!f) return null;
    const lh = f.off;
    if (z.dv.getUint32(lh, true) !== 0x04034b50) throw new Error('Unreadable .xlsx (bad local header).');
    const start = lh + 30 + z.dv.getUint16(lh + 26, true) + z.dv.getUint16(lh + 28, true);
    const data = z.u8.subarray(start, start + f.csize);
    const out = f.method === 0 ? data : f.method === 8 ? await inflateRaw(data) : null;
    if (!out) throw new Error('Unsupported compression in .xlsx.');
    return new TextDecoder().decode(out);
  }

  const parseXml = (t) => new DOMParser().parseFromString(t, 'application/xml');
  const byTag = (node, tag) => node.getElementsByTagNameNS('*', tag);
  function textOf(node) { let s = ''; const ts = byTag(node, 't'); for (let i = 0; i < ts.length; i++) s += ts[i].textContent; return s; }
  function colIndex(ref) { let n = 0; for (let i = 0; i < ref.length; i++) { const c = ref.charCodeAt(i); if (c < 65 || c > 90) break; n = n * 26 + (c - 64); } return n - 1; }

  async function readXlsx(buf) {
    const z = readZipIndex(buf);
    // First sheet in workbook order
    const wb = parseXml(await readEntry(z, 'xl/workbook.xml'));
    const sheet = byTag(wb, 'sheet')[0];
    if (!sheet) throw new Error('No worksheet found in this file.');
    const rid = sheet.getAttribute('r:id') || sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
    let target = 'worksheets/sheet1.xml';
    const relsTxt = await readEntry(z, 'xl/_rels/workbook.xml.rels');
    if (relsTxt) {
      const rels = byTag(parseXml(relsTxt), 'Relationship');
      for (let i = 0; i < rels.length; i++) if (rels[i].getAttribute('Id') === rid) target = rels[i].getAttribute('Target');
    }
    const path = target.charAt(0) === '/' ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    const shared = [];
    const ssTxt = await readEntry(z, 'xl/sharedStrings.xml');
    if (ssTxt) { const si = byTag(parseXml(ssTxt), 'si'); for (let i = 0; i < si.length; i++) shared.push(textOf(si[i])); }
    const sx = await readEntry(z, path);
    if (!sx) throw new Error('Worksheet missing from this file.');
    const rowsEl = byTag(parseXml(sx), 'row');
    const grid = [];
    for (let r = 0; r < rowsEl.length; r++) {
      const cells = byTag(rowsEl[r], 'c'), row = [];
      for (let i = 0; i < cells.length; i++) {
        const c = cells[i], t = c.getAttribute('t'), ref = c.getAttribute('r');
        const idx = ref ? colIndex(ref) : row.length;
        const vEl = byTag(c, 'v')[0], v = vEl ? vEl.textContent : null;
        let val;
        if (t === 's') val = v === null ? '' : shared[Number(v)];
        else if (t === 'inlineStr') val = textOf(c);
        else if (t === 'str' || t === 'e') val = v === null ? '' : v;
        else if (t === 'b') val = v === '1';
        else val = v === null || v === '' ? '' : Number(v);
        row[idx] = val;
      }
      grid.push(row);
    }
    return toObjects(grid);
  }

  function parseCsv(text) {
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; }
      else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return toObjects(rows);
  }

  function toObjects(grid) {
    const head = (grid[0] || []).map((h) => String(h === undefined ? '' : h).trim());
    const out = [];
    for (let r = 1; r < grid.length; r++) {
      const g = grid[r]; if (!g || !g.some((x) => x !== undefined && x !== '')) continue;
      const o = {}; head.forEach((h, i) => { if (h) o[h] = g[i] === undefined ? '' : g[i]; }); out.push(o);
    }
    return { headers: head.filter(Boolean), rows: out };
  }

  async function readFile(file) {
    const name = (file.name || '').toLowerCase();
    if (name.endsWith('.csv')) return parseCsv(await file.text());
    if (typeof DecompressionStream === 'undefined') throw new Error('This browser can’t read .xlsx files. Use a current Chrome, Edge, Safari or Firefox, or save the file as CSV.');
    return readXlsx(await file.arrayBuffer());
  }

  root.XlsxReader = { readFile: readFile, readXlsx: readXlsx, parseCsv: parseCsv };
})(window);
