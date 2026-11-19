/* Weekly tracker tab: upload each week's online files, see each KPI against the YTD baseline. */
(function () {
  'use strict';
  const App = window.App, Calc = window.TrackerCalc, Reader = window.XlsxReader;
  const el = App.el;
  const $ = (id) => document.getElementById(id);
  const MARKETS = ['MY', 'SG', 'TH'];
  const NAMES = { MY: 'Malaysia', SG: 'Singapore', TH: 'Thailand' };
  const CUR = { MY: 'RM', SG: 'S$', TH: '฿' };
  const REQUIRED = ['Country', 'Week', 'Transactions', 'Total Unit Sold', 'Total Sales', 'Department', 'Category', 'Brand'];
  const BASE_PERIOD = 'Baseline (YTD, before changes)';
  // Demo-mode copy of the YTD baseline (the live version comes from the Google Sheet).
  const DEMO_BASELINE = [{"Market": "MY", "Channel": "Online", "net_sales": 28856598.58, "gross_sales": 29967272.53, "returns_value": 1110673.95, "return_pct": 3.71, "baskets": 76538, "sale_units": 95367, "atv": 391.53, "upt": 1.246, "aur": 314.23, "single_unit_pct": 84.53, "multi_unit_sales_pct": 23.05, "fw_baskets": 55341, "fw_plus_acc_pct": 2.36, "two_plus_fw_pct": 5.43, "fw_acc_pct": 3.27, "fw_shoecare_pct": 0.17, "fw_socks_pct": 1.79, "fw_apparel_pct": 2.33, "crep_sales": 13419.7, "socks_sales": 171444.05, "acc_sales": 1878650.85, "crep_sales_pct": 0.047, "socks_sales_pct": 0.594, "acc_sales_pct": 6.51}, {"Market": "MY", "Channel": "Stores", "net_sales": 324166507.33, "gross_sales": 327956307.11, "returns_value": 3794607.46, "return_pct": 1.16, "baskets": 922883, "sale_units": 1373641, "atv": 355.36, "upt": 1.4884, "aur": 238.75, "single_unit_pct": 68.19, "multi_unit_sales_pct": 47.19, "fw_baskets": 529956, "fw_plus_acc_pct": 13.56, "two_plus_fw_pct": 6.77, "fw_acc_pct": 23.61, "fw_shoecare_pct": 13.8, "fw_socks_pct": 9.64, "fw_apparel_pct": 4.45, "crep_sales": 7983145.34, "socks_sales": 9943936.87, "acc_sales": 40601740.52, "crep_sales_pct": 2.463, "socks_sales_pct": 3.068, "acc_sales_pct": 12.525}, {"Market": "SG", "Channel": "Online", "net_sales": 4723777.0, "gross_sales": 5107171.02, "returns_value": 383394.02, "return_pct": 7.51, "baskets": 41561, "sale_units": 53113, "atv": 122.88, "upt": 1.278, "aur": 96.16, "single_unit_pct": 82.51, "multi_unit_sales_pct": 27.44, "fw_baskets": 28439, "fw_plus_acc_pct": 2.44, "two_plus_fw_pct": 6.67, "fw_acc_pct": 3.57, "fw_shoecare_pct": 0.13, "fw_socks_pct": 1.99, "fw_apparel_pct": 3.03, "crep_sales": 1881.51, "socks_sales": 53974.52, "acc_sales": 343737.51, "crep_sales_pct": 0.04, "socks_sales_pct": 1.143, "acc_sales_pct": 7.277}, {"Market": "SG", "Channel": "Stores", "net_sales": 34888305.17, "gross_sales": 35377991.24, "returns_value": 490113.72, "return_pct": 1.39, "baskets": 307796, "sale_units": 455930, "atv": 114.94, "upt": 1.4813, "aur": 77.6, "single_unit_pct": 69.58, "multi_unit_sales_pct": 45.64, "fw_baskets": 169607, "fw_plus_acc_pct": 11.26, "two_plus_fw_pct": 6.47, "fw_acc_pct": 20.44, "fw_shoecare_pct": 8.71, "fw_socks_pct": 10.39, "fw_apparel_pct": 4.77, "crep_sales": 510001.41, "socks_sales": 1364264.7, "acc_sales": 4417760.92, "crep_sales_pct": 1.462, "socks_sales_pct": 3.91, "acc_sales_pct": 12.663}, {"Market": "TH", "Channel": "Online", "net_sales": 165061313.56, "gross_sales": 181213365.14, "returns_value": 16152051.58, "return_pct": 8.91, "baskets": 58056, "sale_units": 70876, "atv": 3121.35, "upt": 1.2208, "aur": 2556.77, "single_unit_pct": 86.89, "multi_unit_sales_pct": 21.25, "fw_baskets": 42516, "fw_plus_acc_pct": 1.94, "two_plus_fw_pct": 5.72, "fw_acc_pct": 2.65, "fw_shoecare_pct": 0.12, "fw_socks_pct": 1.43, "fw_apparel_pct": 1.65, "crep_sales": 122741.95, "socks_sales": 1042667.3, "acc_sales": 10644336.29, "crep_sales_pct": 0.074, "socks_sales_pct": 0.632, "acc_sales_pct": 6.449}, {"Market": "TH", "Channel": "Stores", "net_sales": 1671706564.26, "gross_sales": 1688913906.18, "returns_value": 17218994.16, "return_pct": 1.02, "baskets": 583776, "sale_units": 833564, "atv": 2893.09, "upt": 1.4279, "aur": 2026.14, "single_unit_pct": 71.76, "multi_unit_sales_pct": 39.55, "fw_baskets": 354790, "fw_plus_acc_pct": 11.6, "two_plus_fw_pct": 5.43, "fw_acc_pct": 19.08, "fw_shoecare_pct": 10.47, "fw_socks_pct": 7.99, "fw_apparel_pct": 3.25, "crep_sales": 32114615.45, "socks_sales": 48798832.07, "acc_sales": 169877910.02, "crep_sales_pct": 1.921, "socks_sales_pct": 2.919, "acc_sales_pct": 10.162}];

  // better: 'up' = higher is better, 'down' = lower is better
  const METRICS = [
    { k: 'upt', label: 'Items per basket (UPT)', fmt: 'dec2', better: 'up', tile: true },
    { k: 'fw_plus_acc_pct', label: 'Baskets with footwear + an accessory', fmt: 'pct', better: 'up', tile: true },
    { k: 'fw_acc_pct', label: 'Footwear baskets with any accessory', fmt: 'pct', better: 'up', tile: true },
    { k: 'fw_shoecare_pct', label: 'Footwear baskets with shoe care', fmt: 'pct', better: 'up', tile: true },
    { k: 'fw_socks_pct', label: 'Footwear baskets with socks', fmt: 'pct', better: 'up', tile: true },
    { k: 'single_unit_pct', label: 'Single-unit baskets', fmt: 'pct', better: 'down', tile: true },
    { k: 'crep_sales_pct', label: 'Crep Protect, % of online sales', fmt: 'pct2', better: 'up', tile: true },
    { k: 'acc_sales_pct', label: 'Accessories, % of online sales', fmt: 'pct', better: 'up', tile: true },
    { k: 'multi_unit_sales_pct', label: 'Sales from multi-item baskets', fmt: 'pct', better: 'up' },
    { k: 'atv', label: 'Average basket value', fmt: 'money', better: 'up' },
    { k: 'return_pct', label: 'Returns, % of gross sales', fmt: 'pct', better: 'down' }
  ];

  let rows = [];            // every saved KPI row (baseline + weeks)
  let market = 'MY', metric = 'upt';
  let staged = [];          // parsed uploads waiting to be saved
  let loaded = false, loading = false;

  const num = (v) => { const x = Number(v); return isFinite(x) ? x : 0; };
  function fmt(f, v, m) {
    if (v === null || v === undefined || v === '' || !isFinite(Number(v))) return '—';
    v = Number(v);
    if (f === 'pct') return v.toFixed(1) + '%';
    if (f === 'pct2') return v.toFixed(2) + '%';
    if (f === 'dec2') return v.toFixed(2);
    if (f === 'money') return (CUR[m] || '') + Math.round(v).toLocaleString();
    if (f === 'int') return Math.round(v).toLocaleString();
    return String(v);
  }
  function fmtDelta(f, d) {
    const sign = d > 0 ? '+' : d < 0 ? '−' : '±';
    const a = Math.abs(d);
    if (f === 'pct' || f === 'pct2') return sign + a.toFixed(f === 'pct2' ? 2 : 1) + ' pts';
    if (f === 'dec2') return sign + a.toFixed(2);
    if (f === 'money') return sign + Math.round(a).toLocaleString();
    return sign + a;
  }
  const isBase = (r) => num(r.Week) === 0;
  const weekKey = (r) => num(r.Year) * 100 + num(r.Week);
  function weeksFor(m) { return rows.filter((r) => r.Market === m && !isBase(r) && r.Channel === 'Online').sort((a, b) => weekKey(a) - weekKey(b)); }
  function base(m, ch) { return rows.find((r) => r.Market === m && isBase(r) && r.Channel === ch) || null; }

  // ---------- data ----------
  function demoWeeks() { try { return JSON.parse(localStorage.getItem('bb-demo-weeks') || '[]'); } catch (e) { return []; } }
  async function load() {
    if (loading) return; loading = true;
    try {
      if (App.DEMO) {
        rows = DEMO_BASELINE.map((b) => Object.assign({ Year: '', Week: 0, Period: BASE_PERIOD }, b)).concat(demoWeeks());
      } else {
        if (!App.getKey()) return;
        const res = await fetch(App.URL + '?action=tracker&key=' + encodeURIComponent(App.getKey()) + '&t=' + Date.now(), { cache: 'no-store' });
        const r = await res.json();
        if (!r.ok) { if (r.error === 'unauthorised') App.onAuthFail(); throw new Error(r.error || 'Load failed'); }
        rows = r.kpis;
      }
      loaded = true; render();
    } catch (e) {
      App.toast('Couldn’t load the tracker: ' + e.message, true);
    } finally { loading = false; }
  }

  async function saveGroup(g) {
    const payload = { market: g.market, year: g.year, week: g.week, file: g.files.join(', '), kpis: g.result.kpis, categories: g.result.categories };
    if (App.DEMO) {
      const list = demoWeeks().filter((r) => !(r.Market === g.market && num(r.Year) === g.year && num(r.Week) === g.week));
      list.push(Object.assign({ Market: g.market, Year: g.year, Week: g.week, Channel: 'Online', Period: 'Wk ' + g.week + ' ' + g.year, Uploaded: new Date().toISOString().slice(0, 19), 'Uploaded by': App.getUser(), File: payload.file }, g.result.kpis));
      localStorage.setItem('bb-demo-weeks', JSON.stringify(list));
      return;
    }
    const res = await fetch(App.URL, { method: 'POST', body: JSON.stringify({ key: App.getKey(), user: App.getUser(), action: 'saveWeek', week: payload }) });
    const r = await res.json();
    if (!r.ok) throw new Error(r.error || 'Save failed');
  }

  // ---------- upload ----------
  async function onFiles(files) {
    const year = Math.round(num($('tkYear').value)) || new Date().getFullYear();
    const box = $('tkStaged'); box.replaceChildren(el('p', { class: 'tk-note', text: 'Reading ' + files.length + ' file(s)…' }));
    const groups = new Map(); const problems = [];
    for (const f of files) {
      try {
        const x = await Reader.readFile(f);
        const missing = REQUIRED.filter((h) => x.headers.indexOf(h) < 0);
        if (missing.length) { problems.push(f.name + ': missing column(s) ' + missing.join(', ')); continue; }
        let skipped = 0;
        x.rows.forEach((r) => {
          const m = String(r.Country || '').trim().toUpperCase(), w = Math.round(num(r.Week));
          if (MARKETS.indexOf(m) < 0 || w < 1 || w > 53) { skipped++; return; }
          const key = m + '|' + w;
          if (!groups.has(key)) groups.set(key, { market: m, week: w, year: year, lines: [], files: [], notOnline: 0 });
          const g = groups.get(key); g.lines.push(r); if (g.files.indexOf(f.name) < 0) g.files.push(f.name);
          if (String(r.Transactions || '').toUpperCase().indexOf('ECOMS') < 0) g.notOnline++;
        });
        if (skipped) problems.push(f.name + ': ' + skipped + ' line(s) skipped (Country not MY/SG/TH, or Week not 1–53).');
        if (!x.rows.length) problems.push(f.name + ': no rows found.');
      } catch (e) { problems.push(f.name + ': ' + e.message); }
    }
    staged = Array.from(groups.values()).map((g) => Object.assign(g, { result: Calc.compute(g.lines) }))
      .sort((a, b) => a.market.localeCompare(b.market) || a.week - b.week);
    renderStaged(problems);
  }

  function renderStaged(problems) {
    const box = $('tkStaged'); box.replaceChildren();
    (problems || []).forEach((p) => box.appendChild(el('p', { class: 'tk-warn', text: p })));
    if (!staged.length) { if (!(problems || []).length) box.appendChild(el('p', { class: 'tk-note', text: 'No weeks found in the file(s).' })); return; }
    const table = el('table', { class: 'tk-table' }, [el('tr', {}, ['Market', 'Week', 'Lines', 'Baskets', 'Net sales', 'UPT', 'Footwear + acc.', 'Status'].map((h) => el('th', { text: h })))]);
    staged.forEach((g) => {
      const exists = rows.some((r) => r.Market === g.market && num(r.Year) === g.year && num(r.Week) === g.week && r.Channel === 'Online');
      const k = g.result.kpis;
      table.appendChild(el('tr', {}, [
        el('td', { text: g.market }), el('td', { text: 'Wk ' + g.week + ' ' + g.year }), el('td', { text: g.lines.length.toLocaleString() }),
        el('td', { text: fmt('int', k.baskets) }), el('td', { text: fmt('money', k.net_sales, g.market) }), el('td', { text: fmt('dec2', k.upt) }),
        el('td', { text: fmt('pct', k.fw_plus_acc_pct) }),
        el('td', { class: exists ? 'tk-replace' : 'tk-new', text: exists ? 'Replaces saved week' : 'New week' })]));
      if (g.notOnline) box.appendChild(el('p', { class: 'tk-warn', text: g.market + ' wk ' + g.week + ': ' + g.notOnline + ' line(s) don’t look like online (ECOMS) transactions. Check this is the online file. They’re included.' }));
    });
    box.appendChild(table);
    const btn = el('button', { class: 'btn', type: 'button', text: 'Save ' + staged.length + ' week' + (staged.length > 1 ? 's' : '') + ' to the tracker' });
    btn.addEventListener('click', async () => {
      btn.disabled = true; let done = 0;
      try {
        for (const g of staged) { btn.textContent = 'Saving ' + (++done) + ' of ' + staged.length + '…'; await saveGroup(g); }
        App.toast('Saved ' + staged.length + ' week(s)');
        const last = staged[staged.length - 1]; market = last.market;
        staged = []; $('tkFile').value = ''; box.replaceChildren(); await load();
      } catch (e) { App.toast('Not saved: ' + e.message, true); btn.disabled = false; btn.textContent = 'Try again'; }
    });
    box.appendChild(el('div', { class: 'actions' }, [btn]));
  }

  // ---------- views ----------
  function tile(m, def, latest, prev, bOn, bSt) {
    const v = latest ? num(latest[def.k]) : null, b = bOn ? num(bOn[def.k]) : null, st = bSt ? num(bSt[def.k]) : null;
    const kids = [el('div', { class: 'plab', text: def.label }), el('div', { class: 'pbig', text: latest ? fmt(def.fmt, v, m) : '—' })];
    if (latest && b !== null) {
      const d = v - b, good = def.better === 'up' ? d > 0 : d < 0;
      kids.push(el('div', { class: 'tk-delta ' + (Math.abs(d) < 1e-9 ? '' : good ? 'good' : 'bad'), text: fmtDelta(def.fmt, d) + ' vs baseline' }));
      if (st !== null && Math.abs(st - b) > 1e-9) {
        const closed = (v - b) / (st - b), w = Math.max(0, Math.min(1, closed));
        kids.push(el('div', { class: 'bar' }, [el('i', { style: 'width:' + (100 * w).toFixed(1) + '%' })]));
        kids.push(el('div', { class: 'psub', text: (closed < 0 ? 'Behind the baseline' : Math.round(100 * closed) + '% of the gap to stores closed') }));
      }
    }
    kids.push(el('div', { class: 'psub', text: 'Baseline ' + fmt(def.fmt, b, m) + ' · Stores ' + fmt(def.fmt, st, m) + (prev ? ' · Last wk ' + fmt(def.fmt, prev[def.k], m) : '') }));
    return el('div', { class: 'card tk-tile' + (metric === def.k ? ' on' : ''), role: 'button', tabindex: '0', title: 'Show the trend for ' + def.label, onclick: () => { metric = def.k; render(); } }, kids);
  }

  function chart(m) {
    const def = METRICS.find((x) => x.k === metric), ws = weeksFor(m), bOn = base(m, 'Online'), bSt = base(m, 'Stores');
    const wrap = el('div', { class: 'tk-chart' });
    const W = Math.max(320, Math.min(1180, (document.querySelector('.wrap') || document.body).clientWidth - 48)), H = 280, L = 64, R = 128, T = 18, B = 38;
    const vals = ws.map((r) => num(r[def.k])).concat([bOn ? num(bOn[def.k]) : null, bSt ? num(bSt[def.k]) : null].filter((x) => x !== null));
    let lo = Math.min.apply(null, vals), hi = Math.max.apply(null, vals);
    const nonNeg = isFinite(lo) && lo >= 0;
    if (!isFinite(lo)) { lo = 0; hi = 1; } const pad = (hi - lo) * 0.12 || Math.abs(hi) * 0.1 || 1; lo -= pad; hi += pad;
    if (nonNeg) lo = Math.max(0, lo);
    const n = Math.max(ws.length, 1);
    const X = (i) => L + (n === 1 ? (W - L - R) / 2 : (W - L - R) * i / (n - 1));
    const Y = (v) => T + (H - T - B) * (1 - (v - lo) / (hi - lo));
    const NS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H); svg.setAttribute('width', '100%'); svg.setAttribute('role', 'img');
    svg.setAttribute('aria-label', def.label + ' by week for ' + NAMES[m] + ', against the online baseline and the store level');
    const add = (tag, a, text) => { const e = document.createElementNS(NS, tag); for (const k in a) e.setAttribute(k, a[k]); if (text !== undefined) e.textContent = text; svg.appendChild(e); return e; };
    for (let i = 0; i <= 4; i++) {
      const v = lo + (hi - lo) * i / 4, y = Y(v);
      add('line', { x1: L, x2: W - R, y1: y, y2: y, stroke: '#2B2B2B' });
      add('text', { x: L - 8, y: y + 4, 'text-anchor': 'end', class: 'tk-ax' }, fmt(def.fmt, v, m).replace(' pts', ''));
    }
    [[bSt, 'Stores', '#F2F1EC'], [bOn, 'Online baseline', '#7C7B76']].forEach(([b, lab, col]) => {
      if (!b) return; const y = Y(num(b[def.k]));
      add('line', { x1: L, x2: W - R, y1: y, y2: y, stroke: col, 'stroke-dasharray': '6 6', 'stroke-width': 2 });
      add('text', { x: W - R + 8, y: y + 4, class: 'tk-lab', fill: col }, lab + ' ' + fmt(def.fmt, b[def.k], m));
    });
    if (ws.length > 1) add('polyline', { points: ws.map((r, i) => X(i) + ',' + Y(num(r[def.k]))).join(' '), fill: 'none', stroke: '#FAE31D', 'stroke-width': 3 });
    const tip = el('div', { class: 'tk-tip', hidden: true });
    ws.forEach((r, i) => {
      const cx = X(i), cy = Y(num(r[def.k]));
      add('circle', { cx: cx, cy: cy, r: 5, fill: '#FAE31D' });
      if (ws.length <= 16 || i % Math.ceil(ws.length / 16) === 0 || i === ws.length - 1) add('text', { x: cx, y: H - 12, 'text-anchor': 'middle', class: 'tk-ax' }, 'Wk ' + r.Week);
      const hit = add('circle', { cx: cx, cy: cy, r: 16, fill: 'transparent' });
      hit.addEventListener('pointerenter', () => {
        tip.textContent = 'Wk ' + r.Week + ' ' + r.Year + ': ' + fmt(def.fmt, r[def.k], m) + (bOn ? ' (' + fmtDelta(def.fmt, num(r[def.k]) - num(bOn[def.k])) + ' vs baseline)' : '');
        tip.hidden = false; tip.style.left = Math.min(cx / W * 100, 70) + '%'; tip.style.top = (cy / H * 280 - 44) + 'px';
      });
      hit.addEventListener('pointerleave', () => { tip.hidden = true; });
    });
    if (!ws.length) add('text', { x: (W - R + L) / 2, y: H / 2, 'text-anchor': 'middle', class: 'tk-lab', fill: '#A3A29C' }, 'No weeks uploaded for ' + NAMES[m] + ' yet');
    wrap.append(svg, tip);
    return wrap;
  }

  function table(m) {
    const ws = weeksFor(m).slice().reverse(), bOn = base(m, 'Online'), bSt = base(m, 'Stores');
    const cols = [['Baskets', 'baskets', 'int'], ['Net sales', 'net_sales', 'money'], ['UPT', 'upt', 'dec2'], ['Footwear + acc.', 'fw_plus_acc_pct', 'pct'], ['FW → any acc.', 'fw_acc_pct', 'pct'],
      ['FW → shoe care', 'fw_shoecare_pct', 'pct'], ['FW → socks', 'fw_socks_pct', 'pct'], ['Single-unit', 'single_unit_pct', 'pct'], ['Crep % sales', 'crep_sales_pct', 'pct2'], ['Acc. % sales', 'acc_sales_pct', 'pct'], ['Returns %', 'return_pct', 'pct']];
    const t = el('table', { class: 'tk-table' }, [el('tr', {}, [el('th', { text: 'Period' })].concat(cols.map((c) => el('th', { text: c[0] }))).concat([el('th', { text: 'Uploaded by' })]))]);
    const better = (k) => (METRICS.find((x) => x.k === k) || {}).better;
    const rowEl = (r, label, cls) => el('tr', { class: cls || '' }, [el('td', { text: label })].concat(cols.map((c) => {
      let klass = '';
      const b = better(c[1]);
      if (cls === undefined && bOn && b) { const d = num(r[c[1]]) - num(bOn[c[1]]); klass = Math.abs(d) < 1e-9 ? '' : (b === 'up' ? d > 0 : d < 0) ? 'good' : 'bad'; }
      return el('td', { class: klass, text: (c[1] === 'baskets' || c[1] === 'net_sales') && isBase(r) ? 'YTD' : fmt(c[2], r[c[1]], m) });
    })).concat([el('td', { class: 'muted', text: isBase(r) ? '' : (r['Uploaded by'] || '') })]));
    if (bSt) t.appendChild(rowEl(bSt, 'Stores · baseline', 'tk-base'));
    if (bOn) t.appendChild(rowEl(bOn, 'Online · baseline', 'tk-base'));
    ws.forEach((r) => t.appendChild(rowEl(r, 'Wk ' + r.Week + ' ' + r.Year)));
    return el('div', { class: 'tw' }, [t]);
  }

  function render() {
    const v = $('tkBody'); if (!v) return;
    document.querySelectorAll('.tk-mkt').forEach((b) => b.setAttribute('aria-pressed', String(b.dataset.m === market)));
    v.replaceChildren();
    if (!loaded) { v.appendChild(el('p', { class: 'tk-note', text: 'Loading…' })); return; }
    const ws = weeksFor(market), latest = ws[ws.length - 1] || null, prev = ws[ws.length - 2] || null;
    const bOn = base(market, 'Online'), bSt = base(market, 'Stores');
    v.appendChild(el('h2', { class: 'tk-h', text: latest ? NAMES[market] + ' · week ' + latest.Week + ' ' + latest.Year + ' vs baseline' : NAMES[market] + ' · no weeks uploaded yet' }));
    if (!latest) v.appendChild(el('p', { class: 'tk-note', text: 'Upload this market’s weekly online file above. Until then, the tiles show the baseline only.' }));
    v.appendChild(el('div', { class: 'tk-tiles' }, METRICS.filter((d) => d.tile).map((d) => tile(market, d, latest, prev, bOn, bSt))));
    const sel = el('select', { class: 'tk-sel', 'aria-label': 'Metric to chart' }, METRICS.map((d) => el('option', { value: d.k, text: d.label, selected: d.k === metric })));
    sel.addEventListener('change', () => { metric = sel.value; render(); });
    v.appendChild(el('div', { class: 'tk-row' }, [el('h2', { class: 'tk-h', text: 'Trend by week' }), sel]));
    v.appendChild(el('div', { class: 'card' }, [chart(market)]));
    v.appendChild(el('h2', { class: 'tk-h', text: 'Every week' }));
    v.appendChild(el('p', { class: 'tk-note', text: 'Green is better than the online baseline and red is worse. Baseline rows are year to date, so their basket and sales totals aren’t weekly.' }));
    v.appendChild(table(market));
  }

  // ---------- wiring ----------
  $('tkYear').value = new Date().getFullYear();
  $('tkFile').addEventListener('change', (e) => { if (e.target.files.length) onFiles(Array.from(e.target.files)); });
  const drop = $('tkDrop');
  ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => { const f = Array.from(e.dataTransfer.files || []); if (f.length) onFiles(f); });
  document.querySelectorAll('.tk-mkt').forEach((b) => b.addEventListener('click', () => { market = b.dataset.m; render(); }));
  let resizeT; window.addEventListener('resize', () => { clearTimeout(resizeT); resizeT = setTimeout(() => { if (!$('viewTracker').hidden) render(); }, 200); });

  window.Tracker = { show: () => { if (!loaded) { render(); load(); } else render(); }, reload: load };
})();
