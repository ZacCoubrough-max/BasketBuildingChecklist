/* Weekly basket KPIs: the same definitions as the original YTD online vs stores analysis.
 * Input: transaction lines, each { 'Article Code', Brand, Department, Category, Country, Week,
 *        Franchise, Gender, 'Total Unit Sold', Transactions, 'Total Sales' }.
 * Lines are collated into baskets by the Transactions column.
 * Works in the browser (window.TrackerCalc) and in Node (module.exports). */
(function (root) {
  'use strict';
  const SHOE_CARE_CATS = ['PROOFERS', 'SHOE LACE', 'SHOE BAG'];

  // Order matters: this is the column order of the "Weekly KPIs" tab.
  const KPI_FIELDS = [
    'net_sales', 'gross_sales', 'returns_value', 'return_pct', 'baskets', 'sale_units', 'atv', 'upt', 'aur',
    'single_unit_pct', 'multi_unit_sales_pct', 'fw_baskets', 'fw_plus_acc_pct', 'two_plus_fw_pct',
    'fw_acc_pct', 'fw_shoecare_pct', 'fw_socks_pct', 'fw_apparel_pct',
    'crep_sales', 'socks_sales', 'acc_sales', 'crep_sales_pct', 'socks_sales_pct', 'acc_sales_pct'
  ];

  const s = (v) => (v === null || v === undefined ? '' : String(v)).trim();
  const n = (v) => { const x = typeof v === 'number' ? v : parseFloat(String(v).replace(/,/g, '')); return isFinite(x) ? x : 0; };
  const pct = (a, b) => (b ? 100 * a / b : 0);
  const round = (x, d) => { const f = Math.pow(10, d); return Math.round(x * f) / f; };

  function classify(line) {
    const dept = s(line.Department).toUpperCase(), cat = s(line.Category).toUpperCase(), brand = s(line.Brand).toUpperCase();
    const fw = dept.indexOf('FOOTWEAR') >= 0, app = dept.indexOf('TEXTILES') >= 0, acc = dept === 'ACCESSORIES';
    return {
      fw: fw, app: app, acc: acc,
      care: acc && (brand === 'CREP PROTECT' || SHOE_CARE_CATS.indexOf(cat) >= 0),
      socks: acc && cat === 'SOCKS',
      crep: brand === 'CREP PROTECT',
      dept: dept, cat: cat
    };
  }

  /** Returns { kpis, categories, info } for one market-week of lines. */
  function compute(lines) {
    let net = 0, gross = 0, returns = 0, saleUnits = 0, crep = 0, socks = 0, acc = 0;
    const baskets = new Map();   // txn -> { units, value, fw, fwUnits, acc, care, socks, app }
    const cats = new Map();      // dept|cat -> { units, sales, txns:Set }
    const markets = new Set(), weeks = new Set();
    for (const line of lines) {
      const units = n(line['Total Unit Sold']), sales = n(line['Total Sales']), txn = s(line.Transactions);
      if (!txn) continue;
      const c = classify(line);
      if (line.Country !== undefined) markets.add(s(line.Country).toUpperCase());
      if (line.Week !== undefined && s(line.Week) !== '') weeks.add(n(line.Week));
      net += sales;
      if (c.crep) crep += sales;
      if (c.socks) socks += sales;
      if (c.acc) acc += sales;
      const key = c.dept + '|' + c.cat;
      let cr = cats.get(key); if (!cr) { cr = { dept: c.dept, cat: c.cat, units: 0, sales: 0, txns: new Set() }; cats.set(key, cr); }
      cr.units += units; cr.sales += sales;
      if (units < 0 || sales < 0) { returns += -Math.min(sales, 0); continue; }
      if (!(units > 0 && sales > 0)) continue;   // $0 gift-with-purchase lines are left out of baskets
      cr.txns.add(txn);
      gross += sales; saleUnits += units;
      let b = baskets.get(txn);
      if (!b) { b = { units: 0, value: 0, fw: false, fwUnits: 0, acc: false, care: false, socks: false, app: false }; baskets.set(txn, b); }
      b.units += units; b.value += sales;
      if (c.fw) { b.fw = true; b.fwUnits += units; }
      if (c.acc) b.acc = true; if (c.care) b.care = true; if (c.socks) b.socks = true; if (c.app) b.app = true;
    }
    let nb = 0, single = 0, multiSales = 0, fwB = 0, fwAcc = 0, fwCare = 0, fwSocks = 0, fwApp = 0, fwPlusAcc = 0, twoFw = 0;
    baskets.forEach((b) => {
      nb++;
      if (b.units === 1) single++; else multiSales += b.value;
      if (b.fwUnits >= 2) twoFw++;
      if (b.fw) {
        fwB++;
        if (b.acc) { fwAcc++; fwPlusAcc++; }
        if (b.care) fwCare++; if (b.socks) fwSocks++; if (b.app) fwApp++;
      }
    });
    const k = {
      net_sales: round(net, 2), gross_sales: round(gross, 2), returns_value: round(returns, 2), return_pct: round(pct(returns, gross), 2),
      baskets: nb, sale_units: saleUnits, atv: round(nb ? gross / nb : 0, 2), upt: round(nb ? saleUnits / nb : 0, 4), aur: round(saleUnits ? gross / saleUnits : 0, 2),
      single_unit_pct: round(pct(single, nb), 2), multi_unit_sales_pct: round(pct(multiSales, gross), 2), fw_baskets: fwB,
      fw_plus_acc_pct: round(pct(fwPlusAcc, nb), 2), two_plus_fw_pct: round(pct(twoFw, nb), 2),
      fw_acc_pct: round(pct(fwAcc, fwB), 2), fw_shoecare_pct: round(pct(fwCare, fwB), 2), fw_socks_pct: round(pct(fwSocks, fwB), 2), fw_apparel_pct: round(pct(fwApp, fwB), 2),
      crep_sales: round(crep, 2), socks_sales: round(socks, 2), acc_sales: round(acc, 2),
      crep_sales_pct: round(pct(crep, net), 3), socks_sales_pct: round(pct(socks, net), 3), acc_sales_pct: round(pct(acc, net), 3)
    };
    const categories = Array.from(cats.values()).map((c) => ({ Department: c.dept, Category: c.cat, Units: c.units, 'Net sales': round(c.sales, 2), Baskets: c.txns.size }))
      .sort((a, b) => b['Net sales'] - a['Net sales']);
    return { kpis: k, categories: categories, info: { lines: lines.length, markets: Array.from(markets), weeks: Array.from(weeks) } };
  }

  const api = { compute: compute, KPI_FIELDS: KPI_FIELDS, classify: classify };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.TrackerCalc = api;
})(typeof window !== 'undefined' ? window : this);
