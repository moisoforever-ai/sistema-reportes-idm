'use strict';
(function () {
  // ---------- utilidades ----------
  const $ = (s, el) => (el || document).querySelector(s);
  const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
  const nf0 = new Intl.NumberFormat('es-VE', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('es-VE', { maximumFractionDigits: 1 });
  const nfc = new Intl.NumberFormat('es-VE', { notation: 'compact', maximumFractionDigits: 1 });
  const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const monthLabel = (m) => { const p = m.split('-'); return MESES[+p[1] - 1] + ' ' + p[0]; };
  const monthShort = (m) => { const p = m.split('-'); return MESES[+p[1] - 1].slice(0, 3) + ' ' + p[0].slice(2); };

  const METRICS = {
    unidades: { label: 'Unidades', col: 'total', usd: false },
    facturas: { label: 'Facturas', col: 'factura', usd: false },
    monto: { label: 'Monto ($)', col: 'proyeccion', usd: true },
    visitas: { label: 'Visitas', col: 'visitas', usd: false }
  };
  const fmtM = (metric, v) => (METRICS[metric].usd ? '$' : '') + nf0.format(v);
  const fmtMc = (metric, v) => (METRICS[metric].usd ? '$' : '') + nfc.format(v);

  const EMP_IDX = { DAMASCO: 0, DAKA: 1, MULTIMAX: 2, 'CIRO SANCHEZ': 3, VENELECTRONICS: 4, 'SOY TECHNO': 5 };
  const empIdx = (e, i) => (e in EMP_IDX ? EMP_IDX[e] : 6 + ((i || 0) % 3));
  const empVar = (e, i) => 'var(--e' + empIdx(e, i) + ')';
  const empColor = (e, i) => getComputedStyle(document.documentElement).getPropertyValue('--e' + empIdx(e, i)).trim() || '#555';

  const LS_PCT = 'proy.pct.v1';
  const LS_DATA = 'proy.data.v1';
  const lsGet = (k) => { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } };
  const lsDel = (k) => { try { localStorage.removeItem(k); } catch (e) { /* sin almacenamiento */ } };

  // ---------- estado ----------
  const state = {
    rows: [], report: null, fileName: '', loadedAt: null, detail: null,
    pct: {}, defaultPct: 0,
    f: { mes: null, empresas: new Set(), metric: 'unidades' },
    sort: { key: 'proj', dir: -1 }, search: '', searchSedes: ''
  };
  let charts = { emp: null, trend: null };
  let cache = null;

  const keyOf = (emp, sede) => emp + '||' + sede;

  function pctInfo(emp, sede) {
    const k = keyOf(emp, sede);
    if (Object.prototype.hasOwnProperty.call(state.pct, k) && typeof state.pct[k] === 'number') {
      const same = window.DEFAULT_PCT && window.DEFAULT_PCT[k] === state.pct[k];
      return { value: state.pct[k], known: true, origin: same ? 'Excel anterior' : 'Editado' };
    }
    if (window.DEFAULT_PCT && typeof window.DEFAULT_PCT[k] === 'number') return { value: window.DEFAULT_PCT[k], known: true, origin: 'Excel anterior' };
    return { value: state.defaultPct, known: false, origin: 'Sin %' };
  }

  // ---------- lectura del consolidado ----------
  const normHead = (h) => String(h == null ? '' : h).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const cleanText = (s) => String(s == null ? '' : s).replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
  const pad = (n) => String(n).padStart(2, '0');
  const iso = (y, m, d) => y + '-' + pad(m) + '-' + pad(d);

  function parseDate(v) {
    if (v instanceof Date && !isNaN(v)) return { iso: iso(v.getFullYear(), v.getMonth() + 1, v.getDate()) };
    if (typeof v === 'number') {
      if (v < 36526 || v > 73050) return { bad: true };
      const d = new Date(Math.round((v - 25569) * 86400000));
      return { iso: iso(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()) };
    }
    if (typeof v === 'string') {
      const s = v.trim();
      let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (m) return { iso: iso(+m[1], +m[2], +m[3]) };
      m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d+)$/);
      if (m) {
        if (m[3].length === 4) return { iso: iso(+m[3], +m[2], +m[1]) };
        if (m[3].length === 2) return { iso: iso(2000 + +m[3], +m[2], +m[1]) };
        return { truncated: true, d: +m[1], m: +m[2], prefix: m[3] };
      }
    }
    return { bad: true };
  }

  function validIso(s) {
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return false;
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3];
  }

  function parseConsolidado(aoa) {
    let hi = -1, idx = {};
    for (let i = 0; i < Math.min(aoa.length, 15); i++) {
      const hs = (aoa[i] || []).map(normHead);
      if (hs.includes('empresa') && hs.includes('sucursal') && hs.includes('fecha')) {
        hi = i;
        hs.forEach((h, j) => { if (h && !(h in idx)) idx[h] = j; });
        break;
      }
    }
    if (hi < 0) return null;
    const need = ['factura', 'visitas', 'proyeccion', 'total'];
    const missing = need.filter((c) => !(c in idx));
    const raw = [];
    for (let i = hi + 1; i < aoa.length; i++) {
      const r = aoa[i];
      if (!r || r.every((c) => c == null || c === '')) continue;
      const emp = cleanText(r[idx.empresa]).toUpperCase();
      const sede = cleanText(r[idx.sucursal]).toUpperCase();
      if (!emp || !sede) continue;
      const num = (c) => { const x = Number(r[idx[c]]); return isFinite(x) ? x : 0; };
      raw.push({
        id: idx.id != null ? r[idx.id] : i + 1, fila: i + 1, empresa: emp, sede: sede,
        factura: num('factura'), visitas: num('visitas'), proyeccion: num('proyeccion'), total: num('total'),
        fechaRaw: r[idx.fecha], pd: parseDate(r[idx.fecha])
      });
    }
    return { raw, missing };
  }

  function buildDataset(raw) {
    const report = { leidos: raw.length, fechasCorregidas: [], fechasInvalidas: [], duplicados: [], mismoDia: [], totalMenorFactura: [] };
    // año más frecuente entre las fechas válidas
    const cnt = {};
    raw.forEach((r) => { if (r.pd.iso) { const y = r.pd.iso.slice(0, 4); cnt[y] = (cnt[y] || 0) + 1; } });
    const modeYear = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
    const ok = [];
    raw.forEach((r) => {
      let f = r.pd.iso;
      if (!f && r.pd.truncated && modeYear && modeYear.startsWith(r.pd.prefix)) {
        f = iso(+modeYear, r.pd.m, r.pd.d);
        if (validIso(f)) report.fechasCorregidas.push({ id: r.id, empresa: r.empresa, sede: r.sede, antes: String(r.fechaRaw), despues: f });
        else f = null;
      }
      if (!f || !validIso(f)) { report.fechasInvalidas.push({ id: r.id, empresa: r.empresa, sede: r.sede, valor: String(r.fechaRaw) }); return; }
      ok.push({ id: r.id, empresa: r.empresa, sede: r.sede, factura: r.factura, visitas: r.visitas, proyeccion: r.proyeccion, total: r.total, fecha: f });
    });
    const seen = new Map(); const rows = [];
    ok.forEach((r) => {
      const k = [r.empresa, r.sede, r.factura, r.visitas, r.proyeccion, r.total, r.fecha].join('|');
      if (seen.has(k)) report.duplicados.push({ id: r.id, igualA: seen.get(k), empresa: r.empresa, sede: r.sede, fecha: r.fecha });
      else { seen.set(k, r.id); rows.push(r); }
    });
    const byDay = new Map();
    rows.forEach((r) => { const k = r.empresa + '||' + r.sede + '||' + r.fecha; (byDay.get(k) || byDay.set(k, []).get(k)).push(r); });
    byDay.forEach((list) => { if (list.length > 1) report.mismoDia.push({ empresa: list[0].empresa, sede: list[0].sede, fecha: list[0].fecha, ids: list.map((x) => x.id).join(', '), totales: list.map((x) => x.total).join(' / ') }); });
    rows.forEach((r) => { if (r.total < r.factura) report.totalMenorFactura.push({ id: r.id, empresa: r.empresa, sede: r.sede, fecha: r.fecha, factura: r.factura, total: r.total }); });
    return { rows, report };
  }

  async function loadConsolidado(file) {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      let parsed = null;
      for (const name of wb.SheetNames) {
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
        parsed = parseConsolidado(aoa);
        if (parsed) break;
      }
      if (!parsed) return showMsg('No encontré las columnas Empresa, Sucursal y Fecha en el archivo. Revisa que la primera hoja tenga encabezados.', true);
      if (parsed.missing.length) return showMsg('Faltan columnas en el archivo: ' + parsed.missing.join(', ') + '.', true);
      if (!parsed.raw.length) return showMsg('El archivo no tiene filas de datos.', true);
      const ds = buildDataset(parsed.raw);
      if (!ds.rows.length) return showMsg('Ninguna fila tiene una fecha válida.', true);
      state.rows = ds.rows; state.report = ds.report; state.fileName = file.name; state.loadedAt = new Date().toISOString();
      cache = null;
      const saved = lsSet(LS_DATA, { rows: state.rows, report: state.report, fileName: state.fileName, loadedAt: state.loadedAt });
      initFilters(true);
      renderAll();
      const r = state.report;
      showMsg('Cargué ' + nf0.format(r.leidos) + ' registros de "' + file.name + '". Quité ' + r.duplicados.length + ' duplicados exactos' +
        (r.fechasCorregidas.length ? ', corregí ' + r.fechasCorregidas.length + ' fecha(s)' : '') +
        (r.fechasInvalidas.length ? ' y dejé fuera ' + r.fechasInvalidas.length + ' con fecha inválida' : '') + '.' +
        (saved ? '' : ' No se pudo guardar en este navegador; tendrás que cargar el archivo otra vez al recargar la página.'), false);
    } catch (e) {
      showMsg('No pude leer el archivo: ' + e.message, true);
    }
  }

  function showMsg(t, err) {
    const el = $('#msg');
    el.textContent = t; el.className = 'msg' + (err ? ' err' : ''); el.hidden = false;
  }

  // ---------- cálculo ----------
  const daysIn = (mes) => { const p = mes.split('-').map(Number); return new Date(Date.UTC(p[0], p[1], 0)).getUTCDate(); };

  function compute() {
    if (cache) return cache;
    const map = new Map();
    state.rows.forEach((r) => {
      const mes = r.fecha.slice(0, 7); const k = r.empresa + '||' + r.sede + '||' + mes;
      let g = map.get(k);
      if (!g) { g = { empresa: r.empresa, sede: r.sede, mes: mes, n: 0, days: daysIn(mes), obs: { facturas: 0, unidades: 0, monto: 0, visitas: 0 } }; map.set(k, g); }
      g.n++; g.obs.facturas += r.factura; g.obs.unidades += r.total; g.obs.monto += r.proyeccion; g.obs.visitas += r.visitas;
    });
    const groups = Array.from(map.values());
    groups.forEach((g) => {
      const p = pctInfo(g.empresa, g.sede); g.k = p.value; g.known = p.known; g.proj = {};
      Object.keys(METRICS).forEach((m) => {
        const S = g.obs[m]; const total = S * (1 + g.k); const prom = total / (g.n + 1);
        g.proj[m] = total + prom * Math.max(0, g.days - g.n);
      });
    });
    const sedes = new Map();
    state.rows.forEach((r) => sedes.set(keyOf(r.empresa, r.sede), { empresa: r.empresa, sede: r.sede }));
    cache = { groups: groups, sedes: Array.from(sedes.values()), meses: Array.from(new Set(groups.map((g) => g.mes))).sort(), empresas: [] };
    const ce = {}; state.rows.forEach((r) => { ce[r.empresa] = (ce[r.empresa] || 0) + 1; });
    cache.empresas = Object.keys(ce).sort((a, b) => ce[b] - ce[a]);
    return cache;
  }
  const invalidate = () => { cache = null; };

  // ---------- filtros ----------
  function initFilters(reset) {
    const c = compute();
    if (reset || !state.f.mes || !c.meses.includes(state.f.mes)) state.f.mes = c.meses[c.meses.length - 1];
    if (reset || !state.f.empresas.size) state.f.empresas = new Set(c.empresas);
    else state.f.empresas = new Set(Array.from(state.f.empresas).filter((e) => c.empresas.includes(e)));
    if (!state.f.empresas.size) state.f.empresas = new Set(c.empresas);
  }

  function renderFilters() {
    const c = compute();
    $('#selMes').innerHTML = c.meses.map((m) => '<option value="' + m + '"' + (m === state.f.mes ? ' selected' : '') + '>' + monthLabel(m) + '</option>').join('');
    $('#segMetric').innerHTML = Object.keys(METRICS).map((m) =>
      '<label><input type="radio" name="metric" value="' + m + '"' + (m === state.f.metric ? ' checked' : '') + '><span>' + METRICS[m].label + '</span></label>').join('');
    $('#chipsEmp').innerHTML = c.empresas.map((e, i) =>
      '<label class="chip"><input type="checkbox" value="' + esc(e) + '"' + (state.f.empresas.has(e) ? ' checked' : '') + '><i style="background:' + empVar(e, i) + '"></i>' + esc(e) + '</label>').join('');
  }

  // ---------- resumen ----------
  function selectedGroups() {
    const c = compute();
    return c.groups.filter((g) => state.f.empresas.has(g.empresa));
  }

  function renderKpis() {
    const m = state.f.metric; const mes = state.f.mes;
    const inMes = selectedGroups().filter((g) => g.mes === mes);
    const sum = (f) => inMes.reduce((a, g) => a + f(g), 0);
    const proj = sum((g) => g.proj[m]); const obs = sum((g) => g.obs[m]);
    const pf = sum((g) => g.proj.facturas), pv = sum((g) => g.proj.visitas), pm = sum((g) => g.proj.monto);
    const sedesSel = compute().sedes.filter((s) => state.f.empresas.has(s.empresa)).length;
    const estudiadas = inMes.length; const dias = sum((g) => g.n);
    const sinPct = inMes.filter((g) => !g.known).length;
    const cards = [
      { l: 'Proyectado en ' + monthLabel(mes), v: fmtM(m, proj), s: METRICS[m].label },
      { l: 'Observado en días estudiados', v: fmtM(m, obs), s: nf0.format(dias) + ' días estudiados' },
      { l: 'Sedes estudiadas', v: nf0.format(estudiadas) + ' de ' + nf0.format(sedesSel), s: sinPct ? sinPct + ' sin % asignado' : 'todas con %' },
      { l: 'Conversión (facturas ÷ visitas)', v: pv ? nf1.format(pf / pv * 100) + '%' : '-', s: 'sobre lo proyectado' },
      { l: 'Monto por factura', v: pf ? '$' + nf0.format(pm / pf) : '-', s: 'sobre lo proyectado' }
    ];
    $('#kpis').innerHTML = cards.map((c) => '<div class="kpi"><div class="l">' + esc(c.l) + '</div><div class="v">' + esc(c.v) + '</div><div class="s">' + esc(c.s) + '</div></div>').join('');
  }

  function chartColors() {
    const cs = getComputedStyle(document.documentElement);
    return { text: cs.getPropertyValue('--text-2').trim() || '#5b6676', grid: cs.getPropertyValue('--border').trim() || '#dde1e7' };
  }

  function renderCharts() {
    if (typeof Chart === 'undefined') return;
    renderPies();
    const m = state.f.metric; const mes = state.f.mes; const c = compute(); const col = chartColors();
    Chart.defaults.color = col.text; Chart.defaults.font.family = 'system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif';
    const emps = c.empresas.filter((e) => state.f.empresas.has(e));
    const groups = selectedGroups();
    $('#tChartTrend').textContent = METRICS[m].label + ' proyectado por mes';
    const yTick = (v) => fmtMc(m, v);
    if (charts.trend) charts.trend.destroy();
    charts.trend = new Chart($('#chartTrend'), {
      type: 'bar',
      data: {
        labels: c.meses.map(monthShort),
        datasets: emps.map((e) => ({
          label: e, backgroundColor: empColor(e, c.empresas.indexOf(e)),
          data: c.meses.map((mm) => groups.filter((g) => g.empresa === e && g.mes === mm).reduce((a, g) => a + g.proj[m], 0))
        }))
      },
      options: {
        responsive: true, maintainAspectRatio: false, animation: false,
        plugins: { legend: { position: 'bottom', labels: { boxWidth: 12 } }, tooltip: { callbacks: { label: (x) => x.dataset.label + ': ' + fmtM(m, x.parsed.y) } } },
        scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, beginAtZero: true, grid: { color: col.grid }, ticks: { callback: yTick } } }
      }
    });
  }

  function tableRows() {
    const c = compute(); const m = state.f.metric; const mes = state.f.mes;
    const gm = new Map();
    c.groups.filter((g) => g.mes === mes).forEach((g) => gm.set(keyOf(g.empresa, g.sede), g));
    const rows = c.sedes.filter((s) => state.f.empresas.has(s.empresa)).map((s) => {
      const g = gm.get(keyOf(s.empresa, s.sede)); const p = pctInfo(s.empresa, s.sede);
      return { empresa: s.empresa, sede: s.sede, k: p.value, known: p.known, n: g ? g.n : 0, obs: g ? g.obs[m] : 0, proj: g ? g.proj[m] : 0 };
    });
    const total = rows.reduce((a, r) => a + r.proj, 0);
    rows.forEach((r) => { r.share = total ? r.proj / total : 0; });
    const q = state.search.trim().toUpperCase();
    const filtered = q ? rows.filter((r) => r.sede.includes(q) || r.empresa.includes(q)) : rows;
    const key = state.sort.key, dir = state.sort.dir;
    filtered.sort((a, b) => {
      const x = a[key], y = b[key];
      const r = typeof x === 'string' ? x.localeCompare(y, 'es') : x - y;
      return (r || a.sede.localeCompare(b.sede, 'es')) * dir;
    });
    return { rows: filtered, total: total };
  }

  const COLS = [
    { key: 'empresa', label: 'Empresa' }, { key: 'sede', label: 'Sede' },
    { key: 'k', label: '% sede', num: true }, { key: 'n', label: 'Días estudiados', num: true },
    { key: 'obs', label: 'Observado', num: true }, { key: 'proj', label: 'Proyectado', num: true },
    { key: 'share', label: '% del total', num: true }
  ];

  function renderTable() {
    const m = state.f.metric;
    $('#tTable').textContent = 'Detalle por sede, ' + monthLabel(state.f.mes) + ' (' + METRICS[m].label + ')';
    $('#tbl thead').innerHTML = '<tr>' + COLS.map((c) => {
      const act = state.sort.key === c.key; const arrow = act ? (state.sort.dir > 0 ? ' ▲' : ' ▼') : '';
      return '<th class="' + (c.num ? 'num' : '') + '" aria-sort="' + (act ? (state.sort.dir > 0 ? 'ascending' : 'descending') : 'none') + '"><button type="button" data-sort="' + c.key + '">' + esc(c.label) + arrow + '</button></th>';
    }).join('') + '</tr>';
    const t = tableRows();
    $('#tbl tbody').innerHTML = t.rows.map((r) =>
      '<tr><td>' + esc(r.empresa) + '</td><td>' + esc(r.sede) +
      (r.n === 0 ? '<span class="tag warn">Sin estudio</span>' : '') + (!r.known ? '<span class="tag bad">Sin %</span>' : '') +
      '</td><td class="num">' + (r.known ? nf0.format(r.k * 100) + '%' : '-') + '</td><td class="num">' + r.n +
      '</td><td class="num">' + fmtM(m, r.obs) + '</td><td class="num">' + fmtM(m, r.proj) + '</td><td class="num">' + nf1.format(r.share * 100) + '%</td></tr>').join('') ||
      '<tr><td colspan="7">No hay sedes con ese filtro.</td></tr>';
    const sumObs = t.rows.reduce((a, r) => a + r.obs, 0), sumProj = t.rows.reduce((a, r) => a + r.proj, 0), sumN = t.rows.reduce((a, r) => a + r.n, 0);
    let tf = $('#tbl tfoot'); if (!tf) { tf = document.createElement('tfoot'); $('#tbl').appendChild(tf); }
    tf.innerHTML = '<tr><td colspan="3">Total (' + t.rows.length + ' sedes)</td><td class="num">' + nf0.format(sumN) + '</td><td class="num">' + fmtM(m, sumObs) + '</td><td class="num">' + fmtM(m, sumProj) + '</td><td class="num"></td></tr>';
  }

  function downloadBlob(name, text, type) {
    const b = new Blob([text], { type: type });
    const a = document.createElement('a'); a.href = URL.createObjectURL(b); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }

  function exportCsv() {
    const m = state.f.metric; const t = tableRows();
    const q = (s) => '"' + String(s).replace(/"/g, '""') + '"';
    const lines = [['Empresa', 'Sede', 'Porcentaje sede', 'Dias estudiados', 'Observado', 'Proyectado', 'Porcentaje del total'].map(q).join(',')];
    t.rows.forEach((r) => lines.push([q(r.empresa), q(r.sede), r.known ? (r.k * 100).toFixed(0) : '', r.n, r.obs.toFixed(2), r.proj.toFixed(2), (r.share * 100).toFixed(2)].join(',')));
    downloadBlob('proyeccion_' + state.f.mes + '_' + m + '.csv', '﻿' + lines.join('\r\n'), 'text/csv;charset=utf-8');
  }

  // ---------- sedes y % ----------
  function renderSedes() {
    const c = compute();
    $('#defPct').value = Math.round(state.defaultPct * 1000) / 10;
    const seen = new Set(c.sedes.map((s) => keyOf(s.empresa, s.sede)));
    const days = {}; state.rows.forEach((r) => { const k = keyOf(r.empresa, r.sede); days[k] = (days[k] || 0) + 1; });
    const list = c.sedes.slice();
    const q = state.searchSedes.trim().toUpperCase();
    const rows = (q ? list.filter((s) => s.sede.includes(q) || s.empresa.includes(q)) : list).sort((a, b) => a.empresa.localeCompare(b.empresa, 'es') || a.sede.localeCompare(b.sede, 'es'));
    const sinPct = list.filter((s) => !pctInfo(s.empresa, s.sede).known).length;
    $('#defPctInfo').textContent = sinPct ? sinPct + ' sede(s) sin % usan este valor.' : 'Todas las sedes tienen %.';
    $('#tblSedes tbody').innerHTML = rows.length ? rows.map((s) => {
      const p = pctInfo(s.empresa, s.sede); const k = keyOf(s.empresa, s.sede);
      const tag = p.origin === 'Sin %' ? '<span class="tag bad">Sin %</span>' : p.origin === 'Editado' ? '<span class="tag ok">Editado</span>' : '<span class="tag">Excel anterior</span>';
      return '<tr><td>' + esc(s.empresa) + '</td><td>' + esc(s.sede) + '</td><td class="num"><input type="number" min="0" max="500" step="1" data-k="' + esc(k) + '" value="' + (p.known ? Math.round(p.value * 1000) / 10 : '') + '" placeholder="-" aria-label="Porcentaje de ' + esc(s.sede) + '"> %</td><td>' + tag + '</td><td class="num">' + (days[k] || 0) + '</td></tr>';
    }).join('') : '<tr><td colspan="5">' + (state.rows.length ? 'No hay sedes con ese filtro.' : 'Carga el consolidado para ver las sedes.') + '</td></tr>';
    void seen;
  }

  function savePct() { lsSet(LS_PCT, { pct: state.pct, defaultPct: state.defaultPct }); if (puenteOn()) schedulePctPush(); }

  // ---------- calidad ----------
  function qTable(title, cols, rows, note) {
    if (!rows.length) return '';
    const shown = rows.slice(0, 60);
    return '<div class="card"><h3>' + esc(title) + ' (' + rows.length + ')</h3>' + (note ? '<p class="hint">' + esc(note) + '</p>' : '') +
      '<div class="tbl-wrap"><table class="tbl"><thead><tr>' + cols.map((c) => '<th>' + esc(c[1]) + '</th>').join('') + '</tr></thead><tbody>' +
      shown.map((r) => '<tr>' + cols.map((c) => '<td>' + esc(r[c[0]]) + '</td>').join('') + '</tr>').join('') +
      '</tbody></table></div>' + (rows.length > shown.length ? '<p class="hint">Se muestran los primeros ' + shown.length + '.</p>' : '') + '</div>';
  }

  function detQuality() {
    const d = state.detail; if (!d) return '';
    const r = d.report;
    const pm = new Map(), dm = new Map(); const sm = new Set();
    state.rows.forEach((x) => { const k = x.empresa + '|' + x.sede + '|' + x.fecha; pm.set(k, (pm.get(k) || 0) + x.total); });
    d.rows.forEach((x) => { const k = x.empresa + '|' + x.sede + '|' + x.fecha; dm.set(k, (dm.get(k) || 0) + x.cant); sm.add(x.empresa + '||' + x.sede + '||' + x.fecha.slice(0, 7)); });
    const ratios = []; dm.forEach((v, k) => { if (pm.has(k) && v > 0) ratios.push(pm.get(k) / v); });
    ratios.sort((a, b) => a - b); const med = ratios.length ? ratios[Math.floor(ratios.length / 2)] : null;
    const have = new Set(compute().groups.map((g) => g.empresa + '||' + g.sede + '||' + g.mes));
    let sinPortal = 0; sm.forEach((k) => { if (!have.has(k)) sinPortal++; });
    const fechas = d.rows.map((x) => x.fecha).sort();
    const cards = [['Filas de detalle leídas', r.leidos], ['Filas usadas', d.rows.length], ['Duplicados exactos quitados', r.duplicados || 0], ['Fechas inválidas (excluidas)', r.fechasInvalidas],
      ['Sedes que no coinciden con el portal', r.sinSede.length], ['Filas sin categoría', r.sinCategoria], ['Filas sin marca reconocida', r.sinMarca || 0],
      ['Sede-días en ambos archivos', ratios.length], ['Total del portal ÷ unidades del detalle (mediana)', med == null ? '-' : nf1.format(med)], ['Meses de sede sin consolidado', sinPortal]];
    return '<h3 style="margin:1.2rem 0 .6rem">Detalle de productos</h3><p class="hint">Fechas del detalle: ' + (fechas.length ? fechas[0] + ' a ' + fechas[fechas.length - 1] : '-') + (r.marcaInferida ? '. La marca se deduce del nombre del producto con la lista de marcas del sistema.' : '.') + '</p><div class="qgrid">' + cards.map((c) => '<div class="kpi"><div class="l">' + esc(c[0]) + '</div><div class="v">' + (typeof c[1] === 'number' ? nf0.format(c[1]) : esc(c[1])) + '</div></div>').join('') + '</div>' +
      qTable('Sedes del detalle que no coinciden con el portal', [['sede', 'Empresa | Sede en el detalle'], ['filas', 'Filas']], r.sinSede, 'Estas filas no tienen estudio de sede en el portal y no reciben proyección. Revisa que el nombre coincida con el del consolidado.');
  }

  function renderCalidad() {
    const el = $('#calidadBody'); const r = state.report;
    if (!r) { el.innerHTML = '<div class="card"><p class="hint">Carga el consolidado para ver el control de calidad.</p></div>'; return; }
    const cards = [
      ['Registros leídos', r.leidos], ['Duplicados exactos quitados', r.duplicados.length], ['Fechas corregidas', r.fechasCorregidas.length],
      ['Fechas inválidas (excluidas)', r.fechasInvalidas.length], ['Misma sede y día, valores distintos', r.mismoDia.length], ['Total menor que facturas', r.totalMenorFactura.length]
    ];
    el.innerHTML = detQuality() + '<h3 style="margin:1.2rem 0 .6rem">Consolidado del portal</h3><div class="qgrid">' + cards.map((c) => '<div class="kpi"><div class="l">' + esc(c[0]) + '</div><div class="v">' + nf0.format(c[1]) + '</div></div>').join('') + '</div>' +
      qTable('Duplicados exactos quitados', [['id', 'id'], ['igualA', 'Igual al id'], ['empresa', 'Empresa'], ['sede', 'Sede'], ['fecha', 'Fecha']], r.duplicados, 'Mismo contenido que otro registro. Se conservó el primero.') +
      qTable('Fechas corregidas', [['id', 'id'], ['empresa', 'Empresa'], ['sede', 'Sede'], ['antes', 'Decía'], ['despues', 'Se tomó como']], r.fechasCorregidas, 'La fecha venía cortada. Se completó con el año más frecuente del archivo.') +
      qTable('Fechas inválidas', [['id', 'id'], ['empresa', 'Empresa'], ['sede', 'Sede'], ['valor', 'Valor']], r.fechasInvalidas, 'Estos registros no se usan en el cálculo hasta corregir la fecha en el portal.') +
      qTable('Misma sede y mismo día con valores distintos', [['empresa', 'Empresa'], ['sede', 'Sede'], ['fecha', 'Fecha'], ['ids', 'ids'], ['totales', 'Totales']], r.mismoDia, 'Se conservan como estudios separados, y cada uno cuenta como un día estudiado.') +
      qTable('Total menor que facturas', [['id', 'id'], ['empresa', 'Empresa'], ['sede', 'Sede'], ['fecha', 'Fecha'], ['factura', 'Facturas'], ['total', 'Total']], r.totalMenorFactura, 'Cada factura tiene al menos una unidad, así que el total no debería ser menor. Revisa estos registros en el portal. Se usan tal como están.');
  }

  // ---------- convertidor ----------
  const conv = { wb: null, name: '', records: [], cols: [] };
  function convRebuild() {
    if (!conv.wb) return;
    const ws = conv.wb.Sheets[$('#convSheet').value];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const useHead = $('#optHeader').checked, dates = $('#optDates').checked, trim = $('#optTrim').checked;
    let head, body;
    const width = aoa.reduce((a, r) => Math.max(a, r.length), 0);
    if (useHead && aoa.length) {
      const used = {};
      head = Array.from({ length: width }, (_, i) => {
        let h = cleanText(aoa[0][i]) || 'columna' + (i + 1);
        if (used[h]) { used[h]++; h = h + '_' + used[h]; } else used[h] = 1;
        return h;
      });
      body = aoa.slice(1);
    } else { head = Array.from({ length: width }, (_, i) => 'columna' + (i + 1)); body = aoa; }
    body = body.filter((r) => r && r.some((c) => c != null && c !== ''));
    const dateCol = head.map((h) => /fecha|date/i.test(h));
    const recs = body.map((r) => {
      const o = {};
      head.forEach((h, i) => {
        let v = r[i] == null ? null : r[i];
        if (typeof v === 'string' && trim) v = v.replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
        if (dates && dateCol[i] && v != null) { const p = parseDate(v); if (p.iso) v = p.iso; }
        o[h] = v;
      });
      return o;
    });
    conv.records = recs; conv.cols = head;
    $('#convOut').hidden = false;
    $('#convInfo').textContent = nf0.format(recs.length) + ' filas y ' + head.length + ' columnas.';
    $('#convTbl thead').innerHTML = '<tr>' + head.map((h) => '<th>' + esc(h) + '</th>').join('') + '</tr>';
    $('#convTbl tbody').innerHTML = recs.slice(0, 15).map((o) => '<tr>' + head.map((h) => '<td>' + esc(o[h]) + '</td>').join('') + '</tr>').join('');
    $('#convJson').textContent = JSON.stringify(recs.slice(0, 20), null, 2) + (recs.length > 20 ? '\n... (' + (recs.length - 20) + ' filas más en el archivo descargado)' : '');
  }

  async function convLoad(file) {
    try {
      const buf = await file.arrayBuffer();
      conv.wb = XLSX.read(buf, { type: 'array' }); conv.name = file.name;
      const sel = $('#convSheet');
      sel.innerHTML = conv.wb.SheetNames.map((n) => '<option>' + esc(n) + '</option>').join(''); sel.disabled = false;
      $('#convName').textContent = file.name;
      convRebuild();
    } catch (e) { showMsg('No pude leer el archivo: ' + e.message, true); }
  }

  // ---------- detalle de productos ----------
  const LS_DET = 'proy.det.v1';
  const DET_ALIAS = {
    fecha: ['fecha', 'date', 'dia', 'fechaestudio'],
    empresa: ['empresa', 'compania', 'company'],
    sede: ['sucursal', 'sede', 'tienda', 'store', 'sucursales'],
    categoria: ['categoria', 'category', 'linea', 'familia', 'rubro', 'tipo'],
    marca: ['marca', 'brand'],
    producto: ['producto', 'articulo', 'item', 'product', 'descripcion', 'nombre'],
    modelo: ['modelo', 'model', 'sku', 'codigo'],
    cant: ['cantidad', 'unidades', 'qty', 'cant', 'vendidos', 'unidadesvendidas'],
    monto: ['monto', 'importe', 'montousd', 'usd', 'dolares', 'valor', 'ventausd'],
    precio: ['precio', 'price', 'preciounitario']
  };
  const estadoOf = (sede) => String(sede).split(' - ')[0];

  function detectDet(headers) {
    const hs = headers.map(normHead); const idx = {};
    Object.keys(DET_ALIAS).forEach((f) => {
      for (const a of DET_ALIAS[f]) { const j = hs.indexOf(a); if (j >= 0) { idx[f] = j; break; } }
    });
    return idx;
  }

  function makeBrandFinder() {
    const list = (window.BRANDS || []).slice().sort((a, b) => b.length - a.length || a.localeCompare(b));
    const res = list.map((b) => new RegExp('(?<![A-Z0-9])' + b.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?![A-Z0-9])'));
    const memo = new Map();
    return (name) => {
      if (memo.has(name)) return memo.get(name);
      const u = String(name).replace(/\([^)]*\)/g, ' ');
      let found = '';
      for (let i = 0; i < res.length; i++) if (res[i].test(u)) { found = list[i]; break; }
      memo.set(name, found); return found;
    };
  }

  function parseDetalle(aoa) {
    let hi = -1, idx = null;
    for (let i = 0; i < Math.min(aoa.length, 15); i++) {
      const d = detectDet((aoa[i] || []).map((x) => (x == null ? '' : x)));
      if (d.fecha != null && d.sede != null && d.cant != null) { hi = i; idx = d; break; }
    }
    if (hi < 0) return null;
    const c = compute();
    const byKey = new Set(c.sedes.map((s) => keyOf(s.empresa, s.sede)));
    const cnt = {}; state.rows.forEach((r) => { const y = r.fecha.slice(0, 4); cnt[y] = (cnt[y] || 0) + 1; });
    const modeYear = Object.keys(cnt).sort((a, b) => cnt[b] - cnt[a])[0];
    const rep = { leidos: 0, fechasInvalidas: 0, sinSede: new Map(), sinCategoria: 0, sinMonto: idx.monto == null && idx.precio == null, duplicados: 0, marcaInferida: idx.marca == null, sinMarca: 0 };
    const rows = []; const seenDet = new Set();
    const inferBrand = makeBrandFinder();
    const resolve = (emp, sede) => {
      if (emp && byKey.has(keyOf(emp, sede))) return { empresa: emp, sede: sede, ok: true };
      const cand = c.sedes.filter((s) => (!emp || s.empresa === emp) && (s.sede === sede || s.sede.endsWith(' - ' + sede)));
      if (cand.length === 1) return { empresa: cand[0].empresa, sede: cand[0].sede, ok: true };
      return { empresa: emp || '(sin empresa)', sede: sede, ok: false };
    };
    for (let i = hi + 1; i < aoa.length; i++) {
      const r = aoa[i];
      if (!r || r.every((x) => x == null || x === '')) continue;
      rep.leidos++;
      const pd = parseDate(r[idx.fecha]); let f = pd.iso;
      if (!f && pd.truncated && modeYear && modeYear.startsWith(pd.prefix)) { f = iso(+modeYear, pd.m, pd.d); }
      if (!f || !validIso(f)) { rep.fechasInvalidas++; continue; }
      const sedeRaw = cleanText(r[idx.sede]).toUpperCase(); if (!sedeRaw) continue;
      const empRaw = idx.empresa != null ? cleanText(r[idx.empresa]).toUpperCase() : '';
      const rs = resolve(empRaw, sedeRaw);
      if (!rs.ok) { const k = rs.empresa + ' | ' + rs.sede; rep.sinSede.set(k, (rep.sinSede.get(k) || 0) + 1); }
      const nn = (j) => { const x = Number(r[j]); return j != null && isFinite(x) ? x : 0; };
      const cant = nn(idx.cant);
      let monto = idx.monto != null ? nn(idx.monto) : (idx.precio != null ? nn(idx.precio) * cant : 0);
      if (!cant && !monto) continue;
      const up = (j) => (j != null ? cleanText(r[j]).toUpperCase() : '');
      let prod = up(idx.producto); const mod = up(idx.modelo);
      if (prod && mod && prod !== mod && !prod.includes(mod)) prod = prod + ' · ' + mod; else if (!prod) prod = mod;
      const cat = up(idx.categoria); if (!cat) rep.sinCategoria++;
      const dk = [f, rs.empresa, rs.sede, cat, prod, cant, monto].join('|');
      if (seenDet.has(dk)) { rep.duplicados++; continue; }
      seenDet.add(dk);
      const marca = idx.marca != null ? up(idx.marca) : inferBrand(prod);
      if (!marca) rep.sinMarca++;
      rows.push({ fecha: f, empresa: rs.empresa, sede: rs.sede, ok: rs.ok, categoria: cat, marca: marca, producto: prod, cant: cant, monto: monto });
    }
    return { rows: rows, report: rep, cols: { categoria: idx.categoria != null, marca: true, producto: idx.producto != null || idx.modelo != null, monto: idx.monto != null || idx.precio != null } };
  }

  function detToAoa(wb, jsonText) {
    if (jsonText != null) {
      let j = JSON.parse(jsonText); if (j && !Array.isArray(j)) { j = j.rows || j.data || j.registros || []; }
      if (!Array.isArray(j) || !j.length || typeof j[0] !== 'object') throw new Error('el JSON debe ser una lista de registros');
      const keys = []; j.slice(0, 200).forEach((o) => Object.keys(o).forEach((k) => { if (!keys.includes(k)) keys.push(k); }));
      return [keys].concat(j.map((o) => keys.map((k) => (o[k] == null ? null : o[k]))));
    }
    return null;
  }

  function makeDetail(parsed, name) {
    const r = parsed.report;
    return {
      rows: parsed.rows, cols: parsed.cols, fileName: name,
      report: { leidos: r.leidos, fechasInvalidas: r.fechasInvalidas, sinSede: Array.from(r.sinSede.entries()).map((e) => ({ sede: e[0], filas: e[1] })), sinCategoria: r.sinCategoria, sinMonto: r.sinMonto, duplicados: r.duplicados, marcaInferida: r.marcaInferida, sinMarca: r.sinMarca }
    };
  }

  async function loadDetalle(file) {
    try {
      if (!state.rows.length) return showMsg('Carga primero el consolidado del portal; el detalle se alinea con él.', true);
      let parsed = null; let heads = [];
      if (/\.json$/i.test(file.name)) {
        parsed = parseDetalle(detToAoa(null, await file.text()));
      } else {
        const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
        for (const name of wb.SheetNames) {
          const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
          if (aoa.length && !heads.length) heads = (aoa[0] || []).map(cleanText).filter(Boolean);
          parsed = parseDetalle(aoa); if (parsed) break;
        }
      }
      if (!parsed) return showMsg('No encontré las columnas mínimas del detalle: Fecha, Sucursal (o Sede) y Cantidad (o Unidades).' + (heads.length ? ' Encabezados que vi: ' + heads.slice(0, 12).join(', ') + '.' : ''), true);
      if (!parsed.rows.length) return showMsg('El detalle no tiene filas válidas.', true);
      state.detail = makeDetail(parsed, file.name);
      const r = parsed.report;
      const saved = lsSet(LS_DET, state.detail);
      renderAll();
      showMsg('Cargué ' + nf0.format(parsed.rows.length) + ' filas de detalle de "' + file.name + '".' +
        (r.sinSede.size ? ' ' + r.sinSede.size + ' sede(s) del detalle no coinciden con el portal (ver Calidad de datos); se proyectan sin ajuste.' : '') +
        (r.sinMonto ? ' El archivo no trae monto: solo se proyectan unidades.' : '') +
        (saved ? '' : ' No cupo en el almacenamiento del navegador; tendrás que cargarlo otra vez al recargar la página.'), false);
    } catch (e) { showMsg('No pude leer el detalle: ' + e.message, true); }
  }

  // ---------- motor de proyección por filtros ----------
  const hasDet = () => !!(state.detail && state.detail.rows && state.detail.rows.length);
  function overlapFrac(mes, d1, d2) {
    const n = daysIn(mes); const s = mes + '-01', e = mes + '-' + pad(n);
    const a = d1 > s ? d1 : s, b = d2 < e ? d2 : e; if (a > b) return 0;
    return (+b.slice(8) - +a.slice(8) + 1) / n;
  }
  const sel = (set, v) => !set || !set.size || set.has(v);
  function dimKey(o, group, mes) {
    switch (group) {
      case 'categoria': return o.categoria || '(sin categoría)';
      case 'marca': return o.marca || '(sin marca)';
      case 'producto': return o.producto || '(sin producto)';
      case 'empresa': return o.empresa;
      case 'estado': return estadoOf(o.sede);
      case 'sede': return o.empresa + '||' + o.sede;
      default: return mes;
    }
  }
  const DIMS = {
    empresa: [['Empresa']], estado: [['Estado']], sede: [['Empresa'], ['Sede']], mes: [['Mes']],
    categoria: [['Categoría']], marca: [['Marca']], producto: [['Producto / modelo']]
  };
  function dimParts(key, group) {
    if (group === 'sede') return key.split('||');
    if (group === 'mes') return [monthLabel(key)];
    return [key];
  }

  // o: {d1,d2,empresas,estados,sedes,cats,marcas,prods,metric,group}
  function runReport(o) {
    const c = compute(); const m = o.metric; const col = METRICS[m].col;
    const detailMode = (o.cats && o.cats.size) || (o.marcas && o.marcas.size) || (o.prods && o.prods.size) || ['categoria', 'marca', 'producto'].includes(o.group);
    const out = new Map(); let sinAjuste = 0; const notes = [];
    const passSede = (emp, sede) => sel(o.empresas, emp) && sel(o.estados, estadoOf(sede)) && sel(o.sedes, keyOf(emp, sede));
    if (detailMode) {
      if (!hasDet()) return { error: 'Para filtrar o agrupar por categoría, marca o producto primero carga el detalle de productos (botón de arriba).' };
      if (m !== 'unidades' && m !== 'monto') return { error: 'Con categoría, marca o producto solo se puede medir Unidades o Monto.' };
      if (m === 'monto' && state.detail.report && state.detail.report.sinMonto) return { error: 'El detalle cargado no trae monto. Usa la medida Unidades.' };
      const gm = new Map(); c.groups.forEach((g) => gm.set(g.empresa + '||' + g.sede + '||' + g.mes, g));
      const den = new Map();
      state.detail.rows.forEach((r) => { const k = r.empresa + '||' + r.sede + '||' + r.fecha.slice(0, 7); let e = den.get(k); if (!e) { e = { unidades: 0, monto: 0 }; den.set(k, e); } e.unidades += r.cant; e.monto += r.monto; });
      state.detail.rows.forEach((r) => {
        if (!passSede(r.empresa, r.sede)) return;
        if (!sel(o.cats, r.categoria) || !sel(o.marcas, r.marca) || !sel(o.prods, r.producto)) return;
        const mes = r.fecha.slice(0, 7); const frac = overlapFrac(mes, o.d1, o.d2); if (!frac) return;
        const val = m === 'monto' ? r.monto : r.cant;
        const sm = r.empresa + '||' + r.sede + '||' + mes; const g = gm.get(sm); const dn = den.get(sm)[m];
        let contrib = 0;
        if (g && dn > 0) contrib = val / dn * g.proj[m] * frac; else sinAjuste++;
        const k = dimKey(r, o.group, mes); let e = out.get(k); if (!e) { e = { key: k, obs: 0, proj: 0 }; out.set(k, e); }
        if (r.fecha >= o.d1 && r.fecha <= o.d2) e.obs += val;
        e.proj += contrib;
      });
      notes.push('La proyección de cada sede y mes (la del consolidado) se reparte entre sus productos según el peso de cada uno en lo observado en esa sede. Así las categorías, marcas y productos siempre suman lo mismo que el resumen.');
      if (sinAjuste) notes.push(nf0.format(sinAjuste) + ' fila(s) del detalle pertenecen a una sede y mes sin estudio en el consolidado; cuentan como observado pero no tienen proyección.');
    } else {
      const obsMap = new Map();
      state.rows.forEach((r) => {
        if (r.fecha < o.d1 || r.fecha > o.d2 || !passSede(r.empresa, r.sede)) return;
        const k = r.empresa + '||' + r.sede + '||' + r.fecha.slice(0, 7); obsMap.set(k, (obsMap.get(k) || 0) + r[col]);
      });
      c.groups.forEach((g) => {
        if (!passSede(g.empresa, g.sede)) return;
        const frac = overlapFrac(g.mes, o.d1, o.d2); if (!frac) return;
        const k = dimKey(g, o.group, g.mes); let e = out.get(k); if (!e) { e = { key: k, obs: 0, proj: 0 }; out.set(k, e); }
        e.obs += obsMap.get(g.empresa + '||' + g.sede + '||' + g.mes) || 0; e.proj += g.proj[m] * frac;
      });
    }
    notes.push('Si el rango no cubre meses completos, la proyección del mes se reparte en proporción a los días del rango.');
    const rows = Array.from(out.values());
    if (o.group === 'mes') rows.sort((a, b) => a.key.localeCompare(b.key)); else rows.sort((a, b) => b.proj - a.proj);
    const totalProj = rows.reduce((a, r) => a + r.proj, 0), totalObs = rows.reduce((a, r) => a + r.obs, 0);
    rows.forEach((r) => { r.share = totalProj ? r.proj / totalProj : 0; r.parts = dimParts(r.key, o.group); });
    return { rows: rows, totalProj: totalProj, totalObs: totalObs, notes: notes, detailMode: !!detailMode, group: o.group, metric: m };
  }

  // ---------- pasteles ----------
  const hex2rgb = (h) => { h = h.replace('#', ''); if (h.length === 3) h = h.split('').map((x) => x + x).join(''); return [0, 2, 4].map((i) => parseInt(h.substr(i, 2), 16)); };
  const rgb2hex = (a) => '#' + a.map((x) => pad(Math.max(0, Math.min(255, x)).toString(16))).join('');
  function ramp(n, a0, b0) {
    const cs = getComputedStyle(document.documentElement);
    const a = hex2rgb(a0 || cs.getPropertyValue('--p0').trim() || '#111111'), b = hex2rgb(b0 || cs.getPropertyValue('--p1').trim() || '#d4d4d4');
    return Array.from({ length: n }, (_, i) => { const t = n > 1 ? i / (n - 1) : 0; return rgb2hex(a.map((x, j) => Math.round(x + (b[j] - x) * t))); });
  }
  function topN(items, n) {
    const s = items.filter((x) => x.value > 0).sort((a, b) => b.value - a.value);
    if (s.length <= n) return s;
    const head = s.slice(0, n - 1); const rest = s.slice(n - 1).reduce((a, x) => a + x.value, 0);
    return head.concat([{ label: 'Otros (' + (s.length - n + 1) + ')', value: rest }]);
  }
  let pieCharts = [];
  function pieCardHtml(id, title) { return '<div class="card"><h3>' + esc(title) + '</h3><div id="' + id + '"></div></div>'; }
  function drawPie(id, items, fmt, ariaLabel) {
    const host = $('#' + id); if (!host) return;
    const data = topN(items, 8); const total = data.reduce((a, x) => a + x.value, 0);
    if (!data.length) { host.innerHTML = '<p class="pie-empty">Sin datos para este filtro.</p>'; return; }
    const colors = ramp(data.length);
    host.innerHTML = '<div class="pie-body"><div class="pie-cv"><canvas role="img" aria-label="' + esc(ariaLabel) + '"></canvas></div><ol class="pie-leg">' +
      data.map((x, i) => '<li><i style="background:' + colors[i] + '"></i><span class="n" title="' + esc(x.label) + '">' + esc(x.label) + '</span><span class="v">' + esc(fmt(x.value)) + '</span><span class="p">' + nf1.format(x.value / total * 100) + '%</span></li>').join('') + '</ol></div>';
    const surface = getComputedStyle(document.documentElement).getPropertyValue('--surface').trim() || '#fff';
    pieCharts.push(new Chart($('canvas', host), {
      type: 'doughnut',
      data: { labels: data.map((x) => x.label), datasets: [{ data: data.map((x) => x.value), backgroundColor: colors, borderColor: surface, borderWidth: 2 }] },
      options: { responsive: true, maintainAspectRatio: false, animation: false, cutout: '0%', plugins: { legend: { display: false }, tooltip: { callbacks: { label: (x) => x.label + ': ' + fmt(x.parsed) + ' (' + nf1.format(x.parsed / total * 100) + '%)' } } } }
    }));
  }

  function renderPies() {
    const host = $('#pies'); if (!host) return;
    pieCharts.forEach((p) => p.destroy()); pieCharts = [];
    const m = state.f.metric; const mes = state.f.mes; const L = METRICS[m].label;
    const groups = selectedGroups().filter((g) => g.mes === mes);
    const agg = (kf) => { const mp = new Map(); groups.forEach((g) => { const k = kf(g); mp.set(k, (mp.get(k) || 0) + g.proj[m]); }); return Array.from(mp.entries()).map((e) => ({ label: e[0], value: e[1] })); };
    const fm = (v) => fmtM(m, v);
    const dm = m === 'monto' ? 'monto' : 'unidades'; const dL = METRICS[dm].label; const fd = (v) => fmtM(dm, v);
    const cards = [
      ['pieEmp', L + ' proyectado por empresa'], ['pieEstado', L + ' proyectado por estado'], ['pieSede', 'Sedes con mayor proyección (' + L + ')']
    ];
    if (hasDet()) { cards.push(['pieCat', dL + ' proyectado por categoría'], ['pieMarca', dL + ' proyectado por marca']); }
    else cards.push(['pieNoDet', 'Categorías y marcas']);
    host.innerHTML = cards.map((c) => pieCardHtml(c[0], c[1])).join('');
    drawPie('pieEmp', agg((g) => g.empresa), fm, 'Por empresa');
    drawPie('pieEstado', agg((g) => estadoOf(g.sede)), fm, 'Por estado');
    drawPie('pieSede', agg((g) => g.sede + ' (' + g.empresa + ')'), fm, 'Por sede');
    if (hasDet()) {
      const base = { d1: mes + '-01', d2: mes + '-' + pad(daysIn(mes)), empresas: state.f.empresas, metric: dm };
      const rc = runReport(Object.assign({ group: 'categoria' }, base)); const rm = runReport(Object.assign({ group: 'marca' }, base));
      drawPie('pieCat', rc.rows.map((r) => ({ label: r.key, value: r.proj })), fd, 'Por categoría');
      drawPie('pieMarca', rm.rows.map((r) => ({ label: r.key, value: r.proj })), fd, 'Por marca');
    } else {
      $('#pieNoDet').innerHTML = '<p class="pie-empty">Carga el detalle de productos (botón de arriba) para ver las categorías y marcas de mayor a menor venta.</p>';
    }
    renderTopProd();
  }

  function renderTopProd() {
    const host = $('#topProd'); if (!host) return;
    if (!hasDet()) { host.innerHTML = ''; return; }
    const mes = state.f.mes; const dm = state.f.metric === 'monto' ? 'monto' : 'unidades';
    const r = runReport({ group: 'producto', d1: mes + '-01', d2: mes + '-' + pad(daysIn(mes)), empresas: state.f.empresas, metric: dm });
    if (r.error) { host.innerHTML = ''; return; }
    host.innerHTML = '<div class="card"><h3>Productos y modelos con mayor proyección, ' + esc(monthLabel(mes)) + ' (' + METRICS[dm].label + ')</h3><div class="tbl-wrap"><table class="tbl"><thead><tr><th>#</th><th>Producto / modelo</th><th class="num">Observado</th><th class="num">Proyectado</th><th class="num">% del total</th></tr></thead><tbody>' +
      r.rows.slice(0, 10).map((x, i) => '<tr><td>' + (i + 1) + '</td><td>' + esc(x.key) + '</td><td class="num">' + fmtM(dm, x.obs) + '</td><td class="num">' + fmtM(dm, x.proj) + '</td><td class="num">' + nf1.format(x.share * 100) + '%</td></tr>').join('') +
      '</tbody></table></div></div>';
  }

  // ---------- reporte ----------
  let repLast = null;
  const REP = { group: [['sede', 'Sede'], ['empresa', 'Empresa'], ['estado', 'Estado'], ['mes', 'Mes'], ['categoria', 'Categoría'], ['marca', 'Marca'], ['producto', 'Producto / modelo']] };
  const repSel = (id) => new Set(Array.from($(id).selectedOptions).map((o) => o.value));
  const fmtDate = (s) => s.slice(8) + '/' + s.slice(5, 7) + '/' + s.slice(0, 4);

  function fillMulti(id, values, labelFn, keep) {
    const el = $(id); const prev = keep ? repSel(id) : new Set();
    el.innerHTML = values.map((v) => '<option value="' + esc(v) + '"' + (prev.has(v) ? ' selected' : '') + '>' + esc(labelFn ? labelFn(v) : v) + '</option>').join('');
  }

  function renderRepSedes(keep) {
    const c = compute(); const emps = repSelEmp(); const est = repSel('#repEstado');
    const list = c.sedes.filter((s) => emps.has(s.empresa) && (!est.size || est.has(estadoOf(s.sede)))).sort((a, b) => a.sede.localeCompare(b.sede, 'es') || a.empresa.localeCompare(b.empresa));
    const el = $('#repSede'); const prev = keep ? repSel('#repSede') : new Set();
    el.innerHTML = list.map((s) => { const k = keyOf(s.empresa, s.sede); return '<option value="' + esc(k) + '"' + (prev.has(k) ? ' selected' : '') + '>' + esc(s.sede + ' (' + s.empresa + ')') + '</option>'; }).join('');
  }
  const repSelEmp = () => new Set($$('#repEmp input:checked').map((i) => i.value));

  function renderRepProds(keep) {
    if (!hasDet()) { $('#repProd').innerHTML = ''; return; }
    const q = $('#repProdQ').value.trim().toUpperCase(); const prev = keep ? repSel('#repProd') : new Set();
    const cnt = new Map(); state.detail.rows.forEach((r) => { if (r.producto) cnt.set(r.producto, (cnt.get(r.producto) || 0) + r.cant); });
    let list = Array.from(cnt.keys()).filter((p) => !q || p.includes(q)).sort((a, b) => cnt.get(b) - cnt.get(a));
    prev.forEach((p) => { if (!list.includes(p) && cnt.has(p)) list.unshift(p); });
    list = list.slice(0, 300);
    $('#repProd').innerHTML = list.map((p) => '<option value="' + esc(p) + '"' + (prev.has(p) ? ' selected' : '') + '>' + esc(p) + '</option>').join('');
  }

  function renderRepForm(reset) {
    const has = state.rows.length > 0;
    $('#repEmpty').hidden = has; $('#repBody').hidden = !has; if (!has) return;
    const c = compute(); const dates = state.rows.map((r) => r.fecha).sort();
    if (reset || !$('#repDesde').value) { $('#repDesde').value = dates[0]; $('#repHasta').value = dates[dates.length - 1]; }
    $('#repDesde').min = $('#repHasta').min = dates[0]; $('#repDesde').max = $('#repHasta').max = dates[dates.length - 1];
    const prevM = $('#repMedida').value || 'unidades'; const prevG = $('#repGrupo').value || 'sede';
    $('#repMedida').innerHTML = Object.keys(METRICS).map((k) => '<option value="' + k + '"' + (k === prevM ? ' selected' : '') + '>' + METRICS[k].label + '</option>').join('');
    const det = hasDet();
    $('#repGrupo').innerHTML = REP.group.map((g) => '<option value="' + g[0] + '"' + (g[0] === prevG ? ' selected' : '') + (!det && ['categoria', 'marca', 'producto'].includes(g[0]) ? ' disabled' : '') + '>' + g[1] + '</option>').join('');
    const prevE = reset || !$('#repEmp').children.length ? null : repSelEmp();
    $('#repEmp').innerHTML = c.empresas.map((e, i) => '<label class="chip"><input type="checkbox" value="' + esc(e) + '"' + (!prevE || prevE.has(e) ? ' checked' : '') + '><i style="background:' + empVar(e, i) + '"></i>' + esc(e) + '</label>').join('');
    fillMulti('#repEstado', Array.from(new Set(c.sedes.map((s) => estadoOf(s.sede)))).sort((a, b) => a.localeCompare(b, 'es')), null, !reset);
    renderRepSedes(!reset);
    const cats = det ? Array.from(new Set(state.detail.rows.map((r) => r.categoria).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'es')) : [];
    const marcas = det ? Array.from(new Set(state.detail.rows.map((r) => r.marca).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'es')) : [];
    fillMulti('#repCat', cats, null, !reset); fillMulti('#repMarca', marcas, null, !reset); renderRepProds(!reset);
    ['#repCat', '#repMarca', '#repProd', '#repProdQ'].forEach((s) => { $(s).disabled = !det; });
    $('#repDetHint').textContent = det ? 'Detalle cargado: ' + nf0.format(state.detail.rows.length) + ' filas de "' + state.detail.fileName + '". Mantén Ctrl (o Cmd) para elegir varias opciones.'
      : 'Categorías, marcas y productos se activan al cargar el detalle de productos (botón de arriba). Mientras tanto puedes proyectar por fechas, empresas, estados y sedes.';
  }

  function repOptions() {
    const emps = repSelEmp(); const all = compute().empresas;
    return {
      d1: $('#repDesde').value, d2: $('#repHasta').value, metric: $('#repMedida').value, group: $('#repGrupo').value,
      empresas: emps.size === all.length ? null : emps, estados: repSel('#repEstado'), sedes: repSel('#repSede'),
      cats: repSel('#repCat'), marcas: repSel('#repMarca'), prods: repSel('#repProd'), top: Number($('#repTop').value) || 0
    };
  }

  function repFilterLines(o) {
    const lst = (s, all) => (s && s.size ? Array.from(s).map((x) => x.replace('||', ' - ')).join(', ') : all);
    const L = [['Período', fmtDate(o.d1) + ' al ' + fmtDate(o.d2)], ['Medida', METRICS[o.metric].label], ['Agrupado por', (REP.group.find((g) => g[0] === o.group) || [])[1]],
      ['Empresas', lst(o.empresas, 'Todas')], ['Estados', lst(o.estados, 'Todos')], ['Sedes', lst(o.sedes, 'Todas')]];
    if (hasDet()) L.push(['Categorías', lst(o.cats, 'Todas')], ['Marcas', lst(o.marcas, 'Todas')], ['Productos / modelos', lst(o.prods, 'Todos')]);
    return L;
  }

  function generateReport() {
    const o = repOptions(); const out = $('#repOut');
    if (!o.d1 || !o.d2 || o.d1 > o.d2) { out.innerHTML = '<div class="msg err">Revisa las fechas: "Desde" debe ser igual o anterior a "Hasta".</div>'; repLast = null; setRepBtns(false); return; }
    const r = runReport(o);
    if (!r.error) r.shown = o.top ? r.rows.slice(0, o.top) : r.rows;
    if (r.error) { out.innerHTML = '<div class="msg err">' + esc(r.error) + '</div>'; repLast = null; setRepBtns(false); return; }
    if (!r.rows.length) { out.innerHTML = '<div class="msg err">No hay datos con esos filtros.</div>'; repLast = null; setRepBtns(false); return; }
    repLast = { o: o, r: r, filters: repFilterLines(o), when: new Date() };
    setRepBtns(true); renderRepOut();
  }
  const setRepBtns = (on) => { $('#btnRepXlsx').disabled = !on; $('#btnRepPdf').disabled = !on; };

  function repCols(r) {
    const dims = DIMS[r.group].map((d, i) => ({ label: d[0], get: (x) => x.parts[i], num: false }));
    return dims.concat([
      { label: 'Observado', get: (x) => x.obs, num: true, fmt: (v) => fmtM(r.metric, v) },
      { label: 'Proyectado', get: (x) => x.proj, num: true, fmt: (v) => fmtM(r.metric, v) },
      { label: '% del total', get: (x) => x.share, num: true, fmt: (v) => nf1.format(v * 100) + '%' }
    ]);
  }

  function renderRepOut() {
    const { o, r, filters } = repLast; const cols = repCols(r); const out = $('#repOut');
    pieCharts.forEach((p) => { if (p.canvas && p.canvas.id === 'repPieCv') p.destroy(); });
    out.innerHTML = '<div class="kpis"><div class="kpi"><div class="l">Proyectado</div><div class="v">' + fmtM(o.metric, r.totalProj) + '</div><div class="s">' + METRICS[o.metric].label + '</div></div>' +
      '<div class="kpi"><div class="l">Observado en días estudiados</div><div class="v">' + fmtM(o.metric, r.totalObs) + '</div><div class="s">dentro del período</div></div>' +
      '<div class="kpi"><div class="l">Filas</div><div class="v">' + nf0.format(r.rows.length) + '</div><div class="s">agrupado por ' + esc((REP.group.find((g) => g[0] === o.group) || [])[1].toLowerCase()) + '</div></div></div>' +
      '<div class="card"><h3>Filtros aplicados</h3><ul class="rep-filters">' + filters.map((f) => '<li><b>' + esc(f[0]) + ':</b> ' + esc(f[1]) + '</li>').join('') + '</ul>' +
      r.notes.map((n) => '<p class="hint">' + esc(n) + '</p>').join('') + (r.shown.length < r.rows.length ? '<p class="hint">Se muestran las ' + nf0.format(r.shown.length) + ' primeras de ' + nf0.format(r.rows.length) + ' filas. Los totales incluyen todas.</p>' : '') + '</div>' +
      '<div class="card"><h3>Distribución de lo proyectado</h3><div id="repPie"></div></div>' +
      '<div class="card"><h3>Resultado</h3><div class="tbl-wrap"><table class="tbl"><thead><tr>' + cols.map((c) => '<th class="' + (c.num ? 'num' : '') + '">' + esc(c.label) + '</th>').join('') + '</tr></thead><tbody>' +
      r.shown.map((x) => '<tr>' + cols.map((c) => '<td class="' + (c.num ? 'num' : '') + '">' + esc(c.num ? c.fmt(c.get(x)) : c.get(x)) + '</td>').join('') + '</tr>').join('') +
      '</tbody><tfoot><tr><td colspan="' + DIMS[r.group].length + '">Total</td><td class="num">' + fmtM(o.metric, r.totalObs) + '</td><td class="num">' + fmtM(o.metric, r.totalProj) + '</td><td class="num">100%</td></tr></tfoot></table></div></div>';
    const lbl = (x) => (r.group === 'sede' ? x.parts[1] + ' (' + x.parts[0] + ')' : x.parts.join(' '));
    drawPie('repPie', r.shown.map((x) => ({ label: lbl(x), value: x.proj })), (v) => fmtM(o.metric, v), 'Distribución del reporte');
  }

  function repFileBase() { const { o } = repLast; return 'reporte_proyeccion_' + o.d1 + '_a_' + o.d2 + '_' + o.group; }

  function exportXlsx() {
    if (!repLast) return; const { o, r, filters, when } = repLast; const cols = repCols(r);
    const aoa = [['Reporte de proyección de ventas - IDM'], ['Generado', when.toLocaleString('es-VE')], []].concat(filters.map((f) => [f[0], f[1]]), [[]]);
    const hdrRow = aoa.length;
    aoa.push(cols.map((c) => c.label));
    r.shown.forEach((x) => aoa.push(cols.map((c) => (c.num ? Math.round(c.get(x) * 100) / 100 : c.get(x)))));
    const totRow = aoa.length;
    const tot = cols.map(() => ''); tot[0] = 'Total'; tot[DIMS[r.group].length] = Math.round(r.totalObs * 100) / 100; tot[DIMS[r.group].length + 1] = Math.round(r.totalProj * 100) / 100; tot[DIMS[r.group].length + 2] = 1;
    aoa.push(tot);
    r.notes.forEach((n) => aoa.push(['Nota', n]));
    if (r.shown.length < r.rows.length) aoa.push(['Nota', 'Se muestran las ' + r.shown.length + ' primeras de ' + r.rows.length + ' filas; el total incluye todas.']);
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const nd = DIMS[r.group].length;
    for (let i = hdrRow + 1; i <= totRow; i++) {
      [[nd, '#,##0'], [nd + 1, '#,##0'], [nd + 2, '0.0%']].forEach((p) => { const a = XLSX.utils.encode_cell({ r: i, c: p[0] }); if (ws[a]) ws[a].z = p[1]; });
    }
    ws['!cols'] = cols.map((c, i) => ({ wch: i < nd ? (r.group === 'sede' && i === 1 ? 44 : 30) : 16 }));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
    XLSX.writeFile(wb, repFileBase() + '.xlsx');
  }

  function exportPdf() {
    if (!repLast) return; const { o, r, filters, when } = repLast; const cols = repCols(r);
    const J = window.jspdf && window.jspdf.jsPDF; if (!J) return showMsg('No se cargó la librería de PDF.', true);
    const doc = new J({ unit: 'pt', format: 'a4' }); const W = doc.internal.pageSize.getWidth(); const M = 40;
    try { if (window.LOGO_PNG) doc.addImage(window.LOGO_PNG, 'PNG', M, 30, 86, 26); } catch (e) { /* sin logo */ }
    doc.setTextColor(17); doc.setFont('helvetica', 'bold'); doc.setFontSize(15); doc.text('Reporte de proyección de ventas', W - M, 42, { align: 'right' });
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(110); doc.text('Generado el ' + when.toLocaleString('es-VE'), W - M, 56, { align: 'right' });
    doc.setDrawColor(200); doc.line(M, 68, W - M, 68);
    let y = 86; doc.setFontSize(9);
    filters.forEach((f) => {
      doc.setFont('helvetica', 'bold'); doc.setTextColor(17); doc.text(f[0] + ':', M, y);
      doc.setFont('helvetica', 'normal'); doc.setTextColor(70);
      const lines = doc.splitTextToSize(String(f[1]), W - 2 * M - 110); doc.text(lines, M + 110, y); y += 12 * lines.length;
    });
    y += 6;
    const kp = [['Proyectado', fmtM(o.metric, r.totalProj)], ['Observado', fmtM(o.metric, r.totalObs)], ['Filas', nf0.format(r.rows.length)]];
    const bw = (W - 2 * M - 20) / 3;
    kp.forEach((k, i) => {
      const x = M + i * (bw + 10); doc.setDrawColor(210); doc.setFillColor(246, 246, 246); doc.rect(x, y, bw, 42, 'FD');
      doc.setFontSize(8); doc.setTextColor(110); doc.text(k[0].toUpperCase(), x + 10, y + 15);
      doc.setFontSize(15); doc.setFont('helvetica', 'bold'); doc.setTextColor(17); doc.text(k[1], x + 10, y + 34); doc.setFont('helvetica', 'normal');
    });
    y += 58;
    // pastel
    try {
      const data = topN(r.shown.map((x) => ({ label: r.group === 'sede' ? x.parts[1] + ' (' + x.parts[0] + ')' : x.parts.join(' '), value: x.proj })), 8);
      if (data.length && typeof Chart !== 'undefined') {
        const cv = document.createElement('canvas'); cv.width = 1000; cv.height = 380;
        const colors = ramp(data.length, '#111111', '#d4d4d4'); const tot = data.reduce((a, x) => a + x.value, 0);
        const ch = new Chart(cv, {
          type: 'doughnut',
          data: { labels: data.map((x) => x.label + '  -  ' + nf1.format(x.value / tot * 100) + '%'), datasets: [{ data: data.map((x) => x.value), backgroundColor: colors, borderColor: '#fff', borderWidth: 2 }] },
          options: { responsive: false, animation: false, cutout: '0%', layout: { padding: 10 }, plugins: { legend: { position: 'right', labels: { color: '#111', font: { size: 20 }, boxWidth: 22 } } } },
          plugins: [{ id: 'bg', beforeDraw: (c) => { const x = c.ctx; x.save(); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.restore(); } }]
        });
        const img = cv.toDataURL('image/png'); ch.destroy();
        const iw = W - 2 * M; const ih = iw * 380 / 1000;
        if (y + ih > 780) { doc.addPage(); y = 50; }
        doc.addImage(img, 'PNG', M, y, iw, ih); y += ih + 12;
      }
    } catch (e) { /* sin gráfico */ }
    const nd = DIMS[r.group].length;
    const body = r.shown.map((x) => cols.map((c) => (c.num ? c.fmt(c.get(x)) : c.get(x))));
    const foot = cols.map(() => ''); foot[0] = 'Total'; foot[nd] = fmtM(o.metric, r.totalObs); foot[nd + 1] = fmtM(o.metric, r.totalProj); foot[nd + 2] = '100%';
    const colStyles = {}; cols.forEach((c, i) => { if (c.num) colStyles[i] = { halign: 'right' }; });
    doc.autoTable({
      startY: y, head: [cols.map((c) => c.label)], body: body, foot: [foot], margin: { left: M, right: M, bottom: 40 },
      styles: { font: 'helvetica', fontSize: 8.5, cellPadding: 4, textColor: [30, 30, 30], lineColor: [225, 225, 225], lineWidth: 0.4 },
      headStyles: { fillColor: [17, 17, 17], textColor: 255, fontStyle: 'bold' }, footStyles: { fillColor: [235, 235, 235], textColor: 17, fontStyle: 'bold' },
      alternateRowStyles: { fillColor: [248, 248, 248] }, columnStyles: colStyles, showFoot: 'lastPage'
    });
    let fy = doc.lastAutoTable.finalY + 14; doc.setFontSize(8); doc.setTextColor(110);
    const notasPdf = r.notes.concat(r.shown.length < r.rows.length ? ['Se muestran las ' + r.shown.length + ' primeras de ' + r.rows.length + ' filas; el total incluye todas.'] : []);
    notasPdf.forEach((n) => { const ls = doc.splitTextToSize('Nota: ' + n, W - 2 * M); if (fy + 11 * ls.length > 800) { doc.addPage(); fy = 50; } doc.text(ls, M, fy); fy += 11 * ls.length; });
    const pages = doc.getNumberOfPages();
    for (let p = 1; p <= pages; p++) { doc.setPage(p); doc.setFontSize(8); doc.setTextColor(140); doc.text('IDM - Proyecciones de ventas   |   Página ' + p + ' de ' + pages, W / 2, 822, { align: 'center' }); }
    doc.save(repFileBase() + '.pdf');
  }


  // ---------- conexión con Google Sheets ----------
  const LS_CONN = 'proy.conn.v1';
  const conn = { url: '', clave: '', on: false, when: null, busy: false, meta: null, pctSaved: null, error: '' };
  const connUrl = () => String(window.PUENTE_URL || '').trim() || conn.url;
  const puenteOn = () => !!(conn.on && connUrl() && conn.clave);

  async function puentePost(payload) {
    let r;
    try {
      r = await fetch(connUrl(), { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(Object.assign({ clave: conn.clave }, payload)) });
    } catch (e) { throw new Error('No pude comunicarme con el puente. Revisa tu internet y la dirección.'); }
    if (!r.ok) throw new Error('El puente respondió con el error HTTP ' + r.status + '.');
    let j; try { j = await r.json(); } catch (e) { throw new Error('La dirección no responde como un puente. Debe terminar en /exec y estar publicada con acceso para "Cualquier usuario".'); }
    if (j.error === 'clave') { const e = new Error('La clave de acceso es incorrecta.'); e.clave = true; throw e; }
    if (!j.ok) throw new Error('El puente respondió con un error (' + (j.error || 'desconocido') + (j.detalle ? ': ' + j.detalle : '') + ').');
    return j;
  }

  function renderConn() {
    const fixed = String(window.PUENTE_URL || '').trim();
    $('#fUrl').hidden = !!fixed;
    if (!fixed && document.activeElement !== $('#connUrl')) $('#connUrl').value = conn.url;
    $('#btnDisconnect').hidden = !conn.on;
    $('#btnSync').hidden = !puenteOn();
    $('#btnSync').disabled = conn.busy; $('#btnConnect').disabled = conn.busy;
    $('#btnSync').textContent = conn.busy ? 'Actualizando…' : 'Actualizar datos';
    const fmt = (iso) => { const d = new Date(iso); return isNaN(d) ? '-' : d.toLocaleString('es-VE', { dateStyle: 'short', timeStyle: 'short' }); };
    const t = [];
    if (conn.busy) t.push('Actualizando desde la hoja de Google…');
    else if (puenteOn()) {
      t.push('Conectado.' + (conn.when ? ' Última actualización en este navegador: ' + conn.when.toLocaleString('es-VE', { dateStyle: 'short', timeStyle: 'short' }) + '.' : ''));
      if (conn.meta) {
        if (conn.meta.ultima_carga_detalle) t.push('Último envío del scraping (detalle): ' + fmt(conn.meta.ultima_carga_detalle) + '.');
        if (conn.meta.ultima_carga_consolidado) t.push('Último envío automático del consolidado: ' + fmt(conn.meta.ultima_carga_consolidado) + '.');
        if (conn.meta.ultimo_cambio_pct) t.push('Último cambio de %: ' + fmt(conn.meta.ultimo_cambio_pct) + '.');
      }
    } else t.push('Sin conexión. Escribe la clave de acceso y pulsa "Conectar y actualizar".');
    if (conn.error) t.push('Último error: ' + conn.error);
    $('#connStatus').textContent = t.join(' ');
  }

  async function syncRemote() {
    if (conn.busy || !puenteOn()) return;
    conn.busy = true; conn.error = ''; renderConn();
    try {
      const j = await puentePost({ accion: 'leer' });
      const parsed = j.consolidado && j.consolidado.length > 1 ? parseConsolidado(j.consolidado) : null;
      if (!parsed) throw new Error('La pestaña CONSOLIDADO de la hoja está vacía o no tiene los encabezados esperados (id, Empresa, Sucursal, Factura, Visitas, Proyeccion, Total, Fecha).');
      if (parsed.missing.length) throw new Error('A la pestaña CONSOLIDADO le faltan columnas: ' + parsed.missing.join(', ') + '.');
      const ds = buildDataset(parsed.raw);
      if (!ds.rows.length) throw new Error('La pestaña CONSOLIDADO no tiene filas con fecha válida.');
      const first = !state.rows.length;
      state.rows = ds.rows; state.report = ds.report; state.fileName = 'Google Sheets'; state.loadedAt = new Date().toISOString(); cache = null;
      const rp = j.pct || {};
      if (Object.keys(rp).length) { state.pct = rp; if (typeof j.defaultPct === 'number') state.defaultPct = j.defaultPct; lsSet(LS_PCT, { pct: state.pct, defaultPct: state.defaultPct }); }
      else if (Object.keys(state.pct).length) { schedulePctPush(); }
      state.detail = null;
      let detMsg = 'sin detalle de productos';
      if (j.detalle && j.detalle.length > 1) {
        const pd = parseDetalle(j.detalle);
        if (pd && pd.rows.length) { state.detail = makeDetail(pd, 'Google Sheets'); detMsg = nf0.format(pd.rows.length) + ' filas de detalle'; }
        else detMsg = 'el detalle de la hoja no se pudo leer (revisa las columnas)';
      }
      conn.meta = j.meta || null; conn.when = new Date();
      initFilters(first); renderAll();
      showMsg('Datos actualizados desde Google Sheets: ' + nf0.format(state.rows.length) + ' registros del consolidado y ' + detMsg + '.', false);
    } catch (e) {
      conn.error = e.message; showMsg(e.message, true);
      if (e.clave) { conn.on = false; lsDel(LS_CONN); setTab('datos'); }
    } finally { conn.busy = false; renderConn(); renderStatus(); }
  }

  let pctTimer = null;
  function schedulePctPush() { clearTimeout(pctTimer); pctTimer = setTimeout(pushPct, 900); }
  async function pushPct() {
    if (!puenteOn()) return;
    try { await puentePost({ accion: 'guardar_pct', pct: state.pct, defaultPct: state.defaultPct }); conn.pctSaved = new Date(); conn.error = ''; }
    catch (e) { showMsg('No pude guardar los % en la hoja: ' + e.message, true); }
  }

  async function connectNow() {
    const url = String(window.PUENTE_URL || '').trim() || $('#connUrl').value.trim();
    const key = $('#connKey').value.trim();
    if (!url || !/^https:\/\/|^http:\/\/localhost/.test(url)) return showMsg('Escribe la dirección del puente (empieza con https:// y termina en /exec).', true);
    if (!key) return showMsg('Escribe la clave de acceso.', true);
    conn.url = url; conn.clave = key; conn.on = true;
    lsSet(LS_CONN, { url: conn.url, clave: conn.clave });
    $('#connKey').value = '';
    await syncRemote();
    if (conn.on && state.rows.length) setTab('resumen');
  }

  function disconnect() {
    conn.on = false; conn.clave = ''; conn.meta = null; lsDel(LS_CONN); showMsg('Desconectado. Los datos que ya estaban cargados se mantienen en este navegador.', false); renderConn(); renderStatus();
  }

  // ---------- render general ----------
  function renderStatus() {
    const el = $('#status');
    if (!state.rows.length) { el.textContent = 'Sin datos cargados'; return; }
    const c = compute(); const d = state.rows.map((r) => r.fecha).sort();
    el.textContent = nf0.format(state.rows.length) + ' registros · ' + d[0] + ' a ' + d[d.length - 1] + ' · ' + c.empresas.length + ' empresas · ' + c.sedes.length + ' sedes' + (state.fileName ? ' · ' + state.fileName : '') +
      (puenteOn() && conn.when ? ' · actualizado ' + conn.when.toLocaleTimeString('es-VE', { hour: '2-digit', minute: '2-digit' }) : '');
  }

  function renderAll() {
    const has = state.rows.length > 0;
    $('#empty').hidden = has; $('#dash').hidden = !has; $('#btnClear').hidden = !has;
    renderStatus();
    if (has) { renderFilters(); renderKpis(); renderCharts(); renderTable(); }
    renderSedes(); renderCalidad(); renderRepForm(true); repLast = null; setRepBtns(false); if ($('#repOut')) $('#repOut').innerHTML = '';
    if (window.__proy && window.__proy.hooks) window.__proy.hooks.forEach((fn) => { try { fn(); } catch (e) { /* informe */ } });
  }

  function updateDefInfo() {
    const sinPct = compute().sedes.filter((s) => !pctInfo(s.empresa, s.sede).known).length;
    $('#defPctInfo').textContent = sinPct ? sinPct + ' sede(s) sin % usan este valor.' : 'Todas las sedes tienen %.';
  }

  function refreshCalc(keepSedesTable) {
    invalidate();
    if (state.rows.length) { renderKpis(); renderCharts(); renderTable(); }
    if (keepSedesTable) updateDefInfo(); else renderSedes();
  }

  // ---------- eventos ----------
  function setTab(name) {
    $$('.tab').forEach((t) => t.setAttribute('aria-selected', String(t.dataset.tab === name)));
    $$('.panel').forEach((p) => { p.hidden = p.id !== 'tab-' + name; });
    if (name === 'resumen' && state.rows.length) renderCharts();
    if (name === 'reporte') renderRepForm(false);
    if (name === 'datos') renderConn();
  }

  function bind() {
    $$('.tab').forEach((t) => t.addEventListener('click', () => setTab(t.dataset.tab)));
    $('#fileMain').addEventListener('change', (e) => { if (e.target.files[0]) loadConsolidado(e.target.files[0]); e.target.value = ''; });
    const drop = $('#empty');
    ['dragover', 'dragenter'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', (e) => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f) loadConsolidado(f); });
    $('#btnClear').addEventListener('click', () => {
      if (!window.confirm('Se borra el consolidado y el detalle guardados en este navegador. Los % de las sedes se conservan. ¿Continuar?')) return;
      state.rows = []; state.report = null; state.detail = null; lsDel(LS_DATA); lsDel(LS_DET); invalidate(); $('#msg').hidden = true; renderAll();
    });
    $('#selMes').addEventListener('change', (e) => { state.f.mes = e.target.value; renderKpis(); renderCharts(); renderTable(); });
    $('#segMetric').addEventListener('change', (e) => { if (e.target.name === 'metric') { state.f.metric = e.target.value; renderKpis(); renderCharts(); renderTable(); } });
    $('#chipsEmp').addEventListener('change', (e) => {
      const v = e.target.value; if (e.target.checked) state.f.empresas.add(v); else state.f.empresas.delete(v);
      if (!state.f.empresas.size) { state.f.empresas.add(v); e.target.checked = true; }
      renderKpis(); renderCharts(); renderTable();
    });
    $('#search').addEventListener('input', (e) => { state.search = e.target.value; renderTable(); });
    $('#tbl').addEventListener('click', (e) => {
      const b = e.target.closest('button[data-sort]'); if (!b) return;
      const k = b.dataset.sort; state.sort = { key: k, dir: state.sort.key === k ? -state.sort.dir : (k === 'empresa' || k === 'sede' ? 1 : -1) }; renderTable();
    });
    $('#btnCsv').addEventListener('click', exportCsv);
    $('#btnConnect').addEventListener('click', connectNow);
    $('#connKey').addEventListener('keydown', (e) => { if (e.key === 'Enter') connectNow(); });
    $('#btnDisconnect').addEventListener('click', disconnect);
    $('#btnSync').addEventListener('click', syncRemote);
    $('#btnGoDatos').addEventListener('click', () => setTab('datos'));
    $('#repGrupo').addEventListener('change', (e) => { $('#repTop').value = e.target.value === 'producto' ? '50' : '0'; });
    document.addEventListener('visibilitychange', () => { if (!document.hidden && puenteOn() && conn.when && Date.now() - conn.when.getTime() > 10 * 60 * 1000) syncRemote(); });
    $('#fileDet').addEventListener('change', (e) => { if (e.target.files[0]) loadDetalle(e.target.files[0]); e.target.value = ''; });
    $('#btnRepGen').addEventListener('click', generateReport);
    $('#btnRepXlsx').addEventListener('click', exportXlsx);
    $('#btnRepPdf').addEventListener('click', exportPdf);
    $('#btnRepClear').addEventListener('click', () => { renderRepForm(true); repLast = null; setRepBtns(false); $('#repOut').innerHTML = ''; $('#repProdQ').value = ''; renderRepProds(false); });
    $('#repEmp').addEventListener('change', () => { if (!repSelEmp().size) { const f = $('#repEmp input'); if (f) f.checked = true; } renderRepSedes(true); });
    $('#repEstado').addEventListener('change', () => renderRepSedes(true));
    $('#repProdQ').addEventListener('input', () => renderRepProds(true));
    $('#searchSedes').addEventListener('input', (e) => { state.searchSedes = e.target.value; renderSedes(); });
    $('#tblSedes').addEventListener('change', (e) => {
      const inp = e.target.closest('input[data-k]'); if (!inp) return;
      const k = inp.dataset.k; const v = inp.value.trim();
      const parts = k.split('||');
      if (v === '') delete state.pct[k];
      else {
        const n = Number(v);
        if (!isFinite(n) || n < 0) { const cur = pctInfo(parts[0], parts[1]); inp.value = cur.known ? Math.round(cur.value * 1000) / 10 : ''; return; }
        state.pct[k] = Math.round(n * 10) / 1000;
      }
      savePct();
      // se actualiza solo la fila editada: redibujar la tabla mientras el campo tiene el foco provoca errores en el navegador
      const p = pctInfo(parts[0], parts[1]);
      inp.value = p.known ? Math.round(p.value * 1000) / 10 : '';
      inp.closest('tr').children[3].innerHTML = p.origin === 'Sin %' ? '<span class="tag bad">Sin %</span>' : p.origin === 'Editado' ? '<span class="tag ok">Editado</span>' : '<span class="tag">Excel anterior</span>';
      refreshCalc(true);
    });
    $('#defPct').addEventListener('change', (e) => {
      const n = Number(e.target.value); state.defaultPct = isFinite(n) && n >= 0 ? n / 100 : 0; savePct(); refreshCalc();
    });
    $('#btnPctReset').addEventListener('click', () => {
      if (!window.confirm('Se descartan tus cambios de % y se vuelve a los valores del Excel anterior. ¿Continuar?')) return;
      state.pct = {}; state.defaultPct = 0; savePct(); refreshCalc();
    });
    $('#btnPctExport').addEventListener('click', () => {
      const c = compute(); const out = {};
      c.sedes.forEach((s) => { const p = pctInfo(s.empresa, s.sede); if (p.known) out[keyOf(s.empresa, s.sede)] = p.value; });
      Object.keys(state.pct).forEach((k) => { out[k] = state.pct[k]; });
      downloadBlob('porcentajes_sedes.json', JSON.stringify({ version: 1, defaultPct: state.defaultPct, pct: out }, null, 2), 'application/json');
    });
    $('#filePct').addEventListener('change', async (e) => {
      const f = e.target.files[0]; e.target.value = ''; if (!f) return;
      try {
        const j = JSON.parse(await f.text());
        if (!j || typeof j.pct !== 'object') throw new Error('el archivo no tiene el campo "pct"');
        let n = 0; Object.keys(j.pct).forEach((k) => { const v = Number(j.pct[k]); if (isFinite(v) && v >= 0) { state.pct[k] = v; n++; } });
        if (typeof j.defaultPct === 'number' && j.defaultPct >= 0) state.defaultPct = j.defaultPct;
        savePct(); refreshCalc(); showMsg('Importé ' + n + ' porcentajes.', false);
      } catch (err) { showMsg('No pude importar los porcentajes: ' + err.message, true); }
    });
    $('#fileConv').addEventListener('change', (e) => { if (e.target.files[0]) convLoad(e.target.files[0]); e.target.value = ''; });
    ['#convSheet', '#optHeader', '#optDates', '#optTrim'].forEach((s) => $(s).addEventListener('change', convRebuild));
    $('#btnConvDl').addEventListener('click', () => downloadBlob((conv.name.replace(/\.[^.]+$/, '') || 'datos') + '.json', JSON.stringify(conv.records, null, 2), 'application/json'));
    $('#btnConvCopy').addEventListener('click', async () => {
      try { await navigator.clipboard.writeText(JSON.stringify(conv.records, null, 2)); showMsg('JSON copiado.', false); }
      catch (e) { showMsg('No pude copiar. Usa "Descargar .json".', true); }
    });
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { if (state.rows.length) renderCharts(); });
  }

  // ---------- inicio ----------
  function init() {
    const sp = lsGet(LS_PCT);
    if (sp && typeof sp.pct === 'object' && sp.pct) { state.pct = sp.pct; state.defaultPct = Number(sp.defaultPct) >= 0 ? Number(sp.defaultPct) : 0; }
    const sd = lsGet(LS_DATA);
    if (sd && Array.isArray(sd.rows) && sd.rows.length) {
      state.rows = sd.rows; state.report = sd.report || null; state.fileName = sd.fileName || ''; state.loadedAt = sd.loadedAt || null;
      initFilters(true);
    }
    const dd = lsGet(LS_DET);
    if (state.rows.length && dd && Array.isArray(dd.rows) && dd.rows.length) state.detail = dd;
    const faltan = [['XLSX', 'xlsx.full.min.js'], ['Chart', 'chart.umd.js'], ['jspdf', 'jspdf.umd.min.js']].filter((l) => typeof window[l[0]] === 'undefined').map((l) => l[1]);
    const sc = lsGet(LS_CONN);
    if (sc && sc.clave) { conn.url = sc.url || ''; conn.clave = sc.clave; conn.on = !!connUrl(); }
    bind(); renderAll(); renderConn();
    if (puenteOn()) syncRemote();
    if (faltan.length) showMsg('Faltan archivos en el sitio: ' + faltan.join(', ') + '. Súbelos al repositorio junto a index.html y vuelve a desplegar.', true);
    window.__proy = { state: state, compute: compute, runReport: runReport, conn: conn, pctInfo: pctInfo, daysIn: daysIn, showMsg: showMsg, hooks: [], setTab: setTab };
  }
  init();
})();
