/* Informe de productos: motor por variante + PDF/Excel con el formato del informe mensual. */
(function () {
  'use strict';
  const V = window.Variantes; const P = () => window.__proy;
  const $ = (s, el) => (el || document).querySelector(s);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const mLabel = (m) => { const p = m.split('-'); return MESES[+p[1] - 1] + ' ' + p[0]; };
  const mName = (m) => MESES[+m.split('-')[1] - 1];
  const prevMonth = (m) => { let [y, mo] = m.split('-').map(Number); mo--; if (!mo) { mo = 12; y--; } return y + '-' + String(mo).padStart(2, '0'); };
  const f0 = (x) => Math.round(x).toLocaleString('de-DE');
  const usd = (x) => '$' + f0(x);
  const pc = (x) => Math.round(x * 100) + '%';
  const EMP_FIRST = ['DAKA', 'DAMASCO', 'MULTIMAX'];
  const EMP_COL = { DAKA: [255, 212, 0], DAMASCO: [244, 164, 96], MULTIMAX: [120, 144, 156] };
  const empCol = (e) => EMP_COL[e] || [110, 110, 110];
  const empTxt = (e) => (EMP_COL[e] ? [17, 17, 17] : [255, 255, 255]);
  const PAL = ['#2f5d9b', '#e8801a', '#9aa0a6', '#f2c200', '#4a8f5b', '#7a5cc4', '#c0392b', '#1aa3b8', '#b5835a', '#5b6770', '#d36ba2', '#8aa832'];
  const GCOL = ['#2f5d9b', '#e8801a', '#f2c200', '#4a8f5b', '#7a5cc4', '#c0392b'];

  let model = null; const keep = { prices: {}, narr: {} };

  function monthsInDetail() {
    const d = P().state.detail; if (!d || !d.rows || !d.rows.length) return [];
    return Array.from(new Set(d.rows.map((r) => r.fecha.slice(0, 7)))).sort();
  }

  // ---- motor: proyección por variante, sede y mes (misma fórmula del Excel) ----
  function raw(mes, imput, empSet) {
    const p = P(); const rows = p.state.detail.rows.filter((r) => r.fecha.slice(0, 7) === mes);
    const days = p.daysIn(mes); const dates = new Map();
    rows.forEach((r) => { const k = r.empresa + '||' + r.sede; let s = dates.get(k); if (!s) dates.set(k, s = new Set()); s.add(r.fecha); });
    const fac = new Map();
    dates.forEach((s, k) => { const a = k.split('||'); fac.set(k, { n: s.size, f: (1 + p.pctInfo(a[0], a[1]).value) * (days + 1) / (s.size + 1) }); });
    const out = new Map(), unc = new Map(), sedes = new Map(); let sinCat = 0;
    rows.forEach((r) => {
      if (empSet && !empSet.has(r.empresa)) return;
      const sk = r.empresa + '||' + r.sede; let ss = sedes.get(r.empresa); if (!ss) sedes.set(r.empresa, ss = new Set()); ss.add(sk);
      const c = V.classify(r.categoria, r.producto, r.empresa);
      if (!c.cat) { sinCat++; return; }
      if (c.fail) { const k = r.empresa + '||' + r.categoria + '||' + r.producto; unc.set(k, (unc.get(k) || 0) + r.cant); return; }
      let ec = out.get(r.empresa); if (!ec) out.set(r.empresa, ec = new Map());
      let cv = ec.get(c.cat); if (!cv) ec.set(c.cat, cv = new Map());
      const vk = c.group + '|' + c.label; let v = cv.get(vk);
      if (!v) cv.set(vk, v = { key: vk, group: c.group, label: c.label, ord: c.ord, obs: 0, pure: 0, imp: 0, monto: 0, sedes: new Map(), brands: new Map() });
      const f = fac.get(sk).f; v.obs += r.cant; v.pure += r.cant * f; v.monto += r.monto || 0;
      v.sedes.set(sk, (v.sedes.get(sk) || 0) + r.cant * f);
      const b = r.marca || '(sin marca)'; v.brands.set(b, (v.brands.get(b) || 0) + r.cant * f);
    });
    if (imput) {
      out.forEach((ec, emp) => {
        const all = Array.from(sedes.get(emp) || []); const N = all.length || 1;
        ec.forEach((cv) => cv.forEach((v) => {
          all.forEach((sk) => { if (!(v.sedes.get(sk) > 0)) v.imp += v.pure / N / days * Math.max(0, days - fac.get(sk).n); });
        }));
      });
    }
    return { out: out, unc: unc, days: days, sinCat: sinCat, nRows: rows.length };
  }

  const gIdx = (def, g) => Math.max(0, def.groups.indexOf(g));
  function build(o) {
    const r = raw(o.mes, o.imput, o.empSet); const prevM = prevMonth(o.mes);
    const hasPrev = monthsInDetail().includes(prevM); const rp = hasPrev ? raw(prevM, o.imput, o.empSet) : null;
    const emps = [];
    r.out.forEach((ec, emp) => {
      const cats = [];
      V.CATS.forEach((def) => {
        const cv = ec.get(def.id); if (!cv) return;
        const pcv = rp && rp.out.get(emp) ? rp.out.get(emp).get(def.id) : null;
        const variants = Array.from(cv.values()).sort((a, b) => gIdx(def, a.group) - gIdx(def, b.group) || a.ord - b.ord);
        variants.forEach((v) => {
          v.qty = v.pure + v.imp;
          const pv = pcv ? pcv.get(v.key) : null; v.prevQty = pcv ? (pv ? pv.pure + pv.imp : 0) : null;
          const pk = emp + '|' + def.id + '|' + v.key;
          const auto = v.obs > 0 && v.monto > 0 ? Math.max(5, Math.round(v.monto / v.obs / 5) * 5) : 0;
          v.price = keep.prices[pk] != null ? keep.prices[pk] : auto; v.pk = pk; v.auto = auto;
        });
        // variantes que existían el mes anterior y ya no
        const gone = []; if (pcv) pcv.forEach((pv, k) => { if (!cv.has(k)) gone.push({ label: pv.label, group: pv.group, prev: pv.pure + pv.imp }); });
        const groups = def.groups.map((g) => ({ name: g, qty: variants.filter((v) => v.group === g).reduce((a, v) => a + v.qty, 0), prev: pcv ? Array.from(pcv.values()).filter((v) => v.group === g).reduce((a, v) => a + v.pure + v.imp, 0) : null })).filter((g) => g.qty > 0 || (g.prev || 0) > 0);
        const brands = new Map(); variants.forEach((v) => v.brands.forEach((q, b) => brands.set(b, (brands.get(b) || 0) + q)));
        const cat = { id: def.id, def: def, variants: variants, groups: groups, gone: gone, brands: brands };
        cat.qty = variants.reduce((a, v) => a + v.qty, 0);
        cat.prevQty = pcv ? Array.from(pcv.values()).reduce((a, v) => a + v.pure + v.imp, 0) : null;
        cats.push(cat);
      });
      if (cats.length) emps.push({ name: emp, cats: cats });
    });
    emps.sort((a, b) => { const ia = EMP_FIRST.indexOf(a.name), ib = EMP_FIRST.indexOf(b.name); return (ia < 0 ? 9 : ia) - (ib < 0 ? 9 : ib) || a.name.localeCompare(b.name); });
    const m = { mes: o.mes, prev: hasPrev ? prevM : null, days: r.days, imput: o.imput, emps: emps, unc: r.unc, sinCat: r.sinCat, nRows: r.nRows };
    recalc(m);
    emps.forEach((e) => e.cats.forEach((c) => { const k = e.name + '|' + c.id; c.narr = keep.narr[k] != null ? keep.narr[k] : narrative(m, e, c); c.nk = k; }));
    return m;
  }
  function recalc(m) {
    m.emps.forEach((e) => {
      e.cats.forEach((c) => { c.variants.forEach((v) => { v.usd = v.qty * v.price; }); c.usd = c.variants.reduce((a, v) => a + v.usd, 0); });
      e.usd = e.cats.reduce((a, c) => a + c.usd, 0);
    });
  }

  // ---- narrativa automática ----
  function narrative(m, e, c) {
    const nm = c.def.titulo.toLowerCase(), mes = mName(m.mes), pm = m.prev ? mName(m.prev) : null; const par = [];
    if (c.prevQty != null && c.prevQty > 0) {
      const d = (c.qty - c.prevQty) / c.prevQty; const w = Math.abs(d) < 0.02 ? 'se mantuvo estable' : (d > 0 ? 'presentó un aumento del (' + pc(d) + ')' : 'presentó una disminución del (' + pc(d) + ')');
      par.push('La categoría ' + nm + ' ' + w + ' en su oferta estimada, pasando de ' + f0(c.prevQty) + ' unidades proyectadas en ' + pm + ' a ' + f0(c.qty) + ' unidades en ' + mes + '.');
    } else par.push('La categoría ' + nm + ' proyecta ' + f0(c.qty) + ' unidades para ' + mes + (pm ? ' (no hay unidades de ' + pm + ' para comparar).' : ' (no hay un mes anterior cargado para comparar).'));
    if (c.groups.length > 1) {
      const g = c.groups.slice().sort((a, b) => b.qty - a.qty)[0];
      par.push('El grupo ' + g.name + ' concentra el ' + pc(g.qty / c.qty) + ' de la categoría con ' + f0(g.qty) + ' unidades proyectadas.');
    }
    const top = c.variants.slice().sort((a, b) => b.qty - a.qty)[0];
    par.push('La variante ' + top.group + ' ' + top.label + ' obtuvo la mayor representación con el ' + pc(top.qty / c.qty) + ' de la categoría.');
    if (c.prevQty != null) {
      const ch = c.variants.filter((v) => v.prevQty > 0 && v.qty > 0).map((v) => ({ v: v, d: (v.qty - v.prevQty) / v.prevQty, abs: v.qty - v.prevQty }))
        .filter((x) => Math.abs(x.abs) >= Math.max(30, c.qty * 0.01)).sort((a, b) => Math.abs(b.abs) - Math.abs(a.abs)).slice(0, 3);
      if (ch.length) par.push(ch.map((x) => 'La variante ' + x.v.group + ' ' + x.v.label + (x.d > 0 ? ' aumentó un (' + pc(x.d) + ')' : ' disminuyó un (' + pc(x.d) + ')') + ' frente a ' + pm + ', pasando de ' + f0(x.v.prevQty) + ' a ' + f0(x.v.qty) + ' unidades.').join(' '));
      const nuevas = c.variants.filter((v) => !(v.prevQty > 0) && v.qty >= Math.max(30, c.qty * 0.01)).slice(0, 3);
      if (nuevas.length) par.push('Variantes sin proyección en ' + pm + ' que aparecen en ' + mes + ': ' + nuevas.map((v) => v.label + ' (' + f0(v.qty) + ')').join(', ') + '.');
      if (c.gone.length) { const gg = c.gone.filter((x) => x.prev >= Math.max(30, c.prevQty * 0.01)).slice(0, 3); if (gg.length) par.push('No se observó proyección en ' + mes + ' para: ' + gg.map((x) => x.label + ' (' + f0(x.prev) + ' en ' + pm + ')').join(', ') + '.'); }
    }
    const br = Array.from(c.brands.entries()).sort((a, b) => b[1] - a[1]); const tb = br.reduce((a, x) => a + x[1], 0);
    if (br.length && tb) par.push('La marca ' + br[0][0] + ' lidera la categoría con el ' + pc(br[0][1] / tb) + ' de las unidades' + (br[1] ? ', seguida de ' + br[1][0] + ' con el ' + pc(br[1][1] / tb) + '.' : '.'));
    return par.join('\n---\n');
  }

  // ---- gráficos como imagen (Chart.js fuera de pantalla) ----
  function chartImg(cfg, w, h) {
    const S = 2, cv = document.createElement('canvas'); cv.width = w * S; cv.height = h * S;
    cfg.options = Object.assign({ responsive: false, animation: false, devicePixelRatio: 1 }, cfg.options || {});
    cfg.plugins = (cfg.plugins || []).concat([{ id: 'bg', beforeDraw: (ch) => { const x = ch.ctx; x.save(); x.fillStyle = '#fff'; x.fillRect(0, 0, ch.width, ch.height); x.restore(); } }]);
    const ch = new Chart(cv, cfg); const url = cv.toDataURL('image/jpeg', 0.88); ch.destroy(); return url;
  }
  const FS = 22;
  function pieImg(items, w, h, o) {
    o = o || {}; items = items.filter((x) => x.value > 0); const tot = items.reduce((a, x) => a + x.value, 0) || 1;
    let list = items.slice().sort((a, b) => b.value - a.value); const max = o.max || 7;
    if (list.length > max) { const rest = list.slice(max - 1); list = list.slice(0, max - 1); list.push({ label: 'OTROS', value: rest.reduce((a, x) => a + x.value, 0) }); }
    return chartImg({
      type: 'pie',
      data: { labels: list.map((x) => x.label + ' ' + pc(x.value / tot)), datasets: [{ data: list.map((x) => x.value), backgroundColor: list.map((x, i) => (o.colors && o.colors[x.label]) || PAL[i % PAL.length]), borderColor: '#fff', borderWidth: list.length > 1 ? 3 : 0 }] },
      options: { layout: { padding: 6 }, plugins: { legend: { position: o.legend || 'right', labels: { boxWidth: 22, font: { size: o.fs || FS }, padding: 10 } } } }
    }, w, h);
  }
  function barImg(c, w, h) {
    const vs = c.variants; const gi = (g) => c.def.groups.indexOf(g);
    const lab = { id: 'vals', afterDatasetsDraw: (ch) => { const x = ch.ctx; const m = ch.getDatasetMeta(0); x.save(); x.fillStyle = '#222'; x.font = '26px Helvetica,Arial,sans-serif'; x.textBaseline = 'middle'; m.data.forEach((b, i) => { x.fillText(f0(vs[i].qty), b.x + 6, b.y); }); x.restore(); } };
    return chartImg({
      type: 'bar', data: { labels: vs.map((v) => (c.groups.length > 1 ? v.group.slice(0, 4) + ' ' : '') + v.label), datasets: [{ data: vs.map((v) => v.qty), backgroundColor: vs.map((v) => GCOL[gi(v.group) % GCOL.length]), borderWidth: 0 }] },
      plugins: [lab], options: { indexAxis: 'y', layout: { padding: { right: 110 } }, plugins: { legend: { display: false } }, scales: { x: { ticks: { font: { size: 26 } }, grid: { color: '#e5e5e5' } }, y: { ticks: { font: { size: 26 } }, grid: { display: false } } } }
    }, w, h);
  }
  const brandColors = (names) => { const o = {}; names.forEach((n, i) => { o[n] = PAL[i % PAL.length]; }); return o; };

  // ---- PDF ----
  function exportPdf() {
    if (!model) return; const J = window.jspdf && window.jspdf.jsPDF; if (!J || typeof Chart === 'undefined') return P().showMsg('No se cargó la librería de PDF o de gráficos.', true);
    const doc = new J({ unit: 'pt', format: 'letter', orientation: 'landscape', compress: true }); const W = 792, H = 612, M = 30;
    const mes = mLabel(model.mes); const when = new Date();
    const title = (txt, y) => { doc.setFont('helvetica', 'bold'); doc.setFontSize(12); doc.setTextColor(17); doc.text(txt, W / 2, y, { align: 'center' }); const tw = doc.getTextWidth(txt); doc.setLineWidth(0.8); doc.setDrawColor(17); doc.line(W / 2 - tw / 2, y + 2, W / 2 + tw / 2, y + 2); };
    const logo = () => { try { if (window.LOGO_PNG) doc.addImage(window.LOGO_PNG, 'PNG', M, 16, 70, 21); } catch (e) { /* sin logo */ } };
    // portada
    try { if (window.LOGO_PNG) doc.addImage(window.LOGO_PNG, 'PNG', M, 26, 120, 36); } catch (e) { /* sin logo */ }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(32); doc.setTextColor(17);
    doc.text('Proyección de Productos', W / 2, 250, { align: 'center' }); doc.text('Principales en competencia', W / 2, 296, { align: 'center' });
    doc.text(mName(model.mes) + ' – ' + model.mes.slice(0, 4), W / 2, 342, { align: 'center' });
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(110);
    doc.text('Generado el ' + when.toLocaleString('es-VE') + '  ·  Departamento IDM', W / 2, 560, { align: 'center' });
    // general
    for (let i = 0; i < model.emps.length; i += 3) {
      doc.addPage(); logo(); title('General – ' + mes, 40);
      const chunk = model.emps.slice(i, i + 3); const cw = (W - 2 * M - 20 * (chunk.length - 1)) / chunk.length;
      chunk.forEach((e, j) => {
        const x = M + j * (cw + 20); const rows = e.cats.slice().sort((a, b) => b.usd - a.usd).map((c) => [c.def.nombre, usd(c.usd)]);
        doc.autoTable({
          startY: 70, margin: { left: x, right: W - x - cw }, tableWidth: cw, theme: 'grid',
          head: [[{ content: 'GENERAL REPRESENTACION ' + e.name + ' ($) – ' + mName(model.mes).toUpperCase(), colSpan: 2, styles: { fillColor: empCol(e.name), textColor: empTxt(e.name), halign: 'center' } }], ['CATEGORIA', 'PROYECCION']],
          body: rows.concat([[{ content: 'TOTAL', styles: { fontStyle: 'bold' } }, { content: usd(e.usd), styles: { fontStyle: 'bold' } }]]),
          styles: { fontSize: 8, cellPadding: 3, halign: 'center', textColor: 20, lineColor: [150, 150, 150], lineWidth: 0.4 }, headStyles: { fillColor: [235, 235, 235], textColor: 20 }
        });
        const y = doc.lastAutoTable.finalY + 14; doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(17);
        doc.text('REPRESENTACION ' + e.name + ' ($) – ' + mName(model.mes).toUpperCase(), x + cw / 2, y, { align: 'center' });
        const ph = Math.min(260, cw * 0.95);
        doc.addImage(pieImg(e.cats.map((c) => ({ label: c.def.nombre, value: c.usd })), 520, Math.round(520 * ph / cw), { legend: 'bottom', fs: 30 }), 'JPEG', x, y + 6, cw, ph);
      });
    }
    // categorías por empresa
    model.emps.forEach((e) => e.cats.forEach((c) => {
      doc.addPage(); logo(); title('Proyección ' + c.def.titulo + ' ' + e.name + ' – ' + mName(model.mes) + ':', 40);
      let y = 56; const vs = c.variants; const per = 13; const names = vs.map((v) => v.label);
      for (let s = 0; s < vs.length; s += per) {
        const part = vs.slice(s, s + per); const grp = []; part.forEach((v) => { const l = grp[grp.length - 1]; if (l && l.name === v.group) l.n++; else grp.push({ name: v.group, n: 1 }); });
        const first = s === 0;
        doc.autoTable({
          startY: y, margin: { left: M, right: M }, theme: 'grid',
          head: (first ? [[{ content: e.name, colSpan: part.length + 1, styles: { fillColor: empCol(e.name), textColor: empTxt(e.name), halign: 'center', fontStyle: 'bold' } }]] : [])
            .concat(c.groups.length > 1 || c.def.id === 'AA' ? [[{ content: 'PRODUCTOS', styles: { halign: 'left' } }].concat(grp.map((g) => ({ content: g.name, colSpan: g.n })))] : [])
            .concat([[''].concat(part.map((v) => v.label))]),
          body: [['PROYECCION ESTIMADA (CANTIDAD)'].concat(part.map((v) => f0(v.qty))), ['PRECIO REFERENCIA'].concat(part.map((v) => usd(v.price))), ['PROYECCION ESTIMADA ($)'].concat(part.map((v) => usd(v.usd)))],
          styles: { fontSize: 7, cellPadding: 2.5, halign: 'center', textColor: 20, lineColor: [150, 150, 150], lineWidth: 0.4, overflow: 'linebreak' },
          headStyles: { fillColor: [220, 230, 245], textColor: 20, fontStyle: 'bold' }, columnStyles: { 0: { halign: 'left', fontStyle: 'bold', cellWidth: 118 } }
        });
        y = doc.lastAutoTable.finalY + 8;
      }
      doc.setFillColor(255, 235, 120); doc.setDrawColor(120); doc.rect(M, y, 330, 15, 'FD'); doc.rect(M + 330, y, 90, 15, 'FD');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(17);
      doc.text('PROYECCION ESTIMADA GENERAL ' + c.def.nombre.toUpperCase() + ' ' + e.name + ' ($)', M + 6, y + 10.5); doc.text(usd(c.usd), M + 375, y + 10.5, { align: 'center' });
      y += 24;
      const room = H - 30 - y;
      if (room < 190) { doc.addPage(); logo(); y = 50; }
      const r2 = H - 30 - y; const colA = 250, colB = 240, colC = W - 2 * M - colA - colB - 20;
      // totales por grupo
      const gt = c.groups.map((g) => [g.name, f0(g.qty)]);
      doc.autoTable({ startY: y, margin: { left: M, right: W - M - colA }, tableWidth: colA, theme: 'grid', head: [[{ content: 'TOTAL ' + c.def.nombre.toUpperCase() + ' ' + e.name + ' – ' + mName(model.mes).toUpperCase(), colSpan: 2, styles: { halign: 'center' } }]], body: gt.concat([[{ content: 'TOTAL (CANTIDAD)', styles: { fontStyle: 'bold' } }, { content: f0(c.qty), styles: { fontStyle: 'bold' } }]]), styles: { fontSize: 7, cellPadding: 2, textColor: 20, lineColor: [150, 150, 150], lineWidth: 0.4 }, headStyles: { fillColor: [235, 235, 235], textColor: 20 }, columnStyles: { 1: { halign: 'right' } } });
      const by = doc.lastAutoTable.finalY + 6; const bh = Math.max(80, H - 30 - by);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.text('VENTA CANTIDAD ' + c.def.nombre.toUpperCase() + ' ' + e.name + ' – ' + mName(model.mes).toUpperCase(), M + colA / 2, by + 6, { align: 'center' });
      doc.addImage(barImg(c, 500, Math.round(500 * (bh - 10) / colA)), 'JPEG', M, by + 9, colA, bh - 10);
      // pasteles
      const xB = M + colA + 10; const hp = (r2 - 14) / 2;
      if (c.groups.length > 1) {
        doc.setFontSize(7.5); doc.text('CATEGORIAS ' + c.def.nombre.toUpperCase() + ' ' + e.name, xB + colB / 2, y + 6, { align: 'center' });
        doc.addImage(pieImg(c.groups.map((g) => ({ label: g.name, value: g.qty })), 480, Math.round(480 * (hp - 10) / colB), { legend: 'bottom', fs: 28 }), 'JPEG', xB, y + 9, colB, hp - 10);
      }
      const yv = c.groups.length > 1 ? y + hp + 6 : y; const hv = c.groups.length > 1 ? hp : r2;
      doc.setFontSize(7.5); doc.text('VARIANTES ' + c.def.nombre.toUpperCase() + ' ' + e.name, xB + colB / 2, yv + 6, { align: 'center' });
      doc.addImage(pieImg(vs.map((v) => ({ label: v.group.slice(0, 4) + ' ' + v.label, value: v.qty })), 480, Math.round(480 * (hv - 10) / colB), { legend: 'bottom', fs: 26, max: 8 }), 'JPEG', xB, yv + 9, colB, hv - 10);
      // narrativa
      const xC = xB + colB + 10; doc.setDrawColor(40); doc.setLineWidth(0.8); doc.rect(xC, y, colC, r2);
      doc.setFont('helvetica', 'normal'); doc.setTextColor(17); let fs = 8.5, lines;
      for (; fs >= 5.5; fs -= 0.5) { doc.setFontSize(fs); lines = doc.splitTextToSize(c.narr.replace(/\n---\n/g, '\n---\n'), colC - 14); if (lines.length * (fs + 2.2) <= r2 - 12) break; }
      doc.text(lines, xC + colC / 2, y + 12, { align: 'center', lineHeightFactor: 1.25, maxWidth: colC - 14 });
    }));
    // marcas por variante
    model.emps.forEach((e) => e.cats.forEach((c) => {
      const allB = Array.from(c.brands.entries()).sort((a, b) => b[1] - a[1]).map((x) => x[0]); const col = brandColors(allB.concat(['OTROS']));
      const cells = [{ t: 'GENERAL MARCAS ' + c.def.nombre.toUpperCase() + ' ' + e.name, m: c.brands }].concat(c.variants.filter((v) => v.qty > 0 && v.brands.size).map((v) => ({ t: (c.groups.length > 1 ? v.group + ' ' : '') + v.label, m: v.brands })));
      const per = 12, cols = 4, cw = (W - 2 * M - 30) / cols, chh = 150;
      for (let s = 0; s < cells.length; s += per) {
        doc.addPage(); logo(); title('Porcentaje Marcas ' + c.def.titulo + ' ' + e.name + ' – ' + mName(model.mes) + ':', 40);
        cells.slice(s, s + per).forEach((cell, i) => {
          const x = M + (i % cols) * (cw + 10), yy = 56 + Math.floor(i / cols) * (chh + 22);
          doc.setDrawColor(190); doc.setLineWidth(0.5); doc.rect(x, yy, cw, chh + 16);
          doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.setTextColor(17); doc.text(doc.splitTextToSize(cell.t, cw - 8)[0], x + cw / 2, yy + 10, { align: 'center' });
          doc.addImage(pieImg(Array.from(cell.m.entries()).map((b) => ({ label: b[0], value: b[1] })), 440, Math.round(440 * (chh - 4) / cw), { legend: 'right', fs: 30, max: 5, colors: col }), 'JPEG', x + 2, yy + 14, cw - 4, chh - 4);
        });
      }
    }));
    // notas
    doc.addPage(); logo(); title('Notas del informe', 40);
    const notes = ['Mes proyectado: ' + mes + ' (' + model.days + ' días). Datos tomados del detalle de productos cargado (' + f0(model.nRows) + ' filas del mes).',
      'Fórmula por sede, mes y variante: proyectado = Suma de lo observado x (1 + % de la sede) x (días del mes + 1) / (días estudiados + 1). Los días estudiados son los días distintos del detalle de esa sede en el mes.',
      model.imput ? 'Variantes no observadas: en las sedes donde una variante existe en otras sedes pero no se observó, se agrega el piso del Excel anterior: total de la variante / número de sedes de la empresa / días del mes x (días del mes − días estudiados).' : 'Variantes no observadas: no se agregó ningún piso; solo se proyecta lo observado.',
      'Precio de referencia: precio promedio ponderado observado por variante y empresa, redondeado a múltiplos de 5, salvo que se haya editado a mano. $ = cantidad proyectada x precio de referencia.',
      model.prev ? 'Comparación contra ' + mLabel(model.prev) + ' calculada con el mismo método y el mismo detalle.' : 'No hay datos del mes anterior en el detalle, por eso la redacción no compara contra otro mes.',
      'Filas del detalle sin clasificar en una variante (no se incluyen): ' + f0(Array.from(model.unc.values()).reduce((a, x) => a + x, 0)) + ' unidades en ' + model.unc.size + ' productos.'];
    doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(40); let ny = 70;
    notes.forEach((t) => { const l = doc.splitTextToSize('• ' + t, W - 2 * M); doc.text(l, M, ny); ny += l.length * 13 + 6; });
    // pie de página
    const n = doc.getNumberOfPages();
    for (let p = 2; p <= n; p++) { doc.setPage(p); doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(130); doc.text('Proyección de productos – ' + mes + '  ·  Página ' + p + ' de ' + n, W - M, H - 14, { align: 'right' }); }
    doc.save('informe_productos_' + model.mes + '.pdf');
  }

  // ---- Excel ----
  function exportXlsx() {
    if (!model) return; const wb = XLSX.utils.book_new(); const mes = mLabel(model.mes);
    const g = [['Empresa', 'Categoría', 'Cantidad proyectada', 'Proyección ($)', 'Mes anterior (cant.)', 'Variación']];
    model.emps.forEach((e) => e.cats.forEach((c) => g.push([e.name, c.def.nombre, Math.round(c.qty), Math.round(c.usd), c.prevQty == null ? '' : Math.round(c.prevQty), c.prevQty ? (c.qty - c.prevQty) / c.prevQty : ''])));
    const ws = XLSX.utils.aoa_to_sheet(g); ws['!cols'] = [{ wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 16 }, { wch: 20 }, { wch: 10 }];
    for (let r = 2; r <= g.length; r++) { const c = ws['F' + r]; if (c && c.v !== '') c.z = '0%'; ws['D' + r].z = '$#,##0'; }
    XLSX.utils.book_append_sheet(wb, ws, 'General');
    V.CATS.forEach((def) => {
      const a = [['Empresa', 'Grupo', 'Variante', 'Proyección (cantidad)', 'Precio referencia', 'Proyección ($)', 'Mes anterior (cant.)', 'Variación', 'Observado (unid.)', 'Piso agregado (unid.)']]; let has = false;
      model.emps.forEach((e) => { const c = e.cats.find((x) => x.id === def.id); if (!c) return; has = true;
        c.variants.forEach((v) => { const r = a.length + 1; a.push([e.name, v.group, v.label, Math.round(v.qty), v.price, { t: 'n', v: Math.round(v.qty) * v.price, f: 'D' + r + '*E' + r }, v.prevQty == null ? '' : Math.round(v.prevQty), v.prevQty ? (v.qty - v.prevQty) / v.prevQty : '', v.obs, Math.round(v.imp)]); }); });
      if (!has) return; const w = XLSX.utils.aoa_to_sheet(a); w['!cols'] = [{ wch: 12 }, { wch: 16 }, { wch: 18 }, { wch: 20 }, { wch: 16 }, { wch: 16 }, { wch: 18 }, { wch: 10 }, { wch: 16 }, { wch: 18 }];
      for (let r = 2; r <= a.length; r++) { w['E' + r].z = '$#,##0'; w['F' + r].z = '$#,##0'; if (w['H' + r] && w['H' + r].v !== '') w['H' + r].z = '0%'; }
      XLSX.utils.book_append_sheet(wb, w, def.nombre.replace('/', '').slice(0, 28));
    });
    const b = [['Empresa', 'Categoría', 'Variante', 'Marca', 'Unidades proyectadas (observado)', '% de la variante']];
    model.emps.forEach((e) => e.cats.forEach((c) => c.variants.forEach((v) => { const tot = Array.from(v.brands.values()).reduce((x, y) => x + y, 0) || 1; Array.from(v.brands.entries()).sort((x, y) => y[1] - x[1]).forEach((m) => b.push([e.name, c.def.nombre, v.group + ' ' + v.label, m[0], Math.round(m[1]), m[1] / tot])); })));
    const wb2 = XLSX.utils.aoa_to_sheet(b); for (let r = 2; r <= b.length; r++) wb2['F' + r].z = '0%'; wb2['!cols'] = [{ wch: 12 }, { wch: 16 }, { wch: 24 }, { wch: 22 }, { wch: 28 }, { wch: 16 }];
    XLSX.utils.book_append_sheet(wb, wb2, 'Marcas');
    const nr = [['Narrativa']]; model.emps.forEach((e) => e.cats.forEach((c) => { nr.push([e.name + ' – ' + c.def.titulo]); c.narr.split('\n').forEach((l) => nr.push([l])); nr.push(['']); }));
    const wn = XLSX.utils.aoa_to_sheet(nr); wn['!cols'] = [{ wch: 130 }]; XLSX.utils.book_append_sheet(wb, wn, 'Narrativa');
    const un = [['Empresa', 'Categoría', 'Producto sin clasificar', 'Unidades observadas']]; Array.from(model.unc.entries()).sort((a, b) => b[1] - a[1]).forEach((x) => { const k = x[0].split('||'); un.push([k[0], k[1], k[2], x[1]]); });
    const wu = XLSX.utils.aoa_to_sheet(un); wu['!cols'] = [{ wch: 12 }, { wch: 20 }, { wch: 70 }, { wch: 18 }] ; XLSX.utils.book_append_sheet(wb, wu, 'Sin clasificar');
    XLSX.writeFile(wb, 'informe_productos_' + model.mes + '.xlsx');
  }

  // ---- interfaz ----
  function renderBox() {
    const box = $('#infBox'); if (!box) return; const ms = monthsInDetail(); const p = P();
    if (!ms.length) { box.hidden = false; $('#infForm').hidden = true; $('#infEmpty').hidden = false; $('#infOut').innerHTML = ''; return; }
    box.hidden = false; $('#infForm').hidden = false; $('#infEmpty').hidden = true;
    const cur = $('#infMes').value; $('#infMes').innerHTML = ms.slice().reverse().map((m) => '<option value="' + m + '">' + esc(mLabel(m)) + '</option>').join(''); $('#infMes').value = ms.includes(cur) ? cur : ms[ms.length - 1];
    const emps = Array.from(new Set(p.state.detail.rows.map((r) => r.empresa))).sort((a, b) => { const ia = EMP_FIRST.indexOf(a), ib = EMP_FIRST.indexOf(b); return (ia < 0 ? 9 : ia) - (ib < 0 ? 9 : ib) || a.localeCompare(b); });
    const prev = new Set(Array.from(document.querySelectorAll('#infEmp input:checked')).map((i) => i.value)); const had = document.querySelectorAll('#infEmp input').length > 0;
    $('#infEmp').innerHTML = emps.map((e) => '<label class="chip"><input type="checkbox" value="' + esc(e) + '"' + (!had || prev.has(e) ? ' checked' : '') + '>' + esc(e) + '</label>').join('');
  }
  function prepare() {
    const empSet = new Set(Array.from(document.querySelectorAll('#infEmp input:checked')).map((i) => i.value));
    if (!empSet.size) return P().showMsg('Elige al menos una empresa.', true);
    keep.narr = {}; model = build({ mes: $('#infMes').value, imput: $('#infImput').checked, empSet: empSet });
    renderPreview();
  }
  function renderPreview() {
    const out = $('#infOut'); if (!model || !model.emps.length) { out.innerHTML = '<div class="card"><p class="hint">No hay productos clasificables en ese mes para las empresas elegidas.</p></div>'; $('#btnInfPdf').disabled = true; $('#btnInfXlsx').disabled = true; return; }
    const open = new Set(Array.from(out.querySelectorAll('details[open]')).map((d) => d.dataset.k));
    let h = '<div class="card"><h3>Vista previa · ' + esc(mLabel(model.mes)) + '</h3><p class="hint">Puedes corregir el precio de referencia (en $) y el texto de cada página. Se aplica al PDF y al Excel. '
      + (model.prev ? 'Se compara contra ' + esc(mLabel(model.prev)) + '.' : 'No hay mes anterior en el detalle, no se compara.') + '</p>';
    h += '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Empresa</th><th>Categoría</th><th class="num">Cantidad</th><th class="num">Proyección ($)</th></tr></thead><tbody>';
    model.emps.forEach((e) => e.cats.forEach((c) => { h += '<tr><td>' + esc(e.name) + '</td><td>' + esc(c.def.nombre) + '</td><td class="num">' + f0(c.qty) + '</td><td class="num">' + usd(c.usd) + '</td></tr>'; }));
    h += '</tbody></table></div></div>';
    model.emps.forEach((e) => e.cats.forEach((c) => {
      const k = e.name + '|' + c.id;
      h += '<details class="card" data-k="' + esc(k) + '"' + (open.has(k) ? ' open' : '') + '><summary><b>' + esc(e.name) + ' · ' + esc(c.def.titulo) + '</b> — ' + f0(c.qty) + ' unidades · ' + usd(c.usd) + '</summary>';
      h += '<div class="tbl-wrap"><table class="tbl"><thead><tr><th>Variante</th><th class="num">Cantidad</th><th class="num">Mes ant.</th><th class="num">Precio ref. ($)</th><th class="num">Proyección ($)</th></tr></thead><tbody>';
      c.variants.forEach((v) => { h += '<tr><td>' + esc(v.group + ' ' + v.label) + '</td><td class="num">' + f0(v.qty) + (v.imp > 0.5 ? ' <span class="hint" title="Incluye ' + f0(v.imp) + ' unidades de piso por variante no observada en algunas sedes">(+' + f0(v.imp) + ')</span>' : '') + '</td><td class="num">' + (v.prevQty == null ? '—' : f0(v.prevQty)) + '</td><td class="num"><input type="number" min="0" step="5" style="width:90px" data-pk="' + esc(v.pk) + '" value="' + v.price + '"></td><td class="num">' + usd(v.usd) + '</td></tr>'; });
      h += '</tbody></table></div><div class="f" style="margin-top:10px"><label>Texto de la página</label><textarea data-nk="' + esc(c.nk) + '" rows="8" style="width:100%">' + esc(c.narr) + '</textarea></div></details>';
    }));
    const un = Array.from(model.unc.entries()).sort((a, b) => b[1] - a[1]);
    if (un.length) { h += '<details class="card"><summary><b>Productos que no se pudieron clasificar</b> — ' + un.length + ' (no entran al informe)</summary><div class="tbl-wrap"><table class="tbl"><thead><tr><th>Empresa</th><th>Producto</th><th class="num">Unid.</th></tr></thead><tbody>' + un.slice(0, 60).map((x) => { const q = x[0].split('||'); return '<tr><td>' + esc(q[0]) + '</td><td>' + esc(q[2]) + '</td><td class="num">' + f0(x[1]) + '</td></tr>'; }).join('') + '</tbody></table></div><p class="hint">Suelen ser nombres sin tamaño (ej. "LAVADORA SEMIAUTOMATICA DAMASCO") o productos mal categorizados en el portal.</p></details>'; }
    out.innerHTML = h; $('#btnInfPdf').disabled = false; $('#btnInfXlsx').disabled = false;
  }
  function bind() {
    const box = $('#infBox'); if (!box || box.dataset.b) return; box.dataset.b = '1';
    $('#btnInfGen').addEventListener('click', prepare);
    $('#btnInfPdf').addEventListener('click', () => { try { exportPdf(); } catch (e) { console.error(e); P().showMsg('No se pudo generar el PDF: ' + e.message, true); } });
    $('#btnInfXlsx').addEventListener('click', () => { try { exportXlsx(); } catch (e) { console.error(e); P().showMsg('No se pudo generar el Excel: ' + e.message, true); } });
    $('#infOut').addEventListener('change', (ev) => {
      const t = ev.target; if (t.dataset && t.dataset.pk) { keep.prices[t.dataset.pk] = Math.max(0, +t.value || 0); model.emps.forEach((e) => e.cats.forEach((c) => c.variants.forEach((v) => { if (v.pk === t.dataset.pk) v.price = keep.prices[v.pk]; }))); recalc(model); renderPreview(); }
    });
    $('#infOut').addEventListener('input', (ev) => { const t = ev.target; if (t.dataset && t.dataset.nk) { keep.narr[t.dataset.nk] = t.value; model.emps.forEach((e) => e.cats.forEach((c) => { if (c.nk === t.dataset.nk) c.narr = t.value; })); } });
  }
  function refresh() { bind(); model = null; keep.narr = {}; renderBox(); const o = $('#infOut'); if (o) o.innerHTML = ''; ['#btnInfPdf', '#btnInfXlsx'].forEach((s) => { const b = $(s); if (b) b.disabled = true; }); }
  window.Informe = { build: build, raw: raw, refresh: refresh, _model: () => model };
  const start = () => { if (!window.__proy) return setTimeout(start, 30); window.__proy.hooks.push(refresh); refresh(); };
  start();
})();
