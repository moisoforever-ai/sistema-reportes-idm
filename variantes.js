/* Clasificador de variantes (tipo / tamaño) por categoría, a partir del nombre del producto. */
(function (root) {
  'use strict';
  const up = (s) => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
  const num = (s) => parseFloat(String(s).replace(',', '.'));

  const CATS = [
    { id: 'AA', nombre: 'A/A', titulo: 'Aires acondicionados', test: /AIRE|A\/A/, groups: ['PORTATIL', 'SPLIT', 'VENTANA', 'TON'], unidad: 'unidades' },
    { id: 'TV', nombre: 'Televisores', titulo: 'Televisores', test: /TELEVIS/, groups: ['TV'], unidad: 'unidades' },
    { id: 'NEV', nombre: 'Neveras', titulo: 'Neveras', test: /NEVERA/, groups: ['EJECUTIVO', 'TOP FREEZER', 'BOTTOM FREEZER', 'SBS'], unidad: 'unidades' },
    { id: 'LAV', nombre: 'Lavadoras', titulo: 'Lavadoras', test: /LAVA/, groups: ['SEMIAUTOMATICA', 'AUTOMATICA'], unidad: 'unidades' },
    { id: 'CONG', nombre: 'Congeladores', titulo: 'Congeladores', test: /CONGEL/, groups: ['HORIZONTAL', 'VERTICAL'], unidad: 'unidades' }
  ];
  function catOf(categoria) { const c = up(categoria); return CATS.find((x) => x.test.test(c)) || null; }

  function btu(n) {
    let m = n.match(/(\d{1,2})\s*[.,]?\s*000\s*(?:BTU|BU)/) || n.match(/(\d{1,2})\s*K\s*(?:BTU)?\b/) || n.match(/(\d{1,2})\s*MIL/) || n.match(/(\d{4,5})\s*(?:BTU|BU)/);
    if (!m) return null;
    let v = +m[1]; if (v >= 1000) v = Math.round(v / 1000); return (v >= 4 && v <= 60) ? v : null;
  }
  function clsAA(n, emp) {
    const ton = n.match(/(\d(?:[.,]\d)?)\s*(?:TON\b|TONELADA)/);
    if (ton) return { group: 'TON', label: num(ton[1]) + ' TON', ord: num(ton[1]) };
    const b = btu(n) || (/PORTATIL/.test(n) ? 12 : null); if (!b) return null;
    const v = /\b220\s*V?\b/.test(n) ? '220V' : (/\b110\s*V?\b/.test(n) ? '110V' : '');
    if (/PORTATIL/.test(n)) return { group: 'PORTATIL', label: b + ' MIL PORTATIL', ord: b };
    if (/SPLIT/.test(n)) {
      const volt = (emp === 'DAMASCO' && b === 12 && v) ? ' ' + v : '';
      return { group: 'SPLIT', label: b + ' MIL' + volt, ord: b + (v === '220V' ? 0.2 : 0) };
    }
    const volt = (b === 12 && v) ? ' (' + v + ')' : '';
    return { group: 'VENTANA', label: b + ' MIL' + volt, ord: b + (v === '220V' ? 0.2 : 0) };
  }
  function clsTV(n) {
    const m = n.match(/\b(\d{2,3})\s*(?:"|''|”|“|'|PULG|IN\b)/) || n.match(/\bTV\s*(\d{2})\b/) || n.match(/\bDE\s+(\d{2})\b/) || n.match(/\b(\d{2})\s+(?:SMART|LED|DLED|QLED|SLIM|UHD|HD|FHD|4K|GOOGLE)/);
    if (!m) return null; const s = +m[1]; if (s < 19 || s > 110) return null;
    const smart = /SMART|GOOGLE|ANDROID|TIZEN|WEBOS|ROKU|FIRE/.test(n);
    return { group: 'TV', label: s + '" ' + (smart ? 'S' : 'NS'), ord: s + (smart ? 0.5 : 0) };
  }
  function clsLav(n) {
    if (/SECADORA|HIDROLAV|CENTRO DE LAVADO/.test(n) && !/LAVADORA/.test(n.replace(/HIDROLAVADORA/g, ''))) return { skip: true };
    if (/HIDROLAV/.test(n)) return { skip: true };
    const k = n.match(/(\d{1,2})\s*K(?:G)?\b/); if (!k) return null; const kg = +k[1]; if (kg < 4 || kg > 25) return null;
    const g = /SEMI/.test(n) ? 'SEMIAUTOMATICA' : (/AUTOMATIC|LAVASECADORA|LAVA-SECADORA|CARGA FRONTAL|CARGA SUPERIOR/.test(n) ? 'AUTOMATICA' : (kg <= 8 ? 'SEMIAUTOMATICA' : 'AUTOMATICA'));
    return { group: g, label: kg + 'KG', ord: kg };
  }
  const CONG_STD = [99, 142, 170, 200, 226, 300];
  function snapL(l) { let best = null; CONG_STD.forEach((s) => { if (Math.abs(l - s) / s <= 0.08 && (best == null || Math.abs(l - s) < Math.abs(l - best))) best = s; }); return best || Math.round(l); }
  function clsCong(n) {
    let l = null; const L = n.match(/(\d{2,3})\s*(?:L\b|LT|LTS|LITROS)/); const P = n.match(/(\d{1,2}(?:[.,]\d)?)\s*(?:P\b|PIE)/);
    if (L) l = +L[1]; else if (P) l = Math.round(num(P[1]) * 28.3); else { const k = n.match(/\bBD[- ]?(\d{2,3})\b/); if (k) l = +k[1]; }
    if (!l) return null; l = snapL(l);
    return { group: /VERTICAL/.test(n) ? 'VERTICAL' : 'HORIZONTAL', label: l + 'L', ord: l };
  }
  function clsNev(n) {
    let p = null; const P = n.match(/(\d{1,2}(?:[.,]\d)?)\s*(?:P\b|PIE|FT|'|’)/); const L = n.match(/(\d{2,3})\s*(?:L\b|LT|LTS|LITROS)/);
    if (P) p = num(P[1]); else if (L) p = Math.max(1, +L[1] / 28.3);
    let g = 'TOP FREEZER';
    if (/EJECUTIV/.test(n)) g = 'EJECUTIVO'; else if (/SBS|SIDE BY SIDE|SIDE-BY-SIDE|FRENCH|DOOR/.test(n)) g = 'SBS'; else if (/BOTTOM|\bBF\b|\bBM\b/.test(n)) g = 'BOTTOM FREEZER';
    if (g === 'EJECUTIVO' && p == null) p = 1.6;
    if (p == null) return null;
    if (g === 'EJECUTIVO' && p < 3) p = 1.6; else p = Math.round(p);
    return { group: g, label: p + 'P', ord: p };
  }
  const FN = { AA: clsAA, TV: clsTV, LAV: clsLav, CONG: clsCong, NEV: clsNev };
  // devuelve {cat, group, label, ord} o null (sin categoría de informe / no clasificable → motivo)
  function classify(categoria, producto, empresa) {
    const c = catOf(categoria); if (!c) return { cat: null };
    const v = FN[c.id](up(producto), up(empresa));
    if (v && v.skip) return { cat: null, skip: true };
    return v ? { cat: c.id, group: v.group, label: v.label, ord: v.ord } : { cat: c.id, fail: true };
  }
  root.Variantes = { CATS: CATS, classify: classify, catOf: catOf, up: up };
  if (typeof module !== 'undefined') module.exports = root.Variantes;
})(typeof window !== 'undefined' ? window : globalThis);
