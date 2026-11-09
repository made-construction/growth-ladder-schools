/* charts.js – رسوم SVG خفيفة بهوية سُلّم النمو (بدون مكتبات خارجية) */
(function (root) {
  'use strict';
  const C = { petrol: '#0F4C5C', saffron: '#E9A23B', aqua: '#5BA4A4', dark: '#0A3540', grey: '#4E6168', night: '#10242B', neg: '#B5473A', pos: '#2F7D5B', sand: '#F3D9A8', mist: '#DCEBEB' };
  const PAL = [C.petrol, C.saffron, C.aqua, C.dark, '#8FBFBF', '#C98A2B', '#9AAAAE', '#3F7F8C'];
  const NS = 'http://www.w3.org/2000/svg';
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  let tipEl;
  function tip() { if (!tipEl) { tipEl = document.createElement('div'); tipEl.className = 'tip'; document.body.appendChild(tipEl); } return tipEl; }
  function bindTips(svg) {
    svg.querySelectorAll('[data-tip]').forEach(el => {
      el.addEventListener('mousemove', e => { const t = tip(); t.innerHTML = el.getAttribute('data-tip'); t.classList.add('show'); const w = t.offsetWidth; let x = e.clientX + 14; if (x + w > innerWidth - 8) x = e.clientX - w - 14; t.style.left = x + 'px'; t.style.top = (e.clientY + 14) + 'px'; });
      el.addEventListener('mouseleave', () => tip().classList.remove('show'));
      el.addEventListener('touchstart', e => { const t = tip(); t.innerHTML = el.getAttribute('data-tip'); t.classList.add('show'); const r = el.getBoundingClientRect(); t.style.left = Math.max(8, Math.min(innerWidth - t.offsetWidth - 8, r.left)) + 'px'; t.style.top = (r.top - t.offsetHeight - 8) + 'px'; setTimeout(() => t.classList.remove('show'), 2200); }, { passive: true });
    });
  }

  function niceScale(min, max, ticks = 5) {
    if (min === max) { max = min + 1; }
    const span = max - min, step0 = span / ticks, mag = Math.pow(10, Math.floor(Math.log10(step0)));
    const step = [1, 2, 2.5, 5, 10].map(m => m * mag).find(s => span / s <= ticks) || 10 * mag;
    return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
  }
  const shortNum = v => { const a = Math.abs(v); if (a >= 1e6) return (v / 1e6).toFixed(a >= 1e7 ? 0 : 1).replace(/\.0$/, '') + ' م'; if (a >= 1e3) return (v / 1e3).toFixed(0) + ' أ'; return String(Math.round(v * 100) / 100); };

  function legend(series) {
    return '<div class="lgd">' + series.map(s => `<span><i style="background:${s.color}${s.type === 'line' ? ';height:3px;border-radius:2px' : ''}"></i>${esc(s.name)}</span>`).join('') + '</div>';
  }

  /* أعمدة مجمعة/مكدسة + خطوط (يدعم القيم السالبة) */
  function combo(el, o) {
    const W = o.width || 720, H = o.height || 300, padL = 58, padR = 14, padT = 14, padB = 34;
    const fmt = o.fmt || shortNum, tfmt = o.tipFmt || fmt;
    const series = o.series.map((s, i) => ({ type: 'bar', color: PAL[i % PAL.length], ...s }));
    const bars = series.filter(s => s.type === 'bar'), lines = series.filter(s => s.type === 'line');
    let lo = 0, hi = 0;
    o.cats.forEach((_, i) => {
      if (o.stacked) { let p = 0, n = 0; bars.forEach(s => { const v = s.values[i] || 0; if (v >= 0) p += v; else n += v; }); hi = Math.max(hi, p); lo = Math.min(lo, n); }
      else bars.forEach(s => { const v = s.values[i] || 0; hi = Math.max(hi, v); lo = Math.min(lo, v); });
      lines.forEach(s => { const v = s.values[i]; if (v != null && isFinite(v)) { hi = Math.max(hi, v); lo = Math.min(lo, v); } });
    });
    const sc = niceScale(lo, hi);
    const y = v => padT + (sc.max - v) / (sc.max - sc.min) * (H - padT - padB);
    const n = o.cats.length, band = (W - padL - padR) / n, bw = o.stacked ? band * 0.5 : Math.min(42, band * 0.72 / Math.max(1, bars.length));
    let g = `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.title || '')}"><g class="grid">`;
    for (let v = sc.min; v <= sc.max + 1e-9; v += sc.step) g += `<line x1="${padL}" x2="${W - padR}" y1="${y(v)}" y2="${y(v)}"/><text x="${padL - 8}" y="${y(v) + 4}" text-anchor="end">${esc(fmt(v))}</text>`;
    g += `</g><line class="zero" x1="${padL}" x2="${W - padR}" y1="${y(0)}" y2="${y(0)}"/>`;
    o.cats.forEach((c, i) => {
      const cx = padL + band * i + band / 2;
      g += `<text x="${cx}" y="${H - 10}" text-anchor="middle">${esc(c)}</text>`;
      let pos = 0, neg = 0;
      bars.forEach((s, j) => {
        const v = s.values[i] || 0; let x, y0, y1;
        if (o.stacked) { x = cx - bw / 2; if (v >= 0) { y0 = y(pos + v); y1 = y(pos); pos += v; } else { y0 = y(neg); y1 = y(neg + v); neg += v; } }
        else { x = cx - bars.length * bw / 2 + j * bw; y0 = y(Math.max(0, v)); y1 = y(Math.min(0, v)); }
        const col = (s.negColor && v < 0) ? s.negColor : s.color;
        g += `<rect x="${x + 1}" y="${y0}" width="${Math.max(1, bw - 2)}" height="${Math.max(0.5, y1 - y0)}" rx="3" fill="${col}" data-tip="${esc(c)} · ${esc(s.name)}<br><b>${esc(tfmt(v))}</b>"/>`;
      });
    });
    lines.forEach(s => {
      const pts = s.values.map((v, i) => v == null ? null : [padL + band * i + band / 2, y(v)]);
      const d = pts.filter(Boolean).map((p, i) => (i ? 'L' : 'M') + p[0].toFixed(1) + ' ' + p[1].toFixed(1)).join(' ');
      g += `<path d="${d}" fill="none" stroke="${s.color}" stroke-width="2.5" ${s.dash ? 'stroke-dasharray="6 5"' : ''}/>`;
      pts.forEach((p, i) => { if (p) g += `<circle cx="${p[0]}" cy="${p[1]}" r="4.5" fill="#fff" stroke="${s.color}" stroke-width="2.5" data-tip="${esc(o.cats[i])} · ${esc(s.name)}<br><b>${esc(tfmt(s.values[i]))}</b>"/>`; });
    });
    g += '</svg>';
    el.innerHTML = `<div class="chart">${g}${o.legend === false ? '' : legend(series)}</div>`;
    bindTips(el.querySelector('svg'));
  }

  /* أعمدة أفقية (تورنادو للحساسية أو ترتيب) */
  function hbars(el, o) {
    const items = o.items, W = o.width || 720, rowH = o.rowH || 34, padT = 8, labelW = o.labelW || 190, padR = 70;
    const H = padT * 2 + rowH * items.length;
    const fmt = o.fmt || shortNum;
    let lo = Math.min(0, ...items.map(i => i.value)), hi = Math.max(0, ...items.map(i => i.value));
    if (lo === hi) hi = lo + 1;
    // RTL: التسميات يمين، الأعمدة تمتد لليسار للقيم الموجبة (اتجاه القراءة)
    const x0 = W - labelW, x1 = padR;
    const x = v => x0 - (v - lo) / (hi - lo) * (x0 - x1);
    let g = `<svg viewBox="0 0 ${W} ${H}">`;
    g += `<line class="zero" x1="${x(0)}" x2="${x(0)}" y1="${padT - 4}" y2="${H - padT + 4}"/>`;
    items.forEach((it, i) => {
      const yy = padT + i * rowH, a = x(Math.max(0, it.value)), b = x(Math.min(0, it.value));
      const col = it.color || (it.value < 0 ? C.neg : C.petrol);
      g += `<text x="${W - 4}" y="${yy + rowH / 2 + 4}" text-anchor="end" style="font-size:12px;fill:#10242B">${esc(it.label)}</text>`;
      g += `<rect x="${Math.min(a, b)}" y="${yy + 7}" width="${Math.max(1, Math.abs(b - a))}" height="${rowH - 14}" rx="4" fill="${col}" data-tip="${esc(it.label)}<br><b>${esc((o.tipFmt || fmt)(it.value))}</b>"/>`;
      const tx = it.value < 0 ? Math.max(a, b) + 6 : Math.min(a, b) - 6;
      g += `<text x="${tx}" y="${yy + rowH / 2 + 4}" text-anchor="${it.value < 0 ? 'start' : 'end'}" style="font-weight:700;fill:${col}">${esc(fmt(it.value))}</text>`;
    });
    g += '</svg>';
    el.innerHTML = `<div class="chart">${g}</div>`;
    bindTips(el.querySelector('svg'));
  }

  /* حلقة */
  function donut(el, o) {
    const items = o.items.filter(i => i.value > 0), tot = items.reduce((s, i) => s + i.value, 0) || 1;
    const S = 220, R = 92, r = 60, cx = S / 2, cy = S / 2;
    let a = -Math.PI / 2, g = `<svg viewBox="0 0 ${S} ${S}" style="max-width:240px;margin:auto">`;
    items.forEach((it, i) => {
      const col = it.color || PAL[i % PAL.length], da = it.value / tot * Math.PI * 2, a2 = a + da, large = da > Math.PI ? 1 : 0;
      const p = (rr, t) => [cx + rr * Math.cos(t), cy + rr * Math.sin(t)];
      const [x1, y1] = p(R, a), [x2, y2] = p(R, a2), [x3, y3] = p(r, a2), [x4, y4] = p(r, a);
      const d = da >= Math.PI * 2 - 1e-6 ? `M${cx - R} ${cy}a${R} ${R} 0 1 0 ${2 * R} 0a${R} ${R} 0 1 0 ${-2 * R} 0M${cx - r} ${cy}a${r} ${r} 0 1 1 ${2 * r} 0a${r} ${r} 0 1 1 ${-2 * r} 0` : `M${x1} ${y1}A${R} ${R} 0 ${large} 1 ${x2} ${y2}L${x3} ${y3}A${r} ${r} 0 ${large} 0 ${x4} ${y4}Z`;
      g += `<path d="${d}" fill="${col}" stroke="#fff" stroke-width="1.5" data-tip="${esc(it.label)}<br><b>${esc((o.fmt || shortNum)(it.value))}</b> · ${(it.value / tot * 100).toFixed(1)}%"/>`;
      a = a2;
    });
    if (o.center) g += `<text x="${cx}" y="${cy - 4}" text-anchor="middle" style="font-size:20px;font-weight:800;fill:#10242B">${esc(o.center[0])}</text><text x="${cx}" y="${cy + 16}" text-anchor="middle" style="font-size:11px">${esc(o.center[1] || '')}</text>`;
    g += '</svg>';
    const lg = '<div class="lgd" style="flex-direction:column;gap:6px">' + items.map((it, i) => `<span style="display:flex;justify-content:space-between;gap:10px"><span><i style="background:${it.color || PAL[i % PAL.length]}"></i>${esc(it.label)}</span><b class="num">${(it.value / tot * 100).toFixed(1)}%</b></span>`).join('') + '</div>';
    el.innerHTML = `<div class="chart" style="display:grid;grid-template-columns:minmax(160px,240px) 1fr;gap:18px;align-items:center">${g}${lg}</div>`;
    bindTips(el.querySelector('svg'));
  }

  root.GLCharts = { combo, hbars, donut, C, PAL, shortNum };
})(window);
