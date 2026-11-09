/* app.js – واجهة الموقع التفاعلي لدراسة جدوى مدارس سُلّم النمو
 * كل رقم معروض يُقرأ من محرك المعادلات (engine.js) الذي يشغّل معادلات ملف Excel v2 نفسها.
 * لا توجد نتائج محسوبة مكتوبة يدوياً في هذا الملف.
 */
(function () {
  'use strict';
  const M = window.GL_MODEL, G = window.GLEngine, CH = window.GLCharts;
  const E = new G.Engine(M);
  const STORE = 'gl-study-v2';           // حالة العمل الحالية
  const SNAPS = 'gl-study-v2-snapshots'; // النسخ المحفوظة بأسماء

  /* ================= أدوات عامة ================= */
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const SHEET = {}; M.sheets.forEach((n, i) => { SHEET[n.split(' ')[0]] = i; });
  const sname = code => M.sheets[SHEET[code]];
  const K = (code, a) => SHEET[code] + '!' + a;            // مفتاح خلية: رمز الورقة + العنوان
  const V = (code, a) => E.getKey(K(code, a));              // قيمة خلية
  const N = (code, a) => { const v = V(code, a); return typeof v === 'number' ? v : NaN; };
  const NM = name => E.val(name);                           // قيمة اسم معرّف
  const isNum = v => typeof v === 'number' && isFinite(v);
  const YRS = ['السنة 1', 'السنة 2', 'السنة 3', 'السنة 4', 'السنة 5'];
  const YC = ['E', 'F', 'G', 'H', 'I'];                     // أعمدة السنوات في ورقة 17

  const nf0 = new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('en-US', { maximumFractionDigits: 1, minimumFractionDigits: 1 });
  const nf2 = new Intl.NumberFormat('en-US', { maximumFractionDigits: 2, minimumFractionDigits: 2 });
  const fInt = v => isNum(v) ? nf0.format(Math.round(v)) : esc(v);
  const fSAR = v => isNum(v) ? nf0.format(Math.round(v)) + ' ر.س' : esc(v);
  const fM = v => isNum(v) ? nf2.format(v / 1e6) + ' مليون' : esc(v);
  const fPct = (v, d = 1) => isNum(v) ? (d ? nf1 : nf0).format(v * 100) + '%' : esc(v);
  // يعزل الأرقام فقط باتجاه LTR ويترك الكلمات العربية في سياقها
  const numSpan = s => String(s).replace(/-?\d[\d,.]*%?/g, m => `<span class="num">${m}</span>`);

  /* تنسيق قيمة خلية حسب تنسيق Excel الخاص بها */
  function fmtCell(v, fmt) {
    if (v === null || v === undefined) return '';
    if (G.isErr(v)) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v !== 'number') return String(v);
    if (fmt === 'yyyy-mm-dd') { const d = G.toDate(v); return d.toISOString().slice(0, 10); }
    if (!fmt || fmt === 'General') {
      if (Number.isInteger(v)) return nf0.format(v);
      const a = Math.abs(v);
      if (a < 1 && a > 0) return String(+v.toPrecision(4));
      return new Intl.NumberFormat('en-US', { maximumFractionDigits: a >= 100 ? 0 : 2 }).format(v);
    }
    // تنسيقات الأعداد الصحيحة تُخفي كسوراً مهمة في القيم الصغيرة (مثل 4.70 م²/طالب) فتُعرض بمنزلتين
    if (/^#,##0[^.]*$/.test(fmt.split(';')[0]) && !Number.isInteger(v) && Math.abs(v) < 100 && Math.abs(v - Math.round(v)) > 0.004) {
      const lit = fmt.split(';')[0].replace('#,##0', nf2.format(Math.abs(v))).replace(/"/g, '');
      return (v < 0 ? '-' : '') + lit;
    }
    return G.fmtNum(v, fmt);
  }

  /* ================= الحالة والحفظ ================= */
  function saveState() {
    try { localStorage.setItem(STORE, JSON.stringify({ over: E.over, at: new Date().toISOString(), source: M.source })); } catch (e) { }
    updateChangedBadge();
  }
  function loadState() {
    try { const s = JSON.parse(localStorage.getItem(STORE) || 'null'); if (s && s.over) { E.over = filterOver(s.over); E.cache.clear(); } } catch (e) { }
  }
  const filterOver = o => { const r = {}; for (const [k, v] of Object.entries(o || {})) if (M.inp[k]) r[k] = v; return r; };
  const changedCount = () => Object.keys(E.over).length;
  function updateChangedBadge() {
    const n = changedCount(), b = $('#chg');
    b.textContent = n ? `${fInt(n)} مدخل معدّل` : ''; b.classList.toggle('show', n > 0);
    $('#btn-reset').disabled = !n;
  }
  function setInput(key, val) { E.set(key, val); saveState(); rerender(); toast('تم تحديث النتائج'); }
  function resetAll() { if (!confirm('إعادة جميع المدخلات إلى القيم الأصلية في دراسة v2؟ (لن تُحذف النسخ المحفوظة)')) return; E.reset(); saveState(); rerender(); toast('أُعيدت القيم الأصلية'); }

  const snaps = () => { try { return JSON.parse(localStorage.getItem(SNAPS) || '[]'); } catch (e) { return []; } };
  function snapSave() {
    const name = prompt('اسم النسخة المحفوظة (مثال: سيناريو رسوم أعلى):', 'نسخة ' + new Date().toLocaleString('ar-SA-u-nu-latn'));
    if (!name) return;
    const list = snaps(); list.unshift({ name, at: new Date().toISOString(), over: E.over, scenario: NM('Scenario') });
    localStorage.setItem(SNAPS, JSON.stringify(list.slice(0, 30))); toast('حُفظت النسخة في هذا المتصفح'); openSnaps();
  }
  function openSnaps() {
    const list = snaps();
    modal(`<h3>حفظ التعديلات واستعادتها</h3>
      <div class="note info small">تُحفظ التعديلات <b>محلياً في هذا المتصفح وعلى هذا الجهاز فقط</b> (Local Storage)، ولا تُزامَن تلقائياً بين الأجهزة أو المتصفحات، وتُفقد إذا مُسحت بيانات المتصفح. لنقلها إلى جهاز آخر استخدم «تصدير ملف» ثم «استيراد ملف».</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin:14px 0">
        <button class="btn pri" data-a="snap-save">حفظ الوضع الحالي كنسخة</button>
        <button class="btn" data-a="export">تصدير ملف (JSON)</button>
        <label class="btn" style="cursor:pointer">استيراد ملف<input type="file" accept=".json,application/json" data-a="import" hidden></label>
      </div>
      ${list.length ? `<div class="tbl-wrap"><table class="t"><thead><tr><th>النسخة</th><th>التاريخ</th><th>المدخلات المعدلة</th><th></th></tr></thead><tbody>${list.map((s, i) => `<tr><td><b>${esc(s.name)}</b><div class="small muted">السيناريو: ${esc(s.scenario || '')}</div></td><td class="small">${esc(new Date(s.at).toLocaleString('ar-SA-u-nu-latn'))}</td><td class="n">${Object.keys(s.over || {}).length}</td><td style="white-space:nowrap"><button class="btn" data-a="snap-load" data-i="${i}">استعادة</button> <button class="btn" data-a="snap-del" data-i="${i}">حذف</button></td></tr>`).join('')}</tbody></table></div>` : '<p class="muted small">لا توجد نسخ محفوظة بعد.</p>'}
      <div style="margin-top:14px;text-align:left"><button class="btn" data-a="close">إغلاق</button></div>`);
  }
  function exportJSON() {
    const data = { app: 'growth-ladder-study', source: M.source, exported: new Date().toISOString(), scenario: NM('Scenario'), over: E.over,
      labels: Object.fromEntries(Object.keys(E.over).map(k => [k, labelOf(k)])) };
    const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }));
    a.download = 'سلم_النمو_مدخلات_' + new Date().toISOString().slice(0, 10) + '.json'; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
  function importJSON(file) {
    const r = new FileReader(); r.onload = () => { try { const d = JSON.parse(r.result); E.over = filterOver(d.over); E.cache.clear(); saveState(); rerender(); closeModal(); toast('استُوردت المدخلات (' + changedCount() + ')'); } catch (e) { alert('ملف غير صالح'); } }; r.readAsText(file);
  }

  /* وصف خلية إدخال: نص البند في نفس الصف */
  function labelOf(key) {
    const [s, a] = key.split('!'); const r = +a.match(/\d+/)[0];
    const name = Object.entries(M.names).find(([, k]) => k === key);
    for (const c of ['B', 'C']) { const v = M.v[s + '!' + c + r]; if (typeof v === 'string' && !M.inp[s + '!' + c + r]) return v + (name ? ` (${name[0]})` : ''); }
    return M.sheets[s] + ' – ' + a;
  }

  /* ================= واجهات مساعدة ================= */
  let toastT;
  function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('show'), 1600); }
  function modal(html) { const m = $('#modal'); $('.box', m).innerHTML = html; m.classList.add('show'); }
  function closeModal() { $('#modal').classList.remove('show'); }

  /* زر «طريقة الحساب» لكل مؤشر: يعرض المعادلة الأصلية ومصدرها في ملف Excel */
  function howBtn(key, extra) { return `<button class="how no-print" data-how="${esc(key)}" data-extra="${esc(extra || '')}" title="طريقة الحساب ومصدر الرقم">؟</button>`; }
  function showHow(key, extra) {
    const k = E.resolve(key), [s, a] = k.split('!');
    const fx = M.fx[k]; const name = Object.entries(M.names).find(([, kk]) => kk === k);
    const v = E.getKey(k);
    modal(`<h3>طريقة الحساب</h3>
      <p class="small muted">كل مؤشر في الموقع مرتبط بخلية في ملف الدراسة v2 ويُعاد حسابه بنفس معادلتها عند تغيير المدخلات.</p>
      <table class="t"><tbody>
      <tr><th>الورقة</th><td>${esc(M.sheets[s])}</td></tr>
      <tr><th>الخلية</th><td><span class="code">${esc(a)}</span>${name ? ` · الاسم المعرّف <span class="code">${esc(name[0])}</span>` : ''}</td></tr>
      <tr><th>البند</th><td>${esc(labelOf(k))}</td></tr>
      <tr><th>القيمة الحالية</th><td><b>${esc(fmtCell(v, E.fmt(k)))}</b></td></tr>
      ${fx ? `<tr><th>المعادلة</th><td><span class="code block">${esc(fx)}</span></td></tr>` : '<tr><th>النوع</th><td>مدخل (قابل للتعديل)</td></tr>'}
      ${extra ? `<tr><th>ملاحظة</th><td>${esc(extra)}</td></tr>` : ''}
      </tbody></table>
      <div style="margin-top:14px;text-align:left"><button class="btn" data-a="close">إغلاق</button></div>`);
  }

  /* بطاقة مؤشر */
  function kpi({ label, value, unit, foot, key, tone, bar, extra }) {
    return `<div class="card kpi ${tone || ''}"><div class="lb">${esc(label)}${key ? howBtn(key, extra) : ''}</div>
      <div class="vl">${value}${unit ? `<small>${esc(unit)}</small>` : ''}</div>
      ${foot ? `<div class="ft">${foot}</div>` : ''}${bar != null ? `<div class="bar"><i style="width:${Math.max(0, Math.min(100, bar * 100))}%"></i></div>` : ''}</div>`;
  }

  /* وسوم تصنيف موثوقية البيانات */
  function classTag(txt) {
    const t = String(txt || '');
    if (!t) return '';
    let c = 'c-est';
    if (/اشتراط رسمي|ساري|مؤكد|رسمي/.test(t) && !/يحتاج|مسودة|سابق/.test(t)) c = 'c-ok';
    else if (/يحتاج|تحقق|اعتماد|مطابقة|مسودة|سابق/.test(t)) c = 'c-ver';
    else if (/المستثمر|مطلوب من/.test(t)) c = 'c-inv';
    else if (/افتراض/.test(t)) c = 'c-as';
    else if (/غير مطابق|عجز|مفتوح/.test(t)) c = 'c-neg';
    return `<span class="tag ${c}">${esc(t)}</span>`;
  }
  const trustLegend = () => `<div class="legend no-print-x"><b style="color:var(--night)">تصنيف البيانات:</b>
      <span class="tag c-ok">اشتراط رسمي / مؤكد</span><span class="tag c-as">افتراض تصميمي أو مالي</span>
      <span class="tag c-est">تقدير محسوب</span><span class="tag c-ver">يحتاج تحقق أو اعتماد</span><span class="tag c-inv">مطلوب من المستثمر</span></div>`;

  /* ================= عرض ورقة Excel حية ================= */
  function renderSheet(code, opt = {}) {
    const si = SHEET[code]; if (si === undefined) return '';
    const dim = M.dims[si];
    const r1 = opt.from || 3, r2 = Math.min(opt.to || dim.r, dim.r);
    let c1 = opt.c1 || 2, c2 = opt.c2 || dim.c;
    const merged = {}, covered = new Set();
    dim.merged.forEach(m => {
      const mm = m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/); if (!mm) return;
      const a1 = G.colNum(mm[1]), b1 = +mm[2], a2 = G.colNum(mm[3]), b2 = +mm[4];
      merged[mm[1] + b1] = { cs: Math.min(a2, c2) - a1 + 1, rs: b2 - b1 + 1 };
      for (let r = b1; r <= b2; r++) for (let c = a1; c <= a2; c++) if (r !== b1 || c !== a1) covered.add(G.colStr(c) + r);
    });
    let html = '<div class="tbl-wrap sheet"><table><tbody>', blank = 0;
    for (let r = r1; r <= r2; r++) {
      let row = '', any = false;
      for (let c = c1; c <= c2; c++) {
        const A = G.colStr(c) + r; if (covered.has(A)) continue;
        const k = si + '!' + A, mg = merged[A];
        const span = mg ? `${mg.cs > 1 ? ` colspan="${mg.cs}"` : ''}${mg.rs > 1 ? ` rowspan="${mg.rs}"` : ''}` : '';
        const inp = M.inp[k], fx = M.fx[k], st = M.st[k] || '';
        let v = E.getKey(k);
        if (v !== null && v !== '' && v !== undefined) any = true;
        if (inp) {
          any = true;
          const fmt = E.fmt(k), mod = k in E.over;
          const pct = /%/.test(fmt), date = /yyyy/.test(fmt);
          let ctl;
          if (inp.list) ctl = `<select data-k="${k}">${(inp.list.includes(String(v)) || v == null ? inp.list : [String(v), ...inp.list]).map(o => `<option${String(v) === o ? ' selected' : ''}>${esc(o)}</option>`).join('')}${v == null || v === '' ? '<option selected value="">—</option>' : ''}</select>`;
          else if (date && isNum(v)) ctl = `<input type="date" data-k="${k}" data-t="date" value="${G.toDate(v).toISOString().slice(0, 10)}">`;
          else if (typeof v === 'number' || v === null) ctl = `<input inputmode="decimal" data-k="${k}" data-t="${pct ? 'pct' : 'num'}" value="${v === null ? '' : esc(pct ? +(v * 100).toPrecision(10) : +v.toPrecision(12))}"${pct ? ' title="بالنسبة المئوية %"' : ''}>`;
          else ctl = `<input data-k="${k}" data-t="text" value="${esc(v)}">`;
          row += `<td class="in${mod ? ' mod' : ''}"${span}>${ctl}${pct ? '' : ''}</td>`;
          continue;
        }
        const isN = typeof v === 'number';
        let txt = esc(fmtCell(v, E.fmt(k)));
        if (G.isErr(v)) { row += `<td class="err"${span}>${txt}</td>`; continue; }
        if (M.url[k]) txt = `<a href="${esc(M.url[k])}" target="_blank" rel="noopener">${txt}</a>`;
        else if (typeof v === 'string' && /^https?:\/\//.test(v)) txt = `<a href="${esc(v)}" target="_blank" rel="noopener">${txt}</a>`;
        const cls = [st, isN ? 'v-n' : '', fx ? 'fx' : ''].filter(Boolean).join(' ');
        row += `<td${cls ? ` class="${cls}"` : ''}${span}${fx ? ` title="${esc(fx)}"` : ''}>${txt}</td>`;
      }
      if (!any) { if (!blank++) html += `<tr class="sp"><td colspan="${c2 - c1 + 1}"></td></tr>`; continue; }
      blank = 0; html += `<tr>${row}</tr>`;
    }
    return html + '</tbody></table></div>';
  }
  /* ورقة كاملة داخل عنصر قابل للطي */
  function sheetBlock(code, opt = {}) {
    const si = SHEET[code]; const title = M.sheets[si]; const desc = M.v[si + '!B2'];
    return `<details class="sx${opt.noprint ? ' noprint' : ''}"${opt.open ? ' open' : ''}><summary>${esc(opt.title || ('الورقة ' + title))}<span class="s">${opt.hint || 'الجدول الكامل من الدراسة – الخلايا الصفراء قابلة للتعديل'}</span></summary>
      <div class="body">${desc && !opt.noDesc ? `<p class="small muted" style="margin-top:0">${esc(desc)}</p>` : ''}${renderSheet(code, opt)}</div></details>`;
  }

  /* صفوف حسب التسمية (للمرونة عند تحديث الدراسة) */
  function findRow(code, label, col = 'B', from = 1) {
    const si = SHEET[code], d = M.dims[si];
    for (let r = from; r <= d.r; r++) { const v = M.v[si + '!' + col + r]; if (typeof v === 'string' && v.trim().startsWith(label)) return r; }
    return null;
  }
  const yrRow = (code, r, cols = YC) => cols.map(c => N(code, c + r));

  /* ================= المؤشرات المالية الرئيسية ================= */
  function core() {
    const L = a => V('17', a);
    return {
      scenario: NM('Scenario'),
      invest: L('L7'), capex: L('L8'), conWorks: NM('CON_SUB'), boq: NM('CAPEX_TOTAL'),
      st1: L('E9'), st5: L('L9'), occ5: L('L10'),
      rev1: L('L11'), rev5: L('L12'), cashExp5: L('L13'), np1: L('L14'), np5: L('L15'), npAvg: L('L16'), ebitda5: L('L17'),
      roi: L('L18'), npv: L('L20'), npvNoTV: L('L21'), irr: L('L22'), irrNoTV: L('L23'), payback: L('L25'),
      be5: L('L26'), beOcc5: L('L27'), be1: L('L28'), feeNeeded: L('L30'),
      capProp: NM('Cap_Proposed'), capTheory: NM('Cap_Theory'), capLand: NM('Cap_Land'), capOff: NM('Cap_Official'), capBus: NM('Cap_Bus'),
      classes: NM('Classrooms_Avail'), kgRooms: NM('KG_Rooms'), primRooms: NM('Prim_Rooms'),
      busReq: NM('Bus_Req'), busDrawn: NM('Bus_Drawn'), parkReq: NM('Parking_Req'), parkProv: NM('Parking_Prov'),
      critOpen: NM('Crit_Open'), discRate: NM('Disc_Rate'),
    };
  }
  const yearly = () => ({
    seats: yrRow('17', 8), students: yrRow('17', 9), occ: yrRow('17', 7), fee: yrRow('17', 10),
    gross: yrRow('17', 11), disc: yrRow('17', 12), exempt: yrRow('17', 13), tuition: yrRow('17', 14), other: yrRow('17', 15), revenue: yrRow('17', 16),
    salTeach: yrRow('17', 17), varEdu: yrRow('17', 18), direct: yrRow('17', 19), gross_profit: yrRow('17', 20), salAdm: yrRow('17', 21), opexFixed: yrRow('17', 22), badDebt: yrRow('17', 23),
    preop: yrRow('17', 24, ['D', ...YC]), ebitda: yrRow('17', 25), dep: yrRow('17', 26), ebit: yrRow('17', 27), ebt: yrRow('17', 28), tax: yrRow('17', 32), zakat: yrRow('17', 33), net: yrRow('17', 34),
    capexY: yrRow('17', 35, ['D', ...YC]), wc: yrRow('17', 37, ['D', ...YC]), fcf: yrRow('17', 38, ['D', ...YC]), cum: yrRow('17', 39, ['D', ...YC]), tv: yrRow('17', 40, ['D', ...YC]), fcfTV: yrRow('17', 41, ['D', ...YC]),
    be: yrRow('17', 44), beOcc: yrRow('17', 45),
  });

  /* ================= الصفحات ================= */
  const PAGES = [
    { id: 'home', t: 'الرئيسية', r: pHome },
    { id: 'summary', t: 'الملخص التنفيذي', r: pSummary },
    { id: 'market', t: 'دراسة السوق والطلب', r: pMarket },
    { id: 'engineering', t: 'الدراسة الفنية والهندسية', r: pEngineering },
    { id: 'capacity', t: 'الطاقة الاستيعابية', r: pCapacity },
    { id: 'capex', t: 'تكاليف الإنشاء والتجهيز', r: pCapex },
    { id: 'opex', t: 'المصاريف التشغيلية', r: pOpex },
    { id: 'revenue', t: 'الإيرادات المتوقعة', r: pRevenue },
    { id: 'cashflow', t: 'التدفقات النقدية', r: pCashflow },
    { id: 'pnl', t: 'الأرباح والخسائر', r: pPnl },
    { id: 'breakeven', t: 'نقطة التعادل', r: pBreakeven },
    { id: 'returns', t: 'العائد على الاستثمار', r: pReturns },
    { id: 'timeline', t: 'الجدول الزمني للتنفيذ', r: pTimeline },
    { id: 'risks', t: 'المخاطر والافتراضات', r: pRisks },
    { id: 'refs', t: 'المراجع والاعتمادات', r: pRefs },
  ];

  const head = (eyebrow, title, desc) => `<div class="ph"><div class="eyebrow">${esc(eyebrow)}</div><h1>${esc(title)}</h1>${desc ? `<p>${desc}</p>` : ''}</div>`;
  const scenarioNote = () => `<span class="pill" style="background:var(--sand);color:#8A5A10">السيناريو النشط: ${esc(NM('Scenario'))}</span>`;
  const chartBox = (id, title, sub, cls = '') => `<div class="card ${cls}"><h3>${esc(title)}</h3>${sub ? `<p class="sub">${sub}</p>` : ''}<div id="${id}"></div></div>`;
  const after = []; // رسوم تُرسم بعد إدراج HTML
  const draw = (id, fn) => after.push(() => { const el = document.getElementById(id); if (el) fn(el); });

  function verdictBox(c) {
    const bad = isNum(c.npv) && c.npv < 0;
    const t = V('01', 'B29');
    return `<div class="verdict ${bad ? 'bad' : 'good'}"><div class="ic">${bad ? '!' : '✓'}</div><div>
      <h4>${bad ? 'الحكم المبدئي: غير مجدٍ مالياً بصيغته الحالية' : 'الحكم المبدئي: مؤشرات إيجابية في السيناريو النشط'} <span class="small muted">(السيناريو ${esc(c.scenario)})</span></h4>
      <p>${esc(t)}</p></div></div>`;
  }

  /* ---------- 1) الرئيسية ---------- */
  function pHome() {
    const c = core(), y = yearly();
    const bus = c.busDrawn < c.busReq;
    draw('ch-home-fin', el => CH.combo(el, { cats: YRS, series: [
      { name: 'الإيرادات', values: y.revenue, color: CH.C.petrol },
      { name: 'المصروفات النقدية', values: YC.map((_, i) => -(y.direct[i] + y.salAdm[i] + y.opexFixed[i] + y.badDebt[i])), color: CH.C.aqua },
      { name: 'صافي الربح', type: 'line', values: y.net, color: CH.C.saffron }], tipFmt: fSAR }));
    draw('ch-home-cum', el => CH.combo(el, { cats: ['التأسيس', ...YRS], series: [
      { name: 'التدفق النقدي الحر', values: y.fcf, color: CH.C.aqua, negColor: '#D9A49B' },
      { name: 'التدفق التراكمي', type: 'line', values: y.cum, color: CH.C.neg }], tipFmt: fSAR }));
    draw('ch-home-cap', el => CH.hbars(el, { items: [
      { label: 'حسب الفصول (التصميمية)', value: c.capTheory, color: CH.C.petrol },
      { label: 'حسب مساحة الأرض', value: c.capLand, color: CH.C.aqua },
      { label: 'مقاعد خطة التشغيل', value: NM('Max_Seats'), color: '#8FBFBF' },
      { label: 'تخدمه الحافلات المرسومة', value: c.capBus, color: CH.C.neg },
      { label: 'الطاقة المقترحة', value: c.capProp, color: CH.C.saffron }], fmt: fInt, labelW: 200 }));
    draw('ch-home-st', el => CH.combo(el, { cats: YRS, series: [
      { name: 'المقاعد المتاحة', values: y.seats, color: '#CFE2E2' },
      { name: 'الطلاب المتوقعون', values: y.students, color: CH.C.petrol },
      { name: 'طلاب التعادل', type: 'line', values: y.be, color: CH.C.neg, dash: true }], fmt: fInt, tipFmt: v => fInt(v) + ' طالب' }));
    return `
    <div class="hero"><div class="eyebrow">دراسة الجدوى الفنية والهندسية والمالية · الإصدار v2</div>
      <h1>مدارس سُلّم النمو <span class="en" style="font-weight:400;opacity:.8;font-size:.6em">Growth Ladder Schools</span></h1>
      <p>مدرسة أهلية بالمنهج الوطني في الرياض: تمهيدي وابتدائي (الصفوف 1–6)، بقسمين منفصلين للبنين والبنات، على أرض مملوكة للمستثمر بمساحة تقديرية ${fInt(NM('Land_Area'))} م²، ومبنى جديد من ${fInt(NM('Floors'))} أدوار مع قبو مواقف.</p>
      <div class="meta"><span class="pill">نرتقي بالعلم، وننمو بالمستقبل</span><span class="pill">السيناريو النشط: ${esc(c.scenario)}</span><span class="pill">${fInt(c.classes)} فصلاً</span><span class="pill">افتتاح متوقع ${esc(fmtCell(NM('Opening_Date'), 'mm/yyyy'))}</span></div></div>
    ${verdictBox(c)}
    <div style="height:16px"></div>
    <div class="grid g4">
      ${kpi({ label: 'إجمالي تكلفة المشروع', value: numSpan(fM(c.invest)), unit: 'ر.س', foot: 'إنشاء وتجهيز + ما قبل التشغيل + رأس مال عامل (بدون الأرض)', key: K('17', 'L7') })}
      ${kpi({ label: 'تكاليف الإنشاء والتجهيز', value: numSpan(fM(c.capex)), unit: 'ر.س', foot: `أعمال الإنشاء والتنفيذ ${numSpan(fM(c.conWorks))} · مسعّر بعروض فعلية ${numSpan(fPct(NM('Quote_Pct'), 0))}`, key: K('17', 'L8') })}
      ${kpi({ label: 'الطاقة الاستيعابية المقترحة', value: numSpan(fInt(c.capProp)), unit: 'طالب', foot: `تصميمية ${numSpan(fInt(c.capTheory))} · حسب الأرض ${numSpan(fInt(c.capLand))} · المعتمدة: ${c.capOff === null || c.capOff === '' ? 'لم تصدر' : fInt(c.capOff)}`, key: 'Cap_Proposed', tone: 'acc' })}
      ${kpi({ label: 'عدد الفصول', value: numSpan(fInt(c.classes)), unit: 'فصل', foot: `تمهيدي ${numSpan(fInt(c.kgRooms))} + ابتدائي ${numSpan(fInt(c.primRooms))}`, key: 'Classrooms_Avail' })}
      ${kpi({ label: 'عدد الطلاب المتوقع', value: numSpan(fInt(c.st5)), unit: 'طالب (س5)', foot: `السنة 1: ${numSpan(fInt(c.st1))} طالب · إشغال س5 ${numSpan(fPct(c.occ5, 0))}`, key: K('17', 'L9'), bar: c.occ5 })}
      ${kpi({ label: 'الإيرادات السنوية (السنة 5)', value: numSpan(fM(c.rev5)), unit: 'ر.س', foot: `السنة 1: ${numSpan(fM(c.rev1))}`, key: K('17', 'L12') })}
      ${kpi({ label: 'المصاريف التشغيلية النقدية (س5)', value: numSpan(fM(c.cashExp5)), unit: 'ر.س', foot: 'رواتب + تكاليف تعليمية + مصروفات ثابتة + مخصص التعثر', key: K('17', 'L13') })}
      ${kpi({ label: 'صافي الأرباح (السنة 5)', value: numSpan(fM(c.np5)), unit: 'ر.س', foot: `السنة 1: ${numSpan(fM(c.np1))} · المتوسط ${numSpan(fM(c.npAvg))}`, key: K('17', 'L15'), tone: c.np5 < 0 ? 'neg' : 'pos' })}
      ${kpi({ label: 'نقطة التعادل (السنة 5)', value: numSpan(fInt(c.be5)), unit: 'طالب', foot: `إشغال ${numSpan(fPct(c.beOcc5, 0))} · أو رسوم ≈ ${numSpan(fSAR(c.feeNeeded))} للطالب`, key: K('17', 'L26'), tone: c.be5 > c.st5 ? 'neg' : 'pos' })}
      ${kpi({ label: 'العائد على الاستثمار', value: numSpan(fPct(c.roi)), unit: 'ROI', foot: `NPV (${fPct(c.discRate, 0)}): ${numSpan(fM(c.npv))} · IRR: ${isNum(c.irr) ? numSpan(fPct(c.irr)) : esc(c.irr)}`, key: K('17', 'L18'), tone: c.roi < 0 ? 'neg' : 'pos' })}
      ${kpi({ label: 'فترة استرداد رأس المال', value: isNum(c.payback) ? numSpan(nf1.format(c.payback)) : `<span style="font-size:20px">${esc(c.payback)}</span>`, unit: isNum(c.payback) ? 'سنة' : '', foot: 'من التدفق النقدي التراكمي دون القيمة المتبقية', key: K('17', 'L25'), tone: isNum(c.payback) ? 'pos' : 'neg' })}
      ${kpi({ label: 'البنود الحرجة المفتوحة', value: numSpan(fInt(c.critOpen)), unit: 'بند', foot: bus ? `أهمها عجز مواقف الحافلات: المرسوم ${fInt(c.busDrawn)} من ${fInt(c.busReq)} مطلوبة` : 'ورقة 01أ', key: 'Crit_Open', tone: 'neg' })}
    </div>
    <div style="height:16px"></div>
    <div class="grid g2">
      ${chartBox('ch-home-fin', 'الإيرادات والمصروفات وصافي الربح', 'السنوات الخمس الأولى من التشغيل – ر.س')}
      ${chartBox('ch-home-cum', 'التدفق النقدي الحر والتراكمي', 'يشمل الاستثمار في سنة التأسيس – ر.س')}
      ${chartBox('ch-home-st', 'الطلاب المتوقعون مقابل نقطة التعادل', 'الفجوة بين الخط المتقطع والأعمدة الداكنة هي عدد الطلاب الناقص للتعادل المحاسبي')}
      ${chartBox('ch-home-cap', 'قيود الطاقة الاستيعابية', 'الطاقة المقترحة = أدنى القيود والخطة؛ قيد الحافلات غير مطبق حالياً (عجز غير محلول)')}
    </div>
    <div style="height:16px"></div>
    <div class="grid g2">
      <div class="card"><h3>مقارنة السيناريوهات</h3><p class="sub">تُحسب السيناريوهات الثلاثة بالتوازي بنفس منطق النموذج المالي (ورقة 18)</p>${scenarioTable()}</div>
      <div class="card"><h3>ما يجب حسمه قبل القرار</h3><p class="sub">من ورقة «01أ الملاحظات الحرجة» – ${fInt(c.critOpen)} بنداً مفتوحاً</p>${critList(6)}<p style="margin:10px 0 0"><a href="#risks">عرض كل البنود الحرجة ←</a></p></div>
    </div>`;
  }

  function scenarioTable() {
    const mn = v => isNum(v) ? nf2.format(v / 1e6) : esc(v);
    const rows = [[20, mn, 'إجمالي الاستثمار (مليون ر.س)'], [21, fInt, 'عدد الطلاب – س5'], [23, mn, 'الإيرادات – س5 (مليون ر.س)'], [26, mn, 'صافي الربح – س5 (مليون ر.س)'], [28, fPct, 'العائد على الاستثمار ROI'], [29, fInt, 'طلاب التعادل – س5'], [31, mn, 'NPV (مليون ر.س)'], [32, v => isNum(v) ? fPct(v) : esc(v), 'IRR']];
    const act = NM('Scenario');
    return `<div class="tbl-wrap"><table class="t"><thead><tr><th>المؤشر</th>${['متحفظ', 'متوسط', 'متفائل'].map(s => `<th class="n"${s === act ? ' style="background:var(--saffron);color:var(--night)"' : ''}>${s}</th>`).join('')}</tr></thead><tbody>
      ${rows.map(([r, f, l]) => `<tr><td>${l}</td>${['C', 'D', 'E'].map(cc => { const v = V('18', cc + r); return `<td class="n${isNum(v) && v < 0 ? ' neg' : ''}">${f(v)}</td>`; }).join('')}</tr>`).join('')}
    </tbody></table></div>`;
  }
  function critRows() {
    const si = SHEET['01أ'], out = [];
    for (let r = 5; r <= M.dims[si].r; r++) { const n = V('01أ', 'B' + r), t = V('01أ', 'C' + r); if (isNum(n) && t) out.push({ n, t, now: V('01أ', 'D' + r), imp: V('01أ', 'E' + r), doc: V('01أ', 'F' + r), who: V('01أ', 'G' + r), own: V('01أ', 'H' + r), st: V('01أ', 'I' + r), r }); }
    return out;
  }
  function critList(max) {
    return '<ol class="list">' + critRows().filter(x => x.st !== 'مغلق').slice(0, max).map(x => `<li><b>${esc(x.t)}</b> <span class="muted small">– ${esc(x.now)}</span></li>`).join('') + '</ol>';
  }

  /* ---------- 2) الملخص التنفيذي ---------- */
  function pSummary() {
    const si = SHEET['01'], items = [];
    for (let r = 3; r <= M.dims[si].r; r += 2) { const h = V('01', 'B' + r), t = V('01', 'B' + (r + 1)); if (h && t) items.push([h, t]); }
    return head('القسم 2', 'الملخص التنفيذي', 'نص الملخص مرتبط بالنموذج ويتحدث تلقائياً مع السيناريو والمدخلات.') + verdictBox(core()) + `<div style="height:16px"></div>
      <div class="grid g2">${items.map(([h, t]) => `<div class="card"><h3>${esc(h)}</h3><p style="margin:0">${esc(t)}</p></div>`).join('')}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('02', { title: 'التعريف بالمشروع (ورقة 02)', open: true })}${sheetBlock('00', { title: 'لوحة المستثمر الهندسية (ورقة 00)', to: 42 })}${sheetBlock('16', { title: 'لوحة الجدوى المالية (ورقة 16)' })}</div>`;
  }

  /* ---------- 3) السوق ---------- */
  function pMarket() {
    const si = SHEET['25'], comp = [];
    for (let r = 5; r <= M.dims[si].r; r++) { const nm = V('25', 'B' + r), g1 = V('25', 'H' + r), kg = V('25', 'G' + r); if (nm && (isNum(g1) || isNum(kg))) comp.push({ nm: String(nm), g1, kg }); }
    const ourFee = yearly().fee[0];
    draw('ch-mkt', el => CH.hbars(el, { items: [...comp.filter(c => isNum(c.g1)).map(c => ({ label: c.nm.length > 34 ? c.nm.slice(0, 33) + '…' : c.nm, value: c.g1, color: CH.C.aqua })), { label: 'سُلّم النمو – متوسط الرسوم (س1)', value: ourFee, color: CH.C.saffron }].sort((a, b) => b.value - a.value), fmt: fInt, tipFmt: fSAR, labelW: 250 }));
    return head('القسم 3', 'دراسة السوق والطلب', 'تحليل المنافسين والرسوم في الرياض. <b>تحليل مؤقت على مستوى المدينة</b> لأن الحي لم يُحدد بعد؛ يُحدَّث بعد تحديد موقع الأرض.') +
      `<div class="note" style="margin-bottom:16px">بيانات المنافسين من مصادر عامة (مواقع المدارس وأدلة المدارس) بتاريخ الاطلاع المذكور، وموثوقيتها موضحة في الجدول. الرسوم المعلنة قد تختلف عن الفعلية بعد الخصومات.</div>
      <div class="grid g2">${chartBox('ch-mkt', 'رسوم الصف الأول لدى المنافسين مقابل متوسط رسوم المشروع', 'ر.س سنوياً – متوسط رسوم المشروع من ورقة الإيرادات (قبل الخصم، السيناريو النشط)', 'span2')}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('25', { title: 'السوق والمنافسون (ورقة 25)', open: true })}${sheetBlock('26', { title: 'خطة التسويق (ورقة 26)' })}${sheetBlock('27', { title: 'الهوية التجارية (ورقة 27)' })}</div>`;
  }

  /* ---------- 4) الهندسية ---------- */
  const DRAW = [['A-01_site_plan', 'A-01 المخطط العام للموقع'], ['A-02_ground', 'A-02 الدور الأرضي'], ['A-03_first', 'A-03 الدور الأول'], ['A-04_second', 'A-04 الدور الثاني'], ['A-05_basement', 'A-05 القبو'], ['A-06_circulation', 'A-06 الحركة والإخلاء'], ['A-07_section', 'A-07 القطاع']];
  let engTab = '04';
  function pEngineering() {
    const tabs = [['04', 'الأرض والموقع'], ['05', 'الاشتراطات النظامية'], ['06', 'البرنامج الوظيفي'], ['07', 'التصور المعماري'], ['08', 'الإنشائية'], ['09', 'الكهروميكانيكية'], ['10', 'السلامة والوصول'], ['30', 'مراجعة المكتب الهندسي']];
    const c = core();
    const changed = changedCount() > 0;
    return head('القسم 4', 'الدراسة الفنية والهندسية', 'الأرض والاشتراطات والبرنامج الوظيفي والأنظمة والسلامة، مع اللوحات التصورية v2 المطابقة لملف Excel. <b>تصورات أولية وليست مخططات تنفيذية أو معتمدة.</b>') +
      (changed ? `<div class="note" style="margin-bottom:14px">⚠ عدّلت مدخلات في هذه الجلسة. اللوحات مرسومة على المدخلات الافتراضية (أرض 50×50م، 3 أدوار…) ولا تتحدث تلقائياً؛ الجداول والأرقام هي التي تتحدث.</div>` : '') +
      `<div class="grid g4">
        ${kpi({ label: 'مساحة الأرض', value: numSpan(fInt(NM('Land_Area'))), unit: 'م²', foot: classTag('تقديرية – تحتاج صك وكروكي'), key: 'Land_Area' })}
        ${kpi({ label: 'داخل الارتدادات / البصمة', value: numSpan(fInt(NM('Envelope_Area'))) + ' / ' + numSpan(fInt(NM('Footprint'))), unit: 'م²', foot: `نسبة البناء ${numSpan(fPct(NM('Coverage'), 0))} ${classTag('يحتاج تحقق')}`, key: 'Footprint' })}
        ${kpi({ label: 'المسطحات فوق الأرض + القبو', value: numSpan(fInt(NM('GFA_Above'))) + ' + ' + numSpan(fInt(NM('Basement_Area'))), unit: 'م²', foot: `${fInt(NM('Floors'))} أدوار · البرنامج ${numSpan(fInt(NM('GFA_Req')))} م²`, key: 'GFA_Total' })}
        ${kpi({ label: 'ملاءمة الأرض', value: `<span style="font-size:16px">${esc(NM('Site_Verdict'))}</span>`, key: 'Site_Verdict', tone: /غير/.test(NM('Site_Verdict')) ? 'neg' : 'pos' })}
        ${kpi({ label: 'مواقف السيارات', value: numSpan(fInt(c.parkProv)) + ' / ' + numSpan(fInt(c.parkReq)), unit: 'متوفر / مطلوب', foot: classTag('اشتراط رسمي') + ' جدول المواقف 1447هـ', key: 'Parking_Prov', tone: c.parkProv >= c.parkReq ? 'pos' : 'neg' })}
        ${kpi({ label: 'مواقف الحافلات', value: numSpan(fInt(c.busDrawn)) + ' / ' + numSpan(fInt(c.busReq)), unit: 'مرسوم / مطلوب', foot: c.busDrawn < c.busReq ? '<span class="tag c-neg">عجز غير محلول</span> يمنع المطابقة' : 'مطابق', key: 'Bus_Drawn', tone: c.busDrawn < c.busReq ? 'neg' : 'pos' })}
        ${kpi({ label: 'التكييف / الكهرباء', value: numSpan(fInt(NM('Cooling_TR'))) + ' / ' + numSpan(fInt(NM('Elec_kVA'))), unit: 'طن تبريد / ك.ف.أ', foot: `محول ${fInt(NM('Transformer_kVA'))} ك.ف.أ ${classTag('تقدير أولي')}`, key: 'Cooling_TR' })}
        ${kpi({ label: 'خزان الحريق / الرشاشات', value: numSpan(fInt(NM('Fire_Tank'))), unit: 'م³', foot: `الرشاشات: ${esc(NM('Sprinkler_Req'))} ${classTag('أساس كود – يحتاج مطابقة SBC')}`, key: 'Fire_Tank' })}
      </div>
      <div style="height:16px"></div>
      <div class="card"><h3>اللوحات التصورية v2</h3><p class="sub">7 لوحات مطابقة لملف Excel v2 (ورقة 07 – 14 بنداً متطابقاً). اضغط على اللوحة للتكبير.</p>
        <div class="drawings">${DRAW.map(([f, t]) => `<figure data-img="assets/drawings/${f}.webp"><img loading="lazy" src="assets/drawings/${f}_t.webp" alt="${esc(t)}"><figcaption>${esc(t)}</figcaption></figure>`).join('')}</div></div>
      <div style="height:16px"></div>
      <div class="card"><div class="tabs no-print">${tabs.map(([k, l]) => `<button data-engtab="${k}" class="${k === engTab ? 'on' : ''}">${esc(l)}</button>`).join('')}</div>
      <h3>${esc(sname(engTab))}</h3><p class="sub">${esc(M.v[SHEET[engTab] + '!B2'] || '')}</p>${renderSheet(engTab)}</div>
      <div class="print-only">${tabs.filter(([k]) => k !== engTab).map(([k]) => `<div class="card" style="margin-top:12px"><h3>${esc(sname(k))}</h3>${renderSheet(k)}</div>`).join('')}</div>`;
  }

  /* ---------- 5) الطاقة ---------- */
  function pCapacity() {
    const c = core(), y = yearly();
    draw('ch-cap', el => CH.hbars(el, { items: [
      { label: '(1) حسب الفصول المتاحة', value: c.capTheory, color: CH.C.petrol }, { label: '(2) حسب مساحة الأرض', value: c.capLand, color: CH.C.aqua },
      { label: 'مقاعد الخطة (س5)', value: NM('Max_Seats'), color: '#8FBFBF' }, { label: 'تخدمه الحافلات المرسومة', value: c.capBus, color: CH.C.neg },
      { label: 'الطاقة المقترحة', value: c.capProp, color: CH.C.saffron }], fmt: fInt, labelW: 200 }));
    draw('ch-cap-y', el => CH.combo(el, { cats: YRS, stacked: true, series: [
      { name: 'مقاعد التمهيدي', values: yrRow('23', 26), color: CH.C.saffron }, { name: 'مقاعد الابتدائي', values: yrRow('23', 27), color: CH.C.petrol },
      { name: 'الطلاب المتوقعون', type: 'line', values: y.students, color: CH.C.night }], fmt: fInt, tipFmt: v => fInt(v) }));
    return head('القسم 5', 'الطاقة الاستيعابية', 'الطاقة النظرية والمقترحة والمعتمدة رسمياً معروضة منفصلة. <b>الطاقة الرسمية تحددها وزارة التعليم في الترخيص ولم تصدر بعد.</b>') +
      `<div class="grid g4">
        ${kpi({ label: 'الطاقة التصميمية', value: numSpan(fInt(c.capTheory)), unit: 'طالب', foot: 'الفصول × سعة الفصل ' + classTag('افتراض'), key: 'Cap_Theory' })}
        ${kpi({ label: 'الطاقة حسب الأرض', value: numSpan(fInt(c.capLand)), unit: 'طالب', foot: `${nf2.format(N('11', 'C7'))} م²/طالب (${esc(NM('Land_Basis'))}) ${classTag('اشتراط رسمي')}`, key: 'Cap_Land' })}
        ${kpi({ label: 'الطاقة المقترحة', value: numSpan(fInt(c.capProp)), unit: 'طالب', foot: 'أدنى القيود والخطة', key: 'Cap_Proposed', tone: 'acc' })}
        ${kpi({ label: 'الطاقة المعتمدة رسمياً', value: c.capOff === null || c.capOff === '' ? '<span style="font-size:20px">لم تصدر</span>' : numSpan(fInt(c.capOff)), foot: classTag('يحتاج اعتماد وزارة التعليم'), key: 'Cap_Official' })}
      </div><div style="height:16px"></div>
      <div class="note bad" style="margin-bottom:16px"><b>قيد الحافلات غير المحلول:</b> المواقف المرسومة (${fInt(c.busDrawn)} من ${fInt(c.busReq)}) تخدم نحو ${fInt(c.capBus)} طالباً فقط. القيد غير مطبق على الطاقة المقترحة حالياً (خيار «تطبيق قيد الحافلات» = ${esc(NM('Bus_Cap_Apply'))} في الافتراضات)، ويحتاج قراراً معتمداً.</div>
      <div class="grid g2">${chartBox('ch-cap', 'القيود الحاكمة للطاقة', 'طالب')}${chartBox('ch-cap-y', 'خطة الافتتاح التدريجي', 'المقاعد المتاحة حسب المرحلة والطلاب المتوقعون في السيناريو النشط')}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('11', { title: 'حاسبة الطاقة الاستيعابية (ورقة 11)', open: true })}${sheetBlock('23', { title: 'النموذج التعليمي وخطة الافتتاح (ورقة 23)' })}</div>`;
  }

  /* ---------- 6) تكاليف الإنشاء ---------- */
  function pCapex() {
    const groups = []; for (let r = 94; r <= 109; r++) { const l = V('13', 'B' + r), v = N('13', 'C' + r); if (l && isNum(v)) groups.push({ label: l, value: v }); }
    const cats = []; for (let r = 114; r <= 117; r++) cats.push({ label: V('13', 'B' + r), value: N('13', 'C' + r) });
    draw('ch-capex-g', el => CH.hbars(el, { items: groups.slice().sort((a, b) => b.value - a.value).map((g, i) => ({ ...g, color: i < 3 ? CH.C.petrol : CH.C.aqua })), fmt: v => nf2.format(v / 1e6) + ' م', tipFmt: fSAR, labelW: 210 }));
    draw('ch-capex-c', el => CH.donut(el, { items: cats, fmt: fSAR, center: [nf2.format(NM('CAPEX_TOTAL') / 1e6), 'مليون ر.س'] }));
    return head('القسم 6', 'تكاليف الإنشاء والتجهيز', 'جدول كميات تقديري مرتبط بالمسطحات وحسابات الإنشاء والأنظمة. <b>الأسعار تقديرات أولية للسوق السعودي 2026 وليست عروض أسعار.</b> أدخل سعر العرض الفعلي في عمود «سعر العرض» ليحل محل التقدير تلقائياً.') +
      `<div class="grid g4">
        ${kpi({ label: 'أعمال الإنشاء والتنفيذ', value: numSpan(fM(NM('CON_SUB'))), unit: 'ر.س', foot: `${numSpan(fInt(N('13', 'C119')))} ر.س/م² مبني شامل القبو`, key: 'CON_SUB' })}
        ${kpi({ label: 'إجمالي جدول الكميات', value: numSpan(fM(NM('CAPEX_TOTAL'))), unit: 'ر.س', foot: 'شامل الدراسات والتجهيز وما قبل التشغيل والاحتياطي', key: 'CAPEX_TOTAL' })}
        ${kpi({ label: 'تكلفة التأسيس لكل مقعد', value: numSpan(fInt(N('13', 'C120'))), unit: 'ر.س', foot: 'حسب الطاقة المقترحة', key: K('13', 'C120') })}
        ${kpi({ label: 'المسعّر بعروض فعلية', value: numSpan(fPct(NM('Quote_Pct'), 0)), foot: '<span class="tag c-ver">التكلفة تقديرية بالكامل</span>', key: 'Quote_Pct', tone: 'neg' })}
      </div><div style="height:16px"></div>
      <div class="grid g2">${chartBox('ch-capex-g', 'التكاليف حسب مجموعات الأعمال', 'مليون ر.س')}${chartBox('ch-capex-c', 'التوزيع حسب فئة الإهلاك', 'يغذي النموذج المالي')}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('13', { title: 'جدول الكميات والتكاليف BOQ (ورقة 13)', open: true })}${sheetBlock('08', { title: 'الدراسة الإنشائية (ورقة 08)' })}${sheetBlock('09', { title: 'الأنظمة الكهروميكانيكية (ورقة 09)' })}</div>`;
  }

  /* ---------- 7) المصاريف التشغيلية ---------- */
  function pOpex() {
    const y = yearly();
    const neg = a => a.map(v => -v);
    draw('ch-opex', el => CH.combo(el, { cats: YRS, stacked: true, series: [
      { name: 'رواتب الكادر التعليمي', values: neg(y.salTeach), color: CH.C.petrol }, { name: 'رواتب الإدارة والإشراف والمساندة', values: neg(y.salAdm), color: CH.C.aqua },
      { name: 'تكاليف تعليمية متغيرة', values: neg(y.varEdu), color: CH.C.saffron }, { name: 'مصروفات تشغيلية ثابتة', values: neg(y.opexFixed), color: CH.C.dark },
      { name: 'مخصص التعثر', values: neg(y.badDebt), color: '#9AAAAE' }], tipFmt: fSAR }));
    const hc = yrRow('24', 29, ['E', 'F', 'G', 'H', 'I']);
    const sal5 = -(y.salTeach[4] + y.salAdm[4]), tot5 = -(y.direct[4] + y.salAdm[4] + y.opexFixed[4] + y.badDebt[4]);
    return head('القسم 7', 'المصاريف التشغيلية', 'الرواتب من ورقة الموارد البشرية (مرتبطة بعدد الفصول المفتوحة)، والمصروفات الثابتة والمتغيرة من ورقة OPEX. <b>الرواتب افتراضات استرشادية وليست عروض سوق.</b>') +
      `<div class="grid g4">
        ${kpi({ label: 'المصروفات النقدية – السنة 5', value: numSpan(fM(tot5)), unit: 'ر.س', key: K('17', 'L13') })}
        ${kpi({ label: 'الرواتب والمزايا – السنة 5', value: numSpan(fM(sal5)), unit: 'ر.س', foot: `${numSpan(fPct(sal5 / tot5, 0))} من المصروفات النقدية`, key: K('17', 'I17') })}
        ${kpi({ label: 'عدد الموظفين', value: numSpan(fInt(hc[0])) + ' ← ' + numSpan(fInt(hc[4])), unit: 'س1 ← س5', key: K('24', 'I29') })}
        ${kpi({ label: 'نسبة الطلاب إلى المعلمين', value: esc(fmtCell(V('24', 'I36'), E.fmt(K('24', 'I36')))), unit: 'س5', key: K('24', 'I36') })}
      </div><div style="height:16px"></div>
      <div class="grid g2">${chartBox('ch-opex', 'هيكل المصروفات التشغيلية النقدية', 'ر.س سنوياً – السيناريو النشط', 'span2')}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('22', { title: 'المصروفات التشغيلية OPEX (ورقة 22)', open: true })}${sheetBlock('24', { title: 'الهيكل التنظيمي والموارد البشرية (ورقة 24)' })}</div>`;
  }

  /* ---------- 8) الإيرادات ---------- */
  function pRevenue() {
    const y = yearly();
    draw('ch-rev', el => CH.combo(el, { cats: YRS, stacked: true, series: [
      { name: 'صافي الرسوم الدراسية', values: y.tuition, color: CH.C.petrol }, { name: 'إيرادات أخرى', values: y.other, color: CH.C.saffron },
      { name: 'الرسوم قبل الخصم', type: 'line', values: y.gross, color: CH.C.aqua, dash: true }], tipFmt: fSAR }));
    draw('ch-rev-st', el => CH.combo(el, { cats: YRS, series: [{ name: 'متوسط الرسوم الإجمالية للطالب', values: y.fee, color: CH.C.aqua }], fmt: fInt, tipFmt: fSAR, legend: false }));
    return head('القسم 8', 'الإيرادات المتوقعة', 'الرسوم السنوية لكل صف وقسم قابلة للتعديل (قبل ضريبة القيمة المضافة). <b>الرسوم المقترحة ليست معتمدة نظامياً.</b> الإيراد = الطلاب × متوسط الرسوم − الخصومات والإعفاءات + الإيرادات الأخرى.') +
      `<div class="grid g4">
        ${kpi({ label: 'الإيرادات – السنة 1', value: numSpan(fM(y.revenue[0])), unit: 'ر.س', foot: `${numSpan(fInt(y.students[0]))} طالب`, key: K('17', 'E16') })}
        ${kpi({ label: 'الإيرادات – السنة 5', value: numSpan(fM(y.revenue[4])), unit: 'ر.س', foot: `${numSpan(fInt(y.students[4]))} طالب`, key: K('17', 'I16') })}
        ${kpi({ label: 'متوسط الرسوم الإجمالية (س5)', value: numSpan(fInt(y.fee[4])), unit: 'ر.س/طالب', foot: `زيادة سنوية ${numSpan(fPct(NM('Fee_Incr'), 0))} ${classTag('افتراض')}`, key: K('17', 'I10') })}
        ${kpi({ label: 'الخصومات + الإعفاءات', value: numSpan(fPct(NM('Disc_Pct') + NM('Exempt_Pct'), 0)), unit: 'من الرسوم', foot: `تحصيل ${numSpan(fPct(NM('Collect_Pct'), 0))}`, key: 'Disc_Pct' })}
      </div><div style="height:16px"></div>
      <div class="grid g2">${chartBox('ch-rev', 'مكونات الإيرادات', 'ر.س – السيناريو النشط')}${chartBox('ch-rev-st', 'متوسط الرسوم الإجمالية للطالب', 'ر.س – يشمل معامل رسوم السيناريو')}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('21', { title: 'الإيرادات والرسوم الدراسية (ورقة 21)', open: true })}</div>`;
  }

  /* جدول سنوي من ورقة 17 */
  function finTable(rows, withY0 = false) {
    const cols = withY0 ? ['D', ...YC] : YC, hd = withY0 ? ['التأسيس', ...YRS] : YRS;
    return `<div class="tbl-wrap"><table class="t"><thead><tr><th>البند</th>${hd.map(h => `<th class="n">${h}</th>`).join('')}</tr></thead><tbody>
      ${rows.map(r => { const tot = M.st[K('17', 'B' + r)] === 'tot' || /tot/.test(M.st[K('17', 'B' + r)] || ''); return `<tr class="${tot ? 'tot' : ''}"><td>${esc(V('17', 'B' + r))}${howBtn(K('17', 'I' + r))}</td>${cols.map(c => { const v = V('17', c + r); const fmt = E.fmt(K('17', c + r)); return `<td class="n${isNum(v) && v < 0 ? ' neg' : ''}">${v === null ? '' : esc(fmtCell(v, fmt === 'General' ? '#,##0;(#,##0);"-"' : fmt))}</td>`; }).join('')}</tr>`; }).join('')}
    </tbody></table></div>`;
  }

  /* ---------- 9) التدفقات النقدية ---------- */
  function pCashflow() {
    const y = yearly(), c = core();
    draw('ch-cf', el => CH.combo(el, { cats: ['التأسيس', ...YRS], series: [
      { name: 'الإنفاق الرأسمالي وما قبل التشغيل ورأس المال العامل', values: ['D', ...YC].map((cc, i) => (y.capexY[i] || 0) + (y.preop[i] || 0) + (y.wc[i] || 0)), color: '#D9A49B' },
      { name: 'التدفق النقدي الحر', values: y.fcf, color: CH.C.petrol },
      { name: 'التدفق التراكمي', type: 'line', values: y.cum, color: CH.C.neg }], tipFmt: fSAR }));
    return head('القسم 9', 'التدفقات النقدية', 'التدفق النقدي منفصل عن قائمة الدخل: الإهلاك غير نقدي، وتكلفة الإنشاء تظهر في التدفق النقدي فقط. <b>المشروع ممول ذاتياً بلا فوائد</b> (يُعدَّل إن وُجد تمويل).') +
      `<div class="grid g4">
        ${kpi({ label: 'الاستثمار في سنة التأسيس', value: numSpan(fM(-y.fcf[0])), unit: 'ر.س', key: K('17', 'D38') })}
        ${kpi({ label: 'التدفق الحر – السنة 5', value: numSpan(fM(y.fcf[5])), unit: 'ر.س', key: K('17', 'I38'), tone: y.fcf[5] < 0 ? 'neg' : 'pos' })}
        ${kpi({ label: 'التدفق التراكمي – نهاية س5', value: numSpan(fM(y.cum[5])), unit: 'ر.س', key: K('17', 'I39'), tone: y.cum[5] < 0 ? 'neg' : 'pos' })}
        ${kpi({ label: 'القيمة المتبقية نهاية س5', value: numSpan(fM(y.tv[5])), unit: 'ر.س', foot: `الطريقة: ${esc(NM('TV_Method'))} ${classTag('افتراض')}`, key: K('17', 'I40') })}
      </div><div style="height:16px"></div>
      ${chartBox('ch-cf', 'التدفقات النقدية السنوية والتراكمية', 'ر.س')}
      <div style="height:16px"></div><div class="card"><h3>قائمة التدفق النقدي</h3><p class="sub">من ورقة 17 – السيناريو النشط (${esc(c.scenario)})</p>${finTable([25, 24, 32, 33, 35, 36, 37, 38, 39, 40, 41], true)}</div>`;
  }

  /* ---------- 10) الأرباح والخسائر ---------- */
  function pPnl() {
    const y = yearly();
    draw('ch-pnl', el => CH.combo(el, { cats: YRS, series: [
      { name: 'الإيرادات', values: y.revenue, color: CH.C.petrol }, { name: 'EBITDA', values: y.ebitda, color: CH.C.aqua, negColor: '#D9A49B' },
      { name: 'صافي الربح', type: 'line', values: y.net, color: CH.C.saffron }], tipFmt: fSAR }));
    draw('ch-margin', el => CH.combo(el, { cats: YRS, series: [{ name: 'هامش صافي الربح', values: y.net.map((v, i) => y.revenue[i] ? v / y.revenue[i] : 0), color: CH.C.neg, negColor: CH.C.neg }], fmt: v => Math.round(v * 100) + '%', tipFmt: v => fPct(v), legend: false }));
    return head('القسم 10', 'الأرباح والخسائر', 'قائمة الدخل لخمس سنوات للسيناريو النشط، شاملة الإهلاك والزكاة وضريبة الدخل على حصة الشريك الأجنبي مع ترحيل الخسائر. <b>الضريبة والزكاة تقريبية وتحتاج مستشاراً ضريبياً.</b>') +
      `<div class="grid g2">${chartBox('ch-pnl', 'الإيرادات وEBITDA وصافي الربح', 'ر.س')}${chartBox('ch-margin', 'هامش صافي الربح', '% من الإيرادات')}</div>
      <div style="height:16px"></div><div class="card"><h3>قائمة الدخل</h3><p class="sub">ورقة 17 – السيناريو النشط (${esc(NM('Scenario'))})</p>${finTable([9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 32, 33, 34])}</div>
      <div style="height:16px"></div>${sheetBlock('17', { title: 'النموذج المالي الكامل (ورقة 17)', noprint: true })}`;
  }

  /* ---------- 11) التعادل ---------- */
  function pBreakeven() {
    const y = yearly(), c = core();
    draw('ch-be', el => CH.combo(el, { cats: YRS, series: [
      { name: 'الطلاب المتوقعون', values: y.students, color: CH.C.petrol }, { name: 'طلاب التعادل', values: y.be, color: '#D9A49B' },
      { name: 'المقاعد المتاحة', type: 'line', values: y.seats, color: CH.C.aqua, dash: true }], fmt: fInt, tipFmt: v => fInt(v) + ' طالب' }));
    return head('القسم 11', 'نقطة التعادل', 'التعادل المحاسبي: عدد الطلاب الذي يغطي الرواتب والمصروفات الثابتة والإهلاك بمتوسط صافي الرسوم للطالب.') +
      `<div class="grid g4">
        ${kpi({ label: 'طلاب التعادل – السنة 5', value: numSpan(fInt(c.be5)), unit: 'طالب', foot: `المتوقع ${numSpan(fInt(c.st5))} · الفارق ${numSpan(fInt(c.st5 - c.be5))}`, key: K('17', 'L26'), tone: c.be5 > c.st5 ? 'neg' : 'pos' })}
        ${kpi({ label: 'إشغال التعادل – السنة 5', value: numSpan(fPct(c.beOcc5, 0)), foot: `مقابل إشغال متوقع ${numSpan(fPct(c.occ5, 0))}`, key: K('17', 'L27'), bar: c.beOcc5 })}
        ${kpi({ label: 'طلاب التعادل – السنة 1', value: numSpan(fInt(c.be1)), unit: 'طالب', foot: `المقاعد المتاحة ${numSpan(fInt(y.seats[0]))}`, key: K('17', 'L28') })}
        ${kpi({ label: 'الرسوم المطلوبة للتعادل (س5)', value: numSpan(fInt(c.feeNeeded)), unit: 'ر.س/طالب', foot: `مقابل ${numSpan(fInt(y.fee[4]))} في النموذج`, key: K('17', 'L30'), tone: c.feeNeeded > y.fee[4] ? 'neg' : 'pos' })}
      </div><div style="height:16px"></div>
      ${chartBox('ch-be', 'الطلاب المتوقعون مقابل طلاب التعادل', 'طالب – السيناريو النشط')}
      <div style="height:16px"></div><div class="stack">${sheetBlock('20', { title: 'التعادل والعائد (ورقة 20)', open: true })}</div>`;
  }

  /* ---------- 12) العائد ---------- */
  function pReturns() {
    const c = core();
    const sens = []; for (let r = 16; r <= 21; r++) sens.push({ label: V('19', 'B' + r), value: N('19', 'J' + r) });
    draw('ch-sens', el => CH.hbars(el, { items: sens.sort((a, b) => a.value - b.value), fmt: v => nf2.format(v / 1e6) + ' م', tipFmt: fSAR, labelW: 200 }));
    const sc = ['C', 'D', 'E'].map(cc => N('18', cc + 31));
    draw('ch-scn', el => CH.combo(el, { cats: ['متحفظ', 'متوسط', 'متفائل'], series: [{ name: 'NPV', values: sc, color: CH.C.petrol, negColor: CH.C.neg }, { name: 'صافي الربح س5', values: ['C', 'D', 'E'].map(cc => N('18', cc + 26)), color: CH.C.saffron, negColor: '#F0C98A' }], tipFmt: fSAR }));
    return head('القسم 12', 'العائد على الاستثمار', `مؤشرات العائد للسيناريو النشط بمعدل خصم ${fPct(c.discRate, 0)}، ومقارنة السيناريوهات، وتحليل الحساسية لكل متغير منفرداً.`) +
      `<div class="grid g4">
        ${kpi({ label: 'العائد على الاستثمار ROI', value: numSpan(fPct(c.roi)), foot: 'متوسط صافي الربح السنوي ÷ إجمالي الاستثمار', key: K('17', 'L18'), tone: c.roi < 0 ? 'neg' : 'pos' })}
        ${kpi({ label: 'صافي القيمة الحالية NPV', value: numSpan(fM(c.npv)), unit: 'ر.س', foot: `مع القيمة المتبقية · بدونها ${numSpan(fM(c.npvNoTV))}`, key: K('17', 'L20'), tone: c.npv < 0 ? 'neg' : 'pos' })}
        ${kpi({ label: 'معدل العائد الداخلي IRR', value: isNum(c.irr) ? numSpan(fPct(c.irr)) : esc(c.irr), foot: `بدون القيمة المتبقية: ${isNum(c.irrNoTV) ? fPct(c.irrNoTV) : esc(c.irrNoTV)}`, key: K('17', 'L22'), tone: !isNum(c.irr) || c.irr < c.discRate ? 'neg' : 'pos' })}
        ${kpi({ label: 'فترة الاسترداد', value: isNum(c.payback) ? numSpan(nf1.format(c.payback)) : `<span style="font-size:20px">${esc(c.payback)}</span>`, unit: isNum(c.payback) ? 'سنة' : '', key: K('17', 'L25'), tone: isNum(c.payback) ? 'pos' : 'neg' })}
      </div><div style="height:16px"></div>
      <div class="grid g2">${chartBox('ch-scn', 'السيناريوهات الثلاثة', 'NPV وصافي ربح السنة 5 – ر.س')}${chartBox('ch-sens', 'حساسية NPV لكل صدمة', 'التغير في NPV عن الحالة الأساسية – عدّل حجم الصدمات في جدول ورقة 19 أدناه')}</div>
      <div style="height:16px"></div><div class="card"><h3>مقارنة السيناريوهات</h3><p class="sub">عدّل معاملات كل سيناريو من الجدول الكامل أدناه</p>${scenarioTable()}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('18', { title: 'معاملات السيناريوهات والمقارنة (ورقة 18)', to: 35, open: true })}${sheetBlock('19', { title: 'تحليل الحساسية (ورقة 19)', to: 23, open: true })}${sheetBlock('18', { title: 'حسابات السيناريوهات التفصيلية (ورقة 18)', from: 36, hint: 'كتل الحساب الكاملة لكل سيناريو', noprint: true })}${sheetBlock('19', { title: 'حسابات الحساسية التفصيلية (ورقة 19)', from: 24, hint: 'تعاد حسابات النموذج كاملة لكل صدمة', noprint: true })}</div>`;
  }

  /* ---------- 13) الجدول الزمني ---------- */
  function pTimeline() {
    const si = SHEET['12'], acts = [];
    for (let r = 5; r <= 29; r++) { const n = V('12', 'B' + r), t = V('12', 'C' + r), s = N('12', 'H' + r), e = N('12', 'I' + r); if (isNum(n) && t && isNum(s)) acts.push({ n, t, type: V('12', 'D' + r), s, e: isNum(e) ? e : s + N('12', 'G' + r) - 1, d: N('12', 'G' + r) }); }
    const maxM = Math.max(...acts.map(a => a.e), 1);
    const base = { 'دراسات': CH.C.aqua, 'اعتمادات': CH.C.saffron, 'تصميم': '#8FBFBF', 'تعاقد': CH.C.grey, 'تنفيذ': CH.C.petrol, 'تجهيز': '#3F7F8C', 'تسليم': CH.C.dark };
    const col = {}; acts.forEach((a, i) => { if (!col[a.type]) col[a.type] = base[a.type] || CH.PAL[Object.keys(col).length % CH.PAL.length]; });
    const pct = m => (m - 1) / maxM * 100;
    const gantt = `<div class="gantt"><table><tbody>${acts.map(a => `<tr><td style="width:30px" class="muted num">${fInt(a.n)}</td><td style="width:30%;white-space:normal">${esc(a.t)}</td><td class="small muted" style="width:64px">${esc(a.type)}</td>
      <td class="bar-cell">${Array.from({ length: Math.ceil(maxM / 3) }, (_, i) => `<span class="grid-m" style="inset-inline-start:${pct(i * 3 + 1)}%"></span>`).join('')}<span class="bar" style="inset-inline-start:${pct(a.s)}%;width:${Math.max(1.2, a.d / maxM * 100)}%;background:${col[a.type] || CH.C.grey}" data-tip="${esc(a.t)}<br>الشهر ${fInt(a.s)} → ${fInt(a.e)} (${fInt(a.d)} شهر)" title="الشهر ${a.s} – ${a.e}"></span></td></tr>`).join('')}
      <tr><td></td><td></td><td></td><td class="bar-cell small muted" style="height:20px">${Array.from({ length: Math.ceil(maxM / 3) + 1 }, (_, i) => `<span style="position:absolute;inset-inline-start:${pct(i * 3 + 1)}%;top:2px;font-size:10.5px">ش${i * 3 + 1}</span>`).join('')}</td></tr></tbody></table>
      <div class="lgd chart" style="margin-top:6px"><div class="lgd">${Object.entries(col).map(([k, c]) => `<span><i style="background:${c}"></i>${k}</span>`).join('')}</div></div></div>`;
    return head('القسم 13', 'الجدول الزمني للتنفيذ', 'شهر البداية يُحسب تلقائياً من النشاط السابق وفترة التداخل. <b>المدد تقديرية وليست ضماناً لمواعيد الجهات.</b>') +
      `<div class="grid g4">
        ${kpi({ label: 'مدة المشروع الكلية', value: numSpan(fInt(NM('Exec_Months'))), unit: 'شهر', key: 'Exec_Months' })}
        ${kpi({ label: 'مدة الإنشاء', value: numSpan(fInt(NM('Constr_Months'))), unit: 'شهر', key: 'Constr_Months' })}
        ${kpi({ label: 'تاريخ البدء', value: numSpan(fmtCell(NM('Project_Start'), 'yyyy-mm-dd')), key: 'Project_Start', foot: classTag('افتراض') })}
        ${kpi({ label: 'الافتتاح المتوقع', value: numSpan(fmtCell(NM('Opening_Date'), 'mm/yyyy')), key: 'Opening_Date', foot: 'بداية العام الدراسي التالي لاكتمال التنفيذ', tone: 'acc' })}
      </div><div style="height:16px"></div>
      <div class="card"><h3>مخطط جانت</h3><p class="sub">الأشهر من تاريخ بدء المشروع – عدّل المدد والتداخل من الجدول أدناه</p>${gantt}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('12', { title: 'خطة التنفيذ (ورقة 12)', c2: 9, hint: 'المدد والتداخل قابلة للتعديل' })}${sheetBlock('14', { title: 'خريطة التراخيص والاعتمادات (ورقة 14)' })}</div>`;
  }

  /* ---------- 14) المخاطر والافتراضات ---------- */
  let riskTab = 'crit';
  function pRisks() {
    const risks = []; for (let r = 5; r <= 19; r++) { const t = V('15', 'C' + r), p = N('15', 'E' + r), i = N('15', 'F' + r); if (t && isNum(p)) risks.push({ n: V('15', 'B' + r), t, p, i, cat: V('15', 'D' + r), lvl: V('15', 'H' + r) }); }
    const cell = (p, i) => risks.filter(x => x.p === p && x.i === i);
    const heat = `<div class="tbl-wrap"><table class="t" style="table-layout:fixed;text-align:center"><thead><tr><th style="width:90px">الاحتمالية ↓ / الأثر ←</th>${[1, 2, 3, 4, 5].map(i => `<th class="n" style="text-align:center">${i}</th>`).join('')}</tr></thead><tbody>
      ${[5, 4, 3, 2, 1].map(p => `<tr><th style="background:var(--mist);color:var(--night)">${p}</th>${[1, 2, 3, 4, 5].map(i => { const s = p * i, xs = cell(p, i); const bg = s >= 15 ? '#F3C9C1' : s >= 8 ? '#F8E2B8' : '#DDEEE7'; return `<td style="background:${bg};text-align:center;font-weight:800">${xs.map(x => `<span title="${esc(x.t)}" data-tip="${esc(x.t)}">${fInt(x.n)}</span>`).join(' · ')}</td>`; }).join('')}</tr>`).join('')}</tbody></table></div>`;
    const crit = critRows();
    const asmTable = assumptionsTable();
    const tabs = [['crit', `الملاحظات الحرجة (${crit.length})`], ['asm', 'الافتراضات والمدخلات'], ['risk', `سجل المخاطر (${risks.length})`], ['eng', 'مراجعة المكتب الهندسي']];
    let body = '';
    if (riskTab === 'crit') body = `<div class="tbl-wrap"><table class="t"><thead><tr><th>#</th><th>البند الحرج</th><th>الوضع الحالي</th><th>الأثر إن لم يُعالج</th><th>المستند / القرار المطلوب</th><th>الجهة</th><th>المسؤول</th><th>الحالة</th></tr></thead><tbody>${crit.map(x => `<tr><td class="n">${fInt(x.n)}</td><td><b>${esc(x.t)}</b></td><td class="small">${esc(x.now)}</td><td class="small">${esc(x.imp)}</td><td class="small">${esc(x.doc)}</td><td class="small">${esc(x.who)}</td><td class="small">${esc(x.own)}</td><td class="in"><select class="crit-st" data-k="${K('01أ', 'I' + x.r)}">${(M.inp[K('01أ', 'I' + x.r)]?.list || ['مفتوح', 'قيد المعالجة', 'مغلق']).map(o => `<option${o === x.st ? ' selected' : ''}>${esc(o)}</option>`).join('')}</select></td></tr>`).join('')}</tbody></table></div>`;
    else if (riskTab === 'asm') body = asmTable;
    else if (riskTab === 'risk') body = `<div class="grid g2" style="margin-bottom:14px"><div>${heat}</div><div class="note info small">الدرجة = الاحتمالية × الأثر (1–5). الأرقام في المصفوفة هي أرقام المخاطر في السجل. عدّل الاحتمالية والأثر من الجدول أدناه.</div></div>${renderSheet('15')}`;
    else body = renderSheet('30');
    return head('القسم 14', 'المخاطر والافتراضات', 'البنود الحرجة قبل الاعتماد، وكل الافتراضات مصنفة حسب مصدرها، وسجل المخاطر، وقائمة مراجعة المكتب الهندسي. <b>لا تُعتمد الدراسة قبل إغلاق البنود الحرجة.</b>') +
      `<div class="card" style="margin-bottom:16px">${trustLegend()}</div>
      <div class="card"><div class="tabs no-print">${tabs.map(([k, l]) => `<button data-risktab="${k}" class="${k === riskTab ? 'on' : ''}">${esc(l)}</button>`).join('')}</div>${body}</div>
      <div class="print-only">${riskTab !== 'asm' ? `<div class="card" style="margin-top:12px"><h3>الافتراضات</h3>${asmTable}</div>` : ''}</div>`;
  }

  /* جدول الافتراضات (ورقة 03) مع التصنيف */
  function assumptionsTable() {
    const si = SHEET['03'], d = M.dims[si]; let html = '<div class="tbl-wrap sheet"><table><tbody>';
    for (let r = 5; r <= d.r; r++) {
      const b = V('03', 'B' + r); if (!b) continue;
      const k = K('03', 'C' + r);
      if (!M.inp[k] && !M.fx[k] && M.v[k] === undefined) { html += `<tr><td class="sec" colspan="5">${esc(b)}</td></tr>`; continue; }
      const cellHTML = renderSheet('03', { from: r, to: r, c1: 3, c2: 3 }).replace(/^.*?<tr>|<\/tr>.*$/gs, '');
      html += `<tr><td class="b" style="max-width:340px">${esc(b)}</td>${cellHTML}<td class="mut">${esc(V('03', 'D' + r) || '')}</td><td>${classTag(V('03', 'E' + r))}</td><td class="mut small" style="max-width:420px">${esc(V('03', 'F' + r) || '')}</td></tr>`;
    }
    return html + '</tbody></table></div>';
  }

  /* ---------- 15) المراجع ---------- */
  function pRefs() {
    return head('القسم 15', 'المراجع والاعتمادات', 'المرجع النظامي المعتمد، والمصادر، والمستندات والاعتمادات المطلوبة، وتقرير المراجعة وسجل التعديلات. <b>لا يدّعي الموقع أي اعتماد رسمي أو هندسي غير صادر.</b>') +
      `<div class="grid g2">
        <div class="card"><h3>المرجع النظامي المعتمد</h3><ul class="list">
          <li><b>اشتراطات المباني التعليمية الخاصة 1447هـ</b> – وزارة البلديات والإسكان، قرار وزاري رقم 4600425374/1 بتاريخ 1447-06-13 (ساري). <a href="https://momah.gov.sa/sites/default/files/2025-12/alnskht%20alm%60tmdt%20lashtratat%20almbany%20alt%60lymyt.pdf" target="_blank" rel="noopener">الوثيقة الرسمية</a></li>
          <li><b>كود البناء السعودي 2024</b> – إلزامي منذ 30 يونيو 2025.</li>
          <li>أي قيمة مبنية على كود دولي (IBC/IFC/NFPA/ASHRAE/IPC/ASCE 7) معلَّمة «أساس كود – يحتاج مطابقة SBC 2024 واعتماد الجهة».</li>
          <li>لا تُعتمد الأخبار الصحفية أو المسودات مصدراً لأي متطلب نهائي.</li></ul></div>
        <div class="card"><h3>حالة الاعتمادات</h3>
          <table class="t"><tbody>
          <tr><td>اكتمال المستندات (ورقة 28)</td><td class="n">${fPct(NM('Docs_Pct'), 0)}</td></tr>
          <tr><td>اكتمال التراخيص والاعتمادات (ورقة 14)</td><td class="n">${fPct(NM('Lic_Pct'), 0)}</td></tr>
          <tr><td>البنود المسعّرة بعروض فعلية</td><td class="n">${fPct(NM('Quote_Pct'), 0)}</td></tr>
          <tr><td>البنود الحرجة المفتوحة (ورقة 01أ)</td><td class="n">${fInt(NM('Crit_Open'))}</td></tr>
          <tr><td>الطاقة المعتمدة من وزارة التعليم</td><td>${NM('Cap_Official') === null || NM('Cap_Official') === '' ? 'لم تصدر' : fInt(NM('Cap_Official'))}</td></tr>
          <tr><td>الاشتراطات: مطابق / غير مطابق / يحتاج تحقق</td><td class="n">${fInt(NM('Reg_OK'))} / ${fInt(NM('Reg_NO'))} / ${fInt(NM('Reg_TBC'))}</td></tr>
          </tbody></table><p class="small muted" style="margin-bottom:0">الدراسة تقديرية لأغراض القرار المبدئي، وتحتاج اعتماد مكتب هندسي ومختصين (ورقة 30) وإغلاق البنود الحرجة (ورقة 01أ).</p></div>
      </div>
      <div style="height:16px"></div>
      <div class="card"><h3>تقرير المراجعة والتدقيق قبل v2</h3><p class="sub">وافق عليه المستثمر، ونُفذت تعديلاته في v2 (31 تعديلاً موثقاً في ورقة 31)</p>${reviewSummary()}
        <p style="margin:12px 0 0"><a class="btn" href="docs/تقرير_المراجعة_والتدقيق_قبل_v2.md" target="_blank">النص الكامل لتقرير المراجعة</a></p></div>
      <div style="height:16px"></div>
      <div class="stack">${sheetBlock('31', { title: 'سجل التعديلات v2 (ورقة 31)', open: true })}${sheetBlock('29', { title: 'المصادر والمراجع (ورقة 29)' })}${sheetBlock('28', { title: 'المستندات المطلوبة (ورقة 28)' })}${sheetBlock('05', { title: 'سجل الاشتراطات النظامية (ورقة 05)' })}</div>`;
  }
  function reviewSummary() {
    return `<div class="grid g2"><div><h4 style="margin:0 0 6px;color:var(--neg)">الأخطاء الجوهرية التي صُححت</h4><ul class="list small">
      <li>نصيب الطالب من الأرض: تمهيدي 5.00 وابتدائي 4.40 م² (متوسط 4.70) بدل 4.0/4.5 → الطاقة حسب الأرض 531 بدل 566.</li>
      <li>المواقف وفق الجدول الرسمي: 38 سيارة و8 حافلات على 22 فصلاً (بدل 60 و7).</li>
      <li>خزان الحريق: 113 م³ (مواقف القبو خطر عادي – NFPA 13) بدل 31 م³.</li>
      <li>منحدر القبو: 28م في الشريط الشرقي بدل 9م غير قابل للتنفيذ.</li>
      <li>توحيد مساحات Excel مع المخططات (14 بنداً متطابقاً) وتصحيح ازدواجات جدول الكميات وإضافة البنود الناقصة.</li>
      <li>تصحيح فحص نقاط التجمع، وشرط الرشاشات على منطقة الحريق، والمستودع 100 م² حداً أدنى.</li></ul></div>
      <div><h4 style="margin:0 0 6px;color:var(--petrol)">القرارات التي وافق عليها المستثمر</h4><ol class="list small">
      <li>اعتماد 4.70 م² للطالب (متوسط المراحل)، و4.80 «مجمع تعليمي» بديلاً متحفظاً.</li>
      <li>حساب المواقف والحافلات على فصول التصميم (22).</li>
      <li>إبقاء القبو بعد إعادة تحجيمه (32 م²/موقف) ونقل المنحدر للشريط الشرقي.</li>
      <li>أسوار مصمتة بارتفاع 2.40م لقسم البنات.</li></ol>
      <div class="note bad small" style="margin-top:10px"><b>بند مفتوح أهم:</b> عجز مواقف الحافلات (${fInt(NM('Bus_Drawn'))} من ${fInt(NM('Bus_Req'))}). الحلول: موافقة على مواقف خارج الموقع أو جدولة، قبو حافلات أعمق، ضم أرض مجاورة، أو خفض الفصول.</div></div></div>`;
  }

  /* ================= لوحة المدخلات الرئيسية ================= */
  function drawerHTML() {
    const si = SHEET['03'], d = M.dims[si]; let html = '', open = false;
    for (let r = 5; r <= d.r; r++) {
      const b = V('03', 'B' + r); if (!b) continue; const k = K('03', 'C' + r);
      if (!M.inp[k] && !M.fx[k] && M.v[k] === undefined) { html += (open ? '</div>' : '') + `<div class="grp"><h4>${esc(b)}</h4>`; open = true; continue; }
      if (!M.inp[k]) continue;
      const v = E.getKey(k), fmt = E.fmt(k), pct = /%/.test(fmt), date = /yyyy/.test(fmt), mod = k in E.over, inp = M.inp[k];
      let ctl;
      if (inp.list) ctl = `<select data-k="${k}">${inp.list.map(o => `<option${String(v) === o ? ' selected' : ''}>${esc(o)}</option>`).join('')}${v == null || v === '' ? '<option selected value="">—</option>' : ''}</select>`;
      else if (date && isNum(v)) ctl = `<input type="date" data-k="${k}" data-t="date" value="${G.toDate(v).toISOString().slice(0, 10)}">`;
      else if (typeof v === 'number' || v === null) ctl = `<input inputmode="decimal" data-k="${k}" data-t="${pct ? 'pct' : 'num'}" value="${v === null ? '' : (pct ? +(v * 100).toPrecision(10) : +v.toPrecision(12))}">`;
      else ctl = `<input data-k="${k}" data-t="text" value="${esc(v)}">`;
      const unit = V('03', 'D' + r);
      html += `<div class="field${mod ? ' mod' : ''}"><label>${esc(b)}</label><div>${ctl}<div class="u">${pct ? '%' : esc(unit || '')}${mod ? ` · الأصل: ${esc(fmtCell(M.v[k], fmt))}` : ''}</div></div>
        <div class="meta">${classTag(V('03', 'E' + r))}<span>${esc(V('03', 'F' + r) || '')}</span></div></div>`;
    }
    return html + (open ? '</div>' : '');
  }
  function openDrawer() {
    $('#drawer .body').innerHTML = `<div class="note info small">عدّل أي افتراض وستتحدث كل الصفحات والرسوم فوراً. القيم الأصلية من دراسة v2 تبقى محفوظة ويمكن استعادتها بزر «القيم الأصلية». بقية المدخلات (الرسوم، الرواتب، أسعار البنود، معاملات السيناريوهات…) قابلة للتعديل داخل جداول كل قسم (الخلايا الصفراء).</div>
      <div class="grp"><h4>السيناريو النشط</h4><div class="field"><label>اختر السيناريو</label><div><select data-k="${E.resolve('Scenario')}">${['متحفظ', 'متوسط', 'متفائل'].map(s => `<option${s === NM('Scenario') ? ' selected' : ''}>${s}</option>`).join('')}</select></div></div></div>` + drawerHTML();
    $('#drawer').classList.add('open'); $('#scrim').classList.add('show');
  }
  function closeDrawer() { $('#drawer').classList.remove('open'); $('#scrim').classList.remove('show'); $('#side').classList.remove('open'); }

  /* ================= التوجيه والعرض ================= */
  let current = 'home';
  function nav() {
    $('#nav').innerHTML = PAGES.map((p, i) => `<a href="#${p.id}" data-id="${p.id}" class="${p.id === current ? 'on' : ''}"><span class="n">${i + 1}</span>${esc(p.t)}</a>`).join('');
  }
  function render(id, keepScroll) {
    const p = PAGES.find(x => x.id === id) || PAGES[0]; current = p.id;
    const y = scrollY; after.length = 0;
    $('#page').innerHTML = p.r();
    after.forEach(f => f());
    $('#crumb').innerHTML = `<button class="btn menu no-print" id="btn-menu" aria-label="القائمة">☰</button>${esc(p.t)}`;
    $$('#nav a').forEach(a => a.classList.toggle('on', a.dataset.id === p.id));
    $$('.seg button').forEach(b => b.classList.toggle('on', b.dataset.s === NM('Scenario')));
    document.title = p.t + ' – مدارس سُلّم النمو | دراسة الجدوى';
    if (keepScroll) scrollTo(0, y); else scrollTo(0, 0);
  }
  function rerender() { render(current, true); if ($('#drawer').classList.contains('open')) { const b = $('#drawer .body'), st = b.scrollTop; openDrawer(); b.scrollTop = st; } updateChangedBadge(); }

  /* تقرير كامل للطباعة / PDF */
  function printAll() {
    const keep = current; after.length = 0;
    const html = `<div class="cover"><div class="top"><img src="assets/brand/logo_arabic_color.svg" alt="سُلّم النمو"></div>
      <div><div style="color:var(--aqua);font-weight:700">مدرسة أهلية – الرياض</div><h1>الدراسة الفنية والهندسية والجدوى المالية</h1>
      <p style="font-size:13pt;color:var(--grey);margin:0">مدارس سُلّم النمو | <span class="en">Growth Ladder Schools</span><br>نرتقي بالعلم، وننمو بالمستقبل</p></div>
      <div><div class="band"></div><p class="small muted">الإصدار v2 · السيناريو: ${esc(NM('Scenario'))} · ${changedCount() ? `<b>يتضمن ${changedCount()} مدخلاً معدلاً عن الدراسة الأصلية</b>` : 'بالمدخلات الأصلية للدراسة'} · طُبع في ${esc(new Date().toLocaleDateString('ar-SA-u-nu-latn-ca-gregory'))}<br>دراسة تقديرية لأغراض القرار المبدئي؛ ليست اعتماداً رسمياً أو هندسياً.</p></div></div>` +
      PAGES.map(p => `<section class="psec">${p.r()}</section>`).join('');
    $('#page').innerHTML = html; after.forEach(f => f());
    $$('#page details').forEach(d => d.open = true);
    setTimeout(() => { window.print(); }, 400);
    const restore = () => { window.removeEventListener('afterprint', restore); render(keep, false); };
    window.addEventListener('afterprint', restore);
  }

  /* ================= الأحداث ================= */
  function parseInput(el) {
    const t = el.dataset.t, raw = el.value.trim();
    if (el.tagName === 'SELECT') return raw === '' ? null : raw;
    if (t === 'date') { if (!raw) return null; const [yy, mm, dd] = raw.split('-').map(Number); return (Date.UTC(yy, mm - 1, dd) - Date.UTC(1899, 11, 30)) / 86400000; }
    if (t === 'text') return raw;
    if (raw === '') return null;
    const n = Number(raw.replace(/,/g, '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace('٫', '.').replace('%', ''));
    if (!isFinite(n)) { toast('أدخل رقماً صحيحاً'); return undefined; }
    return t === 'pct' ? n / 100 : n;
  }
  document.addEventListener('change', e => {
    const el = e.target; if (!el.dataset || !el.dataset.k) return;
    if (el.dataset.a === 'import') return;
    const v = parseInput(el); if (v === undefined) return;
    setInput(el.dataset.k, v);
  });
  document.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.dataset && e.target.dataset.k && e.target.tagName === 'INPUT') e.target.blur(); if (e.key === 'Escape') { closeModal(); closeDrawer(); $('#lightbox').classList.remove('show'); } });
  document.addEventListener('click', e => {
    const t = e.target.closest('[data-how],[data-a],[data-engtab],[data-risktab],[data-s],figure[data-img],#btn-menu,#nav a');
    if (!t) { if (e.target.id === 'modal') closeModal(); return; }
    if (t.matches('#nav a')) { $('#side').classList.remove('open'); $('#scrim').classList.remove('show'); return; }
    if (t.id === 'btn-menu') { $('#side').classList.add('open'); $('#scrim').classList.add('show'); return; }
    if (t.dataset.how) return showHow(t.dataset.how, t.dataset.extra);
    if (t.dataset.engtab) { engTab = t.dataset.engtab; return render(current, true); }
    if (t.dataset.risktab) { riskTab = t.dataset.risktab; return render(current, true); }
    if (t.dataset.s) { setInput(E.resolve('Scenario'), t.dataset.s); return; }
    if (t.dataset.img) { $('#lightbox img').src = t.dataset.img; $('#lightbox').classList.add('show'); return; }
    const a = t.dataset.a;
    if (a === 'close') closeModal();
    else if (a === 'inputs') openDrawer();
    else if (a === 'close-drawer') closeDrawer();
    else if (a === 'reset') resetAll();
    else if (a === 'snaps') openSnaps();
    else if (a === 'snap-save') snapSave();
    else if (a === 'export') exportJSON();
    else if (a === 'print') printAll();
    else if (a === 'print-page') window.print();
    else if (a === 'help') showHelp();
    else if (a === 'snap-load') { const s = snaps()[+t.dataset.i]; E.over = filterOver(s.over); E.cache.clear(); saveState(); rerender(); closeModal(); toast('استُعيدت النسخة: ' + s.name); }
    else if (a === 'snap-del') { const l = snaps(); l.splice(+t.dataset.i, 1); localStorage.setItem(SNAPS, JSON.stringify(l)); openSnaps(); }
  });
  document.addEventListener('change', e => { if (e.target.dataset && e.target.dataset.a === 'import' && e.target.files[0]) importJSON(e.target.files[0]); }, true);
  $('#scrim').addEventListener('click', closeDrawer);
  $('#lightbox').addEventListener('click', e => { if (e.target.tagName !== 'IMG') $('#lightbox').classList.remove('show'); });
  window.addEventListener('hashchange', () => render(location.hash.slice(1) || 'home'));

  function showHelp() {
    modal(`<h3>كيف تستخدم الموقع</h3><ol class="list">
      <li>ابدأ من <b>الرئيسية</b>: الحكم المبدئي وأهم المؤشرات والرسوم.</li>
      <li>بدّل <b>السيناريو</b> (متحفظ / متوسط / متفائل) من الشريط العلوي.</li>
      <li>افتح <b>الافتراضات</b> لتعديل المدخلات الرئيسية، أو عدّل الخلايا الصفراء داخل جداول أي قسم. تتحدث كل الأرقام والرسوم فوراً.</li>
      <li>زر <b>؟</b> بجانب أي مؤشر يعرض معادلته الأصلية ومكانها في ملف الدراسة.</li>
      <li><b>حفظ / استعادة</b>: يحفظ التعديلات في هذا المتصفح فقط، ويمكن تصديرها ملفاً لنقلها لجهاز آخر.</li>
      <li><b>القيم الأصلية</b>: يعيد كل المدخلات إلى دراسة v2.</li>
      <li><b>تصدير PDF</b>: يجهز تقريراً كاملاً بكل الأقسام؛ اختر «حفظ كـ PDF» في نافذة الطباعة.</li></ol>
      <div class="note info small">التعديلات تُحفظ تلقائياً محلياً في هذا المتصفح وعلى هذا الجهاز فقط، ولا تتم مزامنتها بين الأجهزة.</div>
      <div style="margin-top:14px;text-align:left"><button class="btn" data-a="close">إغلاق</button></div>`);
  }

  /* ================= التشغيل ================= */
  loadState(); nav(); updateChangedBadge();
  $('#build').textContent = `المصدر: ${M.source} · ${fInt(Object.keys(M.f).length)} معادلة · ${fInt(Object.keys(M.inp).length)} مدخلاً`;
  render(location.hash.slice(1) || 'home');
  if (new URLSearchParams(location.search).has('print')) setTimeout(printAll, 600);
  window.GL = { E, M, core, yearly, render, PAGES }; // للاختبار والتطوير
})();
