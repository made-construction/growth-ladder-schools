/* app.js – واجهة الموقع التفاعلي لدراسة جدوى مدارس سُلّم النمو
 * كل رقم معروض يُقرأ من محرك المعادلات (engine.js) الذي يشغّل معادلات ملف Excel v2 نفسها.
 * لا توجد نتائج محسوبة مكتوبة يدوياً في هذا الملف.
 */
(function () {
  'use strict';
  const M = window.GL_MODEL, G = window.GLEngine, CH = window.GLCharts;
  const E = new G.Engine(M);
  const STORE = 'gl-study-v3-nonfin';     // حالة العمل الحالية (النسخة غير المالية)
  const SNAPS = 'gl-study-v3-nonfin-snapshots'; // النسخ المحفوظة بأسماء

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
  // نوافذ تأكيد وإدخال داخل الصفحة (بدل confirm/prompt التي تمنعها بعض بيئات العرض)
  function resetAll() {
    if (!changedCount()) { toast('المدخلات مطابقة للدراسة الأصلية'); return; }
    modal(`<h3>إعادة القيم الأصلية</h3><p>إعادة جميع المدخلات (${fInt(changedCount())}) إلى القيم الأصلية في دراسة v2؟ لن تُحذف النسخ المحفوظة.</p>
      <div style="display:flex;gap:8px;justify-content:flex-end"><button class="btn" data-a="close">إلغاء</button><button class="btn pri" data-a="reset-yes">نعم، أعد القيم الأصلية</button></div>`);
  }
  function resetYes() { E.reset(); saveState(); closeModal(); rerender(); toast('أُعيدت القيم الأصلية'); }

  const snaps = () => { try { return JSON.parse(localStorage.getItem(SNAPS) || '[]'); } catch (e) { return []; } };
  function snapSave() {
    const el = document.getElementById('snap-name');
    const name = (el && el.value.trim()) || ('نسخة ' + new Date().toLocaleString('ar-SA-u-nu-latn'));
    const list = snaps(); list.unshift({ name, at: new Date().toISOString(), over: E.over });
    try { localStorage.setItem(SNAPS, JSON.stringify(list.slice(0, 30))); toast('حُفظت النسخة في هذا المتصفح'); } catch (e) { toast('تعذّر الحفظ في هذا المتصفح'); }
    openSnaps();
  }
  function openSnaps() {
    const list = snaps();
    modal(`<h3>حفظ التعديلات واستعادتها</h3>
      <div class="note info small">تُحفظ التعديلات <b>محلياً في هذا المتصفح وعلى هذا الجهاز فقط</b> (Local Storage)، ولا تُزامَن تلقائياً بين الأجهزة أو المتصفحات، وتُفقد إذا مُسحت بيانات المتصفح. لنقلها إلى جهاز آخر استخدم «تصدير ملف» ثم «استيراد ملف».</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin:14px 0">
        <input id="snap-name" placeholder="اسم النسخة (مثال: طاقة 450 طالباً)" style="flex:1;min-width:200px;border:1px solid var(--line);border-radius:999px;padding:7px 14px">
        <button class="btn pri" data-a="snap-save">حفظ الوضع الحالي كنسخة</button>
        <button class="btn" data-a="export">تصدير ملف (JSON)</button>
        <label class="btn" style="cursor:pointer">استيراد ملف<input type="file" accept=".json,application/json" data-a="import" hidden></label>
      </div>
      ${list.length ? `<div class="tbl-wrap"><table class="t"><thead><tr><th>النسخة</th><th>التاريخ</th><th>المدخلات المعدلة</th><th></th></tr></thead><tbody>${list.map((s, i) => `<tr><td><b>${esc(s.name)}</b></td><td class="small">${esc(new Date(s.at).toLocaleString('ar-SA-u-nu-latn'))}</td><td class="n">${Object.keys(s.over || {}).length}</td><td style="white-space:nowrap"><button class="btn" data-a="snap-load" data-i="${i}">استعادة</button> <button class="btn" data-a="snap-del" data-i="${i}">حذف</button></td></tr>`).join('')}</tbody></table></div>` : '<p class="muted small">لا توجد نسخ محفوظة بعد.</p>'}
      <div style="margin-top:14px;text-align:left"><button class="btn" data-a="close">إغلاق</button></div>`);
  }
  function exportJSON() {
    const data = { app: 'growth-ladder-study', source: M.source, exported: new Date().toISOString(), over: E.over,
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

  /* ================= المؤشرات الفنية الرئيسية (غير مالية) ================= */
  function core() {
    return {
      capProp: NM('Cap_Proposed'), capTheory: NM('Cap_Theory'), capLand: NM('Cap_Land'), capOff: NM('Cap_Official'), capBus: NM('Cap_Bus'),
      classes: NM('Classrooms_Avail'), kgRooms: NM('KG_Rooms'), primRooms: NM('Prim_Rooms'), maxSeats: NM('Max_Seats'),
      busReq: NM('Bus_Req'), busDrawn: NM('Bus_Drawn'), parkReq: NM('Parking_Req'), parkProv: NM('Parking_Prov'),
      critOpen: NM('Crit_Open'),
    };
  }

  /* ================= الصفحات ================= */
  const PAGES = [
    { id: 'home', t: 'الرئيسية', r: pHome },
    { id: 'summary', t: 'الملخص التنفيذي', r: pSummary },
    { id: 'market', t: 'دراسة السوق والطلب', r: pMarket },
    { id: 'engineering', t: 'الدراسة الفنية والهندسية', r: pEngineering },
    { id: 'property', t: 'معايير المباني التعليمية', r: pProperty },
    { id: 'regulations', t: 'اشتراطات المباني التعليمية', r: pRegulations },
    { id: 'capacity', t: 'الطاقة الاستيعابية', r: pCapacity },
    { id: 'timeline', t: 'الجدول الزمني للتنفيذ', r: pTimeline },
    { id: 'risks', t: 'المخاطر والافتراضات', r: pRisks },
    { id: 'refs', t: 'المراجع والاعتمادات', r: pRefs },
    { id: 'identity', t: 'الهوية المؤسسية والمراجع البصرية', r: pIdentity },
  ];

  let rendering = 'home';
  const head = (eyebrow, title, desc) => `<div class="ph"><div class="eyebrow">${esc(/^القسم \d+$/.test(eyebrow) ? 'القسم ' + (PAGES.findIndex(x => x.id === rendering) + 1) : eyebrow)}</div><h1>${esc(title)}</h1>${desc ? `<p>${desc}</p>` : ''}</div>`;
  const chartBox = (id, title, sub, cls = '') => `<div class="card ${cls}"><h3>${esc(title)}</h3>${sub ? `<p class="sub">${sub}</p>` : ''}<div id="${id}"></div></div>`;
  const after = []; // رسوم تُرسم بعد إدراج HTML
  const draw = (id, fn) => after.push(() => { const el = document.getElementById(id); if (el) fn(el); });

  const refBanner = (what) => `<div class="note" style="margin-bottom:14px"><b>مرجعي:</b> ${what} أُعدّ لتصور مبنى جديد على أرض 50×50م، ويُعرض مرجعاً للبرنامج الوظيفي والمتطلبات. المشروع في مبنى قائم سيتم استئجاره وتجهيزه للاستخدام التعليمي، والمتطلبات النظامية في قسم <a href="#regulations">اشتراطات المباني التعليمية</a>.</div>`;

  const TAGLINE = 'صرح تعليمي يصنع جيلًا من المعرفة والتميّز، ويهيئ أبناءنا وبناتنا لمستقبل واعد.';
  const DESCRIPTION = 'مدارس أهلية في مدينة الرياض، تقدم التعليم بالمنهج الوطني للمرحلتين التمهيدية والابتدائية، ضمن بيئة تعليمية محفزة تجمع بين جودة التعليم وتنمية المهارات.';

  /* ---------- 1) الرئيسية ---------- */
  function pHome() {
    const c = core(), bus = c.busDrawn < c.busReq, P = propsSummary();
    draw('ch-home-cap', el => CH.hbars(el, { items: [
      { label: 'حسب الفصول (التصميمية)', value: c.capTheory, color: CH.C.petrol },
      { label: 'حسب مساحة الأرض', value: c.capLand, color: CH.C.aqua },
      { label: 'مقاعد خطة التشغيل', value: c.maxSeats, color: '#8FBFBF' },
      { label: 'تخدمه الحافلات المرسومة', value: c.capBus, color: CH.C.neg },
      { label: 'الطاقة المقترحة', value: c.capProp, color: CH.C.saffron }], fmt: fInt, labelW: 200 }));
    draw('ch-home-seats', el => CH.combo(el, { cats: YRS, stacked: true, series: [
      { name: 'مقاعد التمهيدي', values: yrRow('23', 26), color: CH.C.saffron }, { name: 'مقاعد الابتدائي', values: yrRow('23', 27), color: CH.C.petrol }], fmt: fInt, tipFmt: v => fInt(v) + ' مقعد' }));
    return `
    <div class="hero">
      <h1>مدارس سُلّم النمو <span class="en" style="font-weight:400;opacity:.8;font-size:.6em">Growth Ladder Schools</span></h1>
      <p style="font-size:17px;color:#fff;font-weight:700;margin-bottom:6px">${esc(TAGLINE)}</p>
      <p>${esc(DESCRIPTION)}</p>
      <div class="meta"><span class="pill">نرتقي بالعلم، وننمو بالمستقبل</span><span class="pill">مجمع تعليمي</span><span class="pill">مبنى قائم مستأجر</span><span class="pill">الرياض</span></div></div>
    <div class="grid g4">
      ${kpi({ label: 'العقارات المقيّمة', value: numSpan(fInt(P.n)), unit: 'عقار', foot: P.n ? `مناسب مبدئياً ${numSpan(fInt(P.ok))} · يحتاج معالجة ${numSpan(fInt(P.treat))} · غير مناسب ${numSpan(fInt(P.fail))} · غير مكتمل ${numSpan(fInt(P.miss))}` : 'لم يُضف أي عقار بعد – <a href="#property">معايير المباني التعليمية</a>', tone: 'acc' })}
      ${kpi({ label: 'الطاقة الاستيعابية المقترحة (المرجعية)', value: numSpan(fInt(c.capProp)), unit: 'طالب', foot: `تصميمية ${numSpan(fInt(c.capTheory))} · حسب الأرض ${numSpan(fInt(c.capLand))} · المعتمدة: ${c.capOff === null || c.capOff === '' ? 'لم تصدر' : fInt(c.capOff)}`, key: 'Cap_Proposed' })}
      ${kpi({ label: 'عدد الفصول (البرنامج المرجعي)', value: numSpan(fInt(c.classes)), unit: 'فصل', foot: `تمهيدي ${numSpan(fInt(c.kgRooms))} + ابتدائي ${numSpan(fInt(c.primRooms))}`, key: 'Classrooms_Avail' })}
      ${kpi({ label: 'المقاعد المخططة (السنة 5)', value: numSpan(fInt(c.maxSeats)), unit: 'مقعد', foot: 'من خطة الافتتاح التدريجي (ورقة 23)', key: 'Max_Seats' })}
      ${kpi({ label: 'مواقف السيارات المطلوبة (المرجعية)', value: numSpan(fInt(c.parkReq)), unit: 'موقف', foot: 'على 22 فصلاً – جدول (2) الفئة (أ)', key: 'Parking_Req' })}
      ${kpi({ label: 'مواقف الحافلات المطلوبة (المرجعية)', value: numSpan(fInt(c.busReq)), unit: 'موقف', foot: bus ? `في التصور السابق: المرسوم ${fInt(c.busDrawn)} فقط` : '', key: 'Bus_Req', tone: bus ? 'neg' : '' })}
      ${kpi({ label: 'البنود الحرجة المفتوحة', value: numSpan(fInt(c.critOpen)), unit: 'بند', foot: 'ورقة 01أ – بعد حذف البندين الماليين', key: 'Crit_Open', tone: 'neg' })}
      ${kpi({ label: 'التصنيف المعتمد', value: '<span style="font-size:20px">مجمع تعليمي</span>', foot: `نصيب الطالب المستخدم ${numSpan(nf2.format(N('11', 'C7')))} م² – جدول (3) الفئة (أ)`, key: 'Land_Basis' })}
    </div>
    <div style="height:16px"></div>
    <div class="grid g2">
      ${chartBox('ch-home-cap', 'قيود الطاقة الاستيعابية (البرنامج المرجعي)', 'أدنى القيود والخطة؛ تُستبدل بقيم العقار المختار بعد تقييمه')}
      ${chartBox('ch-home-seats', 'خطة الافتتاح التدريجي – المقاعد', 'المقاعد المتاحة حسب المرحلة لكل سنة (ورقة 23)')}
    </div>
    <div style="height:16px"></div>
    <div class="grid g2">
      <div class="card"><h3>ما يجب حسمه قبل القرار</h3><p class="sub">من ورقة «01أ الملاحظات الحرجة» – ${fInt(c.critOpen)} بنداً مفتوحاً</p>${critList(6)}<p style="margin:10px 0 0"><a href="#risks">عرض كل البنود الحرجة ←</a></p></div>
      <div class="card"><h3>اشتراطات المباني التعليمية</h3><p class="sub">أبرز المتطلبات النظامية للمبنى في الرياض (الفئة أ)</p><ul class="list small">
        <li>مبنى تعليمي مستقل (4.1 بند 3).</li><li>شارعان أحدهما بعرض 25 م على الأقل، أو ثلاثة شوارع أحدها 20 م (جدول 1).</li><li>الأدوار: الابتدائي حتى أرضي + 3، ورياض الأطفال حتى أرضي + 1 (جدول 4).</li><li>نصيب الطالب من الأرض: رياض الأطفال 5.00 م² والابتدائي 4.40 م² (جدول 3).</li><li>المواقف: موقفان لكل فصل ابتدائي، وموقف لكل 3 فصول تمهيدي، وموقف حافلة لكل 3 فصول (جدول 2).</li></ul>
        <p style="margin:10px 0 0"><a class="btn pri" href="#regulations">اشتراطات المباني التعليمية ←</a></p></div>
    </div>`;
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
    return head('القسم 2', 'الملخص التنفيذي', 'نص الملخص مرتبط بالنموذج ويتحدث تلقائياً مع المدخلات.') + `<div class="hero" style="margin-bottom:16px"><h1>مدارس سُلّم النمو <span class="en" style="font-weight:400;opacity:.8;font-size:.6em">Growth Ladder Schools</span></h1>
        <p style="font-size:17px;color:#fff;font-weight:700;margin-bottom:6px">${esc(TAGLINE)}</p><p>${esc(DESCRIPTION)}</p></div>
      <div class="card"><h3>التعريف بالمشروع</h3><div class="tbl-wrap"><table class="t"><tbody>
        <tr><th>اسم المشروع</th><td>مدارس سُلّم النمو</td></tr><tr><th>المدينة</th><td>الرياض</td></tr><tr><th>نوع المدرسة</th><td>أهلية بمنهج وطني</td></tr>
        <tr><th>المراحل التعليمية</th><td>التمهيدي والابتدائي</td></tr><tr><th>طبيعة العقار</th><td>مبنى قائم سيتم استئجاره وتجهيزه للاستخدام التعليمي</td></tr><tr><th>الجهة المقدمة للدراسة</th><td>شركة ميد للمقاولات</td></tr>
      </tbody></table></div></div><div style="height:16px"></div>
      <div class="grid g2">${items.map(([h, t]) => `<div class="card"><h3>${esc(h)}</h3><p style="margin:0">${esc(t)}</p></div>`).join('')}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('02', { title: 'التعريف بالمشروع (ورقة 02)', open: true })}${sheetBlock('00', { title: 'لوحة المستثمر الهندسية (ورقة 00)', to: 42 })}</div>`;
  }

  /* ---------- 3) السوق ---------- */
  function pMarket() {
    return head('القسم 3', 'دراسة السوق والطلب', 'المدارس المنافسة وتحليل الطلب في الرياض. <b>تحليل مؤقت على مستوى المدينة</b> لأن الحي لم يُحدد بعد؛ يُحدَّث بعد تحديد العقار المختار.') +
      `<div class="note" style="margin-bottom:16px">بيانات المنافسين من مصادر عامة (مواقع المدارس وأدلة المدارس) بتاريخ الاطلاع المذكور، وموثوقيتها موضحة في الجدول. بيانات الرسوم حُذفت من هذه النسخة غير المالية.</div>
      <div class="stack">${sheetBlock('25', { title: 'السوق والمنافسون (ورقة 25)', open: true })}${sheetBlock('26', { title: 'خطة التسويق (ورقة 26)' })}${sheetBlock('27', { title: 'الهوية التجارية (ورقة 27)' })}</div>`;
  }

  /* ---------- 4) الهندسية ---------- */
  const DRAW = [['A-01_site_plan', 'A-01 المخطط العام للموقع'], ['A-02_ground', 'A-02 الدور الأرضي'], ['A-03_first', 'A-03 الدور الأول'], ['A-04_second', 'A-04 الدور الثاني'], ['A-05_basement', 'A-05 القبو'], ['A-06_circulation', 'A-06 الحركة والإخلاء'], ['A-07_section', 'A-07 القطاع']];
  let engTab = '04';
  function pEngineering() {
    const tabs = [['04', 'الأرض والموقع'], ['06', 'البرنامج الوظيفي'], ['07', 'التصور المعماري'], ['08', 'الإنشائية'], ['09', 'الكهروميكانيكية'], ['10', 'السلامة والوصول'], ['30', 'مراجعة المكتب الهندسي']];
    const c = core();
    const changed = changedCount() > 0;
    return head('القسم 4', 'الدراسة الفنية والهندسية', 'الأرض والاشتراطات والبرنامج الوظيفي والأنظمة والسلامة، مع اللوحات التصورية v2 المطابقة لملف Excel. <b>تصورات أولية وليست مخططات تنفيذية أو معتمدة.</b>') + refBanner('هذا القسم (الأرض والبرنامج والتصور المعماري والإنشائي والأنظمة واللوحات)') +
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
    const c = core();
    draw('ch-cap', el => CH.hbars(el, { items: [
      { label: '(1) حسب الفصول المتاحة', value: c.capTheory, color: CH.C.petrol }, { label: '(2) حسب مساحة الأرض', value: c.capLand, color: CH.C.aqua },
      { label: 'مقاعد الخطة (س5)', value: NM('Max_Seats'), color: '#8FBFBF' }, { label: 'تخدمه الحافلات المرسومة', value: c.capBus, color: CH.C.neg },
      { label: 'الطاقة المقترحة', value: c.capProp, color: CH.C.saffron }], fmt: fInt, labelW: 200 }));
    draw('ch-cap-y', el => CH.combo(el, { cats: YRS, stacked: true, series: [
      { name: 'مقاعد التمهيدي', values: yrRow('23', 26), color: CH.C.saffron }, { name: 'مقاعد الابتدائي', values: yrRow('23', 27), color: CH.C.petrol }], fmt: fInt, tipFmt: v => fInt(v) }));
    return head('القسم 5', 'الطاقة الاستيعابية', 'الطاقة النظرية والمقترحة والمعتمدة رسمياً معروضة منفصلة. <b>الطاقة الرسمية تحددها وزارة التعليم في الترخيص ولم تصدر بعد.</b>') + refBanner('حساب الطاقة في هذا القسم') +
      `<div class="note info" style="margin-bottom:14px"><b>تصنيف المشروع في الدراسة: مجمع تعليمي</b> – نصيب الطالب ${nf2.format(N('11', 'C7'))} م² (جدول 3، الفئة أ). وتعرّف الاشتراطات المجمع التعليمي بمرحلتين دراسيتين أو أكثر باستثناء الحضانة ورياض الأطفال (2.1)، ومعدلا المرحلتين منفردتين: رياض الأطفال 5.00 م² والابتدائي 4.40 م². طاقة كل عقار مرشح تُحسب في قسم <a href="#property">معايير المباني التعليمية</a>.</div>
      <div class="grid g4">
        ${kpi({ label: 'الطاقة التصميمية', value: numSpan(fInt(c.capTheory)), unit: 'طالب', foot: 'الفصول × سعة الفصل ' + classTag('افتراض'), key: 'Cap_Theory' })}
        ${kpi({ label: 'الطاقة حسب الأرض', value: numSpan(fInt(c.capLand)), unit: 'طالب', foot: `${nf2.format(N('11', 'C7'))} م²/طالب (${esc(NM('Land_Basis'))}) ${classTag('اشتراط رسمي')}`, key: 'Cap_Land' })}
        ${kpi({ label: 'الطاقة المقترحة', value: numSpan(fInt(c.capProp)), unit: 'طالب', foot: 'أدنى القيود والخطة', key: 'Cap_Proposed', tone: 'acc' })}
        ${kpi({ label: 'الطاقة المعتمدة رسمياً', value: c.capOff === null || c.capOff === '' ? '<span style="font-size:20px">لم تصدر</span>' : numSpan(fInt(c.capOff)), foot: classTag('يحتاج اعتماد وزارة التعليم'), key: 'Cap_Official' })}
      </div><div style="height:16px"></div>
      <div class="note bad" style="margin-bottom:16px"><b>قيد الحافلات غير المحلول:</b> المواقف المرسومة (${fInt(c.busDrawn)} من ${fInt(c.busReq)}) تخدم نحو ${fInt(c.capBus)} طالباً فقط. القيد غير مطبق على الطاقة المقترحة حالياً (خيار «تطبيق قيد الحافلات» = ${esc(NM('Bus_Cap_Apply'))} في الافتراضات)، ويحتاج قراراً معتمداً.</div>
      <div class="grid g2">${chartBox('ch-cap', 'القيود الحاكمة للطاقة', 'طالب')}${chartBox('ch-cap-y', 'خطة الافتتاح التدريجي', 'المقاعد المتاحة حسب المرحلة لكل سنة')}</div>
      <div style="height:16px"></div><div class="stack">${sheetBlock('11', { title: 'حاسبة الطاقة الاستيعابية (ورقة 11)', open: true })}${sheetBlock('23', { title: 'النموذج التعليمي وخطة الافتتاح (ورقة 23)' })}</div>`;
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
    return head('القسم 13', 'الجدول الزمني للتنفيذ', 'شهر البداية يُحسب تلقائياً من النشاط السابق وفترة التداخل. <b>المدد تقديرية وليست ضماناً لمواعيد الجهات.</b>') + refBanner('هذا الجدول الزمني (التصميم والإنشاء ومدته)') +
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
      `${officialRefsCard()}
      <div style="height:16px"></div>
      <div class="grid g2">
        <div class="card"><h3>المرجع النظامي المعتمد</h3><ul class="list">
          <li><b>اشتراطات المباني التعليمية الخاصة 1447هـ</b> – وزارة البلديات والإسكان، قرار وزاري رقم 4600425374/1 بتاريخ 1447-06-13 (ساري). <a href="https://momah.gov.sa/sites/default/files/2025-12/alnskht%20alm%60tmdt%20lashtratat%20almbany%20alt%60lymyt.pdf" target="_blank" rel="noopener">الوثيقة الرسمية</a></li>
          <li><b>كود البناء السعودي 2024</b> – إلزامي منذ 30 يونيو 2025.</li>
          <li>أي قيمة مبنية على كود دولي (IBC/IFC/NFPA/ASHRAE/IPC/ASCE 7) معلَّمة «أساس كود – يحتاج مطابقة SBC 2024 واعتماد الجهة».</li>
          <li>لا تُعتمد الأخبار الصحفية أو المسودات مصدراً لأي متطلب نهائي.</li></ul></div>
        <div class="card"><h3>حالة الاعتمادات</h3>
          <table class="t"><tbody>
          <tr><td>اكتمال المستندات (ورقة 28)</td><td class="n">${fPct(NM('Docs_Pct'), 0)}</td></tr>
          <tr><td>اكتمال التراخيص والاعتمادات (ورقة 14)</td><td class="n">${fPct(NM('Lic_Pct'), 0)}</td></tr>
          <tr><td>البنود الحرجة المفتوحة (ورقة 01أ)</td><td class="n">${fInt(NM('Crit_Open'))}</td></tr>
          <tr><td>الطاقة المعتمدة من وزارة التعليم</td><td>${NM('Cap_Official') === null || NM('Cap_Official') === '' ? 'لم تصدر' : fInt(NM('Cap_Official'))}</td></tr>
          <tr><td>الاشتراطات: مطابق / غير مطابق / يحتاج تحقق</td><td class="n">${fInt(NM('Reg_OK'))} / ${fInt(NM('Reg_NO'))} / ${fInt(NM('Reg_TBC'))}</td></tr>
          </tbody></table><p class="small muted" style="margin-bottom:0">الدراسة تقديرية لأغراض القرار المبدئي، وتحتاج اعتماد مكتب هندسي ومختصين (ورقة 30) وإغلاق البنود الحرجة (ورقة 01أ).</p></div>
      </div>
      <div style="height:16px"></div>
      <div class="card"><h3>تقرير المراجعة والتدقيق قبل v2</h3><p class="sub">وافق عليه المستثمر، ونُفذت تعديلاته في v2 (31 تعديلاً موثقاً في ورقة 31)</p>${reviewSummary()}
        <p style="margin:12px 0 0"><a class="btn" href="docs/تقرير_المراجعة_والتدقيق_قبل_v2.md" target="_blank">النص الكامل لتقرير المراجعة</a></p></div>
      <div style="height:16px"></div>
      <div class="stack">${sheetBlock('31', { title: 'سجل التعديلات v2 (ورقة 31)', open: true })}${sheetBlock('29', { title: 'المصادر والمراجع (ورقة 29)' })}${sheetBlock('28', { title: 'المستندات المطلوبة (ورقة 28)' })}</div>`;
  }
  function reviewSummary() {
    return `<div class="grid g2"><div><h4 style="margin:0 0 6px;color:var(--neg)">الأخطاء الجوهرية التي صُححت</h4><ul class="list small">
      <li>نصيب الطالب من الأرض: تمهيدي 5.00 وابتدائي 4.40 م² (متوسط 4.70) بدل 4.0/4.5 → الطاقة حسب الأرض 531 بدل 566.</li>
      <li>المواقف وفق الجدول الرسمي: 38 سيارة و8 حافلات على 22 فصلاً (بدل 60 و7).</li>
      <li>خزان الحريق: 113 م³ (مواقف القبو خطر عادي – NFPA 13) بدل 31 م³.</li>
      <li>منحدر القبو: 28م في الشريط الشرقي بدل 9م غير قابل للتنفيذ.</li>
      <li>توحيد مساحات Excel مع المخططات (14 بنداً متطابقاً).</li>
      <li>تصحيح فحص نقاط التجمع، وشرط الرشاشات على منطقة الحريق، والمستودع 100 م² حداً أدنى.</li></ul></div>
      <div><h4 style="margin:0 0 6px;color:var(--petrol)">القرارات التي وافق عليها المستثمر</h4><ol class="list small">
      <li>اعتماد 4.70 م² للطالب (متوسط المراحل)، و4.80 «مجمع تعليمي» بديلاً متحفظاً.</li>
      <li>حساب المواقف والحافلات على فصول التصميم (22).</li>
      <li>إبقاء القبو بعد إعادة تحجيمه (32 م²/موقف) ونقل المنحدر للشريط الشرقي.</li>
      <li>أسوار مصمتة بارتفاع 2.40م لقسم البنات.</li></ol>
      <div class="note small" style="margin-top:10px">تحديث نوفمبر 2026: اعتمد المستثمر استئجار مبنى قائم وتصنيف «مجمع تعليمي»، فالقراران 1 و3 يخصان خيار المبنى الجديد السابق.</div>
      <div class="note bad small" style="margin-top:10px"><b>بند مفتوح أهم:</b> عجز مواقف الحافلات (${fInt(NM('Bus_Drawn'))} من ${fInt(NM('Bus_Req'))}). الحلول: موافقة على مواقف خارج الموقع أو جدولة، قبو حافلات أعمق، ضم أرض مجاورة، أو خفض الفصول.</div></div></div>`;
  }

  /* ---------- 16) الهوية المؤسسية والمراجع البصرية (قسم مضاف) ----------
   * كل الصور من حزمة الهوية الرسمية v1.3 «حزمة_هوية_سلم_النمو» كما هي، والمسار الأصلي لكل صورة مذكور تحتها.
   * حسب الحزمة نفسها: المعرّفات (@) مقترحة والحسابات لم تُنشأ بعد، فلا تُنسب أي صورة لحساب منشور. */
  function pIdentity() {
    const PK = 'حزمة_هوية_سلم_النمو (الإصدار 1.3)';
    const D = 'assets/identity/';
    const fig = (n, cap, src, ext = 'webp') => `<figure class="id-fig" data-img="${D}${n}.${ext}"><img loading="lazy" src="${D}${n}${ext === 'webp' ? '_t' : ''}.${ext}" alt="${esc(cap)}"><figcaption><b>${esc(cap)}</b><span class="id-src">المصدر: ${esc(PK)} ← ${esc(src)}</span></figcaption></figure>`;
    const gal = (items, cls = '') => `<div class="id-gal ${cls}">${items.join('')}</div>`;
    const COLORS = [
      ['البترولي', '#0F4C5C', '15, 76, 92', 'اللون الأساسي', 'الشعار، العناوين، الخلفيات الرئيسية', '#fff'],
      ['الزعفراني', '#E9A23B', '233, 162, 59', 'اللون الثانوي', 'نقطة الطالب، الإبراز، الأزرار الثانوية', '#10242B'],
      ['الأكوا', '#5BA4A4', '91, 164, 164', 'اللون المساعد', 'الأنماط، الرسوم، العناصر المساندة', '#fff'],
      ['البترولي الداكن', '#0A3540', '10, 53, 64', 'مساعد داكن', 'الخلفيات الداكنة، التذييلات', '#fff'],
      ['الضبابي', '#F4F7F7', '244, 247, 247', 'لون الخلفيات', 'الخلفية العامة للمطبوعات والرقمي', '#10242B'],
      ['الرملي الفاتح', '#FBF1DE', '251, 241, 222', 'خلفية دافئة', 'مواد التمهيدي والمناسبات', '#10242B'],
      ['الليلي', '#10242B', '16, 36, 43', 'لون النصوص', 'النصوص الأساسية', '#fff'],
      ['الرمادي', '#4E6168', '78, 97, 104', 'نص ثانوي', 'الشروح والتعليقات', '#fff'],
    ];
    const FONTS = [
      ['الشعار المكتوب «سُلّم النمو»', 'ExtraBold', 'محوّل إلى مسارات متجهة؛ لا يُكتب كنص'],
      ['العبارة التعريفية داخل الشعار', 'Regular', '≈ 25% من ارتفاع الكتابة'],
      ['العبارة التعريفية في المنشورات والأغلفة', 'ExtraBold', 'عرض كبير 72–112 بكسل'],
      ['عناوين المنشورات والأغلفة والإعلانات', 'Bold / ExtraBold', '64–96 بكسل (منشور 1080)'],
      ['العناوين الفرعية والتصنيفات', 'Bold', '28–40 بكسل'],
      ['النصوص الأساسية والمطبوعات', 'Regular', '24–36 بكسل · 10–12 نقطة طباعة'],
      ['التعليقات والحواشي والبيانات النظامية', 'Light', '18–22 بكسل · 7–8 نقاط طباعة'],
      ['اللوحات الخارجية والإرشادية', 'Bold', 'ارتفاع الحرف ≥ 25 مم لكل 8 م'],
      ['الموقع والعروض التقديمية', 'Bold للعناوين · Regular للنص', 'حسب سلّم الأحجام'],
    ];
    const BG = [['#FFFFFF', 'أبيض', 'logo_arabic_color'], ['#F4F7F7', 'ضبابي', 'logo_arabic_color'], ['#FBF1DE', 'رملي فاتح', 'logo_arabic_color'], ['#0F4C5C', 'بترولي', 'logo_arabic_white'], ['#0A3540', 'بترولي داكن', 'logo_arabic_white'], ['#E9A23B', 'زعفراني (أحادي داكن)', 'logo_arabic_mono']];
    const LOGOS = [['logo_arabic_color', 'الشعار العربي – ملون'], ['logo_bilingual_color', 'الشعار ثنائي اللغة – ملون'], ['logo_english_color', 'الشعار الإنجليزي – ملون'], ['logo_stacked_color', 'الشعار العمودي – ملون'], ['icon_color', 'الرمز وحده – ملون'], ['logo_arabic_mono', 'الشعار العربي – أحادي'], ['icon_mono', 'الرمز – أحادي']];
    const LOGOS_W = [['logo_arabic_white', 'الشعار العربي – أبيض'], ['logo_bilingual_white', 'ثنائي اللغة – أبيض'], ['logo_stacked_white', 'العمودي – أبيض'], ['icon_white', 'الرمز – أبيض']];
    const POSTS = ['مرحباً بكم في سُلّم النمو', 'لماذا «سُلّم»؟ – فلسفتنا', 'نرتقي بالعلم، وننمو بالمستقبل', 'رؤيتنا ورسالتنا', 'قيمنا', 'البيئة التعليمية', 'مهارات المستقبل', 'الأنشطة والإبداع', 'التميز الأكاديمي', 'أولياء الأمور – شركاء في كل درجة', 'ترقّبوا – قريباً', 'التسجيل والقبول'];
    const TPL = ['نصيحة تربوية', 'إعلان', 'خبر أو فعالية', 'تهنئة بمناسبة', 'إنجاز طالب/ة', 'اقتباس تربوي', 'التسجيل مفتوح'];
    const APPS = [['A01_stationery', 'بطاقة التعريف والأوراق الرسمية والمراسلات'], ['A02_profile', 'الملف التعريفي للمدرسة'], ['A03_uniform', 'الزي المدرسي'], ['A04_signage', 'اللوحات الخارجية والإرشادية'], ['A05_website', 'الموقع الإلكتروني (تصور)'], ['A06_deck', 'العرض التعريفي لأولياء الأمور والمستثمرين']];
    const GUIDE = [['G00_cover', 'غلاف دليل الهوية'], ['G01_story', '01 فكرة الرمز وقصته'], ['G02_versions', '02 نسخ الشعار'], ['G03_rules', '03 مساحة الأمان والأحجام والخلفيات'], ['G04_colors', '04 لوحة الألوان الرسمية'], ['G05_type', '05 الخطوط ومواضع استخدامها'], ['G06_patterns', '06 الأنماط والعناصر المساندة'], ['G07_photo', '07 أسلوب الصور والتصوير'], ['G08_tone', '08 شخصية العلامة ونبرة التواصل']];
    const DIG = [['digital_identity_overview', 'صور الحسابات والأغلفة والنبذة التعريفية'], ['accounts_mockup', 'شكل الحسابات بعد اكتمال الهوية (تصور)'], ['profile_instagram_tiktok_snapchat_1080x1080', 'صورة الحساب – إنستغرام وتيك توك وسناب شات 1080×1080'], ['profile_x_400x400', 'صورة الحساب – X بمقاس 400×400'], ['x_header_1500x500', 'غلاف X بمقاس 1500×500'], ['tiktok_snapchat_video_cover_1080x1920', 'غلاف فيديو عمودي 1080×1920']];
    const HL = ['عن المدرسة', 'القبول', 'الأنشطة', 'أولياء الأمور', 'الفعاليات', 'تواصل'];

    return head('القسم 16', 'الهوية المؤسسية والمراجع البصرية', 'الهوية البصرية المعتمدة لمدارس سُلّم النمو (الإصدار 1.3): الشعار والألوان والخطوط وأسلوب التطبيق، مع المراجع البصرية الرسمية من حزمة الهوية.') +
      `<div class="note info" style="margin-bottom:16px"><b>مصدر كل الصور في هذا القسم:</b> حزمة الهوية البصرية والتسويقية الرسمية «حزمة_هوية_سلم_النمو» – الإصدار 1.3 (8 نوفمبر 2026)، وتُعرض كما هي دون تعديل. المسار الأصلي لكل صورة مكتوب أسفلها.</div>

      <div class="card"><h3>1. الشعار</h3><p class="sub">الاسم: سُلّم النمو | Growth Ladder Schools · العبارة: نرتقي بالعلم، وننمو بالمستقبل</p>
        <div class="grid g2 id-hero"><div class="id-logo-main"><img src="${D}logo_bilingual_color.svg" alt="شعار مدارس سُلّم النمو"></div>
        <div><h4 class="id-h4">فكرة الرمز</h4><ul class="list small">
          <li><b>حرف السين من «سُلّم»:</b> الرمز مبني من حرف «س»، أول حروف الاسم، بأسنانه الثلاث المميزة، فيرتبط الرمز بالاسم ارتباطاً بصرياً مباشراً.</li>
          <li><b>ثلاث درجات صاعدة:</b> ترتفع الأسنان الثلاث من اليمين إلى اليسار باتجاه القراءة العربية، كدرجات سُلّم، كل درجة تُبنى على ما قبلها.</li>
          <li><b>ميلان النبتة:</b> تنحني الدرجة الأعلى انحناءة خفيفة كساق نبتة تنمو نحو الضوء.</li>
          <li><b>نقطة الطالب:</b> النقطة الزعفرانية هي الطالب عند القمة، وهي العنصر اللوني المميز في الأنماط والمنشورات.</li>
          <li><b>الشعار المكتوب:</b> بخط المراعي ExtraBold محوّلاً إلى مسارات متجهة؛ لا يُعاد كتابته كنص ولا يُعدَّل تباعده.</li></ul></div></div>
        <h4 class="id-h4">نسخ الشعار</h4>
        <div class="id-logos">${LOGOS.map(([n, c]) => `<figure class="id-logo" data-img="${D}${n}.svg"><div><img loading="lazy" src="${D}${n}.svg" alt="${esc(c)}"></div><figcaption>${esc(c)}<span class="id-src">02_الشعار/SVG/${n}.svg</span></figcaption></figure>`).join('')}
          ${LOGOS_W.map(([n, c]) => `<figure class="id-logo dark" data-img="${D}${n}.svg"><div><img loading="lazy" src="${D}${n}.svg" alt="${esc(c)}"></div><figcaption>${esc(c)}<span class="id-src">02_الشعار/SVG/${n}.svg</span></figcaption></figure>`).join('')}</div>
      </div>
      <div style="height:16px"></div>

      <div class="card"><h3>2. الألوان الرسمية المعتمدة</h3><p class="sub">من دليل الهوية – لوحة الألوان الرسمية (04)</p>
        <div class="id-colors">${COLORS.map(([n, hex, rgb, role, use, fg]) => `<div class="id-sw"><div class="id-chip" style="background:${hex};color:${fg}">${esc(role)}</div><div class="id-info"><b>${esc(n)}</b><span class="code">${hex}</span><span class="small muted num">RGB ${rgb}</span><span class="small">${esc(use)}</span></div></div>`).join('')}</div>
        <h4 class="id-h4">نسب الاستخدام</h4>
        <div class="id-ratio"><i style="flex:50;background:#F4F7F7;color:#10242B">خلفيات 50%</i><i style="flex:30;background:#0F4C5C">بترولي 30%</i><i style="flex:10;background:#5BA4A4">أكوا 10%</i><i style="flex:10;background:#E9A23B;color:#10242B">زعفراني 10%</i></div>
        <p class="small muted" style="margin:10px 0 0">التباين: النص الأبيض على البترولي والبترولي الداكن والرمادي يحقق معيار WCAG AA. الزعفراني والأكوا للإبراز والأشكال فقط، ولا يُكتب بهما نص صغير على خلفية بيضاء؛ النص على الزعفراني يكون بالليلي.</p>
      </div>
      <div style="height:16px"></div>

      <div class="card"><h3>3. الخطوط المستخدمة</h3><p class="sub">الخط العربي الأساسي: المراعي (Almarai) – معتمد منذ الإصدار 1.3 · الخط الإنجليزي: Outfit</p>
        <div class="grid g2">
          <div><div class="id-weights">
            <div style="font-weight:300">نرتقي بالعلم<span>Light · 300</span></div>
            <div style="font-weight:400">نرتقي بالعلم<span>Regular · 400</span></div>
            <div style="font-weight:700">نرتقي بالعلم<span>Bold · 700</span></div>
            <div style="font-weight:800">نرتقي بالعلم<span>ExtraBold · 800</span></div></div>
            <div class="id-en en"><b>Growth Ladder Schools</b><span>Outfit Bold للاسم والعناوين · Regular للنصوص (Google Fonts – OFL)</span></div>
            <p class="small muted">أوزان المراعي المتوفرة: Light 300 · Regular 400 · Bold 700 · ExtraBold 800، ولا يوجد وزن Medium. الرخصة: SIL OFL 1.1، وتسمح بالاستخدام التجاري وتضمين الخط في المطبوعات والمواقع والتطبيقات، ولا يُباع الخط منفرداً.</p></div>
          <div class="tbl-wrap"><table class="t"><thead><tr><th>الموضع</th><th>الوزن</th><th>الحجم المرجعي</th></tr></thead><tbody>${FONTS.map(([a, b, c]) => `<tr><td>${esc(a)}</td><td><b class="en">${esc(b)}</b></td><td class="small">${esc(c)}</td></tr>`).join('')}</tbody></table></div>
        </div>
      </div>
      <div style="height:16px"></div>

      <div class="card"><h3>4. أسلوب الهوية وتطبيقها في المواد التعريفية والتسويقية</h3><p class="sub">من دليل الهوية: الشخصية ونبرة التواصل (08)، والصور (07)، والأنماط (06)، وقواعد الاستخدام (03)</p>
        <div class="grid g4">
          <div class="id-trait"><b>طموحة</b><span>نتحدث عن المستقبل بثقة، ونربط كل خطوة بغاية أكبر.</span></div>
          <div class="id-trait"><b>دافئة</b><span>قريبة من الأسرة والطفل، بلغة إنسانية لا رسمية جافة.</span></div>
          <div class="id-trait"><b>واضحة</b><span>جمل قصيرة ومعلومة محددة، بلا مبالغة أو وعود عامة.</span></div>
          <div class="id-trait"><b>أصيلة</b><span>عربية أولاً، تعتز بالهوية الوطنية وتنفتح على العالم.</span></div></div>
        <div class="grid g2" style="margin-top:14px">
          <div class="note info small"><b>هكذا نتحدث:</b> «خطوة اليوم… أساس الغد.» · «نعرف كل طالب باسمه، ونحتفي بتقدّمه مهما كان صغيراً.» · «شركاء في كل درجة.»<br><b>جوهر العلامة:</b> نرتقي بالعلم، وننمو بالمستقبل — كل رسالة تجيب: كيف يرتقي الطالب درجةً اليوم؟</div>
          <div class="note bad small"><b>ونتجنب:</b> «الأفضل في المملكة!!!» والمبالغات وعلامات التعجب المتكررة · المعلومات غير المؤكدة عن الرسوم أو المواعيد أو الاعتمادات · العامية المفرطة أو المصطلحات المعقدة.</div></div>
        <div class="grid g2" style="margin-top:14px">
          <div><h4 class="id-h4">قواعد استخدام الشعار</h4><ul class="list small">
            <li><b>مساحة الأمان:</b> س = ارتفاع نقطة الطالب × 2، ويُترك فراغ لا يقل عن (س) من كل جهة.</li>
            <li><b>الحد الأدنى:</b> الشعار الكامل 30 مم / 120 بكسل · الرمز وحده 8 مم / 24 بكسل · مع العبارة التعريفية 60 مم / 240 بكسل.</li>
            <li><b>الخلفيات المسموح بها:</b> أبيض، ضبابي، رملي فاتح، بترولي، بترولي داكن، زعفراني (بالنسخة الأحادية الداكنة).</li>
            <li><b>غير مسموح:</b> مطّ الشعار أو ضغطه، تغيير ألوانه، تدويره، إضافة ظلال أو مؤثرات، وضعه على خلفية متداخلة، عكس ترتيب الرمز والكتابة.</li></ul></div>
          <div><h4 class="id-h4">أسلوب الصور</h4><ul class="list small">
            <li><b>حتى الافتتاح تُستخدم الرسوم والأشكال المساندة بدلاً من الصور، ولا تُستخدم صور عشوائية لا تمثل المدرسة.</b></li>
            <li>بعد الافتتاح: لحظات تعلّم حقيقية، ضوء طبيعي ودافئ، تفاصيل اليد والأداة، والمعلم بجانب الطالب.</li>
            <li>نفعل: تصوير فعلي داخل المدرسة بعد الافتتاح، وموافقة خطية من ولي الأمر قبل نشر أي صورة لطالب، وتأطير الصور بإطار القوس أو زاوية الدرجة، ومراعاة الضوابط الشرعية والذوق العام في تصوير الطالبات.</li>
            <li>لا نفعل: صور مخزنة نمطية أو أجنبية الطابع، وصور أطفال حقيقيين دون إذن، وفلاتر مشبعة، وصور مولّدة توحي بأنها من داخل المدرسة.</li></ul></div></div>
      </div>
      <div style="height:16px"></div>

      <div class="card"><h3>5. نماذج تطبيقية للشعار والألوان والخطوط</h3><p class="sub">الشعار على الخلفيات المسموح بها، ونموذج منشور وبطاقة مبنيان من الألوان والخطوط الرسمية</p>
        <div class="id-bgs">${BG.map(([bg, l, n]) => `<div class="id-bg" style="background:${bg}"><img src="${D}${n}.svg" alt="الشعار على خلفية ${esc(l)}"><span style="color:${/0F4C5C|0A3540/.test(bg) ? '#DCEBEB' : '#4E6168'}">✓ ${esc(l)}</span></div>`).join('')}</div>
        <div class="grid g2" style="margin-top:14px">
          <div class="id-mock post"><div class="ey">مرحباً بكم في</div><img src="${D}logo_arabic_color.svg" alt="" class="wm"><p>مدرسة أهلية في الرياض، تؤمن أن لكل طالب درجته الأولى نحو مستقبل أكبر.</p><div class="steps"><i style="height:64px;background:#E9A23B"></i><i style="height:44px;background:#0F4C5C"></i><i style="height:28px;background:#5BA4A4"></i></div><div class="ft"><img src="${D}icon_color.svg" alt="">نرتقي بالعلم، وننمو بالمستقبل</div></div>
          <div class="id-mock card2"><img src="${D}logo_bilingual_white.svg" alt="" class="lg"><div class="ln"><b>الاسم الكامل</b><span>المسمى الوظيفي · Bold / Light</span></div><div class="dots"><i style="background:#E9A23B"></i><i style="background:#5BA4A4"></i><i style="background:#F4F7F7"></i></div><p>نموذج بطاقة تعريف: خلفية بترولية، شعار أبيض، نص ضبابي، ونقطة زعفرانية للإبراز.</p></div></div>
        <p class="small muted" style="margin:10px 0 0">النموذجان أعلاه يوضحان طريقة تطبيق القواعد فقط، وليسا تصاميم معتمدة. التصاميم المعتمدة هي صور الحزمة في الفقرات التالية.</p>
      </div>
      <div style="height:16px"></div>

      <div class="card"><h3>6. دليل الهوية البصرية</h3><p class="sub">صفحات الدليل الرسمي (9 صفحات) – اضغط على الصفحة للتكبير</p>
        ${gal(GUIDE.map(([n, c]) => fig(n, c, '06_دليل_الهوية/' + n + '.png')), 'wide')}</div>
      <div style="height:16px"></div>

      <div class="card"><h3>7. الصور والمراجع البصرية – الحسابات الرقمية</h3>
        <div class="note" style="margin:6px 0 14px"><b>تنبيه مهم:</b> بحسب حزمة الهوية نفسها، <b>المعرّفات (@) مقترحة ويجب حجزها والتحقق من توفرها، والحسابات لم تُنشأ بعد</b>، ولا يوجد حتى الآن موقع إلكتروني رسمي منشور للمدارس. لذلك لا توجد صور منشورة فعلياً على حسابات رسمية يمكن الاستشهاد بها أو الربط بأصلها. الصور أدناه هي <b>التصاميم الرسمية المعتمدة للحسابات</b> من حزمة الهوية، وتُستبدل بروابط الحسابات الفعلية وصورها بعد إنشائها والترخيص.</div>
        ${gal(DIG.slice(0, 2).map(([n, c]) => fig(n, c, '03_الحسابات_الرقمية/' + n + '.png')), 'big')}
        <div style="height:14px"></div>
        ${gal(DIG.slice(2).map(([n, c]) => fig(n, c, '03_الحسابات_الرقمية/' + n + '.png')))}
        <h4 class="id-h4">أغلفة Highlights لإنستغرام</h4>
        ${gal(HL.map((c, i) => fig('highlight_0' + (i + 1), c, '03_الحسابات_الرقمية/instagram_highlights/highlight_0' + (i + 1) + '.png')), 'hl')}
        <div class="tbl-wrap" style="margin-top:14px"><table class="t"><thead><tr><th>المنصة</th><th>المعرّف المقترح</th><th>الحالة</th></tr></thead><tbody>
          <tr><td>Instagram</td><td class="en">growthladder.schools</td><td><span class="tag c-ver">مقترح – يحتاج حجزاً وتحققاً</span></td></tr>
          <tr><td>X</td><td class="en">@growthladder</td><td><span class="tag c-ver">مقترح – يحتاج حجزاً وتحققاً</span></td></tr>
          <tr><td>TikTok</td><td class="en">@growthladder</td><td><span class="tag c-ver">مقترح – يحتاج حجزاً وتحققاً</span></td></tr>
          <tr><td>Snapchat</td><td>سُلّم النمو</td><td><span class="tag c-ver">مقترح – يحتاج حجزاً وتحققاً</span></td></tr>
          <tr><td>الموقع الإلكتروني الرسمي</td><td>—</td><td><span class="tag c-est">لا يوجد بعد (تصور في التطبيقات A05)</span></td></tr>
        </tbody></table></div>
        <p class="small muted" style="margin:8px 0 0">المعرّفات منقولة من لوحة «شكل الحسابات بعد اكتمال الهوية» في الحزمة.</p>
      </div>
      <div style="height:16px"></div>

      <div class="card"><h3>8. المنشورات التعريفية والتسويقية</h3><p class="sub">12 منشوراً من الحزمة (04_المنشورات). نصوص المنشورات صياغات مقترحة، ومنشورا التسجيل والقبول بعد الترخيص.</p>
        ${gal(POSTS.map((c, i) => { const n = 'post_' + String(i + 1).padStart(2, '0'); return fig(n, c, '04_المنشورات/' + n + '.png'); }))}</div>
      <div style="height:16px"></div>

      <div class="card"><h3>9. القوالب</h3><p class="sub">7 قوالب قابلة لإعادة الاستخدام من الحزمة (05_القوالب)</p>
        ${gal(TPL.map((c, i) => { const n = 'template_0' + (i + 1); return fig(n, c, '05_القوالب/' + n + '.png'); }))}</div>
      <div style="height:16px"></div>

      <div class="card"><h3>10. تطبيقات الهوية</h3><p class="sub">تطبيق الهوية على المطبوعات والزي واللوحات والموقع والعروض (07_التطبيقات)</p>
        ${gal(APPS.map(([n, c]) => fig(n, c, '07_التطبيقات/' + n + '.png')), 'wide')}</div>
      <div style="height:16px"></div>
      <div class="note small"><b>بنود قائمة من الحزمة:</b> الاسم والشعار يحتاجان فحصاً لدى وزارة التجارة والهيئة السعودية للملكية الفكرية · المعرّفات (@) مقترحة · نصوص المنشورات صياغات مقترحة · منشورا التسجيل والقبول بعد الترخيص.</div>`;
  }

  /* ================= معايير المباني التعليمية (5) واشتراطات المباني التعليمية (6) =================
   * المصدران المعتمدان (قُرئا كاملين):
   *  A: «اشتراطات المباني التعليمية الخاصة 1447هـ – 2025م» – وزارة البلديات والإسكان، قرار وزاري 4600425374/1 وتاريخ 1447/06/13هـ
   *     (تحل محل اشتراطات المدارس الأهلية الصادرة بالقرار 4100128804 وتاريخ 1441/04/11هـ، ويُعمل بها من تاريخ نشرها).
   *  B: «اللائحة المنظمة للمؤسسات التعليمية الخاصة» – وزارة التعليم، منشورة في جريدة أم القرى 1448/5/19هـ (2026/10/30م)،
   *     أقرها مجلس شؤون التعليم العام بالقرار (ق-ت 26/1/4) وتاريخ 1448/4/30هـ، ويُعمل بها من تاريخ نفاذ نظام التعليم العام (المادة 25).
   * أرقام البنود والصفحات منقولة من الوثيقتين كما هي. لا توجد في هذا الملف قيمة نظامية من خارجهما. */
  const DOC = {
    A: { s: 'اشتراطات المباني التعليمية الخاصة', e: 'وزارة البلديات والإسكان' },
    B: { s: 'اللائحة المنظمة للمؤسسات التعليمية الخاصة', e: 'وزارة التعليم' },
  };
  const NAT = { m: ['c-ok', 'إلزامي'], a: ['c-inv', 'يتطلب موافقة جهة مختصة'], g: ['c-as', 'توجيهي / مسموح'] };
  const natTag = k => `<span class="tag ${NAT[k][0]}">${NAT[k][1]}</span>`;
  const TOPICS = ['موقع المبنى واستعمال الأرض', 'ملاءمة الموقع للنشاط التعليمي', 'مساحة الأرض والمبنى والمساحات الداخلية', 'الأدوار والارتفاعات والارتدادات', 'الفصول الدراسية والطاقة الاستيعابية', 'الممرات والسلالم والمخارج', 'السلامة ومكافحة الحريق والإخلاء', 'المرافق الصحية', 'التهوية والإضاءة', 'مواقف السيارات ومداخل ومخارج المبنى', 'الوصول الشامل', 'التجهيزات والخدمات التعليمية', 'المستندات والتراخيص'];
  /* [موضوع، نص المتطلب، الوثيقة، البند، الصفحة، الطبيعة، الانطباق على المبنى القائم المستأجر، الجهة (اختياري)] */
  const REQS = [
    [0, 'تطبق الاشتراطات على المباني التعليمية الخاصة المنشأة على الأراضي المخصصة مرفقاً تعليمياً خاصاً أو مرفقاً تعليمياً حكومياً مستثمراً من القطاع الخاص، أو الأراضي المخصصة للاستعمال التجاري (أو السكني لدار الحضانة ورياض الأطفال).', 'A', '1.3 نطاق التطبيق', 5, 'm', 'يُفحص استعمال أرض العقار في الصك ورخصة البناء: تعليمي أو تجاري. الاستعمال السكني لا يصلح إلا لدار الحضانة ورياض الأطفال في مبنى مستقل.'],
    [0, 'عند إنشاء المبنى التعليمي على أرض مخصصة مرفقاً تعليمياً خاصاً أو حكومياً مستثمراً من القطاع الخاص يجب الالتزام بضوابط الاستفادة من المرافق التعليمية وتعديل النوع والمرحلة التعليمية للمرافق التعليمية المملوكة للقطاع الخاص (أراضي أو مباني) الصادرة من وزارة التعليم وما يطرأ عليها من تعديلات.', 'A', 'الباب الرابع – أولاً', 13, 'a', 'ينطبق إذا كان العقار المستأجر على أرض مرفق تعليمي.', 'وزارة التعليم'],
    [0, 'عند إنشاء المبنى التعليمي على أرض مخصصة للاستعمال التجاري يجب الالتزام بمتطلبات الباب الرابع (المتطلبات المكانية).', 'A', 'الباب الرابع – ثانياً', 13, 'm', 'ينطبق إذا كانت أرض العقار تجارية.'],
    [0, 'يجب أن يكون المبنى التعليمي مستقلاً. ويُستثنى نشاط المدارس الأجنبية (الجاليات) في جزء من مبنى قائم بشروط (ألا يتجاوز 170 طالباً في المبنى التجاري، ومدخل ومخرج مستقلان…).', 'A', '4.1 بند 3', 14, 'm', 'المبنى المستأجر يجب أن يكون مستقلاً بالكامل للمدرسة؛ الاستثناء لا ينطبق على مدرسة أهلية بالمنهج الوطني.'],
    [0, 'يجب أن تكون المباني التعليمية مستقلة، ويسمح بإقامة دار الحضانة ورياض الأطفال في جزء من مبنى قائم (تعليمي/تجاري) في الدور الأرضي بمدخل ومخرج مستقل ومهيأ للأطفال، على ألا تقل المساحة المخصصة للنشاط عن 50 م².', 'A', '4.1 بند 4', 14, 'm', 'خاص برياض الأطفال إذا كانت في جزء من مبنى قائم؛ قسم التمهيدي في المشروع جزء من المبنى المدرسي المستقل.'],
    [0, 'يسمح بإنشاء دار الحضانة ورياض الأطفال على الأراضي السكنية بشرط أن تكون في مبنى مستقل (لا يتطلب تعديل رخصة البناء إلى تجاري أو تعليمي).', 'A', '4.1 بند 5', 14, 'g', 'لا ينطبق على المرحلة الابتدائية.'],
    [1, 'الشارع: شارعان لا يقل عرض أحدهما عن 25 م، أو في حال الموقع على ثلاثة شوارع يُشترط ألا يقل عرض أحد الشوارع عن 20 م. ويسمح أن تكون دار الحضانة/رياض الأطفال على شارع واحد لا يقل عرضه عن 15 م إذا كانت في مبنى مستقل.', 'A', '4.1 جدول (1)', 13, 'm', 'ينطبق على العقار المرشح كما هو؛ لا يمكن معالجته بالتأهيل.'],
    [1, 'تكون المداخل والمخارج وفقاً لدراسة التأثير المروري المعتمدة للموقع.', 'A', '4.1 بند 9', 14, 'a', 'ينطبق؛ وتعديل مداخل ومخارج مبنى تعليمي قائم يستلزم دراسة تأثير مروري (7.1 بند 6).'],
    [1, 'ألا تقل المسافة الفاصلة بين المبنى التعليمي ومحطات الوقود أو الغاز عن القيم المحددة في اشتراطات مراكز الخدمة ومحطات الوقود.', 'A', '4.1 بند 10', 14, 'm', 'ينطبق؛ القيم في اشتراطات مراكز الخدمة ومحطات الوقود.'],
    [1, 'ألا تقل المسافة بين روضة الأطفال أو دار الحضانة وأقرب محطة وقود أو محل بيع الغاز عن 50 م (قطرياً).', 'A', '4.1 بند 11', 14, 'm', 'ينطبق لوجود مرحلة رياض الأطفال.'],
    [1, 'ألا تقل المسافة بين موقع روضة الأطفال أو دار الحضانة وتقاطعات الشوارع التجارية والرئيسة عن 50 م (طولياً).', 'A', '4.1 بند 12', 14, 'm', 'ينطبق لوجود مرحلة رياض الأطفال.'],
    [1, 'الالتزام بالقرارات الصادرة من الجهات ذات العلاقة بشأن المسافات المحددة بين روضة الأطفال أو دار الحضانة وأقرب مركز ضيافة أطفال أو روضة أو دار حضانة أو أي أنشطة أخرى.', 'A', '4.1 بند 13', 14, 'a', 'ينطبق لوجود مرحلة رياض الأطفال.'],
    [1, 'يقدم المكتب الهندسي الاستشاري المعتمد من الوزارة دراسة الحركة المرورية لأثر المشروع التعليمي – ماعدا مرحلتي الحضانة ورياض الأطفال في المباني المستقلة – على شبكة الطرق المحيطة، وفقاً للدليل الإجرائي لدراسات التأثير المروري الصادر بقرار الوزارة 4500956243/1 وتاريخ 1447/2/18هـ.', 'A', '4.5', 16, 'm', 'ينطبق على المدرسة (ابتدائي + تمهيدي).'],
    [2, 'الحد الأدنى لمساحة الأرض لكل طالب – الفئة (أ) الرياض: دار حضانة/رياض أطفال 5.00 م²، الطفولة المبكرة 4.70، الابتدائي 4.40، المتوسط 4.70، الثانوي 5.25، المجمع التعليمي 4.80 م².', 'A', '4.2 جدول (3)', 14, 'm', 'تحدد الحد الأعلى لعدد الطلاب في العقار حسب مساحة أرضه.'],
    [2, 'مساحة الأرض للمبنى التعليمي (مرحلة واحدة) = معدل مساحة الأرض للطالب × عدد الطلاب؛ ومساحة الأرض للمجمع التعليمي = (متوسط معدلات الأرض المطلوبة لكل مرحلة ÷ عدد أنواع المراحل بالمبنى) × العدد الكلي للطلاب.', 'A', '4.2 بند 2', 14, 'm', 'طريقة الحساب المستخدمة في أداة الفحص.'],
    [2, 'تعد المدرسة مجمعاً تعليمياً إذا كانت تحتوي على مرحلتين دراسيتين أو أكثر (بنين أو بنات) باستثناء مرحلتي الحضانة ورياض الأطفال.', 'A', '2.1 التعريفات', 9, 'm', 'مدرسة تضم رياض أطفال وابتدائي لا تُعد مجمعاً تعليمياً وفق هذا التعريف.'],
    [2, 'يتم تحديد المساحات الداخلية للمبنى التعليمي حسب متطلبات الجهة المشرفة.', 'A', '5.2 بند 1', 18, 'a', 'مساحات الفصول والمرافق الداخلية تعتمدها وزارة التعليم.', 'وزارة التعليم'],
    [2, 'يسمح بتوفير مخزن للمواد المرتبطة بالأنشطة التعليمية داخل المبنى التعليمي بمساحة (100 م²) كحد أقصى، ويمنع تخزين المواد التي ليس لها علاقة بالنشاط.', 'A', '5.2 بند 8', 18, 'g', 'المخزن مسموح بحد أقصى 100 م²، ولا يُشترط وجوده.'],
    [2, 'الالتزام بالحد الأدنى لمساحة الساحات الخارجية/الداخلية المخصصة للطلاب وفق متطلبات الجهة المشرفة، وألا تقل تغطية الساحات عن 75% من المساحة المخصصة لها مع ضمان الإضاءة والتهوية الطبيعية فيها.', 'A', '5.2 بند 9', 18, 'm', 'ينطبق؛ مساحة الساحات تحددها وزارة التعليم، والتغطية 75% حد أدنى.'],
    [2, 'يسمح بدور الميزانين والملحق في المباني التعليمية – حسب نظام البناء المعتمد في المنطقة – على أن تستخدم للخدمات المساندة، ويمنع استخدامها للفصول التعليمية.', 'A', '5.2 بند 4', 18, 'm', 'لا تُحتسب فصول في الميزانين أو الملحق.'],
    [2, 'يسمح بتوفير غرفة أو أكثر لسكن الحارس بمسطح لا يزيد عن 5% من مساحة الأرض في حدود الارتداد.', 'A', '5.2 بند 7', 18, 'g', 'مسموح.'],
    [3, 'عدد الأدوار – لا يتجاوز عدد أدوار المبنى التعليمي: دار الحضانة/رياض الأطفال أرضي + 1، الطفولة المبكرة أرضي + 2، الابتدائي والمتوسط والثانوي والمجمع التعليمي أرضي + 3؛ على ألا يتجاوز عدد الأدوار المسموح به بنظام البناء المعتمد.', 'A', '4.3 جدول (4) + جدول (1)', 15, 'm', 'ينطبق على العقار كما هو؛ رياض الأطفال لا تتجاوز الدور الأول.'],
    [3, 'في حالة وجود الموقع على شارعين أو أكثر يتم احتساب الأدوار من منسوب الشارع الأعلى.', 'A', '4.3 بند 2', 15, 'm', 'ينطبق.'],
    [3, 'يسمح بدور إضافي غير الأدوار التعليمية – حسب نظام البناء المعتمد وبعد موافقة الجهة المشرفة – على أرض تجارية، لمكاتب إدارة الشركة التعليمية فقط، بمدخل ومخرج مستقل ومع الخصوصية البصرية التامة لساحات الطلاب.', 'A', '4.3 بند 3', 15, 'a', 'خيار لاستخدام دور زائد في عقار تجاري.'],
    [3, 'في المباني التعليمية التي يزيد عدد أدوارها عن عدد أدوار المباني المجاورة يجب وضع معالجات معمارية لواجهات الأدوار المتكررة المطلة على المجاورين.', 'A', '4.3 بند 4', 15, 'm', 'ينطبق عند زيادة الأدوار عن المجاورين.'],
    [3, 'الارتدادات: حسب نظام البناء المعتمد في المنطقة؛ وفي المناطق التي ليس لها نظام بناء معتمد يكون الارتداد الأمامي خُمس عرض الشارع (3 م حداً أدنى إذا كان عرض الشارع 15 م فأقل، و6 م حداً أدنى إذا كان 30 م فأكثر)، وألا يقل الارتداد الجانبي والخلفي عن 3 م.', 'A', '4.1 جدول (1)', 13, 'm', 'يُتحقق من مطابقة المبنى القائم لرخصة البناء ونظام البناء.'],
    [3, 'يسمح باستخدام الدور الأول لفصول رياض الأطفال والحضانة فقط في حال كان صعود ونزول الأطفال بواسطة منحدر بين الدور الأرضي والدور الأول، ولا يسمح نهائياً باستخدام السلم (الدرج) للأطفال.', 'A', '5.2 بند 5', 18, 'm', 'فصول التمهيدي في الأرضي، أو في الأول بمنحدر فقط.'],
    [4, 'تقديم تقرير من مكتب هندسي مرخص يثبت جاهزية المبنى والمرافق التابعة له للاستخدام التعليمي، ويحدد الطاقة الاستيعابية القصوى وفق الاشتراطات المعتمدة.', 'B', 'المادة 6 فقرة 4 (د)', 6, 'm', 'شرط لمنح الترخيص؛ يحدد الطاقة القصوى للمبنى المستأجر.'],
    [4, 'يصنف المبنى التعليمي من قبل الجهة المشرفة على أساس إشغال المراحل التعليمية فيه.', 'A', '6.1 بند 4', 22, 'a', 'ينطبق.', 'وزارة التعليم'],
    [4, 'تستثنى المباني التعليمية القائمة أو التي صدر لها رخصة بناء قبل تاريخ اعتماد هذه الاشتراطات من تطبيق المتطلبات المكانية، وهي ملزمة بما ورد في اشتراطات المدارس الأهلية الصادرة بالقرار 4100128804 وتاريخ 1441/4/11هـ، وتجدد لها الرخصة التشغيلية وفقاً لذلك ما لم يتم تغيير المرحلة التعليمية لها أو زيادة الطاقة الاستيعابية؛ وفي حال تعديل المداخل والمخارج يلزم تقديم دراسة تأثير مروري.', 'A', '7.1 بند 6', 26, 'm', 'مهم للمبنى المستأجر: إن كان مبنى تعليمياً قائماً قبل الاشتراطات ولم تتغير مرحلته أو طاقته، يُستثنى من المتطلبات المكانية (الباب الرابع).'],
    [5, 'الالتزام بتحقيق المواصفات الفنية الواردة في كود البناء السعودي العام (SBC 201) في تنفيذ الأرصفة الخارجية والدرج والمنحدرات في المبنى التعليمي.', 'A', '5.4 بند 1', 19, 'm', 'ينطبق؛ تُفحص الأدراج والمنحدرات القائمة.'],
    [5, 'يجب أن يكون الدرج الخارجي للمبنى مصنوعاً من مواد آمنة ومقاومة للانزلاق وشديدة التحمل.', 'A', '5.4 بند 2', 19, 'm', 'ينطبق.'],
    [5, 'يجب وضع درج مدخل المبنى التعليمي أو المنحدر ضمن حدود المنشأة.', 'A', '5.4 بند 6', 19, 'm', 'ينطبق.'],
    [5, 'يجب أن يخلو الدرج والأرصفة الخارجية ومنحدر الأشخاص ذوي الإعاقة من أي تركيبات أو معدات تعيق دخول المبنى التعليمي أو الخروج أو الإخلاء السريع في حالات الحوادث أو الطوارئ أو الإصابات.', 'A', '5.4 بند 7', 19, 'm', 'ينطبق.'],
    [6, 'يجب الحصول على موافقة المديرية العامة للدفاع المدني قبل إصدار الترخيص.', 'A', '3.1 بند 2', 11, 'a', 'ينطبق.', 'المديرية العامة للدفاع المدني'],
    [6, 'وجود شهادة أو تصريح سلامة سارٍ من المديرية العامة للدفاع المدني، أو الجهة المختصة.', 'B', 'المادة 6 فقرة 4 (ج)', 6, 'm', 'شرط لمنح ترخيص وزارة التعليم.', 'المديرية العامة للدفاع المدني'],
    [6, 'يجب الرجوع لكافة المتطلبات الواردة في كود البناء السعودي للحرائق (SBC-801).', 'A', '5.8', 20, 'm', 'ينطبق؛ يُقيّم المبنى القائم من مختص سلامة.'],
    [6, 'الالتزام بتطبيق متطلبات كود البناء السعودي (ومنها الإنشائية والكهربائية والميكانيكية والتهوية والتكييف والصحية وترشيد الطاقة والمياه) وجميع اشتراطات الجهة المشرفة، مع تنفيذ الأعمال والتجهيزات والمواد طبقاً للمواصفات القياسية المعتمدة.', 'A', '5.1 بند 2', 18, 'm', 'ينطبق على أعمال التأهيل والمبنى القائم.'],
    [6, 'يمنع استخدام خامة الأسبستوس (Asbestos) أو أي مادة يدخل الأسبستوس في تركيبها.', 'A', '5.1 بند 4', 18, 'm', 'يُفحص المبنى القائم.'],
    [6, 'يجب أن تكون الأسقف من مواد مقاومة للحريق ولا تتأثر بالرطوبة، والألواح المعدنية أو الهياكل الحديدية معالجة بمواد مقاومة للصدأ.', 'A', '5.6 بند 1 و3', 20, 'm', 'ينطبق.'],
    [6, 'يجب تركيب كاميرات مراقبة أمنية وفقاً لما ورد في نظام استخدام كاميرات المراقبة الأمنية ولائحته التنفيذية.', 'A', '6.1 بند 10', 22, 'm', 'ينطبق.'],
    [6, 'يسمح بإقامة دور القبو على كامل حدود الملكية، ويسمح أن يستخدم لمواقف السيارات أو غرف الخدمات الكهروميكانيكية والخدمات المساندة للعملية التعليمية مثل المعامل والمختبرات وغيرها، أو الأنشطة الترفيهية مثل الملاعب والمسابح – ماعدا مرحلتي الحضانة ورياض الأطفال – بعد موافقة الجهة المشرفة والمديرية العامة للدفاع المدني، مع ضمان التهوية والإضاءة الطبيعية للقبو وتوفير مخارج الطوارئ.', 'A', '5.2 بند 6', 18, 'a', 'ينطبق إذا كان في العقار قبو.'],
    [7, 'الالتزام بالمتطلبات الصحية في كود البناء السعودي (ضمن متطلبات 5.1)، ومنها الكود السعودي الصحي (SBC-701) في الأدلة المرجعية.', 'A', '5.1 بند 2 و8.1', 18, 'm', 'عدد دورات المياه وتوزيعها ضمن المساحات الداخلية التي تحددها وزارة التعليم.'],
    [7, 'يجب أن تكون جميع فتحات التصريف وخزانات المياه والصرف الصحي وغرف التفتيش محكمة الإغلاق.', 'A', '5.5 بند 2', 19, 'm', 'ينطبق.'],
    [7, 'في حال وجود دورات مياه أو مغاسل أو مصليات يوصى بالحفاظ على نظافتها الدورية، وتوفير صابون الأيدي، والمناشف الورقية أو أجهزة التجفيف، وأوعية نفايات لا تفتح باليد.', 'A', '6.5', 24, 'g', 'متطلب تشغيلي (يوصى).'],
    [8, 'يجب تزويد المبنى التعليمي ومرافقه بما فيها المخزن بوسائل ومعدات إضاءة وفقاً للمواصفات القياسية المعتمدة، والتأكد من كفاءة توزيع الضوء وتغطيته لجميع المساحات.', 'A', '5.7 بند 1 و2', 20, 'm', 'ينطبق.'],
    [8, 'يمنع استخدام الإنارة الغازية خارج المبنى التعليمي.', 'A', '5.7 بند 3', 20, 'm', 'ينطبق.'],
    [8, 'ضمان الإضاءة والتهوية الطبيعية في الساحات (5.2 بند 9) وفي القبو عند استخدامه (5.2 بند 6)، وتطبيق متطلبات التهوية والتكييف في كود البناء السعودي (5.1 بند 2).', 'A', '5.1 و5.2', 18, 'm', 'ينطبق.'],
    [8, 'يمنع تركيب الأسلاك الكهربائية والأنابيب الصحية وغيرها من تمديدات الخدمة ووحدات التكييف الخارجية على واجهة المبنى التعليمي أو بشكل ظاهر على السطح.', 'A', '5.3 بند 2', 19, 'm', 'يُفحص المبنى القائم؛ يحتاج معالجة إن وُجدت.'],
    [9, 'الحد الأدنى للمواقف – الفئة (أ): الحضانات ورياض الأطفال موقف لكل 3 فصول، الابتدائية موقفان لكل فصل، المتوسطة موقفان لكل فصل، الثانوية 4 مواقف لكل فصل، ومواقف الحافلات موقف لكل 3 فصول.', 'A', '4.1 جدول (2)', 13, 'm', 'تُحسب على فصول العقار لكل مرحلة.'],
    [9, 'تكون متطلبات المواقف للمبنى التعليمي حسب مكونات المبنى من مراحل تعليمية، وتبقى المساحة المخصصة لمواقف السيارات والحافلات والتنزيل والتحميل بدون أسوار ضمن حدود الموقع؛ وعند إضافة دور لمكاتب إدارة الشركة التعليمية يلزم 4 مواقف لكل 100 م² من مساحته.', 'A', 'ملاحظات جدول (2)', 13, 'm', 'ينطبق.'],
    [9, 'يكون مجمل مساحة المواقف المطلوبة ضمن مسطح الأرض داخل حدود الملكية عدا المواقف العامة ضمن قطاعات الشوارع المحيطة.', 'A', '4.4 بند 3', 16, 'm', 'ينطبق.'],
    [9, 'فصل حركة مواقف السيارات والحافلات عن المساحة المخصصة للطلاب بأسوار نافذة داخلية.', 'A', '4.4 بند 4', 16, 'm', 'قد يحتاج أعمال تأهيل.'],
    [9, 'تكون بداية المنحدر الخاص بدخول القبو بعد الانتهاء من حد الرصيف الخارجي للمبنى التعليمي.', 'A', '4.4 بند 5', 16, 'm', 'ينطبق إذا كان في العقار قبو مواقف.'],
    [9, 'إحاطة المبنى بأسوار ارتفاعها 2.40 م كحد أدنى، ويسمح أن تكون الأسوار من النوع النافذ للمباني التعليمية المخصصة للبنين، ويجب أن تكون مصمتة (غير نافذة) للمباني التعليمية المخصصة للبنات.', 'A', '4.4 بند 6', 16, 'm', 'ينطبق؛ قسم البنات يحتاج سوراً مصمتاً.'],
    [10, 'يلزم توفير مواقف سيارات للأشخاص ذوي الإعاقة بالقرب من المداخل الرئيسية وفق المعدلات والمواصفات الواردة في كود البناء السعودي العام (SBC 201).', 'A', '4.4 بند 2', 16, 'm', 'ينطبق.'],
    [10, 'يجب توفير منحدر للأشخاص ذوي الإعاقة في حدود المبنى التعليمي (ويستثنى من ذلك المنشآت المتواجدة في مبانٍ توفر متطلبات الأشخاص ذوي الإعاقة).', 'A', '5.4 بند 3', 19, 'm', 'ينطبق.'],
    [10, 'يجب الالتزام بتحقيق متطلبات الأشخاص ذوي الإعاقة طبقاً للفصل (1009) والباب (11) من كود البناء السعودي (SBC 201) وحسب دليل معايير الوصول الشامل للمنشآت، وأن يكون المنحدر خالياً من العيوب ولا يؤثر على الرصيف الخارجي، والالتزام بالميول المسموح بها لمنحدر الرصيف.', 'A', '5.4 بنود 4 و5 و8', 19, 'm', 'ينطبق؛ قد يحتاج أعمال تأهيل.'],
    [11, 'يسمح بجزء مستقل من المبنى التعليمي أو ملحقاً به بمدخل ومخرج خاص يستخدم لمكاتب إدارة الإشراف العام على المبنى التعليمي.', 'A', '5.2 بند 3', 18, 'g', 'مسموح.'],
    [11, 'يسمح بممارسة الأنشطة الإضافية المعتمدة من الوزارة والجهة المشرفة: ألا يتجاوز مجموع مساحاتها 50% من المساحة الإجمالية الواردة في الترخيص، مع المحافظة على الحد الأدنى لمساحة النشاط الرئيسي، وألا تقل مساحة النشاط الإضافي عن 6 م² (ويستثنى من ذلك مكائن الخدمة الذاتية).', 'A', '4.6', 16, 'a', 'ينطبق عند إضافة أنشطة.'],
    [11, 'يجوز للمؤسسة التعليمية الخاصة ممارسة أنشطة مساندة أو مكملة للتعليم (الرياضية والثقافية والفنية والتقنية وخدمات النقل والتغذية والسكن وغيرها) متى كانت مرتبطة بنشاطها التعليمي ومستوفية للتراخيص والموافقات اللازمة؛ ولا يجوز استخدام مرافقها بصورة تتعارض مع أغراضها التعليمية أو تؤثر في سلامة الطلاب أو جودة العملية التعليمية.', 'B', 'المادة 8 فقرتا 2 و3', 7, 'a', 'ينطبق.'],
    [11, 'تلتزم المؤسسة التعليمية الخاصة بحماية حقوق الطلاب وسلامتهم ورفاههم وتوفير بيئة تعليمية آمنة وصحية ومحفزة، ووضع خطة لاستمرارية التعليم في حالات الطوارئ أو الكوارث أو الأعطال.', 'B', 'المادة 13 فقرتا 12 و16', 11, 'm', 'ينطبق طوال مدة الترخيص.'],
    [12, 'يجب أن تكون إجراءات الإصدار والإلغاء والتجديد والتعديل والإيقاف للترخيص وفقاً لنظام إجراءات التراخيص البلدية ولائحته التنفيذية؛ وتقديم سجل تجاري ساري المفعول يحتوي على النشاط المطلوب ترخيصه؛ وموافقة الجهة المشرفة على النشاط.', 'A', '3.1 بنود 1 و3 و4', 11, 'm', 'ينطبق.'],
    [12, 'يجب الحصول على الترخيص البلدي قبل البدء في أعمال التشغيل، والحصول على ترخيص تشغيلي من الجهة المشرفة قبل البدء في أعمال التشغيل، ويمنع ممارسة نشاط مخالف أو إضافي غير مدرج في الترخيص.', 'A', '6.1 بنود 1 و2 و5 و7', 22, 'm', 'ينطبق.'],
    [12, 'لا يجوز لأي شخص فتح مؤسسة تعليمية خاصة أو تشغيلها إلا بعد الحصول على الترخيص اللازم من وزارة التعليم، ولا يجوز الإعلان قبل الحصول على الترخيص، ولا يعد أي ترخيص أو موافقة من جهة أخرى بديلاً عن ترخيص الوزارة.', 'B', 'المادة 4 فقرات 1–3', 3, 'm', 'ينطبق.'],
    [12, 'يمر طلب فتح المؤسسة التعليمية الخاصة بمرحلتين: التأهيل ثم الترخيص. تصدر وثيقة التأهيل خلال مدة لا تتجاوز 10 أيام عمل من اكتمال الطلب، وتكون سارية لمدة سنة، ولا تعد ترخيصاً، ويجوز لحاملها الإعلان عن المشروع لأغراض الاستثمار أو التوظيف أو استكمال الإنشاء مع الإفصاح بأنه في مرحلة التأهيل.', 'B', 'المادة 5', 4, 'm', 'الخطوة الأولى لدى وزارة التعليم.'],
    [12, 'يشترط لمنح الترخيص: (أ) سجل تجاري أو وثيقة نظامية سارية، (ب) رخصة سارية للنشاط والموقع صادرة من وزارة البلديات والإسكان وفق الاشتراطات المعتمدة للمباني التعليمية الخاصة، (ج) شهادة أو تصريح سلامة سارٍ من الدفاع المدني، (د) تقرير مكتب هندسي مرخص بجاهزية المبنى والطاقة القصوى، (هـ) خطة دراسية معتمدة من المركز الوطني للمناهج، (و) تحديد مواد الهوية الوطنية وآلية تدريسها. وتبلغ الوزارة قرارها خلال 30 يوماً.', 'B', 'المادة 6 فقرتا 4 و5', 5, 'm', 'المستندات المطلوبة قبل التشغيل.'],
    [12, 'يُمنح الترخيص لمدة لا تقل عن سنة، وتتقدم المؤسسة بطلب التجديد قبل انتهائه بما لا يقل عن 90 يوماً.', 'B', 'المادة 7 فقرتا 1 و3', 6, 'm', 'ينطبق.'],
    [12, 'الالتزام باعتماد كافة أعمال التصميم والإشراف على التنفيذ من قبل مكاتب هندسية معتمدة لدى الوزارة، وأن يتم التنفيذ بواسطة مقاولين معتمدين لدى الوزارة.', 'A', '5.1 بندا 5 و6', 18, 'm', 'ينطبق على أعمال تأهيل المبنى المستأجر.'],
    [12, 'يجب الالتزام بالاشتراطات الفنية لاستعراض الرمز الإلكتروني الموحد (QR-Code) على واجهة المبنى التعليمي، والقيام بالصيانة الدورية والوقائية للعناصر الإنشائية والتجهيزات والتمديدات وأنظمة السلامة.', 'A', '6.2 و6.3', 22, 'm', 'متطلب تشغيلي.'],
    [12, 'تطبق هذه الاشتراطات اعتباراً من تاريخ نشرها، وتحل محل اشتراطات المدارس الأهلية الصادرة بالقرار الوزاري 4100128804 وتاريخ 1441/04/11هـ؛ ويُعمل باللائحة المنظمة للمؤسسات التعليمية الخاصة من تاريخ نفاذ نظام التعليم العام.', 'A', 'القرار الوزاري (ثانياً وثالثاً)', 2, 'm', 'تاريخ السريان.'],
  ];
  const refCell = r => `${esc(DOC[r[2]].s)}<div class="small muted">البند: ${esc(r[3])} · صفحة ${fInt(r[4])}</div>`;

  /* ---------- 6) اشتراطات المباني التعليمية ---------- */
  let regTopic = -1;
  function officialRefsCard() {
    return `<div class="grid g2">
      <div class="card"><h3>اشتراطات المباني التعليمية الخاصة 1447هـ – 2025م</h3><table class="t"><tbody>
        <tr><th>الجهة</th><td>وزارة البلديات والإسكان</td></tr>
        <tr><th>أداة الإصدار</th><td>قرار وزاري رقم 4600425374/1 وتاريخ 1447/06/13هـ</td></tr>
        <tr><th>السريان</th><td>يُعمل بها من تاريخ نشرها على الموقع الإلكتروني للوزارة، وتحل محل اشتراطات المدارس الأهلية الصادرة بالقرار 4100128804 وتاريخ 1441/04/11هـ</td></tr>
        <tr><th>الصفحة الرسمية</th><td><a href="https://momah.gov.sa/ar/node/16243" target="_blank" rel="noopener">momah.gov.sa/ar/node/16243</a></td></tr>
        <tr><th>المحتوى</th><td>29 صفحة: التصنيف، الترخيص، المتطلبات المكانية، المتطلبات الفنية، المتطلبات التشغيلية، الأحكام العامة، الأدلة المرجعية</td></tr></tbody></table></div>
      <div class="card"><h3>اللائحة المنظمة للمؤسسات التعليمية الخاصة</h3><table class="t"><tbody>
        <tr><th>الجهة</th><td>وزارة التعليم</td></tr>
        <tr><th>أداة الإصدار</th><td>صدرت استناداً إلى الفقرة (5) من المادة (الثامنة) من نظام التعليم العام الصادر بالمرسوم الملكي رقم (م/36) وتاريخ 1448/1/27هـ، وأقرها مجلس شؤون التعليم العام بالقرار (ق-ت 26/1/4) وتاريخ 1448/4/30هـ الموافق 2026/10/11م</td></tr>
        <tr><th>النشر</th><td>جريدة أم القرى – 1448/5/19هـ الموافق 2026/10/30م</td></tr>
        <tr><th>السريان</th><td>يُعمل بها من تاريخ نفاذ نظام التعليم العام (المادة 25)</td></tr>
        <tr><th>الصفحة الرسمية</th><td><a href="https://www.uqn.gov.sa/decisions-and-regulations/4002248" target="_blank" rel="noopener">uqn.gov.sa/decisions-and-regulations/4002248</a></td></tr></tbody></table></div></div>`;
  }
  function pRegulations() {
    const list = REQS.map((r, i) => ({ r, i })).filter(x => regTopic < 0 || x.r[0] === regTopic);
    const cnt = k => REQS.filter(r => r[5] === k).length;
    const rows = list.map(({ r }) => `<tr><td class="small muted">${esc(TOPICS[r[0]])}</td><td>${esc(r[1])}</td><td class="small">${esc(r[7] || DOC[r[2]].e)}</td><td class="small">${refCell(r)}</td><td>${natTag(r[5])}</td><td class="small">${esc(r[6])}</td></tr>`).join('');
    return head('القسم 6', 'اشتراطات المباني التعليمية', 'المتطلبات النظامية المستخلصة من «اشتراطات المباني التعليمية الخاصة» (وزارة البلديات والإسكان) و«اللائحة المنظمة للمؤسسات التعليمية الخاصة» (وزارة التعليم)، مع مدى انطباق كل متطلب على مبنى قائم سيتم استئجاره وتجهيزه للاستخدام التعليمي في الرياض (الفئة أ).') +
      officialRefsCard() + `<div style="height:16px"></div>
      <div class="card"><h3>ملاحظات جوهرية للمشروع</h3><ul class="list small">
        <li><b>التصنيف الجغرافي:</b> الرياض ضمن <b>الفئة (أ)</b> (2.2، صفحة 9)، وعليها تُبنى معدلات مساحة الأرض والمواقف.</li>
        <li><b>المجمع التعليمي:</b> تعد المدرسة مجمعاً تعليمياً إذا احتوت مرحلتين دراسيتين أو أكثر <b>باستثناء مرحلتي الحضانة ورياض الأطفال</b> (2.1، صفحة 9)؛ فمدرسة تضم التمهيدي والابتدائي لا تُعد مجمعاً بهذا التعريف.</li>
        <li><b>المباني التعليمية القائمة:</b> تُستثنى المباني التعليمية القائمة أو التي صدرت لها رخصة بناء قبل اعتماد الاشتراطات من المتطلبات المكانية ما لم تتغير مرحلتها التعليمية أو تزد طاقتها الاستيعابية، وتبقى ملزمة باشتراطات 1441هـ (7.1 بند 6، صفحة 26).</li>
        <li><b>استقلال المبنى:</b> يجب أن يكون المبنى التعليمي مستقلاً (4.1 بند 3، صفحة 14).</li>
        <li><b>مسار الترخيص لدى وزارة التعليم:</b> وثيقة التأهيل ثم الترخيص، ومن شروطه رخصة البلدية للنشاط والموقع وشهادة الدفاع المدني وتقرير المكتب الهندسي بالجاهزية والطاقة القصوى (اللائحة، المادتان 5 و6).</li></ul></div>
      <div style="height:16px"></div>
      <div class="card"><h3>جدول الاشتراطات (${fInt(REQS.length)} متطلباً)</h3>
        <div class="legend" style="margin-bottom:10px"><b>طبيعة المتطلب:</b>${['m', 'a', 'g'].map(k => natTag(k) + ` <span class="small muted">${fInt(cnt(k))}</span>`).join(' ')}</div>
        <div class="tabs no-print"><button data-regtopic="-1" class="${regTopic < 0 ? 'on' : ''}">الكل</button>${TOPICS.map((t, i) => `<button data-regtopic="${i}" class="${regTopic === i ? 'on' : ''}">${esc(t)}</button>`).join('')}</div>
        <div class="tbl-wrap"><table class="t"><thead><tr><th>الموضوع</th><th>نص المتطلب أو مضمونه</th><th>الجهة المختصة</th><th>المرجع</th><th>الطبيعة</th><th>الانطباق على المبنى القائم المستأجر</th></tr></thead><tbody>${rows}</tbody></table></div></div>
      <div style="height:16px"></div>
      <div class="card"><h3>الأدلة والأكواد المرجعية</h3><p class="sub">كما وردت في الباب الثامن (8.1، صفحة 28)</p><ol class="list small">
        <li>نظام إجراءات التراخيص البلدية ولائحته التنفيذية.</li><li>لائحة الجزاءات عن المخالفات البلدية.</li><li>الأنظمة واللوائح والقواعد التنفيذية والفنية الصادرة من وزارة البلديات والإسكان.</li><li>الدليل الإرشادي للوحات التجارية.</li><li>الدليل الإجرائي لدراسات التأثير المروري للمشاريع العمرانية (قرار 4500956243/1 وتاريخ 1447/2/18هـ).</li><li>نظام استخدام كاميرات المراقبة الأمنية ولائحته التنفيذية.</li><li>الدليل التنظيمي الصادر من الأمانة التي يقع النشاط داخل حدودها.</li><li>ضوابط الاستفادة من المرافق التعليمية وتعديل النوع والمرحلة التعليمية للمرافق التعليمية المملوكة للقطاع الخاص – وزارة التعليم.</li><li>الدليل الاسترشادي لمعايير تصميم المدارس الأهلية والعالمية – وزارة التعليم.</li><li>الأنظمة واللوائح الفنية الصادرة عن الهيئة السعودية للمواصفات والمقاييس والجودة.</li><li>الدليل الإرشادي للوصول الشامل في البيئة العمرانية – مركز الملك سلمان لأبحاث الإعاقة.</li><li>المتطلبات الوقائية للحماية من الحريق – الدفاع المدني.</li><li>كود البناء السعودي: SBC-201، SBC 301-306، SBC-401، SBC-501، SBC-601، SBC-701، SBC-801.</li></ol></div>`;
  }

  /* ---------- 5) معايير المباني التعليمية: فحص العقارات المرشحة ---------- */
  const PROP_KEY = 'gl-properties-v2';
  const ls = (k, d) => { try { const v = JSON.parse(localStorage.getItem(k) || 'null'); return v == null ? d : v; } catch (e) { return d; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { toast('تعذّر الحفظ في هذا المتصفح'); } };
  const PROP = { list: () => ls(PROP_KEY, []), save: l => lsSet(PROP_KEY, l) };
  /* القيم النظامية المستخدمة في الفحص – الفئة (أ) – مع مراجعها */
  const RULE = {
    lpsKG: [5.00, '4.2 جدول (3) ص14'], lpsPrim: [4.40, '4.2 جدول (3) ص14'], lpsComplex: [4.80, '4.2 جدول (3) ص14'],
    parkKGper: [3, '4.1 جدول (2) ص13'], parkPrim: [2, '4.1 جدول (2) ص13'], busPer: [3, '4.1 جدول (2) ص13'],
    floorsPrim: [4, '4.3 جدول (4) ص15'], floorsKG: [2, '4.3 جدول (4) ص15'], fence: [2.40, '4.4 بند 6 ص16'], shade: [0.75, '5.2 بند 9 ص18'],
  };
  const YN = ['', 'نعم', 'لا', 'غير معروف'];
  const FIELDS = [
    ['g', 'بيانات العقار'], ['name', 'اسم العقار / الوصف', 'text'], ['district', 'الحي / الموقع', 'text'],
    ['landUse', 'استعمال الأرض (حسب الصك/رخصة البناء)', ['', 'مرفق تعليمي', 'تجاري', 'سكني', 'غير ذلك', 'غير معروف']],
    ['independent', 'هل المبنى مستقل بالكامل للمدرسة؟', YN],
    ['existingEdu', 'هل هو مبنى تعليمي قائم/مرخص قبل اعتماد الاشتراطات (1447هـ)؟', YN],
    ['stageChange', 'هل سيتغير مرحلته التعليمية أو تزيد طاقته عن الترخيص السابق؟', YN],
    ['g', 'الموقع والشوارع'], ['streets', 'عدد الشوارع المحيطة', 'num'], ['w1', 'عرض أعرض شارع (م)', 'num'],
    ['fuel50', 'التمهيدي: المسافة لأقرب محطة وقود/محل غاز ≥ 50 م؟', YN], ['inter50', 'التمهيدي: المسافة لتقاطع شارع تجاري/رئيس ≥ 50 م؟', YN],
    ['traffic', 'دراسة تأثير مروري معتمدة؟', YN],
    ['g', 'المساحات والأدوار'], ['landArea', 'مساحة الأرض (م²)', 'num'], ['bldArea', 'مساحة البناء الإجمالية (م²)', 'num'], ['floors', 'عدد الأدوار فوق الأرض (من منسوب الشارع الأعلى)', 'num'],
    ['kgFloor', 'موقع فصول التمهيدي', ['', 'الدور الأرضي', 'الدور الأول بمنحدر', 'الدور الأول بدرج فقط', 'أعلى من الدور الأول']],
    ['shadePct', 'نسبة تغطية الساحات (%)', 'num'], ['setbacks', 'المبنى مطابق لرخصة البناء والارتدادات النظامية؟', YN],
    ['g', 'الفصول والطلاب'], ['calc', 'أساس حساب مساحة الأرض', ['مجمع تعليمي (4.80)', 'حسب كل مرحلة (5.00 / 4.40)']],
    ['kgC', 'فصول التمهيدي', 'num'], ['prC', 'فصول الابتدائي', 'num'], ['kgS', 'طلاب التمهيدي المتوقعون', 'num'], ['prS', 'طلاب الابتدائي المتوقعون', 'num'],
    ['g', 'المواقف والأسوار والوصول'], ['parking', 'مواقف السيارات داخل الملكية', 'num'], ['bus', 'مواقف الحافلات', 'num'],
    ['fenceH', 'ارتفاع السور الخارجي (م)', 'num'], ['girlsSolid', 'سور قسم البنات مصمت (غير نافذ)؟', YN], ['innerFence', 'أسوار داخلية نافذة تفصل حركة السيارات والحافلات عن الطلاب؟', YN],
    ['accPark', 'مواقف ذوي الإعاقة قرب المداخل الرئيسية؟', YN], ['ramp', 'منحدر لذوي الإعاقة ضمن حدود المبنى؟', YN],
    ['g', 'السلامة والأنظمة'], ['asbestos', 'المبنى خالٍ من الأسبستوس؟', YN], ['facade', 'الواجهات خالية من التمديدات ووحدات التكييف الظاهرة؟', YN], ['cameras', 'كاميرات مراقبة أمنية وفق النظام؟', YN],
  ];
  const DOCS = [['deed', 'صك الملكية وعقد الإيجار'], ['permit', 'رخصة البناء'], ['occupancy', 'شهادة الإشغال'], ['cr', 'سجل تجاري يتضمن النشاط'], ['civil', 'موافقة/شهادة سلامة الدفاع المدني'], ['eng', 'تقرير مكتب هندسي مرخص بالجاهزية والطاقة القصوى'], ['qual', 'وثيقة التأهيل من وزارة التعليم'], ['muni', 'الترخيص البلدي للنشاط والموقع']];
  const nz = v => v === '' || v === null || v === undefined ? null : Number(v);

  function evalProp(p) {
    const out = [], add = (lvl, text, ref) => out.push({ lvl, text, ref });
    const kgC = nz(p.kgC) || 0, prC = nz(p.prC) || 0, kgS = nz(p.kgS) || 0, prS = nz(p.prS) || 0;
    const classes = kgC + prC, students = kgS + prS, m = { classes, students };
    const miss = [];
    [['landUse', 'استعمال الأرض'], ['independent', 'استقلال المبنى'], ['streets', 'عدد الشوارع'], ['w1', 'عرض أعرض شارع'], ['landArea', 'مساحة الأرض'], ['bldArea', 'مساحة البناء'], ['floors', 'عدد الأدوار'], ['parking', 'عدد المواقف'], ['bus', 'مواقف الحافلات']]
      .forEach(([f, l]) => { const v = p[f]; if (v === undefined || v === null || v === '' || v === 'غير معروف') miss.push(l); });
    if (!classes) miss.push('عدد الفصول'); if (!students) miss.push('عدد الطلاب');
    if (miss.length) add('miss', 'بيانات ناقصة: ' + miss.join('، '), '');
    // استثناء المباني التعليمية القائمة من المتطلبات المكانية
    const exempt = p.existingEdu === 'نعم' && p.stageChange === 'لا';
    if (exempt) add('info', 'مبنى تعليمي قائم قبل الاشتراطات دون تغيير المرحلة أو زيادة الطاقة: يُستثنى من المتطلبات المكانية (الباب الرابع) ويلتزم باشتراطات المدارس الأهلية 1441هـ؛ وتعديل المداخل والمخارج يستلزم دراسة تأثير مروري.', '7.1 بند 6 ص26');
    const spatial = (lvl, text, ref) => add(exempt ? 'info' : lvl, (exempt ? '(مستثنى) ' : '') + text, ref);
    // الاستعمال والاستقلال
    if (p.landUse === 'سكني') spatial('fail', 'الاستعمال سكني: الاشتراطات تسري على أراضي المرافق التعليمية والتجارية، والسكني لدار الحضانة ورياض الأطفال فقط في مبنى مستقل؛ لا يصلح للمرحلة الابتدائية.', '1.3 ص5 و4.1 بند 5 ص14');
    else if (p.landUse === 'غير ذلك') spatial('fail', 'استعمال الأرض خارج نطاق تطبيق الاشتراطات (مرفق تعليمي أو تجاري).', '1.3 ص5');
    else if (p.landUse === 'مرفق تعليمي') add('info', 'أرض مرفق تعليمي: يلزم الالتزام بضوابط الاستفادة من المرافق التعليمية الصادرة من وزارة التعليم.', 'الباب الرابع أولاً ص13');
    if (p.independent === 'لا') add('fail', 'المبنى غير مستقل: يجب أن يكون المبنى التعليمي مستقلاً.', '4.1 بند 3 ص14');
    // الشوارع
    const st = nz(p.streets), w = nz(p.w1);
    if (st != null && w != null) {
      const okSt = (st >= 3 && w >= 20) || (st === 2 && w >= 25) || (st > 3 && w >= 20);
      if (!okSt) spatial('fail', st < 2 ? `العقار على ${fInt(st)} شارع: المطلوب شارعان أحدهما ≥ 25 م، أو ثلاثة شوارع أحدها ≥ 20 م.` : `أعرض شارع ${fInt(w)} م: المطلوب ≥ 25 م لموقع على شارعين، أو ≥ 20 م لموقع على ثلاثة شوارع.`, '4.1 جدول (1) ص13');
    }
    // مساحة الأرض للطالب
    if (nz(p.landArea) != null && students) {
      const complex = !p.calc || p.calc.startsWith('مجمع');
      const req = complex ? students * RULE.lpsComplex[0] : kgS * RULE.lpsKG[0] + prS * RULE.lpsPrim[0];
      m.landReq = req; m.capLand = complex ? Math.floor(nz(p.landArea) / RULE.lpsComplex[0]) : null;
      if (nz(p.landArea) < req) spatial('treat', `مساحة الأرض ${fInt(nz(p.landArea))} م² أقل من المطلوب ${fInt(req)} م² لعدد ${fInt(students)} طالباً (${complex ? '4.80 م²/طالب – مجمع' : 'التمهيدي 5.00 والابتدائي 4.40 م²/طالب'}): يلزم خفض عدد الطلاب.`, '4.2 جدول (3) ص14');
      if (complex && kgS) add('info', 'حُسبت المساحة بمعدل المجمع التعليمي، والتعريف (2.1) يستثني مرحلتي الحضانة ورياض الأطفال من مراحل المجمع.', '2.1 ص9');
    }
    // الأدوار وموقع التمهيدي
    const fl = nz(p.floors);
    if (fl != null && fl > RULE.floorsPrim[0]) spatial('fail', `${fInt(fl)} أدوار: الحد الأعلى للمبنى الابتدائي أرضي + 3 (أي 4 أدوار)، ما لم يكن الدور الإضافي غير تعليمي ومسموحاً بنظام البناء بعد موافقة الجهة المشرفة.`, '4.3 جدول (4) وبند 3 ص15');
    if (kgS || kgC) {
      if (p.kgFloor === 'الدور الأول بدرج فقط' || p.kgFloor === 'أعلى من الدور الأول') add('fail', 'فصول التمهيدي: الدور الأول مسموح فقط بمنحدر، ويمنع استخدام الدرج للأطفال نهائياً، ولا تتجاوز رياض الأطفال أرضي + 1.', '5.2 بند 5 ص18 و4.3 جدول (4)');
      else if (!p.kgFloor) add('miss', 'موقع فصول التمهيدي غير محدد.', '5.2 بند 5 ص18');
      if (p.fuel50 === 'لا') spatial('fail', 'التمهيدي: المسافة لأقرب محطة وقود أو محل غاز أقل من 50 م.', '4.1 بند 11 ص14');
      if (p.inter50 === 'لا') spatial('fail', 'التمهيدي: المسافة لتقاطع الشوارع التجارية والرئيسة أقل من 50 م.', '4.1 بند 12 ص14');
    }
    // الفصول والطاقة
    const kgSize = NM('KG_ClassSize'), prSize = NM('Prim_ClassSize');
    if (kgS && kgC && kgS > kgC * kgSize) add('treat', `طلاب التمهيدي ${fInt(kgS)} أكثر من ${fInt(kgC)} فصول × ${fInt(kgSize)} (سعة الفصل في افتراضات الدراسة).`, 'افتراض الدراسة – المساحات الداخلية تعتمدها وزارة التعليم (5.2 بند 1)');
    if (prS && prC && prS > prC * prSize) add('treat', `طلاب الابتدائي ${fInt(prS)} أكثر من ${fInt(prC)} فصلاً × ${fInt(prSize)} (سعة الفصل في افتراضات الدراسة).`, 'افتراض الدراسة – المساحات الداخلية تعتمدها وزارة التعليم (5.2 بند 1)');
    // المواقف والحافلات
    m.parkReq = Math.ceil(kgC / RULE.parkKGper[0]) + prC * RULE.parkPrim[0]; m.busReq = Math.ceil(classes / RULE.busPer[0]);
    if (nz(p.parking) != null && classes && nz(p.parking) < m.parkReq) spatial('treat', `المواقف ${fInt(nz(p.parking))} أقل من الحد الأدنى ${fInt(m.parkReq)} (التمهيدي موقف/3 فصول، الابتدائي موقفان/فصل) داخل حدود الملكية.`, '4.1 جدول (2) ص13 و4.4 بند 3');
    if (nz(p.bus) != null && classes && nz(p.bus) < m.busReq) spatial('treat', `مواقف الحافلات ${fInt(nz(p.bus))} أقل من الحد الأدنى ${fInt(m.busReq)} (موقف لكل 3 فصول).`, '4.1 جدول (2) ص13');
    // الأسوار والساحات والوصول
    if (nz(p.fenceH) != null && nz(p.fenceH) < RULE.fence[0]) spatial('treat', `ارتفاع السور ${nf2.format(nz(p.fenceH))} م أقل من 2.40 م.`, '4.4 بند 6 ص16');
    if (p.girlsSolid === 'لا') spatial('treat', 'سور قسم البنات يجب أن يكون مصمتاً.', '4.4 بند 6 ص16');
    if (p.innerFence === 'لا') spatial('treat', 'يلزم فصل حركة السيارات والحافلات عن مساحة الطلاب بأسوار نافذة داخلية.', '4.4 بند 4 ص16');
    if (nz(p.shadePct) != null && nz(p.shadePct) < RULE.shade[0] * 100) add('treat', `تغطية الساحات ${fInt(nz(p.shadePct))}% أقل من 75%.`, '5.2 بند 9 ص18');
    if (p.accPark === 'لا') spatial('treat', 'يلزم توفير مواقف لذوي الإعاقة قرب المداخل الرئيسية (SBC 201).', '4.4 بند 2 ص16');
    if (p.ramp === 'لا') add('treat', 'يجب توفير منحدر لذوي الإعاقة ضمن حدود المبنى وفق SBC 201.', '5.4 بنود 3 و4 ص19');
    if (p.setbacks === 'لا') spatial('treat', 'المبنى غير مطابق لرخصة البناء أو الارتدادات النظامية.', '4.1 جدول (1) ص13');
    if (p.traffic === 'لا') add('treat', 'تلزم دراسة تأثير مروري معتمدة (المداخل والمخارج وفقها).', '4.5 ص16 و4.1 بند 9');
    // السلامة
    if (p.asbestos === 'لا') add('fail', 'يمنع استخدام الأسبستوس أو أي مادة يدخل في تركيبها.', '5.1 بند 4 ص18');
    if (p.facade === 'لا') add('treat', 'يمنع تركيب التمديدات ووحدات التكييف الخارجية على الواجهة أو بشكل ظاهر على السطح.', '5.3 بند 2 ص19');
    if (p.cameras === 'لا') add('treat', 'يجب تركيب كاميرات مراقبة أمنية وفق النظام.', '6.1 بند 10 ص22');
    // المستندات
    const docs = p.docs || {}, missDocs = DOCS.filter(([k]) => !docs[k]).map(([, l]) => l);
    if (missDocs.length) add('treat', 'مستندات وتراخيص غير متوفرة بعد: ' + missDocs.join('، '), '3.1 ص11 واللائحة المادة 6 فقرة 4');
    m.docs = DOCS.length - missDocs.length;
    const has = l => out.some(x => x.lvl === l);
    return { status: has('fail') ? 'fail' : has('miss') ? 'miss' : has('treat') ? 'treat' : 'ok', reasons: out, m, exempt };
  }
  const STATUS = { ok: ['good', 'مناسب مبدئياً'], treat: ['mid', 'يحتاج معالجة'], fail: ['bad', 'غير مناسب وفق المعلومات المتاحة'], miss: ['mid', 'غير مكتمل البيانات'] };
  const stTag = st => `<span class="tag ${st === 'ok' ? 'c-ok' : st === 'fail' ? 'c-neg' : st === 'miss' ? 'c-est' : 'c-ver'}">${STATUS[st][1]}</span>`;
  function propsSummary() {
    const l = PROP.list(), ev = l.map(p => ({ p, e: evalProp(p) })), cnt = st => ev.filter(x => x.e.status === st).length;
    return { n: l.length, ok: cnt('ok'), treat: cnt('treat'), fail: cnt('fail'), miss: cnt('miss') };
  }

  let propOpen = null;
  function pProperty() {
    const list = PROP.list(), ev = list.map(p => evalProp(p));
    const field = (p, [f, l, t]) => {
      if (f === 'g') return `<div class="pg">${esc(l)}</div>`;
      const v = p[f] == null ? '' : p[f];
      const ctl = Array.isArray(t) ? `<select data-pid="${p.id}" data-f="${f}">${t.map(o => `<option${String(v) === o ? ' selected' : ''} value="${esc(o)}">${o ? esc(o) : '—'}</option>`).join('')}</select>`
        : `<input data-pid="${p.id}" data-f="${f}" ${t === 'num' ? 'inputmode="decimal"' : ''} value="${esc(v)}">`;
      return `<label class="pf"><span>${esc(l)}</span>${ctl}</label>`;
    };
    const LV = { fail: 'يمنع', treat: 'يحتاج معالجة', miss: 'ناقص', info: 'ملاحظة', ok: 'مستوفى' };
    const card = (p, i) => { const e = ev[i], st = STATUS[e.status]; return `<div class="card pcard">
      <div class="pc-head"><h3>${esc(p.name || 'عقار ' + (i + 1))}</h3>${stTag(e.status)}
        <span style="margin-inline-start:auto;display:flex;gap:6px" class="no-print"><button class="btn" data-pa="toggle" data-pid="${p.id}">${propOpen === p.id ? 'إخفاء البيانات' : 'تعديل البيانات'}</button><button class="btn" data-pa="dup" data-pid="${p.id}">نسخ</button><button class="btn" data-pa="del" data-pid="${p.id}">حذف</button></span></div>
      ${propOpen === p.id ? `<div class="pform">${FIELDS.map(f => field(p, f)).join('')}</div>
        <div class="pdocs"><b>المستندات والتراخيص المتوفرة:</b>${DOCS.map(([k, l]) => `<label><input type="checkbox" data-pid="${p.id}" data-f="docs.${k}"${p.docs && p.docs[k] ? ' checked' : ''}> ${esc(l)}</label>`).join('')}</div>` : ''}
      <div class="verdict ${st[0]}" style="margin-top:10px"><div class="ic">${e.status === 'ok' ? '✓' : '!'}</div><div><h4>${st[1]} <span class="small muted">(فحص مبدئي – لا يغني عن الترخيص البلدي وترخيص وزارة التعليم واعتماد الدفاع المدني)</span></h4>
        <ul class="list small" style="margin:6px 0 0">${e.reasons.map(r => `<li><b class="rl-${r.lvl}">${LV[r.lvl]}:</b> ${esc(r.text)}${r.ref ? ` <span class="muted">[${esc(r.ref)}]</span>` : ''}</li>`).join('') || '<li>البنود المدخلة مستوفية للمتطلبات المفحوصة.</li>'}</ul></div></div></div>`; };
    const cmpRows = [['النتيجة', (p, e) => stTag(e.status)], ['استعمال الأرض', p => esc(p.landUse || '—')], ['مستقل', p => esc(p.independent || '—')], ['مبنى تعليمي قائم مستثنى', (p, e) => e.exempt ? 'نعم' : 'لا'], ['الشوارع / أعرض شارع (م)', p => `${p.streets ?? '—'} / ${p.w1 ?? '—'}`],
      ['مساحة الأرض / البناء (م²)', p => `${p.landArea ?? '—'} / ${p.bldArea ?? '—'}`], ['الأدوار', p => esc(p.floors ?? '—')], ['الفصول (تمهيدي / ابتدائي)', p => `${p.kgC || 0} / ${p.prC || 0}`], ['الطلاب', (p, e) => fInt(e.m.students)],
      ['الأرض المطلوبة للطلاب (م²)', (p, e) => e.m.landReq != null ? fInt(e.m.landReq) : '—'], ['المواقف: متوفر / مطلوب', (p, e) => `${p.parking ?? '—'} / ${fInt(e.m.parkReq)}`], ['الحافلات: متوفر / مطلوب', (p, e) => `${p.bus ?? '—'} / ${fInt(e.m.busReq)}`], ['المستندات المتوفرة', (p, e) => `${e.m.docs} من ${DOCS.length}`]];
    const cmp = `<div class="tbl-wrap"><table class="t"><thead><tr><th>البند</th>${list.map((p, i) => `<th>${esc(p.name || 'عقار ' + (i + 1))}</th>`).join('')}</tr></thead><tbody>${cmpRows.map(([l, f]) => `<tr><td><b>${l}</b></td>${list.map((p, i) => `<td>${f(p, ev[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
    const crit = TOPICS.map((t, ti) => { const rs = REQS.filter(r => r[0] === ti); return `<details class="sx"><summary>${esc(t)}<span class="s">${fInt(rs.length)} متطلباً</span></summary><div class="body"><ul class="list small">${rs.map(r => `<li>${esc(r[1])} ${natTag(r[5])} <span class="muted">[${esc(r[2] === 'A' ? 'الاشتراطات' : 'اللائحة')} – ${esc(r[3])}، ص${fInt(r[4])}]</span><div class="muted">${esc(r[6])}</div></li>`).join('')}</ul></div></details>`; }).join('');
    return head('القسم 5', 'معايير المباني التعليمية', 'معايير عملية لفحص المباني القائمة المرشحة للاستئجار وتجهيزها للاستخدام التعليمي، مستخلصة من «اشتراطات المباني التعليمية الخاصة» و«اللائحة المنظمة للمؤسسات التعليمية الخاصة»، للمرحلتين التمهيدية والابتدائية في الرياض (الفئة أ). النصوص الكاملة ومراجعها في قسم <a href="#regulations">اشتراطات المباني التعليمية</a>.') +
      `<div class="card"><h3>أداة فحص العقارات المرشحة</h3>
        <p class="sub">أدخل بيانات كل عقار لمقارنته بالمتطلبات ومعرفة مدى ملاءمته مبدئياً. تُحفظ البيانات في هذا المتصفح وعلى هذا الجهاز فقط، ولا تُزامَن بين الأجهزة؛ استخدم «تصدير» لنقلها.</p>
        <div style="display:flex;gap:8px;flex-wrap:wrap" class="no-print"><button class="btn pri" data-pa="add">+ إضافة عقار</button><button class="btn" data-pa="export">تصدير العقارات (JSON)</button><label class="btn" style="cursor:pointer">استيراد<input type="file" accept=".json,application/json" data-pa-import hidden></label></div>
        <div class="legend" style="margin-top:12px"><b>النتيجة:</b>${['ok', 'treat', 'fail', 'miss'].map(stTag).join('')}</div>
        <p class="small muted" style="margin:8px 0 0">«غير مناسب» عند مخالفة متطلب إلزامي لا يعالَج بالتأهيل (الاستعمال، الاستقلال، الشوارع، الأدوار، موقع التمهيدي، المسافات، الأسبستوس) · «يحتاج معالجة» عند نقص يمكن معالجته أو يتطلب موافقة جهة · «غير مكتمل» عند نقص بيانات أساسية.</p></div>
      <div style="height:16px"></div>
      <div class="stack">${list.map(card).join('') || '<div class="card"><p class="muted">لا توجد عقارات بعد. اضغط «إضافة عقار».</p></div>'}</div>
      ${list.length > 1 ? `<div style="height:16px"></div><div class="card"><h3>مقارنة العقارات</h3>${cmp}</div>` : ''}
      <div style="height:16px"></div>
      <div class="card"><h3>المعايير حسب الموضوع</h3><p class="sub">${fInt(REQS.length)} متطلباً مصنفة: ${['m', 'a', 'g'].map(natTag).join(' ')}</p><div class="stack">${crit}</div></div>
      <div style="height:16px"></div>
      <div class="card"><h3>القيم المستخدمة في الفحص الآلي</h3><div class="tbl-wrap"><table class="t"><thead><tr><th>القيمة</th><th>المقدار</th><th>المرجع</th></tr></thead><tbody>
        <tr><td>نصيب الطالب من الأرض – رياض الأطفال / الابتدائي / المجمع</td><td class="n">5.00 / 4.40 / 4.80 م²</td><td class="small">الاشتراطات ${RULE.lpsKG[1]}</td></tr>
        <tr><td>المواقف – رياض الأطفال / الابتدائي / الحافلات</td><td>موقف/3 فصول · موقفان/فصل · موقف/3 فصول</td><td class="small">الاشتراطات ${RULE.parkPrim[1]}</td></tr>
        <tr><td>الأدوار – رياض الأطفال / الابتدائي</td><td>أرضي + 1 · أرضي + 3</td><td class="small">الاشتراطات ${RULE.floorsPrim[1]}</td></tr>
        <tr><td>الشوارع</td><td>شارعان أحدهما ≥ 25 م، أو ثلاثة أحدها ≥ 20 م</td><td class="small">الاشتراطات 4.1 جدول (1) ص13</td></tr>
        <tr><td>السور / تغطية الساحات</td><td class="n">2.40 م · 75%</td><td class="small">الاشتراطات 4.4 بند 6 · 5.2 بند 9</td></tr>
        <tr><td>المسافات للتمهيدي (محطات الوقود والغاز، التقاطعات)</td><td class="n">50 م</td><td class="small">الاشتراطات 4.1 بندا 11 و12 ص14</td></tr>
        <tr><td>سعة الفصل (تمهيدي / ابتدائي)</td><td class="n">${fInt(NM('KG_ClassSize'))} / ${fInt(NM('Prim_ClassSize'))}</td><td class="small">افتراض الدراسة (ورقة 03)؛ المساحات الداخلية تحددها وزارة التعليم (5.2 بند 1)</td></tr>
      </tbody></table></div></div>`;
  }

  /* أحداث أداة العقارات وتبويب الاشتراطات */
  const newProp = () => ({ id: 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), name: '', docs: {}, calc: 'مجمع تعليمي (4.80)' });
  document.addEventListener('change', e => {
    const el = e.target;
    if (el.dataset && el.dataset.pid && el.dataset.f) {
      const l = PROP.list(), p = l.find(x => x.id === el.dataset.pid); if (!p) return;
      const f = el.dataset.f;
      if (f.startsWith('docs.')) { p.docs = p.docs || {}; p.docs[f.slice(5)] = el.checked; }
      else { const fd = FIELDS.find(x => x[0] === f); let v = el.value.trim(); if (fd && fd[2] === 'num') { v = v === '' ? null : Number(v.replace(/,/g, '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace('٫', '.')); if (v !== null && !isFinite(v)) { toast('أدخل رقماً صحيحاً'); return; } } p[f] = v; }
      PROP.save(l); render(current, true); return;
    }
    if (el.hasAttribute && el.hasAttribute('data-pa-import') && el.files[0]) {
      const r = new FileReader(); r.onload = () => { try { const d = JSON.parse(r.result); const arr = Array.isArray(d) ? d : d.properties; if (!Array.isArray(arr)) throw 0; PROP.save(arr); render(current, true); toast('استُوردت العقارات (' + arr.length + ')'); } catch (x) { alert('ملف غير صالح'); } }; r.readAsText(el.files[0]);
    }
  });
  document.addEventListener('click', e => {
    const rt = e.target.closest('[data-regtopic]'); if (rt) { regTopic = +rt.dataset.regtopic; render(current, true); return; }
    const t = e.target.closest('[data-pa]'); if (!t) return;
    const a = t.dataset.pa, l = PROP.list();
    if (a === 'add') { const p = newProp(); p.name = 'عقار ' + (l.length + 1); l.push(p); PROP.save(l); propOpen = p.id; }
    else if (a === 'toggle') propOpen = propOpen === t.dataset.pid ? null : t.dataset.pid;
    else if (a === 'dup') { const p = l.find(x => x.id === t.dataset.pid); if (p) { l.push({ ...JSON.parse(JSON.stringify(p)), id: newProp().id, name: (p.name || 'عقار') + ' (نسخة)' }); PROP.save(l); } }
    else if (a === 'del') { const i = l.findIndex(x => x.id === t.dataset.pid); if (i >= 0) { l.splice(i, 1); PROP.save(l); } }
    else if (a === 'export') { const data = { app: 'growth-ladder-properties', exported: new Date().toISOString(), properties: l }; const x = document.createElement('a'); x.href = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' })); x.download = 'سلم_النمو_العقارات_' + new Date().toISOString().slice(0, 10) + '.json'; x.click(); return; }
    render(current, true);
  });

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
    $('#drawer .body').innerHTML = `<div class="note info small">عدّل أي افتراض وستتحدث كل الصفحات والرسوم فوراً. القيم الأصلية من دراسة v2 تبقى محفوظة ويمكن استعادتها بزر «القيم الأصلية». بقية المدخلات (الفصول، الكادر، المدد، المخاطر…) قابلة للتعديل داخل جداول كل قسم (الخلايا الصفراء).</div>` + drawerHTML();
    $('#drawer').classList.add('open'); $('#scrim').classList.add('show');
  }
  function closeDrawer() { $('#drawer').classList.remove('open'); $('#scrim').classList.remove('show'); $('#side').classList.remove('open'); }

  /* ================= التوجيه والعرض ================= */
  let current = 'home';
  function nav() {
    $('#nav').innerHTML = PAGES.map((p, i) => `<a href="#${p.id}" data-id="${p.id}" class="${p.id === current ? 'on' : ''}"><span class="n">${i + 1}</span>${esc(p.t)}</a>`).join('');
  }
  function render(id, keepScroll) {
    const p = PAGES.find(x => x.id === id) || PAGES[0]; current = p.id; rendering = p.id;
    const y = scrollY; after.length = 0;
    $('#page').innerHTML = p.r();
    after.forEach(f => f());
    $('#crumb').innerHTML = `<button class="btn menu no-print" id="btn-menu" aria-label="القائمة">☰</button>${esc(p.t)}`;
    $$('#nav a').forEach(a => a.classList.toggle('on', a.dataset.id === p.id));
    document.title = p.t + ' – مدارس سُلّم النمو | دراسة الجدوى';
    if (keepScroll) scrollTo(0, y); else scrollTo(0, 0);
  }
  function rerender() { render(current, true); if ($('#drawer').classList.contains('open')) { const b = $('#drawer .body'), st = b.scrollTop; openDrawer(); b.scrollTop = st; } updateChangedBadge(); }

  /* تقرير كامل للطباعة / PDF */
  function printAll() {
    const keep = current; after.length = 0;
    const html = `<div class="cover"><div class="top"><img src="assets/brand/logo_arabic_color.svg" alt="سُلّم النمو"></div>
      <div><div style="color:var(--aqua);font-weight:700">مدرسة أهلية – الرياض</div><h1>الدراسة الفنية والهندسية</h1>
      <p style="font-size:13pt;color:var(--grey);margin:0">مدارس سُلّم النمو | <span class="en">Growth Ladder Schools</span><br>نرتقي بالعلم، وننمو بالمستقبل</p></div>
      <div><div class="band"></div><p class="small muted">النسخة غير المالية (v3) · ${changedCount() ? `<b>يتضمن ${changedCount()} مدخلاً معدلاً عن الدراسة الأصلية</b>` : 'بالمدخلات الأصلية للدراسة'} · طُبع في ${esc(new Date().toLocaleDateString('ar-SA-u-nu-latn-ca-gregory'))}<br>دراسة تقديرية لأغراض القرار المبدئي؛ ليست اعتماداً رسمياً أو هندسياً.</p></div></div>` +
      PAGES.map(p => { rendering = p.id; return `<section class="psec">${p.r()}</section>`; }).join('');
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
    const t = e.target.closest('[data-how],[data-a],[data-engtab],[data-risktab],figure[data-img],#btn-menu,#nav a');
    if (!t) { if (e.target.id === 'modal') closeModal(); return; }
    if (t.matches('#nav a')) { $('#side').classList.remove('open'); $('#scrim').classList.remove('show'); return; }
    if (t.id === 'btn-menu') { $('#side').classList.add('open'); $('#scrim').classList.add('show'); return; }
    if (t.dataset.how) return showHow(t.dataset.how, t.dataset.extra);
    if (t.dataset.engtab) { engTab = t.dataset.engtab; return render(current, true); }
    if (t.dataset.risktab) { riskTab = t.dataset.risktab; return render(current, true); }
    if (t.dataset.img) { $('#lightbox img').src = t.dataset.img; $('#lightbox').classList.add('show'); return; }
    const a = t.dataset.a;
    if (a === 'close') closeModal();
    else if (a === 'inputs') openDrawer();
    else if (a === 'close-drawer') closeDrawer();
    else if (a === 'reset') resetAll();
    else if (a === 'reset-yes') resetYes();
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
      <li>ابدأ من <b>الرئيسية</b>: وضع المشروع وأهم المؤشرات الفنية.</li>
      <li>افحص العقارات المرشحة للاستئجار في قسم <b>معايير المباني التعليمية</b>، وراجع المتطلبات النظامية في قسم <b>اشتراطات المباني التعليمية</b>.</li>
      <li>افتح <b>الافتراضات</b> لتعديل المدخلات الرئيسية، أو عدّل الخلايا الصفراء داخل جداول أي قسم. تتحدث كل الأرقام والرسوم فوراً.</li>
      <li>زر <b>؟</b> بجانب أي مؤشر يعرض معادلته الأصلية ومكانها في ملف الدراسة.</li>
      <li><b>حفظ / استعادة</b>: يحفظ التعديلات في هذا المتصفح فقط، ويمكن تصديرها ملفاً لنقلها لجهاز آخر.</li>
      <li><b>القيم الأصلية</b>: يعيد كل المدخلات إلى قيم الدراسة (النسخة غير المالية).</li>
      <li><b>تصدير PDF</b>: يجهز تقريراً كاملاً بكل الأقسام؛ اختر «حفظ كـ PDF» في نافذة الطباعة.</li></ol>
      <div class="note info small">التعديلات تُحفظ تلقائياً محلياً في هذا المتصفح وعلى هذا الجهاز فقط، ولا تتم مزامنتها بين الأجهزة.</div>
      <div style="margin-top:14px;text-align:left"><button class="btn" data-a="close">إغلاق</button></div>`);
  }

  /* ================= التشغيل ================= */
  loadState(); nav(); updateChangedBadge();
  $('#build').textContent = `المصدر: ${M.source} · ${fInt(Object.keys(M.f).length)} معادلة · ${fInt(Object.keys(M.inp).length)} مدخلاً`;
  render(location.hash.slice(1) || 'home');
  if (new URLSearchParams(location.search).has('print')) setTimeout(printAll, 600);
  window.GL = { E, M, core, render, PAGES, evalProp, PROP }; // للاختبار والتطوير
})();
