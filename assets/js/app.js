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
    { id: 'property', t: 'معايير اختيار وتقييم المبنى المدرسي المستأجر', r: pProperty },
    { id: 'capacity', t: 'الطاقة الاستيعابية', r: pCapacity },
    { id: 'timeline', t: 'الجدول الزمني للتنفيذ', r: pTimeline },
    { id: 'risks', t: 'المخاطر والافتراضات', r: pRisks },
    { id: 'refs', t: 'المراجع والاعتمادات', r: pRefs },
    { id: 'identity', t: 'الهوية المؤسسية والمراجع البصرية', r: pIdentity },
  ];

  const head = (eyebrow, title, desc) => `<div class="ph"><div class="eyebrow">${esc(eyebrow)}</div><h1>${esc(title)}</h1>${desc ? `<p>${desc}</p>` : ''}</div>`;
  const chartBox = (id, title, sub, cls = '') => `<div class="card ${cls}"><h3>${esc(title)}</h3>${sub ? `<p class="sub">${sub}</p>` : ''}<div id="${id}"></div></div>`;
  const after = []; // رسوم تُرسم بعد إدراج HTML
  const draw = (id, fn) => after.push(() => { const el = document.getElementById(id); if (el) fn(el); });

  /* وضع المشروع بعد اعتماد خيار المبنى المستأجر (بيانات المستثمر المعتمدة – نوفمبر 2026) */
  function statusBox() {
    return `<div class="verdict mid"><div class="ic">!</div><div>
      <h4>وضع المشروع: مرحلة اختيار المبنى المستأجر وتقييمه</h4>
      <p>المعتمد الآن: مدرسة أهلية بمنهج وطني في الرياض (تمهيدي + ابتدائي) بتصنيف «مجمع تعليمي»، في <b>مبنى قائم يُستأجر ويُؤهَّل</b>، ولا توجد نية لبناء مبنى جديد. الدراسة الهندسية السابقة (مبنى جديد على أرض 50×50م) تبقى <b>مرجعاً</b> للبرنامج الوظيفي والطاقة والمتطلبات. الخطوة التالية: تقييم العقارات المرشحة في قسم «معايير اختيار وتقييم المبنى المدرسي المستأجر» بعد التحقق من الاشتراطات السارية.</p>
      <p style="margin-top:8px"><a class="btn pri no-print" href="#property">تقييم العقارات المرشحة ←</a></p></div></div>`;
  }
  const refBanner = (what) => `<div class="note" style="margin-bottom:14px"><b>مرجعي – خيار سابق:</b> ${what} أُعدّ لخيار <b>مبنى جديد على أرض 50×50م</b>، وهو خيار لم يعد معتمداً بعد قرار استئجار مبنى قائم. يُعرض مرجعاً للبرنامج الوظيفي والمتطلبات فقط. قيم الاشتراطات المذكورة من سجل الدراسة v2 ولم يُتحقق منها من النص الرسمي الكامل الساري.</div>`;

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
    <div class="hero"><div class="eyebrow">الدراسة الفنية والهندسية · خيار المبنى المستأجر</div>
      <h1>مدارس سُلّم النمو <span class="en" style="font-weight:400;opacity:.8;font-size:.6em">Growth Ladder Schools</span></h1>
      <p>مدرسة أهلية بالمنهج الوطني في الرياض: تمهيدي وابتدائي (الصفوف 1–6)، بقسمين منفصلين للبنين والبنات، بتصنيف «مجمع تعليمي»، في مبنى قائم سيتم استئجاره وتحويله إلى مدرسة.</p>
      <div class="meta"><span class="pill">نرتقي بالعلم، وننمو بالمستقبل</span><span class="pill">مجمع تعليمي</span><span class="pill">مبنى قائم مستأجر</span><span class="pill">الرياض</span></div></div>
    ${statusBox()}
    <div style="height:16px"></div>
    <div class="grid g4">
      ${kpi({ label: 'العقارات المقيّمة', value: numSpan(fInt(P.n)), unit: 'عقار', foot: P.n ? `مناسب مبدئياً ${numSpan(fInt(P.ok))} · يحتاج معالجة ${numSpan(fInt(P.treat))} · غير مناسب ${numSpan(fInt(P.fail))} · غير مكتمل ${numSpan(fInt(P.miss))}` : 'لم يُضف أي عقار بعد – <a href="#property">أضف عقاراً</a>', tone: 'acc' })}
      ${kpi({ label: 'الطاقة الاستيعابية المقترحة (المرجعية)', value: numSpan(fInt(c.capProp)), unit: 'طالب', foot: `تصميمية ${numSpan(fInt(c.capTheory))} · حسب الأرض ${numSpan(fInt(c.capLand))} · المعتمدة: ${c.capOff === null || c.capOff === '' ? 'لم تصدر' : fInt(c.capOff)}`, key: 'Cap_Proposed' })}
      ${kpi({ label: 'عدد الفصول (البرنامج المرجعي)', value: numSpan(fInt(c.classes)), unit: 'فصل', foot: `تمهيدي ${numSpan(fInt(c.kgRooms))} + ابتدائي ${numSpan(fInt(c.primRooms))}`, key: 'Classrooms_Avail' })}
      ${kpi({ label: 'المقاعد المخططة (السنة 5)', value: numSpan(fInt(c.maxSeats)), unit: 'مقعد', foot: 'من خطة الافتتاح التدريجي (ورقة 23)', key: 'Max_Seats' })}
      ${kpi({ label: 'مواقف السيارات المطلوبة (المرجعية)', value: numSpan(fInt(c.parkReq)), unit: 'موقف', foot: 'على 22 فصلاً – القيمة المرجعية من سجل v2 غير متحقق منها', key: 'Parking_Req' })}
      ${kpi({ label: 'مواقف الحافلات المطلوبة (المرجعية)', value: numSpan(fInt(c.busReq)), unit: 'موقف', foot: bus ? `في التصور السابق: المرسوم ${fInt(c.busDrawn)} فقط` : '', key: 'Bus_Req', tone: bus ? 'neg' : '' })}
      ${kpi({ label: 'البنود الحرجة المفتوحة', value: numSpan(fInt(c.critOpen)), unit: 'بند', foot: 'ورقة 01أ – بعد حذف البندين الماليين', key: 'Crit_Open', tone: 'neg' })}
      ${kpi({ label: 'التصنيف المعتمد', value: '<span style="font-size:20px">مجمع تعليمي</span>', foot: `نصيب الطالب المستخدم ${numSpan(nf2.format(N('11', 'C7')))} م² – يحتاج تحقق`, key: 'Land_Basis' })}
    </div>
    <div style="height:16px"></div>
    <div class="grid g2">
      ${chartBox('ch-home-cap', 'قيود الطاقة الاستيعابية (البرنامج المرجعي)', 'أدنى القيود والخطة؛ تُستبدل بقيم العقار المختار بعد تقييمه')}
      ${chartBox('ch-home-seats', 'خطة الافتتاح التدريجي – المقاعد', 'المقاعد المتاحة حسب المرحلة لكل سنة (ورقة 23)')}
    </div>
    <div style="height:16px"></div>
    <div class="grid g2">
      <div class="card"><h3>ما يجب حسمه قبل القرار</h3><p class="sub">من ورقة «01أ الملاحظات الحرجة» – ${fInt(c.critOpen)} بنداً مفتوحاً</p>${critList(6)}<p style="margin:10px 0 0"><a href="#risks">عرض كل البنود الحرجة ←</a></p></div>
      <div class="card"><h3>تقييم العقارات المرشحة</h3><p class="sub">أداة استرشادية لمقارنة المباني المعروضة للإيجار</p>${P.n ? P.table : '<p class="muted small">أضف العقارات المرشحة وبياناتها لتظهر نتيجتها هنا.</p>'}<p style="margin:10px 0 0"><a href="#property">فتح أداة التقييم ←</a></p></div>
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
    return head('القسم 2', 'الملخص التنفيذي', 'نص الملخص مرتبط بالنموذج ويتحدث تلقائياً مع المدخلات.') + statusBox() + `<div style="height:16px"></div>
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
    const tabs = [['04', 'الأرض والموقع'], ['05', 'الاشتراطات النظامية'], ['06', 'البرنامج الوظيفي'], ['07', 'التصور المعماري'], ['08', 'الإنشائية'], ['09', 'الكهروميكانيكية'], ['10', 'السلامة والوصول'], ['30', 'مراجعة المكتب الهندسي']];
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
      `<div class="note info" style="margin-bottom:14px"><b>التصنيف المعتمد: مجمع تعليمي.</b> نصيب الطالب المستخدم ${nf2.format(N('11', 'C7'))} م² من سجل v2 (غير متحقق منه؛ الخبر الرسمي يذكر حتى 5 م² للمجمعات). ويذكر سجل v2 أن تعريف المجمع يستثني الحضانة ورياض الأطفال – يحتاج تأكيد الأمانة. طاقة العقار المستأجر تُحسب لكل عقار في أداة التقييم.</div>
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
      `${regStatusCard()}
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
      <div class="stack">${sheetBlock('31', { title: 'سجل التعديلات v2 (ورقة 31)', open: true })}${sheetBlock('29', { title: 'المصادر والمراجع (ورقة 29)' })}${sheetBlock('28', { title: 'المستندات المطلوبة (ورقة 28)' })}${sheetBlock('05', { title: 'سجل الاشتراطات النظامية (ورقة 05)' })}</div>`;
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

  /* ================= 5) معايير اختيار وتقييم المبنى المدرسي المستأجر (قسم جديد) =================
   * أداة استرشادية: لا تُعد موافقة أو ترخيصاً. القيم المرجعية التنظيمية منقولة من سجل الاشتراطات في الدراسة v2
   * (ورقة 05) وهي غير متحقق منها من النص الرسمي الساري؛ يمكن تعديلها وتعليمها «تم التحقق» بعد مراجعة النص. */
  const PROP_KEY = 'gl-properties-v1', REF_KEY = 'gl-property-refs-v1';
  const CAT = {
    reg: ['c-ver', 'اشتراط نظامي – بانتظار التحقق من النص الساري'],
    lic: ['c-inv', 'متطلب ترخيص وتشغيل'],
    eng: ['c-as', 'معيار هندسي وفني استرشادي'],
    unv: ['c-neg', 'لم يُتحقق منه'],
  };
  const catTag = k => `<span class="tag ${CAT[k][0]}">${CAT[k][1]}</span>`;
  /* القيم المرجعية الافتراضية – المصدر لكل قيمة مذكور، ولا شيء منها مُعلَّم «متحقق منه» */
  const REF_DEFAULT = {
    streetsMin: { v: 2, l: 'أقل عدد للشوارع المحيطة', u: 'شارع', src: 'خبر واس 21/12/2025 + سجل الاشتراطات v2 (ورقة 05)' },
    streetWidthMin: { v: 25, l: 'أقل عرض لأعرض شارع', u: 'م', src: 'خبر واس 21/12/2025 + سجل الاشتراطات v2' },
    lpsComplex: { v: 4.8, l: 'نصيب الطالب من الأرض – مجمع تعليمي', u: 'م²/طالب', src: 'سجل v2 جدول 3 الفئة (أ) – يتعارض مع الخبر الرسمي «حتى 5 م² للمجمعات»' },
    parkKGper: { v: 3, l: 'رياض الأطفال: موقف واحد لكل … فصول', u: 'فصل', src: 'سجل v2 جدول المواقف الفئة (أ)' },
    parkPrim: { v: 2, l: 'الابتدائي: مواقف لكل فصل', u: 'موقف/فصل', src: 'سجل v2 جدول المواقف الفئة (أ)' },
    parkInter: { v: 2, l: 'المتوسط: مواقف لكل فصل', u: 'موقف/فصل', src: 'سجل v2 جدول المواقف الفئة (أ)' },
    busPer: { v: 3, l: 'موقف حافلة لكل … فصول', u: 'فصل', src: 'سجل v2 (جميع الفئات)' },
    floorsMax: { v: 4, l: 'أقصى عدد أدوار تعليمية (أرضي + 3) للابتدائي والمتوسط والمجمع', u: 'دور', src: 'سجل v2 (جدول 4) – الخبر: «حسب كود البناء المعتمد»' },
    kgClass: { v: null, l: 'سعة فصل التمهيدي (افتراض الدراسة)', u: 'طفل', src: 'افتراض تصميمي من ورقة 03 – ليس اشتراطاً', name: 'KG_ClassSize' },
    primClass: { v: null, l: 'سعة فصل الابتدائي (افتراض الدراسة)', u: 'طالب', src: 'افتراض تصميمي من ورقة 03 – ليس اشتراطاً', name: 'Prim_ClassSize' },
    interClass: { v: null, l: 'سعة فصل المتوسط', u: 'طالب', src: 'لا توجد قيمة في الدراسة – أدخلها إن وُجدت مرحلة متوسطة' },
  };
  const ls = (k, d) => { try { const v = JSON.parse(localStorage.getItem(k) || 'null'); return v == null ? d : v; } catch (e) { return d; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { toast('تعذّر الحفظ في هذا المتصفح'); } };
  const PROP = {
    list: () => ls(PROP_KEY, []),
    save: l => lsSet(PROP_KEY, l),
    refs() { const o = ls(REF_KEY, {}); const r = {}; for (const [k, d] of Object.entries(REF_DEFAULT)) { const def = d.name ? NM(d.name) : d.v; r[k] = { ...d, v: o[k] && o[k].v !== undefined ? o[k].v : def, ok: !!(o[k] && o[k].ok) }; } return r; },
    saveRef(k, patch) { const o = ls(REF_KEY, {}); o[k] = { ...(o[k] || {}), ...patch }; lsSet(REF_KEY, o); },
  };
  const DOCS = [['deed', 'صك الملكية / وثيقة العقار'], ['permit', 'رخصة البناء'], ['plan', 'الكروكي / المخطط المعتمد واستعمال الأرض'], ['completion', 'شهادة إتمام البناء'], ['structural', 'تقرير سلامة إنشائية من مكتب معتمد'], ['civil', 'تقرير أو شهادة الدفاع المدني'], ['asbuilt', 'مخططات المبنى القائم (As-built)'], ['lease', 'مسودة عقد الإيجار ومدته']];
  const FIELDS = [
    ['name', 'اسم العقار / الوصف', 'text'], ['district', 'الحي / الموقع', 'text'],
    ['landUse', 'نوع استعمال الأرض', ['', 'تعليمي (مرفق تعليمي)', 'تجاري', 'سكني', 'سكني تجاري / مختلط', 'غير معروف']],
    ['landArea', 'مساحة الأرض (م²)', 'num'], ['bldArea', 'مساحة البناء الإجمالية (م²)', 'num'], ['floors', 'عدد الأدوار فوق الأرض', 'num'],
    ['streets', 'عدد الشوارع المحيطة', 'num'], ['w1', 'عرض الشارع الأعرض (م)', 'num'], ['w2', 'عرض الشارع الثاني (م)', 'num'],
    ['kgC', 'فصول التمهيدي', 'num'], ['prC', 'فصول الابتدائي', 'num'], ['imC', 'فصول المتوسط (إن وُجد)', 'num'],
    ['kgS', 'طلاب التمهيدي المتوقعون', 'num'], ['prS', 'طلاب الابتدائي المتوقعون', 'num'], ['imS', 'طلاب المتوسط المتوقعون (إن وُجد)', 'num'],
    ['parking', 'عدد المواقف داخل العقار', 'num'], ['bus', 'عدد مواقف الحافلات', 'num'],
    ['prevSchool', 'هل سبق استخدامه كمدرسة؟', ['', 'نعم', 'لا', 'غير معروف']], ['independent', 'هل المبنى مستقل بالكامل للمدرسة؟', ['', 'نعم', 'لا', 'غير معروف']],
  ];
  const nz = v => v === '' || v === null || v === undefined ? null : Number(v);

  /* تقييم عقار: يعيد الحالة والأسباب والمؤشرات */
  function evalProp(p, R) {
    R = R || PROP.refs();
    const out = [], add = (lvl, text, cat) => out.push({ lvl, text, cat });
    const kgC = nz(p.kgC) || 0, prC = nz(p.prC) || 0, imC = nz(p.imC) || 0, kgS = nz(p.kgS) || 0, prS = nz(p.prS) || 0, imS = nz(p.imS) || 0;
    const classes = kgC + prC + imC, students = kgS + prS + imS;
    const miss = [];
    if (!p.landUse || p.landUse === 'غير معروف') miss.push('نوع استعمال الأرض');
    if (nz(p.landArea) == null) miss.push('مساحة الأرض');
    if (nz(p.bldArea) == null) miss.push('مساحة البناء');
    if (nz(p.floors) == null) miss.push('عدد الأدوار');
    if (nz(p.streets) == null) miss.push('عدد الشوارع');
    if (nz(p.w1) == null) miss.push('عرض الشارع الأعرض');
    if (!classes) miss.push('عدد الفصول');
    if (!students) miss.push('عدد الطلاب المتوقع');
    if (nz(p.parking) == null) miss.push('عدد المواقف');
    if (nz(p.bus) == null) miss.push('عدد مواقف الحافلات');
    if (miss.length) add('miss', 'بيانات أساسية ناقصة: ' + miss.join('، '), 'eng');
    const m = { classes, students };
    // 1) الاستعمال
    if (p.landUse === 'سكني') add('fail', 'الاستعمال سكني: سجل الاشتراطات v2 يقصر المباني التعليمية على الأراضي المخصصة مرفقاً تعليمياً والتجارية، ويسمح على السكني بالحضانة فقط في مبنى مستقل.', 'reg');
    else if (p.landUse === 'سكني تجاري / مختلط') add('treat', 'استعمال مختلط: يلزم تقرير اشتراطات القطعة من بلدي لتأكيد جواز النشاط التعليمي.', 'reg');
    else if (p.landUse === 'تجاري') add('info', 'استعمال تجاري: مشمول بالاشتراطات وفق ضوابط محددة (خبر واس) – يُتحقق من الضوابط في النص الساري.', 'reg');
    else if (p.landUse && p.landUse.startsWith('تعليمي')) add('info', 'أرض مرفق تعليمي: يُتحقق من أثر قرار مجلس الوزراء 704 (أم القرى، أغسطس 2026) على استفادة المدارس الأهلية من أراضي المرافق التعليمية – لم يُتحقق من نصه.', 'unv');
    // 2) الشوارع
    const wmax = Math.max(nz(p.w1) || 0, nz(p.w2) || 0);
    if (nz(p.streets) != null && nz(p.streets) < R.streetsMin.v) add('fail', `العقار على ${fInt(nz(p.streets))} شارع؛ المرجع: عدد الشوارع لا يقل عن ${fInt(R.streetsMin.v)}.`, 'reg');
    if (nz(p.w1) != null && wmax < R.streetWidthMin.v) add('fail', `أعرض شارع ${fInt(wmax)} م؛ المرجع: لا يقل أحدها عن ${fInt(R.streetWidthMin.v)} م.`, 'reg');
    // 3) نصيب الطالب من الأرض
    if (nz(p.landArea) != null && R.lpsComplex.v > 0) {
      m.capLand = Math.floor(nz(p.landArea) / R.lpsComplex.v);
      if (students > m.capLand) add('treat', `الطلاب المتوقعون ${fInt(students)} يتجاوزون الطاقة حسب الأرض ${fInt(m.capLand)} (${nf2.format(R.lpsComplex.v)} م²/طالب – مجمع تعليمي).`, 'reg');
    }
    // 4) الفصول مقابل الطلاب
    const cap = (s, c, size, label) => { if (!s) return; if (!c) { add('miss', `لا توجد فصول ${label} مدخلة مقابل ${fInt(s)} طالباً.`, 'eng'); return; } if (!(size > 0)) { add('miss', `سعة فصل ${label} غير محددة – أدخلها في القيم المرجعية.`, 'eng'); return; } if (s > c * size) add('treat', `طلاب ${label} ${fInt(s)} أكثر من سعة الفصول ${fInt(c * size)} (${fInt(c)} × ${fInt(size)}).`, 'eng'); };
    cap(kgS, kgC, R.kgClass.v, 'التمهيدي'); cap(prS, prC, R.primClass.v, 'الابتدائي'); cap(imS, imC, R.interClass.v, 'المتوسط');
    m.capRooms = (kgC * (R.kgClass.v || 0)) + (prC * (R.primClass.v || 0)) + (imC * (R.interClass.v || 0));
    // 5) المواقف والحافلات
    m.parkReq = Math.ceil(kgC / R.parkKGper.v) + prC * R.parkPrim.v + imC * R.parkInter.v;
    m.busReq = Math.ceil(classes / R.busPer.v);
    if (nz(p.parking) != null && classes && nz(p.parking) < m.parkReq) add('treat', `المواقف ${fInt(nz(p.parking))} أقل من المرجع ${fInt(m.parkReq)} (داخل حدود الملكية).`, 'reg');
    if (nz(p.bus) != null && classes && nz(p.bus) < m.busReq) add('treat', `مواقف الحافلات ${fInt(nz(p.bus))} أقل من المرجع ${fInt(m.busReq)} (حافلة لكل ${fInt(R.busPer.v)} فصول).`, 'reg');
    // 6) الأدوار
    if (nz(p.floors) != null && nz(p.floors) > R.floorsMax.v) add('treat', `${fInt(nz(p.floors))} أدوار؛ المرجع للأدوار التعليمية ${fInt(R.floorsMax.v)} (أرضي + 3) – الأدوار الزائدة لاستخدام غير تعليمي إن سُمح.`, 'reg');
    if (kgS) add('info', 'التمهيدي: يُتحقق من الدور المسموح له ومن منع الدرج للأطفال (سجل v2: الدور الأرضي، أو الأول بمنحدر).', 'reg');
    // 7) الاستقلال والاستخدام السابق
    if (p.independent === 'لا') add('treat', 'المبنى غير مستقل للمدرسة: لم نتحقق من نص يجيز أو يمنع المدرسة في جزء من مبنى – يلزم تأكيد الجهات.', 'unv');
    else if (p.independent !== 'نعم') add('info', 'استقلال المبنى غير محدد – يلزم تأكيده.', 'unv');
    if (p.prevSchool === 'نعم') add('ok', 'سبق استخدامه كمدرسة: يُتحقق من سريان الترخيص السابق وشهادة الدفاع المدني ومطابقته للاشتراطات الحالية.', 'lic');
    else if (p.prevSchool === 'لا') add('info', 'لم يُستخدم كمدرسة: يلزم تأهيل كامل وتغيير استخدام وتقييم إنشائي للإشغال التعليمي.', 'lic');
    // 8) المراحل
    if (imS || imC) add('info', 'المرحلة المتوسطة خارج المراحل المعتمدة للمشروع (تمهيدي + ابتدائي) – تُقيَّم فقط إن أُضيفت رسمياً.', 'unv');
    if (kgS) add('info', 'تصنيف «مجمع تعليمي» مع التمهيدي: سجل v2 يذكر أن تعريف المجمع يستثني الحضانة ورياض الأطفال – يحتاج تأكيد الأمانة.', 'unv');
    // 9) المستندات
    const docs = p.docs || {}, missingDocs = DOCS.filter(([k]) => !docs[k]).map(([, l]) => l);
    const keyDocs = ['deed', 'permit', 'plan', 'structural', 'civil'].filter(k => !docs[k]).map(k => DOCS.find(d => d[0] === k)[1]);
    if (keyDocs.length) add('treat', 'مستندات أساسية غير متوفرة: ' + keyDocs.join('، '), 'lic');
    m.docs = DOCS.length - missingDocs.length;
    const has = l => out.some(x => x.lvl === l);
    const status = has('fail') ? 'fail' : has('miss') ? 'miss' : has('treat') ? 'treat' : 'ok';
    return { status, reasons: out, m };
  }
  const STATUS = { ok: ['good', 'مناسب مبدئياً'], treat: ['mid', 'يحتاج معالجة'], fail: ['bad', 'غير مناسب وفق المعلومات المتاحة'], miss: ['mid', 'غير مكتمل البيانات'] };
  const stTag = st => `<span class="tag ${st === 'ok' ? 'c-ok' : st === 'fail' ? 'c-neg' : st === 'miss' ? 'c-est' : 'c-ver'}">${STATUS[st][1]}</span>`;
  function propsSummary() {
    const l = PROP.list(), R = PROP.refs(), ev = l.map(p => ({ p, e: evalProp(p, R) }));
    const cnt = st => ev.filter(x => x.e.status === st).length;
    const table = `<div class="tbl-wrap"><table class="t"><thead><tr><th>العقار</th><th>النتيجة (استرشادية)</th></tr></thead><tbody>${ev.map(x => `<tr><td>${esc(x.p.name || 'عقار بلا اسم')}</td><td>${stTag(x.e.status)}</td></tr>`).join('')}</tbody></table></div>`;
    return { n: l.length, ok: cnt('ok'), treat: cnt('treat'), fail: cnt('fail'), miss: cnt('miss'), table };
  }

  /* بطاقة حالة التحقق من المرجع النظامي (تُعرض في هذا القسم وفي المراجع) */
  function regStatusCard() {
    return `<div class="card"><h3>حالة التحقق من الاشتراطات النظامية (نوفمبر 2026)</h3>
      <div class="note bad small" style="margin-bottom:10px"><b>لم يُتحقق من أحدث نسخة سارية من النص الرسمي.</b> صفحة الوزارة <a href="https://momah.gov.sa/ar/node/16243" target="_blank" rel="noopener">momah.gov.sa/ar/node/16243</a> وملف PDF المنشور تعذّر فتحهما من بيئة إعداد الموقع، وكذلك موقع جريدة أم القرى. لذلك كل قيمة نظامية في هذا القسم معلّمة «بانتظار التحقق»، ولا تُعد نتيجة نهائية.</div>
      <table class="t"><tbody>
        <tr><th>الوثيقة</th><td>اشتراطات المباني التعليمية الخاصة 1447هـ – 2025م (وزارة البلديات والإسكان)</td></tr>
        <tr><th>رقم القرار وتاريخه</th><td>4600425374/1 بتاريخ 1447/06/13هـ – كما ورد في سجل الدراسة v2 (ورقة 29)، ولم يُطابق مع النص الرسمي في هذا التحديث</td></tr>
        <tr><th>تاريخ الإعلان</th><td>1 رجب 1447هـ الموافق 21 ديسمبر 2025م – <a href="https://www.spa.gov.sa/N2472560" target="_blank" rel="noopener">خبر وكالة الأنباء السعودية</a> (ظهر في نتائج البحث)</td></tr>
        <tr><th>آخر تحديث للنسخة</th><td>غير متحقق منه – تعذّر فتح صفحة الوزارة</td></tr>
        <tr><th>مستجد يحتاج تحقق</th><td>نتيجة بحث في <a href="https://www.uqn.gov.sa/details?p=27099" target="_blank" rel="noopener">أم القرى (18 أغسطس 2026)</a> تشير إلى قرار مجلس الوزراء رقم 704 بإيقاف العمل بضوابط استفادة المدارس الأهلية من أراضي المرافق التعليمية، وتكليف وزارة التعليم بضوابط جديدة بالتنسيق مع وزارة البلديات والإسكان – لم يُقرأ نص القرار</td></tr>
        <tr><th>تعارض قائم</th><td>سجل v2 (جدول 3 الفئة أ): المجمع 4.80 م²/طالب؛ خبر واس وعكاظ: «من 4 م² لرياض الأطفال حتى 5 م² للمجمعات» – يُحسم من النص الرسمي</td></tr>
        <tr><th>المطلوب</th><td>رفع نسخة PDF من الاشتراطات السارية (أو رابط بديل يمكن فتحه) لمطابقة كل القيم وتحديث هذا القسم</td></tr>
      </tbody></table></div>`;
  }

  const CRITERIA = [
    ['الاستعمال النظامي للعقار', ['reg', 'unv'], 'استعمال الأرض في المخطط المعتمد ورخصة البناء، وجواز النشاط التعليمي على القطعة.', 'سجل v2: تسري الاشتراطات على الأراضي المخصصة مرفقاً تعليمياً والأراضي التجارية، والحضانة فقط على السكني بمبنى مستقل. خبر واس: النطاق يشمل أراضي المرافق التعليمية الخاصة والأراضي التجارية وفق ضوابط. مستجد غير متحقق: قرار مجلس الوزراء 704 (أغسطس 2026).', 'صك، رخصة بناء، تقرير اشتراطات القطعة (بلدي)'],
    ['إمكانية الترخيص التعليمي', ['lic', 'unv'], 'قابلية حصول المبنى على ترخيص وزارة التعليم للمراحل المطلوبة.', 'الترخيص من وزارة التعليم مطلوب (خبر عكاظ). نظام التعليم العام الجديد يسري من 20/1/2027 وتبقى لائحة المدارس الأهلية حتى صدور البدائل (ورقة 29). متطلبات الترخيص الخاصة بالمبنى المستأجر لم يُتحقق منها.', 'خطاب مبدئي من إدارة التعليم الأهلي'],
    ['استقلال المبنى', ['unv'], 'هل المبنى مخصص بالكامل للمدرسة، وهل يوجد مدخل ومخرج مستقلان.', 'لم نجد نصاً متحققاً منه للمدارس الأهلية. الخبر يذكر السماح لمدارس الجاليات بجزء من مبنى قائم بشروط، ولا ينطبق مباشرة على المشروع.', 'عقد الإيجار ومخطط المبنى'],
    ['موقع العقار والشوارع وعروضها', ['reg'], 'عدد الشوارع المحيطة وعرض كل منها ومداخل الطوارئ والحافلات.', 'شارعان لا يقل عرض أحدهما عن 25 م (خبر واس 21/12/2025 وسجل v2) – النص الكامل لم يُراجع.', 'كروكي، رفع مساحي'],
    ['مساحة الأرض ومساحة البناء', ['reg', 'lic'], 'مساحة الأرض حسب الصك، والمسطحات المبنية لكل دور، والفراغات المتاحة.', 'نصيب الطالب من الأرض – سجل v2: المجمع 4.80 م²، والخبر الرسمي: حتى 5 م² للمجمعات (تعارض). المساحات الداخلية تحددها وزارة التعليم (الدليل غير منشور).', 'صك، رخصة بناء، مخططات As-built'],
    ['عدد الأدوار', ['reg'], 'عدد الأدوار فوق الأرض والقبو، ومنسوب الأرضي بالنسبة للشارع.', 'سجل v2 (جدول 4): الابتدائي والمتوسط والمجمع حتى أرضي + 3، ورياض الأطفال أرضي + 1. الخبر: حسب كود البناء المعتمد.', 'رخصة البناء'],
    ['الطاقة الاستيعابية للابتدائي والمتوسط', ['eng', 'unv'], 'أدنى قيمة بين: الأرض ÷ نصيب الطالب، والفصول × سعة الفصل، وما تسمح به المواقف والمرافق.', 'المراحل المعتمدة تمهيدي وابتدائي؛ المتوسط ورد في الطلب ولا يدخل ضمن المراحل المعتمدة – يحتاج قراراً. الطاقة الرسمية تحددها وزارة التعليم.', '—'],
    ['عدد الفصول الممكنة', ['eng', 'unv'], 'عدد الغرف الصالحة للتحويل إلى فصول، ومساحة كل غرفة وإضاءتها وتهويتها.', 'مساحة الفصل وسعته افتراضات الدراسة (التمهيدي 20 والابتدائي 25 طالباً) حتى صدور دليل وزارة التعليم.', 'مخططات As-built'],
    ['المواقف ومواقف الحافلات', ['reg'], 'عدد المواقف داخل حدود الملكية، ومواقف الحافلات ومسار الإنزال.', 'سجل v2: رياض الأطفال موقف لكل 3 فصول، الابتدائي والمتوسط موقفان لكل فصل، وموقف حافلة لكل 3 فصول، داخل حدود الملكية – النص الكامل لم يُراجع.', 'مخطط الموقع'],
    ['الساحات والمرافق التعليمية', ['reg', 'lic'], 'مساحة الساحات وتظليلها، والمصلى والمقصف والمكتبة والمستودع وغرف الإدارة.', 'سجل v2: تغطية الساحات 75% على الأقل مع إضاءة وتهوية طبيعية، والمستودع 100 م² حداً أدنى؛ مساحة الساحات تحددها وزارة التعليم.', 'مخططات As-built'],
    ['المختبرات والمعامل', ['lic', 'eng'], 'إمكانية تخصيص مختبر علوم ومعمل حاسب ومصادر تعلم بالتمديدات اللازمة.', 'حسب البرنامج المعتمد من وزارة التعليم؛ البرنامج المرجعي في ورقة 06.', '—'],
    ['دورات المياه', ['eng', 'unv'], 'العدد والتوزيع لكل قسم ودور، ودورات الكادر وذوي الإعاقة.', 'النسب تحددها الجهة المشرفة؛ قيم الدراسة مبنية على IPC (أساس كود – يحتاج مطابقة SBC).', 'مخططات As-built'],
    ['التكييف والتهوية والإضاءة', ['eng'], 'حالة منظومة التكييف وعمرها وقدرتها، والتهوية والإضاءة الطبيعية للفصول، والقدرة الكهربائية.', 'تقييم فني من مهندس مختص؛ كفاءة الطاقة وفق SBC 601/602 (أساس كود).', 'تقرير فحص الأنظمة'],
    ['السلامة الإنشائية', ['eng', 'lic'], 'سلامة الهيكل وقدرته على أحمال الإشغال التعليمي بعد تغيير الاستخدام.', 'تقرير سلامة إنشائية من مكتب هندسي معتمد؛ مراجعة الأحمال وفق SBC 301 (أساس كود – يحتاج مطابقة).', 'تقرير سلامة إنشائية'],
    ['الحريق ومخارج الطوارئ', ['lic', 'eng'], 'أنظمة الإنذار والإطفاء والرشاشات، وعدد المخارج والأدراج وتباعدها.', 'اعتماد الدفاع المدني مطلوب؛ متطلبات المخارج والرشاشات وفق SBC 201/801 (قيم الدراسة أساس IBC/IFC – تحتاج مطابقة).', 'شهادة الدفاع المدني'],
    ['سهولة الوصول لذوي الإعاقة', ['eng'], 'المداخل والمنحدرات والمصاعد ودورات المياه والمواقف المخصصة.', 'وفق SBC 201 (أساس كود – يحتاج مطابقة).', 'مخططات As-built'],
    ['التعديلات اللازمة قبل التشغيل', ['eng', 'lic'], 'قائمة الأعمال: فصل قسمي البنين والبنات، الأسوار، تجهيز الفصول والساحات والتظليل، المختبرات، دورات المياه، الحريق، الوصول الشامل، اللوحات، والتمهيدي في الدور المسموح.', 'الأسوار في سجل v2: 2.40 م مصمتة لمباني البنات (غير متحقق من النص). تُعتمد قائمة الأعمال بعد زيارة فنية للعقار.', 'تقرير الزيارة الفنية'],
  ];

  let propOpen = null;
  function pProperty() {
    const R = PROP.refs(), list = PROP.list(), ev = list.map(p => evalProp(p, R));
    const unverified = Object.values(R).filter(r => !r.ok && !/افتراض|لا توجد/.test(r.src)).length;
    const field = (p, [f, l, t]) => {
      const v = p[f] == null ? '' : p[f];
      const ctl = Array.isArray(t) ? `<select data-pid="${p.id}" data-f="${f}">${t.map(o => `<option${String(v) === o ? ' selected' : ''} value="${esc(o)}">${o ? esc(o) : '—'}</option>`).join('')}</select>`
        : `<input data-pid="${p.id}" data-f="${f}" ${t === 'num' ? 'inputmode="decimal"' : ''} value="${esc(v)}">`;
      return `<label class="pf"><span>${esc(l)}</span>${ctl}</label>`;
    };
    const card = (p, i) => { const e = ev[i], st = STATUS[e.status]; return `<div class="card pcard">
      <div class="pc-head"><h3>${esc(p.name || 'عقار ' + (i + 1))}</h3>${stTag(e.status)}<span class="small muted">استرشادي</span>
        <span style="margin-inline-start:auto;display:flex;gap:6px" class="no-print"><button class="btn" data-pa="toggle" data-pid="${p.id}">${propOpen === p.id ? 'إخفاء البيانات' : 'تعديل البيانات'}</button><button class="btn" data-pa="dup" data-pid="${p.id}">نسخ</button><button class="btn" data-pa="del" data-pid="${p.id}">حذف</button></span></div>
      ${propOpen === p.id ? `<div class="pform">${FIELDS.map(f => field(p, f)).join('')}</div>
        <div class="pdocs"><b>المستندات المتوفرة:</b>${DOCS.map(([k, l]) => `<label><input type="checkbox" data-pid="${p.id}" data-f="docs.${k}"${p.docs && p.docs[k] ? ' checked' : ''}> ${esc(l)}</label>`).join('')}</div>` : ''}
      <div class="verdict ${st[0]}" style="margin-top:10px"><div class="ic">${e.status === 'ok' ? '✓' : '!'}</div><div><h4>${st[1]} <span class="small muted">(نتيجة استرشادية – ليست موافقة أو ترخيصاً)</span></h4>
        <ul class="list small" style="margin:6px 0 0">${e.reasons.filter(r => r.lvl !== 'ok' || true).map(r => `<li><b class="rl-${r.lvl}">${{ fail: 'يمنع', treat: 'يحتاج معالجة', miss: 'ناقص', info: 'للتحقق', ok: 'إيجابي' }[r.lvl]}:</b> ${esc(r.text)} ${catTag(r.cat)}</li>`).join('') || '<li>لا توجد ملاحظات وفق القيم المرجعية الحالية.</li>'}</ul></div></div></div>`; };
    const cmpRows = [['الحالة', (p, e) => stTag(e.status)], ['استعمال الأرض', p => esc(p.landUse || '—')], ['مساحة الأرض / البناء (م²)', p => `${p.landArea || '—'} / ${p.bldArea || '—'}`], ['الأدوار', p => esc(p.floors || '—')],
      ['الشوارع / أعرض شارع (م)', p => `${p.streets || '—'} / ${Math.max(nz(p.w1) || 0, nz(p.w2) || 0) || '—'}`], ['الفصول (تمهيدي / ابتدائي / متوسط)', p => `${p.kgC || 0} / ${p.prC || 0} / ${p.imC || 0}`],
      ['الطلاب المتوقعون', (p, e) => fInt(e.m.students)], ['الطاقة حسب الأرض', (p, e) => e.m.capLand != null ? fInt(e.m.capLand) : '—'], ['سعة الفصول', (p, e) => fInt(e.m.capRooms)],
      ['المواقف: متوفر / مرجع', (p, e) => `${p.parking ?? '—'} / ${fInt(e.m.parkReq)}`], ['الحافلات: متوفر / مرجع', (p, e) => `${p.bus ?? '—'} / ${fInt(e.m.busReq)}`],
      ['سبق استخدامه كمدرسة', p => esc(p.prevSchool || '—')], ['مبنى مستقل', p => esc(p.independent || '—')], ['المستندات المتوفرة', (p, e) => `${e.m.docs} من ${DOCS.length}`]];
    const cmp = list.length ? `<div class="tbl-wrap"><table class="t"><thead><tr><th>البند</th>${list.map((p, i) => `<th>${esc(p.name || 'عقار ' + (i + 1))}</th>`).join('')}</tr></thead><tbody>${cmpRows.map(([l, f]) => `<tr><td><b>${l}</b></td>${list.map((p, i) => `<td>${f(p, ev[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>` : '';
    const refs = `<div class="tbl-wrap"><table class="t"><thead><tr><th>القيمة المرجعية</th><th>القيمة</th><th>المصدر</th><th>تم التحقق من النص الساري</th></tr></thead><tbody>${Object.entries(R).map(([k, r]) => `<tr><td>${esc(r.l)}</td><td class="in"><input data-ref="${k}" inputmode="decimal" value="${r.v == null ? '' : esc(r.v)}" style="width:90px"> <span class="small muted">${esc(r.u)}</span></td><td class="small">${esc(r.src)}</td><td>${/افتراض|لا توجد/.test(r.src) ? '<span class="small muted">ليس اشتراطاً</span>' : `<label class="small"><input type="checkbox" data-refok="${k}"${r.ok ? ' checked' : ''}> ${r.ok ? '<span class="tag c-ok">متحقق منه</span>' : '<span class="tag c-ver">غير متحقق</span>'}</label>`}</td></tr>`).join('')}</tbody></table></div>`;
    return head('القسم 5', 'معايير اختيار وتقييم المبنى المدرسي المستأجر', 'معايير تقييم المباني القائمة المعروضة للإيجار قبل التعاقد، وأداة لمقارنة العقارات المرشحة. <b>التقييم داخل الموقع استرشادي، ولا يُعد موافقة أو ترخيصاً رسمياً.</b>') +
      `<div class="note bad" style="margin-bottom:16px"><b>تنبيه:</b> لا تُعتمد أي نتيجة نهائية لتقييم عقار قبل التحقق من الاشتراطات السارية من النص الرسمي، واعتماد الجهات (الأمانة/بلدي، وزارة التعليم، الدفاع المدني) ومكتب هندسي معتمد. ${unverified ? `${fInt(unverified)} من القيم المرجعية التنظيمية المستخدمة في الأداة <b>غير متحقق منها</b> بعد.` : 'عُلّمت كل القيم المرجعية بأنها متحقق منها من قبل المستخدم.'}</div>
      ${regStatusCard()}
      <div style="height:16px"></div>
      <div class="card"><h3>تصنيف المعايير</h3><div class="legend">${Object.keys(CAT).map(catTag).join('')}</div>
        <ul class="list small" style="margin-top:10px">
          <li><b>اشتراط نظامي:</b> من اشتراطات المباني التعليمية الخاصة (وزارة البلديات والإسكان) كما وردت في سجل الدراسة v2 والخبر الرسمي – بانتظار المطابقة مع النص الساري.</li>
          <li><b>متطلب ترخيص وتشغيل:</b> ما تطلبه وزارة التعليم والدفاع المدني والأمانة للترخيص والتشغيل، دون أرقام غير موثقة.</li>
          <li><b>معيار هندسي وفني:</b> فحوص فنية مهنية تساعد في تقييم المبنى قبل استئجاره، وليست اشتراطات نظامية.</li>
          <li><b>لم يُتحقق منه:</b> بنود لا يوجد لها مصدر متحقق منه في هذا التحديث.</li></ul></div>
      <div style="height:16px"></div>
      <div class="card"><h3>المعايير السبعة عشر</h3><p class="sub">ما يُفحص في كل عقار، وتصنيفه، والمرجع المتاح، والمستندات المطلوبة</p>
        <div class="tbl-wrap"><table class="t"><thead><tr><th>#</th><th>المعيار</th><th>التصنيف</th><th>ما يُتحقق منه</th><th>المرجع المتاح وحالته</th><th>المستندات</th></tr></thead><tbody>
        ${CRITERIA.map(([t, cats, chk, ref, doc], i) => `<tr><td class="n">${i + 1}</td><td><b>${esc(t)}</b></td><td>${cats.map(catTag).join(' ')}</td><td class="small">${esc(chk)}</td><td class="small">${esc(ref)}</td><td class="small">${esc(doc)}</td></tr>`).join('')}
        </tbody></table></div></div>
      <div style="height:16px"></div>
      <div class="card"><h3>أداة تقييم العقارات المرشحة</h3>
        <p class="sub">أضف كل عقار معروض وبياناته، وتظهر النتيجة وأسبابها فوراً. تُحفظ البيانات <b>في هذا المتصفح وعلى هذا الجهاز فقط</b>، ولا تُزامَن بين الأجهزة؛ استخدم «تصدير» لنقلها.</p>
        <div style="display:flex;gap:8px;flex-wrap:wrap" class="no-print"><button class="btn pri" data-pa="add">+ إضافة عقار</button><button class="btn" data-pa="export">تصدير العقارات (JSON)</button><label class="btn" style="cursor:pointer">استيراد<input type="file" accept=".json,application/json" data-pa-import hidden></label></div>
        <div class="legend" style="margin-top:12px"><b>النتائج:</b>${['ok', 'treat', 'fail', 'miss'].map(stTag).join('')}</div>
      </div>
      <div style="height:16px"></div>
      <div class="stack">${list.map(card).join('') || '<div class="card"><p class="muted">لا توجد عقارات بعد. اضغط «إضافة عقار».</p></div>'}</div>
      ${list.length > 1 ? `<div style="height:16px"></div><div class="card"><h3>مقارنة العقارات</h3>${cmp}</div>` : ''}
      <div style="height:16px"></div>
      <details class="sx"><summary>القيم المرجعية المستخدمة في الأداة<span class="s">قابلة للتعديل بعد مراجعة النص الرسمي</span></summary><div class="body">
        <p class="small muted">كل قيمة تنظيمية هنا منقولة من سجل الاشتراطات في الدراسة v2 أو من الخبر الرسمي، وليست متحققاً منها من النص الساري. بعد مراجعة النص عدّل القيمة إن لزم وعلّم «تم التحقق». الحساب: مواقف التمهيدي = ⌈الفصول ÷ القيمة⌉، والابتدائي والمتوسط = الفصول × القيمة، والحافلات = ⌈إجمالي الفصول ÷ القيمة⌉، والطاقة حسب الأرض = ⌊مساحة الأرض ÷ نصيب الطالب⌋.</p>${refs}</div></details>`;
  }

  /* أحداث أداة العقارات */
  const newProp = () => ({ id: 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), name: '', docs: {} });
  document.addEventListener('change', e => {
    const el = e.target;
    if (el.dataset && el.dataset.pid && el.dataset.f) {
      const l = PROP.list(), p = l.find(x => x.id === el.dataset.pid); if (!p) return;
      const f = el.dataset.f;
      if (f.startsWith('docs.')) { p.docs = p.docs || {}; p.docs[f.slice(5)] = el.checked; }
      else { const fd = FIELDS.find(x => x[0] === f); let v = el.value.trim(); if (fd && fd[2] === 'num') { v = v === '' ? null : Number(v.replace(/,/g, '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))); if (v !== null && !isFinite(v)) { toast('أدخل رقماً صحيحاً'); return; } } p[f] = v; }
      PROP.save(l); render(current, true); return;
    }
    if (el.dataset && el.dataset.ref) { const v = el.value.trim(); PROP.saveRef(el.dataset.ref, { v: v === '' ? null : Number(v) }); render(current, true); return; }
    if (el.dataset && el.dataset.refok) { PROP.saveRef(el.dataset.refok, { ok: el.checked }); render(current, true); return; }
    if (el.hasAttribute && el.hasAttribute('data-pa-import') && el.files[0]) {
      const r = new FileReader(); r.onload = () => { try { const d = JSON.parse(r.result); const arr = Array.isArray(d) ? d : d.properties; if (!Array.isArray(arr)) throw 0; PROP.save(arr); if (d.refs) lsSet(REF_KEY, d.refs); render(current, true); toast('استُوردت العقارات (' + arr.length + ')'); } catch (x) { alert('ملف غير صالح'); } }; r.readAsText(el.files[0]);
    }
  });
  document.addEventListener('click', e => {
    const t = e.target.closest('[data-pa]'); if (!t) return;
    const a = t.dataset.pa, l = PROP.list();
    if (a === 'add') { const p = newProp(); p.name = 'عقار ' + (l.length + 1); l.push(p); PROP.save(l); propOpen = p.id; }
    else if (a === 'toggle') propOpen = propOpen === t.dataset.pid ? null : t.dataset.pid;
    else if (a === 'dup') { const p = l.find(x => x.id === t.dataset.pid); if (p) { const c = { ...JSON.parse(JSON.stringify(p)), id: newProp().id, name: (p.name || 'عقار') + ' (نسخة)' }; l.push(c); PROP.save(l); } }
    else if (a === 'del') { const i = l.findIndex(x => x.id === t.dataset.pid); if (i >= 0) { l.splice(i, 1); PROP.save(l); } }
    else if (a === 'export') { const data = { app: 'growth-ladder-properties', exported: new Date().toISOString(), properties: l, refs: ls(REF_KEY, {}) }; const x = document.createElement('a'); x.href = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' })); x.download = 'سلم_النمو_العقارات_' + new Date().toISOString().slice(0, 10) + '.json'; x.click(); return; }
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
    const p = PAGES.find(x => x.id === id) || PAGES[0]; current = p.id;
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
      <div><div style="color:var(--aqua);font-weight:700">مدرسة أهلية – الرياض</div><h1>الدراسة الفنية والهندسية – خيار المبنى المستأجر</h1>
      <p style="font-size:13pt;color:var(--grey);margin:0">مدارس سُلّم النمو | <span class="en">Growth Ladder Schools</span><br>نرتقي بالعلم، وننمو بالمستقبل</p></div>
      <div><div class="band"></div><p class="small muted">النسخة غير المالية (v3) · ${changedCount() ? `<b>يتضمن ${changedCount()} مدخلاً معدلاً عن الدراسة الأصلية</b>` : 'بالمدخلات الأصلية للدراسة'} · طُبع في ${esc(new Date().toLocaleDateString('ar-SA-u-nu-latn-ca-gregory'))}<br>دراسة تقديرية لأغراض القرار المبدئي؛ ليست اعتماداً رسمياً أو هندسياً.</p></div></div>` +
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
      <li>قيّم العقارات المرشحة للاستئجار في قسم <b>معايير اختيار وتقييم المبنى المدرسي المستأجر</b>؛ النتيجة استرشادية وليست موافقة رسمية.</li>
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
