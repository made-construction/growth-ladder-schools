// يقارن كل معادلات النموذج المترجم بقيم Excel المعاد حسابها (LibreOffice) المصدّرة إلى JSON
import fs from 'fs'; import vm from 'vm'; import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const [,, modelPath, refPath] = process.argv;
const ctx = { window: {} }; vm.createContext(ctx);
vm.runInContext(fs.readFileSync(modelPath, 'utf8'), ctx);
const M = ctx.window.GL_MODEL;
const { Engine, isErr } = require('../assets/js/engine.js');
const ref = JSON.parse(fs.readFileSync(refPath, 'utf8'));
const E = new Engine(M);
let ok = 0, bad = [], errs = [];
const t0 = Date.now();
for (const k of Object.keys(M.f)) {
  let v; try { v = E.getKey(k); } catch (e) { bad.push([k, 'EXC ' + e.message, ref[k]]); continue; }
  if (isErr(v)) errs.push([k, v.code, M.fx[k]]);
  const r = ref[k];
  let same;
  if (typeof r === 'number' && typeof v === 'number') same = Math.abs(r - v) <= 1e-6 * Math.max(1, Math.abs(r));
  else if (typeof r === 'boolean') same = v === r;
  else if (r === null || r === undefined || r === '') same = v === '' || v === 0 || v === null;
  else if (typeof r === 'string' && r.startsWith('#')) same = isErr(v) && v.code === r;
  else same = String(v) === String(r);
  if (same) ok++; else bad.push([k, v && v.code ? v.code : v, r, M.fx[k]]);
}
console.log(`formulas=${Object.keys(M.f).length} match=${ok} mismatch=${bad.length} errorsInModel=${errs.length} time=${Date.now()-t0}ms`);
for (const b of bad.slice(0, 40)) console.log('MISMATCH', JSON.stringify(b));
for (const e of errs.slice(0, 10)) console.log('ERR', JSON.stringify(e));
process.exit(bad.length ? 1 : 0);
