#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx2js.py – يحوّل ملف دراسة سُلّم النمو (Excel) إلى نموذج JavaScript حي.

كل معادلة في المصنف تُترجم إلى دالة JavaScript تُقيَّم في المتصفح بنفس منطق Excel،
فيتحدث الموقع تلقائياً عند تعديل أي مدخل، دون أي رقم محسوب مكتوب يدوياً.

الاستخدام:
    python tools/xlsx2js.py <ملف_الدراسة.xlsx> assets/js/model.js

المخرجات: assets/js/model.js يعرّف window.GL_MODEL:
    sheets  : أسماء الأوراق بالترتيب
    names   : الأسماء المعرّفة -> مفتاح الخلية "رقم_الورقة!A1"
    v       : القيم الثابتة (نصوص وأرقام وتواريخ كأرقام تسلسلية)
    f       : المعادلات مترجمة إلى دوال
    fx      : نص المعادلة الأصلي (للشفافية وعرض "طريقة الحساب")
    inp     : خلايا الإدخال (نص أزرق على خلفية صفراء) مع قوائم الاختيار
    nf      : تنسيق الأرقام لكل خلية (فهرس في جدول nfs)
    dims    : أبعاد كل ورقة ودمج الخلايا
"""
import sys, re, json, datetime
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

EPOCH = datetime.datetime(1899, 12, 30)


def serial(d):
    if isinstance(d, datetime.datetime):
        return (d - EPOCH).total_seconds() / 86400.0
    if isinstance(d, datetime.date):
        return (datetime.datetime(d.year, d.month, d.day) - EPOCH).days
    return d


# ---------------------------------------------------------------- tokenizer
TOK = re.compile(r"""
 (?P<ws>\s+)
|(?P<str>"(?:[^"]|"")*")
|(?P<qsheet>'(?:[^']|'')+'!)
|(?P<num>\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+)
|(?P<ident>[A-Za-z_\\$][A-Za-z0-9_.$]*)
|(?P<op><>|<=|>=|[-+*/^&=<>%(),:!;{}])
""", re.X | re.U)

CELL = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")


def tokenize(s):
    out, i = [], 0
    while i < len(s):
        m = TOK.match(s, i)
        if not m:
            raise ValueError("tokenize error at %r" % s[i:])
        i = m.end()
        k = m.lastgroup
        if k == 'ws':
            continue
        out.append((k, m.group(k)))
    return out


class Parser:
    def __init__(self, toks, ctx):
        self.t, self.i, self.ctx = toks, 0, ctx

    def peek(self, o=0):
        j = self.i + o
        return self.t[j] if j < len(self.t) else (None, None)

    def take(self, v=None):
        tok = self.peek()
        if v is not None and tok[1] != v:
            raise ValueError("expected %s got %s" % (v, tok))
        self.i += 1
        return tok

    # precedence climbing ------------------------------------------------
    def parse(self):
        e = self.comparison()
        if self.i != len(self.t):
            raise ValueError("trailing tokens %s" % (self.t[self.i:],))
        return e

    def comparison(self):
        a = self.concat()
        while self.peek()[1] in ('=', '<>', '<', '>', '<=', '>='):
            op = self.take()[1]
            b = self.concat()
            a = ('bin', op, a, b)
        return a

    def concat(self):
        a = self.additive()
        while self.peek()[1] == '&':
            self.take()
            b = self.additive()
            a = ('bin', '&', a, b)
        return a

    def additive(self):
        a = self.mult()
        while self.peek()[1] in ('+', '-'):
            op = self.take()[1]
            b = self.mult()
            a = ('bin', op, a, b)
        return a

    def mult(self):
        a = self.power()
        while self.peek()[1] in ('*', '/'):
            op = self.take()[1]
            b = self.power()
            a = ('bin', op, a, b)
        return a

    def power(self):
        a = self.unary()
        while self.peek()[1] == '^':
            self.take()
            b = self.unary()
            a = ('bin', '^', a, b)
        return a

    def unary(self):
        if self.peek()[1] == '-':
            self.take()
            return ('neg', self.unary())
        if self.peek()[1] == '+':
            self.take()
            return self.unary()
        return self.postfix()

    def postfix(self):
        a = self.primary()
        while self.peek()[1] == '%':
            self.take()
            a = ('bin', '/', a, ('num', 100.0))
        return a

    def ref(self, sheet_idx):
        k, v = self.take()
        m = CELL.match(v) if k == 'ident' else None
        if not m:
            raise ValueError("bad ref %r" % v)
        c1, r1 = column_index_from_string(m.group(1)), int(m.group(2))
        if self.peek()[1] == ':':
            self.take()
            k2, v2 = self.take()
            m2 = CELL.match(v2)
            c2, r2 = column_index_from_string(m2.group(1)), int(m2.group(2))
            return ('rng', sheet_idx, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
        return ('cell', sheet_idx, r1, c1)

    def primary(self):
        k, v = self.peek()
        if k == 'num':
            self.take()
            return ('num', float(v))
        if k == 'str':
            self.take()
            return ('str', v[1:-1].replace('""', '"'))
        if k == 'qsheet':
            self.take()
            name = v[1:-2].replace("''", "'")
            return self.ref(self.ctx.sheet_index(name))
        if v == '(':
            self.take()
            e = self.comparison()
            self.take(')')
            return e
        if k == 'ident':
            nxt = self.peek(1)[1]
            if nxt == '(':
                self.take()
                self.take('(')
                args = []
                if self.peek()[1] != ')':
                    while True:
                        if self.peek()[1] in (',', ')'):
                            args.append(('blank',))
                        else:
                            args.append(self.comparison())
                        if self.peek()[1] == ',':
                            self.take()
                            continue
                        break
                self.take(')')
                return ('fn', v.upper(), args)
            if nxt == '!':
                self.take()
                self.take('!')
                return self.ref(self.ctx.sheet_index(v))
            up = v.upper()
            if up in ('TRUE', 'FALSE'):
                self.take()
                return ('bool', up == 'TRUE')
            if v in self.ctx.names:
                self.take()
                return self.ctx.names[v]
            if CELL.match(v):
                return self.ref(self.ctx.cur)
            raise ValueError("unknown identifier %r" % v)
        raise ValueError("unexpected token %r" % (v,))


def key(si, r, c):
    return "%d!%s%d" % (si, get_column_letter(c), r)


LAZY = {'IF', 'IFERROR', 'AND', 'OR'}


def emit(e):
    t = e[0]
    if t == 'num':
        x = e[1]
        return repr(int(x)) if x == int(x) and abs(x) < 1e15 else repr(x)
    if t == 'str':
        return json.dumps(e[1], ensure_ascii=False)
    if t == 'bool':
        return 'true' if e[1] else 'false'
    if t == 'blank':
        return 'null'
    if t == 'cell':
        return 'g(%s)' % json.dumps(key(e[1], e[2], e[3]))
    if t == 'rng':
        return 'R(%d,%d,%d,%d,%d)' % e[1:]
    if t == 'neg':
        return '$.neg(%s)' % emit(e[1])
    if t == 'bin':
        op = {'+': 'add', '-': 'sub', '*': 'mul', '/': 'div', '^': 'pow', '&': 'cat',
              '=': 'eq', '<>': 'ne', '<': 'lt', '>': 'gt', '<=': 'le', '>=': 'ge'}[e[1]]
        return '$.%s(%s,%s)' % (op, emit(e[2]), emit(e[3]))
    if t == 'fn':
        name, args = e[1], e[2]
        fn = name.replace('.', '_')
        if name in LAZY:
            return '$.%s(%s)' % (fn, ','.join('()=>' + emit(a) for a in args))
        return '$.%s(%s)' % (fn, ','.join(emit(a) for a in args))
    raise ValueError(e)


class Ctx:
    def __init__(self, wb):
        self.sheets = wb.sheetnames
        self.cur = 0
        self.names = {}
        for n, d in wb.defined_names.items():
            txt = d.attr_text
            m = re.match(r"^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", txt)
            if not m:
                continue
            si = self.sheet_index(m.group(1).replace("''", "'"))
            c1, r1 = column_index_from_string(m.group(2)), int(m.group(3))
            if m.group(4):
                c2, r2 = column_index_from_string(m.group(4)), int(m.group(5))
                self.names[n] = ('rng', si, r1, c1, r2, c2)
            else:
                self.names[n] = ('cell', si, r1, c1)

    def sheet_index(self, name):
        return self.sheets.index(name)


def is_input(c):
    try:
        fc = (c.font.color.rgb or '') if c.font and c.font.color else ''
        fill = c.fill.fgColor.rgb or ''
    except Exception:
        return False
    return isinstance(fc, str) and fc.endswith('0000FF') and isinstance(fill, str) and fill.endswith('FFF2CC')


def main(src, dst):
    wb = openpyxl.load_workbook(src)
    ctx = Ctx(wb)
    V, F, FX, INP, NF, NFS, DIMS = {}, {}, {}, {}, {}, [], []
    nfi = {}
    errors = []
    for si, ws in enumerate(wb):
        ctx.cur = si
        dv = {}
        for d in ws.data_validations.dataValidation:
            if d.type == 'list' and d.formula1:
                opts = [o.strip() for o in d.formula1.strip('"').split(',')]
                for rng in str(d.sqref).split():
                    for row in ws[rng] if ':' in rng else [[ws[rng]]]:
                        for c in row:
                            dv[c.coordinate] = opts
        widths = {}
        for col, cd in ws.column_dimensions.items():
            if cd.width:
                widths[col] = round(cd.width, 1)
        DIMS.append({'r': ws.max_row, 'c': ws.max_column,
                     'merged': [str(m) for m in ws.merged_cells.ranges], 'w': widths})
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                k = "%d!%s" % (si, c.coordinate)
                fmt = c.number_format or 'General'
                if fmt != 'General':
                    if fmt not in nfi:
                        nfi[fmt] = len(NFS)
                        NFS.append(fmt)
                    NF[k] = nfi[fmt]
                if isinstance(v, str) and v.startswith('='):
                    try:
                        ast = Parser(tokenize(v[1:]), ctx).parse()
                        F[k] = emit(ast)
                        FX[k] = v
                    except Exception as ex:
                        errors.append((k, v, str(ex)))
                else:
                    if isinstance(v, (datetime.datetime, datetime.date)):
                        v = serial(v)
                    V[k] = v
                    if is_input(c):
                        INP[k] = {'list': dv.get(c.coordinate)} if c.coordinate in dv else {}
    if errors:
        for e in errors[:30]:
            print('PARSE ERROR', e)
        raise SystemExit(1)
    names = {n: (key(*e[1:]) if e[0] == 'cell' else
                 "%d!%s%d:%s%d" % (e[1], get_column_letter(e[3]), e[2], get_column_letter(e[5]), e[4]))
             for n, e in ctx.names.items()}
    with open(dst, 'w', encoding='utf-8') as f:
        f.write('/* مولَّد آلياً بواسطة tools/xlsx2js.py من: %s – لا تعدّله يدوياً */\n' % src.split('/')[-1])
        f.write('window.GL_MODEL={\n')
        f.write('source:%s,\n' % json.dumps(src.split('/')[-1], ensure_ascii=False))
        f.write('built:%s,\n' % json.dumps(datetime.date.today().isoformat()))
        f.write('sheets:%s,\n' % json.dumps(ctx.sheets, ensure_ascii=False))
        f.write('names:%s,\n' % json.dumps(names, ensure_ascii=False))
        f.write('dims:%s,\n' % json.dumps(DIMS, ensure_ascii=False))
        f.write('nfs:%s,\n' % json.dumps(NFS, ensure_ascii=False))
        f.write('nf:%s,\n' % json.dumps(NF, ensure_ascii=False))
        f.write('inp:%s,\n' % json.dumps(INP, ensure_ascii=False))
        f.write('v:%s,\n' % json.dumps(V, ensure_ascii=False))
        f.write('fx:%s,\n' % json.dumps(FX, ensure_ascii=False))
        f.write('f:{\n')
        for k, code in F.items():
            f.write('%s:($,g,R)=>%s,\n' % (json.dumps(k), code))
        f.write('}};\n')
    print('sheets=%d constants=%d formulas=%d inputs=%d names=%d -> %s' %
          (len(ctx.sheets), len(V), len(F), len(INP), len(names), dst))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
