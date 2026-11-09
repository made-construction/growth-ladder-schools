# يصدّر قيم كل خلايا المعادلات من نسخة معاد حسابها بـ LibreOffice إلى JSON للمقارنة
import sys, json, datetime, openpyxl
EPOCH = datetime.datetime(1899, 12, 30)
src, rc, out = sys.argv[1:4]
wf = openpyxl.load_workbook(src); wv = openpyxl.load_workbook(rc, data_only=True)
ref = {}
for si, ws in enumerate(wf):
    vs = wv[ws.title]
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                v = vs[c.coordinate].value
                if isinstance(v, datetime.datetime): v = (v - EPOCH).total_seconds() / 86400
                ref["%d!%s" % (si, c.coordinate)] = v
json.dump(ref, open(out, 'w'), ensure_ascii=False)
print(len(ref))
