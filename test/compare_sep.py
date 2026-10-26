import sys, openpyxl, collections
out, fin = sys.argv[1], sys.argv[2]
ws = openpyxl.load_workbook(out, data_only=True)['Summary']
hdr = [c.value for c in ws[4]]
print('title:', ws['A2'].value); print('header:', hdr)
tool = {}
for r in ws.iter_rows(min_row=5, values_only=True):
    if r[1] is None: continue
    tool[r[1]] = [float(x or 0) for x in r[3:11]]
fw = openpyxl.load_workbook(fin, data_only=True)
fs = fw['Summary']
raj = collections.defaultdict(float)
for r in fw['Provision_Rajeev'].iter_rows(min_row=4, values_only=True):
    if isinstance(r[2], (int, float)) and isinstance(r[8], (int, float)): raj[int(r[2])] += r[8]
finance = {}
for r in fs.iter_rows(min_row=5, values_only=True):
    if not isinstance(r[1], (int, float)): continue
    c = int(r[1]); v = [float(x) if isinstance(x, (int, float)) else 0.0 for x in r[3:9]]
    v += [raj.get(c, 0.0), v[4] + raj.get(c, 0.0)]
    finance[c] = v
cols = ['FY27 Pro','FY26 Pro','FY27 Bk','FY26 Bk','FY27 CF','FY26 CF','Rajeev','YTD']
print('centres tool', len(tool), 'finance', len(finance), 'only tool', set(tool)-set(finance), 'only finance', set(finance)-set(tool))
diffs = 0
for c in finance:
    d = [round(a-b, 2) for a, b in zip(tool[c], finance[c])]
    if any(abs(x) >= 0.01 for x in d):
        diffs += 1; print('DIFF', c, {k: x for k, x in zip(cols, d) if abs(x) >= 0.01})
print('centres with any difference >= 1 paisa:', diffs)
