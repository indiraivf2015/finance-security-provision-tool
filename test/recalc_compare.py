import sys, openpyxl
a = openpyxl.load_workbook(sys.argv[1], data_only=True)   # tool's own cached values
f = openpyxl.load_workbook(sys.argv[1], data_only=False)  # formulas
b = openpyxl.load_workbook(sys.argv[2], data_only=True)   # LibreOffice recalculated
n = bad = 0
for ws in f.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('=') and 'XLOOKUP' not in c.value.upper():
                n += 1
                va = a[ws.title][c.coordinate].value; vb = b[ws.title][c.coordinate].value
                if isinstance(va, (int, float)) and isinstance(vb, (int, float)):
                    if abs(va - vb) > 0.005: bad += 1; print('MISMATCH', ws.title, c.coordinate, c.value[:60], va, vb)
                elif va != vb and not (va in (None, 0) and vb in (None, 0)):
                    bad += 1; print('MISMATCH', ws.title, c.coordinate, c.value[:60], repr(va), repr(vb))
print('formula cells checked', n, 'mismatches', bad)
