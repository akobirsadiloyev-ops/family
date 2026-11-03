import sys, openpyxl
src, dst = sys.argv[1], sys.argv[2]
DROP = {'F.I.Sh.','JSHSHIR','Telefon raqami','Manzil','Diplom raqami','Mahalla','Tashkilot STIR','ID','T/R'}
wb = openpyxl.load_workbook(src, read_only=True)
out = openpyxl.Workbook(); out.remove(out.active)
for ws in wb:
    rows = list(ws.iter_rows(values_only=True))
    h = rows[0]
    keep = [i for i,x in enumerate(h) if x is not None and x not in DROP]
    o = out.create_sheet(ws.title)
    for r in rows:
        o.append([r[i] if i < len(r) else None for i in keep])
out.save(dst)
