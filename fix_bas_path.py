import io
p = r"C:\Users\dai86\.zcode\workspace\default\capula_rv\DataUpdater.bas"
with open(p, "rb") as f:
    b = f.read()
old = br'"C:\Users\dai86\.zcode\workspace\default\capula_rv\vba_progress.log"'
new = br'ThisWorkbook.Path & "\vba_progress.log"'
n = b.count(old)
b2 = b.replace(old, new)
with open(p, "wb") as f:
    f.write(b2)
print("replaced", n, "occurrences")
# verify no hardcoded dev path remains
with open(p, "rb") as f:
    c = f.read()
print("dai86 left:", c.count(b"dai86"))
print("ThisWorkbook.Path count:", c.count(b"ThisWorkbook.Path"))
