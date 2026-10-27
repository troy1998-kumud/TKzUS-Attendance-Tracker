"""Build preview/preview.html: the real page wired to an in-browser fake of Google Sheets.
Run:  python3 preview/build_preview.py   then open preview/preview.html in a browser.
Data lives only in that browser tab and resets on reload. PINs are printed in the browser console."""
import pathlib
root = pathlib.Path(__file__).resolve().parent.parent
page = (root / "src" / "Index.html").read_text(encoding="utf-8")
mock = (root / "preview" / "gasmock.js").read_text(encoding="utf-8")
code = (root / "src" / "Code.gs").read_text(encoding="utf-8")
boot = "setup(); console.log('Staff PINs', __sheets.Employees._d.slice(1).map(r => r[0] + ' ' + r[6]).join(', '), '| HR PIN', __sheets.Settings._d[2][1]);"
inject = "<script>" + mock + "</script><script>" + code + "\n" + boot + "</script>"
out = page.replace('<base target="_top">', '<base target="_top">' + inject, 1)
(root / "preview" / "preview.html").write_text(out, encoding="utf-8")
print("Wrote", root / "preview" / "preview.html")
