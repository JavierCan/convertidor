from pathlib import Path
import re

ROOT = Path(__file__).parent
SRC = ROOT / "src"
DIST = ROOT / "dist"

html = (SRC / "index.html").read_text(encoding="utf-8")
css = (SRC / "styles.css").read_text(encoding="utf-8")
js = (SRC / "app.js").read_text(encoding="utf-8")

css_compact = re.sub(r"/\*.*?\*/", "", css, flags=re.S)
css_compact = re.sub(r"\s+", " ", css_compact)
css_compact = re.sub(r"\s*([{}:;,>])\s*", r"\1", css_compact).strip()

# El JavaScript permanece legible dentro del único HTML para evitar
# transformaciones inseguras de expresiones regulares y plantillas.
html = html.replace("/*__CSS__*/", css_compact)
html = html.replace("/*__JS__*/", js.strip())

DIST.mkdir(parents=True, exist_ok=True)
(DIST / "app.min.html").write_text(html, encoding="utf-8")

print("Generado:", DIST / "app.min.html")
