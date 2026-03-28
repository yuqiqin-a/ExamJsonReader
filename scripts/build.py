from pathlib import Path
import shutil

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / "dist"
ASSETS = DIST / "assets"

if DIST.exists():
    shutil.rmtree(DIST)

ASSETS.mkdir(parents=True, exist_ok=True)

index_src = (ROOT / "index.html").read_text(encoding="utf-8")
index_out = index_src.replace("./src/styles.css", "./assets/styles.css").replace("./src/main.js", "./assets/main.js")

(DIST / "index.html").write_text(index_out, encoding="utf-8")
shutil.copy2(ROOT / "src" / "styles.css", ASSETS / "styles.css")
shutil.copy2(ROOT / "src" / "main.js", ASSETS / "main.js")

print(f"Build complete: {DIST}")
