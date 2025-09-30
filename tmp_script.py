from pathlib import Path
lines = Path("index.js").read_text(encoding="utf-8").splitlines()
print(lines[15])
