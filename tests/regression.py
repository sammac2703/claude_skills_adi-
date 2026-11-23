"""Regression: build the spec deck, run the QA gate, write previews. Run after changing any script.
    python tests/regression.py [outdir]     exit 0 = built, gate passed"""
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SC = HERE.parent / "scripts"
out = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "out"
out.mkdir(parents=True, exist_ok=True)
deck = out / "regression.pptx"
for cmd in ([sys.executable, str(SC / "build_from_spec.py"), str(HERE / "regression_spec.json"), str(deck)],
            [sys.executable, str(SC / "qa_layout.py"), str(deck), "-q"],
            [sys.executable, str(SC / "preview_png.py"), str(deck), str(out / "png")]):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    print((r.stdout or "").strip().splitlines()[-1] if r.stdout.strip() else "", r.stderr.strip()[-400:])
    if r.returncode:
        print("FAILED:", " ".join(Path(c).name for c in cmd[:2]))
        sys.exit(1)
print("regression OK ->", deck)
