"""Find Brand Icon Library icons by keyword. python scripts/find_icon.py team meeting   (run extract_icons.py once first)"""
import json
import sys
from pathlib import Path

idx = Path(__file__).resolve().parent.parent / "assets" / "icons" / "index.json"
cat = json.loads(idx.read_text(encoding="utf-8"))
q = [w.lower() for w in sys.argv[1:]]
hits = sorted(((sum(w in v.lower() for w in q), v) for v in cat.values() if any(w in v.lower() for w in q)), reverse=True)
print(", ".join(v for _, v in hits[:15]) or f"no match in {len(cat)} icons")
