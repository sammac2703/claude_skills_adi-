"""Find Slide Library layouts by need (one line each). python scripts/find_layout.py "four numbers" [-v]"""
import json
import sys
from pathlib import Path

F = json.loads((Path(__file__).resolve().parent.parent / "assets" / "layout_facts.json").read_text(encoding="utf-8"))


def main():
    q = [w.lower() for w in sys.argv[1:] if not w.startswith("-")]
    v = "-v" in sys.argv
    scored = []
    for name, f in F.items():
        hay = f"{name} {f['use']}".lower()
        s = sum(w in hay for w in q)
        if s or not q:
            scored.append((s, name, f))
    for s, name, f in sorted(scored, key=lambda t: -t[0])[:8]:
        print(f"{name} | {'dark' if f['dark'] else 'light'} (logo {f['logo']}, icons {f['icons']}) | {f['use']}")
        if v:
            print("   idx:", ", ".join(f"{p['idx']}@{p['x']},{p['y']} {p['w']}x{p['h']}" for p in f["placeholders"]))


main()
