#!/usr/bin/env python3
"""
build.py: stamp the version everywhere, then inline the design system.

The version lives in one place, the VERSION file. This script writes it into:

  index.html        <meta name="version">, every <span class="tc-version">,
                    and the ?v= on every script it loads
  canvsyllabus.js   const VERSION, and the ?v= on pdf.js and its worker

The ?v= query strings are the cache-buster. When the version changes, the
browser sees new URLs for the scripts and fetches them fresh, so a hard
refresh of the page is enough to get a whole new build, not a new page
running last week's JavaScript.

Then it runs design/tc_build.py, which inlines design/tc.css.

Run from the repository root:

    python3 build.py            stamp and rebuild
    python3 build.py --check    exit 1 if anything is out of date

Run --check before you push.
"""

from pathlib import Path
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
VERSION = (ROOT / "VERSION").read_text(encoding="utf-8").strip()

if not re.fullmatch(r"\d+\.\d+\.\d+", VERSION):
    raise SystemExit(f"VERSION should look like 1.2.3, found {VERSION!r}")

RULES = {
    "index.html": [
        (r'(<meta name="version" content=")[^"]*(")', rf"\g<1>{VERSION}\g<2>"),
        (r'(<span class="tc-version">)[^<]*(</span>)', rf"\g<1>{VERSION}\g<2>"),
        (r'(\.(?:js|mjs)\?v=)[0-9.]+', rf"\g<1>{VERSION}"),
    ],
    "canvsyllabus.js": [
        (r'(const VERSION = ")[^"]*(")', rf"\g<1>{VERSION}\g<2>"),
        (r'(\.(?:js|mjs)\?v=)[0-9.]+', rf"\g<1>{VERSION}"),
    ],
}


def main(argv: list[str]) -> int:
    check = "--check" in argv
    stale = []
    for rel, rules in RULES.items():
        path = ROOT / rel
        before = path.read_text(encoding="utf-8")
        after = before
        for pattern, repl in rules:
            after, n = re.subn(pattern, repl, after)
            if n == 0:
                raise SystemExit(f"{rel}: nothing matched {pattern!r}. Did a marker get edited away?")
        if before == after:
            print(f"ok       {rel} ({VERSION})")
        elif check:
            stale.append(rel)
            print(f"STALE    {rel}")
        else:
            path.write_text(after, encoding="utf-8")
            print(f"stamped  {rel} ({VERSION})")
    if stale:
        print("\nRun python3 build.py and commit the result.", file=sys.stderr)
        return 1
    tc = [sys.executable, str(ROOT / "design" / "tc_build.py")] + (["--check"] if check else [])
    return subprocess.call(tc, cwd=ROOT)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
