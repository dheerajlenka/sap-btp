#!/usr/bin/env python3
"""Inject data.json into the dashboard template to produce a self-contained dist/index.html.

Usage:
    python build/build.py [--data data/data.json] [--template src/dashboard.template.html]
                          [--out dist/index.html]

The output is a single HTML file with the data baked in as window.__DATA__ and
no external/CDN references — safe to publish as a Claude Artifact or on GitHub Pages.
"""
import argparse
import json
import os
import re
import sys

DEFAULTS = dict(
    data="data/data.json",
    template="src/dashboard.template.html",
    out="dist/index.html",
)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--data", default=DEFAULTS["data"])
    ap.add_argument("--template", default=DEFAULTS["template"])
    ap.add_argument("--out", default=DEFAULTS["out"])
    args = ap.parse_args()

    with open(args.data, encoding="utf-8") as f:
        data = json.load(f)
    with open(args.template, encoding="utf-8") as f:
        template = f.read()

    if "{{DATA}}" not in template:
        sys.exit("Template is missing the {{DATA}} placeholder.")

    # Escape </script> so an inline JSON string can't terminate the script tag.
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    html = template.replace("{{DATA}}", payload)

    # Guard: the artifact must be self-contained (no external CDN references).
    external = re.findall(r"""(?:src|href)\s*=\s*['"]https?://[^'"]+""", html)
    if external:
        sys.exit("Refusing to build: external references found:\n  " + "\n  ".join(external))

    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)

    kb = len(html.encode("utf-8")) / 1024
    print(f"Wrote {args.out} ({kb:.0f} KB, self-contained, {len(data['agents'])} agents)")


if __name__ == "__main__":
    main()
