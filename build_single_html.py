#!/usr/bin/env python3
"""Bundle index copy.html + drama-quotes-data.js + logo into one shareable HTML file."""

import base64
import re
from pathlib import Path

BASE = Path(__file__).resolve().parent
HTML_IN = BASE / "index copy.html"
JS_IN = BASE / "drama-quotes-data.js"
LOGO_IN = BASE / "cover_layout" / "夜市人生大標去背.png"
HTML_OUT = BASE / "夜市人生台語金句學習器.html"


def main():
    if not HTML_IN.exists():
        raise SystemExit(f"Missing {HTML_IN}")
    if not JS_IN.exists():
        raise SystemExit(f"Missing {JS_IN}")
    if not LOGO_IN.exists():
        raise SystemExit(f"Missing {LOGO_IN}")

    html = HTML_IN.read_text(encoding="utf-8")
    js = JS_IN.read_text(encoding="utf-8").strip()
    if js.endswith(";"):
        js = js[:-1].rstrip()

    logo_b64 = base64.b64encode(LOGO_IN.read_bytes()).decode("ascii")
    logo_data_uri = f"data:image/png;base64,{logo_b64}"

    html = html.replace(
        'src="cover_layout/夜市人生大標去背.png"',
        f'src="{logo_data_uri}"',
    )

    html = re.sub(
        r'<script src="drama-quotes-data\.js"></script>\s*',
        f"<script>\n{js}\n</script>\n",
        html,
        count=1,
    )

    html = html.replace(
        "<title>夜市人生 · 台語金句長輩圖學習站</title>",
        "<title>夜市人生 · 台語金句線上學習器</title>",
    )

    HTML_OUT.write_text(html, encoding="utf-8")
    size_mb = HTML_OUT.stat().st_size / (1024 * 1024)
    print(f"Wrote {HTML_OUT}")
    print(f"  Size: {size_mb:.2f} MB")
    print("  Share this single file — open via http://localhost or any web server.")
    print("  YouTube embed still needs internet; double-click (file://) shows preview link only.")


if __name__ == "__main__":
    main()
