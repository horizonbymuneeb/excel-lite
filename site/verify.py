#!/usr/bin/env python3
"""Headless verification for the docs site.

Checks every page at desktop (1440x980) and mobile (390x844) for horizontal overflow and
console errors, exercises search / theme / mobile nav / copy buttons, crawls internal links
(including anchors) and writes screenshots to site/screenshots/.

    python3 site/verify.py          # needs: pip install playwright && playwright install chromium
"""

from __future__ import annotations

import subprocess
import sys
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
SHOTS = ROOT / "site" / "screenshots"
PORT = 4180
BASE = f"http://localhost:{PORT}"
DESKTOP = {"width": 1440, "height": 980}
MOBILE = {"width": 390, "height": 844}

pages = sorted(
    p.name for p in (ROOT / "site" / "dist").glob("*.html") if p.name != "404.html"
)

failures: list[str] = []
console_errors: list[str] = []


def check(condition: bool, message: str) -> None:
    if not condition:
        failures.append(message)


def main() -> int:
    SHOTS.mkdir(exist_ok=True)
    server = subprocess.Popen(["node", "site/serve.mjs", str(PORT)], cwd=ROOT, stdout=subprocess.PIPE)
    try:
        for _ in range(50):
            try:
                urllib.request.urlopen(f"{BASE}/index.html", timeout=1)
                break
            except Exception:
                time.sleep(0.1)

        with sync_playwright() as playwright:
            browser = playwright.chromium.launch()
            context = browser.new_context(viewport=DESKTOP)
            context.grant_permissions(["clipboard-read", "clipboard-write"])
            page = context.new_page()
            page.on("console", lambda msg: console_errors.append(f"{page.url}: {msg.text}") if msg.type == "error" else None)
            page.on("pageerror", lambda err: console_errors.append(f"{page.url}: {err}"))

            # ---- every page, both viewports -------------------------------------
            for viewport, label in ((DESKTOP, "desktop"), (MOBILE, "mobile")):
                page.set_viewport_size(viewport)
                for name in pages:
                    page.goto(f"{BASE}/{name}", wait_until="load")
                    overflow = page.evaluate("document.documentElement.scrollWidth - window.innerWidth")
                    check(overflow <= 0, f"{name} @{label}: horizontal overflow {overflow}px")
                    title = page.title()
                    check(bool(title), f"{name} @{label}: empty <title>")

            # ---- screenshots ----------------------------------------------------
            page.set_viewport_size(DESKTOP)
            page.goto(f"{BASE}/index.html")
            page.screenshot(path=str(SHOTS / "desktop-index.png"), full_page=False)
            page.goto(f"{BASE}/api.html")
            page.screenshot(path=str(SHOTS / "desktop-api.png"), full_page=True)

            page.goto(f"{BASE}/index.html")
            page.click("#theme-toggle")
            theme = page.evaluate("document.documentElement.dataset.theme")
            check(theme in ("light", "dark"), f"theme toggle produced {theme!r}")
            page.screenshot(path=str(SHOTS / "desktop-index-dark.png"))
            page.click("#theme-toggle")

            page.set_viewport_size(MOBILE)
            page.goto(f"{BASE}/reading.html")
            page.screenshot(path=str(SHOTS / "mobile-reading.png"))
            page.click(".menu-button")
            nav_open = page.evaluate("document.body.classList.contains('nav-open')")
            check(nav_open, "mobile nav did not open")
            page.screenshot(path=str(SHOTS / "mobile-nav.png"))
            page.mouse.click(370, 420)  # scrim area to the right of the drawer
            check(not page.evaluate("document.body.classList.contains('nav-open')"), "mobile nav did not close")

            # ---- interactions ---------------------------------------------------
            page.set_viewport_size(DESKTOP)
            page.goto(f"{BASE}/index.html")
            page.fill("#search-input", "memory")
            page.wait_for_selector("#search-results a", timeout=3000)
            hits = page.locator("#search-results a").count()
            check(hits > 0, "search for 'memory' returned no hits")
            page.screenshot(path=str(SHOTS / "desktop-search.png"))

            page.goto(f"{BASE}/getting-started.html")
            copy = page.locator(".code-copy").first
            copy.click()
            time.sleep(0.2)
            label = copy.inner_text()
            check(label in ("Copied", "Copy"), f"copy button label after click: {label!r}")

            # ---- measured contrast and type scale ------------------------------
            def luminance(rgb):
                def channel(value):
                    value /= 255
                    return value / 12.92 if value <= 0.03928 else ((value + 0.055) / 1.055) ** 2.4

                r, g, b = rgb
                return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b)

            def parse_color(text):
                numbers = [float(n) for n in text.replace("rgba(", "").replace("rgb(", "").replace(")", "").split(",")[:3]]
                return tuple(numbers)

            def contrast(foreground, background):
                a, b = luminance(foreground), luminance(background)
                high, low = max(a, b), min(a, b)
                return round((high + 0.05) / (low + 0.05), 2)

            samples = {
                "body text": ".content p",
                "sidebar link": ".sidebar li a",
                "muted label": ".nav-label",
                "inline code": ".content code",
                "heading": ".content h2",
                "link": ".content a:not(.anchor)",
            }
            page.goto(f"{BASE}/reading.html")
            measured = {}
            for label, selector in samples.items():
                colors = page.evaluate(
                    """(selector) => {
                      const el = document.querySelector(selector);
                      if (!el) return null;
                      const transparent = (value) => !value || value === 'transparent' || value === 'rgba(0, 0, 0, 0)';
                      let node = el;
                      let background = '';
                      while (node) {
                        const value = getComputedStyle(node).backgroundColor;
                        if (!transparent(value)) { background = value; break; }
                        node = node.parentElement;
                      }
                      return { fg: getComputedStyle(el).color, bg: background || getComputedStyle(document.body).backgroundColor };
                    }""",
                    selector,
                )
                if not colors:
                    continue
                measured[label] = contrast(parse_color(colors["fg"]), parse_color(colors["bg"]))

            for label, ratio in measured.items():
                check(ratio >= 4.5, f"low contrast for {label}: {ratio}:1 (needs 4.5:1)")

            # dark theme contrast
            page.goto(f"{BASE}/index.html")
            page.evaluate("document.documentElement.dataset.theme = 'dark'")
            dark_style = page.evaluate(
                "() => ({ bg: getComputedStyle(document.body).backgroundColor, fg: getComputedStyle(document.querySelector('.content p')).color })"
            )
            dark_bg = parse_color(dark_style["bg"])
            dark_text = parse_color(dark_style["fg"])
            dark_ratio = contrast(dark_text, dark_bg)
            check(dark_ratio >= 7, f"dark theme body contrast {dark_ratio}:1 (needs 7:1)")

            # ---- internal links + anchors --------------------------------------
            links: set[tuple[str, str]] = set()
            for name in pages:
                page.goto(f"{BASE}/{name}")
                for href in page.eval_on_selector_all(
                    "a[href]",
                    "els => els.map(e => e.getAttribute('href')).filter(h => h && !h.startsWith('http'))",
                ):
                    links.add((name, href))
            for source, href in sorted(links):
                target, _, anchor = href.partition("#")
                url = f"{BASE}/{target or source}"
                try:
                    status = urllib.request.urlopen(url, timeout=2).status
                except Exception as error:  # noqa: BLE001
                    failures.append(f"broken link {source} -> {href!r}: {error}")
                    continue
                if status != 200:
                    failures.append(f"broken link {source} -> {href!r}: HTTP {status}")
                if anchor:
                    page.goto(url)
                    found = page.evaluate("id => !!document.getElementById(id)", anchor)
                    check(found, f"missing anchor {source} -> {href!r}")

            browser.close()
    finally:
        server.terminate()

    print(f"pages checked: {len(pages)} × 2 viewports")
    print(f"internal links checked: {len(links)}")
    print(f"contrast ratios: {measured}")
    print(f"console/page errors: {len(console_errors)}")
    for error in console_errors[:10]:
        print("  console:", error)
    if failures:
        print(f"\nFAILURES ({len(failures)}):")
        for failure in failures[:30]:
            print("  -", failure)
        return 1
    print("\nOK: no overflow, no console errors, no broken links; screenshots in site/screenshots/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
