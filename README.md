<div align="center">

<h1>html2pptx</h1>

**Convert HTML slides to native PowerPoint -- not screenshots, actual editable shapes and text.**

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue?style=flat-square)](LICENSE)
[![Playwright](https://img.shields.io/badge/Playwright-Chromium-2EAD33?style=flat-square&logo=playwright&logoColor=white)](https://playwright.dev/python/)
[![python-pptx](https://img.shields.io/badge/python--pptx-native%20PPTX-B7472A?style=flat-square)](https://python-pptx.readthedocs.io/)

</div>

---

html2pptx renders your HTML slides in headless Chromium, extracts every DOM
element's position and computed style, then rebuilds the slide as native
PowerPoint shapes, text boxes, and images. The result is a `.pptx` you can
edit in PowerPoint -- not a stack of rasterized screenshots.

## Why html2pptx?

- **Native elements** -- text stays editable, shapes stay resizable. No giant images per slide.
- **Pixel-perfect positioning** -- DOM coordinates convert to PPTX EMUs at canonical PowerPoint scale (12,192,000 EMU at 16:9).
- **Font-aware** -- web fonts are measured against system fallbacks in the browser; cached across slides.
- **CSS L3 + L4 colors** -- `rgb()` / `rgba()` / `hsl()` / `hsla()` / hex; CSS L4 (`oklch`, `oklab`, `color()`) detected and logged.
- **Z-index aware** -- CSS `z-index` participates in paint order.
- **Per-axis border-radius** -- pill buttons render correctly regardless of aspect ratio.
- **Smart container detection** -- finds the slide root by viewport-intersected bounding box; rejects off-screen overlays.
- **Hardened by default** -- per-slide BrowserContext isolation, network-egress allowlist, optional `--no-javascript`, atomic save.

## Quick start

```bash
pip install -e .
playwright install chromium

python html_to_pptx.py -i presentazione_html -o slides.pptx
```

## Usage

```
python html_to_pptx.py [-i INPUT] [-o OUTPUT] [--width W] [--height H] [-s LEVEL]
                       [--strict] [--no-javascript] [--allow-network HOST]
                       [--block-network] [--tailwind-wait-ms MS]
                       [--max-slides N] [-v|-vv]
```

| Flag | Description | Default |
|------|-------------|---------|
| `-i`, `--input` | Directory with HTML files | `presentazione_html` |
| `-o`, `--output` | Output PPTX path | `Slides1.pptx` |
| `--width` / `--height` | Viewport in CSS px | 1280 / 720 |
| `-s`, `--simplify 0-10` | CSS simplification level | 0 |
| `--strict` | Exit code 2 on any warning or partial save | off |
| `--no-javascript` | Disable JS in headless Chromium (safest for untrusted decks) | off |
| `--allow-network HOST` | Hostname substring permitted during render (repeatable) | tailwind / Google Fonts |
| `--block-network` | Block ALL network (overrides `--allow-network`) | off |
| `--tailwind-wait-ms MS` | Cap on Tailwind JIT readiness wait | 1500 |
| `--max-slides N` | Hard cap on slides processed | 5000 |
| `-v` / `-vv` | Increase log verbosity | warning |

```bash
# 4:3 deck
python html_to_pptx.py -i my_slides --width 1024 --height 768

# Untrusted deck: fully isolated, no JS, no network
python html_to_pptx.py -i untrusted/ --no-javascript --block-network --strict

# Defaults (16:9, presentazione_html -> Slides1.pptx)
python html_to_pptx.py
```

On Windows, `run.bat` runs the converter and opens the result. On macOS / Linux, `run.sh`.

## Library use

```python
from pathlib import Path
from html_to_pptx import convert

report = convert(
    input_dir=Path("decks"),
    output=Path("out.pptx"),
    width=1280, height=720,
    strict=True,
)
print(report.summary_line())
```

`convert(...)` returns a `ConversionReport` with per-run statistics and accepts the same
options as the CLI. Configuration flows through a `SlideContext` -- no module globals.

<details>
<summary><b>How it works</b></summary>

1. **Render** -- each `.html` file is loaded in headless Chromium via Playwright, in its own isolated `BrowserContext` (cookies / localStorage / service workers do not leak between slides).
2. **Sanitize** -- pre-existing `data-si` / `data-fi` attributes are stripped from input HTML so the walker's tagging cannot be hijacked.
3. **Preprocess** -- DOM is simplified in-place (e.g., `<br>` flattened into block-level wrappers).
4. **Extract** -- JavaScript walks the DOM and collects positions, sizes, colors, fonts, borders, per-axis border-radius, CSS z-index, and multi-line text flag for every visible element.
5. **Map fonts** -- web fonts are measured against system fallbacks in the browser; results cached across slides.
6. **Build PPTX** -- python-pptx creates native shapes (rectangles, rounded rectangles, ovals), text boxes, and images at exact pixel positions, scaled to canonical PowerPoint EMU.
7. **Handle edge cases** -- complex SVGs and Font Awesome icons are screenshotted into in-memory PNGs and embedded; Tailwind v2 `<link>` tags are auto-patched to the v3 Play CDN; multi-subpath SVG paths fall back to screenshot.
8. **Save** -- atomic temp-file write then rename. On failure, partial output goes to `<output>.partial.pptx`; `--strict` returns exit 2.

</details>

## HTML slide format

Each slide is a standalone `.html` file. Name them `1.html`, `2.html`, etc. for ordered output. Files are sorted by natural order: `1, 2, 10` rather than `1, 10, 2`.

Your HTML can use any CSS framework or none at all. The only requirement: the slide content should be inside a fixed-size container matching the viewport dimensions.

## Security

- Untrusted decks should be converted with `--no-javascript` (disables `<script>` execution) and `--block-network` (no outbound HTTP). Default conversion permits Tailwind / Google Fonts only.
- Output paths are validated: no UNC, no Windows ADS, must end in `.pptx`.
- Input directory is required to be a real (non-symlink) directory.
- Per-file size cap of 5 MB.
- Bidi-override and zero-width Unicode is stripped from text runs before embedding (defense-in-depth against homograph URLs in the deliverable).
- Tailwind v2 detection regex is bounded and ReDoS-safe.

## Known limitations

- CSS L4 colors (`oklch()`, `oklab()`, `color()`, `color-mix()`) log at debug level and skip the affected element. Tailwind v3.3+ uses some `oklch` palette colors -- if you see "missing colors", run with `-vv` to confirm.
- Font Awesome glyphs and complex SVGs are embedded as screenshots, not vectors.
- Complex CSS layouts (grid, flex) are approximated with absolute positioning.
- Font rendering depends on which fonts are installed on the viewing machine.
- Pipeline is per-slide sequential. Worker-pool concurrency is a planned follow-up.

## Contributing

```bash
pip install -e .
playwright install chromium
python html_to_pptx.py -i presentazione_html -vv  # debug logging
```

Open a PR with the regenerated `Slides1.pptx` for visual diff.

## License

[MIT](LICENSE) -- Alfio Caprino
