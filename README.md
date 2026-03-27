<div align="center">

<h1>html2pptx</h1>

**Convert HTML slides to native PowerPoint — not screenshots, actual editable shapes and text.**

[![Python 3.13+](https://img.shields.io/badge/Python-3.13%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue?style=flat-square)](LICENSE)
[![Playwright](https://img.shields.io/badge/Playwright-Chromium-2EAD33?style=flat-square&logo=playwright&logoColor=white)](https://playwright.dev/python/)
[![python-pptx](https://img.shields.io/badge/python--pptx-native%20PPTX-B7472A?style=flat-square)](https://python-pptx.readthedocs.io/)

</div>

---

html2pptx renders your HTML slides in headless Chromium, extracts every DOM element's position and computed style, then rebuilds the slide as native PowerPoint shapes, text boxes, and images. The result is a `.pptx` you can edit in PowerPoint — not a stack of rasterized screenshots.

## Why html2pptx?

- **Native elements** — Text stays editable, shapes stay resizable. No giant images per slide.
- **Pixel-perfect positioning** — DOM coordinates are converted to PPTX EMUs with exact scale factors.
- **Font-aware** — Web fonts (Poppins, Inter, Roboto Mono) are mapped to Windows fallbacks with width-compensation ratios computed via fonttools.
- **Works with any HTML** — Smart container detection finds the slide root automatically. Tailwind, vanilla CSS, whatever.
- **Configurable aspect ratio** — 16:9, 4:3, or any custom viewport. Slide dimensions adapt automatically.

## Quick Start

```bash
# Install dependencies
pip install playwright python-pptx
playwright install chromium

# Clone and run
git clone https://github.com/acaprino/html2pptx.git
cd html2pptx
python html_to_pptx.py -i presentazione_html -o slides.pptx
```

Your `.pptx` is ready. Open it in PowerPoint.

## Usage

```
python html_to_pptx.py [-i INPUT] [-o OUTPUT] [--width W] [--height H]
```

| Flag | Description | Default |
|------|-------------|---------|
| `-i`, `--input` | Directory containing HTML files | `presentazione_html` |
| `-o`, `--output` | Output PPTX path | `Slides1.pptx` |
| `--width` | Viewport width in pixels | `1280` |
| `--height` | Viewport height in pixels | `720` |

```bash
# 4:3 presentation
python html_to_pptx.py -i my_slides --width 1024 --height 768

# Custom output path
python html_to_pptx.py -o presentation_final.pptx

# Defaults (16:9, presentazione_html → Slides1.pptx)
python html_to_pptx.py
```

On Windows, `run.bat` runs the converter and opens the result in PowerPoint automatically.

<details>
<summary><b>How it works</b></summary>

1. **Render** — Each `.html` file is loaded in headless Chromium via Playwright
2. **Extract** — JavaScript traverses the DOM and collects positions, sizes, colors, fonts, borders, and z-order for every visible element
3. **Map fonts** — Web fonts are mapped to Windows system fonts with width-compensation ratios (e.g., Poppins → Segoe UI at 1.137x)
4. **Build PPTX** — python-pptx creates native shapes, text boxes, and images at the exact pixel positions, scaled to slide EMUs
5. **Handle edge cases** — SVGs and icon fonts are screenshotted and embedded as images. Tailwind v2 CSS is auto-patched to the Play CDN for arbitrary value support. Overflow is clipped to slide boundaries.

### Container Detection

The converter finds the slide root via cascade:
1. `.w-[1280px]` (Tailwind arbitrary width)
2. `[class*="1280"]` (any class containing the viewport width)
3. Largest visible direct child of `<body>` by area
4. `document.body` as fallback

This means it works with any HTML structure, not just Tailwind-based slides.

### CSS px → PPTX pt

CSS pixels ≠ typographic points. At 96 DPI: `1px = 0.75pt`. The scale factor `SLIDE_W / VP_W` converts pixel coordinates to EMUs, recalculated when `--width`/`--height` change.

</details>

## HTML Slide Format

Each slide is a standalone `.html` file. Name them `1.html`, `2.html`, etc. for ordered output. The converter sorts numerically, then alphabetically for non-numeric names.

Your HTML can use any CSS framework or none at all. The only requirement: the slide content should be inside a fixed-size container matching the viewport dimensions.

## Known Limitations

- Font Awesome icons and SVGs are embedded as screenshots, not vector shapes
- Complex CSS layouts (grid, flex) are approximated with absolute positioning
- Font rendering depends on which fonts are installed on the system
- No automated tests — verification is visual

## Contributing

Contributions are welcome. The project is a single Python file (`html_to_pptx.py`), so diving in is straightforward.

1. Fork the repo
2. Make your changes
3. Test by running the converter and inspecting the output PPTX
4. Open a PR

## License

[MIT](LICENSE) — Alfio Caprino

