# CLAUDE.md -- Tri-Tech TIA Presentation Tools

## Project Purpose

HTML-to-PPTX converter -- originally built for Tri-Tech IA (TIA), now generic.
Core tool: `html_to_pptx.py` -- parses HTML slides via Playwright, extracts DOM positions/styles, creates native PPTX elements via python-pptx.

## Key Files

- `html_to_pptx.py` -- Main converter (Playwright + python-pptx). THE active development focus.
- `presentazione_html/*.html` -- TIA source slides (1280x720, Tailwind CSS + Google Fonts), 11 slides (1.html - 11.html)
- `run.bat` -- Launcher: runs html_to_pptx.py and opens output in PowerPoint

## Tech Stack

- Python 3.13, Playwright (Chromium), python-pptx
- TIA slides use: Tailwind CSS (v2 CDN, patched to Play CDN at runtime), Google Fonts (Poppins, Inter, Roboto Mono), Font Awesome icons
- No test framework yet

## CLI Usage

```
python html_to_pptx.py [html_dir] [output] [--width W] [--height H]
```

- `html_dir` -- Directory with HTML files (default: `presentazione_html`)
- `output` -- Output PPTX path (default: `Slides1.pptx`)
- `--width` -- Viewport width in pixels (default: 1280)
- `--height` -- Viewport height in pixels (default: 720)

Slide dimensions adapt to aspect ratio: height fixed at 7.5", width = `7.5 * width/height` inches (13.333" for 16:9, 10" for 4:3).

## Critical Domain Knowledge

### CSS px to PPTX pt conversion
CSS pixels != typographic points. At 96 DPI: `1px = 0.75pt`. Always convert: `Pt(css_px * 0.75)`.
Scale factor: `SLIDE_W / VP_W` EMU per pixel (recalculated when `--width`/`--height` change).

### Container detection (generic)
The JS extractor finds the slide container via cascade:
1. `.w-[1280px]` (TIA-specific Tailwind class)
2. `[class*="1280"]` (any class containing "1280")
3. Largest visible direct child of `<body>` by area
4. `document.body` as final fallback

This makes the converter work with any HTML, not just TIA slides.

### Font handling
Known web fonts are mapped to Windows fallbacks with width-compensation ratios (computed via fonttools):
- Poppins -> Segoe UI (ratio 1.137)
- Inter -> Segoe UI (ratio 1.08)
- Roboto Mono -> Consolas (ratio 1.092)

Unknown fonts pass through as-is (`FONT_MAP.get(web_font, web_font)`) with ratio 1.0 and 5% safety margin.

### Tailwind CSS patching
Conditional: only triggers when HTML contains a Tailwind v2 `<link>` tag (regex match). Replaces it with Play CDN `<script>` because v2 pre-built CSS doesn't support arbitrary values (`bg-[#060606]`). Non-Tailwind HTML is unaffected.

### Flex overflow fix
Injects `<style>.flex-1{min-height:0!important;min-width:0!important;}</style>` into all slides. Fixes flex-1 items overflowing fixed-height containers. Harmless on non-Tailwind HTML.

## Approach to Bug Fixing

- Solve problems with precise calculations, not hacks (no arbitrary % reductions)
- Every fix must be grounded in measurable data (fonttools metrics, Range API bounds, browser measurements)
- Test every change by regenerating the PPTX and visually inspecting output
- When font metrics differ: compute the exact ratio, don't guess

## Workflow

1. Edit `html_to_pptx.py`
2. Run: `python html_to_pptx.py presentazione_html` (or use `run.bat`)
3. Open output: `start Slides1.pptx`
4. Export PDF from PowerPoint for comparison

## Known Limitations

- Font Awesome icons and SVGs are screenshotted (not native PPTX) -- positioning may be slightly off
- Complex CSS layouts (grid, flex) are approximated with absolute positioning
- PowerPoint font substitution is unpredictable when web fonts are not installed
- If save fails, attempts fallback to `.partial.pptx`
- No automated testing -- verification is visual
