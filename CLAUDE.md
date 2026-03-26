# CLAUDE.md -- Tri-Tech TIA Presentation Tools

## Project Purpose

HTML-to-PPTX converter for the Tri-Tech IA (TIA) program.
Core tool: `html_to_pptx.py` -- parses HTML slides via Playwright, extracts DOM positions/styles, creates native PPTX elements via python-pptx.

## Key Files

- `html_to_pptx.py` -- Main converter (Playwright + python-pptx). THE active development focus.
- `presentazione_html/*.html` -- Source HTML slides (1280x720, Tailwind CSS + Google Fonts), 11 slides (1.html - 11.html)
- `run.bat` -- Launcher: runs html_to_pptx.py and opens output in PowerPoint

## Tech Stack

- Python 3.13, Playwright (Chromium), python-pptx
- HTML slides use: Tailwind CSS (v2 CDN, patched to Play CDN at runtime), Google Fonts (Poppins, Inter, Roboto Mono), Font Awesome icons
- No test framework yet

## Critical Domain Knowledge

### CSS px to PPTX pt conversion
CSS pixels != typographic points. At 96 DPI: `1px = 0.75pt`. Always convert: `Pt(css_px * 0.75)`.
Slide dimensions: 13.333" x 7.5" maps to 1280x720 viewport. Scale: `Inches(13.333) / 1280` EMU per pixel.

### Font metric compensation
Web fonts (Poppins, Inter) have different widths than Windows fallbacks (Segoe UI, Consolas).
Ratios hardcoded in `html_to_pptx.py` (originally computed via fonttools from actual TTF metrics):
- Poppins/Segoe UI: 1.137 (Poppins is 13.7% WIDER)
- Inter/Segoe UI: 1.08
- Roboto Mono/Consolas: 1.092

Use these ratios to adjust textbox widths: `fallback_width = browser_width / ratio * safety_margin`.

### Tailwind CSS patching
The converter patches all HTML slides at runtime via regex: replaces any Tailwind v2 `<link>` with Play CDN `<script src="https://cdn.tailwindcss.com">`. This is needed because v2 pre-built CSS does not support arbitrary values (`bg-[#060606]`, `text-[#FF6200]`).

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

Note: `html_to_pptx.py` defaults to `Slides1/` as input dir, override with first argument. Output defaults to `Slides1.pptx`, override with second argument.

## Known Limitations

- Font Awesome icons and SVGs are screenshotted (not native PPTX) -- positioning may be slightly off
- Complex CSS layouts (grid, flex) are approximated with absolute positioning
- PowerPoint font substitution is unpredictable when web fonts are not installed
- No automated testing -- verification is visual
