# CLAUDE.md -- html2pptx

## Project Purpose

Convert HTML slides to native PPTX (editable shapes / text, not screenshots).
Single-module Python tool: `html_to_pptx.py`. Renders each HTML in a headless
Chromium via Playwright, extracts DOM positions and computed styles via
injected JavaScript, then emits python-pptx native shapes, textboxes, and
embedded images.

Also importable as a library: `from html_to_pptx import convert`.

## Key Files

- `html_to_pptx.py` -- main module. CLI + library entry point `convert(...)`.
- `presentazione_html/*.html` -- sample slides (1280x720, Tailwind CSS, Google Fonts).
- `pyproject.toml` -- pinned dependencies and console-script entry point.
- `run.bat` / `run.sh` -- platform launchers (forward to the Python CLI).
- `.gitignore` -- excludes generated PPTX/PDFs, virtualenvs, IDE folders, review-pipeline workspaces.

## Tech Stack

- Python 3.10+ (3.13 recommended)
- Playwright (Chromium) for rendering
- python-pptx for native shape generation
- lxml (transitively, via python-pptx)

Install: `pip install -e .` then `playwright install chromium`.

## CLI Usage

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
| `--width`, `--height` | Viewport in CSS px | 1280x720 |
| `-s`, `--simplify` | CSS simplification 0-10 | 0 |
| `--strict` | Exit code 2 on any warning or partial save | off |
| `--no-javascript` | Disable JS in headless Chromium (safest for untrusted decks) | off |
| `--allow-network HOST` | Allowlist network host (repeatable) | tailwind/Google Fonts |
| `--block-network` | Block ALL network (overrides allowlist) | off |
| `--tailwind-wait-ms` | Cap on Tailwind JIT readiness wait | 1500 |
| `--max-slides` | Hard cap on slide count | 5000 |
| `-v`, `-vv` | Increase log verbosity | warning |

Library entry point: `convert(input_dir=..., output=..., ...) -> ConversionReport`.

## Architecture (single-file)

- `SlideContext` (dataclass) -- holds viewport / slide / scale; passed to every
  geometry helper. Replaces the legacy mutable module globals.
- `ConversionReport` -- aggregated per-run statistics (skipped slides, failed
  screenshots, timeouts, partial saves).
- `WalkerOutput` / `ShapeEntry` / `TextEntry` / `SvgPrimitive` (TypedDicts) --
  document the JS<->Python contract for the DOM walker output.
- `parse_rgba` / `to_rgb` -- CSS color parsing supporting rgb/rgba (incl. CSS L4
  space-separated), hsl/hsla, hex (3/4/6/8 digits), named colors. CSS L4
  oklch/oklab/color/color-mix log a debug warning and return None.
- `convert(...)` -- per-slide pipeline: HTML patch -> Chromium render -> walk DOM
  -> extract -> screenshot fallback elements -> python-pptx assembly. Each slide
  runs in its own `BrowserContext` so cookies / localStorage / service workers
  cannot leak between attacker-controlled inputs.

## Domain Knowledge

### CSS px to EMU and pt

- `EMU = ctx.px(v) = round(v * SCALE)`. `SCALE = SLIDE_W_EMU / VP_W` and lands
  exactly on PowerPoint's canonical 12,192,000 EMU at 16:9.
- `Pt = v * (72 / 96)` (96 DPI).
- Coordinates use top-left origin in both DOM and PPTX EMU.

### Container detection

- Find the largest `<body>` direct child by viewport-intersected bounding box.
- Filters: `display:none`, `visibility:hidden`, `opacity:0`, `inert`,
  `aria-hidden`, off-screen positioning, sub-pixel size.
- Threshold: at least 100 px^2 (10x10).

### Walker schema

- `out.shapes`: rect/oval entries. Per-axis border-radius `brX`/`brY` plus a
  legacy single `br` for backward compat. CSS `z-index` captured as `z`.
- `out.texts`: text runs with precise tx/tw from Range API + `multiline` flag
  derived from `getClientRects().length > 1`.
- `out.svgs` / `out.icons`: screenshot-fallback entries (with container width
  cw/ch so Python applies the same skip threshold as JS).
- `out.nativeSvgs[N].elements`: per-primitive entries flattened in `build_slide`
  so they can interleave with surrounding HTML shapes in z-order.
- `out.fontRatios`: web font -> Windows fallback width compensation. Cached on
  `window.__h2pFontRatios` so subsequent slides skip the DOM-thrash measurement.

### Sort and clipping

- Top-level sort key: `(z, dp, seq)` ascending. Stable.
- Borders use `dp + 0.5` so they paint between parent (dp) and child (dp+1).
- Clipping happens against the viewport before sorting; text boxes drop only if
  more than 50% overflows by height.

### Paint paths

- Circular elements (per-axis `>= 45%` AND aspect `< 1.15` AND minDim > 20) ->
  `MSO_SHAPE.OVAL` with single oval outline.
- Other rounded elements with single-axis radius > 4 -> `ROUNDED_RECTANGLE`.
- Rectangular borders rendered as 4 thin shapes at `dp + 0.5`.
- Native SVG primitives: circle/ellipse -> OVAL, rect -> rect/rounded, line/path/
  polygon/polyline -> freeform via `build_freeform`. Multi-subpath paths fall
  back to screenshot.
- Hostile-HTML hardening: per-slide `BrowserContext`, network egress allowlist,
  Symbol-keyed hide/restore state (instead of `window.__ssHidden`), strip
  pre-existing `data-si`/`data-fi` from input HTML.

### Output integrity

- Atomic write via `os.replace(tmp, output)`.
- On failure, fall back to `<output>.partial.pptx` and (with `--strict`) exit 2.

## Workflow

1. `pip install -e .` and `playwright install chromium`.
2. `python html_to_pptx.py -i presentazione_html` (or `run.bat` / `run.sh`).
3. `start Slides1.pptx` (Windows) / `open Slides1.pptx` (macOS) / `xdg-open Slides1.pptx` (Linux).
4. For untrusted decks, add `--no-javascript --block-network` for maximum isolation.

## Known Limitations

- CSS `oklch()`, `oklab()`, `color()`, `color-mix()` colors: not converted; logged at debug level.
- Font Awesome glyphs and complex SVGs are screenshotted (rasterized).
- Complex CSS layouts (grid, flex) are approximated with absolute positioning.
- Pipeline is per-slide sequential. Worker-pool concurrency is a planned follow-up.
