#!/usr/bin/env python3
"""
html_to_pptx.py - Parse HTML slides into native PPTX elements.

Renders each HTML in Playwright, extracts DOM element positions and
computed styles via JavaScript, creates python-pptx shapes and textboxes.
SVGs and Font Awesome icons are screenshotted and embedded as images.
"""
import argparse, os, re, sys, tempfile
from pathlib import Path
from playwright.sync_api import sync_playwright
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE

BASE = Path(__file__).resolve().parent
HTML_DIR = BASE / "presentazione_html"
OUTPUT = BASE / "Slides1.pptx"
VP_W, VP_H = 1280, 720
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
SCALE = SLIDE_W / VP_W  # EMU per CSS pixel

# Font scale ratios: web_font_width / fallback_width (computed via fonttools)
# Ratio > 1 means web font is WIDER, so Segoe UI text will be narrower → more room
# Ratio < 1 means web font is NARROWER, so Segoe UI text will be wider → risk of overflow
FONT_RATIOS = {
    'Poppins': 1.137,       # Poppins is 13.7% wider than Segoe UI
    'Inter': 1.08,          # Inter is ~8% wider than Segoe UI (estimated)
    'Roboto Mono': 1.092,   # Roboto Mono is 9.2% wider than Consolas
}
FONT_MAP = {
    'Poppins': 'Segoe UI',
    'Inter': 'Segoe UI',
    'Roboto Mono': 'Consolas',
}

TAILWIND_CDN = '<script src="https://cdn.tailwindcss.com?plugins=forms,typography"></script>'

# Fix flex overflow: flex-1 items have min-height:auto by default, which prevents
# shrinking below content size. This causes content to overflow fixed-height containers.
# Setting min-height:0 allows flex items to actually shrink to fit their parent.
FLEX_FIX_CSS = '<style>.flex-1{min-height:0!important;min-width:0!important;}</style>'


def px(v):
    """CSS pixels to EMU."""
    return int(v * SCALE)


def parse_rgba(s):
    """Parse CSS color: 'rgb(r,g,b)', 'rgba(r,g,b,a)', or CSS L4 'rgb(r g b / a)'."""
    if not s:
        return None
    m = re.match(
        r'rgba?\(\s*([\d.]+)\s*[,\s]\s*([\d.]+)\s*[,\s]\s*([\d.]+)'
        r'(?:\s*[,/]\s*([\d.]+%?))?\s*\)', s
    )
    if not m:
        return None
    r, g, b = int(float(m[1])), int(float(m[2])), int(float(m[3]))
    if m[4]:
        a = float(m[4].rstrip('%')) / 100 if m[4].endswith('%') else float(m[4])
    else:
        a = 1.0
    return (r, g, b, a)


def to_rgb(c):
    return RGBColor(min(255, c[0]), min(255, c[1]), min(255, c[2])) if c else None


ALIGN = {
    'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT,
    'left': PP_ALIGN.LEFT, 'start': PP_ALIGN.LEFT,
    'end': PP_ALIGN.RIGHT, 'justify': PP_ALIGN.JUSTIFY,
}


# ── JavaScript DOM extraction (executed inside Playwright browser) ───

JS = r"""() => {
    // Container detection: find the largest visible top-level element
    const c = (function() {
        var best = null, bestArea = 0;
        var ch = document.body.children;
        for (var i = 0; i < ch.length; i++) {
            var el = ch[i], tag = el.tagName;
            if (!tag || tag === 'SCRIPT' || tag === 'STYLE' || tag === 'LINK' || tag === 'META') continue;
            var s = getComputedStyle(el);
            if (s.display === 'none' || s.visibility === 'hidden') continue;
            var r = el.getBoundingClientRect();
            if (r.width * r.height > bestArea) { bestArea = r.width * r.height; best = el; }
        }
        return bestArea >= 100 ? best : null;
    })();
    if (!c) return null;
    const cr = c.getBoundingClientRect();
    const ox = cr.left, oy = cr.top;
    const out = {
        bg: getComputedStyle(c).backgroundColor,
        bodyBg: getComputedStyle(document.body).backgroundColor,
        shapes: [], texts: [], svgs: [], icons: []
    };

    // Tag SVGs and FA icons with data attributes for later screenshotting
    c.querySelectorAll('svg').forEach((el, i) => el.setAttribute('data-si', String(i)));
    c.querySelectorAll('.fas,.far,.fab,.fa-solid,.fa').forEach((el, i) => el.setAttribute('data-fi', String(i)));

    function gr(el) {
        const b = el.getBoundingClientRect();
        return { x: b.left - ox, y: b.top - oy, w: b.width, h: b.height };
    }
    function vis(el) {
        const b = el.getBoundingClientRect();
        if (b.width < 0.5 || b.height < 0.5) return false;
        const s = getComputedStyle(el);
        return s.display !== 'none' && s.visibility !== 'hidden' && parseFloat(s.opacity) > 0.01;
    }
    function inl(el) { return getComputedStyle(el).display.startsWith('inline'); }
    function isFA(el) {
        if (!el.classList) return false;
        return ['fas','far','fab','fa-solid','fa'].some(function(cn) { return el.classList.contains(cn); });
    }
    function sty(el) {
        const s = getComputedStyle(el);
        return {
            ff: s.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
            fs: parseFloat(s.fontSize),
            fw: parseInt(s.fontWeight) || 400,
            fi: s.fontStyle === 'italic',
            co: s.color,
            tt: s.textTransform
        };
    }

    // seq: monotonic insertion counter for stable sort across element types
    var seq = 0;
    var dp = 0;
    function walk(el) {
        var tag = el.tagName;
        if (!tag || ['SCRIPT','STYLE','LINK','META','HEAD','BR','HR'].indexOf(tag) >= 0) return;
        if (!vis(el)) return;
        var rect = gr(el);
        var cs = getComputedStyle(el);

        // SVG -> record for screenshot
        if (tag === 'svg' || tag === 'SVG') {
            var si = el.getAttribute('data-si');
            if (si !== null) out.svgs.push({ x: rect.x, y: rect.y, w: rect.w, h: rect.h, dp: dp, seq: seq++, i: parseInt(si) });
            return;
        }
        // Font Awesome icon -> record for screenshot
        if (isFA(el)) {
            var fi = el.getAttribute('data-fi');
            if (fi !== null) out.icons.push({ x: rect.x, y: rect.y, w: rect.w, h: rect.h, dp: dp, seq: seq++, i: parseInt(fi) });
            return;
        }

        // Background shape (non-transparent bg color)
        var bg = cs.backgroundColor;
        if (bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent') {
            out.shapes.push({
                x: rect.x, y: rect.y, w: rect.w, h: rect.h,
                dp: dp, seq: seq++, bg: bg, op: parseFloat(cs.opacity),
                br: parseInt(cs.borderRadius) || 0
            });
        }

        // Colored borders >= 1px (table separators, accent bars, card borders)
        // dp + 0.5 places borders between parent (dp) and child (dp+1) depth
        var borderDefs = [
            ['Top',    function(r, bw) { return {x:r.x, y:r.y, w:r.w, h:bw}; }],
            ['Bottom', function(r, bw) { return {x:r.x, y:r.y+r.h-bw, w:r.w, h:bw}; }],
            ['Left',   function(r, bw) { return {x:r.x, y:r.y, w:bw, h:r.h}; }],
            ['Right',  function(r, bw) { return {x:r.x+r.w-bw, y:r.y, w:bw, h:r.h}; }]
        ];
        for (var bi = 0; bi < borderDefs.length; bi++) {
            var prop = borderDefs[bi][0];
            var mkR = borderDefs[bi][1];
            var bw = parseFloat(cs['border' + prop + 'Width']);
            var bs = cs['border' + prop + 'Style'];
            var bc = cs['border' + prop + 'Color'];
            if (bw >= 1 && bs !== 'none' && bc && bc !== 'rgba(0, 0, 0, 0)') {
                var sr = mkR(rect, bw);
                out.shapes.push({
                    x: sr.x, y: sr.y, w: sr.w, h: sr.h,
                    dp: dp + 0.5, seq: seq++, bg: bc, op: parseFloat(cs.opacity), br: 0
                });
            }
        }

        // Collect text runs with PRECISE bounding rects via Range API
        var runs = [];
        var tMinX = 9999, tMinY = 9999, tMaxX = 0, tMaxY = 0;
        var hasTB = false;
        var childNodes = el.childNodes;
        for (var ci = 0; ci < childNodes.length; ci++) {
            var ch = childNodes[ci];
            if (ch.nodeType === 3) {
                var t = ch.textContent.replace(/\s+/g, ' ').trim();
                if (t) {
                    runs.push(Object.assign({ t: t }, sty(el)));
                    var rng = document.createRange();
                    rng.selectNode(ch);
                    var rr = rng.getBoundingClientRect();
                    if (rr.width > 0 && rr.height > 0) {
                        tMinX = Math.min(tMinX, rr.left);
                        tMinY = Math.min(tMinY, rr.top);
                        tMaxX = Math.max(tMaxX, rr.right);
                        tMaxY = Math.max(tMaxY, rr.bottom);
                        hasTB = true;
                    }
                }
            } else if (ch.nodeType === 1 && inl(ch) && vis(ch) && !isFA(ch) &&
                       ch.tagName !== 'svg' && ch.tagName !== 'SVG') {
                // Capture inline element background color on the run data
                var inlBg = getComputedStyle(ch).backgroundColor;
                var t2 = ch.textContent.trim();
                if (t2) {
                    var rd = Object.assign({ t: t2 }, sty(ch));
                    if (inlBg && inlBg !== 'rgba(0, 0, 0, 0)' && inlBg !== 'transparent') {
                        var inlRect = ch.getBoundingClientRect();
                        rd.hlBg = inlBg;
                        rd.hlX = inlRect.left - ox;
                        rd.hlY = inlRect.top - oy;
                        rd.hlW = inlRect.width;
                        rd.hlH = inlRect.height;
                        rd.hlBr = parseInt(getComputedStyle(ch).borderRadius) || 0;
                    }
                    runs.push(rd);
                    var rr2 = ch.getBoundingClientRect();
                    if (rr2.width > 0 && rr2.height > 0) {
                        tMinX = Math.min(tMinX, rr2.left);
                        tMinY = Math.min(tMinY, rr2.top);
                        tMaxX = Math.max(tMaxX, rr2.right);
                        tMaxY = Math.max(tMaxY, rr2.bottom);
                        hasTB = true;
                    }
                }
            }
        }
        if (runs.length > 0) {
            // Precise text Y/H from Range API, container X/W for alignment
            var ty = hasTB ? (tMinY - oy) : rect.y;
            var th = hasTB ? (tMaxY - tMinY) : rect.h;
            var tx = hasTB ? (tMinX - ox) : rect.x;
            var tw = hasTB ? (tMaxX - tMinX) : rect.w;
            out.texts.push({
                x: rect.x, y: ty, w: rect.w, h: th,
                tx: tx, tw: tw,
                dp: dp, seq: seq++, runs: runs, ta: cs.textAlign
            });
        }

        // Recurse into block-level children
        dp++;
        var children = el.children;
        for (var ki = 0; ki < children.length; ki++) {
            var kid = children[ki];
            if (kid.nodeType !== 1) continue;
            if (isFA(kid) || kid.tagName === 'svg' || kid.tagName === 'SVG') {
                walk(kid);
                continue;
            }
            if (!inl(kid)) {
                walk(kid);
            } else {
                // Inline element with block descendants? Recurse.
                var hasBlock = false;
                for (var gi = 0; gi < kid.children.length; gi++) {
                    if (!inl(kid.children[gi])) { hasBlock = true; break; }
                }
                if (hasBlock) walk(kid);
            }
        }
        dp--;
    }

    walk(c);

    // Measure font width ratios: web font vs Windows fallback
    var _fonts = {};
    for (var _ti = 0; _ti < out.texts.length; _ti++) {
        for (var _ri = 0; _ri < out.texts[_ti].runs.length; _ri++) {
            var _ff = out.texts[_ti].runs[_ri].ff;
            if (_ff) _fonts[_ff] = 1;
        }
    }
    var _SYS = {'Segoe UI':1,'Arial':1,'Calibri':1,'Times New Roman':1,
                'Consolas':1,'Courier New':1,'Verdana':1,'Tahoma':1,'Georgia':1,'Trebuchet MS':1};
    var _REF = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    out.fontRatios = {};
    var _fnames = Object.keys(_fonts);
    for (var _fi = 0; _fi < _fnames.length; _fi++) {
        var _fn = _fnames[_fi];
        if (_SYS[_fn]) continue;
        // Measure web font
        var _s1 = document.createElement('span');
        _s1.style.cssText = 'position:absolute;visibility:hidden;white-space:nowrap;font-size:100px;font-family:"' + _fn + '",sans-serif';
        _s1.textContent = _REF;
        document.body.appendChild(_s1);
        // Detect monospace: compare narrow vs wide chars
        var _mi = document.createElement('span');
        _mi.style.cssText = 'position:absolute;visibility:hidden;white-space:nowrap;font-size:100px;font-family:"' + _fn + '"';
        _mi.textContent = 'iiiiii';
        var _mw = document.createElement('span');
        _mw.style.cssText = 'position:absolute;visibility:hidden;white-space:nowrap;font-size:100px;font-family:"' + _fn + '"';
        _mw.textContent = 'MMMMMM';
        document.body.appendChild(_mi);
        document.body.appendChild(_mw);
        var _isMono = Math.abs(_mi.getBoundingClientRect().width - _mw.getBoundingClientRect().width) < 2;
        document.body.removeChild(_mi);
        document.body.removeChild(_mw);
        var _fb = _isMono ? 'Consolas' : 'Segoe UI';
        // Measure fallback font
        var _s2 = document.createElement('span');
        _s2.style.cssText = 'position:absolute;visibility:hidden;white-space:nowrap;font-size:100px;font-family:"' + _fb + '",sans-serif';
        _s2.textContent = _REF;
        document.body.appendChild(_s2);
        var _w1 = _s1.getBoundingClientRect().width;
        var _w2 = _s2.getBoundingClientRect().width;
        document.body.removeChild(_s1);
        document.body.removeChild(_s2);
        if (_w1 > 0 && _w2 > 0) {
            out.fontRatios[_fn] = {r: _w1 / _w2, fb: _fb};
        }
    }

    return out;
}"""


# ── Screenshot SVGs and Font Awesome icons ───────────────────────────

def screenshot_elements(page, data, tmpdir):
    imgs = []
    for svg in data.get('svgs', []):
        # Skip full-slide SVGs (decorative backgrounds like grid/dot patterns).
        # Their screenshots capture text rendered on top, causing double text.
        if svg['w'] > VP_W * 0.8 and svg['h'] > VP_H * 0.8:
            continue
        si = int(svg['i'])
        p = os.path.join(tmpdir, f"svg_{si}.png")
        loc = page.locator(f'[data-si="{si}"]')
        if loc.count() > 0:
            try:
                loc.first.screenshot(path=p)
                imgs.append({**svg, 'path': p})
            except Exception as e:
                print(f"  WARN: screenshot failed for svg {si}: {e}", file=sys.stderr)
    for icon in data.get('icons', []):
        fi = int(icon['i'])
        p = os.path.join(tmpdir, f"fa_{fi}.png")
        loc = page.locator(f'[data-fi="{fi}"]')
        if loc.count() > 0:
            try:
                loc.first.screenshot(path=p)
                imgs.append({**icon, 'path': p})
            except Exception as e:
                print(f"  WARN: screenshot failed for icon {fi}: {e}", file=sys.stderr)
    return imgs


# ── Build PPTX slide from extracted data ─────────────────────────────

def build_slide(prs, data, images):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # Slide background (fallback to body bg if container is transparent)
    bg = parse_rgba(data.get('bg'))
    if not bg or bg[3] < 0.1:
        bg = parse_rgba(data.get('bodyBg'))
    if bg:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = to_rgb(bg)

    # Merge all visual elements; use (dp, seq) for stable depth+insertion-order sort
    elems = []
    for s in data.get('shapes', []):
        # Skip container-sized shapes that match slide background
        if s['w'] > VP_W - 20 and s['h'] > VP_H - 20:
            sbg = parse_rgba(s['bg'])
            if sbg and bg and sbg[:3] == bg[:3]:
                continue
        elems.append(('shape', (s.get('dp', 0), s.get('seq', 0)), s))
    for t in data.get('texts', []):
        elems.append(('text', (t.get('dp', 0), t.get('seq', 0)), t))
    for i in images:
        elems.append(('image', (i.get('dp', 0), i.get('seq', 0)), i))

    # Clip elements to slide boundaries (HTML uses overflow:hidden)
    clipped = []
    for etype, sort_key, ed in elems:
        ey = ed.get('y', 0)
        eh = ed.get('h', 0)
        ex = ed.get('x', 0)
        ew = ed.get('w', 0)
        # Skip elements fully outside the slide
        if ey >= VP_H or ey + eh <= 0 or ex >= VP_W or ex + ew <= 0:
            continue
        # Clip negative coordinates (left/top overflow)
        if ex < 0:
            ed = {**ed, 'w': ew + ex, 'x': 0}
            ew = ed['w']
            ex = 0
        if ey < 0:
            ed = {**ed, 'h': eh + ey, 'y': 0}
            eh = ed['h']
            ey = 0
        # Clip at slide edges: shapes/images fully, text height only (preserve width for wrapping)
        if ey + eh > VP_H:
            ed = {**ed, 'h': VP_H - ey}
        if etype in ('shape', 'image'):
            if ex + ew > VP_W:
                ed = {**ed, 'w': VP_W - ex}
        elif etype == 'text':
            # Skip text boxes mostly outside (>50% overflow by height)
            visible_h = min(VP_H - ey, eh)
            if visible_h < eh * 0.5:
                continue
        clipped.append((etype, sort_key, ed))
    elems = clipped

    elems.sort(key=lambda e: e[1])

    font_ratios = data.get('fontRatios', {})

    for etype, _, ed in elems:
        try:
            if etype == 'shape':
                _add_shape(slide, ed)
            elif etype == 'text':
                _add_text(slide, ed, font_ratios)
            elif etype == 'image':
                _add_image(slide, ed)
        except Exception as e:
            print(f"  WARN: skipped {etype}: {e}", file=sys.stderr)


def _add_shape(slide, s):
    w, h = max(s['w'], 1), max(s['h'], 1)
    # Filter out tiny decorative shapes (borders, dots) that create noise
    if w < 5 and h < 5:
        return
    c = parse_rgba(s['bg'])
    if not c:
        return
    effective_opacity = s.get('op', 1.0) * c[3]
    if effective_opacity < 0.15:
        return
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if s.get('br', 0) > 4 else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, px(s['x']), px(s['y']), px(w), px(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = to_rgb(c)
    shp.line.fill.background()  # no outline


def _add_text(slide, t, font_ratios=None):
    runs = t.get('runs', [])
    if not runs or t['w'] < 2:
        return
    if font_ratios is None:
        font_ratios = {}

    x, y, w, h = t['x'], t['y'], t['w'], t['h']
    ta = t.get('ta', 'left')
    max_fs = max(rd.get('fs', 16) for rd in runs)

    # Use precise text bounds (tx/tw from Range API) for positioning
    tx = t.get('tx', x)  # precise text X from Range API
    tw = t.get('tw', w)  # precise text width from Range API
    # Detect multi-line vs single-line text from browser rendering
    # If text height > 1.8x font size, it's multi-line (paragraph that should wrap)
    is_multiline = h > max_fs * 1.8

    # Font-aware width adjustment (measured ratio > hardcoded > 1.0)
    primary_font = runs[0].get('ff', 'Segoe UI') if runs else 'Segoe UI'
    measured = font_ratios.get(primary_font, {})
    ratio = measured.get('r', FONT_RATIOS.get(primary_font, 1.0))
    fallback_factor = (1.0 / ratio) * 1.05

    if is_multiline:
        # Multi-line text: use CONTAINER width (card boundary), keep wrapping
        # Don't use text width — the text should wrap at the container edge
        w_use = w * fallback_factor
        if ta == 'center':
            extra = w_use - w
            x = max(0, x - extra / 2)
            w = min(VP_W, w_use)
        else:
            x = t['x']  # container X, not text X
            w = min(VP_W - x, w_use)
    else:
        # Single-line text: use precise text width, no wrapping
        if ta == 'center':
            needed_w = w * fallback_factor
            extra = needed_w - w
            x = max(0, x - extra / 2)
            w = min(VP_W, needed_w)
        else:
            x = tx
            w = tw * fallback_factor
            w = min(VP_W - x, w)

    h = max(h, max_fs * 1.2)
    # Cap height at slide boundary (clipping may have reduced h earlier)
    if y + h > VP_H:
        h = max(VP_H - y, 1)

    txBox = slide.shapes.add_textbox(px(x), px(y), px(w), px(h))
    tf = txBox.text_frame
    tf.auto_size = MSO_AUTO_SIZE.NONE  # prevent PowerPoint from resizing

    if is_multiline:
        tf.word_wrap = True  # paragraphs wrap at container boundary
    else:
        tf.word_wrap = False  # labels/headings never wrap
    p = tf.paragraphs[0]
    p.alignment = ALIGN.get(ta, PP_ALIGN.LEFT)

    # First pass: create highlighted text boxes for runs with inline backgrounds
    for rd in runs:
        if rd.get('hlBg'):
            hl_bg = parse_rgba(rd['hlBg'])
            if not hl_bg:
                continue
            hl_x, hl_y = rd['hlX'], rd['hlY']
            hl_w, hl_h = rd['hlW'], rd['hlH']
            hl_text = rd['t']
            tt = rd.get('tt', 'none')
            if tt == 'uppercase': hl_text = hl_text.upper()
            elif tt == 'lowercase': hl_text = hl_text.lower()
            elif tt == 'capitalize': hl_text = hl_text.title()
            # Create a text box with solid fill for the inline highlight
            hlBox = slide.shapes.add_textbox(px(hl_x), px(hl_y), px(hl_w), px(hl_h))
            hlTf = hlBox.text_frame
            hlTf.auto_size = MSO_AUTO_SIZE.NONE
            hlTf.word_wrap = False
            hlTf.margin_left = hlTf.margin_right = hlTf.margin_top = hlTf.margin_bottom = 0
            # White fill background
            hlBox.fill.solid()
            hlBox.fill.fore_color.rgb = to_rgb(hl_bg)
            hlP = hlTf.paragraphs[0]
            hlP.alignment = PP_ALIGN.CENTER
            hlR = hlP.add_run()
            hlR.text = hl_text
            web_font = rd.get('ff', 'Segoe UI')
            m = font_ratios.get(web_font, {})
            hlR.font.name = m.get('fb', FONT_MAP.get(web_font, web_font))
            hlR.font.size = Pt(rd.get('fs', 16) * 0.75)
            hlR.font.bold = rd.get('fw', 400) >= 600
            co = parse_rgba(rd.get('co', ''))
            if co:
                hlR.font.color.rgb = to_rgb(co)

    for i, rd in enumerate(runs):
        text = rd['t']
        # Apply CSS text-transform
        tt = rd.get('tt', 'none')
        if tt == 'uppercase':
            text = text.upper()
        elif tt == 'lowercase':
            text = text.lower()
        elif tt == 'capitalize':
            text = text.title()
        # Space between inline runs (HTML whitespace collapsing)
        if i > 0:
            text = ' ' + text
        r = p.add_run()
        r.text = text
        web_font = rd.get('ff', 'Segoe UI')
        m = font_ratios.get(web_font, {})
        r.font.name = m.get('fb', FONT_MAP.get(web_font, web_font))
        # Convert CSS pixels to typographic points: 96px = 72pt
        r.font.size = Pt(rd.get('fs', 16) * 0.75)
        r.font.bold = rd.get('fw', 400) >= 600
        r.font.italic = rd.get('fi', False)
        co = parse_rgba(rd.get('co', ''))
        if co:
            # If this run has an inline bg, make it invisible in the main textbox
            # (the separate highlighted textbox handles rendering)
            if rd.get('hlBg'):
                hl_bg = parse_rgba(rd['hlBg'])
                if hl_bg:
                    r.font.color.rgb = to_rgb(hl_bg)  # text color = bg color → invisible
                else:
                    r.font.color.rgb = to_rgb(co)
            else:
                r.font.color.rgb = to_rgb(co)


def _add_image(slide, img):
    path = img.get('path', '')
    if not path or not os.path.exists(path):
        return
    if img['w'] < 2 or img['h'] < 2:
        return
    slide.shapes.add_picture(path, px(img['x']), px(img['y']), px(img['w']), px(img['h']))


# ── Main ──────────────────────────────────────────────────────────────

def main():
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')

    parser = argparse.ArgumentParser(description='Convert HTML slides to native PPTX elements')
    parser.add_argument('html_dir', nargs='?', default=str(HTML_DIR),
                        help='Directory with HTML files (default: presentazione_html)')
    parser.add_argument('output', nargs='?', default=str(OUTPUT),
                        help='Output PPTX path (default: Slides1.pptx)')
    parser.add_argument('--width', type=int, default=VP_W,
                        help='Viewport width in pixels (default: %(default)s)')
    parser.add_argument('--height', type=int, default=VP_H,
                        help='Viewport height in pixels (default: %(default)s)')
    args = parser.parse_args()

    html_dir = Path(args.html_dir)
    output = Path(args.output)

    # Update globals so px(), build_slide(), etc. use the right dimensions
    global VP_W, VP_H, SLIDE_W, SLIDE_H, SCALE
    VP_W, VP_H = args.width, args.height
    if VP_W <= 0 or VP_H <= 0:
        parser.error("--width and --height must be positive integers")
    SLIDE_H = Inches(7.5)
    SLIDE_W = Inches(7.5 * VP_W / VP_H)  # preserve aspect ratio
    SCALE = SLIDE_W / VP_W

    files = sorted(html_dir.glob("*.html"),
                   key=lambda f: (0, int(f.stem)) if f.stem.isdigit() else (1, f.stem))
    if not files:
        print(f"No HTML files found in {html_dir}")
        return

    print(f"Converting {len(files)} HTML slides -> PPTX (native elements)\n")
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    with tempfile.TemporaryDirectory() as tmpdir:
        with sync_playwright() as pw:
            browser = pw.chromium.launch()
            try:
                page = browser.new_page(viewport={"width": VP_W, "height": VP_H})

                for idx, hf in enumerate(files):
                    print(f"  [{idx+1}/{len(files)}] {hf.name} ... ", end="", flush=True)

                    # Patch HTML: replace Tailwind v2 CDN with Play CDN
                    # (v2 pre-built CSS doesn't support arbitrary values like bg-[#060606])
                    html_content = hf.read_text(encoding='utf-8')
                    patched = re.sub(
                        r'<link[^>]+tailwindcss@[^>]+/dist/tailwind[^>]*/?>',
                        TAILWIND_CDN,
                        html_content
                    )
                    # Inject flex overflow fix (min-height:0 on flex-1 items)
                    patched = patched.replace('</head>', FLEX_FIX_CSS + '</head>')
                    temp_html = Path(tmpdir) / f"slide_{idx}.html"
                    temp_html.write_text(patched, encoding='utf-8')

                    page.goto(f"file:///{temp_html.resolve().as_posix()}")
                    page.wait_for_load_state("networkidle")
                    page.wait_for_timeout(500)  # Tailwind JIT compile time

                    data = page.evaluate(JS)
                    if not data:
                        print("SKIP (no container)")
                        continue

                    sd = os.path.join(tmpdir, f"s{idx}")
                    os.makedirs(sd, exist_ok=True)
                    images = screenshot_elements(page, data, sd)
                    build_slide(prs, data, images)

                    ns = len(data.get('shapes', []))
                    nt = len(data.get('texts', []))
                    ni = len(images)
                    print(f"{ns} shapes, {nt} texts, {ni} imgs")
            finally:
                browser.close()

    try:
        prs.save(str(output))
    except Exception as e:
        fallback = output.with_suffix('.partial.pptx')
        try:
            prs.save(str(fallback))
            print(f"\nERROR saving {output.name}: {e}", file=sys.stderr)
            print(f"Partial output saved: {fallback.name}")
        except Exception as e2:
            print(f"\nERROR: unable to save any output: {e}; fallback also failed: {e2}", file=sys.stderr)
            sys.exit(1)
        return
    kb = output.stat().st_size // 1024
    print(f"\nSaved: {output} ({len(files)} slides, {kb:,} KB)")


if __name__ == "__main__":
    main()
