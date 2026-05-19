"""Shared geometry helpers for the captioner placement and audit pipelines.

Imported by apply_captions.py, audit_caption_placement.py, and verify.py so
all three agree on slide-band thresholds, footer clearance, and obstacle-rect
semantics. Keeping these constants/functions in one place prevents drift —
which previously caused audit-vs-apply geometric mismatches.

Vertical-only obstacle checking is the v0.2.1 contract. v0.3.0 may upgrade
to 2D box-intersection (placement-budget pre-pass architecture).
"""
import os
from pptx.enum.shapes import PP_PLACEHOLDER

# ---------------------------------------------------------------------------
# Geometry constants — single source of truth for all three pipelines.
# ---------------------------------------------------------------------------
MIN_CAPTION_HEIGHT  = 250_000   # ~0.27 in — absolute floor for a readable caption box.
FOOTER_CLEARANCE_EMU = 91_440   # ~0.10 in — visual gap between caption bottom and footer band.
MIN_CAPTION_WIDTH_EMU = 1_828_800  # ~2.00 in — minimum width so a 30-50 char caption fits in 1 line.
EMU_PER_CHAR_DEFAULT  = 114_000  # ~0.125 in/char — conservative estimate for 10pt italic Calibri.
                                 # See "Fix-D heuristic" in SKILL.md "Known limits."
LOW_BAND_FRACTION   = 0.86       # bottom 14% of slide = "footer band" search zone.

# Placeholder type sets — kept here so apply/audit/verify agree on what counts
# as a "title" obstacle vs a "body" obstacle vs a "footer" land-mine.
TITLE_TYPES  = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
BODY_TYPES   = {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.SUBTITLE,
                PP_PLACEHOLDER.OBJECT}  # CONTENT alias = OBJECT in python-pptx
FOOTER_TYPES = {PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE,
                PP_PLACEHOLDER.SLIDE_NUMBER}

CAPTION_NAME_PREFIXES = ('captioner_caption_', 'captioner_sa_icon_')


# ---------------------------------------------------------------------------
# Footer band: the y-coordinate below which captions must not extend.
# ---------------------------------------------------------------------------
def slide_footer_top(slide, slide_h, exclude_caption_shapes=True):
    """Return the topmost y (EMU) of the reserved bottom band. Considers
    FOOTER/DATE/SLIDE_NUMBER placeholders on slide/layout/master AND any
    low-band text shape (catches branded master footers like "Pearson").
    A footer-typed placeholder PARKED AT THE TOP of a master (top < band)
    is NOT a bottom footer — ignore it (else the band collapses to the
    slide top). When called on a captioned deck, caption shapes are
    excluded so they don't pollute their own footer line.
    """
    limit = slide_h - 274_320  # ~0.30 in universal bottom safety margin
    band  = int(slide_h * LOW_BAND_FRACTION)

    sources = [slide]
    try:
        sources.append(slide.slide_layout)
        sources.append(slide.slide_layout.slide_master)
    except Exception:
        pass

    for src in sources:
        try:
            shapes = list(src.shapes)
        except Exception:
            continue
        for sh in shapes:
            try:
                if exclude_caption_shapes:
                    nm = sh.name or ''
                    if any(nm.startswith(p) for p in CAPTION_NAME_PREFIXES):
                        continue
                top = sh.top
                if top is None or top < band:
                    continue
                is_foot = False
                try:
                    pf = sh.placeholder_format
                    if pf is not None and pf.type in FOOTER_TYPES:
                        is_foot = True
                except Exception:
                    pass
                has_txt = bool(getattr(sh, 'has_text_frame', False)
                               and sh.text_frame.text.strip())
                if is_foot or has_txt:
                    limit = min(limit, int(top))
            except Exception:
                continue
    return max(0, limit)


# ---------------------------------------------------------------------------
# Title obstacle: caption must not visually overlap a slide title.
# ---------------------------------------------------------------------------
def slide_title_rect(slide):
    """Bounding rect (L, T, R, B) EMU enclosing every TITLE/CENTER_TITLE
    placeholder with visible text, wherever positioned. None if no titled text."""
    L = T = R = B = None
    for sh in slide.shapes:
        if (sh.name or '').startswith(CAPTION_NAME_PREFIXES):
            continue
        try:
            pf = sh.placeholder_format
            if pf is None or pf.type not in TITLE_TYPES:
                continue
            if not (getattr(sh, 'has_text_frame', False)
                    and sh.text_frame.text.strip()):
                continue
            if None in (sh.left, sh.top, sh.width, sh.height):
                continue
            l, t = int(sh.left), int(sh.top)
            r, b = l + int(sh.width), t + int(sh.height)
            L = l if L is None else min(L, l)
            T = t if T is None else min(T, t)
            R = r if R is None else max(R, r)
            B = b if B is None else max(B, b)
        except Exception:
            continue
    return None if L is None else (L, T, R, B)


# ---------------------------------------------------------------------------
# Body obstacles: SUBTITLE/BODY/OBJECT placeholders with non-empty text.
# Returns vertical-band tuples (y_top, y_bottom). Vertical-only is the
# v0.2.1 contract — v0.3.0 may upgrade to 2D rects.
# ---------------------------------------------------------------------------
def slide_body_obstacle_bands(slide):
    """List of (y_top, y_bottom) for every body-type placeholder with visible
    text on this slide. Used by the placement obstacle filter. Vertical-only
    semantics (consistent with _voverlap)."""
    bands = []
    for sh in slide.shapes:
        if (sh.name or '').startswith(CAPTION_NAME_PREFIXES):
            continue
        try:
            pf = sh.placeholder_format
            if pf is None or pf.type not in BODY_TYPES:
                continue
            if not (getattr(sh, 'has_text_frame', False)
                    and sh.text_frame.text.strip()):
                continue
            if sh.top is None or sh.height is None:
                continue
            bands.append((int(sh.top), int(sh.top) + int(sh.height)))
        except Exception:
            continue
    return bands


# ---------------------------------------------------------------------------
# Vertical overlap predicates.
# ---------------------------------------------------------------------------
def _voverlap_band(c_top, c_h, y_top, y_bottom, frac=0.15):
    """True if caption interval [c_top, c_top+c_h] vertically overlaps the
    band [y_top, y_bottom] by more than `frac` of caption height."""
    if y_top is None or y_bottom is None:
        return False
    iy = max(0, min(c_top + c_h, y_bottom) - max(c_top, y_top))
    return iy > frac * c_h


def _voverlap(c_top, c_h, title_rect, frac=0.15):
    """Back-compat shim for the original signature. title_rect = (L, T, R, B)."""
    if title_rect is None:
        return False
    return _voverlap_band(c_top, c_h, title_rect[1], title_rect[3], frac)


def _clear_all_obstacles(c_top, c_h, obstacles, frac=0.15):
    """True if caption [c_top, c_top+c_h] clears EVERY obstacle. Obstacles
    accepted in either format:
      - 4-tuple (L, T, R, B): use T, B (vertical-only — v0.2.1 contract)
      - 2-tuple (y_top, y_bottom): use directly
    """
    for obs in obstacles:
        if obs is None:
            continue
        if len(obs) == 4:
            y_top, y_bottom = obs[1], obs[3]
        elif len(obs) == 2:
            y_top, y_bottom = obs
        else:
            continue
        if _voverlap_band(c_top, c_h, y_top, y_bottom, frac):
            return False
    return True


# ---------------------------------------------------------------------------
# 2D rect intersection — used by audit for caption-caption overlap detection.
# ---------------------------------------------------------------------------
def rect_intersect_area(a, b):
    """Area of intersection of two (l, t, w, h) rects. 0 if disjoint."""
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    ix = max(0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0, min(ay + ah, by + bh) - max(ay, by))
    return ix * iy


# ---------------------------------------------------------------------------
# Picture coverage (for informational metadata in overlay-fullbleed rows).
# ---------------------------------------------------------------------------
def visible_coverage(p_left, p_top, p_width, p_height, slide_w, slide_h):
    """Fraction of slide area that the picture's on-slide-visible region
    covers. Clamps to slide bounds (a picture that overflows the slide
    edges doesn't get credit for the off-slide portion)."""
    vis_left   = max(0, p_left)
    vis_top    = max(0, p_top)
    vis_right  = min(slide_w, p_left + p_width)
    vis_bottom = min(slide_h, p_top + p_height)
    vis_w = max(0, vis_right - vis_left)
    vis_h = max(0, vis_bottom - vis_top)
    return (vis_w * vis_h) / (slide_w * slide_h)
