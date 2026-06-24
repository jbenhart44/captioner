#!/usr/bin/env python3
"""fix_captions.py — per-slide, in-place corrections to an already-captioned deck.

The /captioner pipeline (extract -> caption -> apply -> verify) produces a deck where
every caption is a named text box: shape name `captioner_caption_<hash>` (normal) or
`captioner_capband_<hash>` (in-picture bottom band). This script lets a human who is
REVIEWING that finished deck make surgical fixes by SLIDE NUMBER — "slide 5's caption
box is too big", "slide 12's caption overlaps the photo, move it below", "slide 3's
caption is wrong" — WITHOUT re-running the whole pipeline or re-reading any images.

It edits ONLY the caption shapes on the named slides. Every other slide, every other
shape, and the pictures themselves are left byte-for-byte untouched. The original deck
is backed up to `<deck>.pptx.bak.pptx` before the first write (unless --no-backup).

Each adjusted caption is re-checked against the SAME 2D obstacle model the placer and
verify.py use (titles, body text, other pictures, other captions, footer, off-slide).
A correction that would CREATE a defect is refused and reported — unless --force.

------------------------------------------------------------------------------
USAGE — one or more correction groups, each opened by --slide N
------------------------------------------------------------------------------
  python3 fix_captions.py <deck.pptx> \
      --slide 5  --shrink-box \
      --slide 12 --move below \
      --slide 3  --font-size 8 --font-color 222222 \
      --slide 7  --text "Photo: solar farm at dusk"

  # affect a specific caption when a slide has more than one picture/caption:
  --slide 9 --which 2 --move right        # the 2nd caption on slide 9 (top-to-bottom)
  --slide 9 --hash a1b2c3d4e5f6 --shrink-box   # the caption for that image hash

Flags reset at each new --slide, so corrections are independent per group.

Resize        --height EMU | --height-in INCHES | --shrink-box (fit text)
              | --width EMU | --width-in INCHES
Reposition    --move {below|above|left|right|bottom}   (re-placed cleanly next to its picture)
              --nudge-x EMU --nudge-y EMU              (relative shift; +y = down)
Restyle       --font-size PT | --font-color HEX | --italic {true|false}
              --font-name NAME | --bg-color HEX (""=transparent) | --border-color HEX
Retext        --text "new caption wording"

  --which N        target the Nth caption on the slide (1-based, top-to-bottom). Default: all.
  --hash HEX       target only the caption whose shape-name ends with this image hash.
  --force          apply even if the result would overlap an obstacle (still reports it).
  --dry-run        report what WOULD change; write nothing.
  --no-backup      skip the .bak.pptx copy.
  --verify         after saving, run verify.py-equivalent gate on the deck (reports defects).

Exit codes: 0 clean · 4 a requested fix would create/leave a placement defect (and
--force was not given, so it was refused) · 2 usage / IO / no matching caption.
"""
import sys, os, argparse, shutil, copy

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_AUTO_SIZE

import _geometry as G
from _oxml_pics import iter_slide_pics, resolve_blob

CAPTION_PREFIX = 'captioner_caption_'
BAND_PREFIX    = 'captioner_capband_'
ALL_PREFIXES   = (CAPTION_PREFIX, BAND_PREFIX)
EMU_PER_INCH   = 914400


# ---------------------------------------------------------------------------
# Locating caption shapes on a slide
# ---------------------------------------------------------------------------
def caption_shapes(slide):
    """Captioner caption shapes on this slide, sorted top-to-bottom then left-to-right.

    Returns list of dicts: {shape, name, hash, is_band, left, top, width, height}."""
    out = []
    for sh in slide.shapes:
        try:
            name = sh.name or ''
        except Exception:
            continue
        if not name.startswith(ALL_PREFIXES):
            continue
        is_band = name.startswith(BAND_PREFIX)
        h = name[len(BAND_PREFIX if is_band else CAPTION_PREFIX):]
        out.append({
            'shape': sh, 'name': name, 'hash': h, 'is_band': is_band,
            'left': int(sh.left or 0), 'top': int(sh.top or 0),
            'width': int(sh.width or 0), 'height': int(sh.height or 0),
        })
    out.sort(key=lambda c: (c['top'], c['left']))
    return out


def own_picture_rect(slide, cap_hash):
    """(L,T,R,B) of the picture this caption belongs to (matched by image hash),
    or None if it cannot be matched (e.g. hash collision / linked image)."""
    import hashlib
    for pic in iter_slide_pics(slide):
        rid = pic.get('rid')
        ox, oy, cx, cy = (pic.get('off_x'), pic.get('off_y'),
                          pic.get('ext_cx'), pic.get('ext_cy'))
        if None in (ox, oy, cx, cy) or rid is None:
            continue
        try:
            digest = hashlib.sha256(resolve_blob(slide, rid)).hexdigest()[:12]
        except Exception:
            continue
        if digest == cap_hash:
            return (int(ox), int(oy), int(ox) + int(cx), int(oy) + int(cy))
    return None


# ---------------------------------------------------------------------------
# Obstacle model — identical population to the placer / verifier
# ---------------------------------------------------------------------------
def slide_obstacles(slide, slide_h, exclude_caption_shape):
    """All hard obstacles for a caption on this slide: every text region, every
    OTHER caption, every picture (except the caption's own), and the footer band.
    `exclude_caption_shape` is the shape currently being moved (so it doesn't block
    itself). Pictures the moved caption owns are NOT excluded here — repositioning
    'below/above' is allowed to sit adjacent to its own picture, handled by caller."""
    obstacles = list(G.slide_text_obstacle_rects(slide))
    # other captioner captions (compare underlying XML element — python-pptx
    # returns a NEW proxy object on every shapes walk, so `is` on the proxy is
    # unstable; `._element is` is the stable identity for self-exclusion).
    excl_el = getattr(exclude_caption_shape, '_element', None)
    for c in caption_shapes(slide):
        if c['shape']._element is excl_el:
            continue
        obstacles.append((c['left'], c['top'],
                          c['left'] + c['width'], c['top'] + c['height']))
    # footer band as a full-width vertical obstacle
    footer_top = G.slide_footer_top(slide, slide_h)
    if footer_top is not None and footer_top < slide_h:
        obstacles.append((0, footer_top, slide_h, slide_h))  # (L,T,R,B) full-width-ish
    return obstacles


def all_picture_rects(slide, exclude_hash=None):
    """(L,T,R,B) for every picture on the slide, optionally excluding the one whose
    hash matches `exclude_hash` (the caption's own picture)."""
    import hashlib
    rects = []
    for pic in iter_slide_pics(slide):
        ox, oy, cx, cy = (pic.get('off_x'), pic.get('off_y'),
                          pic.get('ext_cx'), pic.get('ext_cy'))
        if None in (ox, oy, cx, cy):
            continue
        rid = pic.get('rid')
        if exclude_hash and rid is not None:
            try:
                if hashlib.sha256(resolve_blob(slide, rid)).hexdigest()[:12] == exclude_hash:
                    continue
            except Exception:
                pass
        rects.append((int(ox), int(oy), int(ox) + int(cx), int(oy) + int(cy)))
    return rects


def rect_off_slide(rect, slide_w, slide_h):
    l, t, r, b = rect
    return l < 0 or t < 0 or r > slide_w or b > slide_h


def defects_for(rect, slide, slide_w, slide_h, exclude_shape, own_hash):
    """List of human-readable defect strings if `rect` (L,T,R,B) is placed on slide.
    Empty list == clean. Mirrors verify.py's pattern set."""
    problems = []
    if rect_off_slide(rect, slide_w, slide_h):
        problems.append('off-slide')
    text_obs = list(G.slide_text_obstacle_rects(slide))
    if not G.clear_all_obstacles_2d(rect, text_obs):
        problems.append('text-overlap')
    cap_obs = []
    excl_el = getattr(exclude_shape, '_element', None)
    for c in caption_shapes(slide):
        if c['shape']._element is excl_el:
            continue
        cap_obs.append((c['left'], c['top'], c['left'] + c['width'], c['top'] + c['height']))
    if not G.clear_all_obstacles_2d(rect, cap_obs):
        problems.append('caption-caption-overlap')
    # other pictures (own picture excluded — a caption may sit adjacent to its own)
    other_pics = all_picture_rects(slide, exclude_hash=own_hash)
    for p in other_pics:
        ia = G.rect_intersect_area((rect[0], rect[1], rect[2] - rect[0], rect[3] - rect[1]),
                                   (p[0], p[1], p[2] - p[0], p[3] - p[1]))
        cap_area = max(1, (rect[2] - rect[0]) * (rect[3] - rect[1]))
        if ia > 0.5 * cap_area:
            problems.append('inside-other-picture')
            break
    footer_top = G.slide_footer_top(slide, slide_h)
    if footer_top is not None and rect[3] > footer_top:
        problems.append('footer-intrusion')
    return problems


# ---------------------------------------------------------------------------
# The corrections
# ---------------------------------------------------------------------------
def estimate_shrunk_height(shape, width):
    """Height a SHAPE_TO_FIT box would take for its current text at `width`."""
    try:
        txt = shape.text_frame.text
    except Exception:
        txt = ''
    return G.estimate_caption_height(txt, width)


def candidate_rect_for_move(cap, pic_rect, direction, gap=50000):
    """Compute a new (L,T,R,B) for moving caption `cap` to `direction` of its picture.
    Preserves current width/height. Returns None if pic_rect unknown."""
    if pic_rect is None:
        return None
    pl, pt, pr, pb = pic_rect
    w, h = cap['width'], cap['height']
    if direction == 'below':
        return (pl, pb + gap, pl + w, pb + gap + h)
    if direction == 'above':
        return (pl, pt - gap - h, pl + w, pt - gap)
    if direction == 'right':
        cy = pt + (pb - pt - h) // 2
        return (pr + gap, cy, pr + gap + w, cy + h)
    if direction == 'left':
        cy = pt + (pb - pt - h) // 2
        return (pl - gap - w, cy, pl - gap, cy + h)
    if direction == 'bottom':  # inside picture bottom band
        return (pl, pb - h, pl + w, pb)
    return None


def apply_correction(slide, cap, g, slide_w, slide_h, force, dry):
    """Apply one correction group `g` (argparse Namespace) to one caption `cap`.
    Returns (changed: bool, messages: list[str], defect_after: list[str])."""
    sh = cap['shape']
    msgs = []
    # work on a trial rect; commit to the shape only if accepted
    L, T = cap['left'], cap['top']
    W, H = cap['width'], cap['height']

    # --- resize ---
    if g.width is not None:
        W = int(g.width)
    if g.width_in is not None:
        W = int(round(g.width_in * EMU_PER_INCH))
    if g.height is not None:
        H = int(g.height)
    if g.height_in is not None:
        H = int(round(g.height_in * EMU_PER_INCH))
    if g.shrink_box:
        H = estimate_shrunk_height(sh, W)
        msgs.append(f"shrink-box -> height {H} EMU (~{H/EMU_PER_INCH:.2f}in)")

    # --- reposition (absolute, by direction relative to own picture) ---
    if g.move:
        pic_rect = own_picture_rect(slide, cap['hash'])
        if pic_rect is None:
            msgs.append(f"WARN: cannot match own picture for hash {cap['hash']}; "
                        f"--move ignored (use --nudge-x/--nudge-y for a relative shift)")
        else:
            # use the (possibly resized) W/H for the moved box
            tmp = dict(cap); tmp['width'] = W; tmp['height'] = H
            rect = candidate_rect_for_move(tmp, pic_rect, g.move)
            if rect:
                L, T = rect[0], rect[1]
                W, H = rect[2] - rect[0], rect[3] - rect[1]
                msgs.append(f"move {g.move} -> ({L},{T})")

    # --- nudge (relative) ---
    if g.nudge_x:
        L += int(g.nudge_x); msgs.append(f"nudge-x {g.nudge_x}")
    if g.nudge_y:
        T += int(g.nudge_y); msgs.append(f"nudge-y {g.nudge_y}")

    new_rect = (L, T, L + W, T + H)

    # --- restyle / retext: these don't move the box; validate geometry only if it changed ---
    geometry_changed = (L, T, W, H) != (cap['left'], cap['top'], cap['width'], cap['height'])
    defect_after = defects_for(new_rect, slide, slide_w, slide_h, sh, cap['hash']) \
        if geometry_changed else []

    if geometry_changed and defect_after and not force:
        msgs.append(f"REFUSED geometry change — would create: {', '.join(defect_after)} "
                    f"(use --force to apply anyway)")
        return (False, msgs, defect_after)

    if dry:
        if geometry_changed:
            msgs.append(f"[dry-run] would set rect ({L},{T}) {W}x{H} EMU")
        return (geometry_changed or _has_style_change(g), msgs, defect_after)

    # commit geometry
    if geometry_changed:
        sh.left, sh.top, sh.width, sh.height = Emu(L), Emu(T), Emu(W), Emu(H)
        if defect_after and force:
            msgs.append(f"FORCED despite: {', '.join(defect_after)}")

    # commit style / text
    changed_style = _apply_style_text(sh, g, msgs)

    return (geometry_changed or changed_style, msgs, defect_after)


def _has_style_change(g):
    return any(v is not None for v in (g.font_size, g.font_color, g.italic,
                                       g.font_name, g.bg_color, g.border_color, g.text))


def _apply_style_text(sh, g, msgs):
    changed = False
    try:
        tf = sh.text_frame
    except Exception:
        if _has_style_change(g):
            msgs.append("WARN: shape has no text frame; style/text changes skipped")
        return False

    if g.text is not None:
        # preserve the first run's formatting; replace text content
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        # clear all paragraphs but the first, set first run text
        p0 = tf.paragraphs[0]
        for extra_p in list(tf.paragraphs[1:]):
            extra_p._p.getparent().remove(extra_p._p)
        if p0.runs:
            p0.runs[0].text = g.text
            for r in list(p0.runs[1:]):
                r._r.getparent().remove(r._r)
        else:
            p0.add_run().text = g.text
        changed = True
        msgs.append(f'text -> "{g.text}"')

    # style applies to every run in the box
    for para in tf.paragraphs:
        for run in para.runs:
            if g.font_size is not None:
                run.font.size = Pt(g.font_size); changed = True
            if g.font_color is not None:
                run.font.color.rgb = RGBColor.from_string(g.font_color); changed = True
            if g.italic is not None:
                run.font.italic = g.italic; changed = True
            if g.font_name is not None:
                run.font.name = g.font_name; changed = True
    if g.font_size is not None:  msgs.append(f"font-size {g.font_size}pt")
    if g.font_color is not None: msgs.append(f"font-color #{g.font_color}")
    if g.italic is not None:     msgs.append(f"italic {g.italic}")
    if g.font_name is not None:  msgs.append(f"font-name {g.font_name}")

    if g.bg_color is not None:
        if g.bg_color == '':
            try:
                sh.fill.background()
            except Exception:
                pass
            msgs.append("bg transparent")
        else:
            sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(g.bg_color)
            msgs.append(f"bg #{g.bg_color}")
        changed = True
    if g.border_color is not None:
        if g.border_color == '':
            try:
                sh.line.fill.background()
            except Exception:
                pass
            msgs.append("border off")
        else:
            sh.line.color.rgb = RGBColor.from_string(g.border_color)
            sh.line.width = Emu(6350)
            msgs.append(f"border #{g.border_color}")
        changed = True
    return changed


# ---------------------------------------------------------------------------
# CLI — parse interleaved --slide groups
# ---------------------------------------------------------------------------
def build_group_parser():
    p = argparse.ArgumentParser(add_help=False)
    p.add_argument('--slide', type=int, required=True)
    p.add_argument('--which', type=int, default=None,
                   help='1-based index of the caption on the slide (top-to-bottom). Default: all.')
    p.add_argument('--hash', dest='hash', default=None,
                   help='Target only the caption whose shape name ends with this image hash.')
    # resize
    p.add_argument('--height', type=int, default=None)
    p.add_argument('--height-in', dest='height_in', type=float, default=None)
    p.add_argument('--width', type=int, default=None)
    p.add_argument('--width-in', dest='width_in', type=float, default=None)
    p.add_argument('--shrink-box', dest='shrink_box', action='store_true')
    # reposition
    p.add_argument('--move', choices=['below', 'above', 'left', 'right', 'bottom'], default=None)
    p.add_argument('--nudge-x', dest='nudge_x', type=int, default=0)
    p.add_argument('--nudge-y', dest='nudge_y', type=int, default=0)
    # restyle
    p.add_argument('--font-size', dest='font_size', type=float, default=None)
    p.add_argument('--font-color', dest='font_color', default=None)
    p.add_argument('--italic', dest='italic', type=lambda v: v.lower() in ('1', 'true', 'yes', 'y'),
                   default=None)
    p.add_argument('--font-name', dest='font_name', default=None)
    p.add_argument('--bg-color', dest='bg_color', default=None)
    p.add_argument('--border-color', dest='border_color', default=None)
    # retext
    p.add_argument('--text', default=None)
    return p


def split_groups(argv):
    """Split argv into per-`--slide` chunks. Returns (deck_path, global_opts_argv, [group_argv,...])."""
    # first non-flag token before the first --slide is the deck path
    groups, cur = [], None
    deck = None
    globals_argv = []
    i = 0
    GLOBAL_FLAGS = {'--force', '--dry-run', '--no-backup', '--verify'}
    while i < len(argv):
        tok = argv[i]
        if tok == '--slide':
            if cur is not None:
                groups.append(cur)
            cur = ['--slide']
            i += 1
            continue
        if cur is None:
            # pre-slide region: deck path or a global flag
            if tok in GLOBAL_FLAGS:
                globals_argv.append(tok)
            elif deck is None and not tok.startswith('-'):
                deck = tok
            else:
                globals_argv.append(tok)
        else:
            if tok in GLOBAL_FLAGS:
                globals_argv.append(tok)
            else:
                cur.append(tok)
        i += 1
    if cur is not None:
        groups.append(cur)
    return deck, globals_argv, groups


def main():
    argv = sys.argv[1:]
    if not argv or argv[0] in ('-h', '--help'):
        print(__doc__); sys.exit(0 if argv else 2)

    deck, globals_argv, group_argvs = split_groups(argv)
    if not deck or not os.path.isfile(deck):
        print(f"ERROR: deck not found: {deck!r}"); sys.exit(2)
    if not group_argvs:
        print("ERROR: no --slide correction groups given (see --help)"); sys.exit(2)

    gp = argparse.ArgumentParser(add_help=False)
    gp.add_argument('--force', action='store_true')
    gp.add_argument('--dry-run', dest='dry_run', action='store_true')
    gp.add_argument('--no-backup', dest='no_backup', action='store_true')
    gp.add_argument('--verify', action='store_true')
    gopts, _ = gp.parse_known_args(globals_argv)

    grp_parser = build_group_parser()
    groups = [grp_parser.parse_args(ga) for ga in group_argvs]

    prs = Presentation(deck)
    slides = list(prs.slides)
    n_slides = len(slides)

    any_change = False
    any_refused = False
    print(f"Deck: {deck}  ({n_slides} slides)")
    print(f"Mode: {'DRY-RUN (no write)' if gopts.dry_run else 'EDIT IN PLACE'}"
          f"{'  [--force]' if gopts.force else ''}\n")

    for g in groups:
        if g.slide < 1 or g.slide > n_slides:
            print(f"slide {g.slide}: OUT OF RANGE (1..{n_slides}) — skipped"); any_refused = True
            continue
        slide = slides[g.slide - 1]
        caps = caption_shapes(slide)
        # filter by hash / which
        if g.hash:
            caps = [c for c in caps if c['hash'].endswith(g.hash) or c['hash'] == g.hash]
        if g.which is not None:
            caps = caps[g.which - 1: g.which] if 1 <= g.which <= len(caps) else []
        if not caps:
            print(f"slide {g.slide}: no matching captioner caption "
                  f"({'hash=' + g.hash + ' ' if g.hash else ''}"
                  f"{'which=' + str(g.which) if g.which is not None else 'any'}) — nothing to fix")
            any_refused = True
            continue
        for c in caps:
            label = f"slide {g.slide} [{c['name']}]"
            try:
                changed, msgs, defect = apply_correction(
                    slide, c, g, prs.slide_width, prs.slide_height,
                    force=gopts.force, dry=gopts.dry_run)
            except Exception as e:
                print(f"  {label}: ERROR {e}"); any_refused = True; continue
            tag = 'CHANGED' if changed else ('REFUSED' if defect else 'no-op')
            if changed:
                any_change = True
            if defect and not changed:
                any_refused = True
            print(f"  {label}: {tag}")
            for m in msgs:
                print(f"      - {m}")

    if any_change and not gopts.dry_run:
        if not gopts.no_backup:
            bak = deck + '.bak.pptx'
            try:
                shutil.copy(deck, bak)
                print(f"\nBackup: {bak}")
            except Exception as e:
                print(f"\nWARN: backup failed ({e}); proceeding to save anyway")
        prs.save(deck)
        print(f"Saved: {deck}")
    elif gopts.dry_run:
        print("\n(dry-run — no file written)")
    else:
        print("\nNo changes applied.")

    # optional post-fix verification gate
    if gopts.verify and not gopts.dry_run:
        defects = run_inline_verify(deck)
        if defects:
            print("\nVERIFY: remaining placement defects after fix:")
            for d in defects:
                print(f"  - {d}")
            sys.exit(4)
        print("\nVERIFY: clean — no caption placement defects.")

    sys.exit(4 if (any_refused and not gopts.force) else 0)


def run_inline_verify(deck):
    """Lightweight re-scan of the saved deck for the same defect classes, slide by
    slide. Returns a list of 'slide N: <defect>' strings (empty == clean)."""
    out = []
    prs = Presentation(deck)
    sw, sh_ = prs.slide_width, prs.slide_height
    for idx, slide in enumerate(prs.slides, 1):
        for c in caption_shapes(slide):
            rect = (c['left'], c['top'], c['left'] + c['width'], c['top'] + c['height'])
            d = defects_for(rect, slide, sw, sh_, c['shape'], c['hash'])
            # band captions intentionally sit in their own picture; suppress that one class
            if c['is_band']:
                d = [x for x in d if x not in ('inside-other-picture',)]
            for x in d:
                out.append(f"slide {idx} [{c['name']}]: {x}")
    return out


if __name__ == '__main__':
    main()
