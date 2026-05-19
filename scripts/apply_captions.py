"""Apply captions from captions.json to every deck listed in manifest.json.

(Side-effect imports: re used by SmartArt icon-name extractor below.)


Usage:
  python3 apply_captions.py <work_dir> [--dry-run] [--font-name FONT]
                            [--font-size PT] [--font-color HEX] [--italic BOOL]
                            [--gap-emu INT] [--height-emu INT]
                            [--update-existing] [--quiet]

Expects:
  <work_dir>/manifest.json    (from extract_images.py)
  <work_dir>/captions.json    (written by Claude after reading images)
      Format: {"<deck>/<hash>.<ext>": "<short caption>", ...}
      Special value "[decorative]" → skip (no visible text added, logged decorative)

Produces:
  <work_dir>/captioned_decks/<deck>_captioned.pptx
  <work_dir>/audit/<deck>_audit.csv

Improvements (from Gemini review v1/v2):
  - --dry-run: emits audit CSV but does NOT modify any .pptx
  - --font-name / --font-size / --font-color / --italic: caption styling
  - --gap-emu / --height-emu: layout offsets
  - Recursive traversal into GROUP shapes
  - Re-run protection: detects `_captioned` suffix on input + warns
  - Idempotency: caption shapes get name prefix "captioner_caption_<hash>" so re-runs can
    detect + with --update-existing flag, remove and re-add (else: append, leading to duplicates)
  - Per-deck progress prints
  - Per-deck try/except — corrupt/password-protected file does not halt batch
"""
import sys, os, json, shutil, csv, hashlib, argparse, re
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu, Pt
from _oxml_pics import iter_slide_pics, resolve_blob, guess_ext

# Fix: pic enumeration switched to the strict
# raw-OOXML slide-spTree `.//p:pic` walk (see _oxml_pics.py), matching
# extract/verify exactly. Placement geometry reads the pic's own
# <a:off>/<a:ext> EMU; verified byte-identical to the old pic.left/.top/
# .width/.height for every deck pic in this corpus (incl. group-nested),
# so caption-card positions do not move. Hashes unchanged.
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE

CAPTION_SHAPE_NAME_PREFIX = 'captioner_caption_'  # idempotency marker
SMARTART_SHAPE_NAME_PREFIX = 'captioner_smartart_'  # idempotency marker for SmartArt captions
SMARTART_ICON_SHAPE_NAME_PREFIX = 'captioner_sa_icon_'  # idempotency marker for per-icon SmartArt captions
MIN_CAPTION_HEIGHT = 250000

NS_A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
NS_DGM_REL = '{http://schemas.openxmlformats.org/drawingml/2006/diagram}'
DIAGRAM_DATA_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData'

# ---------------------------------------------------------------------------
# Spell-check (opt-in, --spellcheck). FLAG-ONLY: this never edits the .pptx and
# never auto-corrects a word. It emits suspected misspellings to a separate
# <deck>_spellcheck.csv for a human to review. A bundled, user-extensible
# whitelist (spellcheck_whitelist.txt) prevents flagging known domain terms,
# abbreviations, and proper nouns — so captioner does not "fix" non-issues.
# ---------------------------------------------------------------------------
WHITELIST_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'spellcheck_whitelist.txt')
# Proper-noun typos a generic dictionary won't suggest well — surfaced explicitly.
# Names here are web-verified canonical spellings (see SKILL.md "Name verification").
SPELL_KNOWN_BAD = {
    'humaoid': 'humanoid', 'humaoids': 'humanoids',
    'appronik': 'Apptronik',          # Apptronik (humanoid-robotics co.) — web-verified
    'geoffery': 'Geoffrey',           # Geoffrey Moore (Crossing the Chasm) — web-verified
    'clayten': 'Clayton',             # Clayton Christensen (HBS) — web-verified
    'siemans': 'Siemens',             # Siemens — web-verified
    'wolfrom': 'Wolfram',             # Wolfram Research — web-verified
}


def _load_whitelist():
    wl = set()
    try:
        with open(WHITELIST_PATH, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                for tok in line.split():
                    wl.add(tok.lower())
    except OSError:
        pass
    return wl


def init_spellcheck(enabled):
    """Return (SpellChecker|None, whitelist set, status str). Degrades gracefully
    if pyspellchecker is not installed (feature is optional)."""
    if not enabled:
        return None, set(), 'disabled'
    wl = _load_whitelist()
    try:
        from spellchecker import SpellChecker
    except ImportError:
        return None, wl, 'unavailable'  # documented optional dependency
    return SpellChecker(distance=1), wl, 'ready'


def spell_scan(text, source, s_idx, sp, wl, seen, rows):
    """Append suspected-misspelling rows for `text`. FLAG-ONLY — no mutation.
    `source` is 'caption' or 'slide-text'. `seen` dedupes (source, slide, word)."""
    if not text:
        return
    low_line = text.lower()
    if 'http' in low_line or 'www.' in low_line or '://' in low_line \
            or any(d in low_line for d in ('.com', '.org', '.ai', '.io',
                                           '.gov', '.edu', '.net', '.co/')):
        return  # URL / citation / product-domain line: not prose
    for tok in re.findall(r"[A-Za-z][A-Za-z'\-]+", text):
        low = tok.lower().strip("'-")
        if low in SPELL_KNOWN_BAD:
            sug, known = SPELL_KNOWN_BAD[low], True
        else:
            if len(low) < 4 or tok.isupper():
                continue
            if any(ch.isdigit() for ch in tok) or "'" in tok or '-' in tok:
                continue
            # Plural-of-acronym (NPVs, IRRs, CEOs, KPIs, MNCs, SVMs, CNNs, GPTs):
            # the singular form is an all-caps acronym -> not a misspelling.
            if low.endswith('s') and len(tok) >= 3 and tok[:-1].isupper():
                continue
            if low in wl or sp is None:
                continue
            if not sp.unknown([low]):
                continue
            sug = sp.correction(low)
            if not sug or sug == low:
                continue
            known = False
        # Likely proper noun (Capitalized, not all-caps) -> the agent must
        # web-verify the canonical spelling before presenting it as a fix.
        verify_name = bool(tok[:1].isupper() and not tok.isupper())
        key = (source, s_idx, low)
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            'slide': s_idx, 'source': source, 'term': tok,
            'suggestion': sug, 'known_bad': known,
            'verify_name': verify_name,
            'context': text.strip()[:110].replace('\n', ' '),
        })


_QC_PLACEHOLDER = ("click to add", "lorem ipsum", "[gap", "tbd", "xxx",
                   "placeholder text", "insert text here", "your text here")
_QC_DOUBLE = re.compile(r"\b(\w{3,})\s+\1\b", re.IGNORECASE)
_QC_DATE = re.compile(r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
                      r"[a-z]*\.?\s+\d{1,2}(?:,?\s*\d{4})?\b")


def qc_scan(text, source, s_idx, seen, rows):
    """READ-ONLY date / doubled-word / leftover-template scan. FLAG-ONLY — never
    edits the .pptx. Generic (no course-specific year/code rules). `seen` dedupes.
    Date strings are tagged 'date-review' (informational, not asserted wrong)."""
    if not text:
        return
    low = text.lower()
    for ph in _QC_PLACEHOLDER:
        if ph in low:
            k = ('placeholder', s_idx, ph)
            if k not in seen:
                seen.add(k)
                rows.append({'slide': s_idx, 'source': source, 'kind': 'leftover-template',
                             'detail': ph, 'context': text.strip()[:110].replace('\n', ' ')})
    for m in set(w.lower() for w in _QC_DOUBLE.findall(text)):
        if m in ('that', 'had', 'is', 'the'):  # common legitimate repeats
            continue
        k = ('double', s_idx, m)
        if k not in seen:
            seen.add(k)
            rows.append({'slide': s_idx, 'source': source, 'kind': 'doubled-word',
                         'detail': f"{m} {m}", 'context': text.strip()[:110].replace('\n', ' ')})
    for m in _QC_DATE.findall(text):
        k = ('date', s_idx, m.strip())
        if k not in seen:
            seen.add(k)
            rows.append({'slide': s_idx, 'source': source, 'kind': 'date-review',
                         'detail': m.strip(), 'context': text.strip()[:110].replace('\n', ' ')})


def resolve_ph_geometry(slide, ph_idx):
    """A picture content-placeholder's geometry is inherited from the layout/
    master; the <p:pic> itself often has NO <a:xfrm>. python-pptx resolves the
    inheritance, so match the slide placeholder by idx and read its effective
    box. Returns (left, top, width, height) EMU or None. (A prior
    off-page-caption bug: placeholder pics fell back to (0, 50000).)"""
    try:
        for ph in slide.placeholders:
            pf = ph.placeholder_format
            if pf is not None and pf.idx == ph_idx:
                if None not in (ph.left, ph.top, ph.width, ph.height):
                    return int(ph.left), int(ph.top), int(ph.width), int(ph.height)
    except Exception:
        pass
    return None


def slide_footer_top(slide, slide_h):
    """Y (EMU) below which captions must NOT extend — the top of the reserved
    bottom band. Considers FOOTER/DATE/SLIDE_NUMBER placeholders on the slide,
    its layout, and master, plus any low-band text shape (catches branded
    footers like a 'Pearson' credit). Always reserves at least the bottom
    ~0.3in so captions never sit in the extreme bottom strip."""
    from pptx.enum.shapes import PP_PLACEHOLDER
    foot_types = {PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE,
                  PP_PLACEHOLDER.SLIDE_NUMBER}
    limit = slide_h - 274320  # ~0.30in universal bottom safety margin
    band = int(slide_h * 0.86)  # "low band" = bottom ~14%
    sources = [slide]
    try:
        sources.append(slide.slide_layout)
        sources.append(slide.slide_layout.slide_master)
    except Exception:
        pass
    for src in sources:
        try:
            shapes = list(src.placeholders) + [s for s in src.shapes
                                               if s not in src.placeholders]
        except Exception:
            try:
                shapes = list(src.shapes)
            except Exception:
                continue
        for sh in shapes:
            try:
                top = sh.top
                if top is None:
                    continue
                is_foot = False
                try:
                    pf = sh.placeholder_format
                    if pf is not None and pf.type in foot_types:
                        is_foot = True
                except Exception:
                    pass
                has_txt = bool(getattr(sh, 'has_text_frame', False)
                               and sh.text_frame.text.strip())
                # A footer is in the BOTTOM band. A footer/date/slide-number
                # placeholder parked at the TOP of a master (top < band) is
                # NOT a bottom footer — ignore it, else the reserved band
                # collapses to the slide top.
                if top >= band and (is_foot or has_txt):
                    limit = min(limit, int(top))
            except Exception:
                continue
    return max(0, limit)


def slide_title_box(slide):
    """Bounding box (l, t, r, b) EMU enclosing every TITLE/CENTER_TITLE
    placeholder that has visible text, wherever it sits on the slide (some
    layouts park the title low, under an image). None if no titled text.
    Captions must avoid this rect when a clear slot exists."""
    from pptx.enum.shapes import PP_PLACEHOLDER
    title_types = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
    L = T = R = B = None
    for sh in slide.shapes:
        if (sh.name or '').startswith('captioner_caption_'):
            continue
        try:
            pf = sh.placeholder_format
            if pf is None or pf.type not in title_types:
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


def _voverlap(c_top, c_h, title):
    """True if a caption [c_top, c_top+c_h] vertically overlaps the title rect
    by a meaningful amount (>15% of caption height)."""
    if title is None:
        return False
    _, tT, _, tB = title
    iy = max(0, min(c_top + c_h, tB) - max(c_top, tT))
    return iy > 0.15 * c_h


def iter_slide_body_text(shapes):
    """Yield instructor-authored text (text frames + tables), recursing groups,
    skipping captioner's own added shapes."""
    for sh in shapes:
        try:
            name = sh.name or ''
        except Exception:
            name = ''
        if name.startswith((CAPTION_SHAPE_NAME_PREFIX, SMARTART_SHAPE_NAME_PREFIX,
                            SMARTART_ICON_SHAPE_NAME_PREFIX)):
            continue
        try:
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                yield from iter_slide_body_text(sh.shapes)
                continue
            if sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            yield cell.text
                continue
        except Exception:
            pass
        if getattr(sh, 'has_text_frame', False):
            t = sh.text_frame.text
            if t and t.strip():
                yield t


def iter_smartart_frames(slide):
    """Yield graphicFrame shapes that contain SmartArt (diagram) content."""
    for sh in slide.shapes:
        try:
            if not sh._element.tag.endswith('}graphicFrame'):
                continue
            for child in sh._element.iter():
                ctag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if ctag == 'graphicData' and 'diagram' in (child.get('uri') or ''):
                    yield sh
                    break
        except Exception:
            continue


def _camel_to_words(s):
    """Icons_VideoCamera -> 'video camera'. CheckMark -> 'check mark'."""
    s = s.replace('Icons_', '').replace('Icon_', '')
    # Insert space before each capital letter (except first), then lowercase
    out = re.sub(r'(?<!^)(?=[A-Z])', ' ', s).lower()
    return out.strip()


def extract_smartart_icons(slide_part, frame):
    """Extract icon names from SVG metadata embedded in the SmartArt's drawing part.

    Returns a deduped list of human-readable icon names in order of first appearance.
    Strategy: follow graphicFrame -> diagramDrawing rel -> read drawing*.xml's rels ->
    for each SVG image relationship, parse the SVG file and extract the `id` attribute
    (PowerPoint stores names like 'Icons_Checkmark' on the root <svg> element).
    """
    # Find rId of the diagramDrawing relationship
    drawing_rid = None
    for child in frame._element.iter():
        if child.tag.endswith('}relIds'):
            for k, v in child.attrib.items():
                # ms diagram-drawing rel is 'dgm' style — older r:dm references the data part,
                # but the drawing part is referenced via a 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing'
                # which lives under the diagramData's rels, NOT the slide's rels.
                if k.endswith('}dm') or k == 'dm':
                    drawing_rid = v
                    break
            if drawing_rid:
                break
    if not drawing_rid:
        return []
    try:
        # Slide -> diagramData part
        data_rel = slide_part.rels[drawing_rid]
        data_part = data_rel.target_part
        # Images may be linked directly from diagramData (common case) OR via a separate
        # diagramDrawing part (less common). Collect candidates from both.
        candidate_parts = [data_part]
        for rel in data_part.rels.values():
            if 'diagramDrawing' in rel.reltype:
                candidate_parts.append(rel.target_part)
        icons = []
        seen_targets = set()
        for cp in candidate_parts:
            for rel in cp.rels.values():
                if rel.reltype.endswith('/image'):
                    target = rel.target_part
                    pname = str(getattr(target, 'partname', '') or '').lower()
                    if not pname.endswith('.svg'):
                        continue
                    if pname in seen_targets:
                        continue
                    seen_targets.add(pname)
                    try:
                        blob = target.blob
                    except Exception:
                        continue
                    text = blob.decode('utf-8', errors='ignore')
                    m = re.search(r'<svg[^>]*\bid="([^"]+)"', text)
                    if not m:
                        continue
                    name = _camel_to_words(m.group(1))
                    if name and name not in icons:
                        icons.append(name)
        return icons
    except Exception:
        return []


def extract_smartart_icon_placements(slide_part, frame):
    """For each icon embedded in this SmartArt, return its name + slide-absolute bounding box.

    Returns list of dicts: {name, x, y, cx, cy} in EMU (slide-absolute).
    Strategy: walk slide.part.rels for the diagramDrawing part (Microsoft cache that
    contains the rendered SmartArt with explicit per-shape geometry), iterate its <dsp:sp>
    elements, follow each shape's <a:blip r:embed=...> to the SVG image, parse the SVG id
    to get the icon name, and compute slide-absolute coords by adding the frame offset.
    """
    NS_DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram'
    DRAWING_RELTYPE = 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing'
    try:
        # Find the diagramDrawing part referenced by this slide (PowerPoint caches it
        # under the slide's rels, not the diagramData's rels)
        drawing_part = None
        for rel in slide_part.rels.values():
            if rel.reltype == DRAWING_RELTYPE:
                drawing_part = rel.target_part
                break
        if drawing_part is None:
            return []
        # Build rId -> svg-id lookup for this drawing part
        rid_to_svg_name = {}
        for rel in drawing_part.rels.values():
            if not rel.reltype.endswith('/image'):
                continue
            tp = rel.target_part
            if not str(getattr(tp, 'partname', '')).lower().endswith('.svg'):
                continue
            try:
                text = tp.blob.decode('utf-8', errors='ignore')
            except Exception:
                continue
            m = re.search(r'<svg[^>]*\bid="([^"]+)"', text)
            if m:
                rid_to_svg_name[rel.rId] = _camel_to_words(m.group(1))
        if not rid_to_svg_name:
            return []
        from lxml import etree
        root = etree.fromstring(drawing_part.blob)
        out = []
        f_left, f_top = frame.left, frame.top
        for sp in root.iter(f'{{{NS_DSP}}}sp'):
            # Prefer <asvg:svgBlip r:embed=rId> (points at the SVG which carries the icon name);
            # fall back to <a:blip> (points at the PNG, which has no name) for shapes without SVG.
            embed_rid = None
            for el in sp.iter():
                ltag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
                if ltag == 'svgBlip':
                    for k, v in el.attrib.items():
                        if k.endswith('}embed') or k == 'embed':
                            embed_rid = v
                            break
                    if embed_rid:
                        break
            if not embed_rid:
                for blip in sp.iter(f'{NS_A}blip'):
                    for k, v in blip.attrib.items():
                        if k.endswith('}embed') or k == 'embed':
                            embed_rid = v
                            break
                    if embed_rid:
                        break
            if not embed_rid or embed_rid not in rid_to_svg_name:
                continue
            # Find the xfrm geometry
            off_x = off_y = ext_cx = ext_cy = None
            for xfrm in sp.iter():
                tag = xfrm.tag.split('}')[-1]
                if tag != 'xfrm':
                    continue
                for child in xfrm:
                    ctag = child.tag.split('}')[-1]
                    if ctag == 'off':
                        off_x = int(child.get('x', 0))
                        off_y = int(child.get('y', 0))
                    elif ctag == 'ext':
                        ext_cx = int(child.get('cx', 0))
                        ext_cy = int(child.get('cy', 0))
                if off_x is not None and ext_cx is not None:
                    break
            if off_x is None or ext_cx is None:
                continue
            out.append({
                'name': rid_to_svg_name[embed_rid],
                'x': f_left + off_x,
                'y': f_top + off_y,
                'cx': ext_cx,
                'cy': ext_cy,
            })
        return out
    except Exception:
        return []


def remove_previous_smartart_icon_caption_shapes(slide):
    spTree = slide.shapes._spTree
    to_remove = []
    for sh in slide.shapes:
        try:
            name = sh.name or ''
        except Exception:
            continue
        if name.startswith(SMARTART_ICON_SHAPE_NAME_PREFIX):
            to_remove.append(sh._element)
    for el in to_remove:
        spTree.remove(el)
    return len(to_remove)


def extract_smartart_text(slide_part, frame):
    """Resolve the SmartArt's diagramData relationship and return list of visible texts.

    Returns [] if the diagram data part cannot be resolved or contains no text.
    """
    # Find the rId of the diagramData relationship referenced by this graphicFrame
    rid = None
    for child in frame._element.iter():
        # dgm:relIds element has r:dm attribute pointing at diagramData rId
        if child.tag.endswith('}relIds'):
            for k, v in child.attrib.items():
                if k.endswith('}dm') or k == 'dm':
                    rid = v
                    break
            if rid:
                break
    if not rid:
        return []
    try:
        rel = slide_part.rels[rid]
        if DIAGRAM_DATA_URI not in rel.reltype:
            return []
        data_part = rel.target_part
        from lxml import etree
        root = etree.fromstring(data_part.blob)
        texts = []
        for t in root.iter(f'{NS_A}t'):
            if t.text and t.text.strip():
                texts.append(t.text.strip())
        # Dedupe consecutive duplicates while preserving order
        out = []
        for t in texts:
            if not out or out[-1] != t:
                out.append(t)
        return out
    except Exception:
        return []


def generate_smartart_caption(texts, icons=None):
    """Build a deterministic accessibility caption from SmartArt content + icon names."""
    icons = icons or []
    if not texts and not icons:
        return None
    parts = []
    if icons:
        parts.append(f"Diagram ({', '.join(icons)} icons):")
    else:
        parts.append("Diagram:")
    if texts:
        parts.append("; ".join(texts) + ".")
    cap = " ".join(parts)
    if len(cap) > 180:
        cap = cap[:177].rstrip() + '...'
    return cap


def remove_previous_smartart_caption_shapes(slide):
    """Remove SmartArt caption text boxes previously added by /captioner."""
    spTree = slide.shapes._spTree
    to_remove = []
    for sh in slide.shapes:
        try:
            name = sh.name or ''
        except Exception:
            continue
        if name.startswith(SMARTART_SHAPE_NAME_PREFIX):
            to_remove.append(sh._element)
    for el in to_remove:
        spTree.remove(el)
    return len(to_remove)


def iter_pictures_recursive(shapes, depth=0):
    """Yield (shape, depth) for every Picture; recurse into GROUP shapes."""
    for sh in shapes:
        try:
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                yield sh, depth
            elif sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                try:
                    yield from iter_pictures_recursive(sh.shapes, depth + 1)
                except NotImplementedError:
                    continue
        except Exception:
            continue


def remove_previous_caption_shapes(slide, target_pic_hash=None):
    """Remove caption text boxes previously added by /captioner.

    If target_pic_hash given, only remove captions matching that picture's hash
    (encoded in the shape name suffix). Else remove all captioner captions on slide.
    Returns count removed.
    """
    spTree = slide.shapes._spTree
    to_remove = []
    for sh in slide.shapes:
        name = ''
        try:
            name = sh.name or ''
        except Exception:
            continue
        if name.startswith(CAPTION_SHAPE_NAME_PREFIX):
            if target_pic_hash is None or name.endswith(target_pic_hash):
                to_remove.append(sh._element)
    for el in to_remove:
        spTree.remove(el)
    return len(to_remove)


def apply_to_deck(deck_info, captions, captioned_dir, audit_dir, style, opts):
    deck = deck_info['deck']
    src = deck_info['deck_path']

    # Re-run protection: warn if source already looks like a captioner output
    src_base = os.path.basename(src)
    if src_base.endswith('_captioned.pptx'):
        print(f"  WARN: input {src_base} already has _captioned suffix; output will be {src_base[:-5]}_captioned.pptx")

    dst_name = f"{deck}_captioned.pptx" if not deck.endswith('_captioned') else f"{deck}.pptx"
    dst = os.path.join(captioned_dir, dst_name)

    if opts['dry_run']:
        # Don't copy; we'll only inspect
        prs = Presentation(src)
    else:
        shutil.copy(src, dst)
        prs = Presentation(dst)

    slide_w = prs.slide_width
    slide_h = prs.slide_height
    audit_rows = []
    n_caption_shapes_removed = 0
    spell_rows = []
    spell_seen = set()
    qc_rows = []
    qc_seen = set()
    sp_engine, sp_wl = opts.get('_sp'), opts.get('_wl', set())

    for s_idx, slide in enumerate(prs.slides, 1):
        if opts['spellcheck'] or opts['dateqc']:
            for body_text in iter_slide_body_text(slide.shapes):
                if opts['spellcheck']:
                    spell_scan(body_text, 'slide-text', s_idx, sp_engine, sp_wl,
                               spell_seen, spell_rows)
                if opts['dateqc']:
                    qc_scan(body_text, 'slide-text', s_idx, qc_seen, qc_rows)
        # Idempotency: if --update-existing AND not dry-run, strip any prior captioner shapes
        # on this slide before adding new ones.
        if opts['update_existing'] and not opts['dry_run']:
            n_caption_shapes_removed += remove_previous_caption_shapes(slide)
            n_caption_shapes_removed += remove_previous_smartart_caption_shapes(slide)
            n_caption_shapes_removed += remove_previous_smartart_icon_caption_shapes(slide)

        # SmartArt frames: caption deterministically from extracted text content
        if opts['caption_smartart']:
            for sa_idx, frame in enumerate(iter_smartart_frames(slide)):
                placements = extract_smartart_icon_placements(slide.part, frame)
                # Per-icon caption boxes: short label directly under each icon.
                # Text-only SmartArts (no icons) get NO caption — the visible text is
                # already accessible via the slide's own text layer.
                if placements and not opts['dry_run']:
                    icon_h = 200000  # ~0.22"
                    icon_gap = 20000  # tiny gap below icon
                    for ic_idx, p in enumerate(placements):
                        ic_top = p['y'] + p['cy'] + icon_gap
                        if ic_top + icon_h > slide_h:
                            ic_top = max(0, p['y'] - icon_gap - icon_h)
                        ic_left = p['x']
                        ic_width = p['cx']
                        if ic_width < 400000:
                            ic_width = 400000
                        itb = slide.shapes.add_textbox(ic_left, ic_top, ic_width, icon_h)
                        try:
                            itb._element.nvSpPr.cNvPr.set('name', f"{SMARTART_ICON_SHAPE_NAME_PREFIX}{sa_idx}_{ic_idx}")
                        except Exception:
                            pass
                        if style['bg_color']:
                            itb.fill.solid()
                            itb.fill.fore_color.rgb = RGBColor.from_string(style['bg_color'])
                        if style['border_color']:
                            itb.line.color.rgb = RGBColor.from_string(style['border_color'])
                            itb.line.width = Emu(6350)
                        itf = itb.text_frame
                        itf.word_wrap = True
                        # Auto-grow vertical to fit wrapped text (icons can be narrow)
                        itf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
                        itf.margin_left = Emu(20000); itf.margin_right = Emu(20000)
                        itf.margin_top = Emu(10000); itf.margin_bottom = Emu(10000)
                        ip = itf.paragraphs[0]
                        ip.alignment = PP_ALIGN.CENTER
                        irun = ip.add_run()
                        irun.text = p['name']
                        irun.font.size = Pt(8)
                        irun.font.italic = style['italic']
                        irun.font.name = style['font_name']
                        irun.font.color.rgb = RGBColor.from_string(style['font_color'])
                        audit_rows.append({
                            'slide': s_idx, 'pic_id': '', 'image_hash': f'smartart_{sa_idx}_icon_{ic_idx}',
                            'caption': p['name'], 'char_len': len(p['name']), 'in_group_depth': 0,
                            'action': 'added-smartart-icon',
                        })
                # Note: SmartArts with NO icons (text-only diagrams) get NO caption.
                # The visible text already lives in the slide's accessible text layer.
                if not placements:
                    audit_rows.append({
                        'slide': s_idx, 'pic_id': '', 'image_hash': f'smartart_{sa_idx}',
                        'caption': '', 'char_len': 0, 'in_group_depth': 0,
                        'action': 'skipped-smartart-text-only',
                    })

        footer_limit = slide_footer_top(slide, slide_h)
        title_box = slide_title_box(slide)
        for pic in iter_slide_pics(slide):
            depth = pic['depth']
            pic_id = pic['pic_id']
            if pic['rid'] is None:
                # Linked/external pic (r:link, no r:embed) or malformed blip:
                # cannot hash bytes -> skip + structured audit row (parity
                # with extract_images.py, which also skips + logs these).
                audit_rows.append({
                    'slide': s_idx, 'pic_id': pic_id, 'image_hash': '',
                    'caption': '', 'char_len': 0, 'in_group_depth': depth,
                    'action': 'skipped-linked-no-embed',
                })
                continue
            try:
                blob = resolve_blob(slide, pic['rid'])
                ext = guess_ext(slide.part.related_part(pic['rid']))
            except Exception:
                audit_rows.append({
                    'slide': s_idx, 'pic_id': pic_id, 'image_hash': '',
                    'caption': '', 'char_len': 0, 'in_group_depth': depth,
                    'action': 'skipped-image-extract-failed',
                })
                continue
            h = hashlib.sha256(blob).hexdigest()[:12]
            key = f"{deck}/{h}.{ext}"
            caption = captions.get(key)

            if caption is None:
                audit_rows.append({
                    'slide': s_idx, 'pic_id': pic_id, 'image_hash': h,
                    'caption': '', 'char_len': 0, 'in_group_depth': depth,
                    'action': 'skipped-no-caption',
                })
                continue
            if caption.strip().lower() in ('[decorative]', 'decorative', ''):
                audit_rows.append({
                    'slide': s_idx, 'pic_id': pic_id, 'image_hash': h,
                    'caption': '[decorative]', 'char_len': 0, 'in_group_depth': depth,
                    'action': 'skipped-decorative',
                })
                continue

            if opts['spellcheck']:
                spell_scan(caption, 'caption', s_idx, sp_engine, sp_wl,
                           spell_seen, spell_rows)
            if opts['dateqc']:
                qc_scan(caption, 'caption', s_idx, qc_seen, qc_rows)

            # Compute placement. Raw <a:off>/<a:ext> EMU — proven byte-identical
            # to the old python-pptx pic.left/.top/.width/.height for every deck
            # pic in this corpus (incl. group-nested), so caption cards do not
            # move. Defensive 0 default if a pic somehow lacks an xfrm (the
            # off-slide fallback-placement logic below then handles it).
            p_left = pic['off_x']
            p_top = pic['off_y']
            p_width = pic['ext_cx']
            p_height = pic['ext_cy']
            # Placeholder-hosted picture: geometry inherited from the layout.
            # Resolve it via the python-pptx placeholder, else the caption
            # lands off-page at (0, 50000).
            if (None in (p_left, p_top, p_width, p_height)
                    and pic.get('is_placeholder') and pic.get('ph_idx') is not None):
                g = resolve_ph_geometry(slide, pic['ph_idx'])
                if g is not None:
                    p_left, p_top, p_width, p_height = g
            # Last-resort: center a default box on the slide rather than the
            # old degenerate top-left stub.
            if None in (p_left, p_top, p_width, p_height):
                p_width = p_width or min(4000000, slide_w)
                p_height = p_height or 0
                p_left = p_left if p_left is not None else (slide_w - p_width) // 2
                p_top = p_top if p_top is not None else int(slide_h * 0.55)

            gap = opts['gap_emu']
            c_height = opts['height_emu']
            c_left = max(0, min(p_left, slide_w - 500000))
            c_width = min(p_width, slide_w - c_left)
            if c_width < 500000:
                c_width = min(500000, slide_w)
            if c_left + c_width > slide_w:
                c_left = max(0, slide_w - c_width)

            # Preferred: directly below the picture, but never into the
            # reserved bottom/footer band.
            below_top = p_top + p_height + gap
            above_top = p_top - gap - c_height
            below_ok = below_top + MIN_CAPTION_HEIGHT <= footer_limit
            below_h = min(c_height, footer_limit - below_top) if below_ok else 0
            forced_h = max(MIN_CAPTION_HEIGHT,
                           min(c_height, footer_limit - p_top - p_height - gap))
            forced_top = max(0, footer_limit - forced_h)
            # Candidate slots in preference order; first that CLEARS the title
            # (wherever it sits) wins. If none clear it, fall back to the first
            # geometrically valid slot (title overlap then truly unavoidable —
            # the picture itself occupies the rest of the slide).
            cands = []
            if below_ok:
                cands.append(('added', below_top, below_h))
            if above_top >= 0:
                cands.append(('fallback-above', above_top, c_height))
            cands.append(('fallback-bottom', forced_top, forced_h))
            pick = next((c for c in cands
                         if not _voverlap(c[1], c[2], title_box)), cands[0])
            action, c_top, c_height = pick
            # Final hard clamp — guarantee fully on-slide, never negative.
            c_top = max(0, min(c_top, slide_h - MIN_CAPTION_HEIGHT))
            if c_top + c_height > slide_h:
                c_height = max(MIN_CAPTION_HEIGHT, slide_h - c_top)

            if opts['dry_run']:
                audit_rows.append({
                    'slide': s_idx, 'pic_id': pic_id, 'image_hash': h,
                    'caption': caption, 'char_len': len(caption), 'in_group_depth': depth,
                    'action': f'dry-run-would-{action}',
                })
                continue

            # Add the text box
            tb = slide.shapes.add_textbox(c_left, c_top, c_width, c_height)
            # Idempotency: tag the shape name with our prefix + image hash
            try:
                tb._element.nvSpPr.cNvPr.set('name', f"{CAPTION_SHAPE_NAME_PREFIX}{h}")
            except Exception:
                pass  # if naming fails, the caption still renders correctly

            # Solid background fill — ensures caption is readable on dark slide backgrounds
            if style['bg_color']:
                tb.fill.solid()
                tb.fill.fore_color.rgb = RGBColor.from_string(style['bg_color'])
            # Thin border so the caption "card" reads cleanly on any background
            if style['border_color']:
                tb.line.color.rgb = RGBColor.from_string(style['border_color'])
                tb.line.width = Emu(6350)  # 0.5pt

            tf = tb.text_frame
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
            tf.margin_left = Emu(50000); tf.margin_right = Emu(50000)
            tf.margin_top = Emu(20000); tf.margin_bottom = Emu(20000)
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = caption
            run.font.size = Pt(style['font_size'])
            run.font.italic = style['italic']
            run.font.name = style['font_name']
            run.font.color.rgb = RGBColor.from_string(style['font_color'])

            audit_rows.append({
                'slide': s_idx, 'pic_id': pic_id, 'image_hash': h,
                'caption': caption, 'char_len': len(caption), 'in_group_depth': depth,
                'action': action,
            })

    if not opts['dry_run']:
        prs.save(dst)
    if audit_rows:
        # Dry-run writes its OWN file so it can never clobber a prior real
        # apply audit.
        audit_name = f"{deck}_audit_dryrun.csv" if opts['dry_run'] else f"{deck}_audit.csv"
        csv_path = os.path.join(audit_dir, audit_name)
        with open(csv_path, 'w', newline='') as f:
            w = csv.DictWriter(f, fieldnames=audit_rows[0].keys())
            w.writeheader()
            for r in audit_rows:
                w.writerow(r)

    # QC artifacts go to their OWN directory, never the caption audit dir.
    qc_dir = opts['qc_dir']
    if opts['spellcheck'] and spell_rows:
        sc_path = os.path.join(qc_dir, f"{deck}_spellcheck.csv")
        with open(sc_path, 'w', newline='') as f:
            w = csv.DictWriter(f, fieldnames=['slide', 'source', 'term',
                                              'suggestion', 'known_bad',
                                              'verify_name', 'context'])
            w.writeheader()
            for r in spell_rows:
                w.writerow(r)
    if opts['dateqc'] and qc_rows:
        qp = os.path.join(qc_dir, f"{deck}_qc.csv")
        with open(qp, 'w', newline='') as f:
            w = csv.DictWriter(f, fieldnames=['slide', 'source', 'kind',
                                              'detail', 'context'])
            w.writeheader()
            for r in qc_rows:
                w.writerow(r)

    return {
        'deck': deck, 'dst': dst if not opts['dry_run'] else None,
        'rows': audit_rows, 'n_caption_shapes_removed': n_caption_shapes_removed,
        'spell_rows': spell_rows, 'qc_rows': qc_rows,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('work_dir', help='Directory containing manifest.json and captions.json')
    ap.add_argument('--dry-run', action='store_true', help='Compute caption placements + write audit CSV but do NOT modify any .pptx')
    ap.add_argument('--font-name', default='Calibri', help='Caption font (default: Calibri)')
    ap.add_argument('--font-size', type=float, default=10.0, help='Caption font size in pt (default: 10)')
    ap.add_argument('--font-color', default='333333', help='Caption hex color, no # (default: 333333 dark gray)')
    ap.add_argument('--bg-color', default='FFFFFF', help='Caption text box fill hex, no # (default: FFFFFF white; "" to disable for transparent)')
    ap.add_argument('--border-color', default='CCCCCC', help='Caption text box border hex, no # (default: CCCCCC light gray; "" to disable)')
    ap.add_argument('--italic', type=lambda v: v.lower() in ('1','true','yes','y'), default=True, help='Caption italic (default: true)')
    ap.add_argument('--gap-emu', type=int, default=50000, help='EMU between picture and caption (default: 50000 = ~0.05")')
    ap.add_argument('--height-emu', type=int, default=400000, help='Caption text box height in EMU (default: 400000 = ~0.44")')
    ap.add_argument('--update-existing', action='store_true', help='Detect and remove previously-added captioner shapes before adding new ones (else re-runs DUPLICATE)')
    ap.add_argument('--no-smartart', action='store_true', help='Disable SmartArt captioning (SmartArt captions are auto-generated from diagram text content, on by default)')
    ap.add_argument('--quiet', action='store_true', help='Suppress per-deck progress lines')
    # QC is ON BY DEFAULT (spell-check + name-verify tagging + date/template QC).
    ap.add_argument('--quick', action='store_true', help='Captioning ONLY — skip ALL QC (spell-check + date/template scan). Use when you explicitly do not want QC.')
    ap.add_argument('--no-spellcheck', action='store_true', help='Disable just the spell-check pass (date/template QC still runs unless --quick).')
    ap.add_argument('--no-dateqc', action='store_true', help='Disable just the date/doubled-word/leftover-template scan (spell-check still runs unless --quick).')
    ap.add_argument('--spellcheck', action='store_true', help=argparse.SUPPRESS)  # back-compat no-op: spell-check is now default-on
    args = ap.parse_args()
    # Default-on QC with granular + master (--quick) toggles.
    qc_on = not args.quick
    do_spell = qc_on and not args.no_spellcheck
    do_dateqc = qc_on and not args.no_dateqc

    work = os.path.abspath(args.work_dir)
    with open(os.path.join(work, 'manifest.json')) as f:
        manifest = json.load(f)
    captions_path = os.path.join(work, 'captions.json')
    if not os.path.exists(captions_path):
        print(f"ERROR: {captions_path} not found. Write it after reading images via Read tool.")
        sys.exit(2)
    with open(captions_path) as f:
        captions = json.load(f)

    captioned_dir = os.path.join(work, 'captioned_decks')
    audit_dir = os.path.join(work, 'audit')
    qc_dir = os.path.join(work, 'qc')
    os.makedirs(captioned_dir, exist_ok=True)
    os.makedirs(audit_dir, exist_ok=True)
    if do_spell or do_dateqc:
        os.makedirs(qc_dir, exist_ok=True)

    style = {
        'font_name': args.font_name, 'font_size': args.font_size,
        'font_color': args.font_color, 'italic': args.italic,
        'bg_color': args.bg_color, 'border_color': args.border_color,
    }
    sp_engine, sp_wl, sp_status = init_spellcheck(do_spell)
    # FAIL LOUD: spell-check is default-on; a missing optional dep must NOT
    # silently no-op. Captioning still
    # proceeds, but we banner it, drop a marker, and exit non-zero.
    spellcheck_degraded = do_spell and sp_status == 'unavailable'
    opts = {
        'dry_run': args.dry_run, 'update_existing': args.update_existing,
        'gap_emu': args.gap_emu, 'height_emu': args.height_emu,
        'caption_smartart': not args.no_smartart,
        'spellcheck': do_spell and sp_status == 'ready',
        'dateqc': do_dateqc, 'qc_dir': qc_dir,
        '_sp': sp_engine, '_wl': sp_wl,
    }

    if args.dry_run:
        print("=== DRY RUN — no .pptx files will be modified ===")
    if args.quick:
        print("=== --quick: captioning ONLY, all QC skipped by request ===")
    if spellcheck_degraded:
        os.makedirs(qc_dir, exist_ok=True)
        with open(os.path.join(qc_dir, 'SPELLCHECK_NOT_RUN.txt'), 'w') as mf:
            mf.write("Spell-check is default-ON but pyspellchecker is NOT importable "
                     "in the Python that ran apply_captions.py.\n"
                     "Captions were applied; SPELL-CHECK DID NOT RUN.\n"
                     "Fix: run via a Python with pyspellchecker installed, or pass "
                     "--no-spellcheck / --quick to acknowledge skipping it.\n")
        print("*" * 72)
        print("*** ERROR: spell-check is ON by default but pyspellchecker is NOT")
        print("*** installed in this Python. Captions WILL be applied, but the")
        print("*** spell-check DID NOT RUN. Marker: qc/SPELLCHECK_NOT_RUN.txt")
        print("*** Re-run with pyspellchecker available, or pass --no-spellcheck")
        print("*** / --quick to explicitly proceed without it. (exit code 3)")
        print("*" * 72)
    elif opts['spellcheck']:
        print(f"=== spell-check ON (default; FLAG-ONLY, whitelist={len(sp_wl)} "
              "terms): captions + slide text scanned; qc/<deck>_spellcheck.csv "
              "emitted; NO .pptx edited, NO auto-correction ===")
    if opts['dateqc']:
        print("=== date/template QC ON (default; FLAG-ONLY): "
              "qc/<deck>_qc.csv emitted ===")

    total_added = total_skipped = total_pics = total_removed = total_errored = 0
    total_spell = total_qc = total_verify = 0
    decks_with_errors = []
    for i, (deck_name, deck_info) in enumerate(manifest.items(), 1):
        if 'error' in deck_info:
            total_errored += 1
            decks_with_errors.append(deck_name)
            if not args.quiet:
                print(f"[{i}/{len(manifest)}] {deck_name:<35} SKIPPED (extract error: {deck_info['error'][:60]})")
            continue
        try:
            result = apply_to_deck(deck_info, captions, captioned_dir, audit_dir, style, opts)
        except Exception as e:
            total_errored += 1
            decks_with_errors.append(deck_name)
            if not args.quiet:
                print(f"[{i}/{len(manifest)}] {deck_name:<35} ERROR: {type(e).__name__}: {str(e)[:60]}")
            continue
        rows = result['rows']
        added = sum(1 for r in rows if r['action'].startswith('added') or r['action'] in ('fallback-above','fallback-bottom') or r['action'].startswith('dry-run-would'))
        skipped = sum(1 for r in rows if r['action'].startswith('skipped'))
        if not args.quiet:
            verb = 'would-add' if args.dry_run else 'added'
            dst_label = os.path.basename(result['dst']) if result['dst'] else '(dry-run, no output)'
            removed_note = f"  removed-prior:{result['n_caption_shapes_removed']}" if result['n_caption_shapes_removed'] else ''
            print(f"[{i}/{len(manifest)}] {deck_name:<35} {added:>3}/{len(rows)} {verb}  {skipped:>3} skip{removed_note}  -> {dst_label}")
        total_added += added; total_skipped += skipped; total_pics += len(rows)
        total_removed += result['n_caption_shapes_removed']
        sr = result.get('spell_rows', [])
        total_spell += len(sr)
        total_verify += sum(1 for r in sr if r.get('verify_name'))
        total_qc += len(result.get('qc_rows', []))

    print(f"\n{'='*72}")
    print(f"TOTAL: {total_added}/{total_pics} captions {'would be added (dry-run)' if args.dry_run else 'added'}, {total_skipped} skipped, {total_errored} deck-level errors")
    if total_removed:
        print(f"Prior captioner shapes removed (idempotency): {total_removed}")
    if opts['spellcheck']:
        print(f"Spell-check flags (review only, no edits made): {total_spell} "
              f"— see qc/<deck>_spellcheck.csv in {qc_dir}")
        if total_verify:
            print(f"  ↳ {total_verify} are likely PROPER NOUNS (verify_name=True): "
                  "web-verify the canonical spelling BEFORE presenting as a fix "
                  "(see SKILL.md “Name verification”).")
    if opts['dateqc']:
        print(f"Date/template QC flags (review only): {total_qc} "
              f"— see qc/<deck>_qc.csv in {qc_dir}")
    if spellcheck_degraded:
        print("SPELL-CHECK DID NOT RUN — pyspellchecker missing "
              "(qc/SPELLCHECK_NOT_RUN.txt). See banner above.")
    if decks_with_errors:
        print(f"Decks with errors: {decks_with_errors}")
    if not args.dry_run:
        print(f"Captioned decks: {captioned_dir}")
    print(f"Audit CSVs:      {audit_dir}")
    if do_spell or do_dateqc:
        print(f"QC CSVs:         {qc_dir}")
    print(f"{'='*72}")
    if spellcheck_degraded:
        sys.exit(3)


if __name__ == '__main__':
    main()
