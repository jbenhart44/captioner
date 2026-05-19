"""Shared raw-OOXML <p:pic> enumeration for the captioner skill.

Replaces the python-pptx `shape_type == MSO_SHAPE_TYPE.PICTURE` walk, which
silently TYPES placeholder-hosted pictures (`<p:pic>` carrying `<p:ph type="pic"/>`)
and OLE-fallback equation bitmaps as PLACEHOLDER (not PICTURE) and skips them.

Contract (adversarially verified):
  * STRICT slide-spTree scope: only `ppt/slides/slideN.xml` shape trees are walked.
    Layout/master pics and `ppt/diagrams/` SmartArt fallback bitmaps are OUT by
    construction (they never appear in a slide's own spTree).
  * `<mc:AlternateContent>`: count the Fallback pic ONCE. Choice subtrees are
    removed before the descendant search, so a (future) Choice branch that also
    carries a `<p:pic>` cannot double-count. Verified: this corpus has 204
    AlternateContent blocks, 0 with a Choice pic, 204 with exactly one Fallback
    pic; oracle counts reproduced exactly.
  * Blob via `<a:blip r:embed>` -> `slide.part.related_part(rId).blob`. This is
    byte-identical to the old `pic.image.blob` (python-pptx resolves the same
    related part internally) -> sha256(blob)[:12] hashes are unchanged, so the
    reuse index / prior captions.json keys still match.
  * `<a:blip>` with NO `r:embed` (e.g. an `r:link` external picture) -> SKIP +
    structured log entry (`linked-no-embed`). None exist in this corpus.
  * `in_group_depth` = number of enclosing `<p:grpSp>` ancestors (NOT incremented
    by `<mc:Fallback>` / `<p:oleObj>` wrappers, so OLE-fallback pics read depth 0).

Geometry: `<a:off>`/`<a:ext>` EMU read from the pic's own `<a:xfrm>`. For a
group-nested pic this is the group-child coordinate (same value python-pptx
`.left/.top` returned for grouped shapes), so caption placement is byte-for-byte
unchanged vs the prior skill. (True slide-absolute de-nesting of grouped pics is
a pre-existing skill limitation, NOT introduced here.)
"""
from lxml import etree

NS = {
    'p':  'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a':  'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}
_R_EMBED = '{%s}embed' % NS['r']
_R_LINK = '{%s}link' % NS['r']


def _ln(el):
    t = el.tag
    return t.split('}', 1)[1] if isinstance(t, str) and '}' in t else t


def _strip_choice(root):
    """Remove every <mc:Choice> subtree so only <mc:Fallback> content remains.

    PowerPoint stores an OLE object as <mc:Choice Requires="v"> (a v:* VML/OLE
    branch with NO <p:pic>) plus an <mc:Fallback> that carries the rendered
    <p:pic> bitmap. Counting the raw tree is correct here, but stripping Choice
    makes the walk robust to a future deck whose Choice ALSO embeds a <p:pic>.
    """
    for ac in root.findall('.//mc:AlternateContent', NS):
        for ch in list(ac.findall('mc:Choice', NS)):
            ac.remove(ch)


def _xfrm_geom(pic_el):
    """Return (off_x, off_y, ext_cx, ext_cy) in EMU from the pic's own xfrm.

    Falls back to (None, None, None, None) component-wise if absent (caller then
    uses python-pptx geometry / the existing fallback-placement logic)."""
    off_x = off_y = ext_cx = ext_cy = None
    xfrm = pic_el.find('.//a:xfrm', NS)
    if xfrm is not None:
        off = xfrm.find('a:off', NS)
        ext = xfrm.find('a:ext', NS)
        if off is not None:
            try:
                off_x = int(off.get('x')); off_y = int(off.get('y'))
            except (TypeError, ValueError):
                off_x = off_y = None
        if ext is not None:
            try:
                ext_cx = int(ext.get('cx')); ext_cy = int(ext.get('cy'))
            except (TypeError, ValueError):
                ext_cx = ext_cy = None
    return off_x, off_y, ext_cx, ext_cy


def _walk(el, depth, out):
    """Depth-first walk. depth = count of enclosing <p:grpSp> ancestors.

    Wrapper containers (<mc:AlternateContent>, <mc:Fallback>, <p:oleObj>,
    <p:spPr>, ...) are descended WITHOUT incrementing depth; only <p:grpSp>
    increments it. A <p:pic> is yielded once when first encountered.
    """
    for child in el:
        ln = _ln(child)
        if ln == 'pic':
            out.append((child, depth))
            # do not recurse into a pic (a pic has no nested pic)
            continue
        if ln == 'grpSp':
            _walk(child, depth + 1, out)
            continue
        # Any other container: recurse, same depth.
        _walk(child, depth, out)


def iter_slide_pics(slide):
    """Yield dicts for every <p:pic> in this slide's spTree (strict).

    Each dict:
      el         lxml element (live python-pptx element copy via fromstring)
      depth      int, enclosing <p:grpSp> count (in_group_depth)
      rid        r:embed value, or None
      linked     True if the blip has r:link but no r:embed (external; skip)
      off_x/off_y/ext_cx/ext_cy   EMU geometry or None
      pic_id/pic_name/old_descr   from <p:cNvPr>
    The caller resolves blob via slide.part.related_part(rid).
    """
    spTree = slide.shapes._spTree
    # Re-parse a detached copy so Choice-stripping never mutates the live deck.
    root = etree.fromstring(etree.tostring(spTree))
    _strip_choice(root)
    raw = []
    _walk(root, 0, raw)
    out = []
    for pic_el, depth in raw:
        cNvPr0 = pic_el.find('.//p:cNvPr', NS)
        # Preserve the prior skill's hidden-shape skip (cNvPr hidden="1").
        # (No hidden pics exist in the current corpus; this keeps parity for
        # future decks and matches the old extract_images.py behavior.)
        if cNvPr0 is not None and cNvPr0.get('hidden') == '1':
            continue
        blip = pic_el.find('.//a:blip', NS)
        rid = blip.get(_R_EMBED) if blip is not None else None
        link = blip.get(_R_LINK) if blip is not None else None
        ox, oy, cx, cy = _xfrm_geom(pic_el)
        cNvPr = pic_el.find('.//p:cNvPr', NS)
        if cNvPr is not None:
            pic_id = cNvPr.get('id', '') or ''
            pic_name = cNvPr.get('name', '') or ''
            old_descr = cNvPr.get('descr', '') or ''
        else:
            pic_id = pic_name = old_descr = ''
        # Picture content-placeholder: <p:nvPicPr><p:nvPr><p:ph .../>.
        # Its geometry is INHERITED from the slide layout/master (the <p:pic>
        # often has no <a:xfrm>), so the caller must resolve it via the
        # python-pptx placeholder by this idx — else caption placement is junk.
        ph = pic_el.find('.//p:nvPicPr/p:nvPr/p:ph', NS)
        ph_idx = ph_type = None
        if ph is not None:
            ph_type = ph.get('type')
            iv = ph.get('idx')
            if iv is not None and iv.isdigit():
                ph_idx = int(iv)
        out.append({
            'el': pic_el, 'depth': depth, 'rid': rid,
            'linked': (rid is None and link is not None),
            'off_x': ox, 'off_y': oy, 'ext_cx': cx, 'ext_cy': cy,
            'pic_id': pic_id, 'pic_name': pic_name, 'old_descr': old_descr,
            'is_placeholder': ph is not None, 'ph_idx': ph_idx, 'ph_type': ph_type,
        })
    return out


def resolve_blob(slide, rid):
    """blob bytes for r:embed rid (byte-identical to old pic.image.blob)."""
    part = slide.part.related_part(rid)
    return part.blob


def guess_ext(part):
    """File extension for a related image part (mirrors python-pptx Image.ext)."""
    pn = str(getattr(part, 'partname', '') or '')
    if '.' in pn:
        return pn.rsplit('.', 1)[1].lower()
    ct = (getattr(part, 'content_type', '') or '').lower()
    return {
        'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif',
        'image/bmp': 'bmp', 'image/tiff': 'tiff', 'image/x-emf': 'emf',
        'image/x-wmf': 'wmf', 'image/svg+xml': 'svg', 'image/webp': 'webp',
    }.get(ct, 'img')
