"""Verify captioned decks.

v0.2.1: in addition to text-presence parity (the original v0.1.x check),
verify.py now performs two coordinate sanity checks:
  - Pattern B: no caption bottom is within FOOTER_CLEARANCE_EMU of the footer.
  - Pattern C: no caption box is geometrically inside its picture.
Audit rows with action='overlay-fullbleed' or 'flagged-no-slot' are treated
as known-missing (intentionally skipped placement), not text-mismatch errors.

Usage:
  python3 verify.py <work_dir>

Reads:
  <work_dir>/captions.json
  <work_dir>/captioned_decks/<deck>_captioned.pptx
  <work_dir>/audit/<deck>_audit.csv  (optional — used to honor known-skip rows)
"""
import sys, os, json, hashlib, glob, csv
from pptx import Presentation
from _oxml_pics import iter_slide_pics, resolve_blob, guess_ext
from _geometry import (
    slide_footer_top, FOOTER_CLEARANCE_EMU, rect_intersect_area,
    CAPTION_NAME_PREFIXES,
)


def load_known_skips(audit_dir, deck_stem):
    """Return a set of image_hash values that the apply pass intentionally
    did NOT caption (overlay-fullbleed or flagged-no-slot). Caller treats
    these as known-missing during text-match verification."""
    skips = set()
    p = os.path.join(audit_dir, f"{deck_stem}_audit.csv")
    if not os.path.exists(p):
        return skips
    try:
        with open(p, encoding="utf-8", newline="") as f:
            for row in csv.DictReader(f):
                act = row.get("action", "")
                if "overlay-fullbleed" in act or "flagged-no-slot" in act:
                    h = row.get("image_hash", "")
                    if h:
                        skips.add(h)
    except Exception:
        pass
    return skips


def caption_shapes(slide):
    """Yield (shape_name, (left, top, width, height)) for every caption-named shape."""
    for sh in slide.shapes:
        nm = sh.name or ""
        if not any(nm.startswith(p) for p in CAPTION_NAME_PREFIXES):
            continue
        if None in (sh.left, sh.top, sh.width, sh.height):
            continue
        yield nm, (int(sh.left), int(sh.top), int(sh.width), int(sh.height))


def main():
    if len(sys.argv) != 2:
        print(__doc__); sys.exit(2)
    work = os.path.abspath(sys.argv[1])
    captions_path = os.path.join(work, "captions.json")
    if not os.path.exists(captions_path):
        print(f"ERROR: {captions_path} not found"); sys.exit(2)
    with open(captions_path) as f:
        captions = json.load(f)
    captioned_dir = os.path.join(work, "captioned_decks")
    audit_dir = os.path.join(work, "audit")

    expected_values = set(c for c in captions.values()
                          if c and c.strip().lower() not in ("[decorative]", "decorative"))

    print(f"{'Deck':<38}{'Slides':>7}{'Pics':>6}{'Cap':>5}{'Skip':>5}{'B':>4}{'C':>4}{'Match':>10}")
    print("=" * 79)

    grand = {"decks": 0, "B_fail": 0, "C_fail": 0, "text_miss": 0, "skips": 0}
    for path in sorted(glob.glob(os.path.join(captioned_dir, "*_captioned.pptx"))):
        deck = os.path.basename(path).replace("_captioned.pptx", "")
        known_skip = load_known_skips(audit_dir, deck)
        prs = Presentation(path)
        H = prs.slide_height
        total_pics = total_match = 0
        b_fail = c_fail = skip_known = 0
        for slide in prs.slides:
            flim = slide_footer_top(slide, H, exclude_caption_shapes=True)
            textboxes = []
            for sh in slide.shapes:
                if sh.has_text_frame:
                    t = (sh.text_frame.text or "").strip()
                    if t:
                        textboxes.append(t)
            # Picture geometry list for Pattern C check on this slide
            slide_pic_rects = []
            for pic in iter_slide_pics(slide):
                if pic["rid"] is None:
                    continue
                total_pics += 1
                try:
                    blob = resolve_blob(slide, pic["rid"])
                    ext = guess_ext(slide.part.related_part(pic["rid"]))
                    h = hashlib.sha256(blob).hexdigest()[:12]
                except Exception:
                    continue
                # Track pic rect for Pattern C check
                ox, oy, cx, cy = (pic.get("off_x"), pic.get("off_y"),
                                  pic.get("ext_cx"), pic.get("ext_cy"))
                if None not in (ox, oy, cx, cy):
                    slide_pic_rects.append((int(ox), int(oy), int(cx), int(cy)))
                key = f"{deck}/{h}.{ext}"
                expected = captions.get(key)
                if not expected or expected.strip().lower() in ("[decorative]", "decorative"):
                    continue
                if h in known_skip:
                    skip_known += 1
                    continue  # intentionally not placed
                if expected in textboxes:
                    total_match += 1
            # Pattern B/C coordinate checks on all caption shapes
            for _, (cl, ct, cw, ch) in caption_shapes(slide):
                # Pattern B: caption bottom too close to footer
                if (ct + ch) > flim - FOOTER_CLEARANCE_EMU:
                    b_fail += 1
                # Pattern C: caption box mostly inside any picture
                for pr in slide_pic_rects:
                    ov = rect_intersect_area((cl, ct, cw, ch), pr)
                    if cw * ch > 0 and ov / (cw * ch) > 0.50:
                        c_fail += 1
                        break

        # Effective text-match: matches + known-skipped should equal pics that
        # have non-decorative captions in captions.json.
        flag = "✓" if total_match + skip_known == total_pics else \
               f"{total_match}/{total_pics}"
        print(f"{deck:<38}{len(prs.slides):>7}{total_pics:>6}{total_match:>5}"
              f"{skip_known:>5}{b_fail:>4}{c_fail:>4}{flag:>10}")
        grand["decks"] += 1
        grand["B_fail"] += b_fail
        grand["C_fail"] += c_fail
        grand["text_miss"] += max(0, total_pics - total_match - skip_known)
        grand["skips"] += skip_known

    print("=" * 79)
    print(f"TOTAL: {grand['decks']} decks | text-miss={grand['text_miss']} | "
          f"known-skips={grand['skips']} | Pattern-B fails={grand['B_fail']} | "
          f"Pattern-C fails={grand['C_fail']}")
    # Exit non-zero on any coord-quality failure (placement regression signal).
    sys.exit(0 if grand["B_fail"] == 0 and grand["C_fail"] == 0 else 4)


if __name__ == "__main__":
    main()
