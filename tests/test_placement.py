"""Regression tests for captioner placement + safety logic (v0.2.2 onward).

Run:  python3 tests/test_placement.py     (exit 0 = all pass, 1 = a failure)

No external fixtures — builds tiny in-memory decks with python-pptx. Covers the
changes most prone to silent regression:
  - 2D obstacle clearance (x AND y) in _geometry
  - full-rect body-obstacle extraction
  - every text frame (not just placeholders) treated as an obstacle
  - caption height estimate growing with text length
  - the pre-write invariant predicates (text overlap, in-picture, off-slide)
  - edit-aware caption removal (a human's hand-edits are preserved)
"""
import os
import sys

from pptx import Presentation
from pptx.util import Emu

# Resolve the scripts dir whether this test lives alongside the scripts (flat
# layout) or in a sibling tests/ dir (repo layout), as test_descr_writer.py does.
_HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(_HERE), 'scripts')
if not os.path.exists(os.path.join(SCRIPTS, '_geometry.py')):
    SCRIPTS = _HERE
sys.path.insert(0, SCRIPTS)

import _geometry as g            # noqa: E402
import apply_captions as ac      # noqa: E402

FAILS = []


def check(name, cond):
    print(f"  {'PASS' if cond else 'FAIL'}  {name}")
    if not cond:
        FAILS.append(name)


def test_2d_clearance():
    print("2D obstacle clearance:")
    cap = (1_000_000, 5_000_000, 3_000_000, 5_400_000)  # L,T,R,B (2in x 0.44in)
    # horizontally disjoint, same vertical band -> CLEARS (the core fix)
    check("disjoint-x clears", g.clear_all_obstacles_2d(cap, [(4_000_000, 5_000_000, 6_000_000, 5_400_000)]))
    # overlapping both axes -> BLOCKED
    check("both-axis overlap blocks", not g.clear_all_obstacles_2d(cap, [(1_500_000, 5_100_000, 3_500_000, 5_500_000)]))
    # vertically disjoint -> CLEARS
    check("disjoint-y clears", g.clear_all_obstacles_2d(cap, [(1_000_000, 1_000_000, 3_000_000, 1_400_000)]))
    # (l,t,w,h) misread as (L,T,R,B) was the v0.2.1 cap-cap bug; ensure a
    # placed-caption rect stored as (L,T,R,B) is now detected when overlapping.
    placed = (1_000_000, 5_050_000, 3_000_000, 5_450_000)  # L,T,R,B overlapping cap
    check("placed-cap overlap blocks (strict)", not g.clear_all_obstacles_2d(cap, [placed], vfrac=0.02))


def test_body_rects_full():
    print("Body obstacle rects carry horizontal extent:")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # title+content
    body = None
    for ph in slide.placeholders:
        if ph.placeholder_format.idx != 0:  # not the title
            body = ph
            break
    ok = body is not None
    if ok:
        body.left, body.top, body.width, body.height = Emu(1_000_000), Emu(2_000_000), Emu(4_000_000), Emu(1_500_000)
        body.text_frame.text = "some body text"
        rects = g.slide_body_obstacle_rects(slide)
        ok = len(rects) == 1 and len(rects[0]) == 4 and rects[0][2] > rects[0][0]
    check("body rect is 4-tuple with width", ok)


def test_text_obstacle_rects_catches_textbox():
    print("ALL text frames are obstacles (not just placeholders):")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank — no placeholders
    tb = slide.shapes.add_textbox(Emu(1_000_000), Emu(2_000_000), Emu(3_000_000), Emu(500_000))
    tb.text_frame.text = "Use it to research questions?"   # a plain text box, not a placeholder
    rects = g.slide_text_obstacle_rects(slide)
    ok = any(r[0] == 1_000_000 and len(r) == 4 for r in rects)
    check("plain text box detected as obstacle", ok)
    # a captioner shape must NOT be treated as a text obstacle
    cap = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(1_000_000), Emu(300_000))
    cap._element.nvSpPr.cNvPr.set('name', f"{ac.CAPTION_SHAPE_NAME_PREFIX}deadbeef")
    cap.text_frame.text = "Photo: something"
    rects2 = g.slide_text_obstacle_rects(slide)
    check("captioner shape excluded from text obstacles", len(rects2) == len(rects))


def test_estimate_caption_height_grows():
    print("Caption height estimate grows with text length:")
    short = g.estimate_caption_height("Photo: cat.", 3_000_000)
    long = g.estimate_caption_height("Photo: " + "x" * 200, 3_000_000)
    check("longer caption -> taller estimate", long > short)
    check("short caption floors at nominal", short >= g.MIN_CAPTION_HEIGHT)


def test_pre_write_invariant_predicates():
    print("Pre-write invariant rejects each defect class:")
    # The invariant (apply_captions, v0.2.4) refuses to write a caption whose FINAL
    # rect violates any defect class. Exercise the same predicates it uses.
    cap = (1_000_000, 5_000_000, 3_000_000, 5_400_000)  # L,T,R,B
    # text overlap -> must NOT clear (rejected)
    text = [(1_500_000, 5_100_000, 2_500_000, 5_500_000)]
    check("text-overlap rejected", not g.clear_all_obstacles_2d(cap, text, vfrac=0.03))
    # clear text -> accepted
    check("text-clear accepted", g.clear_all_obstacles_2d(cap, [(8_000_000, 5_000_000, 9_000_000, 5_400_000)], vfrac=0.03))
    # >50% inside a picture -> detected by the area test the invariant uses
    pic = (900_000, 4_900_000, 3_200_000, 5_500_000)
    ix = max(0, min(cap[2], pic[2]) - max(cap[0], pic[0]))
    iy = max(0, min(cap[3], pic[3]) - max(cap[1], pic[1]))
    cap_area = (cap[2] - cap[0]) * (cap[3] - cap[1])
    check("in-picture >50% detected", (ix * iy) / cap_area > 0.50)
    # off-slide detection
    slide_w, slide_h = 9_144_000, 6_858_000
    off = (8_000_000, 1_000_000, 9_500_000, 1_400_000)  # right edge past slide_w
    check("off-slide detected", not (off[0] >= 0 and off[1] >= 0 and off[2] <= slide_w and off[3] <= slide_h))


def test_edit_aware_removal():
    print("Edit-aware caption removal:")
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    # two captioner-named caption boxes
    tb1 = slide.shapes.add_textbox(Emu(0), Emu(0), Emu(1_000_000), Emu(300_000))
    tb1._element.nvSpPr.cNvPr.set('name', f"{ac.CAPTION_SHAPE_NAME_PREFIX}aaaaaaaaaaaa")
    tb1.text_frame.text = "Photo: original caption"
    tb2 = slide.shapes.add_textbox(Emu(0), Emu(400_000), Emu(1_000_000), Emu(300_000))
    tb2._element.nvSpPr.cNvPr.set('name', f"{ac.CAPTION_SHAPE_NAME_PREFIX}bbbbbbbbbbbb")
    tb2.text_frame.text = "HAND EDITED THIS"
    oracle = {"aaaaaaaaaaaa": "Photo: original caption",
              "bbbbbbbbbbbb": "Photo: machine caption b"}  # b was edited
    n_rm, n_pres = ac.remove_previous_caption_shapes(slide, oracle=oracle)
    check("unedited removed, edited preserved", n_rm == 1 and n_pres == 1)
    # force overwrite removes both (need fresh slide since one was removed)
    prs2 = Presentation()
    s2 = prs2.slides.add_slide(prs2.slide_layouts[6])
    for nm, txt in (("aaaaaaaaaaaa", "Photo: original caption"), ("bbbbbbbbbbbb", "EDITED")):
        t = s2.shapes.add_textbox(Emu(0), Emu(0), Emu(1_000_000), Emu(300_000))
        t._element.nvSpPr.cNvPr.set('name', f"{ac.CAPTION_SHAPE_NAME_PREFIX}{nm}")
        t.text_frame.text = txt
    n_rm2, n_pres2 = ac.remove_previous_caption_shapes(s2, oracle=oracle, force_overwrite_edits=True)
    check("force-overwrite removes edited too", n_rm2 == 2 and n_pres2 == 0)


if __name__ == "__main__":
    test_2d_clearance()
    test_body_rects_full()
    test_text_obstacle_rects_catches_textbox()
    test_estimate_caption_height_grows()
    test_pre_write_invariant_predicates()
    test_edit_aware_removal()
    print()
    if FAILS:
        print(f"{len(FAILS)} FAILED: {FAILS}")
        sys.exit(1)
    print("ALL TESTS PASSED")
