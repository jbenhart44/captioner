"""Verify captioned decks: confirm caption text boxes exist with expected text.

Usage:
  python3 verify.py <work_dir>
"""
import sys, os, json, hashlib, glob
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def main():
    if len(sys.argv) != 2:
        print(__doc__); sys.exit(2)
    work = os.path.abspath(sys.argv[1])
    with open(os.path.join(work, 'captions.json')) as f:
        captions = json.load(f)
    captioned_dir = os.path.join(work, 'captioned_decks')

    print(f"{'Deck':<35}{'Slides':>8}{'Pics':>6}{'Captions':>10}{'Match':>10}")
    print('=' * 70)
    expected_values = set(c for c in captions.values() if c and c.strip().lower() not in ('[decorative]', 'decorative'))

    for path in sorted(glob.glob(os.path.join(captioned_dir, '*_captioned.pptx'))):
        deck = os.path.basename(path).replace('_captioned.pptx', '')
        prs = Presentation(path)
        total_pics = total_match = total_captions_seen = 0
        for slide in prs.slides:
            textboxes = []
            for sh in slide.shapes:
                if sh.has_text_frame:
                    t = (sh.text_frame.text or '').strip()
                    if t:
                        textboxes.append(t)
            for sh in slide.shapes:
                if sh.shape_type != MSO_SHAPE_TYPE.PICTURE:
                    continue
                total_pics += 1
                try:
                    blob = sh.image.blob; ext = sh.image.ext
                    h = hashlib.sha256(blob).hexdigest()[:12]
                    key = f"{deck}/{h}.{ext}"
                    expected = captions.get(key)
                except Exception:
                    continue
                if expected and expected in textboxes:
                    total_match += 1
            for t in textboxes:
                if t in expected_values:
                    total_captions_seen += 1
        flag = '✓' if total_match == total_pics else f'{total_match}/{total_pics}'
        print(f"{deck:<35}{len(prs.slides):>8}{total_pics:>6}{total_captions_seen:>10}{flag:>10}")


if __name__ == '__main__':
    main()
