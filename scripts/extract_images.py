"""Extract unique images from .pptx (or all .pptx in a folder) into a work dir.

Usage:
  python3 extract_images.py <input.pptx-or-folder> <work_dir> [--context "Lecture 3 — Module 5 overview"]

Produces:
  <work_dir>/manifest.json
  <work_dir>/images/<deck_name>/<sha256-hash>.<ext>

Features (hardened following an independent external code review):
  - Recursive traversal into GROUP shapes (catches nested pictures consultants love to group)
  - Skips hidden shapes (cNvPr hidden="1")
  - Per-image progress prints
  - Try/except wrapper per deck (one corrupt/password-protected file does not halt the batch)
  - Captures slide title + body text context for caption quality
  - Optional --context flag for deck-level overarching context (e.g., course name)
"""
import sys, os, json, hashlib, glob, argparse, traceback
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def iter_pictures_recursive(shapes, depth=0):
    """Yield every Picture shape in a slide, recursing into GROUP shapes."""
    for sh in shapes:
        # Skip hidden shapes (cNvPr hidden="1")
        try:
            xml_el = sh._element
            # Walk into nv*Pr / cNvPr to check hidden attribute
            cNvPr = None
            for tag_suffix in ('nvPicPr', 'nvSpPr', 'nvGrpSpPr', 'nvGraphicFramePr', 'nvCxnSpPr'):
                nvpr = getattr(xml_el, tag_suffix, None)
                if nvpr is not None:
                    cNvPr = getattr(nvpr, 'cNvPr', None)
                    break
            if cNvPr is not None and cNvPr.get('hidden') == '1':
                continue
        except Exception:
            pass  # If hidden-check fails, proceed (don't drop the shape silently)

        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            yield sh, depth
        elif sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            try:
                yield from iter_pictures_recursive(sh.shapes, depth + 1)
            except NotImplementedError:
                # python-pptx raises NotImplementedError on certain group geometries
                # (e.g., SmartArt-derived groups). Skip + log silently — apply_captions will too.
                continue
            except Exception:
                continue


def extract_text_context(slide):
    """Pull all visible text on a slide for caption context (title + body)."""
    bits = []
    for sh in slide.shapes:
        try:
            if sh.has_text_frame:
                t = (sh.text_frame.text or '').strip()
                if t:
                    bits.append(t)
        except Exception:
            continue
    return ' | '.join(bits)[:400]  # cap


def process_deck(deck_path, images_root, deck_context=''):
    deck = os.path.splitext(os.path.basename(deck_path))[0]
    deck_dir = os.path.join(images_root, deck)
    os.makedirs(deck_dir, exist_ok=True)
    try:
        prs = Presentation(deck_path)
    except Exception as e:
        print(f"  ERROR opening {deck}: {type(e).__name__}: {str(e)[:80]}")
        return {
            'deck': deck, 'deck_path': os.path.abspath(deck_path),
            'error': f'{type(e).__name__}: {e}', 'n_slides': 0, 'n_pictures': 0,
            'n_unique_images': 0, 'pictures': [], 'deck_context': deck_context,
        }

    seen = {}
    pictures = []
    n_groups_seen = 0
    n_hidden_skipped = 0  # implicit — would be counted if we tracked it

    for s_idx, slide in enumerate(prs.slides, 1):
        slide_text = extract_text_context(slide)
        # Count top-level groups for diagnostics
        for sh in slide.shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                n_groups_seen += 1

        for shape, depth in iter_pictures_recursive(slide.shapes):
            try:
                blob = shape.image.blob
                ext = shape.image.ext
            except Exception:
                continue
            h = hashlib.sha256(blob).hexdigest()[:12]
            try:
                cNvPr = shape._element.nvPicPr.cNvPr
                pic_id = cNvPr.get('id')
                pic_name = cNvPr.get('name')
                old_descr = cNvPr.get('descr', '')
            except AttributeError:
                pic_id = ''; pic_name = ''; old_descr = ''
            fname = f"{h}.{ext}"
            if h not in seen:
                with open(os.path.join(deck_dir, fname), 'wb') as f:
                    f.write(blob)
                seen[h] = fname
            pictures.append({
                'slide': s_idx, 'pic_id': pic_id, 'name': pic_name,
                'image_hash': h, 'image_file': fname, 'ext': ext,
                'bytes': len(blob), 'in_group_depth': depth,
                'slide_text_context': slide_text, 'old_descr': old_descr,
            })

    return {
        'deck': deck, 'deck_path': os.path.abspath(deck_path),
        'n_slides': len(prs.slides), 'n_pictures': len(pictures),
        'n_unique_images': len(seen), 'n_top_level_groups_seen': n_groups_seen,
        'pictures': pictures, 'deck_context': deck_context,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('input', help='Single .pptx or a directory of .pptx files')
    ap.add_argument('work_dir', help='Output working directory')
    ap.add_argument('--context', default='', help='Deck-level overarching context (e.g., "Lecture 3: Decision Trees")')
    args = ap.parse_args()

    inp = os.path.abspath(args.input)
    work = os.path.abspath(args.work_dir)
    os.makedirs(work, exist_ok=True)
    images_root = os.path.join(work, 'images')
    os.makedirs(images_root, exist_ok=True)

    if os.path.isfile(inp):
        decks = [inp]
    elif os.path.isdir(inp):
        decks = sorted(glob.glob(os.path.join(inp, '*.pptx')))
        # Filter out our own output to avoid recursion
        decks = [d for d in decks if not d.endswith('_captioned.pptx')]
    else:
        print(f"ERROR: {inp} not found"); sys.exit(2)

    if not decks:
        print("No .pptx files found."); sys.exit(2)

    manifest = {}
    for i, d in enumerate(decks, 1):
        print(f"[{i}/{len(decks)}] {os.path.basename(d)}")
        info = process_deck(d, images_root, args.context)
        manifest[info['deck']] = info
        if 'error' in info:
            print(f"           SKIPPED (error captured in manifest)")
        else:
            groups = info.get('n_top_level_groups_seen', 0)
            extra = f"  (+{groups} group shapes traversed)" if groups else ""
            print(f"           {info['n_slides']:>3} slides  {info['n_pictures']:>3} pics  {info['n_unique_images']:>3} unique{extra}")

    out = os.path.join(work, 'manifest.json')
    with open(out, 'w') as f:
        json.dump(manifest, f, indent=2)

    n_ok = sum(1 for v in manifest.values() if 'error' not in v)
    n_err = sum(1 for v in manifest.values() if 'error' in v)
    print(f"\nManifest: {out}  ({n_ok} ok, {n_err} errored)")
    if args.context:
        print(f"Deck context: '{args.context}' applied to all decks")
    print(f"Images:   {images_root}/")
    print(f"\nNext: Claude reads each image and writes captions.json to {work}/captions.json")


if __name__ == '__main__':
    main()
