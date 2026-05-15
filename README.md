# Captioner

*A brief, subject-identifying caption under every picture in a PowerPoint deck — with an auditable trail.*

Captioner inserts a short (5–15 word) italic caption directly below every picture in a `.pptx` deck. The caption identifies *what the picture is* — so sighted readers know at a glance, and so a reviewer has a written record of every visual on every slide. Each run is non-destructive (the source `.pptx` is preserved untouched), idempotent (re-runs replace prior captions cleanly), and produces a per-deck audit CSV recording the exact action taken on every shape.

## Requirements

Captioner is a **Claude Code skill**. The caption-generation step uses Claude Code's built-in image reading (vision) — there is no separate vision API to configure, and captioner is **not** a standalone Python package. You need:

- [Claude Code](https://www.anthropic.com/claude-code) (provides the vision capability the workflow depends on)
- Python 3.9+
- `python-pptx` (and `lxml`, `Pillow`) — see `requirements.txt`

## Install

```bash
git clone https://github.com/jbenhart44/captioner.git
cd captioner
git checkout v0.1.1
pip install -r requirements.txt
bash install.sh
```

Then restart Claude Code and run:

```
/captioner <path-to-.pptx-or-folder>
```

> Prefer a versioned archive? Download the [v0.1.1 release](https://github.com/jbenhart44/captioner/releases/tag/v0.1.1) tarball instead of cloning.

## Quick start

```bash
# Smoke test against the bundled synthetic deck
python3 tests/test_smoke.py
```

In Claude Code, point the skill at a deck or a folder of decks:

```
/captioner ~/lectures/module5.pptx
/captioner ~/lectures/            # every .pptx in the folder
```

You get back: `<deck>_captioned.pptx` files plus a per-deck `<deck>_audit.csv`.

## Usage

The full workflow (extract → read → caption → dry-run audit → apply → verify) and every CLI flag is documented in [`SKILL.md`](SKILL.md). Run a dry run first to review the audit CSV before any deck is modified.

<details>
<summary>CLI flag reference</summary>

### `extract_images.py`
| Flag | Default | Effect |
|---|---|---|
| `--context "<text>"` | empty | Deck-level context recorded in the manifest to bias caption generation |

### `apply_captions.py`
| Flag | Default | Effect |
|---|---|---|
| `--dry-run` | off | Emits the audit CSV; does NOT modify any `.pptx` |
| `--font-name` | Calibri | Caption font family |
| `--font-size` | 10 | Caption font size (pt) |
| `--font-color` | 333333 | Caption hex color (no `#`) |
| `--bg-color` | FFFFFF | Caption text-box fill (white card; `""` to disable) |
| `--border-color` | CCCCCC | Caption box border (`""` to disable) |
| `--italic` | true | Italic on/off |
| `--gap-emu` | 50000 | EMU between picture and caption |
| `--height-emu` | 400000 | Caption box height (EMU) |
| `--update-existing` | off | Strip prior captioner shapes before re-applying (idempotent re-run) |
| `--no-smartart` | off | Disable SmartArt icon captioning |
| `--quiet` | off | Suppress per-deck progress |

</details>

## Capabilities

1. **Subject-identifying photo captions** — brief, category-prefixed (`Photo: …`), no descriptive prose, no "image of…" filler.
2. **Category-aware prefixes** — distinguishes `Photo:`, `Illustration:`, `Diagram:`, `Chart:`, `Screenshot:`, `Logo:`, `Map:`, `Icon:` from visual context.
3. **Vision reads in-image text** — reproduces typography rendered as part of an image (a warning sign, a code screenshot).
4. **White-card background fill** — captions stay legible on any slide background, including dark section dividers.
5. **Three positioning fallbacks** — below the picture by default; above, then slide-bottom; the chosen fallback is logged.
6. **SmartArt per-icon captioning** — each embedded PowerPoint Icon gets its own caption from its SVG `id` metadata (no vision call).
7. **Decorative-image triage** — repeating logos and bullet markers are flagged `[decorative]` and skipped.
8. **Text-only SmartArt skip** — diagrams whose only content is text labels are intentionally not captioned.
9. **Deck-level context** — a free-text context string biases caption vocabulary toward the subject.
10. **Audit CSV trail** — every action (`added`, `fallback-above`, `fallback-bottom`, `skipped-decorative`, `added-smartart-icon`, `skipped-smartart-text-only`, …) is logged.
11. **Idempotent re-runs** — captioner shapes carry a content-hash name; `--update-existing` regenerates cleanly with no duplicates.

## What captioner does NOT do

- Write the OOXML `descr` alt-text field (the WCAG-canonical attribute consumed by screen readers). Captioner is a pedagogical-clarity layer; pair it with a full alt-text remediation workflow for WCAG 2.1 AA conformance.
- Read off chart values — captions identify *what the picture is*, not what its data says.
- Modify embedded charts, the slide master, or any non-Picture shape.
- Process `.ppt` legacy format (python-pptx supports `.pptx` only).

## Track record

1,132 captions across 32 PowerPoint decks in three graduate engineering courses (Summer 2026). Source decks preserved unmodified; every run produces a reproducible per-deck audit CSV.

## License

MIT — see [`LICENSE`](LICENSE). Copyright (c) 2026 Jake Benhart.

## Citation

```bibtex
@software{benhart_captioner_2026,
  author  = {Benhart, Jake},
  title   = {Captioner: pedagogical-clarity captions for PowerPoint decks},
  year    = {2026},
  version = {0.1.1},
  url     = {https://github.com/jbenhart44/captioner}
}
```
