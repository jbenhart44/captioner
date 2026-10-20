# Captioner

[![tests](https://github.com/jbenhart44/captioner/actions/workflows/tests.yml/badge.svg)](https://github.com/jbenhart44/captioner/actions/workflows/tests.yml)

*A visible caption under every picture in a PowerPoint deck, plus optional WCAG alt-text (the OOXML `descr` field) for images that lack one — with an auditable trail.*

Captioner inserts a short (5–15 word) italic caption directly below every picture in a `.pptx` deck. The caption identifies *what the picture is* — so sighted readers know at a glance, and so a reviewer has a written record of every visual on every slide.

With `--write-descr` (opt-in), captioner also authors the OOXML `descr` attribute — the alt-text a screen reader announces — for pictures that have none, while *preserving* any alt-text a deck already carries and marking decorative images as decorative. Writing that attribute fills a required field; it does **not** by itself certify WCAG 2.1 conformance, and the authored wording is **not** independently verified — review it before you rely on it.

Each run is non-destructive (the source `.pptx` is preserved untouched), idempotent (re-runs replace prior captions cleanly), and produces a per-deck audit CSV recording the exact action taken on every shape.

## Works across every discipline

Captioner was first built for engineering course decks, but it has no built-in field taxonomy — it
adapts to whatever subject you supply through a one-line `--context` prompt at the start of a run,
so the captions and alt-text it writes use the right vocabulary for **any** academic discipline
(e.g. `--context "Biology 101 — cell mitosis"`, `--context "Art History — Baroque painting"`,
`--context "Constitutional Law — the commerce clause"`). Those examples illustrate the *kind* of
context you can pass; they are not validated test cases or a claim the tool has been run across
those specific fields.

## What captioner reads and writes

Captioner **reads** the `.pptx` file(s) you point it at and the images embedded inside them (via
Claude Code's built-in vision: captioner calls no API of its own and needs no API key, but Claude Code sends
each image to Anthropic's API, or to the provider you run Claude Code through, so the model can read it). Your **original `.pptx` is never
modified**. Each run **writes** into a working directory you choose:

- a new captioned copy at `<work_dir>/captioned_decks/<deck>_captioned.pptx`
- a per-deck audit trail at `<work_dir>/audit/<deck>_audit.csv` (plus `<deck>_descr_audit.csv` when you pass `--write-descr`)
- with the default QC pass, flag-only quality-check files at `<work_dir>/qc/<deck>_spellcheck.csv` and `<work_dir>/qc/<deck>_qc.csv`

The installer itself touches exactly one path on your system, `~/.claude/skills/captioner` (or
`$CLAUDE_SKILLS_DIR/captioner`), and announces what it is about to do there before it acts (see
[Install](#install)).

## Requirements

Captioner is a **Claude Code skill**. The caption-generation step uses Claude Code's built-in image reading (vision) — there is no separate vision API to configure, and captioner is **not** a standalone Python package. You need:

- [Claude Code](https://www.anthropic.com/claude-code) (provides the vision capability the workflow depends on)
- Python 3.9+
- `python-pptx`, `lxml`, `Pillow`, and `pyspellchecker` — all in `requirements.txt` (the installer checks each one and tells you if any is missing)

## Install (one command)

Open your Terminal and run this — it downloads the v0.5.0 release, installs the dependencies, and
sets captioner up:

```bash
curl -L https://github.com/jbenhart44/captioner/releases/download/v0.5.0/captioner-0.5.0.zip -o captioner.zip \
  && unzip -o captioner.zip && cd captioner-0.5.0 \
  && pip install -r requirements.txt && bash install.sh
```

**Where it lands:** the installer symlinks captioner into your Claude Code skills directory at
`~/.claude/skills/captioner` (or `$CLAUDE_SKILLS_DIR/captioner`) — that one path is all it touches,
and it prints exactly what it does before it acts. The download itself unzips into a
`captioner-0.5.0/` folder in whatever directory you ran the command from; you can keep or delete
that folder afterward (the working install lives in the skills directory).

**Then restart Claude Code and use it:**

```
/captioner <path-to-your-deck.pptx-or-a-folder-of-decks>
```

Add `--write-descr` to also write screen-reader alt-text, or `--descr-only` for alt-text with no
visible captions. Captioner reads your `.pptx` and writes a **new** captioned copy plus an audit CSV
into a working folder — **your original file is never changed.**

> **Installing with Claude Code itself?** Just tell Claude: *"install captioner from
> https://github.com/jbenhart44/captioner — run the one-command install from the README."*
> Claude will run the `curl … && … && bash install.sh` line above for you.

> Prefer git? `git clone https://github.com/jbenhart44/captioner.git && cd captioner && git checkout v0.5.0 && pip install -r requirements.txt && bash install.sh` works too.

### Updating to a new version

```bash
cd captioner
git fetch --tags
git checkout <new-version-tag>      # e.g. v0.5.0 — replace with the version you want
pip install -r requirements.txt     # re-run: cheap, and a version bump can add a dependency
bash install.sh                     # safe to re-run — a same-target relink just prints "nothing to do"
```

Then restart Claude Code. Stay on a tagged release — do not run `git checkout main` or
`git pull`, which can pull unreleased work instead of a stable version.

## Quick start

```bash
python3 tests/test_smoke.py          # extract -> apply --dry-run on the bundled synthetic deck
python3 tests/test_descr_writer.py   # alt-text engine: preserve-by-default, decorative marking, coverage gate
python3 tests/test_placement.py      # placement and safety: 2D clearance, footer, edit-aware re-runs
python3 tests/mutation_check.py      # proves each suite catches the defect it guards
```

None of these needs Claude Code or a network connection. CI runs all four on Linux and macOS, Python 3.9 and 3.12, on every push to `main` and every pull request.

In Claude Code, point the skill at a deck or a folder of decks:

```
/captioner ~/lectures/module5.pptx
/captioner ~/lectures/            # every .pptx in the folder
```

You get back: `<deck>_captioned.pptx` files plus a per-deck `<deck>_audit.csv`.

### Correcting captions after review
When a reviewer flags a specific slide — *"slide 5's box is too big"*, *"slide 12's caption overlaps the photo"*, *"slide 3's caption is wrong"* — fix it in place by slide number, without re-captioning the deck:

```bash
python3 scripts/fix_captions.py mydeck_captioned.pptx \
  --slide 5  --shrink-box \
  --slide 12 --move below \
  --slide 3  --text "Photo: solar farm at dusk" \
  --verify
```

Run with `--dry-run` first to preview. Each geometry change is re-checked against the placement rules and refused if it would create an overlap (override with `--force`). The deck is backed up to `<deck>.pptx.bak.pptx` before the first write. See [`SKILL.md`](SKILL.md) for the full review-note → fix table and iteration loop.

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

### `fix_captions.py` (per-slide in-place corrections)
| Flag | Effect |
|---|---|
| `--slide N` | Open a correction group for slide N (required, repeatable; flags reset per group) |
| `--which N` / `--hash HEX` | Target the Nth caption (top-to-bottom) / the caption for a given image hash |
| `--shrink-box` / `--height(-in)` / `--width(-in)` | Resize the caption box |
| `--move below\|above\|left\|right\|bottom` / `--nudge-x` / `--nudge-y` | Reposition the caption |
| `--font-size` / `--font-color` / `--italic` / `--font-name` | Restyle the caption text |
| `--bg-color ""` / `--border-color ""` | Make the caption card 100% transparent (true `noFill`) |
| `--text "..."` | Replace the caption wording |
| `--dry-run` / `--force` / `--no-backup` / `--verify` | Preview / override defect refusal / skip backup / re-scan after saving |

</details>

## Capabilities

1. **Subject-identifying photo captions** — brief, category-prefixed (`Photo: …`), no descriptive prose, no "image of…" filler.
2. **Category-aware prefixes** — distinguishes `Photo:`, `Illustration:`, `Diagram:`, `Chart:`, `Screenshot:`, `Logo:`, `Map:`, `Icon:` from visual context.
3. **Vision reads in-image text** — reproduces typography rendered as part of an image (a warning sign, a code screenshot).
4. **White-card background fill** — captions stay legible on any slide background, including dark section dividers.
5. **Text-aware 2D placement — captions never cover text.** Every text frame (title, body, plain text boxes) is a 2D obstacle the caption routes around: below the picture → above it → a band tucked into the picture's own bottom strip → otherwise the picture is left uncaptioned and flagged. Two captions never touch, a caption never lands inside a *different* picture, nothing intrudes the footer, and box height is sized to the rendered text so a label never spills onto its neighbor. The chosen placement is logged.
6. **SmartArt per-icon captioning** — each embedded PowerPoint Icon gets its own caption from its SVG `id` metadata (no vision call).
7. **Decorative-image triage** — repeating logos and bullet markers are flagged `[decorative]` and skipped.
8. **Text-only SmartArt skip** — diagrams whose only content is text labels are intentionally not captioned.
9. **Deck-level context** — a free-text context string biases caption vocabulary toward the subject.
10. **Audit CSV trail** — every action (`added`, `fallback-above`, `inside-bottom`, `overlay-fullbleed`, `flagged-no-slot`, `skipped-decorative`, `added-smartart-icon`, `skipped-smartart-text-only`, …) is logged.
11. **Idempotent re-runs** — captioner shapes carry a content-hash name; `--update-existing` regenerates cleanly with no duplicates, and a re-run without it on an already-captioned deck is refused (exit 4) rather than captioning it twice.
12. **Per-slide correction loop** — `fix_captions.py` turns a reviewer's plain-language notes into surgical, in-place fixes (resize, reposition, restyle, retext) by slide number, each re-checked against the placement rules, with `--dry-run` preview and `--verify` re-scan.

## What captioner does NOT do

- Write alt-text by default, or overwrite existing alt-text — the OOXML `descr` field is authored only with `--write-descr`, and any alt-text a deck already has is preserved untouched.
- Certify WCAG conformance, or verify that the alt-text it authors is *correct* — it populates the field for images that lacked one; a human should review the wording. Pair it with a full alt-text review for WCAG 2.1 AA conformance.
- Read off chart values — captions identify *what the picture is*, not what its data says.
- Modify embedded charts, the slide master, or any non-Picture shape.
- Process `.ppt` legacy format (python-pptx supports `.pptx` only).

## Reliability

What backs each run is the verification gate, not a track record. Every deck is checked by an independent verifier for six placement defect classes — text-overlap, footer intrusion, caption-inside-picture, caption-on-caption, text overflow, and a band caption burying a small/structural picture — in 2D coordinates, plus (with `--descr`) an alt-text coverage gate that confirms every picture is either given a `descr` value or marked decorative. The verifier exits non-zero on any defect, so a caption over text, off-slide, inside another picture, or overlapping another caption cannot ship silently. Source decks are preserved unmodified, and every run produces a reproducible per-deck audit CSV.

The tests are held to the same standard. `tests/mutation_check.py` injects the defects the suites exist to prevent (a caption allowed to overlap an obstacle, a hand-edited caption deleted on re-run, hand-edited or pre-existing alt-text overwritten) and requires the matching suite to go red and name the check that caught it.

## License

MIT — see [`LICENSE`](LICENSE). Copyright (c) 2026 Jake Benhart.

## Citation

```bibtex
@software{benhart_captioner_2026,
  author  = {Benhart, Jake},
  title   = {Captioner: pedagogical-clarity captions for PowerPoint decks},
  year    = {2026},
  version = {0.5.0},
  url     = {https://github.com/jbenhart44/captioner}
}
```
