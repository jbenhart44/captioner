---
name: captioner
description: Add brief compliance-style captions under every picture in PowerPoint (.pptx) decks. Reads each photo, generates a short subject-identifying caption (5-15 words), and inserts a visible italic text box under each picture. Supports dry-run, idempotent re-runs, recursive group traversal, configurable styling, and deck-level context.
---

# /captioner — PowerPoint Picture Caption Agent

## Purpose
Add **brief, compliance-style captions** under every picture in one or more `.pptx` files. The caption identifies *what the picture is* so a viewer (or accessibility reviewer) immediately knows the subject. Slide body text is assumed to handle the deeper detail.

## Trigger phrases
- "Run /captioner on `<file or folder>`"
- "Caption this deck"
- "Add picture captions to these slides"
- "Tag the pictures in `<deck>`"

## Caption style — STRICT RULES

1. **Brief** — 5 to 15 words target, hard cap 20.
2. **Subject-identifying, not descriptive**.
   - ✗ "Top-down view of a white ceramic mug filled with milky coffee, casting a long shadow on a vivid red surface."
   - ✓ "Photo: coffee mug on red background."
3. **Lead with category** when useful — `Photo:`, `Chart:`, `Diagram:`, `Screenshot:`, `Logo:`, `Map:`, `Illustration:`.
4. **No "image of" / "picture of"** filler.
5. **Decorative images** (small icons, bullet markers, slide-master logos repeated on every slide): caption value `[decorative]` — logged, NO visible text added to slide.
6. **Context-aware**: if you can see slide text or the user passed `--context "Lecture 3 — Module 5 overview"`, use it. Photo of cells on slide titled "Mitosis" → "Diagram: cell mitosis stages" not "Photo: pink shapes on white."

## Workflow (when /captioner is invoked)

### Step 1 — Parse arguments
- Single `.pptx` path: process that one file.
- Folder: process every `.pptx` in it (non-recursive). Files ending `_captioned.pptx` are auto-excluded.
- If user supplied a course / deck-level theme, capture it for the `--context` flag in step 2.
- If no arg: ask the user for a path.

### Step 2 — Extract images + manifest
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/extract_images.py" <input-path> <work-dir> [--context "<deck-context>"]
```
Produces `<work-dir>/manifest.json` and `<work-dir>/images/<deck>/<hash>.<ext>`. **Errored decks (corrupt / password-protected) are logged in the manifest, not crashed on.** Group shapes are traversed recursively — pictures nested inside groups are captured.

### Step 3 — Read images and generate brief captions
For each unique image:
1. Use the `Read` tool on the image file (Claude Code's vision, no external API).
2. Generate a caption per the STRICT RULES above.
3. If the manifest carries a `deck_context` for the picture's deck, use it to bias the caption.
4. Build `<work-dir>/captions.json` mapping `"<deck>/<hash>.<ext>": "<short caption>"`.

**Batch reads aggressively** — many Read calls in one tool message. 20-30 image reads per message is fine.

**Context hygiene between decks** (added 2026-05-12): When processing multiple decks in a single session, drop image reads from working context after each deck is finished. Concretely: complete one deck's `captions.json` end-to-end (read images → write JSON → run `apply_captions.py` → verify), then move to the next deck WITHOUT carrying the prior deck's image reads forward. If you anticipate context pressure across 5+ decks, dispatch each deck (or small batch) to a background subagent instead of doing it inline — the subagent's image reads are scoped to its own context window. Rationale: each deck has 15-50 images at ~50-200KB each; accumulating reads across 10+ decks balloons main-conversation context and can trigger summarization mid-batch, losing your place. Symptom that indicates this rule was violated: the conversation gets compacted while a multi-deck batch is in flight.

### Step 4 — DRY RUN first (recommended)
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/apply_captions.py" <work-dir> --dry-run
```
This emits the audit CSV showing **what would be added**, without modifying any `.pptx`. Surface the audit CSV to the user, ask if they want to proceed, edit `captions.json` if needed.

### Step 5 — Apply captions
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/apply_captions.py" <work-dir> \
  [--font-name Calibri] [--font-size 10] [--font-color 333333] [--italic true] \
  [--update-existing] [--gap-emu 50000] [--height-emu 400000]
```
Produces `<work-dir>/captioned_decks/<deck>_captioned.pptx` and `<work-dir>/audit/<deck>_audit.csv`.

**Re-runs**: if the user re-runs against the same source, the previous captioner text boxes are recognizable by their shape name prefix `captioner_caption_<hash>`. Pass `--update-existing` to strip the prior captions before adding new ones; otherwise re-runs will DUPLICATE captions (added behind the prior text box).

### Step 6 — Verify
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/verify.py" <work-dir>
```
Confirms every expected caption text box landed.

### Step 7 — Report
Tell the user:
- N captions added across M decks
- N skipped (decorative / no-caption / extract-failed)
- N deck-level errors (corrupt / password-protected, names listed)
- Edge cases (fallback-above, fallback-bottom counts)
- Output paths

## CLI flag reference

### `extract_images.py`
| Flag | Default | Effect |
|---|---|---|
| `--context "<text>"` | empty | Deck-level overarching context recorded in manifest for caption-generation step |

### `apply_captions.py`
| Flag | Default | Effect |
|---|---|---|
| `--dry-run` | off | Emits audit CSV; does NOT modify any `.pptx` |
| `--font-name` | Calibri | Caption font family |
| `--font-size` | 10 | Caption font size in pt |
| `--font-color` | 333333 | Caption hex color (no `#`) |
| `--bg-color` | FFFFFF | Caption text-box fill hex (white card behind text — readable on any slide background; `""` to disable for transparent) |
| `--border-color` | CCCCCC | Caption text-box border hex (`""` to disable) |
| `--italic` | true | Italic on/off (`true`/`false`) |
| `--gap-emu` | 50000 | EMU between picture and caption (~0.05") |
| `--height-emu` | 400000 | Caption box height (~0.44") |
| `--update-existing` | off | Strip prior captioner shapes before adding new ones (idempotent re-run) |
| `--no-smartart` | off | Disable SmartArt icon captioning (on by default; auto-extracts icon names from SVG metadata) |
| `--quiet` | off | Suppress per-deck progress |

## What this skill does NOT do
- Write to the WCAG `descr` alt-text XML field (separate workflow).
- Read off chart values — captions identify "this is a chart of X", they don't read figures.
- Modify the picture itself, slide master, or existing slide shapes (other than captioner-prior shapes via `--update-existing`).
- Touch embedded charts (preserved on save; no caption added since they're not Pictures).

## SmartArt handling (added 2026-05-12)

PowerPoint SmartArt diagrams are NOT `MSO_SHAPE_TYPE.PICTURE` shapes — they are `graphicFrame` elements with a `drawingml/2006/diagram` URI. The captioner has explicit handling for them:

1. **Text-only SmartArts** (no embedded icons — just a layout with text labels): **no caption added**. The visible text already lives in the slide's accessible text layer, so an alt-caption would be redundant clutter.
2. **SmartArts with embedded icons** (PowerPoint "Insert > Icons" inserts SVG icons into SmartArt cells): each icon gets its **own short caption directly under it**, positioned via the icon's geometry from `ppt/diagrams/drawing*.xml`. No combined diagram caption is added.

### How icon names are derived (no vision pass needed)
PowerPoint Icons store their human-readable name in the SVG `id` attribute on the root element, e.g. `<svg id="Icons_Checkmark">`, `<svg id="Icons_VideoCamera">`. The captioner walks the slide's `diagramDrawing` part rels, for each `<dsp:sp>` finds the `<asvg:svgBlip r:embed=rIdN>` extension (NOT `<a:blip>` — that points at the PNG fallback, which has no name), parses the SVG's `id`, and converts CamelCase → space-separated lowercase. So `Icons_VideoCamera` becomes the caption `video camera`.

### Caption style for icons
1-3 words, lowercase, no "icon" suffix. Placed in a small white card directly under the icon's bounding box.

### Audit actions for SmartArt
| Action | Meaning |
|---|---|
| `added-smartart-icon` | Per-icon caption placed under an icon inside a SmartArt frame |
| `skipped-smartart-text-only` | SmartArt with no embedded icons — intentionally not captioned |

To disable all SmartArt handling, pass `--no-smartart`.

## Audit CSV columns

`slide, pic_id, image_hash, caption, char_len, in_group_depth, action`

`action` values:
- `added` — caption inserted below picture
- `fallback-above` — caption placed above (no room below)
- `fallback-bottom` — caption stuffed at slide bottom (no room above or below)
- `dry-run-would-{added|fallback-above|fallback-bottom}` — what would happen in non-dry-run
- `skipped-decorative` — caption value was `[decorative]`
- `skipped-no-caption` — no entry in captions.json for this image
- `skipped-image-extract-failed` — python-pptx could not read the image blob
- `added-smartart-icon` — per-icon caption inserted under a SmartArt icon (auto-generated from SVG metadata)
- `skipped-smartart-text-only` — SmartArt frame with no embedded icons; intentionally not captioned

`in_group_depth` shows how deeply nested the picture was inside group shapes (0 = top level).

## Edge-case behavior

| Case | Behavior |
|---|---|
| Corrupt `.pptx` | Caught in extract_images.py; deck recorded in manifest with `error` field; skipped by apply_captions; included in summary |
| Password-protected | Same as corrupt — caught + logged, not crashed |
| `.ppt` legacy format | NOT supported (python-pptx only handles `.pptx`); skipped |
| Picture nested in GROUP shape | Recursively traversed; depth recorded in audit |
| SmartArt-derived group | python-pptx raises NotImplementedError on `.shapes` access; caught + skipped |
| SmartArt frame with icons | Each icon captioned individually under it (auto from SVG `id` metadata); no combined diagram caption |
| SmartArt frame, text-only | Skipped (no caption); slide text already accessible |
| Caption on dark slide background | Caption sits in a white card (`--bg-color FFFFFF` default) with light gray border, so it's readable regardless of slide background |
| Hidden shape (`cNvPr hidden="1"`) | Skipped — invisible shapes don't need visible captions |
| Picture fills entire slide | Caption falls back to slide-bottom; logged as `fallback-bottom` |
| Re-run on `_captioned.pptx` output | Without `--update-existing`: duplicates captions. With: strips prior captioner shapes and re-adds. |
| Re-run on source with prior captioner shapes | Same as above — captioner shapes carry name prefix `captioner_caption_<hash>` |
| Multi-language slide text | Context extraction works; captions still emitted in English |
| Animated picture / video | Treated as Picture if shape_type matches; otherwise skipped silently |

## Provenance

Captioner was built and hardened across a real production run of 1,132 captions across 32 PowerPoint decks in three graduate engineering courses (Summer 2026). It was hardened following an independent external code review (recursive group traversal, dry-run, error handling, idempotency marker, style flags, deck context, hidden-shape skip, progress prints) and extended after production feedback (white-card background fill for dark-slide legibility; SmartArt per-icon captioning from SVG `id` metadata; the context-hygiene workflow rule for large batches). See CHANGELOG.md for the dated release history.
