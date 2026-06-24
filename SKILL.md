---
name: captioner
description: Add brief compliance-style captions under every picture in PowerPoint (.pptx) decks. Reads each photo, generates a short subject-identifying caption (5-15 words), and inserts a visible italic text box under each picture. Supports dry-run, idempotent re-runs, recursive group traversal, configurable styling, deck-level context, and an optional whitelist-aware spell/QC pass.
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
6. **Context-aware**: if you can see slide text or the user passed `--context "Biology 101 Lecture 3"`, use it. Photo of cells on slide titled "Mitosis" → "Diagram: cell mitosis stages" not "Photo: pink shapes on white."

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
Produces `<work-dir>/manifest.json` and `<work-dir>/images/<deck>/<hash>.<ext>`. **Errored decks (corrupt / password-protected) are logged in the manifest, not crashed on.**

**Picture enumeration (raw-OOXML).** extract/apply/verify enumerate every `<p:pic>` in each slide's shape tree via lxml (`scripts/_oxml_pics.py`), NOT python-pptx `shape_type == PICTURE`. This matters: python-pptx classifies a picture content-placeholder (`<p:pic>` carrying `<p:ph type="pic"/>`) as `PLACEHOLDER`, **not** `PICTURE`, so the old walk silently skipped every placeholder-hosted picture (a real-world deck where the extractor saw only ~8% of its pictures). The `.//p:pic` descent also captures group-nested pictures, so the three scripts now enumerate an identical population. `<mc:Choice>` subtrees are stripped so an `<mc:AlternateContent>` picture is counted once (the Fallback); linked pictures with no `r:embed` are skipped + logged; hidden shapes are skipped. The SHA-256 image hash and `captions.json` key format are byte-unchanged (prior reuse indexes still match).

### Step 3 — Read images and generate brief captions
For each unique image:
1. Use the `Read` tool on the image file (Claude Code's vision, no external API).
2. Generate a caption per the STRICT RULES above.
3. If the manifest carries a `deck_context` for the picture's deck, use it to bias the caption.
4. Build `<work-dir>/captions.json` mapping `"<deck>/<hash>.<ext>": "<short caption>"`.

**Batch reads aggressively** — many Read calls in one tool message. 20-30 image reads per message is fine.

**Context hygiene between decks**: When processing multiple decks in a single session, drop image reads from working context after each deck is finished. Concretely: complete one deck's `captions.json` end-to-end (read images → write JSON → run `apply_captions.py` → verify), then move to the next deck WITHOUT carrying the prior deck's image reads forward. If you anticipate context pressure across 5+ decks, dispatch each deck (or small batch) to a background subagent instead of doing it inline — the subagent's image reads are scoped to its own context window. Rationale: each deck has 15-50 images at ~50-200KB each; accumulating reads across 10+ decks balloons main-conversation context and can trigger summarization mid-batch, losing your place. Symptom that indicates this rule was violated: the conversation gets compacted while a multi-deck batch is in flight.

### Step 4 — DRY RUN first (recommended)
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/apply_captions.py" <work-dir> --dry-run
```
This emits `audit/<deck>_audit_dryrun.csv` showing **what would be added**, without modifying any `.pptx`. (Dry-run uses its own filename so it can never overwrite a prior real `_audit.csv`.) Surface it to the user, ask if they want to proceed, edit `captions.json` if needed.

### Step 4.5 — Name verification (REQUIRED before presenting any name fix)
QC runs by default (see Step 5). For every `qc/<deck>_spellcheck.csv` row with **`verify_name=true`** (the flagged term looks like a proper noun — person, company, product, place), you MUST web-verify the canonical spelling with `WebSearch`/`WebFetch` **before** presenting it to the user as a correction. Do not assert a name fix from model memory alone. If the web check confirms the suspected term is actually correct, drop it (do not "fix" a real name); if it confirms a misspelling, present the web-verified spelling with its source. Rationale: an earlier run nearly "corrected" valid names; in-document context + a web check is the bar.

### Step 5 — Apply captions (QC is ON by default)
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/apply_captions.py" <work-dir> \
  [--font-name Calibri] [--font-size 10] [--font-color 333333] [--italic true] \
  [--update-existing] [--gap-emu 50000] [--height-emu 400000] \
  [--quick | --no-spellcheck | --no-dateqc]
```
**Default behaviour: spell-check + date/template QC both run** unless the user asks for the quick (captioning-only) variant. Toggles: `--quick` = captioning only, skip all QC; `--no-spellcheck` / `--no-dateqc` = disable just that one. (The old `--spellcheck` flag is still accepted as a no-op since QC is now default-on.)

Produces `<work-dir>/captioned_decks/<deck>_captioned.pptx`, `<work-dir>/audit/<deck>_audit.csv`, and (QC default-on) `<work-dir>/qc/<deck>_spellcheck.csv` + `<work-dir>/qc/<deck>_qc.csv`. QC artifacts live in their **own `qc/` directory**, never the caption `audit/` dir.

**pyspellchecker is required for the default spell-check.** It is an optional dependency only in the sense that captioning still works without it — but because spell-check is now default-on, a missing `pyspellchecker` is treated as a **loud failure**: captioner prints a banner, drops `qc/SPELLCHECK_NOT_RUN.txt`, and **exits 3** (captions still applied). Either run via a Python that has `pyspellchecker` installed (`pip install pyspellchecker`, or a virtualenv if your system Python is PEP 668 / externally-managed), or pass `--no-spellcheck` / `--quick` to explicitly proceed without it.

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

### Step 8 — Per-slide corrections (review loop, `fix_captions.py`)
After a deck is delivered, whoever reviews it will spot per-slide issues — *"slide 5's caption box is too big"*, *"slide 12's caption overlaps the photo, move it below"*, *"slide 3's caption is wrong"*. `fix_captions.py` applies those fixes **in place, by slide number**, editing ONLY the named caption shapes — no re-extract, no re-read of images, no re-caption of the rest of the deck. It backs up the deck to `<deck>.pptx.bak.pptx` before the first write.

```bash
python3 scripts/fix_captions.py <deck_captioned.pptx> \
  --slide 5  --shrink-box \
  --slide 12 --move below \
  --slide 3  --font-size 8 --font-color 222222 \
  --slide 7  --text "Photo: solar farm at dusk" \
  --verify
```

Each `--slide N` opens a fresh correction group; its flags reset at the next `--slide`, so groups are independent. The four families of fix:

| Fix | Flags |
|---|---|
| **Resize** (box too big/small) | `--shrink-box` (fit to text) · `--height EMU` / `--height-in INCHES` · `--width EMU` / `--width-in INCHES` |
| **Reposition** (wrong spot / overlap) | `--move {below\|above\|left\|right\|bottom}` (re-placed cleanly next to its own picture) · `--nudge-x EMU` / `--nudge-y EMU` (relative; +y = down) |
| **Restyle** (font too big / wrong color) | `--font-size PT` · `--font-color HEX` · `--italic {true\|false}` · `--font-name NAME` · `--bg-color HEX` (`""`=transparent) · `--border-color HEX` (`""`=off) |
| **Retext** (wrong/bad caption) | `--text "new wording"` |

Target selection when a slide has more than one caption: `--which N` (1-based, top-to-bottom) or `--hash HEX` (the image hash in the caption's shape name). Default = every caption on the slide.

**Safety:** every geometry change (resize/move/nudge) is re-checked against the SAME 2D obstacle model the placer and `verify.py` use (title/body text, other pictures, other captions, footer, off-slide). A fix that would CREATE a defect is **refused and reported** — pass `--force` to apply it anyway, or `--dry-run` to preview every group without writing. `--verify` re-scans the saved deck and exits 4 if any caption placement defect remains. **Exit codes:** `0` clean · `4` a fix was refused (would create/leave a defect) or `--verify` found a defect · `2` usage/IO/no-matching-caption. Retext/restyle-only changes don't move the box, so they're never refused on geometry.

### The iteration loop — from "this caption looks wrong" to a clean fix
The tool is built to be driven from plain-language review notes. The reviewer reads the
delivered deck, writes down each problem **by slide number**, and you translate the note
into one `--slide` group. Always work the loop the same way:

1. **Preview, don't commit blind.** Run the fix with `--dry-run` first. It prints, per
   group, exactly what would change and — for any geometry change — whether the new
   position would create a defect. Nothing is written.
2. **Apply.** Drop `--dry-run`; the deck is backed up to `<deck>.pptx.bak.pptx` and saved.
3. **Re-verify.** Add `--verify` (or run `verify.py`) so the saved deck is re-scanned for
   placement defects before you hand it back. A clean exit (0) means ship it.
4. **If a fix is refused** (exit 4: it would create a text/photo/caption overlap), the tool
   tells you which defect. Adjust the note — e.g. switch `--move below` to `--move right`,
   or pair a `--shrink-box` with the move so the smaller box fits — and re-run. Use
   `--force` only when you have *looked at the slide* and decided the overlap is acceptable.
5. **Iterate until `--verify` is clean.** Multiple review notes go in ONE command as
   multiple `--slide` groups; re-running is safe and idempotent (it edits the same named
   shapes in place, never duplicates them).

**Common review note → fix:**

| What the reviewer says | Likely cause | Fix |
|---|---|---|
| "slide 5's caption box is too big / too tall" | nominal box height larger than the text needs | `--slide 5 --shrink-box` |
| "slide 5's caption is too wide / runs off the edge" | wide nominal box | `--slide 5 --width-in 2.5` (then `--verify`) |
| "slide 12's caption overlaps the photo / a number on it" | image-baked text the placer can't see (geometry is blind to text *inside* a picture) | `--slide 12 --move below` (or `right`/`above`); if all sides are crowded, `--shrink-box` then `--move` |
| "slide 8's caption sits on the title / body text" | a layout the placer mis-measured (e.g. a tall multi-line title) | `--slide 8 --move below`, or `--nudge-y 200000` to push it down |
| "slide 3's caption is wrong / describes the wrong thing" | a confidently-wrong vision caption (the highest-visibility, otherwise-silent failure) | `--slide 3 --text "Photo: <correct subject>"` |
| "the caption text is too small / too faint to read" | font size or color | `--slide N --font-size 11 --font-color 000000` |
| "the white box behind the caption looks bad on this slide" | the default white card clashes with the background | `--slide N --bg-color "" --border-color ""` (100% transparent — both the fill and the border become `noFill`) |
| "two captions on one slide overlap each other" | two pictures close together | move one: `--slide N --which 2 --move right` |
| slide has several pictures and you mean a specific one | ambiguous target | add `--which N` (top-to-bottom) or `--hash <image-hash from the audit CSV>` |

**Known limits the loop cannot fully solve (escalate to a manual slide edit):**
- **Text rasterized inside a picture** (a number/label baked into the image, no text frame)
  is invisible to the geometry model, so an auto-placed caption *can* land on it. The fix
  is to `--move` the caption to a side with clear space; if the picture is full-bleed and no
  side is clear, the slide needs a manual layout change — caption coverage there is a
  genuine gap, not something the tool hides.
- **Caption *correctness*** is never auto-verified. `--verify` checks placement, not whether
  the words match the picture. Spot-check wording before any external handoff and fix wrong
  ones with `--text`.
- A picture with **no clean slot at all** was left uncaptioned by design (look for
  `flagged-no-slot` / `overlay-fullbleed` in the audit CSV). `fix_captions.py` edits captions
  that already exist; it does not create a caption for a skipped picture — that one needs a
  manual layout change first.

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
| `--bg-repeat-threshold` | 4 | An identical image on ≥ N slides is a repeated background — not captioned and not a placement obstacle (`0` disables) |
| `--quiet` | off | Suppress per-deck progress |
| `--quick` | off | **Captioning ONLY** — skip ALL QC (spell-check + date/template scan) |
| `--no-spellcheck` | off | Disable just the spell-check pass (date/template QC still runs) |
| `--no-dateqc` | off | Disable just the date/doubled-word/leftover-template scan |
| `--spellcheck` | (no-op) | Back-compat alias — spell-check is now **default-on**; flag accepted, does nothing |

### `fix_captions.py` (per-slide in-place corrections)
Each `--slide N` opens an independent correction group (flags reset at the next `--slide`). See Step 8 for the full workflow.

| Flag | Effect |
|---|---|
| `--slide N` | Open a correction group targeting slide N (required, repeatable) |
| `--which N` | Target the Nth caption on the slide (1-based, top-to-bottom). Default: all |
| `--hash HEX` | Target only the caption whose shape name ends with this image hash |
| `--shrink-box` | Resize the box height to fit its text |
| `--height EMU` / `--height-in IN` | Set box height (EMU or inches) |
| `--width EMU` / `--width-in IN` | Set box width (EMU or inches) |
| `--move {below\|above\|left\|right\|bottom}` | Re-place the caption cleanly relative to its own picture |
| `--nudge-x EMU` / `--nudge-y EMU` | Relative shift (+y = down) |
| `--font-size PT` / `--font-color HEX` / `--italic {true\|false}` / `--font-name NAME` | Restyle the caption text |
| `--bg-color HEX` (`""`=transparent) / `--border-color HEX` (`""`=off) | Restyle the caption card |
| `--text "..."` | Replace the caption wording |
| `--force` | Apply a geometry change even if it would create a placement defect (still reported) |
| `--dry-run` | Preview every group; write nothing |
| `--no-backup` | Skip the `.bak.pptx` copy |
| `--verify` | After saving, re-scan the deck for placement defects (exit 4 if any remain) |

## QC: spell-check + date/template scan (DEFAULT-ON, flag-only)

QC runs **by default** on every apply (the user must opt OUT via `--quick`/`--no-*`, not opt in). Two FLAG-ONLY scanners, written to the **`qc/` directory** (never the caption `audit/` dir):

**`qc/<deck>_spellcheck.csv`** — columns `slide, source, term, suggestion, known_bad, verify_name, context`. Two sources:
- `source=caption` — captions captioner itself generated (its own output quality).
- `source=slide-text` — the instructor's slide title/body/table text (captioner's own added shapes excluded).

**`qc/<deck>_qc.csv`** — columns `slide, source, kind, detail, context`, where `kind` ∈ `doubled-word` | `leftover-template` | `date-review`. Generic only (no course-specific year/code rules); `date-review` rows are informational, not asserted wrong.

Hard guarantees:
- **Flag-only. Captioner never edits the `.pptx` and never auto-corrects.** The CSVs are review aids; a human decides.
- **`verify_name=true`** marks a likely proper noun → the workflow MUST web-verify it before any name fix is presented (Step 4.5). Do not "fix" names from memory.
- **Whitelist-aware so it does NOT flag non-issues.** Bundled, user-extensible `scripts/spellcheck_whitelist.txt` suppresses domain vocab/brands. Plural-of-acronym (`NPVs`, `IRRs`, `KPIs`, `GPTs`) is auto-skipped heuristically. Hard proper-noun typos a dictionary can't suggest are surfaced with `known_bad=true`.
- Lines with URLs/email/product domains skipped; ALL-CAPS acronyms, digit/hyphen/contraction tokens, sub-4-char tokens not flagged.
- Read-only in `--dry-run` and normal mode alike.
- **pyspellchecker is required for the default spell-check.** Missing it is a LOUD failure: banner + `qc/SPELLCHECK_NOT_RUN.txt` + **exit 3** (captions still applied). Use a Python with `pyspellchecker`, or pass `--no-spellcheck`/`--quick` to proceed deliberately without it.

## Placement engine (text-aware 2D)

A caption is **never** placed where it covers text. The placer treats as 2D obstacles
every text frame on the slide — title, body, AND plain text boxes / auto-shapes — plus
every other (non-background) picture and every already-placed caption. Each obstacle is
narrowed to its estimated **visible-text region**, sized **font- and line-break-aware**
(explicit `\n`/`\x0b` breaks and the actual font size are counted), so a multi-line or
large-font title is covered correctly — while a caption below a tall title's single line
of text is still not falsely blocked by the title's empty box. Placement order per picture:

1. **Below** the picture (clean external slot) — clears all obstacles incl. the picture's
   own box; a near-zero tolerance against other captions keeps cards from touching.
2. **Small bottom-of-picture band** (`inside-bottom`) — preferred over above/beside so the
   caption stays *with* its picture rather than drifting up into title/body text. The band
   renders small (`BAND_FONT_PT`) — a thin strip overlapping a sliver of image — which also
   lets it fit tight spots a full-size caption can't. Skipped for a small/structural picture
   where a band would bury its content.
3. **Above / beside** the picture — last resort before skipping (with a horizontal nudge).
4. **Skip + flag** (`flagged-no-slot`) only if even the small band would cover text.

A **repeated background** (the same image hash on ≥ `--bg-repeat-threshold` slides — a
full-slide template texture) is detected, left uncaptioned, AND excluded from the picture-
obstacle set, so it never blocks the real pictures on those slides.

Caption box height is estimated from the text + width (the box auto-sizes) and used for
every placement/overlap/footer decision, so a grown caption never spills onto a neighbour.
The own picture is matched by **geometry, not `pic_id`** (which can repeat across pictures).
`verify.py` enforces five patterns — text-overlap (T), footer (B), in-picture (C; band/icon
exempt), caption-caption (E), and missing captions — exiting non-zero on any defect.

## What this skill does NOT do
- Write to the WCAG `descr` alt-text XML field (separate workflow).
- Auto-correct spelling — the default QC only *reports* suspects (and web-verifies names before suggesting); it never rewrites slide or caption text.
- Read off chart values — captions identify "this is a chart of X", they don't read figures.
- Modify the picture itself, slide master, or existing slide shapes (other than captioner-prior shapes via `--update-existing`).
- Touch embedded charts (preserved on save; no caption added since they're not Pictures).

## SmartArt handling

PowerPoint SmartArt diagrams are NOT `MSO_SHAPE_TYPE.PICTURE` shapes — they are `graphicFrame` elements with a `drawingml/2006/diagram` URI. The captioner has explicit handling for them:

1. **Text-only SmartArts** (no embedded icons — just a layout with text labels): **no caption added**. The visible text already lives in the slide's accessible text layer, so an alt-caption would be redundant clutter.
2. **SmartArts with embedded icons** (PowerPoint "Insert > Icons" inserts SVG icons into SmartArt cells): each icon gets its **own short caption directly under it**, positioned via the icon's geometry from `ppt/diagrams/drawing*.xml`. No combined diagram caption is added.

### How icon names are derived (no vision pass needed)
PowerPoint Icons store their human-readable name in the SVG `id` attribute on the root element, e.g. `<svg id="Icons_Checkmark">`, `<svg id="Icons_VideoCamera">`. The captioner walks the slide's `diagramDrawing` part rels, for each `<dsp:sp>` finds the `<asvg:svgBlip r:embed=rIdN>` extension (NOT `<a:blip>` — that points at the PNG fallback, which has no name), parses the SVG's `id`, and converts CamelCase → space-separated lowercase. So `Icons_VideoCamera` becomes the caption `video camera`.

### Caption style for icons
1-3 words, lowercase, no "icon" suffix. Placed in a small white card directly under the icon's bounding box. **Icon-proportional sizing:** the font scales with the icon (5pt for tiny icons, up to 8pt) and the box hugs the text (width capped near the icon, height sized to the wrapped line count), so a caption never dwarfs a small icon. If an icon caption cannot be placed without overlapping another caption (e.g. a dense, near-footer icon strip), it is skipped + flagged rather than overlapped.

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
- `fallback-bottom` — caption placed at slide bottom (footer-safe; no room above or below)
- `inside-bottom` — caption placed in its own picture's bottom strip (no text-clear external slot existed)
- `overlay-fullbleed` — SKIPPED: caption would land inside a picture; flagged for review
- `flagged-no-slot` — SKIPPED: no position clears every obstacle without covering text; flagged
- `dry-run-would-{added|fallback-above|fallback-bottom}` — what would happen in non-dry-run
- `skipped-decorative` — caption value was `[decorative]`
- `skipped-decorative-background` — SKIPPED: a repeated background/template image (same hash on ≥ `--bg-repeat-threshold` slides); not captioned and not treated as a placement obstacle
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
| Picture fills entire slide | No text-clear slot -> caption placed in the picture's bottom strip (`inside-bottom`) or skipped + flagged; never over text |
| Re-run on `_captioned.pptx` output | Without `--update-existing`: duplicates captions. With: strips prior captioner shapes and re-adds. |
| Re-run on source with prior captioner shapes | Same as above — captioner shapes carry name prefix `captioner_caption_<hash>` |
| Multi-language slide text | Context extraction works; captions still emitted in English |
| Animated picture / video | Treated as Picture if shape_type matches; otherwise skipped silently |

## Provenance
Battle-tested on a multi-hundred-deck graduate-course remediation run. Hardened iteratively from real production feedback:

- **Recursive group traversal, dry-run, fail-safe error handling, idempotency marker, style flags, deck context, hidden-shape skip, progress output.**
- **White-card background fill** for caption text boxes (`--bg-color`/`--border-color`, default-on) so captions stay legible on dark/section-divider slide backgrounds.
- **SmartArt icon captioning**: each embedded PowerPoint Icon gets its own short caption, name pulled from SVG `id` metadata (no vision pass); text-only SmartArts skipped.
- **Context-hygiene workflow**: finish one deck end-to-end before reading the next deck's images; for 5+ deck batches, dispatch via background subagents so image reads stay scoped.
- **Raw-OOXML picture enumeration** replacing the python-pptx `shape_type` walk, after a real deck where the old walk saw only ~8% of its pictures (placeholder-hosted `<p:pic>` are typed `PLACEHOLDER`, not `PICTURE`). Adversarially verified; SHA-256 hash + `captions.json` key format unchanged.
- **Default-on whitelist-aware spell/QC pass** with proper-noun web-verification guidance and fail-loud behaviour when `pyspellchecker` is absent.
- **Placement hardening**: placeholder-inherited geometry resolved; captions kept on-slide and clear of footer and title placeholders.
