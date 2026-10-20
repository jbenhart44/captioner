---
name: captioner
description: Add brief subject-identifying visible captions under every picture in PowerPoint (.pptx) decks, and optionally write WCAG alt-text (the OOXML descr attribute) for screen readers. Reads each photo, generates a short caption (5-15 words), inserts a visible italic text box under each picture, and (with --write-descr) authors descr for images that have none while preserving any existing alt-text. Supports dry-run, idempotent re-runs, recursive group traversal, configurable styling, deck-level context, 2D-aware placement (never covers title/body/footer/another picture), and a verification gate for both caption placement and alt-text coverage.
---

# /captioner — PowerPoint Picture Caption Agent

## Purpose
Add **brief, subject-identifying visible captions** under every picture in one or more `.pptx` files. The caption identifies *what the picture is* so a viewer immediately knows the subject. Slide body text is assumed to handle the deeper detail.

**Scope (read first):** captioner writes two independent things — a *visible* pedagogical caption on the slide canvas, and (opt-in, with `--write-descr`) the OOXML `descr` (alt-text) attribute that screen readers consume, for pictures that have none. Alt-text is **preserve-by-default**: existing alt-text is never overwritten, and decorative images are marked decorative rather than described. Populating `descr` fills a required field but does **not** by itself certify WCAG 2.1 conformance (it authors one attribute for a subset of images; it does not touch charts/embedded objects/non-picture shapes), and `verify.py` validates caption **presence/placement** and alt-text **coverage**, not caption or alt-text **correctness** — a human should review authored wording (see "Caption accuracy" and "Known limits" below).

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
6. **Context-aware**: if you can see slide text or the user passed `--context "Biology 101 — cell mitosis and the cell cycle"`, use it. Photo of cells on slide titled "Mitosis" → "Diagram: cell mitosis stages" not "Photo: pink shapes on white."

## Workflow (when /captioner is invoked)

### Step 1 — Parse arguments
- Single `.pptx` path: process that one file.
- Folder: process every `.pptx` in it (non-recursive). Files ending `_captioned.pptx` are auto-excluded.
- If user supplied a course / deck-level theme, capture it for the `--context` flag in step 2.
- If no arg: ask the user for a path.

#### Step 1a — ASK the professor/user for domain context (REQUIRED, added 2026-07-19)
Before extracting, **ask the professor or user who owns the deck for the domain/subject context of the material**, and pass it to every script via `--context`. Do this at the **start of every run** — good captions and (v0.3) authored alt-text are frequently **domain-specific**: the right words for an MRI scan, a perceptron diagram, a supply-chain flowchart, or a course-specific case study come from knowing the subject, not from the pixels alone. A generic caption ("Diagram: boxes and arrows") is worse than a domain-aware one ("Diagram: perceptron with weighted inputs and step activation").

Ask a short, concrete question, e.g.: *"What course and topic are these slides for? A one-line subject context (e.g. 'Biology 101 — cell mitosis and the cell cycle') helps me write domain-accurate captions and alt-text."* Then thread the answer through as `--context "<their answer>"` on `extract_images.py` (and it flows into the manifest's `deck_context`, biasing Step 3's caption generation).

- If the user already gave the context in their request, reuse it — do not re-ask.
- If they decline or have none, proceed without it (the tool still works; captions are just less domain-tuned).
- **This matters most for the images that have NO existing alt-text** — the ones the tool will **author** (`descr-authored`). Images that already carry alt-text are preserved untouched (see Step 5 / v0.3 preserve-by-default), so the context you gather is spent where it actually changes the output.
- CLI fallback: when a human runs `extract_images.py` directly with no `--context`, the script now prompts for this interactively (pass `--no-context-prompt` to suppress in scripted runs). The AI-skill path should still ASK per this step rather than rely on the CLI prompt, since Claude runs the script non-interactively.

#### Step 1b — ASK which output the professor wants (visible captions / alt-text / both)
The visible on-slide caption and the screen-reader alt-text (`descr`) are two independent layers, and the professor may want one, the other, or both. **Ask before running** (unless the request already made it clear), e.g.: *"Do you want (a) the WCAG alt-text only — accessible to screen readers, nothing added to how the slide looks; (b) the visible captions only; or (c) both?"* Then map the answer to `apply_captions.py`:
- **(a) Alt-text only** → `--descr-only` (implies `--write-descr`, adds NO visible caption text boxes). This is the accessibility-only path — the deck looks unchanged but screen readers get the descriptions.
- **(b) Visible captions only** → default (no `--write-descr`): on-slide caption cards, alt-text field left alone.
- **(c) Both** → `--write-descr`: visible captions + authored alt-text.
Reason to ask: adding visible caption cards changes the slide's appearance, which some instructors do not want; others want only the on-slide labels and not the invisible field. Defaulting silently to one mode surprises the other half of users. Existing alt-text is preserved and decorative images are marked decorative regardless of mode.

### Step 2 — Extract images + manifest
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/extract_images.py" <input-path> <work-dir> [--context "<deck-context>"]
```
Produces `<work-dir>/manifest.json` and `<work-dir>/images/<deck>/<hash>.<ext>`. **Errored decks (corrupt / password-protected) are logged in the manifest, not crashed on.**

**Picture enumeration (raw-OOXML, fixed 2026-05-18).** extract/apply/verify enumerate every `<p:pic>` in each slide's shape tree via lxml (`scripts/_oxml_pics.py`), NOT python-pptx `shape_type == PICTURE`. This matters: python-pptx classifies a picture content-placeholder (`<p:pic>` carrying `<p:ph type="pic"/>`) as `PLACEHOLDER`, **not** `PICTURE`, so the old walk silently skipped every placeholder-hosted picture (a prior production run saw ~8% of its real pictures affected). The `.//p:pic` descent also captures group-nested pictures, so the three scripts now enumerate an identical population. `<mc:Choice>` subtrees are stripped so an `<mc:AlternateContent>` picture is counted once (the Fallback); linked pictures with no `r:embed` are skipped + logged; hidden shapes are skipped. The SHA-256 image hash and `captions.json` key format are byte-unchanged (prior reuse indexes still match).

### Step 3 — Read images and generate brief captions
For each unique image:
1. Use the `Read` tool on the image file (Claude Code's vision: no separate API or key, but Claude Code sends the image to its model provider, Anthropic's API by default).
2. Generate a caption per the STRICT RULES above.
3. If the manifest carries a `deck_context` for the picture's deck, use it to bias the caption.
4. Build `<work-dir>/captions.json` mapping `"<deck>/<hash>.<ext>": "<short caption>"`.

**Batch reads aggressively** — many Read calls in one tool message. 20-30 image reads per message is fine.

**Context hygiene between decks** (added 2026-05-12): When processing multiple decks in a single session, drop image reads from working context after each deck is finished. Concretely: complete one deck's `captions.json` end-to-end (read images → write JSON → run `apply_captions.py` → verify), then move to the next deck WITHOUT carrying the prior deck's image reads forward. If you anticipate context pressure across 5+ decks, dispatch each deck (or small batch) to a background subagent instead of doing it inline — the subagent's image reads are scoped to its own context window. Rationale: each deck has 15-50 images at ~50-200KB each; accumulating reads across 10+ decks balloons main-conversation context and can trigger summarization mid-batch, losing your place. Symptom that indicates this rule was violated: the conversation gets compacted while a multi-deck batch is in flight.

### Step 4 — DRY RUN first (recommended)
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/apply_captions.py" <work-dir> --dry-run
```
This emits `audit/<deck>_audit_dryrun.csv` showing **what would be added**, without modifying any `.pptx`. (Dry-run uses its own filename so it can never overwrite a prior real `_audit.csv`.) Surface it to the user, ask if they want to proceed, edit `captions.json` if needed.

### Step 4.5 — Name verification (REQUIRED before presenting any name fix)
QC runs by default (see Step 5). For every `qc/<deck>_spellcheck.csv` row with **`verify_name=true`** (the flagged term looks like a proper noun — person, company, product, place), you MUST web-verify the canonical spelling with `WebSearch`/`WebFetch` **before** presenting it to the user as a correction. Do not assert a name fix from model memory alone. If the web check confirms the suspected term is actually correct, drop it (do not "fix" a real name); if it confirms a misspelling, present the web-verified spelling with its source. Rationale: the 2026-05-18 run nearly "corrected" valid names; in-document context + a web check is the bar.

### Step 5 — Apply captions (QC is ON by default)
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/apply_captions.py" <work-dir> \
  [--font-name Calibri] [--font-size 10] [--font-color 333333] [--italic true] \
  [--update-existing] [--gap-emu 50000] [--height-emu 400000] \
  [--quick | --no-spellcheck | --no-dateqc]
```
**Default behaviour: spell-check + date/template QC both run** unless the user asks for the quick (captioning-only) variant. Toggles: `--quick` = captioning only, skip all QC; `--no-spellcheck` / `--no-dateqc` = disable just that one. (The old `--spellcheck` flag is still accepted as a no-op since QC is now default-on.)

Produces `<work-dir>/captioned_decks/<deck>_captioned.pptx`, `<work-dir>/audit/<deck>_audit.csv`, and (QC default-on) `<work-dir>/qc/<deck>_spellcheck.csv` + `<work-dir>/qc/<deck>_qc.csv`. QC artifacts live in their **own `qc/` directory**, never the caption `audit/` dir.

**pyspellchecker is required for the default spell-check.** It is an optional dependency only in the sense that captioning still works without it — but because spell-check is now default-on, a missing `pyspellchecker` is treated as a **loud failure**: captioner prints a banner, drops `qc/SPELLCHECK_NOT_RUN.txt`, and **exits 3** (captions still applied). Either run via a Python that has `pyspellchecker` installed, or pass `--no-spellcheck` / `--quick` to explicitly proceed without it. (If a system Python blocks `pip install` under PEP 668, install into a virtual environment.)

**Re-runs**: if the user re-runs against the same source, the previous captioner text boxes are recognizable by their shape name prefix `captioner_caption_<hash>`. Pass `--update-existing` to strip the prior captions before adding new ones; without it, a deck that already holds captioner shapes is refused (exit 4, nothing written for it), so captions are never duplicated.

### Step 6 — Verify (6-pattern placement gate + caption presence, + optional alt-text coverage)
```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/verify.py" <work-dir> [--descr]
```
Confirms (a) every expected caption text box landed, and (b) no caption overlaps
a defect class. Checks all six placement patterns in 2D:
**T** text-overlap (title/body/subtitle/text boxes) · **B** footer intrusion ·
**C** caption-inside-picture (SmartArt-icon captions exempt) · **E** caption–caption
overlap · **O** caption text overflow · **P** a band caption burying a
small/structural picture. With `--descr`, a third coverage gate also confirms every
picture is either given a `descr` value or marked decorative. **Exit codes:** `0`
clean · `4` a placement defect (T/B/C/E/O/P) · `5` a missing caption (text-miss) ·
`6` (`--descr`) a picture with no descr and no decorative mark · `2` usage/IO. The
TOTAL line prints every count so nothing hides behind the code. This is the
CI-friendly gate — a nonzero exit means do not ship until investigated.

### Step 6.5 — Caption accuracy spot-check (recommended before any external handoff)
`verify.py` validates presence + placement, **not whether a caption correctly
describes its image**. Captions are vision-generated once and reused by hash. Before
shipping to a faculty/external audience, spot-check a sample: re-open a handful of
images and confirm the caption matches. A confidently-wrong caption (e.g. a
data-viz labeled "hot air balloon") is the highest-visibility, otherwise-undetectable
failure. See "Known limits."

### Step 7 — Report
Tell the user (numbers come from the per-run `coverage_report.csv` + the audit CSVs):
- **N captions placed** across M decks (`added` / `fallback-above` / `fallback-bottom` / `added-smartart-icon`)
- **N pictures with NO clean slot** — intentionally uncaptioned, flagged for human review (`overlay-fullbleed` / `flagged-no-slot`). **This is the real accessibility-coverage gap — report it explicitly**, not lumped into "skipped".
- N decorative skips (`skipped-decorative`) and N text-only-SmartArt skips (`skipped-smartart-text-only`) — NOT gaps (text already accessible)
- N no-caption / extract-failed / linked-image skips (`skipped-no-caption` / `skipped-image-extract-failed` / `skipped-linked-no-embed`)
- N deck-level errors (corrupt / password-protected, names listed)
- N captions PRESERVED as instructor hand-edits (re-runs with `--update-existing`)
- Output paths (captioned decks, audit CSVs, `coverage_report.csv`)

### Step 8 — Per-slide corrections (review loop, `fix_captions.py`)
After the deck is delivered, a human reviewing it will spot per-slide issues — *"slide 5's caption box is too big"*, *"slide 12's caption overlaps the photo, move it below"*, *"slide 3's caption is wrong"*. `fix_captions.py` applies those fixes **in place, by slide number**, editing ONLY the named caption shapes — no re-extract, no re-read of images, no re-caption of the rest of the deck. It backs up the deck to `<deck>.pptx.bak.pptx` before the first write.

```bash
python3 "$CLAUDE_PROJECT_DIR/.claude/skills/captioner/scripts/fix_captions.py" <deck_captioned.pptx> \
  --slide 5  --shrink-box \
  --slide 12 --move below \
  --slide 3  --font-size 8 --font-color 222222 \
  --slide 7  --text "Photo: solar farm at dusk" \
  --verify
```

Each `--slide N` opens a fresh correction group; its flags reset at the next `--slide`, so the groups are independent. The four families of fix:

| Fix | Flags |
|---|---|
| **Resize** (box too big/small) | `--shrink-box` (fit to text) · `--height EMU` / `--height-in INCHES` · `--width EMU` / `--width-in INCHES` |
| **Reposition** (wrong spot / overlap) | `--move {below\|above\|left\|right\|bottom}` (re-placed cleanly next to its own picture) · `--nudge-x EMU` / `--nudge-y EMU` (relative; +y = down) |
| **Restyle** (font too big / wrong color) | `--font-size PT` · `--font-color HEX` · `--italic {true\|false}` · `--font-name NAME` · `--bg-color HEX` (`""`=transparent) · `--border-color HEX` (`""`=off) |
| **Retext** (wrong/bad caption) | `--text "new wording"` |

Target selection when a slide has more than one caption: `--which N` (1-based, top-to-bottom) or `--hash HEX` (the image hash in the caption's shape name). Default = every caption on the slide.

**Safety:** every geometry change (resize/move/nudge) is re-checked against the SAME 2D obstacle model the placer and `verify.py` use (title/body text, other pictures, other captions, footer, off-slide). A fix that would CREATE a defect is **refused and reported** — pass `--force` to apply it anyway, or `--dry-run` to preview every group without writing. `--verify` re-scans the saved deck and exits 4 if any caption placement defect remains. **Exit codes:** `0` clean · `4` a fix was refused (would create/leave a defect) or `--verify` found a defect · `2` usage/IO/no-matching-caption. Retext/restyle-only changes don't move the box, so they're never refused on geometry.

### The iteration loop — from "this caption looks wrong" to a clean fix
The tool is driven from plain-language review notes. The reviewer reads the delivered deck,
writes each problem **by slide number**, and you translate the note into one `--slide` group.
Always work the loop the same way:

1. **Preview, don't commit blind.** Run with `--dry-run` first — it prints, per group, exactly
   what would change and (for geometry changes) whether the new rect would create a defect.
2. **Apply.** Drop `--dry-run`; the deck is backed up to `<deck>.pptx.bak.pptx` and saved.
3. **Re-verify.** Add `--verify` (or run `verify.py`) so the saved deck is re-scanned before handoff.
4. **If a fix is refused** (exit 4: would overlap text/photo/caption), the tool says which defect.
   Adjust the note — switch `--move below` to `--move right`, or pair `--shrink-box` with the move
   so the smaller box fits — and re-run. Use `--force` only after you have LOOKED at the slide.
5. **Iterate until `--verify` is clean.** Multiple notes go in ONE command as multiple `--slide`
   groups; re-running is idempotent (edits the same named shapes in place, never duplicates).

**Common review note → fix:**

| What the reviewer says | Likely cause | Fix |
|---|---|---|
| "slide 5's caption box is too big / too tall" | nominal height larger than the text needs | `--slide 5 --shrink-box` |
| "slide 5's caption is too wide / runs off the edge" | wide nominal box | `--slide 5 --width-in 2.5` then `--verify` |
| "slide 12's caption overlaps the photo / a number on it" | image-baked text the placer can't see (geometry is blind to text *inside* a picture, e.g. the Module 12 slide 34 case) | `--slide 12 --move below` (or `right`/`above`); if all sides crowded, `--shrink-box` then `--move` |
| "slide 8's caption sits on the title / body text" | a layout the placer mis-measured (tall multi-line title) | `--slide 8 --move below`, or `--nudge-y 200000` to push down |
| "slide 3's caption is wrong / describes the wrong thing" | a confidently-wrong vision caption (highest-visibility, otherwise-silent failure) | `--slide 3 --text "Photo: <correct subject>"` |
| "caption text too small / too faint" | font size or color | `--slide N --font-size 11 --font-color 000000` |
| "the white box behind the caption looks bad here" | default white card clashes with background | `--slide N --bg-color "" --border-color ""` (100% transparent — fill AND border become `noFill`) |
| "two captions on one slide overlap each other" | two pictures close together | move one: `--slide N --which 2 --move right` |
| several pictures, you mean a specific one | ambiguous target | add `--which N` (top-to-bottom) or `--hash <hash from audit CSV>` |

**Known limits the loop cannot fully solve (escalate to a manual slide edit):**
- **Text rasterized inside a picture** is invisible to the geometry model, so an auto-placed
  caption *can* land on it. `--move` to a clear side; if full-bleed with no clear side, the
  slide needs a manual layout change (genuine coverage gap, not hidden).
- **Caption *correctness*** is never auto-verified — `--verify` checks placement only. Spot-check
  wording (Step 6.5) and fix wrong ones with `--text`.
- A picture with **no clean slot** (`flagged-no-slot` / `overlay-fullbleed` in the audit CSV) has
  no caption to edit — `fix_captions.py` adjusts existing captions, it does not create one for a
  skipped picture. That needs a manual layout change first.

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
| `--update-existing` | off | Strip prior captioner shapes before adding new ones (idempotent re-run). Without it, a deck that already holds captioner shapes is refused (exit 4) instead of being captioned twice. **Edit-aware**: a caption whose text no longer matches the prior audit CSV is treated as an instructor hand-edit and PRESERVED; backs up the prior output to `<deck>_captioned.pptx.bak.pptx` |
| `--force-overwrite-edits` | off | With `--update-existing`, remove even captions that look instructor-edited |
| `--no-smartart` | off | Disable SmartArt icon captioning (on by default; auto-extracts icon names from SVG metadata) |
| `--bg-repeat-threshold` | 4 | An identical image hash appearing as a structural picture on ≥ this many slides is treated as a repeated background/template and skipped (`skipped-decorative-background`), even if vision captioned it. Set `0` to disable. |
| `--quiet` | off | Suppress per-deck progress |
| `--quick` | off | **Captioning ONLY** — skip ALL QC (spell-check + date/template scan) |
| `--no-spellcheck` | off | Disable just the spell-check pass (date/template QC still runs) |
| `--no-dateqc` | off | Disable just the date/doubled-word/leftover-template scan |
| `--spellcheck` | (no-op) | Back-compat alias — spell-check is now **default-on**; flag accepted, does nothing |

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

## What this skill does NOT do
- Write alt-text unless asked — the OOXML `descr` field is authored only with `--write-descr` or `--descr-only`, and alt-text a deck already has is never overwritten.
- Auto-correct spelling — the default QC only *reports* suspects (and web-verifies names before suggesting); it never rewrites slide or caption text.
- Read off chart values — captions identify "this is a chart of X", they don't read figures.
- Modify the picture itself, slide master, or existing slide shapes (other than captioner-prior shapes via `--update-existing`).
- Touch embedded charts (preserved on save; no caption added since they're not Pictures).

## SmartArt handling (added 2026-05-12)

PowerPoint SmartArt diagrams are NOT `MSO_SHAPE_TYPE.PICTURE` shapes — they are `graphicFrame` elements with a `drawingml/2006/diagram` URI. The captioner has explicit handling for them:

1. **Text-only SmartArts** (no embedded icons — just a layout with text labels): **no caption added**. The visible text already lives in the slide's accessible text layer, so an alt-caption would be redundant clutter.
2. **SmartArts with embedded icons** (PowerPoint "Insert > Icons" inserts SVG icons into SmartArt cells): each icon gets its **own short caption directly under it**, positioned via the icon's geometry from `ppt/diagrams/drawing*.xml`. No combined diagram caption is added. **Icon-proportional sizing (v0.2.5):** the caption font scales with the icon (5pt for tiny icons < ~0.44", 6pt < ~0.71", 7pt < ~0.98", else 8pt). The box **hugs the text in both dimensions** — width has only a small floor and is capped at ~3× the icon width (never narrower than the longest word — it wraps to more lines instead of ballooning width), and height is sized to the wrapped line count (~1.25× line spacing + ~6000-EMU margins) so the card never extends below the text. The width/overflow math and `verify.py`'s overflow gate read the caption's actual font size so down-scaled captions are not false-flagged.

### How icon names are derived (no vision pass needed)
PowerPoint Icons store their human-readable name in the SVG `id` attribute on the root element, e.g. `<svg id="Icons_Checkmark">`, `<svg id="Icons_VideoCamera">`. The captioner walks the slide's `diagramDrawing` part rels, for each `<dsp:sp>` finds the `<asvg:svgBlip r:embed=rIdN>` extension (NOT `<a:blip>` — that points at the PNG fallback, which has no name), parses the SVG's `id`, and converts CamelCase → space-separated lowercase. So `Icons_VideoCamera` becomes the caption `video camera`.

### Caption style for icons
1-3 words, lowercase, no "icon" suffix. Placed in a small white card directly under the icon's bounding box.

### Audit actions for SmartArt
| Action | Meaning |
|---|---|
| `added-smartart-icon` | Per-icon caption placed under an icon inside a SmartArt frame |
| `skipped-smartart-text-only` | SmartArt with no embedded icons — intentionally not captioned |

### Toggling SmartArt captions on/off

SmartArt icon captioning is **on by default**; `--no-smartart` turns it off. The behaviour differs depending on whether the deck is already captioned:

| Goal | Command |
|---|---|
| Turn SmartArt captions **off** on already-captioned decks | `apply_captions.py <work_dir> --no-smartart --update-existing` |
| Turn them **back on** | `apply_captions.py <work_dir> --update-existing` (default = on) |
| Fresh deck, never caption SmartArt | `apply_captions.py <work_dir> --no-smartart` |

**Why `--update-existing` is required to turn them off:** `--no-smartart` alone only stops *adding* new SmartArt captions — it does not remove ones already in the deck. Removal of prior SmartArt caption shapes happens only on the `--update-existing` path. Without it, a re-run with `--no-smartart` leaves the existing SmartArt captions in place.

**Caveat — SmartArt removal is not edit-aware.** Unlike the main picture-caption path (which preserves instructor hand-edits via the audit-CSV oracle), SmartArt caption removal strips by shape-name prefix unconditionally, so `--update-existing` will also remove any hand-edited SmartArt caption.

## Audit CSV columns

`slide, pic_id, image_hash, caption, char_len, in_group_depth, action`

`action` values:
- `added` — caption inserted below picture
- `fallback-above` — caption placed above (no room below)
- `fallback-bottom` — caption placed at slide bottom, footer-safe (no room above or below)
- `overlay-fullbleed` — **SKIPPED**: caption would land >50% inside a picture (its own full-bleed image or another picture); no caption added — flagged for human review
- `flagged-no-slot` — **SKIPPED**: no candidate position clears every obstacle (title/body/footer/other captions/other pictures); no caption added — flagged
- `overflow-risk-noted` — informational note on a PLACED caption whose text may wrap beyond the nominal box (the box auto-sizes to fit; not a defect)
- `dry-run-would-{added|fallback-above|fallback-bottom|overlay-fullbleed|flagged-no-slot}` — what would happen in non-dry-run
- `skipped-decorative` — caption value was `[decorative]`
- `skipped-decorative-background` — **SKIPPED**: the same image hash appears on ≥ `--bg-repeat-threshold` slides, so it is a repeated background/template (e.g. a slide-wide texture). Per-image vision can mislabel these as informative; this deck-level signal overrides a non-decorative caption so the same background is not captioned on every slide. Vision-classified `[decorative]` images are left as `skipped-decorative`. **A repeated-background image is also excluded from the placement-obstacle set** (in both `apply_captions.py` and `verify.py`): a full-bleed background otherwise makes every candidate position look ">50% inside a picture" and forces real photos on that slide to `flagged-no-slot`. Captions may sit over a background; titles, body text, footers, and real (non-repeated) pictures still block them.
- `skipped-no-caption` — no entry in captions.json for this image
- `skipped-image-extract-failed` — could not read the image blob
- `skipped-linked-no-embed` — linked/external image (no embedded bytes to hash)
- `added-smartart-icon` — per-icon caption inserted under a SmartArt icon (auto-generated from SVG metadata)
- `skipped-smartart-text-only` — SmartArt frame with no embedded icons; intentionally not captioned

`in_group_depth` shows how deeply nested the picture was inside group shapes (0 = top level).

## Placement engine (v0.2.3 — 2D, text-aware, with bottom-of-picture fallback)

A caption is **never** placed where it covers text. The placer treats as **2D hard
obstacles**: EVERY text frame on the slide — title, body, AND plain text boxes /
auto-shapes (the question/answer/label text that lives outside placeholders, which
v0.2.2 was blind to) — plus every *other* picture and every already-placed caption.
A candidate is rejected only when it overlaps an obstacle in **both** axes
(`_geometry.clear_all_obstacles_2d`), so a caption beside a narrow box is allowed but
one that would sit on it is not.

Placement preference for each picture:
1. **Below** the picture (clean), then **above** it (clean) — each tested against all
   obstacles *including the picture's own* box, so a caption whose auto-sized height
   would dip into its own picture falls through; a near-zero tolerance against other
   captions keeps two cards from touching; a horizontal nudge is tried before giving up.
2. **Bottom-of-picture band** — if no clean external slot exists, the caption is placed
   in the bottom strip *inside its own picture* (constrained to the picture's width so
   it never spills into an adjacent picture; it grows taller instead). Covering a little
   of the image beats covering text. Recorded as action `inside-bottom`, shape-name
   prefix `captioner_capband_`. **Band captions render small (`BAND_FONT_PT`, 6pt)** so
   the strip is thin (overlaps as little of the photo as possible) and fits tight spots a
   full-size caption can't — maximizing coverage while staying unobtrusive.
3. **Skip + flag** (`flagged-no-slot`) only if even the small band would cover text.

**Text-obstacle sizing (v0.2.6):** the region a title/text box occupies is estimated
font-aware and line-break-aware (`_visible_text_rect`) — it counts explicit `\n`/`\x0b`
breaks and scales line height by the actual font size, so a 2-line or large-font title is
no longer under-measured as one line (which previously let a caption land on the title's
lower lines). **Limitation:** text rasterized *inside* a picture image has no text frame,
so the band can still overlap image-baked text — undetectable by geometry.

Caption box height is **estimated from the text + width** (auto-size grows the box, so
a 2-line caption is taller than nominal) and used for every placement/overlap/footer
decision — so a grown caption never spills onto a neighbor. SmartArt-icon captions get
the same height estimate + mutual-collision nudge. The "own picture" is identified by
**geometry, not pic_id** (`iter_slide_pics` can assign the same pic_id to multiple
pictures, which previously let a caption land on a *different* picture).

The gate patterns are **T** (text overlap), **B**
(footer), **C** (in-picture, band/icon exempt), **E** (caption–caption), plus missing-caption.

### Two more defect classes (v0.2.4) — overflow + structural-picture coverage
- **O — text overflow.** A single word cannot wrap, so a caption box narrower than its
  longest word overflows horizontally (e.g. a SmartArt-icon caption "workflow" /
  "electrician" forced into a ~0.44in box). Fix: every caption/icon box is sized to fit
  its longest word with headroom (`required_caption_width`), re-centered on the icon.
  Detection: `caption_overflows` is a **blocking** gate pattern (verify.py `O`).
- **P — covers picture content.** A bottom-band caption is fine on a big photo (thin
  strip) but buries content on a **small / thin / tall-narrow structural** picture
  (numbered chevron, icon, banner). Fix: placement now tries clean **right-of-picture →
  left-of-picture** slots before the band, and **refuses to band an at-risk picture**
  (`band_covers_structural_picture`) — it skips + flags instead of covering it.
  Detection is a **blocking** gate pattern (verify.py `P`). Residual: a number/arrow
  rendered *inside* an image is invisible to the obstacle model; the size/coverage
  heuristic + side-first placement is the mitigation, and unplaceable ones are flagged.

### Pre-write invariant (v0.2.4 — the standing guarantee)
Immediately before writing each caption, `apply_captions` re-verifies the FINAL clamped
rect against every defect class (off-slide, footer, text-overlap, caption-overlap,
>50%-in-picture). If any upstream step — the clamp, the height estimate, a new/unseen
layout, group-nested coordinate drift — produced a defect, the caption is **refused and
skipped** (`flagged-self-check-<class>`) instead of written. A placement defect is therefore
**impossible to write by construction**; `verify.py` and the auditor are defense-in-depth,
not the only line. In testing the invariant fires 0 times on clean input (it rejects no legitimate
caption). **Honest guarantee scope:** every written caption is clear of the checked defect
classes *at the moment it is written*; this does not cover post-delivery edits, post-write
auto-size growth beyond the height estimate, or guarantee that a group-nested (depth≥1)
picture's caption is visually adjacent to its picture (only that it is defect-free) — those
cases are surfaced as skips in `coverage_report.csv` / the audit CSV, never shipped silently.

## Known limits

- **Caption *correctness* is not verified.** Captions are vision-generated once and
  reused by hash; `verify.py` checks presence + placement only. Spot-check accuracy
  before any external handoff (Step 6.5). A wrong-but-plausible caption is the
  highest-visibility, otherwise-silent failure mode.
- **WCAG alt-text is opt-in, and one attribute only.** With `--write-descr` (or
  `--descr-only`) captioner authors the OOXML `descr` field for pictures that lack it;
  it does not certify WCAG 2.1 conformance and does not touch charts/embedded objects/
  non-picture shapes. Pair it with a full alt-text review rather than treat it as a
  replacement.
- **Pictures with no clean slot are left uncaptioned** (full-bleed images, crowded
  slides) — surfaced as `overlay-fullbleed` / `flagged-no-slot`, not silently dropped.
  If 100% picture coverage is required, those slides need manual layout changes.
- **`OVERFLOW-RISK` is heuristic**, based on a fixed chars-per-line estimate; the
  caption box auto-sizes, so these are informational, not confirmed visual overflow.
- **`.ppt` legacy format is not supported** (python-pptx handles `.pptx` only).
- **Group-nested picture geometry** is read from the picture's own `xfrm`; deeply
  group-transformed coordinates are a known approximation.

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
| Picture fills entire slide | No clean slot → caption SKIPPED + flagged `overlay-fullbleed` (never placed on top of the image) |
| Re-run on `_captioned.pptx` output | Without `--update-existing`: **refused** (exit 4; no deck or audit written for it), so captions are never duplicated. With: strips prior captioner shapes (ordinary and in-picture band captions) and re-adds; hand-edited captions are kept. |
| Re-run on source with prior captioner shapes | Same as above — captioner shapes carry the name prefixes `captioner_caption_`, `captioner_capband_`, `captioner_smartart_` and `captioner_sa_icon_` |
| Multi-language slide text | Context extraction works; captions still emitted in English |
| Animated picture / video | A video's poster frame is extracted and captioned like a picture (measured 2026-10-19 with a python-pptx `add_movie` deck); the audit row does not say it is a video. An animated GIF is captioned as a picture. |

## Provenance
Built 2026-05-10 for lecture-deck accessibility remediation (compliance image-tagging).

Hardened 2026-05-10 (later same day) following an independent review: added recursive group traversal, dry-run, error handling, idempotency marker, style flags, deck context, hidden-shape skip, progress prints.

Extended 2026-05-12 during a production run after concrete user feedback on a captioned-deck output:
- **White-card background fill** for caption text boxes (`--bg-color FFFFFF` + `--border-color CCCCCC`, both default-on). Root cause: default `--font-color 333333` was unreadable on dark-background slides (e.g., section dividers). Card makes captions legible on any slide background.
- **SmartArt icon captioning**: each PowerPoint Icon embedded inside a SmartArt now gets its own short caption directly under it, with the icon's name pulled from SVG `id` metadata (no vision pass). Text-only SmartArts intentionally skipped (no caption clutter when text is already accessible).
- **Context-hygiene rule** (workflow, not code): finish one deck end-to-end before reading the next deck's images; for 5+ deck batches, dispatch via background subagents so image reads stay scoped to each agent's context window. Root cause: a 32-deck batch run mid-session triggered a context compaction because all image reads were accumulating in the main conversation.
