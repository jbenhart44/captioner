# Changelog

All notable changes to captioner are documented here. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.3] - 2026-05-29

Minor release. Text-aware placement: captions are guaranteed never to cover text.

### Added

- **`TEXT-OVERLAP` is now a first-class placement check** in both the auditor and
  `verify.py` (which gains a five-pattern gate: text-overlap, footer, in-picture,
  caption-caption, missing-caption, with distinct non-zero exit codes).
- **Bottom-of-picture band fallback** (`inside-bottom`): when no slot outside the
  picture clears the surrounding text, the caption is placed in the picture's own
  bottom strip (constrained to the picture width) rather than over text or skipped.

### Fixed

- **Captions no longer cover text in plain text boxes / auto-shapes.** The obstacle
  model previously saw only placeholder text (title/body); it now treats *every*
  text frame as a 2D obstacle. Obstacles are narrowed to each shape's estimated
  visible-text region (anchor-aware) so a caption below a tall title's text is not
  falsely blocked by the title's empty box.
- **Auto-sized caption height is accounted for**, so a multi-line caption no longer
  grows onto a neighbouring caption or label.
- **The own picture is identified by geometry, not `pic_id`.** `iter_slide_pics`
  can assign the same `pic_id` to multiple pictures on a slide, which previously let
  a caption land inside a *different* picture stacked nearby.

### Verification

- Across a 49-deck production corpus: 0 text-overlap, 0 caption-in-picture, 0
  caption-caption, 0 off-slide, 0 footer defects. Confirmed by an independent
  adversarial cross-check using separate overlap math.

## [0.2.2] - 2026-05-19

Patch release. One placement improvement on top of v0.2.1.

### Fixed

- **Caption for picture-N no longer lands inside picture-M on multi-picture
  slides.** Every other picture's vertical band on the slide is now part of
  the caption placement obstacle list (the caption's own picture is filtered
  out by `pic_id` so it can still anchor adjacent). On a 49-deck production
  corpus, `FULLBLEED-OVERLAY` audit rows dropped 31 → 19 with this change.

### Behaviour notes

- The caption count placed per deck may drop slightly compared to v0.2.1 when
  a slide has two large pictures whose vertical bands cover most of the slide
  — captions that previously landed inside the neighbouring picture are now
  correctly skipped, surfaced as `overlay-fullbleed` or `flagged-no-slot`
  audit rows for human review.

## [0.2.1] - 2026-05-19

Placement-quality release. Five real-deck failure patterns surfaced by user review of v0.2.0 output are now prevented in placement and surfaced explicitly in the audit. SmartArt-icon caption placement inherits the same guarantees as the main caption path.

### Fixed

- **Body / subtitle overlap (Pattern A)** — caption placement now considers BODY, SUBTITLE, CONTENT, OBJECT placeholders with visible text as obstacles, not just TITLE/CENTER_TITLE. The Fix-A obstacle filter is unified with the title check via `_clear_all_obstacles(c_top, c_h, obstacles)` so future obstacle classes plug in cleanly.
- **Footer abutment (Pattern B)** — new `FOOTER_CLEARANCE_EMU` constant (default 91 440 EMU ≈ 0.10 in) reserves a visible gap between the caption bottom and the master/layout footer band. Configurable via the new `--footer-clearance-emu` CLI flag. The constant is applied to all four uses of `footer_limit` in the placement block (the v0.2.0 patch missed one site, which caused caption-top regression on full-bleed pictures).
- **Full-bleed picture overlay (Pattern C)** — replaces the area-ratio heuristic with deterministic no-clean-slot detection: if the picture leaves no room above OR in the clean-below band, the caption is **skipped** and a `overlay-fullbleed` audit row is emitted with the picture's slide-clamped visible-coverage for human review. No magic threshold; no silent overlay on the picture body.
- **Caption text overflow on narrow pictures (Pattern D)** — caption box now widens (up to slide width) to fit the caption text in ≤2 lines. Never truncates the caption (truncation would diverge the displayed text from the accessibility metadata). Widening events surface as informational audit rows.
- **Caption–caption overlap on multi-picture slides (Pattern E)** — placement tracks previously placed captions on the same slide and rejects candidate positions that would collide. When no clean slot exists, the second caption is **skipped** with an explicit `flagged-no-slot` audit row (never a silent fallback). Includes a horizontal-nudge attempt when only a single nearby caption is in the way.
- **SmartArt-icon caption parity** — the per-icon placement path now inherits Fix-B footer clearance + Fix-D widening + caption-caption obstacle registration, so SmartArt icon captions no longer abut the footer or overflow narrow boxes.

### Added

- `scripts/_geometry.py` — shared geometry helpers (`slide_footer_top`, `slide_title_rect`, `slide_body_obstacle_bands`, `_clear_all_obstacles`, `visible_coverage`, `FOOTER_CLEARANCE_EMU`, `MIN_CAPTION_HEIGHT`, `MIN_CAPTION_WIDTH_EMU`, `EMU_PER_CHAR_DEFAULT`) so apply / verify (and an external auditor, if you have one) agree on band thresholds and clearance constants.
- `verify.py` extended with Pattern B (footer-abutment) and Pattern C (caption-inside-picture) coordinate checks. Audit rows with `action='overlay-fullbleed'` or `flagged-no-slot` are treated as known-skipped (intentionally no caption placed), not text-mismatch failures.

### Changed

- The audit-row CSV writer now unions the field names across all rows (not just the first), so issue-specific columns (e.g. `visible_coverage` on `overlay-fullbleed` rows) round-trip cleanly.
- `--height-emu` is now an *advisory* initial height; Fix-D may widen the box and `MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT` grows it vertically as needed.

### Known limits

- The `EMU_PER_CHAR_DEFAULT = 114 000` heuristic for caption-text overflow detection is conservative for 10 pt italic Calibri but unvalidated across other fonts. Override via `--chars-per-emu` if your captions use a wider face. Long unbreakable tokens (URLs, compound proper nouns) can still overflow even with widening.
- Vertical-only obstacle checking is the v0.2.1 contract — a narrow centered subtitle plus a right-side caption may trigger false-positive avoidance. 2D box-intersection is on the roadmap for a future release.

## [0.2.0] - 2026-05-19

Correctness + QC release. Caption output for decks **without** placeholder-
hosted or group-nested pictures is unchanged; decks that use picture
content-placeholders now get captions they previously missed.

### Fixed
- **Picture enumeration** moved off python-pptx `shape_type == PICTURE` to a
  raw-OOXML `<p:pic>` walk (`scripts/_oxml_pics.py`). python-pptx types a
  picture content-placeholder (`<p:pic>` with `<p:ph type="pic"/>`) as
  `PLACEHOLDER`, not `PICTURE`, so the old walk silently skipped every
  placeholder-hosted picture (observed on a real deck where only ~8% of
  pictures were seen). The new walk also captures group-nested pictures, and
  extract/apply/verify now enumerate an identical population. `<mc:Choice>`
  branches are de-duplicated; linked images without `r:embed` are skipped and
  logged. The SHA-256 image hash and `captions.json` key format are
  byte-unchanged — existing reuse indexes still match.
- **Caption placement**: placeholder-inherited geometry is now resolved (was
  landing off-slide at a degenerate position); captions are clamped fully
  on-slide and kept clear of footer and title placeholders wherever a valid
  slot exists.

### Added
- **Default-on QC**: a whitelist-aware spell-check plus a date / doubled-word
  / leftover-template scan now run on every apply (flag-only — never edits a
  `.pptx`). New `qc/<deck>_spellcheck.csv` and `qc/<deck>_qc.csv`.
- `--quick` (captioning only, skip all QC), `--no-spellcheck`, `--no-dateqc`.
- Suspected proper-noun misspellings are tagged `verify_name` so the spelling
  can be web-verified before being presented as a correction.
- Plural-of-acronym suppression (e.g. `NPVs`, `KPIs`) + expanded whitelist.

### Changed
- QC artifacts live in their own `qc/` directory; dry-run writes
  `audit/<deck>_audit_dryrun.csv` so it can never overwrite a real audit.
- Missing `pyspellchecker` is now a loud failure (banner, marker file, exit
  code 3) instead of a silent skip, because QC is default-on. `pyspellchecker`
  added to `requirements.txt`.
- `--spellcheck` is retained as an accepted no-op (QC is default-on).

## [0.1.1] - 2026-05-15

Documentation-only release. No code or behavior changes; captioning output
is byte-identical to 0.1.0.

### Changed
- README and project landing page now describe only shipped capabilities.
  The speculative "Roadmap" / "Extending captioner" sections (forward-looking
  feature ideas, including a prospective `--descr` mode and PyPI packaging)
  were removed so every documentation surface conveys one consistent message.

## [0.1.0] - 2026-05-14

First public release. Built and hardened across a production run of 1,132
captions over 32 PowerPoint decks in three graduate engineering courses.

### Added
- Core workflow: extract pictures + SmartArt icons from `.pptx`, generate
  brief subject-identifying captions, apply as visible italic text boxes,
  verify placement, emit a per-deck audit CSV.
- Recursive group-shape traversal — pictures nested inside group shapes are
  captioned.
- `--dry-run` mode — emits the audit CSV without modifying any `.pptx`.
- Idempotent re-runs — captioner shapes carry a `captioner_caption_<hash>`
  name prefix; `--update-existing` strips prior captions before re-applying.
- Configurable styling flags — font name/size/color, white-card background
  fill + border, gap and box-height in EMU.
- Deck-level context flag (`--context`) to bias caption vocabulary.
- White-card background fill (default-on) so captions stay legible on dark
  slide backgrounds, including section dividers.
- SmartArt per-icon captioning — each embedded PowerPoint Icon gets its own
  short caption derived deterministically from the icon's SVG `id` metadata
  (no vision pass). Text-only SmartArts are intentionally skipped.
- Positioning fallbacks — caption goes below the picture by default; falls
  back to above, then to slide-bottom, with the chosen fallback recorded in
  the audit CSV.
- Error handling for corrupt / password-protected decks (logged in the
  manifest, not crashed on).
- Context-hygiene workflow guidance for large multi-deck batches.

### Known limitations
- Does not write the OOXML `descr` alt-text field (the WCAG-canonical,
  screen-reader-consumed attribute). Captioner produces a visible
  pedagogical-clarity layer; pair it with a full alt-text remediation
  workflow for WCAG 2.1 AA conformance.
- `.ppt` legacy format is not supported (python-pptx handles `.pptx` only).

[0.2.2]: https://github.com/jbenhart44/captioner/releases/tag/v0.2.2
[0.2.1]: https://github.com/jbenhart44/captioner/releases/tag/v0.2.1
[0.2.0]: https://github.com/jbenhart44/captioner/releases/tag/v0.2.0
[0.1.1]: https://github.com/jbenhart44/captioner/releases/tag/v0.1.1
[0.1.0]: https://github.com/jbenhart44/captioner/releases/tag/v0.1.0
