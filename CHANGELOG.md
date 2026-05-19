# Changelog

All notable changes to captioner are documented here. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[0.2.0]: https://github.com/jbenhart44/captioner/releases/tag/v0.2.0
[0.1.1]: https://github.com/jbenhart44/captioner/releases/tag/v0.1.1
[0.1.0]: https://github.com/jbenhart44/captioner/releases/tag/v0.1.0
