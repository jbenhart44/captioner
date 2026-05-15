# Changelog

All notable changes to captioner are documented here. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[0.1.1]: https://github.com/jbenhart44/captioner/releases/tag/v0.1.1
[0.1.0]: https://github.com/jbenhart44/captioner/releases/tag/v0.1.0
