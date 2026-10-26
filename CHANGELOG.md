# Changelog

All notable changes to this repository's skills are documented here.
Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versioning: [SemVer](https://semver.org) on the plugin manifest
(breaking skill-interface change → major, new skill → minor, fix → patch).

## [Unreleased]

### Added
- `pptx-style-profile`: profiles a .pptx/.potx template into `contract.json` (layouts,
  placeholders with prompt-only detection, theme, master/layout chrome, sample-slide
  patterns with tagged slots and repeat groups, derived rules with evidence, per-part
  hashes) plus a generated `STYLE.md` and an AI-readiness audit; `diff.py` classifies
  template changes as breaking, changed or additive.
- `pptx-style-build`: builds or edits native decks inside a contract from a JSON spec:
  real layouts, cloned sample slides (charts and workbooks copied per clone), run-level
  text, speaker notes, preview gate, edit mode with byte-identical passthrough and
  explicit removals; fit checks with real font metrics, contract lint (also for decks
  made elsewhere), LibreOffice render with contact sheet.
- Generated brand-free test fixtures (`tests/fixtures/make_fixtures.py`) and the
  acceptance suite `tests/test_pptx_style.py`, run by `make test` and CI without
  LibreOffice; trigger-eval ratchet set to 85 % rank-1.
- `docs/setup-prompt.md`: a paste-ready `/goal` that runs profile → plan → preview →
  build → review with a gate per skill.
- README demo as code: `demo/demo.tape` (VHS) and `demo/render-sheet.sh` with their
  committed renders.
- Repository scaffolded from the skills template.
