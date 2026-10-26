# QA: fit, lint, render, review

**Contents:** [Fit](#fit) · [Lint codes](#lint-codes) · [Render](#render) ·
[Review rubric](#review-rubric) · [Package facts](#package-facts)

## Fit

`fit.py` measures text with the theme font's real advance widths (TrueType/OpenType
`cmap` + `hmtx` + `head.unitsPerEm`, read with the stdlib), wraps greedily word by
word, and compares lines needed with lines available: box minus insets and the space
before each paragraph, 1.2 × size × the paragraph's line spacing per line. Kerning,
ligatures and complex-script shaping are ignored, so widths get a 3 % margin and the
result stays an estimate even with the right font. Font lookup:
exact family in the OS font folders and `$PPTX_STYLE_FONT_DIRS`, else a
metric-compatible twin (Carlito ↔ Calibri, Caladea ↔ Cambria, Liberation Sans ↔
Arial, Liberation Serif ↔ Times New Roman, Liberation Mono ↔ Courier New, Gelasio ↔
Georgia), else an average 0.55 em glyph and the result is labelled approximate. Point
`PPTX_STYLE_FONT_DIRS` at a folder of licensed font files to measure (and render) with
a font that is not installed.

Policy, in order: shorten the wording → move detail into the speaker notes →
`allow_split` (continuation slide with the same title + "(cont.)") → choose a pattern
with more room. Never shrink text below the template's sizes; overflow is an error that
names slide, shape and capacity.

Sources: [OpenType hmtx](https://learn.microsoft.com/en-us/typography/opentype/spec/hmtx),
[hhea](https://learn.microsoft.com/en-us/typography/opentype/spec/hhea),
[Metric-compatible fonts](https://wiki.archlinux.org/title/Metric-compatible_fonts).

## Lint codes

Errors (exit 1):

| Code | Fix |
| --- | --- |
| `text-overflow` | shorten, move to notes, split |
| `missing-title` | fill the title placeholder |
| `missing-notes` | add `notes` (exempt: slide 1 and slides on layouts of type `title` or `secHead`) |
| `prompt-only-content` | fill the placeholder whose prompt looks like content (`01`); new specs are rejected earlier, this catches decks edited by hand |
| `leftover-sample-text`, `placeholder-text` | fill or `null` the slot; never ship lorem/TODO/"Click to add" |
| `off-palette-colour`, `off-theme-font`, `text-too-small` | remove hand formatting; styling comes from the template |
| `covers-chrome`, `off-slide` | free shapes may not sit on logo/page number/footer or off the slide; a table grown into the footer band is flagged too |
| `off-contract-layout` | use a contract layout |
| `pkg-*` | package defects (content types, dangling rels, orphans, ids, sections, template content type): a builder bug, report it |
| `duplicate-shape-id` | builder bug, report it |

Warnings: `template-hardcoded-colour` (the template's own off-theme colour),
`no-title-placeholder` (the sample slide keeps its title in a text box),
`empty-placeholder`, `low-contrast`, and `text-overflow` when the font is missing and
the overflow is small.

## Render

`render.py DECK --contract C` converts with LibreOffice in an isolated profile
(`-env:UserInstallation`, so runs never share state), rasterises with `pdftoppm`, and
tiles a `contact-sheet.png`. Without LibreOffice it prints `SKIPPED` and exits 0
(`--require` makes it exit 3): lint still stands, but say the visual check is owed.
LibreOffice substitutes missing fonts and wraps differently from PowerPoint, so the
render is labelled **APPROXIMATE** whenever a theme font is not installed; judge layout
and overlap, not exact line breaks. Fonts in `PPTX_STYLE_FONT_DIRS` are loaded into the
private LibreOffice profile, so the render uses them without a system install.
LibreOffice also applies `normAutofit` shrinking when it opens a file while PowerPoint
only does so after an edit: a render can look tidier than the deck will in PowerPoint,
which is why fit is checked on the file, not on the picture.

Sources: [LibreOffice start parameters](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html).

## Review rubric

Look at the contact sheet fresh (a subagent is ideal), slide by slide.

Hard (fix before delivery): text cut off or overlapping; content over the logo, page
number or footer; broken image; empty placeholder visible; leftover sample text;
missing title; a chart showing sample data.

Soft (fix if cheap): uneven card heights after edits, a removed repeat member leaving a
gap, a one-word last line, the same pattern three slides in a row.

Budget: at most two fix passes, then report what remains. Under-fixing beats
oscillating. Never change theme colours, fonts, layout geometry or chrome to make
content fit.

## Package facts

- Cloned slides get their own chart parts and embedded workbooks; two clones of a chart
  pattern never share data. Images are shared. Notes are created fresh, and a notes
  master is created when the template has none.
- Removed slides are dropped from the relationship graph and every unreachable part is
  deleted, so an icon-bank appendix does not bloat the output. Custom shows and
  sections that point at removed slides are pruned (create mode drops the template's
  sections).
- Deleting a shape from a clone removes that slide's animations (they may target it).
- Sources: [notes slides](https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-notes-slides),
  [python-pptx chart/xlsx parts](https://python-pptx.readthedocs.io/en/stable/dev/analysis/cht-access-xlsx.html),
  [slide deletion leaves parts (python-pptx #67)](https://github.com/scanny/python-pptx/issues/67).
