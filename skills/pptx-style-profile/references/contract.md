# contract.json reference

`profile.py` writes `contract.json` (schema `pptx-style-contract/1`) and renders
`STYLE.md` from it. Read the contract when you need exact keys; read STYLE.md for the
human overview. Never edit either by hand: re-profile instead.

**Contents:** [Top level](#top-level) · [layouts](#layouts) · [patterns](#patterns) ·
[derived](#derived) · [ai_readiness](#ai_readiness) · [Diff classes](#diff-classes)

## Top level

| Key | Meaning |
| --- | --- |
| `source` | file name, `path`, `kind` (`pptx`/`potx`), whole-file `sha256`, per-part sha256 of presentation, masters, theme and layouts |
| `slide_size` | `[cx, cy]` in EMU (914400 per inch, 12700 per pt) |
| `theme.colors` | the 12 scheme slots as uppercase hex |
| `theme.fonts` | `major` (headings) and `minor` (body) Latin typefaces |
| `theme.text_styles` | master title/body/other sizes per level, pt |
| `masters[].chrome` | static shapes on the master: logo pictures, slide-number fields, footers, rules |
| `layouts[]` | see below |
| `patterns[]` | sample slides worth cloning, see below |

Boxes are `[x, y, width, height]` in EMU, absolute on the slide (group transforms
already applied).

## layouts

| Key | Meaning |
| --- | --- |
| `name` | what a build spec names in `"layout"` |
| `type` | PowerPoint layout type (`title`, `obj`, `secHead`, …) when present |
| `placeholders[]` | `key` (fill key: `title`, `subtitle`, `body`, `body2`, `pic`, … in reading order), `type`, `idx`, `box`, `prompt`, `prompt_looks_like_content`, `font_size`, `font`, `spacing` `[pt before each paragraph, line-spacing factor]`, `capacity` `[lines, chars per line]` (a planning estimate at 0.5 em; the build skill's fit check with the real font decides) |
| `signature` | placeholder count per role; `collision_group` lists layouts with the same signature (an agent cannot tell them apart by structure: choose by name). Same signature is not same capacity |
| `chrome` | static non-placeholder shapes on the layout (`where: layout`) |
| `art` | `background` (`image`/`solid`/`none`) and picture count: art on the layout travels with every slide that uses it |
| `has_sample_slide` | whether any sample slide demonstrates the layout |
| `show_master_shapes` | `false` when the layout hides master chrome |

## patterns

One entry per slide in the template file.

| Key | Meaning |
| --- | --- |
| `id` | stable slug from the slide title (`timeline`, `three-ways-we-help`) |
| `slide`, `part`, `layout` | where it lives in the template |
| `slots` | fill keys → `kind` (`text`, `bullets`, `table`, `image`, `chart`, `bar`, `mark`), index `path` into the shape tree, original `text`, `box`, `capacity` |
| `repeats` | repeat groups (`g1`, …): sibling groups with the same structure (cards, KPI tiles); fill as a list of at most `count` items, each an object of the inner slot tags; unused members are deleted whole |
| `dependencies` | images (shared by clones), charts and embeddings (copied per clone) |
| `placeholders_used` | `false` for decorative slides built from text boxes |
| `asset_bank` | `true` when every slot is a picture (icon or photo banks): reuse its images, do not clone it as a content slide |
| `hidden` | the sample slide is hidden in the template (often instructions) |
| `shape_boxes` | top-level shape boxes, used by lint to tell template geometry from free shapes |

Placeholder slots on a pattern reuse the layout's key (`title`, `body3`); other text
shapes get `t1`, `t2`, … in reading order. `bar` slots are text-less auto-shapes over a
table (timeline bars, filled with `{"row", "from", "to"}`); `mark` slots are small ones
(rating dots, kept or deleted).

## derived

Inferred from the sample slides, each with evidence:
`colors_used` (hex → slide numbers), `off_theme_colors`, `fonts_used`, `type_scale`
(sizes per role), `min_font_size` (smallest text in the samples, the lint floor),
`safe_area`, `footer_band`, `footer_placeholders_on_slides`, `contrast` (every theme
colour on lt1/lt2/dk1/dk2 with AA pass/fail).

## ai_readiness

`{code, severity, where, message, fix}` per finding; the codes are listed in
`ai-readiness.md`.

## Diff classes

`diff.py OLD NEW` (either side a contract or a template):

| Class | Examples | Exit |
| --- | --- | --- |
| breaking | layout removed or renamed, placeholder removed or retyped, placeholder moved/resized beyond tolerance (default 0.25 % of slide width, min 1 pt), pattern removed or renamed, slot removed or changed kind, repeat group removed or holding fewer items, slide size changed | 1 |
| changed | theme colour or font value, chrome geometry, a pattern's layout | 0 |
| additive | new layouts, placeholders, patterns | 0 |

A layout rename is detected when a removed and an added layout have identical
placeholder geometry; a pattern rename when a removed and an added pattern have the
same slots and repeat groups (its sample slide's title changed).
