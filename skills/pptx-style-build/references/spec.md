# Build spec reference

A spec is JSON. Every key it may contain is listed here; anything else is rejected
(formatting keys with an "off-contract" message). Keys come from the contract:
layout names and placeholder keys from `layouts[]`, pattern ids, slot tags and repeat
groups from `patterns[]` (STYLE.md lists them).

**Contents:** [Top level](#top-level) · [Slides](#slide-entries) · [Fill values](#fill-values) ·
[Example](#example) · [Edit mode](#edit-mode) · [Exit codes](#exit-codes)

## Top level

| Key | Required | Meaning |
| --- | --- | --- |
| `contract` | yes (or `--contract`) | path to `contract.json`, relative to the spec file |
| `slides` | yes | ordered list of slide entries |
| `output` | no | output path relative to the spec (default `deck.pptx`); `.pptx` output from a `.potx` template is converted to a presentation |
| `template` | no | override the template path recorded in the contract |
| `base` | no | existing deck to edit (edit mode); must be built on the contract's template and differ from `output` |
| `remove` | edit mode | base slide numbers to delete; every base slide must be kept or removed |
| `require_notes` | no | default `true`: content slides without notes are lint errors |

## Slide entries

Exactly one of `layout`, `pattern`, `keep`:

| Key | Meaning |
| --- | --- |
| `layout` | layout name; `fill` keys are its placeholder keys (`title`, `subtitle`, `body`, `body2`, `pic`, …). Unfilled placeholders are omitted, except prompt-only ones (agenda `01`), which need real text: `null` or omission is rejected |
| `pattern` | pattern id; the sample slide is cloned (media shared, charts and workbooks copied, notes fresh) and its slots filled. Every slot that holds sample text must be filled or set to `null` |
| `keep` | 1-based slide of the `base` deck (edit mode only; each base slide at most once) |
| `fill` | object of fill key → value |
| `notes` | speaker notes: string, or list of paragraphs |
| `allow_split` | `true`: list fills that overflow move their tail to continuation slides titled "… (cont.)" |
| `replace` | `keep` slides only: `{"old text": "new text"}`; each old text must sit inside one run (else rejected, never flattened) |
| `comment`, `id` | ignored, for humans |

## Fill values

| Slot kind | Value |
| --- | --- |
| text / bullets / placeholder | `"one line"`, `"para 1\npara 2"`, or a list; items can be `{"text": "…", "level": 1}` for sub-bullets. Paragraph and run formatting of the template is kept |
| any slot | `null` deletes the shape (use it for optional decorative text) |
| repeat group `g1` | list of ≤ `count` objects keyed by the group's own inner tags: `[{"t1": "Hire", "t2": "Add a shift"}]`; unused members are deleted whole. Inner `t1` is not the top-level `t1` |
| table `tbl1` | list of rows (header row included); fewer rows deletes rows, more rows clones the last; never more columns than the sample |
| bar `bar1` | `{"row": 2, "from": 2, "to": 3}`: 1-based table row and first/last column; geometry comes from the table grid |
| mark `mark7` | `false`/`null` deletes it (rating dots) |
| chart `chart1` | `{"categories": ["Q1", "Q2"], "series": [{"name": "Revenue", "values": [3, 5]}]}`; cached values and the embedded workbook are both rewritten; series count cannot grow |
| image `img1`, `pic` | `{"image": "path.png", "alt": "what it shows"}`: PNG, JPEG or GIF, centre-cropped to the slot's aspect ratio (never stretched); alt defaults to the file name. For SVG, EMF or WebP ask for a PNG |

## Example

```json
{
  "contract": ".pptx-style/our-template/contract.json",
  "output": "q3-review.pptx",
  "slides": [
    {"layout": "Title Slide", "fill": {"title": "Q3 operations review", "subtitle": "Board, 14 Nov"}},
    {"layout": "Agenda", "fill": {"title": "Agenda", "body": "01", "body2": "Where we are", "body3": "02", "body4": "The plan"},
     "notes": "Two parts: status, then the decision we need."},
    {"pattern": "timeline", "fill": {
        "title": "Rollout over four quarters",
        "tbl1": [["Stream", "Q1", "Q2", "Q3", "Q4"], ["Pilot", "", "", "", ""], ["Scale", "", "", "", ""]],
        "bar1": {"row": 2, "from": 2, "to": 2}, "bar2": {"row": 3, "from": 3, "to": 4}, "bar3": null},
     "notes": "Pilot in Q1, scale over Q2-Q3."},
    {"layout": "Title and Content", "allow_split": true,
     "fill": {"title": "Risks", "body": ["Hiring lead time", {"text": "Mitigation: agency", "level": 1}]},
     "notes": "Name the owner of each risk."}
  ]
}
```

## Edit mode

Set `base` to the existing deck (built on the same template: its theme must match the
contract) and an `output` that is not the base file. List the output order:
`{"keep": n}` passes slide *n* through byte-identical; add `fill` (placeholder keys of
its layout), `replace` or `notes` to patch it; add `layout`/`pattern` entries for new
slides. Put base slides to delete in `"remove"`. A base slide that is neither kept
nor removed rejects the spec, so nothing is dropped by accident. Sections are kept and
new slides join the section of the slide before them. The build report counts
`unchanged/patched/added/removed` and verifies the unchanged slides byte for byte.

```json
{"contract": ".pptx-style/our-template/contract.json", "base": "board.pptx", "output": "board-v2.pptx",
 "remove": [2, 5],
 "slides": [{"keep": 1}, {"keep": 3, "fill": {"title": "Revenue up 12%"}}, {"keep": 4}, {"keep": 6}]}
```

## Exit codes

| Code | Meaning | Next step |
| --- | --- | --- |
| 0 | built, lint clean | render and review |
| 1 | built, lint/fit errors (listed) | fix the spec: shorten, notes, split; rebuild |
| 2 | spec or contract error, nothing written | read the messages ("did you mean …") |
| 3 | template drifted with breaking changes, nothing written | re-profile with `pptx-style-profile`, update the spec |
