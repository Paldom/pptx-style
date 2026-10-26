# PresentationML facts the profiler relies on

Verified against primary sources on 2026-10-26. Each fact names what the code does with it.

**Contents:** [Inheritance](#placeholder-inheritance) · [Prompt text](#prompt-text) ·
[Slide numbers](#slide-numbers) · [Theme](#theme-colours-and-fonts) · [IDs](#ids) ·
[Templates](#potx-vs-pptx) · [Markup compatibility](#markup-compatibility) ·
[Sections](#sections) · [Contrast](#contrast) · [AI-ready templates](#what-ai-tools-read-in-a-template) ·
[Prior art](#prior-art)

## Placeholder inheritance

A slide placeholder inherits position, size, fill, line and font from the **layout**
placeholder with the same `idx`; a layout placeholder inherits from the **master**
placeholder with the same `type`. A placeholder without `a:xfrm` uses the inherited
position. The profiler resolves boxes and font sizes along that chain; slots record
the resolved geometry.
Sources: [python-pptx: understanding placeholders](https://python-pptx.readthedocs.io/en/latest/user/placeholders-understanding.html),
[Open XML SDK `PlaceholderShape`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.placeholdershape).

## Prompt text

Text inside a layout or master placeholder is *prompt text*: PowerPoint shows it in
Normal view while the slide placeholder is empty and never in Slide Show, print or
PDF. `hasCustomPrompt="1"` marks a custom prompt. A template whose agenda numbers
`01`–`07` exist only as prompt text produces an agenda without numbers unless each
number is typed into the slide. The profiler flags such prompts
(`prompt_looks_like_content`) and the builder refuses to leave them empty.
Sources: [`PlaceholderShape.HasCustomPrompt`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.placeholdershape),
[Microsoft Q&A on prompt text visibility](https://learn.microsoft.com/en-us/answers/questions/5094076/powerpoint-slide-master-text-automatically-viewabl).

## Slide numbers

A slide number is a text field `<a:fld type="slidenum">`, normally inside a
`ph type="sldNum"` placeholder but legal in any text box (cached text `‹#›` on
layouts and masters). The profiler reports both forms as `slide_number` chrome and
never treats a field-bearing shape as a fillable slot.
Source: [`HeaderFooter.SlideNumber`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.headerfooter.slidenumber).

## Theme colours and fonts

`a:clrScheme` holds 12 slots (`dk1 lt1 dk2 lt2 accent1–6 hlink folHlink`), each an
`srgbClr` or a `sysClr` whose `lastClr` carries the hex. `a:fontScheme` has
`majorFont`/`minorFont` with a `latin` typeface; text styles refer to them as
`+mj-lt`/`+mn-lt`. Masters map `bg1/tx1/bg2/tx2` onto scheme slots through `p:clrMap`;
slides and layouts use `p:clrMapOvr`.
Sources: [`ColorMap`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.colormap),
[Structure of a PresentationML document](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document).

## IDs

`p:sldId/@id` is in `256 … 2147483647`; slide-master and slide-layout ids are
`≥ 2147483648`. Every id must be unique in the presentation.
Sources: [ST_SlideId](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_SlideId_topic_ID0EOLIIB.html),
[ST_SlideLayoutId](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_SlideLayoutId_topic_ID0EUPIIB.html).

## .potx vs .pptx

The only required difference is the main part's content type:
`…presentationml.template.main+xml` (template) vs `…presentationml.presentation.main+xml`.
The profiler reads both; the builder rewrites the content type when writing `.pptx`.
Source: [Open XML Package Editor known content types](https://github.com/OfficeDev/Open-XML-Package-Editor-Power-Tool-for-Visual-Studio/blob/master/KnownPackageInfo.cs).

## Markup compatibility

`mc:Ignorable` is a whitespace-separated list of namespace **prefixes**; each must be
declared in scope. Python's ElementTree drops declarations no element uses and renames
prefixes, so a naive round trip leaves `mc:Ignorable="p14 a14"` pointing at nothing and
PowerPoint offers to repair the file. The shared core records every declaration on
parse and re-injects it on write; untouched parts are never re-serialized.
Sources: [ECMA-376 Part 3, Ignorable](https://c-rex.net/samples/ooxml/e1/Part5/OOXML_P5_Markup_Compatibility_and_Extensibility_Ignorable_topic_ID0ERFAG.html),
[Introduction to markup compatibility](https://learn.microsoft.com/en-us/office/open-xml/general/introduction-to-markup-compatibility).

## Sections

Sections are a Microsoft extension in `p:presentation/p:extLst/p:ext[@uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"]/p14:sectionLst`,
each `p14:section` listing `p14:sldId/@id` values that match `p:sldId/@id`. A slide
listed in two sections makes the file unreadable; stale ids are hygiene, not a proven
repair trigger.
Source: [MS-PPTX 2.2.5 sectionLst](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/1f21a089-944d-410b-bd47-4f5e692c2532).

## Contrast

WCAG 2.x contrast ratio `(L1 + 0.05) / (L2 + 0.05)` with sRGB linearisation threshold
0.04045. AA needs 4.5:1 for normal text and 3:1 for large text (≥ 18 pt, or ≥ 14 pt
bold). The contract's contrast table applies this to every theme colour on the four
background slots.
Source: [Understanding SC 1.4.3](https://www.w3.org/WAI/WCAG22/Understanding/contrast-minimum.html).

## What AI tools read in a template

Microsoft's guidance for Copilot-ready templates: use text placeholders instead of
text boxes; provide representative sample slides at several content densities (Copilot
learns mostly from sample slides and falls back to the bare master when there are too
few); keep colours, fonts and logos in the theme and master because "Copilot learns
from theme definitions, not one-off formatting". The *strict brand adherence* setting
restricts Copilot to the template's layouts with no added or removed placeholders.
These map onto the AI-readiness checks (see `ai-readiness.md`).
Sources: [Keep your presentation on-brand with Copilot](https://support.microsoft.com/en-us/powerpoint/copilot/keep-your-presentation-on-brand-with-copilot),
[Manage brand kit template settings](https://support.microsoft.com/en-us/powerpoint/copilot/manage-brand-kit-template-settings-in-powerpoint).

## Prior art

Ideas (no code) came from MIT/Apache-licensed projects: hashed manifests and breaking
vs additive diffs (ThomasRohde/pptx-cli), placeholder signatures and collision groups
(tristan-mcinnis/pptx-from-layouts-skill), index-path slot addressing and per-slide
fill plans (CxyZyr/PPTX-Template-Skills), "a fact about slide 7 vs the reusable rule"
(hugohe3/ppt-master), and the content-area safe box (anthropics/financial-services).
