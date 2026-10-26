# AI-readiness findings

What the profiler reports to the template owner, why it matters to agents and to
Copilot-style tools, and the fix. Severity: `error` blocks reliable automation,
`warning` causes wrong output unless handled, `info` reduces quality.

| Code | Severity | Why it matters | Fix for the template owner |
| --- | --- | --- | --- |
| `duplicate-layout-name` | error | name-based layout selection is ambiguous | rename one layout |
| `multiple-masters` | warning | theme and text styles are read from the first master only; names may repeat across masters | one master per template file, unique layout names |
| `no-sample-slides` | warning | nothing to clone: builds can only fill bare layouts | add sample slides that show intended content |
| `prompt-text-as-content` | warning | the value (e.g. agenda `01`) shows only in edit view, never in the show or PDF | type the value on the slide, or make it real text on the layout |
| `textbox-instead-of-placeholder` | warning | tools fill placeholders; a free text box on a layout is invisible to them | replace with a placeholder (or keep it as deliberate fixed chrome) |
| `duplicate-shape-names` | warning | name-based addressing hits the wrong shape | rename in the Selection Pane; the contract already addresses by index path |
| `hardcoded-colour` | warning | colours outside the theme survive a re-theme and spread when cloned | map to a theme slot |
| `export-artifact-names` | info | names like `Google Shape;3265;p106` carry no meaning | rename to describe the role |
| `decorative-sample-slide` | info | no placeholders: only clone-and-retext can reuse it | rebuild on placeholders if it should be AI-fillable |
| `layout-without-sample` | info | tools fall back to the bare layout without a density example | add a sample slide per layout |
| `localized-layout-name` | info | agents and colleagues may not recognise the name | rename to a descriptive English name |
| `accent-fails-aa-as-text` | info | accents below 4.5:1 make small text unreadable | reserve for fills or large text; small text uses dk1/dk2 |

Facts behind these checks: `ooxml-facts.md` (prompt text, Copilot guidance, WCAG).

Report findings as a list grouped by severity, each with its fix. Findings describe the
template; they never block building inside the contract, but the build skill's lint
honours several of them (prompt-only placeholders must be filled, template-own
off-theme colours stay a warning).
