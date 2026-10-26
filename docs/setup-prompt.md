# Setup prompt: a deck in your template, end to end

The two skills compose into one workflow: **profile** the template once, then
**build** decks inside its contract. Paste the prompt below into Claude Code as a
`/goal` after installing the skills (`npx skills add Paldom/pptx-style`), and fill
in the four inputs. It orders the steps, keeps parallel agents on separate files,
puts a gate after each skill, and never touches git.

```text
/goal Produce the deck described below in our PowerPoint template with the pptx-style skills, and stop only when every gate passes. Work autonomously. Never run git commit, push or any other git write; leave every file in place for me.

Inputs
- TEMPLATE: <path to our .pptx or .potx>
- SOURCE: <notes, outline or document the deck is built from>
- AUDIENCE and GOAL: <who sees it, what they should decide or remember>
- LENGTH: <number of slides>

Steps (in this order)
1. Profile. Run /pptx-style-profile on TEMPLATE. If .pptx-style/ already holds a contract for this template, profile into a new folder and diff against the old one instead; on breaking changes stop and show me the diff.
   Gate P: .pptx-style/<slug>/contract.json and STYLE.md exist; you have read STYLE.md and listed its AI-readiness findings for me (they describe the template, they do not block the build).
2. Plan, two agents in parallel on separate files only:
   - Agent A writes outline.md from SOURCE: one message per slide, titles that tell the story when read alone, 2-5 sentences of speaker notes per content slide.
   - Agent B writes mapping.md from STYLE.md: which layout or pattern suits each kind of content (timeline, comparison, cards, numbers, bullets), with capacities.
   Neither agent edits the other's file, the template or .pptx-style/.
3. Spec, one agent: merge outline.md and mapping.md into deck.json for /pptx-style-build. Use only layout names, pattern ids and fill keys that appear in STYLE.md; use at least three different patterns where the content fits them; notes on every content slide.
4. Preview gate: with /pptx-style-build, build only the cover and the hardest slide, render them and show me the contact sheet. Wait for my approval. If I am not around, approve it yourself and say so in the report.
5. Build: with /pptx-style-build, build the full deck from deck.json.
   Gate B: the build report (<deck>.build.json) shows 0 lint errors, the slide count equals LENGTH (plus continuation slides if you split), and no off-contract formatting was requested. Fix overflow by shortening, moving detail into the notes, or splitting. Never shrink fonts. At most two fix passes.
6. Review: a fresh read-only subagent reviews the rendered contact sheet against the hard/soft rubric of pptx-style-build. Fix hard findings in deck.json and rebuild, at most two passes. If LibreOffice is missing, say the visual review is still owed instead of claiming it passed. If the render is labelled APPROXIMATE, say which font was missing.

Done when Gates P and B pass and the review has no open hard findings.

Report: the deck path, slide count, patterns used, lint result, render label (OK, APPROXIMATE or skipped), the AI-readiness findings worth fixing in the template, and anything left unfixed with the reason.
```

## Why it is shaped this way

- **Profile first, once.** Every build checks the template against the contract's
  hashes; a changed template stops the build (exit 3) instead of drifting silently.
- **Parallel only on disjoint files.** The outline and the layout mapping can be
  written at the same time because they never share a file; the spec, the build and
  the fixes stay with one agent so nothing overwrites anything.
- **Gates are files and exit codes, not impressions.** `contract.json`, the build
  report and the render label are what the gates read.
- **Preview before volume.** Two approved slides are cheaper to fix than eighteen.
