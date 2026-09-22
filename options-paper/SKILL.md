---
name: options-paper
description: >
  Produce a structured options paper in a branded Word format (.docx). Use this skill whenever
  someone asks for "options for", "framing options", "options paper", "what are our options",
  "give me 3/4/5 options", "how could we approach X", or any request to explore multiple strategic
  approaches to a tension, decision, or opportunity. The output is always a branded .docx with a
  comparison table and suggested approach — not just a list of ideas. If someone asks for "pros
  and cons" of multiple approaches, or wants to present choices to a leadership team or working
  group, this is the right skill.
---

# Options Paper

An options paper presents a strategic question, lays out distinct approaches for addressing it,
and gives the reader enough to make a decision — without making the decision for them. It's a
good match whenever you need to present a genuine choice to a leadership team, board, or working
group and want them to decide, not just agree with you.

The format works because it separates the analytical work (what are the options?) from the
political work (which do we choose?). Each option is given a fair hearing. The comparison table
makes trade-offs visible. The suggested approach shows how the options relate to each other —
they're often complementary, not mutually exclusive.

## When to use this vs. other formats

- **Options paper** — when there's a genuine choice between approaches and the reader needs to
  decide. The question is "how should we do this?" not "what should we do?"
- **Project brief** — when the approach is already chosen and you need to scope the work. An
  options paper is the wrong format once a decision has been made.
- **Landscape / trend report** — when the task is understanding a landscape, not choosing between
  approaches.

## Document structure

Every options paper follows this structure. Don't skip sections — the format's value is in its
completeness.

### 1. Title block

```
[Title — short, named tension or opportunity]     ← Title style
[Subtitle — "N options for [doing what]"]          ← Subtitle style
```

The title should name the tension or opportunity, not describe the document. "The Data Centre
Tension" works. "Options Paper About Data Centres" doesn't.

Don't add front matter (For/From/Date/Status). The document speaks for itself — the audience
and author are clear from context. If it matters, put it in the email or the tracking system you
use, not the paper.

### 2. Context

Two paragraphs maximum. Sets up *why this paper exists now*. What's changed, what's been
surfaced, what's the trigger? Reference specific events, feedback, or decisions where possible —
this isn't an abstract policy discussion, it's a response to something concrete.

### 3. The problem / tension / opportunity

Name it clearly. If it's a tension (two things pulling in opposite directions), state both sides.
If it's an opportunity (several assets that could be combined), describe each asset. This section
should make the reader feel the weight of the question — why it can't just be ignored.

### 4. The options

Default: **3 familiar options + 1 "more out there" option** (marked with ★). The user can
override this — they might want 5 options, or 2+1, or all familiar. But the 3+1 default works
well because it gives the reader a range from "we could start this tomorrow" to "this would be
transformative if we committed to it."

Each option follows this sub-structure:

```
# Option N: [Short memorable name]              ← Heading 1

## Core idea                                     ← Heading 2
One paragraph. What is this, in plain language? A reader should understand the option
from this paragraph alone.

## How it works                                  ← Heading 2
One to two paragraphs. The mechanism — what would actually happen, who does what, what
does the process look like? Include analogies or precedents where they help ("this is
similar to how X works in Y context").

## Strengths                                     ← Heading 2
3–5 bullet points. Each starts with a bold two-to-four word label, then explains.
Be specific — "Low cost" is weak; "Low cost because it packages existing work rather
than creating new infrastructure" is strong.

## Risks                                         ← Heading 2
3–5 bullet points. Same format. Be honest — the point is to help the reader make a
real decision, not to sell the option. Name political risks (who might object?) as well
as operational ones.

## Best for                                      ← Heading 2
One paragraph. When would you choose this option? What does it pair well with? What
would need to be true for this to be the right move?
```

**Heading hierarchy matters:** Title → Subtitle → Heading 1 (sections and option names) →
Heading 2 (sub-sections within options). This ensures the document looks right in Word's
navigation pane and table of contents.

**Naming options well matters.** A good option name is memorable, evocative, and tells you
something about the approach. "The Reciprocity Condition", "The Grant Matchmaker", "Build the
Alternative" — these stick. "Option A: Structured Approach" doesn't.

**The "out there" option (★)** should be genuinely ambitious — not reckless, but the kind of
thing that would be transformative if it worked. It's there to stretch thinking and give the
reader permission to be bold. Mark it with ★ in the heading and add an italic note:
*The "more out there" option.*

### 5. Comparison table

A table with columns: **Option | Strength | Main risk | Operational cost | Best pairing**

One row per option. Keep cells short — this is a summary, not a repeat of the full analysis.
The "Best pairing" column is important: it shows the reader how options relate to each other.

### 6. Suggested approach

This is where the author offers a view — but framed as a suggestion, not a recommendation.
Common patterns:

- **Layered approach:** different options serve different purposes (narrative / operational /
  campaign / aspiration)
- **Sequential approach:** options build on each other over time (start with X, add Y in Q2,
  scope Z for next year)
- **Pick one + park one:** recommend one option now, hold another for later

Always acknowledge that the final choice sits with the reader.

### 7. Next steps

5–7 bullet points. Concrete, actionable, assigned where possible. These should flow naturally
from the suggested approach — "if we go with this direction, here's what happens next."

## The writing process

### Pass 1: Builder

Write the strongest, fairest version of each option. Give each one a genuine shot — don't
sandbag options you think are weaker. The reader needs to trust that the analysis is balanced,
not that you've pre-decided.

Draw on whatever context the user has provided. If they've given a braindump, extract the key
tensions and assets from it. If they've pointed to existing documents, use evidence mode — only
state what's in the source material, and flag where you're going beyond it.

If the user has asked you to research, do so — but keep the research targeted and flag which
claims come from research vs. existing context.

### Pass 2: Sceptic

Before presenting to the user, review the full draft with a sceptic's eye:

- Are the risks honest? Or have you softened them to make options look better than they are?
- Is any option a straw man — included just to be knocked down?
- Are there political or relational risks you've missed? (Who would object to this? Whose
  territory does it encroach on?)
- Is the "suggested approach" actually justified by the analysis, or have you smuggled in a
  preference?
- Does the comparison table accurately reflect the full write-ups?
- Are there gaps in evidence that you should flag rather than paper over?

Surface anything significant from the sceptic pass to the user before finalising.

## Producing the .docx

Use the bundled `scripts/build_options_paper.py` script. It takes a JSON content file and a Word
template, and produces a branded `.docx`.

### Step 1: Write the content as JSON

Create a JSON file with this structure:

```json
{
  "title": "The Data Centre Tension",
  "subtitle": "Five options for navigating internal AI use alongside data centre campaigning",
  "context": [
    "First paragraph of context...",
    "Second paragraph of context..."
  ],
  "problem_heading": "The tension",
  "problem_paragraphs": [
    "First paragraph describing the problem..."
  ],
  "problem_bullets": [
    {"bold": "Asset name. ", "text": "Description of the asset..."}
  ],
  "problem_closing": "Optional closing paragraph after bullets.",
  "options": [
    {
      "number": 1,
      "name": "The Reciprocity Condition",
      "star": false,
      "core_idea": "One paragraph...",
      "how_it_works": ["Paragraph 1...", "Paragraph 2..."],
      "strengths": [
        "Bold label. Rest of the bullet point..."
      ],
      "risks": [
        "Bold label. Rest of the bullet point..."
      ],
      "best_for": "One paragraph..."
    }
  ],
  "comparison": {
    "headers": ["Option", "Strength", "Main risk", "Operational cost", "Best pairing"],
    "rows": [
      ["1. Reciprocity", "Values-aligned", "Self-enforcing", "Low", "Pair with 2"]
    ]
  },
  "suggested_approach_intro": "No single option resolves the tension on its own.",
  "suggested_approach_items": [
    {"bold": "Now (Q4): ", "text": "Start with Option 1..."}
  ],
  "next_steps": [
    "First concrete next step...",
    "Second concrete next step..."
  ]
}
```

### Step 2: Run the builder

```bash
python3 <skill-path>/scripts/build_options_paper.py \
  --content content.json \
  --template <path-to-your-org-template.docx> \
  --output <output-path.docx>
```

The script handles all XML construction, branding, header/footer preservation, and table
formatting from whatever template you point it at. You provide the content; it handles the
packaging.

### Finding a template

Point `--template` at any Word document that already defines the styles this skill uses —
`Title`, `Subtitle`, `Heading 1`, `Heading 2`, and a numbered/bulleted list style. A blank
document saved from your organisation's letterhead works well; the script only reads that
file's styles, headers, and footers, so it doesn't need to be a template built specifically for
this skill. If you only have a `.doc` file, convert it first:

```bash
soffice --headless --convert-to docx your-template.doc
```

Exit code 134 is harmless (malloc warning).

## Reference

For a full example of the format in action, read `references/example-corporate-sponsorship.md`
in this skill's directory. It shows the complete structure with worked (fictional) content.
