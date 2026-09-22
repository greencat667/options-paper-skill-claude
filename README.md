# Options Paper

A skill for [Claude Code](https://docs.anthropic.com/en/docs/agents-and-tools/claude-code/overview) and [Cowork](https://claude.ai) that writes structured options papers — a document format for presenting a genuine strategic choice to a leadership team or working group without pre-deciding for them — and exports them as a branded `.docx`.

## What it does

An options paper separates the analytical work (what are the options?) from the political work (which do we choose?). Ask for "options for X" and the skill produces a document with:

- A named tension or opportunity, not a generic title
- 3 familiar options plus one deliberately "more out there" option (★)
- Each option given a fair hearing: core idea, mechanism, strengths, risks, and "best for"
- A comparison table making trade-offs visible at a glance
- A suggested approach — framed as a suggestion, since the final call sits with the reader
- Concrete next steps

It also runs a "sceptic pass" before finalising: are the risks honest, is any option a straw man, has a preference been smuggled into the suggested approach?

See [`options-paper/references/example-corporate-sponsorship.md`](options-paper/references/example-corporate-sponsorship.md) for a full worked example.

## Installation

**Ask Claude to set it up for you.** If you're using Claude Code or Claude Cowork, you can just say something like *"install the options-paper skill from github.com/greencat667/options-paper-skill-claude"* and Claude will clone the repo and put it in the right place — you don't need to do this by hand.

Or do it yourself: copy the `options-paper/` folder into your project's `.claude/skills/` directory:

```bash
git clone https://github.com/YOUR_USERNAME/options-paper-skill-claude.git
cp -r options-paper-skill-claude/options-paper/ your-project/.claude/skills/options-paper/
```

Claude will pick it up automatically from `available_skills` next time you start a session.

## Producing the Word document

The skill drafts the content, then hands it to the bundled `scripts/build_options_paper.py` to
produce the `.docx`:

```bash
python3 options-paper/scripts/build_options_paper.py \
  --content content.json \
  --template your-org-template.docx \
  --output paper.docx
```

`--template` can be any `.docx` that already defines `Title`, `Subtitle`, `Heading 1`, `Heading 2`,
and a list style — a blank document saved from your organisation's letterhead works well. The
script reads that file's styles, headers, and footers and injects the generated content; it
doesn't need to be a template built specifically for this skill.

## Repository structure

```
options-paper-skill-claude/
├── options-paper/
│   ├── SKILL.md                              # Copy this folder to .claude/skills/
│   ├── references/
│   │   └── example-corporate-sponsorship.md  # Full worked (fictional) example
│   └── scripts/
│       └── build_options_paper.py            # JSON + template → branded .docx
├── README.md
├── CONTRIBUTING.md
└── LICENSE
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) — note this repo isn't actively maintained, so response times on issues and PRs will be slow to nonexistent.

## License

MIT — see [LICENSE](LICENSE).
