# pptx-deck plugin

Claude Code plugin that turns an **Excel / Word / PDF** into a pixel-perfect `.pptx` matching a PowerPoint template you already own. An agent team drives the pipeline and picks the right template slide for each piece of content.

Design fidelity is preserved by **cloning the template's slide XML verbatim** and swapping only the text runs — colors, fonts, layout, and logo positions are byte-identical to the source.

## Install

```bash
git clone https://github.com/<your-user>/pptx-deck-plugin.git
cd pptx-deck-plugin
pip install python-pptx openpyxl python-docx pypdf
```

LibreOffice is required for preview rendering:

```bash
brew install --cask libreoffice        # macOS
# apt install libreoffice              # linux
```

Link the plugin into Claude Code by pointing your plugin marketplace / settings at this directory, or copy it under `~/.claude/plugins/`.

## Bring your own template

The design lives entirely in **one file** you provide:

```
assets/
├── template.pptx          ← your template (not committed)
└── classification.json    ← catalog generated from it
```

One-time bootstrap:

1. Drop your template at `assets/template.pptx`.
2. Run the classifier to build `classification.json`:

   ```bash
   python3 scripts/classify_template.py assets/template.pptx assets/classification.json
   ```

   (The classifier identifies each slide as one of `cover / agenda / section / kpi / process / matrix / table / image / content / closing` — rules live in `scripts/classify_template.py`.)

3. Everything downstream reads from `classification.json` — the agents never touch the template directly.

## Usage

```
/make-deck path/to/input.xlsx
/make-deck path/to/input.docx out/my-deck.pptx
/make-deck path/to/report.pdf
```

`deck-orchestrator` runs the pipeline end-to-end and returns the final deck path plus a QA report.

### Excel input convention

| sheet     | rows                                                                                     |
|-----------|------------------------------------------------------------------------------------------|
| `meta`    | `title`, `subtitle`, `company`, `closing` as key/value rows (all optional)               |
| anything else | one sheet per section (sheet name = heading). Item rows start with `KPI`, `PROCESS`, `MATRIX`, `TABLE`, `IMAGE`, or `CONTENT`; remaining cells carry data. Unmarked rows become `CONTENT` items. |

See `scripts/make_sample_xlsx.py` for a working example.

Word uses `Heading 1/2/3` hierarchy; PDF parses per page.

## Agent team

```
             ┌─────────────────────────┐
   user ───▶ │  deck-orchestrator      │  (director)
             └───┬──────┬──────┬───┬───┘
                 │      │      │   │
       content-parser   │      │   qa-verifier
          (xlsx/docx/   │      │   (render + lint)
           pdf → outline)│     │
                  slide-matcher
                  (outline + classification.json
                   → slide_plan)
                               │
                       deck-assembler
                  (clone XML, inject text → deck.pptx)
```

## Layout

```
.claude-plugin/plugin.json
agents/            # 5 agent prompts
commands/          # /make-deck slash command
scripts/           # Python pipeline (parse, match, assemble, render)
assets/            # your template + classification.json (template gitignored)
```

## Pipeline contract (pixel-perfect)

- Slides are **cloned** from the template — no new shapes, layouts, or masters are created.
- Only run-level `.text` is modified. `<a:rPr>` (font, size, color, spacing) is never touched.
- If a placeholder pattern isn't found on a slide, the slide is left as-is and the QA report flags it — we do not "best effort" write into a random text box.

## License

Plugin code: MIT. Your template and any generated decks are yours — nothing from this repo claims rights on them.
