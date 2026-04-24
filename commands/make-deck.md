---
description: Generate a pixel-perfect PPT from an Excel/Word/PDF input using the classified template.
argument-hint: <input-file> [output-path]
---

Build a PPT deck from `$1` (Excel `.xlsx`, Word `.docx`, or `.pdf`) using the `pptx-deck` agent team.

Output path: `$2` if provided, otherwise `out/deck.pptx` inside the plugin.

**Hand off to the `deck-orchestrator` subagent immediately.** Pass it:

- `input_path`: `$1`
- `output_path`: `$2` (optional — orchestrator chooses default)
- plugin root: the directory of this plugin (use it to locate `assets/template.pptx`, `assets/classification.json`, and `scripts/*.py`)

The orchestrator will coordinate `content-parser` → `slide-matcher` → `deck-assembler` → `qa-verifier` and return the final deck path plus any QA notes. Do not re-invoke individual sub-agents yourself.
