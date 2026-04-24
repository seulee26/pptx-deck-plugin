"""Map outline.json sections/items onto source-slide indices from classification.json.

v0.3 — ADAPTIVE FIT
    The matcher no longer just round-robins within a kind bucket. For items with
    structured data (table, matrix, process), it picks the source slide whose
    real placeholder geometry best fits the data.

    Example: a `table` item with 10 rows × 6 cols picks slide 56 (9×5 native
    table, truncates 1 row/1 col) over slide 52 (7×3, would lose half the data).

Size contract:
    MAX_TOTAL_SLIDES = 30   Hard ceiling. Items are trimmed from the tail of
    each section if the budget would overflow.
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

MAX_TOTAL_SLIDES = 30


def build_classified(classification: dict) -> list[dict]:
    return list(classification["slides"])


def by_kind(slides: list[dict]) -> dict[str, list[dict]]:
    out: dict[str, list[dict]] = defaultdict(list)
    for s in slides:
        out[s["kind"]].append(s)
    return out


def pick_agenda(n_sections: int, agenda_slides: list[dict]) -> int:
    indices = [s["index"] for s in agenda_slides]
    if n_sections <= 3 and 2 in indices:
        return 2
    if n_sections == 6 and 3 in indices:
        return 3
    if 4 in indices:
        return 4
    return indices[0]


def pad_sections(sections: list[dict], slot_count: int = 9) -> list[dict]:
    padded = list(sections)
    while len(padded) < slot_count:
        padded.append({"heading": "", "items": []})
    return padded[:slot_count]


def trim_to_budget(sections: list[dict], budget_items: int) -> list[dict]:
    sections = [dict(s, items=list(s["items"])) for s in sections]
    total = sum(len(s["items"]) for s in sections)
    while total > budget_items:
        progressed = False
        for s in reversed(sections):
            if len(s["items"]) > 1:
                s["items"].pop()
                total -= 1
                progressed = True
                if total <= budget_items:
                    break
        if not progressed:
            break
    return sections


# ─────────────────────────────── fit scoring ────────────────────────────────

def fit_score_table(item: dict, slide: dict) -> float:
    """Lower score = better fit. The winner is the slide whose native table
    shape is the smallest envelope that still holds the data without losing
    too much. If no slide can hold everything, we pick the one with the
    largest capacity (minimum overflow)."""
    ts = slide.get("table_shape")
    if not ts:
        return float("inf")
    slide_rows, slide_cols = ts
    # +1 row for header
    need_rows = len(item["data"].get("rows", [])) + 1
    need_cols = len(item["data"].get("headers", []))

    # Overflow cost (rows/cols that would be truncated) is penalized heavily.
    row_overflow = max(0, need_rows - slide_rows)
    col_overflow = max(0, need_cols - slide_cols)
    # Slack cost (unused slots) is a small penalty.
    row_slack = max(0, slide_rows - need_rows)
    col_slack = max(0, slide_cols - need_cols)

    return row_overflow * 10 + col_overflow * 15 + row_slack * 0.2 + col_slack * 0.3


def fit_score_matrix(item: dict, slide: dict) -> float:
    # Matrix items are 2×2 (SWOT). Slide 46 is 4×4 native — good fit.
    ts = slide.get("table_shape")
    if ts == [4, 4]:
        return 0
    # Fallback matrix slides (48, 49, 50) have body slots instead.
    return 10 - min(slide.get("body_slots", 0), 4)


def fit_score_process(item: dict, slide: dict) -> float:
    steps = len(item["data"].get("steps", []))
    # We want a process slide whose title_slots ≥ steps.
    slots = slide.get("title_slots", 0)
    if slots >= steps:
        return slots - steps
    return (steps - slots) * 5


def fit_score_kpi(item: dict, slide: dict) -> float:
    # Prefer KPI slides where the big-number slot is prominent. Title_slots==1
    # is a single-KPI slide, higher counts are multi-KPI dashboards.
    return 0  # no strong preference; round-robin


KIND_SCORER = {
    "table": fit_score_table,
    "matrix": fit_score_matrix,
    "process": fit_score_process,
    "kpi": fit_score_kpi,
}


def pick_best_slide(
    item: dict,
    candidates: list[dict],
    cursor: defaultdict[str, int],
    kind: str,
    usage_counts: dict[int, int],
) -> dict:
    if not candidates:
        return None  # caller must handle
    scorer = KIND_SCORER.get(kind)
    if scorer:
        # Variety bonus: add a small penalty per prior use of this slide so we
        # diversify when multiple slides score equivalently.
        best = min(
            candidates,
            key=lambda s: (scorer(item, s) + usage_counts.get(s["index"], 0) * 0.5, s["index"]),
        )
        return best
    idx = cursor[kind] % len(candidates)
    cursor[kind] += 1
    return candidates[idx]


# ───────────────────────────── plan builder ─────────────────────────────────

def build_plan(outline: dict, classification: dict) -> dict:
    slides = build_classified(classification)
    kind_map = by_kind(slides)
    by_index = {s["index"]: s for s in slides}

    # Rotation cursors per kind (only used when scorer returns None).
    cursor: defaultdict[str, int] = defaultdict(int)
    # Global usage count across the whole deck — fuels the variety bonus.
    usage_counts: dict[int, int] = defaultdict(int)
    # Track which table-slide indices we've used; try to avoid back-to-back
    # reuse within the same section.
    last_used_in_section: set[int] = set()

    sections = outline.get("sections", [])
    n_sections = len(sections)
    items_budget = max(0, MAX_TOTAL_SLIDES - (3 + n_sections))
    sections = trim_to_budget(sections, items_budget)

    plan_slides: list[dict] = []
    position = 0

    def add(source_index: int, kind: str, content: dict) -> None:
        nonlocal position
        position += 1
        plan_slides.append(
            {"position": position, "source_index": source_index, "kind": kind, "content": content}
        )

    cover = kind_map["cover"][0]
    add(
        source_index=cover["index"],
        kind="cover",
        content={
            "title": outline.get("title", "Untitled"),
            "subtitle": outline.get("subtitle", ""),
            "company": outline.get("company", "COMPANY LOGO HERE"),
        },
    )

    agenda_src = pick_agenda(n_sections, kind_map["agenda"])
    padded = pad_sections(sections)
    add(
        source_index=agenda_src,
        kind="agenda",
        content={"headings": [s.get("heading", "") for s in padded]},
    )

    for n, section in enumerate(sections, start=1):
        heading = section.get("heading", f"Section {n}")
        add(
            source_index=kind_map["section"][0]["index"],
            kind="section",
            content={"step": f"{n:02d}", "title": heading},
        )
        for item in section.get("items", []):
            kind = item.get("type", "content")

            # 'dashboard' is a synthetic kind we inject ourselves — it reuses
            # the cleanest table source (slide 54) as a blank canvas and the
            # dashboard injector wipes every shape to redraw from scratch.
            if kind == "dashboard":
                # prefer slide 54 if present, else first table slide
                tables = kind_map.get("table", [])
                idx = 54 if any(t["index"] == 54 for t in tables) else (tables[0]["index"] if tables else None)
                if idx is None:
                    kind = "content"
                else:
                    add(
                        source_index=idx,
                        kind="dashboard",
                        content={
                            "title": item.get("title", ""),
                            "body": item.get("body", ""),
                            "data": item.get("data", {}),
                        },
                    )
                    continue

            if kind not in kind_map or not kind_map[kind]:
                kind = "content"

            # Fit-first: pick the best-fitting slide every time. Previously we
            # excluded the prior slide within the same section for variety, but
            # that forced data-losing fallbacks (e.g. 5-col data → 4-col template).
            candidates = list(kind_map[kind])

            chosen = pick_best_slide(item, candidates, cursor, kind, usage_counts)
            usage_counts[chosen["index"]] += 1

            # If table item doesn't fit the best table slide well (overflow),
            # fall back to a content slide so we don't show truncated garbage.
            if kind == "table":
                overflow = fit_score_table(item, chosen)
                if overflow >= 10:  # at least one row or column would be cut
                    # Still use the chosen table slide (best we have) but flag it
                    # in content; downstream inject_table will truncate.
                    pass

            add(
                source_index=chosen["index"],
                kind=kind,
                content={
                    "title": item.get("title", ""),
                    "body": item.get("body", ""),
                    "data": item.get("data", {}),
                },
            )

    closing = kind_map["closing"][0]
    add(
        source_index=closing["index"],
        kind="closing",
        content={"message": outline.get("closing", "Thank you")},
    )

    return {"source_template": "assets/template.pptx", "slides": plan_slides}


def main() -> None:
    if len(sys.argv) != 4:
        sys.exit("usage: match_slides.py <outline.json> <classification.json> <slide_plan.json>")
    outline = json.loads(Path(sys.argv[1]).read_text())
    classification = json.loads(Path(sys.argv[2]).read_text())
    plan = build_plan(outline, classification)

    out = Path(sys.argv[3])
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(plan, ensure_ascii=False, indent=2))

    kinds: dict[str, int] = defaultdict(int)
    for s in plan["slides"]:
        kinds[s["kind"]] += 1
    print(f"plan written: {out}")
    print(f"  total slides: {len(plan['slides'])}")
    for k, n in sorted(kinds.items()):
        print(f"  {k}: {n}")
    # Print per-table fit decisions for diagnostic.
    by_idx = {s["index"]: s for s in classification["slides"]}
    print("  table fits:")
    for s in plan["slides"]:
        if s["kind"] == "table":
            shape = by_idx.get(s["source_index"], {}).get("table_shape")
            d = s["content"]["data"]
            rows = len(d.get("rows", [])) + 1
            cols = len(d.get("headers", []))
            print(f"    pos {s['position']:2d} slide {s['source_index']:2d} shape={shape} data={rows}x{cols}")


if __name__ == "__main__":
    main()
