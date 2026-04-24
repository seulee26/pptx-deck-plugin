"""Map outline.json sections/items onto source-slide indices from classification.json.

Produces slide_plan.json that deck-assembler consumes. Implements the rules in
agents/slide-matcher.md: cover + agenda + (section divider + items...)×N + closing,
with round-robin inside each kind-bucket.
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path


def build_buckets(classification: dict) -> dict[str, list[int]]:
    buckets: dict[str, list[int]] = defaultdict(list)
    for entry in classification["slides"]:
        buckets[entry["kind"]].append(entry["index"])
    return buckets


def pick_agenda(n_sections: int, agenda_bucket: list[int]) -> int:
    if n_sections <= 3 and 2 in agenda_bucket:
        return 2
    if n_sections == 6 and 3 in agenda_bucket:
        return 3
    if 4 in agenda_bucket:
        return 4
    return agenda_bucket[0]


def pad_sections(sections: list[dict]) -> list[dict]:
    """Agenda slides have 6 placeholder slots. Pad or trim so the count matches."""
    padded = list(sections)
    while len(padded) < 6:
        padded.append({"heading": "", "items": []})
    return padded[:6]


def build_plan(outline: dict, classification: dict) -> dict:
    buckets = build_buckets(classification)
    bucket_cursor: dict[str, int] = defaultdict(int)

    def take(kind: str) -> int:
        bucket = buckets.get(kind) or buckets["content"]
        idx = bucket[bucket_cursor[kind] % len(bucket)]
        bucket_cursor[kind] += 1
        return idx

    plan_slides: list[dict] = []
    position = 0

    def add(source_index: int, kind: str, content: dict) -> None:
        nonlocal position
        position += 1
        plan_slides.append(
            {"position": position, "source_index": source_index, "kind": kind, "content": content}
        )

    add(
        source_index=buckets["cover"][0],
        kind="cover",
        content={
            "title": outline.get("title", "Untitled"),
            "subtitle": outline.get("subtitle", ""),
            "company": outline.get("company", "COMPANY LOGO HERE"),
        },
    )

    sections = outline.get("sections", [])
    agenda_src = pick_agenda(len(sections), buckets["agenda"])
    padded = pad_sections(sections)
    add(
        source_index=agenda_src,
        kind="agenda",
        content={"headings": [s["heading"] for s in padded]},
    )

    for n, section in enumerate(sections, start=1):
        heading = section.get("heading", f"Section {n}")
        add(
            source_index=buckets["section"][0],
            kind="section",
            content={"step": f"{n:02d}", "title": heading},
        )
        for item in section.get("items", []):
            kind = item.get("type", "content")
            if kind not in buckets or not buckets[kind]:
                kind = "content"
            add(
                source_index=take(kind),
                kind=kind,
                content={
                    "title": item.get("title", ""),
                    "body": item.get("body", ""),
                    "data": item.get("data", {}),
                },
            )

    add(
        source_index=buckets["closing"][0],
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

    kinds = defaultdict(int)
    for s in plan["slides"]:
        kinds[s["kind"]] += 1
    print(f"plan written: {out}")
    print(f"  total slides: {len(plan['slides'])}")
    for k, n in sorted(kinds.items()):
        print(f"  {k}: {n}")


if __name__ == "__main__":
    main()
