"""Build the output PPT from slide_plan.json by:
  1. opening the template
  2. cloning the slide XML of each planned source slide (append to end)
  3. swapping placeholder run text per-kind
  4. dropping all original template slides so only the new ones remain, in plan order
  5. saving to --output

Shape geometry, masters/layouts, images, and run formatting are untouched — this
satisfies the pixel-perfect contract in agents/deck-assembler.md.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import shutil
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Pt  # noqa: F401

BIG_NUM_RE = re.compile(r"\d{1,3}(?:,\d{3})+|\d{1,3}%|\b\d{3,}\b|\b\d{2}%")
STEP_RE = re.compile(r"(?<!\d)0[1-9](?!\d)")

TITLE_PLACEHOLDERS = (
    "메인 타이틀을 입력하세요",
    "서브 타이틀을 입력하세요",
    "타이틀 입력",
    "타이틀을 입력하세요",
    "내용 입력",
    "내용을 입력하세요",
    "상세 타이틀",
    "목차를 입력해주세요",
    "목차를 상세하게 입력해주세요",
)

BODY_PLACEHOLDER_MARKERS = (
    "상세 타이틀",
    "내용 입력",
    "내용을 입력",
    "매주 업데이트되는",
    "다양한 주제와 디자인",
    "비즈니스, 교육, 마케팅",
    "저희는 프레젠테이션",
    "세련되고 정교한",
    "안녕하세요. 매주",
)


def is_body_placeholder(text: str) -> bool:
    return any(m in text for m in BODY_PLACEHOLDER_MARKERS)


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def set_runs_text(tf, new_text: str) -> bool:
    """Write `new_text` into the first run of the first paragraph, blanking the rest.
    Returns True if at least one run was modified."""
    if not new_text:
        return False
    runs = [r for p in tf.paragraphs for r in p.runs]
    if not runs:
        return False
    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""
    return True


def replace_text_in_shape(shape, old: str, new: str) -> bool:
    if not shape.has_text_frame or not new:
        return False
    txt = shape.text_frame.text
    if old and old in txt:
        return set_runs_text(shape.text_frame, new)
    return False


def first_matching_shape(slide, predicate):
    for shape in iter_shapes(slide.shapes):
        if predicate(shape):
            return shape
    return None


def _explicit_font_max(tf) -> int:
    """Max explicit run font size; 0 if all runs inherit from master."""
    best = 0
    for p in tf.paragraphs:
        for r in p.runs:
            sz = r.font.size
            if sz is not None:
                best = max(best, sz)
    return best


def inject_cover(slide, content: dict) -> list[str]:
    title = content.get("title", "")
    subtitle = content.get("subtitle", "")
    company = content.get("company", "")

    text_shapes = []
    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text.strip()
        if not t:
            continue
        text_shapes.append((shape, t, _explicit_font_max(shape.text_frame)))

    inherited = [(s, t) for s, t, sz in text_shapes if sz == 0 and "LOGO" not in t]
    title_shape = None
    if inherited:
        title_shape = min(inherited, key=lambda x: len(x[1]))[0]
    else:
        non_logo = [x for x in text_shapes if "LOGO" not in x[1]]
        if non_logo:
            title_shape = max(non_logo, key=lambda x: x[2])[0]

    if title_shape is not None and title:
        set_runs_text(title_shape.text_frame, title)

    if subtitle:
        for shape, t, _ in text_shapes:
            if shape is title_shape or "LOGO" in t:
                continue
            if len(t) >= 15:
                set_runs_text(shape.text_frame, subtitle)
                break

    if company:
        for shape, t, _ in text_shapes:
            if "COMPANY LOGO" in t or "LOGO HERE" in t:
                set_runs_text(shape.text_frame, company)
                break
    return []


def inject_agenda(slide, content: dict) -> list[str]:
    headings = content.get("headings", [])
    slot = 0
    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text.strip()
        if "목차를" in t and "입력" in t:
            heading = headings[slot] if slot < len(headings) else ""
            set_runs_text(shape.text_frame, heading or " ")
            slot += 1
    return []


def inject_section(slide, content: dict) -> list[str]:
    step = content.get("step", "01")
    title = content.get("title", "")
    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text.strip()
        if STEP_RE.fullmatch(t):
            set_runs_text(shape.text_frame, step)
        elif any(p in t for p in ("타이틀 입력", "타이틀을 입력")) and title:
            set_runs_text(shape.text_frame, title)
    return []


def inject_kpi(slide, content: dict) -> list[str]:
    title = content.get("title", "")
    body = content.get("body", "")
    value = (content.get("data") or {}).get("value", "")

    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text.strip()
        if value and BIG_NUM_RE.search(t) and not any(p in t for p in TITLE_PLACEHOLDERS):
            set_runs_text(shape.text_frame, value)
        elif title and any(p in t for p in ("타이틀 입력", "타이틀을 입력")):
            set_runs_text(shape.text_frame, title)
        elif body and any(p in t for p in ("내용 입력", "내용을 입력")):
            set_runs_text(shape.text_frame, body)
    return []


def inject_process(slide, content: dict) -> list[str]:
    steps = (content.get("data") or {}).get("steps", []) or []
    title = content.get("title", "")

    slot = 0
    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text.strip()
        if any(p in t for p in ("타이틀 입력", "타이틀을 입력")) and title and slot == 0:
            set_runs_text(shape.text_frame, title)
        elif any(p in t for p in ("내용 입력", "내용을 입력")) and slot < len(steps):
            set_runs_text(shape.text_frame, steps[slot])
            slot += 1
    return []


def inject_matrix(slide, content: dict) -> list[str]:
    data = content.get("data") or {}
    values = [data.get(k, "") for k in ("S", "W", "O", "T")]
    title = content.get("title", "")
    body = content.get("body", "")

    title_done = False
    sub_done = False
    body_slots: list = []
    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text.strip()
        if not title_done and "메인 타이틀" in t:
            set_runs_text(shape.text_frame, title or " ")
            title_done = True
        elif not sub_done and "서브 타이틀" in t:
            set_runs_text(shape.text_frame, body or " ")
            sub_done = True
        elif is_body_placeholder(t):
            body_slots.append(shape)

    for shape, val in zip(body_slots, values):
        if val:
            set_runs_text(shape.text_frame, val)

    return [] if body_slots else ["matrix-no-slots"]


def inject_table(slide, content: dict) -> list[str]:
    data = content.get("data") or {}
    headers = data.get("headers", [])
    rows = data.get("rows", [])
    title = content.get("title", "")
    body = content.get("body", "")

    title_done = False
    sub_done = False
    for shape in iter_shapes(slide.shapes):
        if shape.has_table:
            tbl = shape.table
            n_cols = len(tbl.columns)
            n_rows = len(tbl.rows)
            for c, header in enumerate(headers[:n_cols]):
                if header:
                    set_runs_text(tbl.rows[0].cells[c].text_frame, str(header))
            for r_idx, row in enumerate(rows[: n_rows - 1], start=1):
                for c_idx, val in enumerate(row[:n_cols]):
                    if val:
                        set_runs_text(tbl.rows[r_idx].cells[c_idx].text_frame, str(val))
        elif shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if not title_done and "메인 타이틀" in t:
                set_runs_text(shape.text_frame, title or " ")
                title_done = True
            elif not sub_done and "서브 타이틀" in t:
                set_runs_text(shape.text_frame, body or " ")
                sub_done = True
            elif t == "타이틀 입력":
                set_runs_text(shape.text_frame, " ")
    return []


def inject_image(slide, content: dict) -> list[str]:
    title = content.get("title", "")
    body = content.get("body", "")
    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text.strip()
        if any(p in t for p in ("타이틀 입력", "타이틀을 입력")) and title:
            set_runs_text(shape.text_frame, title)
        elif any(p in t for p in ("내용 입력", "내용을 입력")) and body:
            set_runs_text(shape.text_frame, body)
    return []


def inject_content(slide, content: dict) -> list[str]:
    title = content.get("title", "")
    body = content.get("body", "")
    used_title = False
    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text.strip()
        if not used_title and title and any(p in t for p in ("타이틀 입력", "타이틀을 입력", "메인 타이틀")):
            set_runs_text(shape.text_frame, title)
            used_title = True
        elif body and any(p in t for p in ("내용 입력", "내용을 입력", "서브 타이틀")):
            set_runs_text(shape.text_frame, body)
    return []


def inject_closing(slide, content: dict) -> list[str]:
    message = content.get("message", "Thank you")
    title_done = False
    sub_done = False
    for shape in iter_shapes(slide.shapes):
        if not shape.has_text_frame:
            continue
        t = shape.text_frame.text.strip()
        if "감사합니다" in t or "Thank" in t or "THANK" in t:
            set_runs_text(shape.text_frame, message)
            return []
        if not title_done and "메인 타이틀" in t:
            set_runs_text(shape.text_frame, message)
            title_done = True
        elif not sub_done and "서브 타이틀" in t:
            set_runs_text(shape.text_frame, " ")
            sub_done = True
    return []


INJECTORS = {
    "cover": inject_cover,
    "agenda": inject_agenda,
    "section": inject_section,
    "kpi": inject_kpi,
    "process": inject_process,
    "matrix": inject_matrix,
    "table": inject_table,
    "image": inject_image,
    "content": inject_content,
    "closing": inject_closing,
}


def duplicate_slide(prs, source_slide):
    """Append a deep copy of `source_slide` to `prs`. Returns the new slide."""
    blank_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(blank_layout)

    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    spTree = new_slide.shapes._spTree
    for child in source_slide.shapes._spTree.iterchildren():
        tag = child.tag
        if tag.endswith("}nvGrpSpPr") or tag.endswith("}grpSpPr"):
            continue
        spTree.append(copy.deepcopy(child))

    for rel in source_slide.part.rels.values():
        if "notesSlide" in rel.reltype:
            continue
        if rel.is_external:
            new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)

    return new_slide


def drop_slide(prs, slide) -> None:
    slide_id = slide.slide_id
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        if int(sldId.get("id")) == slide_id:
            rId = sldId.get(qn("r:id"))
            prs.part.drop_rel(rId)
            sldIdLst.remove(sldId)
            return


def assemble(template_path: Path, plan_path: Path, output_path: Path) -> dict:
    plan = json.loads(plan_path.read_text())
    prs = Presentation(str(template_path))

    original_slides = list(prs.slides)
    source_by_index = {i + 1: slide for i, slide in enumerate(original_slides)}

    report = []
    for entry in plan["slides"]:
        source_idx = entry["source_index"]
        kind = entry["kind"]
        content = entry.get("content", {})

        source_slide = source_by_index.get(source_idx)
        if source_slide is None:
            report.append({"position": entry["position"], "kind": kind, "error": f"missing source slide {source_idx}"})
            continue

        new_slide = duplicate_slide(prs, source_slide)
        injector = INJECTORS.get(kind, inject_content)
        issues = injector(new_slide, content)
        if issues:
            report.append(
                {"position": entry["position"], "source_index": source_idx, "kind": kind, "issues": issues}
            )

    for slide in original_slides:
        drop_slide(prs, slide)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return {"output": str(output_path), "slide_count": len(plan["slides"]), "issues": report}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--plan", required=True)
    ap.add_argument("--output", required=True)
    args = ap.parse_args()

    template = Path(args.template).expanduser().resolve()
    plan = Path(args.plan).expanduser().resolve()
    output = Path(args.output).expanduser().resolve()

    if not template.exists():
        sys.exit(f"template not found: {template}")
    if not plan.exists():
        sys.exit(f"plan not found: {plan}")

    result = assemble(template, plan, output)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    import sys  # noqa
    main()
