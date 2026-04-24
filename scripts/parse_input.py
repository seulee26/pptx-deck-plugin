"""Parse an Excel/Word/PDF input into the outline.json schema consumed by slide-matcher.

Conventions (xlsx):
  - optional sheet `meta` with key/value rows: title, subtitle, company, closing
  - every other sheet = one section; sheet name is the heading
  - rows with first cell in {KPI, PROCESS, MATRIX, TABLE, IMAGE, CONTENT} become items
  - otherwise fall back to `content` items with the row text joined
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

ITEM_TYPES = {"kpi", "process", "matrix", "table", "image", "content"}


def parse_xlsx(path: Path) -> dict:
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True)
    meta: dict[str, str] = {}
    sections: list[dict] = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = [[("" if c is None else str(c).strip()) for c in row] for row in ws.iter_rows(values_only=True)]
        rows = [r for r in rows if any(cell for cell in r)]
        if not rows:
            continue

        if sheet_name.lower() == "meta":
            for row in rows:
                if len(row) >= 2 and row[0]:
                    meta[row[0].lower()] = row[1]
            continue

        items: list[dict] = []
        i = 0
        while i < len(rows):
            row = rows[i]
            marker = row[0].upper() if row else ""
            if marker.lower() in ITEM_TYPES:
                kind = marker.lower()
                title = row[1] if len(row) > 1 else ""
                body = row[2] if len(row) > 2 else ""
                data: dict = {}

                if kind == "kpi":
                    data = {"value": row[3] if len(row) > 3 else "", "delta": row[4] if len(row) > 4 else ""}
                elif kind == "process":
                    data = {"steps": [c for c in row[3:] if c]}
                elif kind == "matrix":
                    data = {
                        "S": row[3] if len(row) > 3 else "",
                        "W": row[4] if len(row) > 4 else "",
                        "O": row[5] if len(row) > 5 else "",
                        "T": row[6] if len(row) > 6 else "",
                    }
                elif kind == "table":
                    headers = row[3:] if len(row) > 3 else []
                    tbl_rows: list[list[str]] = []
                    j = i + 1
                    while j < len(rows) and rows[j] and rows[j][0].upper().lower() not in ITEM_TYPES:
                        tbl_rows.append([c for c in rows[j] if c])
                        j += 1
                    data = {"headers": [h for h in headers if h], "rows": tbl_rows}
                    items.append({"type": kind, "title": title, "body": body, "data": data})
                    i = j
                    continue

                items.append({"type": kind, "title": title, "body": body, "data": data})
            else:
                items.append({"type": "content", "title": row[0], "body": " ".join(row[1:]).strip(), "data": {}})
            i += 1

        sections.append({"heading": sheet_name, "items": items})

    return {
        "title": meta.get("title", path.stem),
        "subtitle": meta.get("subtitle", ""),
        "company": meta.get("company", ""),
        "sections": sections,
        "closing": meta.get("closing", "Thank you"),
    }


def parse_docx(path: Path) -> dict:
    from docx import Document

    doc = Document(path)
    title = ""
    sections: list[dict] = []
    current_section: dict | None = None
    current_item: dict | None = None

    def flush_item() -> None:
        nonlocal current_item
        if current_item and current_section is not None:
            current_section["items"].append(current_item)
        current_item = None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = (para.style.name or "").lower()
        if "heading 1" in style and not title:
            title = text
        elif "heading 2" in style:
            flush_item()
            current_section = {"heading": text, "items": []}
            sections.append(current_section)
        elif "heading 3" in style:
            flush_item()
            current_item = {"type": "content", "title": text, "body": "", "data": {}}
        else:
            if current_section is None:
                current_section = {"heading": "Overview", "items": []}
                sections.append(current_section)
            if current_item is None:
                current_item = {"type": "content", "title": "", "body": "", "data": {}}
            current_item["body"] = (current_item["body"] + "\n" + text).strip()

            stripped = text.replace(",", "").replace("%", "")
            if any(ch.isdigit() for ch in text) and len(text) < 40 and any(s in text for s in ["%", "K", "M", "B"]):
                current_item["type"] = "kpi"
                current_item["data"] = {"value": text, "delta": ""}
    flush_item()

    for tbl in doc.tables:
        rows = [[cell.text.strip() for cell in row.cells] for row in tbl.rows]
        if not rows:
            continue
        item = {
            "type": "table",
            "title": rows[0][0] if rows[0] else "Table",
            "body": "",
            "data": {"headers": rows[0], "rows": rows[1:]},
        }
        if not sections:
            sections.append({"heading": "Data", "items": []})
        sections[-1]["items"].append(item)

    return {
        "title": title or path.stem,
        "subtitle": "",
        "company": "",
        "sections": sections,
        "closing": "Thank you",
    }


def parse_pdf(path: Path) -> dict:
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    pages = [page.extract_text() or "" for page in reader.pages]
    title = path.stem
    if pages:
        first_line = next((ln.strip() for ln in pages[0].splitlines() if ln.strip()), "")
        if first_line:
            title = first_line[:80]

    sections = []
    for idx, text in enumerate(pages[1:], start=1):
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        if not lines:
            continue
        heading = lines[0][:80]
        body = "\n".join(lines[1:])[:800]
        sections.append(
            {
                "heading": heading,
                "items": [{"type": "content", "title": heading, "body": body, "data": {}}],
            }
        )

    if not sections:
        sections = [
            {
                "heading": "Content",
                "items": [{"type": "content", "title": title, "body": "\n".join(pages)[:800], "data": {}}],
            }
        ]

    return {"title": title, "subtitle": "", "company": "", "sections": sections, "closing": "Thank you"}


def parse(path: Path) -> dict:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return parse_xlsx(path)
    if suffix == ".docx":
        return parse_docx(path)
    if suffix == ".pdf":
        return parse_pdf(path)
    raise ValueError(f"unsupported input type: {suffix}")


def main() -> None:
    if len(sys.argv) != 3:
        sys.exit("usage: parse_input.py <input> <outline.json>")
    src = Path(sys.argv[1]).expanduser().resolve()
    dst = Path(sys.argv[2]).expanduser().resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)

    try:
        outline = parse(src)
    except Exception as e:
        outline = {
            "title": src.stem,
            "subtitle": "",
            "company": "",
            "sections": [
                {
                    "heading": "Content",
                    "items": [{"type": "content", "title": src.stem, "body": f"(parse failed: {e})", "data": {}}],
                }
            ],
            "closing": "Thank you",
        }

    dst.write_text(json.dumps(outline, ensure_ascii=False, indent=2))
    n_sections = len(outline["sections"])
    n_items = sum(len(s["items"]) for s in outline["sections"])
    print(f"outline written: {dst}")
    print(f"  sections: {n_sections}  items: {n_items}")


if __name__ == "__main__":
    main()
