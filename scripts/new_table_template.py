#!/usr/bin/env python3
"""Create a table template (templates/tables/<name>.xml) from a column spec — no Hancom needed.

The cells are cloned from ``basic.xml``, so the new template uses exactly the
report's standard table styles (header 26/18/9, body 27/18·22/10, summary 26/18/9).
For a table that needs its own look, design it in Hancom and use
``extract_table_templates.py`` instead.

Usage:
    python new_table_template.py achievement "목표 달성도 (추진목표/달성내용/달성도)" \\
        "추진목표:14000:justify" "달성내용:27849:justify" "달성도(%):6000:center"
    python new_table_template.py sample_plan "실험 시료·반응 계획" --summary \\
        "구분:9000:center" "대상:12849:justify" "시료 수:7000:center" "반응 수:7000:center" "비고:12000:justify"

Column widths should add up to the report table width (47849).
"""
import argparse
import copy
from pathlib import Path

from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parent
TABLES = SCRIPT_DIR.parent / "templates" / "tables"
HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
TABLE_WIDTH = 47849
#: paraPr for a body cell by alignment (header.xml: 18 = CENTER, 22 = JUSTIFY)
ALIGN_PARA = {"center": "18", "justify": "22"}


def make(name, description, cols, summary=False):
    src = etree.parse(str(TABLES / "basic.xml"))
    root = copy.deepcopy(src.getroot())
    meta = root.find("meta")
    meta.find("name").text = name
    meta.find("description").text = description
    total = sum(w for _, w, _ in cols)
    meta.find("table-width").text = str(total)
    cols_el = meta.find("columns")
    for c in list(cols_el):
        cols_el.remove(c)
    for i, (_, w, a) in enumerate(cols):
        etree.SubElement(cols_el, "col", index=str(i), width=str(w), align=a)

    tbl = root.find(f"sample/{{{HP}}}tbl")
    tbl.set("colCnt", str(len(cols)))
    tbl.find(f"{{{HP}}}sz").set("width", str(total))
    rows = tbl.findall(f"{{{HP}}}tr")
    head_cell = copy.deepcopy(rows[0][0])
    body_cell = copy.deepcopy(rows[1][0])
    for tr in rows:
        tbl.remove(tr)

    def row(kind, cell, texts):
        tr = etree.SubElement(tbl, f"{{{HP}}}tr")
        tr.set("row-type", kind)
        for i, ((label, w, a), text) in enumerate(zip(cols, texts)):
            tc = copy.deepcopy(cell)
            tc.find(f"{{{HP}}}cellAddr").set("colAddr", str(i))
            tc.find(f"{{{HP}}}cellSz").set("width", str(w))
            p = tc.find(f".//{{{HP}}}p")
            if kind == "body":
                p.set("paraPrIDRef", ALIGN_PARA[a])
            t = tc.find(f".//{{{HP}}}t")
            t.text = text or None
            tr.append(tc)
        return tr

    row("header", head_cell, [c[0] for c in cols])
    row("body", body_cell, [""] * len(cols))
    if summary:
        s = row("summary", head_cell, ["합  계"] + [""] * (len(cols) - 1))
        for tc in s:
            tc.find(f"{{{HP}}}cellAddr").set("rowAddr", "__ROW_ADDR__")
    out = TABLES / f"{name}.xml"
    etree.ElementTree(root).write(str(out), xml_declaration=True, encoding="UTF-8", pretty_print=True)
    if total != TABLE_WIDTH:
        print(f"note: width {total} != report table width {TABLE_WIDTH}")
    print(f"wrote {out}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("name")
    ap.add_argument("description")
    ap.add_argument("columns", nargs="+", help='"헤더:너비:center|justify"')
    ap.add_argument("--summary", action="store_true", help="add a 합계 row (header style)")
    a = ap.parse_args()
    cols = []
    for spec in a.columns:
        label, w, align = spec.rsplit(":", 2)
        if align not in ALIGN_PARA:
            ap.error(f"align must be center or justify: {spec}")
        cols.append((label, int(w), align))
    make(a.name, a.description, cols, a.summary)


if __name__ == "__main__":
    main()
