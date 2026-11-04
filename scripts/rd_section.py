#!/usr/bin/env python3
"""rd_section.py — 보고서 본문을 XML 대신 함수 호출로 쌓는 작성 도구.

report 양식(templates/report)의 스타일 ID를 이 파일이 알고 있으므로, 쓰는 쪽은
내용만 넘기면 된다. 결과는 build_hwpx.py 와 같은 경로로 조립·검증된다.

    import sys; sys.path.insert(0, f"{SKILL_DIR}/scripts")
    from rd_section import Report

    r = Report(title="소나무 마커 검증 결과보고서")          # 머리말 기관명은 config/company.json
    r.heading("1. 연구개발과제의 개요")
    r.box("NTC 제외 재판정으로 사용 가능 마커 40개 확보")     # □
    r.item("76개 중 40개는 칩 유전형과 10개 시료 중 9개 이상 일치")  # ❍
    r.dash("나머지 36개는 조건부 12, 재설계 12, Sanger 확인 12")  # -
    r.note("NTC: 주형 없는 음성 대조")                       # ※
    r.table("achievement", [["사용 마커 확보", "40개 바로 사용", "53"]])
    r.image("fig/scatter.png", mm=150, caption="그림 1. 마커별 산점도")
    r.page_break()
    r.save("output/결과보고서.hwpx", page_number=True)

스타일 ID 근거: SKILL.md 「report 템플릿 스타일 ID 맵」.
"""
from __future__ import annotations

import json
import sys
import tempfile
from datetime import date
from pathlib import Path

from lxml import etree

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

from build_hwpx import build  # noqa: E402
from table_builder import build_table_paragraph  # noqa: E402

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
SECTION_TEMPLATE = SKILL_DIR / "templates" / "report" / "section0.xml"
COMPANY_FILE = SKILL_DIR / "config" / "company.json"

#: (charPrIDRef, paraPrIDRef) — templates/report/header.xml
HEADING = ("10", "2")        # 1. 2. 소제목, 16pt HY헤드라인M
HEADING_GAP = ("12", "2")    # 소제목 아래 4pt 간격
BODY = ("25", "22")          # □ ❍ - 본문, 13pt 휴먼명조
NOTE = ("20", "21")          # ※ 참고, 13pt 한양중고딕
SECTION_GAP = ("13", "2")    # 섹션 사이 빈 줄, 11pt
CAPTION = ("27", "1")        # 그림·표 제목, 12pt 휴먼명조, 가운데
PICTURE_PARA = ("0", "1")    # [[IMG:…]] 자리 — build_hwpx 가 그림으로 바꾼다


def company() -> dict:
    """config/company.json (없으면 빈 값). 머리말 기관명·작성자 기본값."""
    if COMPANY_FILE.is_file():
        return json.loads(COMPANY_FILE.read_text(encoding="utf-8"))
    return {}


class Report:
    """report 양식 위에 본문 문단을 순서대로 쌓는다."""

    def __init__(self, title: str, org: str | None = None, on: date | None = None):
        cfg = company()
        self.title = title
        self.org = org or cfg.get("report_header") or "부서명"
        self.creator = cfg.get("creator") or self.org
        self.on = on or date.today()
        self.root = etree.fromstring(SECTION_TEMPLATE.read_bytes())
        # 양식의 머리말·제목 도형·제목 아래 빈 줄(앞 3문단)만 남긴다.
        for p in self.root.findall(f"{{{HP}}}p")[3:]:
            self.root.remove(p)
        self._fill_placeholders()
        self._next_id = 1000001000
        self._table_id = 1100000000

    # -- 머리말 -------------------------------------------------------------
    def _fill_placeholders(self) -> None:
        yy = f"{self.on.year % 100:02d}"
        md = f"{self.on.month}. {self.on.day}."
        for t in self.root.iter(f"{{{HP}}}t"):
            if t.text == "YY":
                t.text = yy
            elif t.text == "보고자료 제목":
                t.text = self.title
            elif t.text and ". MM. DD., 부서명)" in t.text:
                t.text = t.text.replace(". MM. DD., 부서명)", f". {md}, {self.org})")

    # -- 문단 ---------------------------------------------------------------
    def _p(self, text: str | None, style, page_break: bool = False):
        cp, pp = style
        self._next_id += 1
        p = etree.SubElement(self.root, f"{{{HP}}}p", id=str(self._next_id), paraPrIDRef=pp,
                             styleIDRef="0", pageBreak="1" if page_break else "0",
                             columnBreak="0", merged="0")
        run = etree.SubElement(p, f"{{{HP}}}run", charPrIDRef=cp)
        t = etree.SubElement(run, f"{{{HP}}}t")
        if text:
            t.text = text
        return p

    def heading(self, text: str, new_page: bool = False) -> None:
        """「1. 연구개발과제의 개요」 같은 소제목. 앞 섹션과는 빈 줄로 띄운다."""
        if len(self.root) > 3 and not new_page:
            self._p(None, SECTION_GAP)
        self._p(text, HEADING, page_break=new_page)
        self._p(None, HEADING_GAP)

    def box(self, text: str) -> None:
        """□ 상위 항목 (들여쓰기 없음)."""
        self._p("□ " + text.lstrip("□ ").strip(), BODY)

    def item(self, text: str) -> None:
        """❍ 하위 항목 (공백 2칸 들여쓰기)."""
        self._p("  ❍ " + text.lstrip("❍○ ").strip(), BODY)

    def dash(self, text: str) -> None:
        """- 세부 항목 (공백 4칸 들여쓰기)."""
        self._p("    - " + text.lstrip("- ").strip(), BODY)

    def note(self, text: str) -> None:
        """※ 참고·주석."""
        self._p("※ " + text.lstrip("※ ").strip(), NOTE)

    def text(self, text: str) -> None:
        """기호 없는 본문 줄."""
        self._p(text, BODY)

    def blank(self) -> None:
        self._p(None, SECTION_GAP)

    def page_break(self) -> None:
        """다음 문단을 새 쪽에서 시작한다."""
        self._p(None, SECTION_GAP, page_break=True)

    def caption(self, text: str) -> None:
        self._p(text, CAPTION)

    # -- 표·그림 ------------------------------------------------------------
    def table(self, template: str, data, headers=None, summary=None, caption: str | None = None) -> None:
        """templates/tables/<template>.xml 표. 빈 셀은 "-" 로 채운다(양식 규칙)."""
        if caption:
            self.caption(caption)
        data = [["-" if (c is None or str(c) == "") else str(c) for c in row] for row in data]
        xml = build_table_paragraph(template, data=data, headers=headers, summary=summary,
                                    start_id=self._table_id)
        self._table_id += 10000
        self.root.append(etree.fromstring(xml.encode("utf-8")))

    def image(self, path: str | Path, mm: float = 150, caption: str | None = None) -> None:
        """그림. 너비 mm(본문폭 상한), 높이는 비율대로. 제목은 그림 아래."""
        path = Path(path).expanduser().resolve()
        if not path.is_file():
            raise FileNotFoundError(path)
        self._p(f"[[IMG:{path}|{mm:g}]]", PICTURE_PARA)
        if caption:
            self.caption(caption)

    # -- 저장 ---------------------------------------------------------------
    def to_xml(self) -> bytes:
        return etree.tostring(self.root, xml_declaration=True, encoding="UTF-8", standalone=True)

    def save(self, output: str | Path, page_number: bool = True, creator: str | None = None) -> Path:
        output = Path(output)
        with tempfile.NamedTemporaryFile("wb", suffix=".xml", delete=False) as fh:
            fh.write(self.to_xml())
            section = Path(fh.name)
        try:
            build(template="report", header_override=None, section_override=section,
                  title=self.title, creator=creator or self.creator, output=output,
                  page_number=page_number)
        finally:
            section.unlink(missing_ok=True)
        return output


if __name__ == "__main__":
    print(__doc__)
