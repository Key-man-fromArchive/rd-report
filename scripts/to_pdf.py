#!/usr/bin/env python3
"""to_pdf.py — Word(.docx)를 PDF로 바꾸고, 원하면 쪽 미리보기 PNG를 만든다.

ONLYOFFICE Document Server 이미지의 docbuilder 를 **일회용 컨테이너**로 돌린다.
운영 중인 ONLYOFFICE 컨테이너에는 아무것도 넣지 않는다.

    python3 scripts/to_pdf.py output/보고서.docx                  # → output/보고서.pdf
    python3 scripts/to_pdf.py output/보고서.docx --png output/preview --pages 1,2,5

필요한 것: docker, 이미지 onlyoffice/documentserver (RD_ONLYOFFICE_IMAGE 로 바꿀 수 있음),
미리보기 PNG는 pypdfium2. 한글 글꼴은 ~/.local/share/rd-report-fonts (RD_REPORT_FONTS)의
TrueType 글꼴을 컨테이너에 붙여 쓴다(scripts/setup_fonts.sh 로 준비) — 문서 글꼴(맑은 고딕·바탕)이 없으면 가장 비슷한 글꼴로 대체되므로, 정확한 모양은
받는 사람 PC의 Word에서 확인한다.

목차(rd_docx.js "toc")가 있으면 한 번 변환한 PDF 에서 각 장이 실린 쪽을 찾아 Word 파일의
PAGEREF 필드 값에 써 넣고(입력 .docx 를 그 자리에서 고침) 다시 변환한다. Word 에서 고친 뒤에는
목차에서 [필드 업데이트](F9)로 다시 맞출 수 있다.

docker 가 없으면 Word에서 [파일 → 다른 이름으로 저장 → PDF] 로 만든다.
"""
import argparse
import html
import re
import os
import shutil
import subprocess
import sys
import uuid
import zipfile
from pathlib import Path

IMAGE = os.environ.get("RD_ONLYOFFICE_IMAGE", "onlyoffice/documentserver:8.2.2")
BIN = "/var/www/onlyoffice/documentserver/server/FileConverter/bin"
# docker 데몬이 볼 수 있는 곳(홈)에 작업 폴더를 둔다. /tmp 는 샌드박스에서 데몬에 안 보일 수 있다.
CACHE = Path.home() / ".cache" / "rd-report"
#: 변환용 한글 글꼴. TrueType(.ttf)이어야 PDF에 쓴 글자만 들어간다 — OTF(CFF) CJK 글꼴은
#: 통째로 들어가 3쪽짜리 PDF가 40MB를 넘는다. scripts/setup_fonts.sh 가 나눔고딕·나눔명조를 받는다.
FONTS = Path(os.environ.get("RD_REPORT_FONTS", Path.home() / ".local" / "share" / "rd-report-fonts"))
#: PDF 변환 때만 쓰는 글꼴 바꿔치기. 맑은 고딕·바탕은 변환 서버에 없어 엉뚱한 글꼴(본문이 고딕,
#: 영문이 Trebuchet)로 대체되므로, 변환용 사본에서 이름을 비슷한 무료 글꼴로 바꾼다. Word 파일은
#: 그대로이고, PDF 에는 실제로 쓴 나눔 글꼴이 그 이름 그대로 들어간다.
PDF_FONT_MAP = {
    "맑은 고딕": "NanumGothic", "Malgun Gothic": "NanumGothic",
    "바탕": "NanumMyeongjo", "Batang": "NanumMyeongjo",
}


def _pdf_copy(src: Path, dst: Path) -> None:
    """src 의 사본을 dst 에 쓰되, 본문·스타일·머리글·바닥글의 글꼴 이름을 PDF_FONT_MAP 대로 바꾼다."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                for old, new in PDF_FONT_MAP.items():
                    text = text.replace(f'"{old}"', f'"{new}"')
                data = text.encode("utf-8")
            zout.writestr(item, data)


def convert(docx: Path, pdf: Path) -> None:
    if not shutil.which("docker"):
        raise SystemExit("docker 가 없습니다. Word에서 'PDF로 저장'을 쓰세요.")
    work = CACHE / uuid.uuid4().hex
    work.mkdir(parents=True)
    try:
        _pdf_copy(docx, work / "in.docx")
        (work / "conv.js").write_text(
            'builder.OpenFile("/data/in.docx");\n'
            'builder.SaveFile("pdf", "/data/out.pdf");\n'
            "builder.CloseFile();\n", encoding="utf-8")
        cmd = ["docker", "run", "--rm", "-v", f"{work}:/data"]
        if FONTS.is_dir():
            cmd += ["-v", f"{FONTS}:/usr/share/fonts/truetype/rd-fonts:ro"]
        cmd += ["--entrypoint", "sh", IMAGE, "-c",
                "/usr/bin/documentserver-generate-allfonts.sh true >/dev/null 2>&1; "
                f"cd {BIN} && ./docbuilder /data/conv.js"]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        out = work / "out.pdf"
        if not out.is_file():
            raise SystemExit(f"PDF 변환 실패:\n{res.stdout[-2000:]}\n{res.stderr[-2000:]}")
        pdf.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(out, pdf)
    finally:
        shutil.rmtree(work, ignore_errors=True)


TOC_FIELD = re.compile(r'(<w:fldSimple w:instr="PAGEREF (_Toc_rd\d+) \\h"[^>]*>.*?<w:t[^>]*>)([^<]*)(</w:t>)', re.S)
TOC_MARK = re.compile(r'<w:bookmarkStart [^>]*w:name="(_Toc_rd\d+)"[^>]*/>(.*?)</w:p>', re.S)


def _squash(t: str) -> str:
    return re.sub(r"\s+", "", t)


def toc_targets(docx: Path) -> list[tuple[str, str]]:
    """목차 북마크 (이름, 제목 글자) — 문서 순서."""
    with zipfile.ZipFile(docx) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    if "PAGEREF _Toc_rd" not in xml:
        return []
    return [(name, html.unescape("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", body))))
            for name, body in TOC_MARK.findall(xml)]


def find_pages(pdf: Path, targets: list[tuple[str, str]]) -> dict[str, int]:
    """제목이 한 줄 전체로 나오는 첫 쪽을 문서 순서대로 앞에서부터 찾는다.
    목차 줄은 점선·쪽 번호가 붙어 한 줄 전체와 같지 않으므로 목차 쪽에는 걸리지 않는다."""
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(str(pdf))
    lines = [{_squash(l) for l in doc[i].get_textpage().get_text_range().splitlines()} for i in range(len(doc))]
    pages, start = {}, 0
    for name, text in targets:
        key = _squash(text)
        for i in range(start, len(lines)):
            if key in lines[i]:
                pages[name] = i + 1
                start = i
                break
        else:
            print(f"목차: 「{text}」 쪽을 PDF 에서 찾지 못함 — 「-」로 둠", file=sys.stderr)
    return pages


def write_toc_pages(docx: Path, pages: dict[str, int]) -> bool:
    """PAGEREF 필드의 표시값을 pages 로 바꾼다. 바뀐 것이 있으면 True."""
    tmp = docx.with_suffix(".toc.tmp")
    changed = False
    with zipfile.ZipFile(docx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                def sub(m):
                    nonlocal changed
                    new = str(pages.get(m.group(2), "-"))
                    changed |= new != m.group(3)
                    return m.group(1) + new + m.group(4)
                data = TOC_FIELD.sub(sub, data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(docx)
    return changed


def render_png(pdf: Path, out_dir: Path, pages: list[int] | None, scale: float) -> list[Path]:
    try:
        import pypdfium2 as pdfium
    except ImportError:
        raise SystemExit("미리보기에는 pypdfium2 가 필요합니다 (pip install pypdfium2).")
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = pdfium.PdfDocument(str(pdf))
    wanted = pages or list(range(1, len(doc) + 1))
    files = []
    for n in wanted:
        if 1 <= n <= len(doc):
            f = out_dir / f"{pdf.stem}_p{n:02d}.png"
            doc[n - 1].render(scale=scale).to_pil().save(f)
            files.append(f)
    return files


def main() -> None:
    ap = argparse.ArgumentParser(description="DOCX → PDF (ONLYOFFICE 일회용 컨테이너) + 쪽 미리보기")
    ap.add_argument("docx", type=Path)
    ap.add_argument("pdf", type=Path, nargs="?")
    ap.add_argument("--png", type=Path, help="쪽 미리보기 PNG 를 저장할 폴더")
    ap.add_argument("--pages", help="미리보기 쪽 번호 (예: 1,2,5). 생략 시 전체")
    ap.add_argument("--scale", type=float, default=1.0)
    a = ap.parse_args()
    pdf = a.pdf or a.docx.with_suffix(".pdf")
    convert(a.docx.resolve(), pdf)
    targets = toc_targets(a.docx)
    if targets:
        # 쪽 번호 글자 수가 바뀌어도 목차 줄 수는 그대로라 두 번째 변환에서 쪽이 밀리지 않는다
        if write_toc_pages(a.docx, find_pages(pdf, targets)):
            convert(a.docx.resolve(), pdf)
        print(f"목차: {len(targets)}개 항목 쪽 번호 기록 ({a.docx})")
    print(f"PDF: {pdf}")
    if a.png:
        pages = [int(x) for x in a.pages.split(",")] if a.pages else None
        for f in render_png(pdf, a.png, pages, a.scale):
            print(f"PNG: {f}")


if __name__ == "__main__":
    main()
