#!/usr/bin/env python3
"""to_pdf.py — Word(.docx)를 PDF로 바꾸고, 원하면 쪽 미리보기 PNG를 만든다.

ONLYOFFICE Document Server 이미지의 docbuilder 를 **일회용 컨테이너**로 돌린다.
운영 중인 ONLYOFFICE 컨테이너에는 아무것도 넣지 않는다.

    python3 scripts/to_pdf.py output/보고서.docx                  # → output/보고서.pdf
    python3 scripts/to_pdf.py output/보고서.docx --png output/preview --pages 1,2,5

필요한 것: docker, 이미지 onlyoffice/documentserver (RD_ONLYOFFICE_IMAGE 로 바꿀 수 있음),
미리보기 PNG는 pypdfium2. 한글 글꼴은 ~/.local/share/fonts 의 글꼴을 컨테이너에 붙여
쓴다 — 문서 글꼴(맑은 고딕·바탕)이 없으면 가장 비슷한 글꼴로 대체되므로, 정확한 모양은
받는 사람 PC의 Word에서 확인한다.

docker 가 없으면 Word에서 [파일 → 다른 이름으로 저장 → PDF] 로 만든다.
"""
import argparse
import os
import shutil
import subprocess
import sys
import uuid
from pathlib import Path

IMAGE = os.environ.get("RD_ONLYOFFICE_IMAGE", "onlyoffice/documentserver:8.2.2")
BIN = "/var/www/onlyoffice/documentserver/server/FileConverter/bin"
# docker 데몬이 볼 수 있는 곳(홈)에 작업 폴더를 둔다. /tmp 는 샌드박스에서 데몬에 안 보일 수 있다.
CACHE = Path.home() / ".cache" / "rd-report"
FONTS = Path.home() / ".local" / "share" / "fonts"


def convert(docx: Path, pdf: Path) -> None:
    if not shutil.which("docker"):
        raise SystemExit("docker 가 없습니다. Word에서 'PDF로 저장'을 쓰세요.")
    work = CACHE / uuid.uuid4().hex
    work.mkdir(parents=True)
    try:
        shutil.copy2(docx, work / "in.docx")
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
    print(f"PDF: {pdf}")
    if a.png:
        pages = [int(x) for x in a.pages.split(",")] if a.pages else None
        for f in render_png(pdf, a.png, pages, a.scale):
            print(f"PNG: {f}")


if __name__ == "__main__":
    main()
