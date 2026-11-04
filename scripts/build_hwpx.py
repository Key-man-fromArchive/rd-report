#!/usr/bin/env python3
"""Build an HWPX document from the report template and optional XML overrides.

Assembles a valid HWPX file by:
1. Copying the internal skeleton (mimetype, META-INF, Preview, content.hpf, …)
2. Overlaying the report template's header.xml and section0.xml
3. Optionally overriding header.xml and/or section0.xml with custom files
4. Optionally setting metadata (title, creator)
5. Validating XML well-formedness
6. Packaging as HWPX (ZIP with mimetype first, ZIP_STORED)

Usage:
    # Default report document (no overrides)
    python build_hwpx.py --output result.hwpx

    # Custom section XML override
    python build_hwpx.py --section my_section0.xml --output result.hwpx

    # Custom header and section
    python build_hwpx.py --header my_header.xml --section my_section0.xml --output result.hwpx

    # With metadata
    python build_hwpx.py --section my.xml --title "제목" --creator "작성자" --output result.hwpx

    # Page numbers "- 1 -" at the bottom centre (공문서 관행)
    python build_hwpx.py --section my.xml --page-number --output result.hwpx

Pictures: a paragraph whose only text is ``[[IMG:path/to/fig.png|150]]`` is replaced
by that picture, centred, 150 mm wide (height follows the aspect ratio; the width is
capped at the body width). A relative path is resolved against the section file's
folder first, then the current directory. PNG and JPEG only.
"""

import argparse
import re
import shutil
import struct
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

from lxml import etree

# Hardened parser: no external entity resolution, no DTD loading, no network
# access, bounded tree size — defends against XXE / billion-laughs payloads
# in supplied XML overrides.
_SAFE_XML_PARSER = etree.XMLParser(
    resolve_entities=False, no_network=True, load_dtd=False, huge_tree=False
)

# Resolve paths relative to this script
SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
TEMPLATES_DIR = SKILL_DIR / "templates"
BASE_DIR = TEMPLATES_DIR / "base"

HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HC_NS = "http://www.hancom.co.kr/hwpml/2011/core"
HH_NS = "http://www.hancom.co.kr/hwpml/2011/head"
OPF_NS = "http://www.idpf.org/2007/opf/"

#: A paragraph whose whole text is this becomes a picture.
IMG_PLACEHOLDER = re.compile(r"^\[\[IMG:(?P<path>[^|\]]+)(?:\|(?P<mm>\d+(?:\.\d+)?))?\]\]$")
#: ext -> (Format attribute, media type)
IMAGE_TYPES = {".png": ("png", "image/png"), ".jpg": ("jpg", "image/jpeg"),
               ".jpeg": ("jpg", "image/jpeg")}
HWPUNIT_PER_MM = 283.465
#: report 양식 본문폭 (표 템플릿과 같은 값). A picture is never wider than this.
BODY_WIDTH = 48190
DEFAULT_IMG_MM = 150
#: report header.xml paraPr 1 = CENTER — the picture paragraph's alignment.
CENTER_PARA_PR = "1"

AVAILABLE_TEMPLATES = ["report"]
DEFAULT_TEMPLATE = "report"
EXCLUDED_PACKAGE_NAMES = {".DS_Store"}
EXCLUDED_PACKAGE_PARTS = {"__MACOSX", "__pycache__"}
EXCLUDED_PACKAGE_SUFFIXES = {".pyc", ".pyo"}


def should_package_file(rel_path: str) -> bool:
    """Return False for local cache/metadata files that do not belong in HWPX."""
    path = Path(rel_path)
    if path.name in EXCLUDED_PACKAGE_NAMES:
        return False
    if any(part in EXCLUDED_PACKAGE_PARTS for part in path.parts):
        return False
    return path.suffix not in EXCLUDED_PACKAGE_SUFFIXES


def validate_xml(filepath: Path) -> None:
    """Check that an XML file is well-formed. Raises on error."""
    try:
        etree.parse(str(filepath), parser=_SAFE_XML_PARSER)
    except etree.XMLSyntaxError as e:
        raise SystemExit(f"Malformed XML in {filepath.name}: {e}")


def update_metadata(content_hpf: Path, title: str | None, creator: str | None) -> None:
    """Update title and/or creator in content.hpf."""
    if not title and not creator:
        return

    tree = etree.parse(str(content_hpf), parser=_SAFE_XML_PARSER)
    root = tree.getroot()
    ns = {"opf": "http://www.idpf.org/2007/opf/"}

    if title:
        title_el = root.find(".//opf:title", ns)
        if title_el is not None:
            title_el.text = title

    now = datetime.now(timezone.utc)
    iso_now = now.strftime("%Y-%m-%dT%H:%M:%SZ")

    for meta in root.findall(".//opf:meta", ns):
        name = meta.get("name", "")
        if creator and name == "creator":
            meta.text = creator
        elif creator and name == "lastsaveby":
            meta.text = creator
        elif name == "CreatedDate":
            meta.text = iso_now
        elif name == "ModifiedDate":
            meta.text = iso_now
        elif name == "date":
            meta.text = now.strftime("%Y년 %m월 %d일")

    etree.indent(root, space="  ")
    tree.write(
        str(content_hpf),
        pretty_print=True,
        xml_declaration=True,
        encoding="UTF-8",
    )


def image_size(data: bytes, ext: str) -> tuple[int, int]:
    """(width, height) in pixels, read from the PNG/JPEG header — no Pillow needed."""
    if ext == ".png":
        if data[:8] != b"\x89PNG\r\n\x1a\n":
            raise SystemExit("Not a PNG file")
        return struct.unpack(">II", data[16:24])
    i = 2  # JPEG: walk the markers to the first SOFn
    while i + 9 < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            h, w = struct.unpack(">HH", data[i + 5:i + 9])
            return w, h
        i += 2 + struct.unpack(">H", data[i + 2:i + 4])[0]
    raise SystemExit("Could not read JPEG size")


def _resolve_image(raw: str, search_dirs: list[Path]) -> Path:
    p = Path(raw.strip()).expanduser()
    candidates = [p] if p.is_absolute() else [d / p for d in search_dirs]
    for c in candidates:
        if c.is_file():
            if c.suffix.lower() not in IMAGE_TYPES:
                raise SystemExit(f"Unsupported image type (PNG/JPEG only): {c}")
            return c
    raise SystemExit(f"Image not found: {raw}")


def _pic_element(bin_id: str, width: int, height: int, pic_id: int):
    """An inline (treat-as-char) picture, the shape Hancom itself writes."""
    hp, hc = f"{{{HP_NS}}}", f"{{{HC_NS}}}"
    pic = etree.Element(hp + "pic", {
        "textWrap": "SQUARE", "textFlow": "BOTH_SIDES", "reverse": "0", "id": str(pic_id),
        "zOrder": "0", "numberingType": "PICTURE", "lock": "0", "dropcapstyle": "None",
        "href": "", "groupLevel": "0", "instid": str(pic_id)})
    sub = lambda parent, tag, **a: etree.SubElement(parent, tag, {k: str(v) for k, v in a.items()})
    sub(pic, hp + "offset", x=0, y=0)
    sub(pic, hp + "orgSz", width=width, height=height)
    sub(pic, hp + "curSz", width=width, height=height)
    sub(pic, hp + "flip", horizontal=0, vertical=0)
    sub(pic, hp + "rotationInfo", angle=0, centerX=width // 2, centerY=height // 2, rotateimage=1)
    ri = sub(pic, hp + "renderingInfo")
    for m in ("transMatrix", "scaMatrix", "rotMatrix"):
        sub(ri, hc + m, e1=1, e2=0, e3=0, e4=0, e5=1, e6=0)
    rect = sub(pic, hp + "imgRect")
    for k, (x, y) in enumerate(((0, 0), (width, 0), (width, height), (0, height))):
        sub(rect, hc + f"pt{k}", x=x, y=y)
    sub(pic, hp + "imgClip", left=0, right=width, top=0, bottom=height)
    sub(pic, hp + "inMargin", left=0, right=0, top=0, bottom=0)
    sub(pic, hp + "imgDim", dimwidth=width, dimheight=height)
    sub(pic, hc + "img", binaryItemIDRef=bin_id, bright=0, contrast=0, effect="REAL_PIC", alpha=0)
    sub(pic, hp + "effects")
    sub(pic, hp + "sz", width=width, widthRelTo="ABSOLUTE", height=height,
        heightRelTo="ABSOLUTE", protect=0)
    sub(pic, hp + "pos", treatAsChar=1, affectLSpacing=0, flowWithText=1, allowOverlap=0,
        holdAnchorAndSO=0, vertRelTo="PARA", horzRelTo="COLUMN", vertAlign="TOP",
        horzAlign="CENTER", vertOffset=0, horzOffset=0)
    sub(pic, hp + "outMargin", left=0, right=0, top=0, bottom=0)
    sub(pic, hp + "shapeComment")
    return pic


def embed_images(work: Path, search_dirs: list[Path]) -> int:
    """Replace every ``[[IMG:path|mm]]`` paragraph with the picture; return the count.

    Each picture is stored three times over, as Hancom does: the file under
    ``BinData/``, an item in content.hpf's manifest, and a ``binItem`` in
    header.xml's ``binDataList`` (created at the end of ``refList`` if absent).
    """
    sec_path = work / "Contents" / "section0.xml"
    sec_tree = etree.parse(str(sec_path), parser=_SAFE_XML_PARSER)
    found = []
    for p in sec_tree.getroot().iter(f"{{{HP_NS}}}p"):
        texts = [t for t in p.iter(f"{{{HP_NS}}}t")]
        text = "".join(t.text or "" for t in texts).strip()
        m = IMG_PLACEHOLDER.match(text)
        if m:
            found.append((p, texts, m))
    if not found:
        return 0

    hdr_path = work / "Contents" / "header.xml"
    hdr_tree = etree.parse(str(hdr_path), parser=_SAFE_XML_PARSER)
    ref_list = hdr_tree.getroot().find(f"{{{HH_NS}}}refList")
    bin_list = ref_list.find(f"{{{HH_NS}}}binDataList")
    if bin_list is None:
        bin_list = etree.SubElement(ref_list, f"{{{HH_NS}}}binDataList", itemCnt="0")
    hpf_path = work / "Contents" / "content.hpf"
    hpf_tree = etree.parse(str(hpf_path), parser=_SAFE_XML_PARSER)
    manifest = hpf_tree.getroot().find(f"{{{OPF_NS}}}manifest")
    (work / "BinData").mkdir(exist_ok=True)

    start = len(bin_list)
    for k, (p, texts, m) in enumerate(found):
        src = _resolve_image(m.group("path"), search_dirs)
        ext = src.suffix.lower()
        fmt, media = IMAGE_TYPES[ext]
        data = src.read_bytes()
        px_w, px_h = image_size(data, ext)
        width = min(int(round(float(m.group("mm") or DEFAULT_IMG_MM) * HWPUNIT_PER_MM)), BODY_WIDTH)
        height = int(round(width * px_h / px_w))
        n = start + k + 1
        bin_id = f"BIN{n:04d}"
        (work / "BinData" / f"{bin_id}.{fmt}").write_bytes(data)
        etree.SubElement(bin_list, f"{{{HH_NS}}}binItem", id=str(n - 1), Type="Embedding",
                         BinData=f"{bin_id}.{fmt}", Format=fmt)
        etree.SubElement(manifest, f"{{{OPF_NS}}}item", id=bin_id,
                         href=f"BinData/{bin_id}.{fmt}", **{"media-type": media, "isEmbeded": "1"})
        # The picture takes the first run's place; the placeholder text goes.
        first_run = texts[0].getparent()
        for t in texts:
            t.getparent().remove(t)
        first_run.append(_pic_element(bin_id, width, height, 1900000000 + n))
        p.set("paraPrIDRef", CENTER_PARA_PR)
        for lsa in p.findall(f"{{{HP_NS}}}linesegarray"):
            p.remove(lsa)
    bin_list.set("itemCnt", str(len(bin_list)))

    for tree, path in ((sec_tree, sec_path), (hdr_tree, hdr_path), (hpf_tree, hpf_path)):
        tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)
    return len(found)


def add_page_number(work: Path) -> None:
    """「쪽 번호 매기기」: "- 1 -" at the bottom centre of every page.

    Stored the way Hancom stores it — one ``hp:pageNum`` control in the first
    paragraph (the one carrying ``secPr``).
    """
    sec_path = work / "Contents" / "section0.xml"
    tree = etree.parse(str(sec_path), parser=_SAFE_XML_PARSER)
    first_p = tree.getroot().find(f"{{{HP_NS}}}p")
    if first_p is None or first_p.find(f".//{{{HP_NS}}}pageNum") is not None:
        return
    run = etree.SubElement(first_p, f"{{{HP_NS}}}run", charPrIDRef="0")
    ctrl = etree.SubElement(run, f"{{{HP_NS}}}ctrl")
    etree.SubElement(ctrl, f"{{{HP_NS}}}pageNum", pos="BOTTOM_CENTER", formatType="DIGIT", sideChar="-")
    # Keep it before any linesegarray, which must close the paragraph.
    lsa = first_p.find(f"{{{HP_NS}}}linesegarray")
    if lsa is not None:
        first_p.remove(lsa)
        first_p.append(lsa)
    tree.write(str(sec_path), xml_declaration=True, encoding="UTF-8", standalone=True)


def pack_hwpx(input_dir: Path, output_path: Path) -> None:
    """Create HWPX archive with mimetype as first entry (ZIP_STORED)."""
    mimetype_file = input_dir / "mimetype"
    if not mimetype_file.is_file():
        raise SystemExit(f"Missing 'mimetype' in {input_dir}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    all_files = sorted(
        p.relative_to(input_dir).as_posix()
        for p in input_dir.rglob("*")
        if p.is_file() and should_package_file(p.relative_to(input_dir).as_posix())
    )

    with ZipFile(output_path, "w", ZIP_DEFLATED) as zf:
        zf.write(mimetype_file, "mimetype", compress_type=ZIP_STORED)
        for rel_path in all_files:
            if rel_path == "mimetype":
                continue
            zf.write(input_dir / rel_path, rel_path, compress_type=ZIP_DEFLATED)


def validate_hwpx(hwpx_path: Path) -> list[str]:
    """Quick structural validation of the output HWPX."""
    errors: list[str] = []
    required = [
        "mimetype",
        "Contents/content.hpf",
        "Contents/header.xml",
        "Contents/section0.xml",
    ]

    try:
        from zipfile import BadZipFile
        zf = ZipFile(hwpx_path, "r")
    except BadZipFile:
        return [f"Not a valid ZIP: {hwpx_path}"]

    with zf:
        names = zf.namelist()
        for r in required:
            if r not in names:
                errors.append(f"Missing: {r}")

        if "mimetype" in names:
            content = zf.read("mimetype").decode("utf-8").strip()
            if content != "application/hwp+zip":
                errors.append(f"Bad mimetype content: {content}")
            if names[0] != "mimetype":
                errors.append("mimetype is not the first ZIP entry")
            info = zf.getinfo("mimetype")
            if info.compress_type != ZIP_STORED:
                errors.append("mimetype is not ZIP_STORED")

        for name in names:
            if name.endswith(".xml") or name.endswith(".hpf"):
                try:
                    etree.fromstring(zf.read(name), parser=_SAFE_XML_PARSER)
                except etree.XMLSyntaxError as e:
                    errors.append(f"Malformed XML: {name}: {e}")

    return errors


def build(
    template: str | None,
    header_override: Path | None,
    section_override: Path | None,
    title: str | None,
    creator: str | None,
    output: Path,
    page_number: bool = False,
) -> None:
    """Main build logic."""

    if not BASE_DIR.is_dir():
        raise SystemExit(f"Base template not found: {BASE_DIR}")

    with tempfile.TemporaryDirectory() as tmpdir:
        work = Path(tmpdir) / "build"

        # 1. Copy base template
        shutil.copytree(BASE_DIR, work)

        # 2. Apply template overlay
        if template:
            overlay_dir = TEMPLATES_DIR / template
            if not overlay_dir.is_dir():
                raise SystemExit(
                    f"Template '{template}' not found. "
                    f"Available: {', '.join(AVAILABLE_TEMPLATES)}"
                )
            for overlay_file in overlay_dir.iterdir():
                if overlay_file.is_file() and overlay_file.suffix == ".xml":
                    dest = work / "Contents" / overlay_file.name
                    shutil.copy2(overlay_file, dest)

        # 3. Apply custom overrides
        if header_override:
            if not header_override.is_file():
                raise SystemExit(f"Header file not found: {header_override}")
            shutil.copy2(header_override, work / "Contents" / "header.xml")

        if section_override:
            if not section_override.is_file():
                raise SystemExit(f"Section file not found: {section_override}")
            shutil.copy2(section_override, work / "Contents" / "section0.xml")

        # 3b. Pictures and page numbers
        search_dirs = [section_override.resolve().parent] if section_override else []
        search_dirs.append(Path.cwd())
        n_images = embed_images(work, search_dirs)
        if page_number:
            add_page_number(work)

        # 4. Update metadata
        update_metadata(work / "Contents" / "content.hpf", title, creator)

        # 5. Validate all XML files
        for xml_file in work.rglob("*.xml"):
            validate_xml(xml_file)
        for hpf_file in work.rglob("*.hpf"):
            validate_xml(hpf_file)

        # 6. Pack
        pack_hwpx(work, output)

    # 7. Final validation
    errors = validate_hwpx(output)
    if errors:
        print(f"WARNING: {output} has issues:", file=sys.stderr)
        for e in errors:
            print(f"  - {e}", file=sys.stderr)
    else:
        print(f"VALID: {output}")
        print(f"  Template: {template}")
        if header_override:
            print(f"  Header: {header_override}")
        if section_override:
            print(f"  Section: {section_override}")
        if n_images:
            print(f"  Pictures: {n_images}")
        if page_number:
            print("  Page numbers: - n -")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build HWPX document from templates and XML overrides"
    )
    parser.add_argument(
        "--template", "-t",
        choices=AVAILABLE_TEMPLATES,
        default=DEFAULT_TEMPLATE,
        help=f"Document template to use as overlay (default: {DEFAULT_TEMPLATE})",
    )
    parser.add_argument(
        "--header",
        type=Path,
        help="Custom header.xml to override",
    )
    parser.add_argument(
        "--section",
        type=Path,
        help="Custom section0.xml to override",
    )
    parser.add_argument(
        "--title",
        help="Document title (updates content.hpf metadata)",
    )
    parser.add_argument(
        "--creator",
        help="Document creator (updates content.hpf metadata)",
    )
    parser.add_argument(
        "--page-number",
        action="store_true",
        help='Add "- n -" page numbers at the bottom centre (공문서 관행)',
    )
    parser.add_argument(
        "--output", "-o",
        type=Path,
        required=True,
        help="Output .hwpx file path",
    )
    args = parser.parse_args()

    build(
        template=args.template,
        header_override=args.header,
        section_override=args.section,
        title=args.title,
        creator=args.creator,
        output=args.output,
        page_number=args.page_number,
    )


if __name__ == "__main__":
    main()
