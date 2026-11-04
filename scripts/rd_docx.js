#!/usr/bin/env node
/*
 * rd_docx.js — 연구개발 보고서를 Word(.docx)로 만든다. 내용은 JSON, 서식은 이 파일.
 *
 *   NODE_PATH=$(npm root -g) node scripts/rd_docx.js report.json output/보고서.docx
 *
 * JSON 형식 (references/rd-report-guide.md 의 목차를 따른다):
 *   {
 *     "title": "○○ 마커 검증 결과보고서",
 *     "org": "주식회사 인바이러스테크",          // 생략 시 config/company.json report_header
 *     "date": "2026-11-04",                      // 생략 시 오늘
 *     "cover": [["과제명", "…"], ["수신", "…"]],   // 생략 가능: 있으면 표지 쪽을 따로 만든다
 *     "blocks": [
 *       {"h": "1. 연구개발과제의 개요"},          // 소제목 ("new_page": true 로 새 쪽)
 *       {"sub": "가. 판정 방법"},                 // 둘째 수준 소제목
 *       {"box": "…"}, {"item": "…"}, {"dash": "…"}, {"note": "…"}, {"text": "…"},
 *       {"table": "achievement", "rows": [[…]], "caption": "목표 달성 수준",
 *        "headers": […], "summary": […], "widths": [3, 6, 1]},   // headers/summary/widths 생략 가능
 *       //   headers: null → 머리줄 없음. "label_col": true → 첫 열을 회색 라벨로. "size": 글자(반 pt)
 *       //   칸 값은 문자열 또는 {"t": "…", "bold": true, "fill": "EEEEEE"}
 *       {"image": "fig/a.png", "mm": 150, "caption": "판정 결과 분포"},
 *       {"grid": [{"title": "#1 …", "sub": "…", "image": "a.png", "frame": "C00000",
 *                  "table": {"headers": […], "rows": [[…]], "size": 14}}], "cols": 2, "mm": 78},
 *       // 그림 격자: 칸마다 제목·부제·그림·작은 표, frame 은 칸 테두리 색(없으면 테두리 없음)
 *       // "rows_per_page": 3, "first_rows": 2 → 그 줄 수마다 쪽을 나눈다(표 안의 표는 프로그램에 따라
 *       // '행 나누지 않기'가 무시되어 칸이 쪽 경계에서 잘리므로, 쪽 나눔을 명시하는 편이 확실하다)
 *       {"pagebreak": true},
 *       {"appendix": "마커별 결론"}              // 새 쪽 + "붙임 n. 마커별 결론"
 *     ]
 *   }
 *
 * 개조식 기호(□ ❍ - ※)는 글자로 넣고 내어쓰기로 정렬한다. 목록 번호 기능을 쓰지 않는 것은
 * 문서를 한글(HWP)·메일 본문으로 옮겨 붙여도 기호가 그대로 남게 하기 위해서다.
 * 표 제목은 위 "<표 n> …", 그림 제목은 아래 "[그림 n] …" 으로 자동 번호를 붙인다.
 */
"use strict";
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, Footer,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, PageNumber, TabStopType,
  TableLayoutType,
} = require("docx");

const SKILL_DIR = path.resolve(__dirname, "..");
const company = (() => {
  const f = path.join(SKILL_DIR, "config", "company.json");
  return fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, "utf8")) : {};
})();

// ---- 서식 (A4, 좌우 20mm) -------------------------------------------------
const PAGE = { width: 11906, height: 16838, top: 1418, bottom: 1134, left: 1134, right: 1134 };
const CONTENT = PAGE.width - PAGE.left - PAGE.right; // 9638 DXA
const FONT_HEAD = company.font_heading || "맑은 고딕";
const FONT_BODY = company.font_body || "바탕";
const SZ = { title: 34, heading: 28, box: 24, body: 23, note: 20, table: 20, caption: 21, meta: 20 };
const GRAY = "D9D9D9";
const LINE = { style: BorderStyle.SINGLE, size: 4, color: "000000" };

const run = (text, o = {}) => new TextRun({ text, font: o.font || FONT_BODY, size: o.size || SZ.body,
  bold: !!o.bold, color: o.color });
// 기호 + 탭 + 내용, 내어쓰기로 둘째 줄을 내용 첫 글자에 맞춘다.
function marked(symbol, text, { left, hanging, font, size, bold, before = 40, after = 40 }) {
  return new Paragraph({
    indent: { left, hanging },
    tabStops: [{ type: TabStopType.LEFT, position: left }],
    spacing: { before, after, line: 300 },
    alignment: AlignmentType.JUSTIFIED,
    children: [run(`${symbol}\t${text}`, { font, size, bold })],
  });
}
const strip = (t, chars) => String(t).replace(new RegExp(`^[${chars}\\s]+`), "").trim();

// ---- 표 양식 (templates/tables 와 같은 이름) ---------------------------------
// [머리말, 너비 비율, 정렬 c|j]
const TABLES = {
  basic: [["연번", 1, "c"], ["구분", 2, "c"], ["내용", 5, "j"], ["비고", 2, "c"]],
  status: [["번호", 1, "c"], ["항목", 3, "c"], ["추진현황", 5, "j"], ["진행률", 1.5, "c"]],
  budget: [["연번", 1, "c"], ["항목", 3, "c"], ["내용", 4, "j"], ["금액(원)", 2, "c"], ["비고", 2, "c"]],
  schedule: [["일자", 2, "c"], ["추진내용", 6, "j"], ["담당", 2, "c"]],
  checklist: [["번호", 1, "c"], ["점검항목", 5, "j"], ["담당", 2, "c"], ["확인", 1, "c"], ["비고", 2, "c"]],
  summary: [["구분", 1, "c"], ["내용", 4, "j"]],
  achievement: [["추진목표", 3, "j"], ["달성내용", 6, "j"], ["달성도(%)", 1.3, "c"]],
  sample_plan: [["구분", 2, "c"], ["대상", 3, "j"], ["시료 수", 1.4, "c"], ["반응 수", 1.4, "c"], ["비고", 2.5, "j"]],
};

function dxaWidths(ratios, totalWidth = CONTENT) {
  const total = ratios.reduce((a, b) => a + b, 0);
  const w = ratios.map((r) => Math.floor((totalWidth * r) / total));
  w[w.length - 1] += totalWidth - w.reduce((a, b) => a + b, 0);
  return w;
}

function cell(value, width, { header = false, align = "c", size = SZ.table, label = false, pad = 60 } = {}) {
  const v = value !== null && typeof value === "object" ? value : { t: value };
  const shaded = header || label || v.fill;
  const bold = header || label || !!v.bold;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: { top: LINE, bottom: LINE, left: LINE, right: LINE },
    shading: shaded ? { fill: v.fill || GRAY, type: ShadingType.CLEAR, color: "auto" } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: pad, bottom: pad, left: Math.min(100, pad + 40), right: Math.min(100, pad + 40) },
    children: String(v.t).split("\n").map((line) => new Paragraph({
      alignment: header || label || align === "c" ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { line: size < SZ.table ? 220 : 260 },
      children: [run(line, { font: header || label ? FONT_HEAD : FONT_BODY, size, bold })],
    })),
  });
}

function table(block, totalWidth = CONTENT) {
  const first = block.headers || block.rows[0];
  const spec = TABLES[block.table] || first.map(() => ["", 1, "c"]);
  const headers = block.headers === null ? null : block.headers || spec.map((c) => c[0]);
  const aligns = block.align || spec.map((c) => c[2]);
  const widths = dxaWidths(block.widths || spec.map((c) => c[1]), totalWidth);
  if (headers && headers.length !== widths.length) throw new Error(`표 ${block.table}: 열 수가 맞지 않음`);
  const size = block.size || SZ.table;
  const pad = size <= 15 ? 20 : size < SZ.table ? 40 : 60;
  const fill = (v) => (v === null || v === undefined || String(v) === "" ? "-" : v);
  const rows = [];
  if (headers) rows.push(new TableRow({ tableHeader: true, cantSplit: true,
    children: headers.map((h, i) => cell(h, widths[i], { header: true, size, pad })) }));
  for (const r of block.rows) {
    if (r.length !== widths.length) throw new Error(`표 ${block.table}: 행의 칸 수 ${r.length} ≠ ${widths.length}`);
    rows.push(new TableRow({ cantSplit: true, children: r.map((v, i) =>
      cell(fill(v), widths[i], { align: aligns[i], size, pad, label: !!block.label_col && i === 0 })) }));
  }
  if (block.summary) rows.push(new TableRow({ cantSplit: true,
    children: block.summary.map((v, i) => cell(fill(v), widths[i], { header: true, size, pad })) }));
  // 고정 폭: '내용에 맞춤'이면 긴 서열이 열을 넓히고 좁은 열의 글자가 세로로 깨진다.
  return new Table({ width: { size: totalWidth, type: WidthType.DXA }, columnWidths: widths,
    layout: TableLayoutType.FIXED, rows });
}

const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
// 표 칸은 문단으로 끝나야 하므로 표 뒤에 두는 빈 문단 — 높이를 거의 0으로 해 쪽을 넘기지 않게 한다.
const tiny = () => new Paragraph({ spacing: { before: 0, after: 0, line: 40, lineRule: "exact" },
  children: [new TextRun({ text: "", size: 2 })] });

// 그림 격자: 칸마다 제목·부제·그림·작은 표. frame 색이 있으면 그 칸만 테두리를 두른다.
function grid(block, baseDir) {
  const cols = block.cols || 2;
  const colW = Math.floor(CONTENT / cols);
  const innerW = colW - 160;
  const items = block.grid;
  const rows = [];
  for (let i = 0; i < items.length; i += cols) {
    rows.push(new TableRow({ cantSplit: true, children: Array.from({ length: cols }, (_, k) => {
      const it = items[i + k];
      const kids = [];
      if (it) {
        if (it.title) kids.push(new Paragraph({ keepNext: true, children: [run(it.title, { font: FONT_HEAD, size: 19, bold: true })] }));
        if (it.sub) kids.push(new Paragraph({ keepNext: true, spacing: { after: 40 }, children: [run(it.sub, { font: FONT_HEAD, size: 15 })] }));
        if (it.image) kids.push(imagePara(path.resolve(baseDir, it.image), Math.min(block.mm || 78, (innerW / 1440) * 25.4), null, 20));
        if (it.table) kids.push(table(it.table, innerW - 120));
      }
      kids.push(tiny());
      const edge = it && it.frame ? { style: BorderStyle.SINGLE, size: 18, color: it.frame } : NONE;
      const inner = new Table({ width: { size: innerW, type: WidthType.DXA }, columnWidths: [innerW],
        layout: TableLayoutType.FIXED, rows: [
        new TableRow({ cantSplit: true, children: [new TableCell({ width: { size: innerW, type: WidthType.DXA },
          borders: { top: edge, bottom: edge, left: edge, right: edge },
          margins: { top: 60, bottom: 20, left: 60, right: 60 }, children: kids })] })] });
      return new TableCell({ width: { size: colW, type: WidthType.DXA },
        borders: { top: NONE, bottom: NONE, left: NONE, right: NONE },
        margins: { top: 40, bottom: 120, left: 40, right: 40 }, children: [inner, tiny()] });
    }) }));
  }
  return new Table({ width: { size: colW * cols, type: WidthType.DXA }, columnWidths: Array(cols).fill(colW),
    layout: TableLayoutType.FIXED, rows });
}

function imagePara(file, mm, caption, before = 160) {
  const buf = fs.readFileSync(file);
  const [pw, ph] = imageSize(buf);
  const maxPx = (CONTENT / 1440) * 96;
  const wPx = Math.min((mm / 25.4) * 96, maxPx);
  return new Paragraph({ alignment: AlignmentType.CENTER, keepNext: !!caption, spacing: { before, after: 60 },
    children: [new ImageRun({ type: /\.jpe?g$/i.test(file) ? "jpg" : "png", data: buf,
      transformation: { width: Math.round(wPx), height: Math.round((wPx * ph) / pw) },
      altText: { title: caption || path.basename(file), description: caption || path.basename(file), name: path.basename(file) } })] });
}

// PNG/JPEG 크기 (픽셀) — 높이를 비율대로 맞추려고 읽는다.
function imageSize(buf) {
  if (buf.readUInt32BE(0) === 0x89504e47) return [buf.readUInt32BE(16), buf.readUInt32BE(20)];
  let i = 2;
  while (i < buf.length) {
    if (buf[i] !== 0xff) { i++; continue; }
    const m = buf[i + 1];
    if (m >= 0xc0 && m <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(m)) return [buf.readUInt16BE(i + 7), buf.readUInt16BE(i + 5)];
    i += 2 + buf.readUInt16BE(i + 2);
  }
  throw new Error("이미지 크기를 읽지 못함 (PNG/JPEG만)");
}

function build(spec, baseDir) {
  const today = spec.date ? new Date(spec.date) : new Date();
  const yy = String(today.getFullYear() % 100).padStart(2, "0");
  const org = spec.org || company.report_header || "";
  const children = [];
  let nTable = 0, nFig = 0, nAppendix = 0;

  if (spec.cover) {
    children.push(new Paragraph({ spacing: { before: 3000, after: 600 }, alignment: AlignmentType.CENTER,
      children: [run(spec.title, { font: FONT_HEAD, size: 40, bold: true })] }));
    children.push(table({ table: "summary", headers: ["구분", "내용"], rows: spec.cover, widths: [1, 3] }));
    children.push(new Paragraph({ spacing: { before: 2400 }, alignment: AlignmentType.CENTER,
      children: [run(`${today.getFullYear()}. ${today.getMonth() + 1}. ${today.getDate()}.`, { font: FONT_HEAD, size: 24 })] }));
    children.push(new Paragraph({ alignment: AlignmentType.CENTER, pageBreakBefore: false,
      children: [run(org, { font: FONT_HEAD, size: 28, bold: true })] }));
    children.push(new Paragraph({ pageBreakBefore: true, children: [] }));
  }

  // 머리말 줄과 제목 상자 (공공기관 보고서 관행)
  children.push(new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 80 },
    children: [run(`(’${yy}. ${today.getMonth() + 1}. ${today.getDate()}.${org ? ", " + org : ""})`, { font: FONT_HEAD, size: SZ.meta })] }));
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 360 },
    border: { top: { style: BorderStyle.DOUBLE, size: 6, color: "000000", space: 8 },
              bottom: { style: BorderStyle.DOUBLE, size: 6, color: "000000", space: 8 } },
    children: [run(spec.title, { font: FONT_HEAD, size: SZ.title, bold: true })] }));

  for (const b of spec.blocks || []) {
    if (b.h !== undefined) {
      children.push(new Paragraph({ pageBreakBefore: !!b.new_page, keepNext: true,
        spacing: { before: b.new_page ? 0 : 320, after: 120 },
        children: [run(b.h, { font: FONT_HEAD, size: SZ.heading, bold: true })] }));
    } else if (b.sub !== undefined) {
      children.push(new Paragraph({ keepNext: true, spacing: { before: 200, after: 80 },
        children: [run(b.sub, { font: FONT_HEAD, size: SZ.box, bold: true })] }));
    } else if (b.grid !== undefined) {
      if (b.caption) {
        nFig++;
        children.push(new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 120, after: 80 },
          children: [run(`[그림 ${nFig}] ${b.caption}`, { font: FONT_HEAD, size: SZ.caption })] }));
      }
      const cols = b.cols || 2;
      const per = (b.rows_per_page || 0) * cols;
      if (!per) {
        children.push(grid(b, baseDir));
      } else {
        let start = 0, take = (b.first_rows || b.rows_per_page) * cols;
        while (start < b.grid.length) {
          if (start > 0) children.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 0 }, children: [] }));
          children.push(grid({ ...b, grid: b.grid.slice(start, start + take) }, baseDir));
          start += take;
          take = per;
        }
      }
    } else if (b.appendix !== undefined) {
      nAppendix++;
      children.push(new Paragraph({ pageBreakBefore: true, keepNext: true, spacing: { after: 200 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "000000", space: 4 } },
        children: [run(`붙임 ${nAppendix}. ${b.appendix}`, { font: FONT_HEAD, size: SZ.heading, bold: true })] }));
    } else if (b.box !== undefined) {
      children.push(marked("□", strip(b.box, "□"), { left: 340, hanging: 340, font: FONT_HEAD, size: SZ.box, bold: true, before: 160 }));
    } else if (b.item !== undefined) {
      children.push(marked("❍", strip(b.item, "❍○"), { left: 680, hanging: 300 }));
    } else if (b.dash !== undefined) {
      children.push(marked("-", strip(b.dash, "\\-"), { left: 960, hanging: 220 }));
    } else if (b.note !== undefined) {
      children.push(marked("※", strip(b.note, "※"), { left: 600, hanging: 260, font: FONT_HEAD, size: SZ.note }));
    } else if (b.text !== undefined) {
      children.push(new Paragraph({ spacing: { before: 40, after: 40, line: 300 }, children: [run(b.text)] }));
    } else if (b.table !== undefined) {
      if (b.caption) {
        nTable++;
        children.push(new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 200, after: 80 },
          children: [run(`<표 ${nTable}> ${b.caption}`, { font: FONT_HEAD, size: SZ.caption, bold: true })] }));
      }
      children.push(table(b));
      children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    } else if (b.image !== undefined) {
      children.push(imagePara(path.resolve(baseDir, b.image), b.mm || 150, b.caption));
      if (b.caption) {
        nFig++;
        children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
          children: [run(`[그림 ${nFig}] ${b.caption}`, { font: FONT_HEAD, size: SZ.caption })] }));
      }
    } else if (b.pagebreak) {
      children.push(new Paragraph({ pageBreakBefore: true, children: [] }));
    } else {
      throw new Error(`알 수 없는 블록: ${JSON.stringify(b)}`);
    }
  }

  return new Document({
    creator: company.creator || org, title: spec.title,
    styles: { default: { document: { run: { font: FONT_BODY, size: SZ.body } } } },
    sections: [{
      // 표지가 있으면 첫 쪽에는 쪽 번호를 찍지 않는다(titlePage + 빈 첫 쪽 바닥글).
      properties: { titlePage: !!spec.cover, page: { size: { width: PAGE.width, height: PAGE.height },
        margin: { top: PAGE.top, bottom: PAGE.bottom, left: PAGE.left, right: PAGE.right } } },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: ["- ", PageNumber.CURRENT, " -"], font: FONT_HEAD, size: 20 })] })] }),
        first: new Footer({ children: [new Paragraph({ children: [] })] }) },
      children,
    }],
  });
}

async function main() {
  const [specPath, outPath] = process.argv.slice(2);
  if (!specPath || !outPath) {
    console.error("usage: node rd_docx.js report.json output.docx");
    process.exit(2);
  }
  const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));
  const doc = build(spec, path.dirname(path.resolve(specPath)));
  fs.mkdirSync(path.dirname(path.resolve(outPath)), { recursive: true });
  fs.writeFileSync(outPath, await Packer.toBuffer(doc));
  console.log(`DOCX: ${outPath}`);
}

if (require.main === module) main().catch((e) => { console.error(e.message); process.exit(1); });
module.exports = { build, TABLES };
