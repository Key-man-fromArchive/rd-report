"""윤문 전후 사실 보존 검사.

    python fact_guard.py 원본.json 윤문본.json [--restore 출력.json]
    python fact_guard.py 원본.md   윤문본.md

소제목 구간(JSON: `h` 블록, Markdown: `#` 줄) 마다 다음이 같은지 본다.
  - 수치 토큰 (쉼표·유니코드 마이너스 정규화, 뒤에 붙은 단위 포함)
  - 영문·숫자 식별자 (마커명, 파일명, 유전자명, P1 같은 라벨)
  - 부정 표현 수 (않·없·못·아니·불가)
  - 유보 표현 수 (추정·가능성·~로 본다/보임·예상·듯)
JSON 은 소제목·표·그림·표지 같은 구조 블록이 글자 그대로 같아야 한다.
위반이 있으면 종료 코드 1. --restore 를 주면 위반 구간만 원본으로 되돌려 저장한다(JSON 만).
"""
from __future__ import annotations

import collections
import json
import re
import sys
from pathlib import Path

TEXT_KEYS = ("box", "item", "dash", "note", "text")
STRUCT_KEYS = ("h", "sub", "appendix", "table", "grid", "image", "pagebreak")
UNIT = r"(?:%|‰|°C|℃|nM|µM|uM|mM|M|nt|bp|kb|Mb|RFU|ng|µg|ug|mg|g|µL|uL|mL|L|s|min|h|배|개|종|웰|회|건|명|판|쪽|일|주|월|년|사이클|시간|분|초|원|만|억)?"
NUM = re.compile(r"(?<![A-Za-z0-9_])[-+−–]?\d[\d,]*(?:\.\d+)?\s?" + UNIT)
IDENT = re.compile(r"[A-Za-z][A-Za-z0-9_.\-/]*[A-Za-z0-9]|[A-Za-z]")
NEG = re.compile(r"않|없|못|아니|불가|미확인|미검출")
HEDGE = re.compile(r"추정|가능성|(?:로|것으로) (?:보임|보인|본다|봄|봤)|예상|추측|불확실|듯")


def norm_num(tok: str) -> str:
    t = tok.replace(",", "").replace("−", "-").replace("–", "-").replace(" ", "")
    return t


def facts(text: str) -> dict:
    nums = collections.Counter(norm_num(m.group()) for m in NUM.finditer(text))
    stripped = NUM.sub(" ", text)
    ids = collections.Counter(m.group() for m in IDENT.finditer(stripped) if len(m.group()) >= 2 or m.group().isupper())
    return dict(nums=nums, ids=ids, neg=len(NEG.findall(text)), hedge=len(HEDGE.findall(text)))


# ---------- 구간 나누기 ----------
def json_sections(doc: dict):
    """[(소제목 or '(머리)', [블록...])]"""
    out = [["(머리)", []]]
    for b in doc.get("blocks", []):
        if "h" in b:
            out.append([b["h"], [b]])
        else:
            out[-1][1].append(b)
    return out


def block_text(b: dict) -> str:
    return " ".join(str(b[k]) for k in TEXT_KEYS if k in b)


def struct_sig(blocks) -> list:
    return [json.dumps({k: b[k] for k in b}, ensure_ascii=False, sort_keys=True)
            for b in blocks if any(k in b for k in STRUCT_KEYS)]


def md_sections(text: str):
    out = [["(머리)", []]]
    in_code = False
    for line in text.splitlines():
        if line.lstrip().startswith("```"):
            in_code = not in_code
        if not in_code and re.match(r"#{1,6}\s", line):
            out.append([line.strip(), []])
        else:
            out[-1][1].append(line)
    return out


# ---------- 비교 ----------
def diff_counter(a: collections.Counter, b: collections.Counter):
    lost = sorted((a - b).elements())
    added = sorted((b - a).elements())
    return lost, added


def compare(name: str, ta: str, tb: str, sa=None, sb=None) -> list[str]:
    v = []
    if sa is not None and sa != sb:
        v.append("구조 블록(소제목·표·그림 등)이 바뀜")
    fa, fb = facts(ta), facts(tb)
    for key, label in (("nums", "수치"), ("ids", "식별자")):
        lost, added = diff_counter(fa[key], fb[key])
        if lost:
            v.append(f"{label} 사라짐: {' '.join(lost)}")
        if added:
            v.append(f"{label} 새로 생김: {' '.join(added)}")
    if fb["neg"] < fa["neg"]:
        v.append(f"부정 표현 감소 {fa['neg']} → {fb['neg']} (의미 반전 확인)")
    if fb["hedge"] < fa["hedge"]:
        v.append(f"유보 표현 감소 {fa['hedge']} → {fb['hedge']} (추정이 단정으로 바뀌었는지 확인)")
    return v


def run(orig: Path, new: Path, restore: Path | None) -> int:
    is_json = orig.suffix.lower() == ".json"
    if is_json:
        A, B = json.loads(orig.read_text(encoding="utf-8")), json.loads(new.read_text(encoding="utf-8"))
        meta = [k for k in set(A) | set(B) if k != "blocks" and A.get(k) != B.get(k)]
        SA, SB = json_sections(A), json_sections(B)
    else:
        meta = []
        SA, SB = md_sections(orig.read_text(encoding="utf-8")), md_sections(new.read_text(encoding="utf-8"))

    problems: dict[int, list[str]] = {}
    if meta:
        problems[-1] = [f"머리 항목이 바뀜: {', '.join(sorted(meta))}"]
    if [s[0] for s in SA] != [s[0] for s in SB]:
        print("소제목 목록이 다르다. 윤문에서 소제목을 바꾸거나 구간을 옮기지 않는다.")
        print("  원본 :", " | ".join(s[0] for s in SA))
        print("  윤문본:", " | ".join(s[0] for s in SB))
        return 1

    for i, ((name, ba), (_, bb)) in enumerate(zip(SA, SB)):
        if is_json:
            v = compare(name, " \n".join(block_text(b) for b in ba), " \n".join(block_text(b) for b in bb),
                        struct_sig(ba), struct_sig(bb))
        else:
            v = compare(name, "\n".join(ba), "\n".join(bb))
        if v:
            problems[i] = v

    print(f"사실 보존 검사: {orig.name} → {new.name} (구간 {len(SA)}개)")
    if not problems:
        print("  위반 없음")
        return 0
    for i, v in sorted(problems.items()):
        print(f"  [{SA[i][0] if i >= 0 else '문서'}]")
        for x in v:
            print(f"    - {x}")

    if restore:
        if not is_json:
            print("--restore 는 JSON 만 지원한다.")
            return 1
        out = dict(B)
        if -1 in problems:
            for k in meta:
                if k in A:
                    out[k] = A[k]
                else:
                    out.pop(k, None)
        blocks = []
        for i, (sa, sb) in enumerate(zip(SA, SB)):
            blocks += sa[1] if i in problems else sb[1]
        out["blocks"] = blocks
        restore.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"  위반 구간 {len([i for i in problems if i >= 0])}개를 원본으로 되돌려 {restore} 에 저장")
    return 1


def main() -> None:
    args = sys.argv[1:]
    restore = None
    if "--restore" in args:
        i = args.index("--restore")
        restore = Path(args[i + 1])
        del args[i:i + 2]
    if len(args) != 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(run(Path(args[0]), Path(args[1]), restore))


if __name__ == "__main__":
    main()
