#!/usr/bin/env python3
"""예제: 연구개발 결과보고서(국가연구개발혁신법 시행규칙 별지 제5호 최종보고서 목차 준용).

내용은 가상의 예시다. 실행하면 output/예제_연구개발결과보고서.hwpx 가 만들어진다.

    python3 examples/rd_result_report.py
"""
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "scripts"))
from rd_section import Report  # noqa: E402

r = Report(title="○○ 분자마커 패널 검증 결과보고서")

r.heading("요약")
r.table("summary", [
    ["과제명", "○○ 분자마커 76종 ASG-PCR 전환 및 검증"],
    ["수행 기간", "‘26.09.14 ~ ‘26.11.04"],
    ["최종 목표", "칩 유전형과 일치하는 실용 마커 60종 이상 확보"],
    ["주요 결과", "바로 사용 40종, 조건부 12종, 재설계 12종, 추가 확인 12종"],
    ["활용 계획", "재설계·재확인 실험 후 품종 판별 패널로 이관"],
])

r.heading("1. 연구개발과제의 개요")
r.box("칩 기반 마커를 현장형 PCR 마커로 전환")
r.item("GWAS로 선발된 76개 SNP를 ASG-PCR 프라이머로 설계하고 칩 유전형과 비교 검증함")
r.dash("마커당 시료 10개(동형 3, 반대 동형 3, 이형 4) 및 NTC 4웰 구성")

r.heading("2. 연구개발과제의 수행 과정 및 수행 내용")
r.box("설계 → 합성 → 실시간 PCR 판독 → 칩 비교 순 진행")
r.item("판독 사이클 36부터 형광값을 기록하고 NTC를 제외한 군집으로 재판정함")
r.note("NTC: 주형 DNA를 넣지 않은 음성 대조")

r.heading("3. 연구개발과제의 수행 결과 및 목표 달성 정도", new_page=True)
r.box("목표 60종 대비 바로 사용 가능 40종 확보(67%)")
r.table("achievement", [
    ["실용 마커 확보", "바로 사용 40종, 조건부 12종", "67"],
    ["판정 기준 정립", "NTC 제외 재판정·보류 규칙 마련", "100"],
], caption="<표 1> 목표 달성 수준")
r.image(HERE / "fig_example.png", mm=140, caption="[그림 1] 마커별 판정 결과 분포")

r.heading("4. 목표 미달 시 원인분석")
r.box("미판정 다발 원인은 설계보다 판정 시점 문제")
r.item("판독 시점에 NTC 신호가 시료 군집과 겹쳐 판정 불가가 늘어난 것으로 추정됨")

r.heading("5. 연구개발성과 및 관련 분야에 대한 기여 정도")
r.box("칩 없이 현장에서 쓸 수 있는 판별 마커 확보")

r.heading("6. 연구개발성과의 관리 및 활용 계획")
r.table("sample_plan", [
    ["재설계 검증", "재설계 12종", "10", "168", "NTC 4웰 포함"],
    ["조건부 재확인", "조건부 12종", "10", "168", "사이클별 판독"],
], summary=["합  계", "24종", "-", "336", "-"], caption="<표 2> 후속 실험 시료·반응 계획")

out = r.save(Path.cwd() / "output" / "예제_연구개발결과보고서.hwpx")
print(out)
