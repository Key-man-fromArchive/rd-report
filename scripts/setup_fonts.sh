#!/bin/sh
# PDF 변환(to_pdf.py)용 한글 글꼴을 받는다: 나눔고딕·나눔명조 (SIL Open Font License).
# TrueType 이라 PDF에 쓴 글자만 들어간다. 문서 글꼴(맑은 고딕·바탕)이 변환 서버에 없을 때
# 이 글꼴로 대체된다.
set -e
DIR="${RD_REPORT_FONTS:-$HOME/.local/share/rd-report-fonts}"
mkdir -p "$DIR"
for p in nanumgothic/NanumGothic-Regular.ttf nanumgothic/NanumGothic-Bold.ttf \
         nanummyeongjo/NanumMyeongjo-Regular.ttf nanummyeongjo/NanumMyeongjo-Bold.ttf; do
  f=$(basename "$p")
  [ -s "$DIR/$f" ] || curl -sfL -o "$DIR/$f" "https://github.com/google/fonts/raw/main/ofl/$p"
done
ls -l "$DIR"
