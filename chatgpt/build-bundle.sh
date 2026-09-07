#!/bin/bash
# references/ 를 한 파일로 묶는다. ChatGPT 프로젝트 「파일」에 이 한 개만 올린다.
set -e
cd "$(dirname "$0")/.."
OUT=chatgpt/인터뷰_참조_묶음.md
{
  echo "# 비즈니스 GTM 코어 — 인터뷰 참조 묶음"
  echo
  echo "*이 파일은 references/ 폴더를 한 파일로 합친 것이다. 질문 전문 · 힌트 · 양식 · 조회 방법 · 채널 세팅 순서가 순서대로 들어 있다. 지침이 \"참조 묶음을 검색해 그 문장 그대로\"라고 하면 이 파일을 본다.*"
  echo
  for f in interview-guide hint-playbook research facts-sheet-guide content-direction intro-image channel-setup-guide project-files-template examples next-steps; do
    echo; echo; echo "---"; echo; echo "<!-- FILE: references/$f.md -->"; echo
    cat "references/$f.md"
  done
} > "$OUT"
echo "made $OUT ($(wc -c < "$OUT") bytes)"
