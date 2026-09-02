---
name: ppt-format
description: 회사 표준 PPT 포맷(templates/company_format.pptx)으로 슬라이드를 제작·수정·검증한다. 슬라이드/PPT/발표자료를 만들어 달라거나, 스크립트·메모를 슬라이드로 옮겨 달라거나, 기존 슬라이드가 회사 포맷에 맞는지 확인해 달라고 할 때 사용한다.
---

# 회사 표준 포맷 PPT 제작

이 저장소는 **회사 표준 포맷을 절대 깨지 않는 것**을 최우선으로 한다.
슬라이드가 예뻐도 포맷이 다르면 실패다.

## 0. 준비 (환경에 한 번)
```bash
pip install python-pptx pillow
apt-get install -y libreoffice-impress poppler-utils fonts-nanum fonts-nanum-extra
mkdir -p ~/.config/fontconfig && cp tools/render-fonts.conf ~/.config/fontconfig/fonts.conf && fc-cache -f
```
마지막 줄은 회사 폰트(`현대하모니 L`)가 없는 리눅스에서 렌더 미리보기를 신뢰할 수 있게 만든다.
PPTX 안의 폰트 지정은 바뀌지 않는다.

## 1. 파이프라인
| 단계 | 에이전트 | 산출물 |
|---|---|---|
| 포맷 등록·추출 | `ppt-format-guardian` | `templates/company_format.pptx`, `format/format_spec.json` |
| 참고 이미지 실측 | `ppt-image-analyst` | 도형 좌표·색·글자 크기(EMU/hex/pt) |
| 스크립트 구조화 | `ppt-text-analyst` | `content/<slide_id>.json`, `content/source_script.txt` |
| 도해 제작 | `ppt-graphic-builder` | `tools/build_slide.py` 의 도형 XML·좌표 상수 |
| 슬라이드 생성 | `ppt-builder` | `output/<이름>.pptx` |
| 최종 검증 | `ppt-verifier` | `output/verify_report.json` (26/26 PASS 여야 종료) |

검증에서 FAIL 이 나오면 **해당 단계로 되돌아가 수정하고 다시 검증한다.** PASS 할 때까지 반복한다.

## 2. 명령
```bash
python3 tools/extract_format.py                                   # 포맷 → 스펙
python3 tools/build_slide.py content/x.json "output/x.pptx"       # 슬라이드 생성
python3 tools/render.py "output/x.pptx" output/preview --dpi 102  # 미리보기 PNG
python3 tools/verify.py "output/x.pptx"                           # 26개 항목 검증
```

## 3. 절대 규칙
1. `templates/company_format.pptx` 는 **원본**이다. 편집하지 않는다.
2. 새 덱은 그 파일을 **열어서** 레이아웃을 상속받아 만든다. 빈 `Presentation()` 금지.
3. 본문 슬라이드 레이아웃은 `3. 본문 및 내용`.
4. 상단 모티프 바(#014099 남색 + #009944 녹색)와 우하단 페이지 번호는 레이아웃이 제공한다.
   슬라이드에 다시 그리지 않는다.
5. 글꼴 `현대하모니 L`, 글자색 테마 `tx1`. 도해 안 흰 글씨만 예외.
6. 제목 런에 `sz`/`b`/`latin`/`solidFill` 을 직접 쓰지 않는다 — 레이아웃에서 상속받는다.
7. 마스터·레이아웃·테마는 수정하지 않는다.

## 4. 수준별 스타일 (format_spec.json `body_levels`)
| 수준 | 불릿 | 크기 | 굵기 | 줄간격 | marL / indent |
|---|---|---|---|---|---|
| lvl1 | ❖ (Wingdings `v`) | 17pt | Bold | 32pt | 285750 / −285750 |
| lvl2 | ● (Wingdings `l`) | 14pt | — | 21.6pt | 541338 / −254000 |
| lvl3 | ➢ (Wingdings `Ø`) | 13pt | — | 21pt | 808038 / −268288 |
| lvl4 | − | 13pt | — | (기본) | 896938 / −134938 |

제목: 현대하모니 L 24pt Bold, 좌측 정렬.

## 5. 분량 예산
본문 가용 높이는 약 311pt(4.3"). `(lvl2 줄수 × 21.6) + (lvl4 줄수 × 17.5) + 문단간격` 으로 미리 계산한다.
넘치면 **글자를 줄이지 말고 문장을 줄인다.**
