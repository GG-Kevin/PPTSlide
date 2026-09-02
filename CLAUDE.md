# PPTSlide

회사 표준 PPT 포맷을 유지하면서 슬라이드를 제작하는 저장소.

## 가장 중요한 것
`templates/company_format.pptx` 가 **회사 표준 포맷 원본**이다. 이 파일은 편집하지 않는다.
모든 슬라이드는 이 파일을 열어 그 안의 레이아웃을 상속받아 만든다.
본문 글꼴은 회사 폰트가 없는 PC를 고려해 테마 폰트인 **맑은 고딕**을 쓴다.

슬라이드 제작·수정·검증 요청이 오면 `ppt-format` 스킬을 먼저 읽는다.
파이프라인·명령·절대 규칙이 전부 거기에 있다.

## 구조
```
templates/company_format.pptx     회사 표준 포맷 원본 (편집 금지)
format/format_spec.json           원본에서 추출한 포맷 스펙 (기계 판독용 기준)
format/reference_intro_slide.png  완성 예시 기준 이미지 (검증 픽셀 기준)
content/<slide_id>.json           슬라이드 콘텐츠 구조 (slides 배열이면 여러 장)
assets/screens/                   프로그램 스크린샷
content/source_script.txt         원본 구술 스크립트 (보존)
tools/extract_format.py           포맷 → 스펙 추출
tools/shapes.py                   도형 XML 헬퍼 (그라데이션·둥근모서리·불릿)
tools/blocks.py                   블록 렌더러 (image·table·text·note·panels·flow)
tools/capture_app.py              배포용 HTML 프로그램 → 화면별 스크린샷
tools/build_slide.py              콘텐츠 + 포맷 → pptx (단일/다중 슬라이드)
tools/render.py                   pptx → PNG 미리보기
tools/verify.py                   29개 항목 포맷 검증
tools/render-fonts.conf           렌더 환경용 한글 대체 폰트 매핑
output/                           산출물
.claude/agents/                   ppt-* 에이전트 6종
.claude/skills/ppt-format/        파이프라인 스킬
```

## 완료 기준
`python3 tools/verify.py output/<이름>.pptx` 가 **전 항목 PASS** 여야 한다.
슬라이드 1장당 구조 15 + 화면 14 항목을 검사한다 (16장이면 408항목).
FAIL 을 남긴 채 완료 보고하지 않는다.
