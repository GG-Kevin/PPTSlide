# PPTSlide

회사 표준 PowerPoint 포맷을 그대로 유지하면서 슬라이드를 만드는 에이전트 파이프라인.

## 무엇이 들어있나

- **포맷 원본** `templates/company_format.pptx` — 마스터/레이아웃/테마/모티프 바/페이지 번호
- **포맷 스펙** `format/format_spec.json` — 원본에서 기계 추출한 값 (수준별 글꼴·크기·불릿·색)
- **에이전트 6종** — 포맷 가디언 / 이미지 분석 / 텍스트 분석 / 도해 제작 / PPT 제작 / 최종 검증
- **툴 4종** — 스펙 추출, 슬라이드 빌드, PNG 렌더, 포맷 검증
- **완성 예시** `output/Introduction_개발일정.pptx`

## 빠른 시작

```bash
pip install python-pptx pillow
apt-get install -y libreoffice-impress poppler-utils fonts-nanum fonts-nanum-extra
mkdir -p ~/.config/fontconfig && cp tools/render-fonts.conf ~/.config/fontconfig/fonts.conf && fc-cache -f

python3 tools/extract_format.py
python3 tools/build_slide.py content/intro_dev_schedule.json "output/Introduction_개발일정.pptx"
python3 tools/verify.py "output/Introduction_개발일정.pptx"
```

## 검증

`tools/verify.py` 는 26개 항목을 본다.

- **구조(XML)** — 슬라이드 크기, 마스터/레이아웃 세트 유지, 본문 레이아웃 사용,
  제목·본문 플레이스홀더 사용, 제목 스타일 상속, 승인 폰트만 사용,
  모티프 바/페이지번호 영역 침범 없음, 슬라이드 경계 이탈 없음
- **화면(렌더)** — 흰 배경, 모티프 바 위치·남색·녹색·분할점, 제목 좌측/세로 위치,
  우하단 페이지 번호, 본문 좌측 여백·우측 침범·하단 넘침

`format/reference_intro_slide.png` 가 픽셀 기준값이다.

## 포맷을 새로 받았을 때

`templates/company_format.pptx` 를 교체하고 `python3 tools/extract_format.py` 를 다시 돌린다.
그러면 빌더와 검증기가 새 스펙을 자동으로 따라간다.
