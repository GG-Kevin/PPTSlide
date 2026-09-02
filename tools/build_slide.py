#!/usr/bin/env python3
"""회사 표준 포맷(templates/company_format.pptx) 위에 슬라이드를 생성한다.

  python3 tools/build_slide.py content/intro_dev_schedule.json output/Introduction_개발일정.pptx

규칙(포맷 유지):
  * 반드시 templates/company_format.pptx 를 열어 그 안의 레이아웃으로 슬라이드를 추가한다.
  * 상단 모티프 바 / 우하단 페이지번호 / 제목 스타일은 레이아웃이 제공하므로 다시 그리지 않는다.
  * 본문 글꼴은 '현대하모니 L', 글자색은 테마 tx1 을 유지한다.
"""
import json
import sys
from pathlib import Path
from xml.sax.saxutils import escape

from pptx import Presentation
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml import parse_xml

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "company_format.pptx"
SPEC = json.loads((ROOT / "format" / "format_spec.json").read_text(encoding="utf-8"))

FONT = "현대하모니 L"
FONT_ATTRS = 'panose="02020603020101020101" pitchFamily="18" charset="-127"'

# ── 기준 이미지(format/reference_intro_slide.png, 1361x761px)에서 실측한 도해 좌표 ──
SCALE_X = 12192000 / 1361
SCALE_Y = 6858000 / 761
BAND_TOP_PX, BAND_BOTTOM_PX = 126.0, 258.5
# (좌, 우) px — 이웃한 셰브런이 촉(tip) 깊이만큼 겹치도록 배치된다
STEP_BOX_PX = [(100, 301), (235, 552), (486, 802), (736, 1053), (987, 1261)]
STEP_LABEL_PT = 1700
# 1번(homePlate)은 촉을 제외한 사각형에 글자가 중앙정렬되어 왼쪽으로 쏠린다 → 광학 보정
STEP1_LINS_PX = 20.0

DETAIL_TOP = 2420000
DETAIL_BOTTOM = 6453188


def px_x(v):
    return int(round(v * SCALE_X))


def px_y(v):
    return int(round(v * SCALE_Y))


def font_tags():
    return (
        f'<a:latin typeface="{FONT}" {FONT_ATTRS}/>'
        f'<a:ea typeface="{FONT}" {FONT_ATTRS}/>'
        f'<a:cs typeface="Arial" pitchFamily="34" charset="0"/>'
    )


def delete_all_slides(prs):
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        prs.part.drop_rel(sld_id.get(qn("r:id")))
        sld_id_lst.remove(sld_id)


def get_layout(prs, name):
    for layout in prs.slide_masters[0].slide_layouts:
        if layout.name == name:
            return layout
    raise SystemExit(f"레이아웃 '{name}' 을 포맷에서 찾을 수 없음")


def chevron_xml(shape_id, idx, step, box_px, band_top, band_h):
    left_px, right_px = box_px
    left, width = px_x(left_px), px_x(right_px - left_px)
    prst = "homePlate" if idx == 0 else "chevron"
    lins = px_x(STEP1_LINS_PX) if idx == 0 else 0
    lines = "".join(
        f'<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="{STEP_LABEL_PT}" b="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>{font_tags()}</a:rPr>'
        f"<a:t>{escape(t)}</a:t></a:r>" + ("<a:br><a:rPr lang=\"ko-KR\"/></a:br>" if i < len(step["label"]) - 1 else "")
        for i, t in enumerate(step["label"])
    )
    return parse_xml(
        f'<p:sp {nsdecls("p", "a")}>'
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="공정 셰브런 {idx + 1}"/>'
        f"<p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        f"<p:spPr>"
        f'<a:xfrm><a:off x="{left}" y="{band_top}"/><a:ext cx="{width}" cy="{band_h}"/></a:xfrm>'
        f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
        f'<a:gradFill flip="none" rotWithShape="1"><a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="{step["top"]}"/></a:gs>'
        f'<a:gs pos="12000"><a:srgbClr val="{step["mid"]}"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="{step["bottom"]}"/></a:gs>'
        f'</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
        f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{step["bottom"]}"/></a:solidFill></a:ln>'
        f"</p:spPr>"
        f'<p:txBody><a:bodyPr lIns="{lins}" tIns="45720" rIns="0" bIns="45720" anchor="ctr"/>'
        f"<a:lstStyle/>"
        f'<a:p><a:pPr algn="ctr"><a:lnSpc><a:spcPct val="105000"/></a:lnSpc>'
        f'<a:buNone/></a:pPr>{lines}</a:p></p:txBody>'
        f"</p:sp>"
    )


def detail_xml(shape_id, items, left, top, width, height):
    lvl = SPEC["body_levels"]
    paras = []
    for n, item in enumerate(items):
        style = lvl[f'lvl{item["lvl"]}']
        if item["lvl"] == 2:
            lnspc, spcbef = 2160, (0 if n == 0 else 600)
            bufont, buchar, sz = "Wingdings", "l", 1400
            bufont_attrs = 'panose="05000000000000000000" pitchFamily="2" charset="2"'
        else:
            lnspc, spcbef = 1750, 200
            bufont, buchar, sz = FONT, "-", 1300
            bufont_attrs = FONT_ATTRS
        b = ' b="1"' if item.get("bold") else ""
        paras.append(
            f'<a:p><a:pPr marL="{style["marL"]}" indent="{style["indent"]}">'
            f'<a:lnSpc><a:spcPts val="{lnspc}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="{spcbef}"/></a:spcBef>'
            f'<a:buFont typeface="{bufont}" {bufont_attrs}/><a:buChar char="{escape(buchar)}"/></a:pPr>'
            f'<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="{sz}"{b} spc="-30" dirty="0">'
            f'<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>{font_tags()}</a:rPr>'
            f"<a:t>{escape(item['text'])}</a:t></a:r></a:p>"
        )
    return parse_xml(
        f'<p:sp {nsdecls("p", "a")}>'
        f'<p:nvSpPr><p:cNvPr id="{shape_id}" name="단계별 세부내용"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f"<p:spPr>"
        f'<a:xfrm><a:off x="{left}" y="{top}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        f"</p:spPr>"
        f'<p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="0" rIns="91440" bIns="0">'
        f"<a:normAutofit/></a:bodyPr><a:lstStyle/>" + "".join(paras) + "</p:txBody></p:sp>"
    )


def build(content_path, out_path):
    content = json.loads(Path(content_path).read_text(encoding="utf-8"))
    prs = Presentation(str(TEMPLATE))
    delete_all_slides(prs)
    layout = get_layout(prs, content.get("layout", SPEC["body_slide_layout"]))
    slide = prs.slides.add_slide(layout)

    # 1) 제목 (레이아웃 스타일 그대로 상속)
    slide.shapes.title.text_frame.paragraphs[0].add_run().text = content["title"]

    # 2) 본문 플레이스홀더 lvl1 = ❖ 대분류
    body = next(
        ph for ph in slide.placeholders if ph.element.find(f".//{qn('p:ph')}") is not None
        and ph.element.find(f".//{qn('p:ph')}").get("idx") == "10"
    )
    tf = body.text_frame
    tf.paragraphs[0].add_run().text = content["heading"]

    # 3) 공정 셰브런 도해
    band_top = px_y(BAND_TOP_PX)
    band_h = px_y(BAND_BOTTOM_PX - BAND_TOP_PX)
    sid = 100
    for i, (step, box) in enumerate(zip(content["process"]["steps"], STEP_BOX_PX)):
        slide.shapes._spTree.append(chevron_xml(sid + i, i, step, box, band_top, band_h))

    # 4) 단계별 세부내용 (본문 플레이스홀더와 좌측 기준선 동일)
    ph_shape = next(
        s for s in layout.shapes
        if s.element.find(f".//{qn('p:ph')}") is not None
        and s.element.find(f".//{qn('p:ph')}").get("idx") == "10"
    )
    slide.shapes._spTree.append(
        detail_xml(200, content["detail"], ph_shape.left, DETAIL_TOP,
                   ph_shape.width, DETAIL_BOTTOM - DETAIL_TOP)
    )

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    print(f"wrote {out_path}")


if __name__ == "__main__":
    build(sys.argv[1], sys.argv[2])
