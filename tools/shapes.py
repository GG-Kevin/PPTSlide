#!/usr/bin/env python3
"""슬라이드 도형을 원시 DrawingML XML 로 만드는 헬퍼.

python-pptx 의 고수준 API 는 그라데이션·둥근모서리 조절값·문단별 불릿을
세밀하게 못 다루므로, 필요한 도형은 XML 로 직접 만든다.
좌표 단위는 전부 EMU.
"""
from xml.sax.saxutils import escape

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

EMU_IN = 914400
PT = 12700  # 1pt = 12700 EMU

_ID = [1000]


def _next_id():
    _ID[0] += 1
    return _ID[0]


def reset_ids():
    _ID[0] = 1000


def font_tags(font):
    return (f'<a:latin typeface="{font}"/><a:ea typeface="{font}"/>'
            f'<a:cs typeface="Arial"/>')


def solid(hex_or_scheme):
    if hex_or_scheme.startswith("+"):
        return f'<a:solidFill><a:schemeClr val="{hex_or_scheme[1:]}"/></a:solidFill>'
    return f'<a:solidFill><a:srgbClr val="{hex_or_scheme}"/></a:solidFill>'


def grad(c1, c2, angle=5400000, stops=None):
    """선형 그라데이션. stops 를 주면 (pos, hex) 목록을 그대로 쓴다."""
    if stops is None:
        stops = [(0, c1), (100000, c2)]
    gs = "".join(f'<a:gs pos="{p}"><a:srgbClr val="{c}"/></a:gs>' for p, c in stops)
    return (f'<a:gradFill flip="none" rotWithShape="1"><a:gsLst>{gs}</a:gsLst>'
            f'<a:lin ang="{angle}" scaled="0"/></a:gradFill>')


def run(text, *, font, size, color, bold=False, spc=-20):
    b = ' b="1"' if bold else ""
    return (f'<a:r><a:rPr lang="ko-KR" altLang="en-US" sz="{size}"{b} spc="{spc}" dirty="0">'
            f"{solid(color)}{font_tags(font)}</a:rPr><a:t>{escape(text)}</a:t></a:r>")


def para(runs, *, align="l", line_pts=None, before_pts=None, marL=0, indent=0,
         bullet=None, bullet_font=None, bullet_color=None, bullet_size=None):
    bits = []
    if line_pts:
        bits.append(f'<a:lnSpc><a:spcPts val="{int(line_pts * 100)}"/></a:lnSpc>')
    if before_pts:
        bits.append(f'<a:spcBef><a:spcPts val="{int(before_pts * 100)}"/></a:spcBef>')
    if bullet:
        if bullet_color:
            bits.append(f'<a:buClr><a:srgbClr val="{bullet_color}"/></a:buClr>')
        if bullet_size:
            bits.append(f'<a:buSzPct val="{bullet_size}"/>')
        bits.append(f'<a:buFont typeface="{bullet_font or "Arial"}"/>'
                    f'<a:buChar char="{escape(bullet)}"/>')
    else:
        bits.append("<a:buNone/>")
    return (f'<a:p><a:pPr marL="{marL}" indent="{indent}" algn="{align}">'
            f'{"".join(bits)}</a:pPr>{"".join(runs)}</a:p>')


def shape(*, name, prst, x, y, cx, cy, fill="", line="", adj=None, body="",
          text_paras=None, anchor="ctr", ins=(0, 0, 0, 0), rot=None,
          shadow=""):
    av = (f'<a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst>'
          if adj is not None else "<a:avLst/>")
    li, ti, ri, bi = ins
    tx = ""
    if text_paras is not None:
        tx = (f'<p:txBody><a:bodyPr wrap="square" lIns="{li}" tIns="{ti}" rIns="{ri}" '
              f'bIns="{bi}" anchor="{anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
              f'{"".join(text_paras)}</p:txBody>')
    else:
        tx = ('<p:txBody><a:bodyPr wrap="none" anchor="ctr"/><a:lstStyle/>'
              '<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:endParaRPr lang="ko-KR"/></a:p></p:txBody>')
    rotattr = f' rot="{rot}"' if rot else ""
    return parse_xml(
        f'<p:sp {nsdecls("p", "a")}>'
        f'<p:nvSpPr><p:cNvPr id="{_next_id()}" name="{escape(name)}"/>'
        f"<p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        f"<p:spPr>"
        f'<a:xfrm{rotattr}><a:off x="{int(x)}" y="{int(y)}"/>'
        f'<a:ext cx="{int(cx)}" cy="{int(cy)}"/></a:xfrm>'
        f'<a:prstGeom prst="{prst}">{av}</a:prstGeom>'
        f'{fill or "<a:noFill/>"}{line or "<a:ln><a:noFill/></a:ln>"}{shadow}'
        f"{body}</p:spPr>{tx}</p:sp>"
    )


def textbox(*, name, x, y, cx, cy, paras, anchor="t", ins=(0, 0, 0, 0)):
    li, ti, ri, bi = ins
    return parse_xml(
        f'<p:sp {nsdecls("p", "a")}>'
        f'<p:nvSpPr><p:cNvPr id="{_next_id()}" name="{escape(name)}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(x)}" y="{int(y)}"/>'
        f'<a:ext cx="{int(cx)}" cy="{int(cy)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square" lIns="{li}" tIns="{ti}" rIns="{ri}" bIns="{bi}" '
        f'anchor="{anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>'
        f'{"".join(paras)}</p:txBody></p:sp>'
    )


def rule(*, name, x, y, cx, color, thickness=9525):
    """가는 구분선 (얇은 사각형으로 그린다 — 렌더러 간 굵기 차가 없다)."""
    return shape(name=name, prst="rect", x=x, y=y, cx=cx, cy=thickness,
                 fill=solid(color))


def soft_shadow(blur=101600, dist=25400, alpha=14000):
    return (f'<a:effectLst><a:outerShdw blurRad="{blur}" dist="{dist}" dir="5400000" '
            f'rotWithShape="0"><a:srgbClr val="1B2A41"><a:alpha val="{alpha}"/>'
            f"</a:srgbClr></a:outerShdw></a:effectLst>")
