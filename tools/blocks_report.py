#!/usr/bin/env python3
"""보고자료용 확장 블록.

blocks.py 의 기본 블록(image·table·text·note·panels·flow)으로는 표현되지 않는
네 가지를 추가한다. 1장 통합본처럼 좁은 열에 밀도를 높여야 할 때 쓴다.

  colhead : 열 소제목 바      — 3열 구성에서 열의 시작을 알린다
  vsteps  : 세로 단계 목록    — 좁은 열에서 흐름·역할을 세로로 쌓는다
  cycle   : 폐쇄형 순환       — 마지막 단계가 처음으로 되돌아가는 것을 보인다
  strip   : 상단 논리 띠      — 문서 전체 스토리라인을 한 줄로 깐다

도형 이름 규칙
  '카드 ' / '패널 ' 접두사를 쓰지 않는다. verify.py 가 그 접두사를 가로 배치로
  간주해 좌우 겹침을 검사하는데, 여기 블록들은 세로로 쌓이므로 오탐이 난다.
"""
import shapes as S
import blocks as B


# ── 열 소제목 바 ──────────────────────────────────────────────────────────
def add_colhead(slide, spec, box):
    x, y, cx, cy = B.emu_box(box)
    color = spec.get("color", B.BRAND_NAVY)
    key = spec.get("key", spec["label"][:8])
    add = slide.shapes._spTree.append
    add(S.shape(name=f"열머리표식 {key}", prst="rect",
                x=x, y=y + int(cy * 0.10), cx=52000, cy=int(cy * 0.60),
                fill=S.solid(color)))
    add(S.textbox(name=f"열머리 {key}", x=x + 100000, y=y,
                  cx=cx - 100000, cy=int(cy * 0.82),
                  paras=[S.para([S.run(spec["label"], font=B.FONT, size=1100,
                                       color=B.INK, bold=True)])],
                  anchor="ctr"))
    add(S.rule(name=f"열머리선 {key}", x=x, y=y + cy - 9525, cx=cx, color=color))


# ── 세로 단계 목록 ────────────────────────────────────────────────────────
def add_vsteps(slide, spec, box):
    x, y, cx, cy = B.emu_box(box)
    items = spec["items"]
    n = len(items)
    pitch = cy / n
    badge = int(min(292000, pitch * 0.54))
    tsz = spec.get("title_size", 1000)
    dsz = spec.get("desc_size", 900)
    add = slide.shapes._spTree.append
    for i, it in enumerate(items):
        iy = int(y + pitch * i)
        color = it.get("color", B.BRAND_NAVY)
        dark = it.get("dark", color)
        if it.get("band"):
            add(S.shape(name=f"스텝바탕 {i + 1}", prst="roundRect", adj=7000,
                        x=x, y=iy, cx=cx, cy=int(pitch - 38000),
                        fill=S.solid(it["band"]),
                        line=f'<a:ln w="9525">{S.solid(it.get("band_ln", B.CARD_LN))}</a:ln>'))
        add(S.shape(name=f"스텝배지 {i + 1}", prst="ellipse",
                    x=x + 40000, y=iy + 46000, cx=badge, cy=badge,
                    fill=S.grad(color, dark, angle=2700000),
                    text_paras=[S.para([S.run(it.get("no", str(i + 1)), font=B.FONT,
                                              size=850, color="FFFFFF", bold=True)],
                                       align="ctr")]))
        tx = x + 40000 + badge + 95000
        tw = x + cx - tx - 50000
        # 설명이 없으면 제목을 세로 가운데에 둔다 (좁은 열에서 제목만 쓰는 경우)
        ty = iy + 32000 if it.get("desc") else int(iy + (pitch - 232000) / 2)
        add(S.textbox(name=f"스텝제목 {i + 1}", x=tx, y=ty, cx=tw, cy=232000,
                      paras=[S.para([S.run(it["title"], font=B.FONT, size=tsz,
                                           color=it.get("title_color", B.INK),
                                           bold=True)], line_pts=13)]))
        if it.get("desc"):
            add(S.textbox(name=f"스텝설명 {i + 1}", x=tx, y=iy + 252000,
                          cx=tw, cy=int(pitch - 292000),
                          paras=[S.para([S.run(it["desc"], font=B.FONT, size=dsz,
                                               color=B.BODY)], line_pts=11.5)]))


# ── 폐쇄형 순환 ───────────────────────────────────────────────────────────
def add_cycle(slide, spec, box):
    """가로 단계 + 마지막에서 처음으로 되돌아오는 회귀 경로.

    직선 흐름도로는 '닫힌 고리'가 보이지 않는다. 회귀 경로를 실제로 그려
    프로젝트 전 기간 반복된다는 것을 도해 자체로 전달한다.
    """
    x, y, cx, cy = B.emu_box(box)
    nodes = spec["nodes"]
    n = len(nodes)
    gap = int(spec.get("gap", 0.38) * B.IN)
    node_h = int(cy * spec.get("node_ratio", 0.62))
    w = (cx - gap * (n - 1)) / n
    add = slide.shapes._spTree.append

    for i, nd in enumerate(nodes):
        nx = x + (w + gap) * i
        color = nd.get("color", B.BRAND_NAVY)
        dark = nd.get("dark", color)
        paras = [S.para([S.run(nd["label"], font=B.FONT, size=1150,
                               color="FFFFFF", bold=True)], align="ctr", line_pts=14.5)]
        if nd.get("sub"):
            paras.append(S.para([S.run(nd["sub"], font=B.FONT, size=880, color="E8EEF7")],
                                align="ctr", line_pts=11.5, before_pts=2))
        add(S.shape(name=f"루프 {i + 1}", prst="roundRect", adj=9000,
                    x=int(nx), y=y, cx=int(w), cy=node_h,
                    fill=S.grad(color, dark, angle=5400000),
                    line=f'<a:ln w="9525">{S.solid(dark)}</a:ln>',
                    text_paras=paras, ins=(45720, 45720, 45720, 45720),
                    shadow=S.soft_shadow()))
        if i < n - 1:
            ax = nx + w + (gap - 150000) / 2
            add(S.shape(name=f"루프화살 {i + 1}", prst="triangle", rot=5400000,
                        x=int(ax), y=int(y + node_h / 2 - 90000), cx=150000, cy=180000,
                        fill=S.solid(spec.get("arrow_color", "A9B4C6"))))

    lc = spec.get("return_color", "8A94A6")
    ry = int(y + node_h + (cy - node_h) * 0.40)
    lx, rx = int(x + w * 0.5), int(x + cx - w * 0.5)
    add(S.rule(name="회귀선", x=lx, y=ry, cx=rx - lx, color=lc))
    add(S.shape(name="회귀연결 우", prst="rect", x=rx, y=y + node_h,
                cx=9525, cy=ry - (y + node_h), fill=S.solid(lc)))
    add(S.shape(name="회귀연결 좌", prst="rect", x=lx, y=y + node_h + 130000,
                cx=9525, cy=ry - (y + node_h) - 130000, fill=S.solid(lc)))
    add(S.shape(name="회귀촉", prst="triangle", x=lx - 68000, y=y + node_h + 20000,
                cx=145000, cy=140000, fill=S.solid(lc)))
    if spec.get("return_label"):
        add(S.textbox(name="회귀라벨", x=lx, y=ry + 46000, cx=rx - lx,
                      cy=min(250000, y + cy - ry - 46000),
                      paras=[S.para([S.run(spec["return_label"], font=B.FONT,
                                           size=900, color=B.BODY, bold=True)],
                                    align="ctr")]))


# ── 상단 논리 띠 ──────────────────────────────────────────────────────────
def add_strip(slide, spec, box):
    x, y, cx, cy = B.emu_box(box)
    items = spec["items"]
    n = len(items)
    gap = int(spec.get("gap", 0.17) * B.IN)
    w = (cx - gap * (n - 1)) / n
    add = slide.shapes._spTree.append
    for i, it in enumerate(items):
        nx = x + (w + gap) * i
        color = it.get("color", B.BRAND_NAVY)
        add(S.shape(name=f"논리띠 {i + 1}", prst="roundRect", adj=11000,
                    x=int(nx), y=y, cx=int(w), cy=cy,
                    fill=S.solid(it.get("bg", "F2F6FC")),
                    line=f'<a:ln w="9525">{S.solid("DDE5F0")}</a:ln>',
                    text_paras=[S.para([
                        S.run(it["no"] + "   ", font=B.FONT, size=950,
                              color=color, bold=True),
                        S.run(it["text"], font=B.FONT, size=950, color=B.INK, bold=True),
                    ], align="ctr")], anchor="ctr", ins=(36000, 0, 36000, 0)))
        if i < n - 1:
            ax = nx + w + (gap - 92000) / 2
            add(S.shape(name=f"논리띠화살 {i + 1}", prst="triangle", rot=5400000,
                        x=int(ax), y=int(y + cy / 2 - 56000), cx=92000, cy=112000,
                        fill=S.solid("A9B4C6")))


RENDERERS = {
    "colhead": add_colhead,
    "vsteps": add_vsteps,
    "cycle": add_cycle,
    "strip": add_strip,
}


def render_block(slide, b):
    RENDERERS[b["type"]](slide, b, b["box"])
