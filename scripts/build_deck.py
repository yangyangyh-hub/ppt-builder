# -*- coding: utf-8 -*-
"""按 JSON 页面单生成 PPT 骨架（内置本套视觉规格）。

用法：
    python build_deck.py spec.json [-o out.pptx]

spec.json 见 references/06-spec-schema.md。不传 template 时用 assets/template.pptx。

内置规格（改这里就等于改全篇，不要在页面代码里硬编码字号）：
    画幅 33.87×19.05   内容区 x=1.70…32.10   y=3.50…18.60
    标题 34 / 副标 24 / 正文 24·20 / 表格 20·18 / 图注·页码 16
"""
import argparse
import json
import math
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Cm, Pt
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
DEFAULT_TEMPLATE = os.path.join(SKILL, "assets", "template.pptx")
DEFAULT_LOGO = os.path.join(SKILL, "assets", "school-logo.gif")

# ---------------------------------------------------------------- 规格常量
PAGE_W, PAGE_H = 33.87, 19.05
MARGIN_L, MARGIN_R = 1.70, 32.10
CONTENT_W = MARGIN_R - MARGIN_L          # 30.40
BODY_TOP = 3.50
LINE_Y = 18.80
BANNER_Y, BANNER_H = 17.45, 0.85
TRI_W, TRI_STEP = 9.55, 10.32            # 三列宽与步长
TWO_W, TWO_STEP = 14.20, 15.15           # 两列宽与步长

INK = RGBColor(0x06, 0x06, 0x07)
RED = RGBColor(0xB5, 0x25, 0x25)
KEY = RGBColor(0xFF, 0x00, 0x00)
GREY = RGBColor(0x8C, 0x8C, 0x8C)
PAGE_GREY = RGBColor(0x9A, 0x9A, 0x9A)
CARD = RGBColor(0xF2, 0xF0, 0xEB)
TINT = RGBColor(0xF6, 0xF5, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
HAIR = RGBColor(0xD8, 0xD8, 0xD8)

FS_TITLE, FS_SUB, FS_BODY, FS_BODY_S = 34, 24, 24, 20
FS_TABLE, FS_TABLE_S, FS_CAP, FS_PAGE = 20, 18, 16, 16

FONT_CN, FONT_EN = "微软雅黑", "Times New Roman"


# ---------------------------------------------------------------- 基础工具
def _style_run(run, size, color=INK, bold=True, font_cn=FONT_CN):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color
    f.name = FONT_EN
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", font_cn)
    rPr.set("lang", "zh-CN")          # 不写会按西文断行，标点甩到行首
    rPr.set("altLang", "en-US")


CUR_PAGE = [0]
WARN = []


def est_height(blocks, w):
    """按 CJK 全宽 / ASCII 半宽估算文本实际占高（cm），用于提前发现溢出。"""
    total = 0.0
    for b in blocks:
        if isinstance(b, str):
            b = {"text": b}
        size = b.get("size", FS_BODY)
        line = b.get("line", 1.5)
        txt = b.get("text") or "".join(r[0] for r in b.get("runs", []))
        units = sum(1.0 if ord(c) > 0x2E80 else 0.55 for c in txt)
        per_line = max(1.0, w / (size * 0.0353 * 1.02))
        lines = max(1, math.ceil(units / per_line))
        total += lines * size * 0.0353 * 1.30 * line + (b.get("space_before") or 0) * 0.0353
    return total


def text_box(slide, l, t, w, h, blocks, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, wrap=True, warn=True):
    """blocks: str | list[str | dict]

    dict 形如 {"text":..., "size":24, "color":"B52525", "bold":True,
               "space_before":6, "line":1.5, "runs":[(text, size, color)]}
    """
    tb = slide.shapes.add_textbox(Cm(l), Cm(t), Cm(w), Cm(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None                     # 关键：不让框自长压住下方
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    if isinstance(blocks, (str, dict)):
        blocks = [blocks]
    for i, b in enumerate(blocks):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", "0")
        pPr.set("indent", "0")
        if isinstance(b, str):
            b = {"text": b}
        p.line_spacing = b.get("line", 1.5)
        if b.get("space_before") is not None:
            p.space_before = Pt(b["space_before"])
        size = b.get("size", FS_BODY)
        color = b.get("color", INK)
        color = RGBColor.from_string(color) if isinstance(color, str) else color
        bold = b.get("bold", True)
        runs = b.get("runs") or [(b.get("text", ""), size, color)]
        for rt in runs:
            r = p.add_run()
            r.text = rt[0]
            _style_run(r, rt[1] if len(rt) > 1 and rt[1] else size,
                       rt[2] if len(rt) > 2 and rt[2] else color, bold)
    need = est_height(blocks, w)
    if warn and h > 0.8 and need > h * 1.20:
        head = (blocks[0].get("text", "") if isinstance(blocks[0], dict) else blocks[0])[:16]
        WARN.append(f"P{CUR_PAGE[0]} 文本可能溢出：需≈{need:.2f}cm / 框高{h:.2f}cm @T{t} 「{head}」")
    return tb


def rect(slide, l, t, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=1.0):
    s = slide.shapes.add_shape(shape, Cm(l), Cm(t), Cm(w), Cm(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
    s.shadow.inherit = False
    return s


def pic_fill(slide, path, l, t, w, h):
    """等比裁切填满目标框（不拉伸）。"""
    pic = slide.shapes.add_picture(path, Cm(l), Cm(t), Cm(w), Cm(h))
    iw, ih = pic.image.size
    ir, tr = iw / ih, w / h
    pic.crop_left = pic.crop_right = pic.crop_top = pic.crop_bottom = 0.0
    if ir > tr:
        pic.crop_left = pic.crop_right = (1 - tr / ir) / 2
    elif ir < tr:
        pic.crop_top = pic.crop_bottom = (1 - ir / tr) / 2
    pic.width, pic.height = Cm(w), Cm(h)
    return pic


def new_slide(prs):
    blank = min(prs.slide_layouts, key=lambda ly: len(ly.placeholders))
    s = prs.slides.add_slide(blank)
    for ph in list(s.placeholders):
        ph._element.getparent().remove(ph._element)
    # 模板的深色底常是母版里的形状而非 bg 填充，只设 background 盖不住。
    # 铺一张白底矩形并压到最底层，才能彻底盖掉母版装饰。
    white = rect(s, 0, 0, PAGE_W, PAGE_H, WHITE)
    spTree = s.shapes._spTree
    spTree.remove(white._element)
    spTree.insert(2, white._element)
    return s


def chrome(slide, page_no, chapter=None, title=None, sub=None, page_on_banner=False):
    """章节块 / 标题 / 副标 / 校徽 / 底线 / 页码。"""
    if chapter is not None:
        rect(slide, 0.33, 0.53, 1.96, 1.24, RED, shape=MSO_SHAPE.PENTAGON)
        text_box(slide, 0.33, 0.53, 1.96, 1.24, [{"text": str(chapter), "size": 30,
                 "line": 1.0}], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        for r in slide.shapes[-1].text_frame.paragraphs[0].runs:
            r.font.color.rgb = WHITE
    if title:
        text_box(slide, 2.77, 0.52, 22.60, 1.30, [{"text": title, "size": FS_TITLE,
                 "line": 1.0}], warn=False)
    if sub:
        text_box(slide, 2.82, 1.95, 22.00, 1.02, [{"text": sub, "size": FS_SUB,
                 "color": RED, "line": 1.0}], warn=False)
    if os.path.exists(DEFAULT_LOGO):
        slide.shapes.add_picture(DEFAULT_LOGO, Cm(26.30), Cm(0.42), Cm(7.16), Cm(2.04))
    ln = rect(slide, 0, LINE_Y, PAGE_W, 0, None)   # 占位后改为线
    ln.width, ln.height = Cm(PAGE_W), Cm(0)
    ln.line.color.rgb = RED
    ln.line.width = Pt(1.0)
    if page_no:
        if page_on_banner:
            # 通栏页：页码压在红条里反白，否则会被红条盖掉（原版就丢在这一步）
            text_box(slide, 29.90, BANNER_Y, 2.20, BANNER_H,
                     [{"text": str(page_no), "size": FS_PAGE, "color": WHITE,
                       "line": 1.0}], align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        else:
            text_box(slide, 30.40, 17.90, 1.80, 0.55,
                     [{"text": str(page_no), "size": FS_PAGE, "color": PAGE_GREY,
                       "line": 1.0}], align=PP_ALIGN.RIGHT)


def frame(slide, logo=True, line=True, logo_box=(1.70, 0.60, 7.40, 2.11)):
    if logo and os.path.exists(DEFAULT_LOGO):
        slide.shapes.add_picture(DEFAULT_LOGO, Cm(logo_box[0]), Cm(logo_box[1]),
                                 Cm(logo_box[2]), Cm(logo_box[3]))
    if line:
        ln = rect(slide, 0, LINE_Y, PAGE_W, 0)
        ln.width, ln.height = Cm(PAGE_W), Cm(0)
        ln.line.color.rgb = RED
        ln.line.width = Pt(1.0)


def banner(slide, text, t=BANNER_Y):
    rect(slide, MARGIN_L, t, CONTENT_W, BANNER_H, RED)
    text_box(slide, MARGIN_L, t, CONTENT_W, BANNER_H,
             [{"text": text, "size": FS_BODY_S, "color": WHITE, "line": 1.0}],
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def caption(slide, l, t, w, text, align=PP_ALIGN.LEFT):
    text_box(slide, l, t, w, 0.85, [{"text": text, "size": FS_CAP, "color": GREY,
             "line": 1.0}], align=align, warn=False)


# ---------------------------------------------------------------- 表格
def _border(cell, color=HAIR, width=0.75, edges=("B",)):
    tcPr = cell._tc.get_or_add_tcPr()
    for e in edges:
        tag = qn(f"a:ln{e}")
        old = tcPr.find(tag)
        if old is not None:
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, tag)
        ln.set("w", str(int(width * 12700)))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        fill = etree.SubElement(ln, qn("a:solidFill"))
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", str(color))


def _cell(cell, text, size, color=INK, bold=False, fill=None, align=PP_ALIGN.CENTER):
    cell.margin_left = Cm(0.18)
    cell.margin_right = Cm(0.18)
    cell.margin_top = Cm(0.05)
    cell.margin_bottom = Cm(0.05)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    p.line_spacing = 1.0
    r = p.add_run()
    r.text = str(text)
    _style_run(r, size, color, bold)
    if fill is None:
        cell.fill.background()
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


def table(slide, spec):
    cols, rows = spec["cols"], spec["rows"]
    l, t = spec.get("l", MARGIN_L), spec.get("t", 3.30)
    w = spec.get("w", CONTENT_W)
    colw = spec.get("colw") or [w / len(cols)] * len(cols)
    k = w / sum(colw)
    colw = [c * k for c in colw]
    rowh = spec.get("rowh") or 0.80
    body_sz = spec.get("size", FS_TABLE)
    head_sz = spec.get("head_size", body_sz)
    # 行高下限：字号换算成厘米 × 1.30，否则 PowerPoint 会自己撑高把下方元素顶掉
    rowh = max(rowh, round(max(body_sz, head_sz) * 0.0353 * 1.30, 2))
    gt = slide.shapes.add_table(len(rows) + 1, len(cols), Cm(l), Cm(t),
                                Cm(sum(colw)), Cm(rowh * (len(rows) + 1)))
    tb = gt.table
    tb.first_row = False
    tb.horz_banding = False
    for i, cw in enumerate(colw):
        tb.columns[i].width = Cm(cw)
    for i in range(len(rows) + 1):
        tb.rows[i].height = Cm(rowh)
    for j, c in enumerate(cols):
        _cell(tb.cell(0, j), c, head_sz, WHITE, True, RED,
              PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER)
        _border(tb.cell(0, j), RED, 1.0, ("B",))
    for i, row in enumerate(rows, 1):
        last = i == len(rows) and spec.get("total_row", True)
        for j, v in enumerate(row):
            al = PP_ALIGN.LEFT if (j == 0 or not str(v).replace("—", "").strip("-").replace(".", "").isdigit()) else PP_ALIGN.CENTER
            _cell(tb.cell(i, j), v, body_sz, INK, last,
                  TINT if last else None, al)
            _border(tb.cell(i, j), HAIR, 0.75, ("B",))
    return gt

# ---------------------------------------------------------------- 页面原型
def pg_cover(prs, slide, sp, no):
    frame(slide, logo=sp.get("logo", True), line=True)
    text_box(slide, 1.70, 5.20, 30.40, 2.20,
             [{"text": sp["title"], "size": 44, "line": 1.15}])
    ln = rect(slide, 1.70, 7.85, 11.20, 0.01, None)
    ln.line.color.rgb = RED
    ln.line.width = Pt(2.0)
    if sp.get("kicker"):
        text_box(slide, 1.70, 8.30, 30.40, 1.10,
                 [{"text": sp["kicker"], "size": 24, "color": GREY, "line": 1.0}])
    y = 12.20
    for line in sp.get("byline", []):
        text_box(slide, 1.70, y, 24.00, 1.00, [{"text": line, "size": 24, "line": 1.0}])
        y += 1.55
    return slide


def pg_toc(prs, slide, sp, no):
    text_box(slide, 1.70, 1.45, 9.00, 2.00,
             [{"text": sp.get("title", "目录"), "size": 54, "line": 1.0}], warn=False)
    text_box(slide, 1.75, 3.45, 6.00, 0.80,
             [{"text": sp.get("label", "C O N T E N T"), "size": FS_BODY_S, "line": 1.0}])
    y = 3.35
    for i, it in enumerate(sp.get("items", []), 1):
        cy = slide.shapes.add_shape(MSO_SHAPE.OVAL, Cm(15.00), Cm(y), Cm(1.10), Cm(1.10))
        cy.fill.solid()
        cy.fill.fore_color.rgb = RED
        cy.line.fill.background()
        cy.shadow.inherit = False
        tf = cy.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = f"{i:02d}"
        _style_run(r, FS_PAGE, WHITE)
        text_box(slide, 17.20, y, 14.00, 1.10,
                 [{"text": it, "size": FS_SUB, "line": 1.0}], anchor=MSO_ANCHOR.MIDDLE)
        y += 2.12
    return slide


def pg_tri(prs, slide, sp, no):
    """三条并列 + 三图横排 + 通栏。"""
    top = sp.get("top", 3.65)
    colh = sp.get("colh", 5.62)
    bar = sp.get("bar", False)   # 三列默认不加竖条：加条会吃掉 0.35cm 字宽，正中一行变 10 字
    for i, c in enumerate(sp.get("cols", [])[:3]):
        l = MARGIN_L + i * TRI_STEP
        if bar:
            rect(slide, l, top, 0.10, colh, RED)
        text_box(slide, l + (0.35 if bar else 0), top, TRI_W - (0.35 if bar else 0), colh, [
            {"text": c.get("lead", ""), "size": FS_BODY, "color": RED, "line": 1.15},
            {"text": c.get("body", ""), "size": FS_BODY, "line": sp.get("body_line", 1.0),
             "space_before": 4},
        ])
    if sp.get("figs"):
        fy = sp.get("fig_top", top + colh + 0.20)
        fh = sp.get("fig_h", sp.get("fig_bottom", BANNER_Y - 1.00) - fy)
        for i, f in enumerate(sp["figs"][:3]):
            l = MARGIN_L + i * TRI_STEP
            if f.get("path") and os.path.exists(f["path"]):
                pic_fill(slide, f["path"], l, fy, TRI_W, fh)
            else:
                rect(slide, l, fy, TRI_W, fh, TINT, HAIR)
                text_box(slide, l, fy, TRI_W, fh,
                         [{"text": f.get("caption", "图位"), "size": FS_BODY_S,
                           "color": GREY}], align=PP_ALIGN.CENTER,
                         anchor=MSO_ANCHOR.MIDDLE)
            caption(slide, l, fy + fh + 0.05, TRI_W, f.get("caption", ""))
    if sp.get("banner"):
        banner(slide, sp["banner"])
    return slide


def pg_flow(prs, slide, sp, no):
    """中心句 + 横向链条 + 2×2 产出块 + 通栏。"""
    if sp.get("claim"):
        text_box(slide, MARGIN_L, sp.get("claim_top", 3.60), CONTENT_W, 1.02,
                 [{"text": sp["claim"], "size": FS_BODY, "color": RED, "line": 1.0}])
    chips = sp.get("chips", [])
    if chips:
        gap = sp.get("chip_gap", 0.30)
        cw = (CONTENT_W - gap * (len(chips) - 1)) / len(chips)
        cy, ch = sp.get("chip_top", 5.00), sp.get("chip_h", 1.15)
        for i, c in enumerate(chips):
            l = MARGIN_L + i * (cw + gap)
            last = i == len(chips) - 1
            rect(slide, l, cy, cw, ch, RED if last else CARD)
            text_box(slide, l, cy, cw, ch,
                     [{"text": c, "size": FS_BODY_S,
                       "color": WHITE if last else INK, "line": 1.0}],
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y = cy + ch + 0.80
    else:
        y = sp.get("block_top", 6.95)
    bw, bh = sp.get("block_w", 14.20), sp.get("block_h", 4.30)
    for i, b in enumerate(sp.get("blocks", [])[:4]):
        l = sp.get("block_l", 2.25) + (i % 2) * sp.get("block_step", 15.15)
        t = y + (i // 2) * (bh + sp.get("block_gap", 1.05))
        rect(slide, l - 0.55, t + 0.10, 0.10, bh - 0.60, RED)
        text_box(slide, l, t, bw, bh, [
            {"text": b.get("lead", ""), "size": FS_BODY, "color": RED, "line": 1.0},
            {"text": b.get("body", ""), "size": FS_BODY_S, "line": 1.20, "space_before": 2},
        ])
    if sp.get("banner"):
        banner(slide, sp["banner"], sp.get("banner_t", 17.55))
    return slide


def pg_blocks(prs, slide, sp, no):
    """2×2 卡片（红短竖条）+ 底部 1–2 张图。"""
    top = sp.get("top", 3.05)
    bw, bh = sp.get("card_w", 14.00), sp.get("card_h", 3.70)
    gap = sp.get("card_gap", 2.00)
    for i, c in enumerate(sp.get("cards", [])[:4]):
        l = sp.get("card_l", 2.05) + (i % 2) * (bw + gap)
        t = top + (i // 2) * (bh + sp.get("card_vgap", 0.95))
        rect(slide, l - 0.55, t + 0.10, 0.10, bh - 0.20, RED)
        text_box(slide, l, t, bw, bh, [
            {"text": c.get("lead", ""), "size": FS_BODY, "color": RED, "line": 1.0},
            {"text": c.get("body", ""), "size": sp.get("body_size", FS_BODY),
             "line": sp.get("body_line", 1.0), "space_before": 4},
        ])
    pics = sp.get("pics", [])
    if pics:
        py = sp.get("pic_top", top + 2 * (bh + sp.get("card_vgap", 1.00)) + 0.20)
        ph = sp.get("pic_h", LINE_Y - 0.90 - py)
        n = len(pics)
        gap = sp.get("pic_gap", 1.20)
        pw = min(sp.get("pic_w", 10.10), (CONTENT_W - gap * (n - 1)) / n)
        total = pw * n + gap * (n - 1)
        x0 = MARGIN_L + (CONTENT_W - total) / 2
        for i, f in enumerate(pics):
            l = x0 + i * (pw + gap)
            path = f if isinstance(f, str) else f.get("path")
            if path and os.path.exists(path):
                pic_fill(slide, path, l, py, pw, ph)
            if isinstance(f, dict) and f.get("caption"):
                caption(slide, l, py + ph + 0.05, pw, f["caption"])
    return slide


def pg_table(prs, slide, sp, no):
    tspec = sp["table"]
    if sp.get("lead"):
        text_box(slide, MARGIN_L, sp.get("lead_top", 2.80), CONTENT_W,
                 sp.get("lead_h", 1.50),
                 [{"text": sp["lead"], "size": FS_BODY, "line": 1.4}])
    table(slide, tspec)
    for i, n in enumerate(sp.get("footnotes", [])):   # 不要用 notes，那是演讲备注
        caption(slide, tspec.get("l", MARGIN_L),
                sp.get("note_top", 15.20) + i * 0.62, tspec.get("w", CONTENT_W), n)
    side = sp.get("side")
    if side:
        # side 可以是单个 dict，也可以是 list（多栏并排）；多栏时不写 l 就按两列网格排
        slist = side if isinstance(side, list) else [side]
        for si, sd in enumerate(slist):
            dflt_l = (18.40 if not isinstance(side, list)
                      else MARGIN_L + si * (14.00 + 1.55))
            text_box(slide, sd.get("l", dflt_l), sd.get("t", 3.20),
                     sd.get("w", 14.00 if isinstance(side, list) else 13.20),
                     sd.get("h", 6.00),
                     [{"text": sd.get("lead", ""), "size": FS_BODY, "color": RED,
                       "line": 1.2}] +
                     [{"text": p, "size": sd.get("size", FS_BODY_S), "line": 1.45,
                       "space_before": 4} for p in sd.get("body", [])])
    for p in sp.get("pics", []):
        pic_fill(slide, p["path"], p.get("l", 20.30), p.get("t", 4.90),
                 p.get("w", 11.40), p.get("h", 7.00))
        if p.get("caption"):
            caption(slide, p.get("l", 20.30), p.get("t", 4.90) + p.get("h", 7.00) + 0.05,
                    p.get("w", 11.40), p["caption"])
    return slide


def pg_image(prs, slide, sp, no):
    w = sp.get("w", 21.98)
    t = sp.get("t", 3.30)
    h = sp.get("h", LINE_Y - 0.95 - t)
    l = sp.get("l", (PAGE_W - w) / 2)
    pic_fill(slide, sp["path"], l, t, w, h)
    if sp.get("caption"):
        caption(slide, l, t + h + 0.05, w, sp["caption"],
                PP_ALIGN.CENTER if sp.get("center_cap") else PP_ALIGN.LEFT)
    return slide


def pg_custom(prs, slide, sp, no):
    for it in sp.get("items", []):
        if it.get("kind") == "text":
            text_box(slide, it.get("l", MARGIN_L), it["t"], it.get("w", CONTENT_W),
                     it.get("h", 1.0),
                     [{"text": it["text"], "size": it.get("size", FS_BODY),
                       "color": it.get("color", "060607"),
                       "bold": it.get("bold", True), "line": it.get("line", 1.0)}],
                     align={"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                            "right": PP_ALIGN.RIGHT}[it.get("align", "left")])
        elif it.get("kind") == "rect":
            rect(slide, it.get("l", MARGIN_L), it["t"], it.get("w", CONTENT_W),
                 it.get("h", 1.0),
                 RGBColor.from_string(it["fill"]) if it.get("fill") else None)
        elif it.get("kind") == "pic":
            pic_fill(slide, it["path"], it.get("l", MARGIN_L), it["t"],
                     it.get("w", 10.0), it.get("h", 6.0))
    return slide


def pg_end(prs, slide, sp, no):
    frame(slide, logo=sp.get("logo", True), line=True)
    text_box(slide, MARGIN_L, 7.60, CONTENT_W, 1.80,
             [{"text": sp.get("title", "敬请各位专家指导！"), "size": 40, "line": 1.0}],
             align=PP_ALIGN.CENTER, warn=False)
    y = 12.20
    for line in sp.get("lines", []):
        text_box(slide, MARGIN_L, y, CONTENT_W, 1.00,
                 [{"text": line, "size": FS_BODY, "line": 1.0}], align=PP_ALIGN.CENTER)
        y += 1.55
    return slide


RENDER = {
    "cover": pg_cover, "toc": pg_toc, "tri": pg_tri, "flow": pg_flow,
    "blocks": pg_blocks, "table": pg_table, "image": pg_image,
    "custom": pg_custom, "end": pg_end,
}


# ---------------------------------------------------------------- 主流程
def drop_all_slides(prs):
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        prs.part.drop_rel(sid.get(qn("r:id")))
        lst.remove(sid)


def build(spec):
    tpl = spec.get("template") or DEFAULT_TEMPLATE
    prs = Presentation(tpl)
    drop_all_slides(prs)
    base = spec.get("start_page", 1)
    for i, s in enumerate(spec["pages"]):
        no = base + i
        CUR_PAGE[0] = no
        slide = new_slide(prs)
        t = s["type"]
        if t not in RENDER:
            raise SystemExit(f"未知页面类型: {t}（可用 {list(RENDER)}）")
        if t in ("cover", "toc", "end"):
            RENDER[t](prs, slide, s, no)
        else:
            chrome(slide, no, s.get("ch"), s.get("title"), s.get("sub"),
                   page_on_banner=bool(s.get("banner")) and t in ("tri", "flow"))
            RENDER[t](prs, slide, s, no)
        if isinstance(s.get("notes"), str) and s["notes"]:
            slide.notes_slide.notes_text_frame.text = s["notes"]
    out = spec.get("out") or "out.pptx"
    prs.save(out)
    if WARN:
        print("⚠ 版式预警（按估算，务必渲染复核）：")
        for w in WARN:
            print("  -", w)
    print(f"OK {len(spec['pages'])} 页 -> {out}")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("-o", "--out")
    a = ap.parse_args()
    spec = json.load(open(a.spec, encoding="utf-8"))
    if a.out:
        spec["out"] = a.out
    build(spec)


if __name__ == "__main__":
    main()
