#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 BTC Momentum 演示 PPT（视觉升级版）。

设计语言：暗色渐变背景 + 圆角毛玻璃卡片 + 阴影景深 + 渐变装饰线 +
          径向光晕 + KPI 大字卡片 + Noto Sans SC 多字重排版。
内容与 slides.html / TALK_SCRIPT.md 一一对应；回测数据冻结于 2026-09-07。

用法：python scripts/build_pptx.py
输出：BTC_Momentum_演示.pptx（项目根目录）
"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 配色
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BG_TOP    = "080C12"
BG_MID    = "0D1117"
BG_BOT    = "131A24"
CARD_TOP  = "161B22"
CARD_BOT  = "1A2030"
BORDER    = "30363D"
FG_HEX    = "E6EDF3"
FG2_HEX   = "8B949E"

FG   = RGBColor(0xE6, 0xED, 0xF3)
FG2  = RGBColor(0x8B, 0x94, 0x9E)
FG3  = RGBColor(0x6E, 0x76, 0x81)

GREEN  = RGBColor(0x3F, 0xB9, 0x50)
RED    = RGBColor(0xF8, 0x51, 0x49)
BLUE   = RGBColor(0x58, 0xA6, 0xFF)
ORANGE = RGBColor(0xD2, 0x99, 0x22)
PURPLE = RGBColor(0xBC, 0x8C, 0xFF)
CYAN   = RGBColor(0x56, 0xD3, 0xF0)

GREEN_HEX  = "3FB950"
RED_HEX    = "F85149"
BLUE_HEX   = "58A6FF"
ORANGE_HEX = "D29922"
PURPLE_HEX = "BC8CFF"
CYAN_HEX   = "56D3F0"

# 字体
FONT_TITLE  = "Noto Sans SC Medium"
FONT_BODY   = "Noto Sans SC DemiLight"
FONT_LIGHT  = "Noto Sans SC Light"
FONT_NUM    = "Segoe UI Semibold"
FONT_MONO   = "Consolas"

EMU_W = Inches(13.333)
EMU_H = Inches(7.5)

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]

slide_counter = [0]  # mutable counter for page numbering


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 底层 XML 工具
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _set_gradient_bg(slide, top=BG_TOP, mid=BG_MID, bot=BG_BOT):
    """给幻灯片设置三段线性渐变背景。"""
    cSld = slide._element
    # remove existing bg if any
    for old in cSld.findall(qn('p:bg')):
        cSld.remove(old)
    bg_xml = f'''
    <p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:bgPr>
        <a:gradFill>
          <a:gsLst>
            <a:gs pos="0"><a:srgbClr val="{top}"/></a:gs>
            <a:gs pos="50000"><a:srgbClr val="{mid}"/></a:gs>
            <a:gs pos="100000"><a:srgbClr val="{bot}"/></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="1"/>
        </a:gradFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>'''
    bg_el = parse_xml(bg_xml)
    cSld.insert(0, bg_el)


def _set_shape_gradient(shape, c_top, c_bot, angle=5400000):
    """给形状设置渐变填充（替换默认 solidFill）。"""
    spPr = shape._element.spPr
    for ch in list(spPr):
        if ch.tag.endswith('}solidFill') or ch.tag.endswith('}gradFill'):
            spPr.remove(ch)
    xml = f'''
    <a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="{c_top}"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="{c_bot}"/></a:gs>
      </a:gsLst>
      <a:lin ang="{angle}" scaled="1"/>
    </a:gradFill>'''
    spPr.insert(0, parse_xml(xml))


def _add_shadow(shape, blur=50800, dist=25400, alpha=35000):
    """给形状加外阴影。"""
    spPr = shape._element.spPr
    # remove old effectLst if present
    for old in spPr.findall(qn('a:effectLst')):
        spPr.remove(old)
    xml = f'''
    <a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:outerShdw blurRad="{blur}" dist="{dist}" dir="5400000" algn="ctr" rotWithShape="0">
        <a:srgbClr val="000000">
          <a:alpha val="{alpha}"/>
        </a:srgbClr>
      </a:outerShdw>
    </a:effectLst>'''
    spPr.append(parse_xml(xml))


def _set_rounded_corners(shape, radius=6000):
    """设置圆角矩形的圆角半径。"""
    sp = shape._element
    prstGeom = sp.find('.//' + qn('a:prstGeom'))
    if prstGeom is None:
        return
    avLst = prstGeom.find(qn('a:avLst'))
    if avLst is None:
        avLst = parse_xml(
            '<a:avLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
        prstGeom.append(avLst)
    else:
        for gd in list(avLst):
            avLst.remove(gd)
    gd = parse_xml(
        f'<a:gd xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        f' name="adj" fmla="val {radius}"/>')
    avLst.append(gd)


def _make_text_gradient(run_element, c1, c2):
    """给文本 run 设置渐变填色（用于大号标题数字）。"""
    rPr = run_element.find(qn('a:rPr'))
    if rPr is None:
        return
    # Remove solidFill
    for sf in rPr.findall(qn('a:solidFill')):
        rPr.remove(sf)
    grad_xml = f'''
    <a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="{c1}"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="{c2}"/></a:gs>
      </a:gsLst>
      <a:lin ang="5400000" scaled="1"/>
    </a:gradFill>'''
    rPr.append(parse_xml(grad_xml))


def _remove_table_borders(cell):
    """删除表格单元格的所有边框。"""
    tc = cell._tc
    tcPr = tc.find(qn('a:tcPr'))
    if tcPr is None:
        tcPr = parse_xml(
            '<a:tcPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
        tc.append(tcPr)
    for side in ['lnL', 'lnR', 'lnT', 'lnB']:
        for old in tcPr.findall(qn(f'a:{side}')):
            tcPr.remove(old)
        no_ln = parse_xml(
            f'<a:{side} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="0">'
            f'<a:noFill/></a:{side}>')
        tcPr.append(no_ln)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 高级组件
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _set_font(run, size, color, bold=False, italic=False, font=None):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font or FONT_BODY


def add_slide(decor=True):
    """新建幻灯片：渐变背景 + 可选装饰光晕。"""
    slide_counter[0] += 1
    s = prs.slides.add_slide(BLANK)
    _set_gradient_bg(s)
    if decor:
        glow_circle(s, Inches(10.5), Inches(-1.5), Inches(4), BLUE_HEX, alpha=8000)
    return s


def glow_circle(slide, left, top, size, color_hex, alpha=10000):
    """添加径向渐变半透明装饰圆。"""
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    shape.line.fill.background()
    spPr = shape._element.spPr
    for ch in list(spPr):
        if ch.tag.endswith('}solidFill'):
            spPr.remove(ch)
    xml = f'''
    <a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="{color_hex}"><a:alpha val="{alpha}"/></a:srgbClr></a:gs>
        <a:gs pos="100000"><a:srgbClr val="{color_hex}"><a:alpha val="0"/></a:srgbClr></a:gs>
      </a:gsLst>
      <a:path path="circle">
        <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
      </a:path>
    </a:gradFill>'''
    spPr.insert(0, parse_xml(xml))
    return shape


def gradient_line(slide, left, top, width, height=Pt(3)):
    """三色渐变装饰线（蓝→紫→淡出）。"""
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    line.line.fill.background()
    spPr = line._element.spPr
    for ch in list(spPr):
        if ch.tag.endswith('}solidFill'):
            spPr.remove(ch)
    xml = f'''
    <a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:gsLst>
        <a:gs pos="0"><a:srgbClr val="{BLUE_HEX}"/></a:gs>
        <a:gs pos="50000"><a:srgbClr val="{PURPLE_HEX}"/></a:gs>
        <a:gs pos="100000"><a:srgbClr val="{PURPLE_HEX}"><a:alpha val="0"/></a:srgbClr></a:gs>
      </a:gsLst>
      <a:lin ang="0" scaled="1"/>
    </a:gradFill>'''
    spPr.insert(0, parse_xml(xml))
    return line


def textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    return tb, tf


def add_title(slide, text, page_num=None):
    """页面标题：大字 + 渐变下划线 + 可选页码。"""
    tb, tf = textbox(slide, Inches(0.7), Inches(0.3), Inches(11.0), Inches(0.9))
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    _set_font(r, 30, FG, bold=True, font=FONT_TITLE)
    gradient_line(slide, Inches(0.7), Inches(1.15), Inches(3.5), Pt(3))
    if page_num is None:
        page_num = slide_counter[0]
    total = 27
    tb2, tf2 = textbox(slide, Inches(11.5), Inches(0.35), Inches(1.5), Inches(0.5))
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run()
    r2.text = f"{page_num:02d} / {total:02d}"
    _set_font(r2, 10, FG3, font=FONT_MONO)
    return tb


def card(slide, left, top, width, height, title, title_color, bullets,
         accent_hex=None, accent_side="top"):
    """圆角卡片 + 微渐变 + 阴影 + 可选高亮条。"""
    # 主卡片
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    _set_rounded_corners(box, 7000)
    _set_shape_gradient(box, CARD_TOP, CARD_BOT)
    box.line.color.rgb = RGBColor(0x30, 0x36, 0x3D)
    box.line.width = Pt(0.75)
    _add_shadow(box)

    # 高亮条
    if accent_hex:
        if accent_side == "top":
            bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                left + Inches(0.15), top, width - Inches(0.3), Pt(3))
            _set_rounded_corners(bar, 50000)
        else:  # left
            bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                left, top + Inches(0.12), Pt(4), height - Inches(0.24))
            _set_rounded_corners(bar, 50000)
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(
            int(accent_hex[0:2], 16), int(accent_hex[2:4], 16), int(accent_hex[4:6], 16))
        bar.line.fill.background()

    # 文本
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.22)
    tf.margin_right = Inches(0.22)
    tf.margin_top = Inches(0.18)

    p0 = tf.paragraphs[0]
    p0.space_after = Pt(8)
    r = p0.add_run()
    r.text = title
    _set_font(r, 15, title_color, bold=True, font=FONT_TITLE)

    for b in bullets:
        p = tf.add_paragraph()
        p.space_after = Pt(5)
        p.space_before = Pt(1)
        r = p.add_run()
        r.text = "  •  " + b
        _set_font(r, 13, FG)
    return box


def kpi_card(slide, left, top, width, height, number, label,
             color=GREEN, accent_hex=GREEN_HEX):
    """KPI 数字卡片：大号数字 + 标签 + 左侧高亮条。"""
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    _set_rounded_corners(box, 8000)
    _set_shape_gradient(box, CARD_TOP, CARD_BOT)
    box.line.color.rgb = RGBColor(0x30, 0x36, 0x3D)
    box.line.width = Pt(0.75)
    _add_shadow(box)

    # 左侧高亮条
    bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left + Pt(2), top + Inches(0.15), Pt(4), height - Inches(0.3))
    _set_rounded_corners(bar, 50000)
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor(
        int(accent_hex[0:2], 16), int(accent_hex[2:4], 16), int(accent_hex[4:6], 16))
    bar.line.fill.background()

    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_top = Inches(0.12)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    r = p0.add_run()
    r.text = number
    _set_font(r, 36, color, bold=True, font=FONT_NUM)

    p1 = tf.add_paragraph()
    p1.alignment = PP_ALIGN.CENTER
    p1.space_before = Pt(4)
    r1 = p1.add_run()
    r1.text = label
    _set_font(r1, 12, FG2)
    return box


def add_table(slide, left, top, width, headers, rows, col_widths=None,
              header_size=12, cell_size=11, row_h=Inches(0.44)):
    """美化表格：无边框 + 表头渐变 + 交替行色。"""
    nrows = len(rows) + 1
    ncols = len(headers)
    height = row_h * nrows
    tbl_shape = slide.shapes.add_table(nrows, ncols, left, top, width, height)
    gtbl = tbl_shape.table

    if col_widths:
        for i, w in enumerate(col_widths):
            gtbl.columns[i].width = w

    # 去除表格默认样式
    tbl_el = gtbl._tbl
    tbl_pr = tbl_el.find(qn('a:tblPr'))
    if tbl_pr is not None:
        tbl_pr.attrib.pop('bandRow', None)
        tbl_pr.attrib.pop('firstRow', None)
        tbl_pr.attrib.pop('lastRow', None)

    for c, h in enumerate(headers):
        cell = gtbl.cell(0, c)
        # 表头渐变
        tc = cell._tc
        tcPr = tc.find(qn('a:tcPr'))
        if tcPr is None:
            tcPr = parse_xml(
                '<a:tcPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
            tc.append(tcPr)
        # Add gradient to header
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0x1A, 0x23, 0x32)
        _remove_table_borders(cell)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = cell.text_frame.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = h
        _set_font(run, header_size, FG2, bold=True, font=FONT_TITLE)

    for r_i, row in enumerate(rows, start=1):
        for c_i, val in enumerate(row):
            if isinstance(val, tuple):
                text, color = val
            else:
                text, color = val, FG
            cell = gtbl.cell(r_i, c_i)
            # 交替行色
            if r_i % 2 == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0x0E, 0x12, 0x19)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0x12, 0x18, 0x20)
            _remove_table_borders(cell)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.CENTER if c_i > 0 else PP_ALIGN.LEFT
            cell.text_frame.margin_left = Inches(0.12)
            run = para.add_run()
            run.text = text
            _set_font(run, cell_size, color)
    return gtbl


def section_divider(num, title, subtitle=None):
    """分节页：大号渐变数字 + 装饰光晕 + 标题。"""
    s = add_slide(decor=False)
    # 装饰光晕
    glow_circle(s, Inches(8.5), Inches(1.0), Inches(5), BLUE_HEX, alpha=6000)
    glow_circle(s, Inches(-1), Inches(3.5), Inches(4), PURPLE_HEX, alpha=5000)

    # 大号数字
    tb, tf = textbox(s, Inches(0), Inches(1.8), EMU_W, Inches(2.0), MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = num
    _set_font(r, 96, BLUE, bold=True, font=FONT_NUM)
    # 给数字加渐变色
    _make_text_gradient(r._r, BLUE_HEX, PURPLE_HEX)

    # 标题
    tb2, tf2 = textbox(s, Inches(0), Inches(4.0), EMU_W, Inches(1.0), MSO_ANCHOR.MIDDLE)
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = title
    _set_font(r2, 34, FG, bold=True, font=FONT_TITLE)

    if subtitle:
        tb3, tf3 = textbox(s, Inches(0), Inches(5.0), EMU_W, Inches(0.7), MSO_ANCHOR.MIDDLE)
        p3 = tf3.paragraphs[0]
        p3.alignment = PP_ALIGN.CENTER
        r3 = p3.add_run()
        r3.text = subtitle
        _set_font(r3, 18, FG2, font=FONT_LIGHT)
    return s


def add_paragraphs(tf, items, first=False):
    """items: [(text, color, bold, size, font_override), ...]"""
    for i, it in enumerate(items):
        text = it[0]
        color = it[1] if len(it) > 1 else FG
        bold = it[2] if len(it) > 2 else False
        size = it[3] if len(it) > 3 and it[3] else 16
        font = it[4] if len(it) > 4 else None
        p = tf.paragraphs[0] if (first and i == 0) else tf.add_paragraph()
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = text
        _set_font(r, size, color, bold=bold, font=font)


# ═══════════════════════════════════════════════════════════
# SLIDE 1 · 封面
# ═══════════════════════════════════════════════════════════
s = add_slide(decor=False)
# 装饰光晕
glow_circle(s, Inches(-2), Inches(-1), Inches(7), BLUE_HEX, alpha=7000)
glow_circle(s, Inches(9), Inches(4), Inches(6), PURPLE_HEX, alpha=6000)
glow_circle(s, Inches(4), Inches(6), Inches(3), CYAN_HEX, alpha=4000)

# 主标题
tb, tf = textbox(s, Inches(0), Inches(2.0), EMU_W, Inches(1.3), MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "BTC Momentum"
_set_font(r, 54, FG, bold=True, font=FONT_TITLE)

# 渐变横线
gradient_line(s, Inches(4), Inches(3.35), Inches(5.333), Pt(3))

# 副标题
tb, tf = textbox(s, Inches(0), Inches(3.7), EMU_W, Inches(0.7), MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "JLST 反转-动量多信号策略系统"
_set_font(r, 22, FG2, font=FONT_LIGHT)

# 出处信息
tb, tf = textbox(s, Inches(0), Inches(4.7), EMU_W, Inches(1.6), MSO_ANCHOR.MIDDLE)
for i, (line, sz, it) in enumerate([
    ("基于 Jegadeesh, Luo, Subrahmanyam & Titman (2025)", 14, False),
    ("Review of Financial Studies — 全球顶级金融学术期刊", 13, True),
    ("回测数据截止 2026-09-07", 13, False),
]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(6)
    r = p.add_run()
    r.text = line
    _set_font(r, sz, FG3, italic=it, font=FONT_LIGHT)


# ═══════════════════════════════════════════════════════════
# SLIDE 2 · 你可能遇到过这些问题
# ═══════════════════════════════════════════════════════════
s = add_slide()
add_title(s, "你可能遇到过这些问题")

problems = [
    "BTC 跌了 10%，朋友说\"抄底\"，你不确定——还会反弹吗？",
    "BTC 涨了一个月，你想追——趋势还能持续吗？",
    "看了一堆指标（RSI、MACD、布林带），信号互相矛盾",
    "想要一个有学术论文支撑、而不是靠\"画线\"的分析工具",
    "找到了好信号，但不知道如何叠加多重确认来提高胜率",
]
for idx, prob in enumerate(problems):
    y = Inches(1.55) + Inches(0.68) * idx
    # 小圆点
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL,
                             Inches(0.85), y + Inches(0.08), Pt(8), Pt(8))
    dot.fill.solid()
    dot.fill.fore_color.rgb = BLUE
    dot.line.fill.background()
    # 文字
    tb, tf = textbox(s, Inches(1.15), y, Inches(11.0), Inches(0.55))
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = prob
    _set_font(r, 17, FG, font=FONT_BODY)

card(s, Inches(0.7), Inches(5.35), Inches(11.9), Inches(1.5),
     "BTC Momentum 的价值", BLUE,
     ["JLST 理论信号 + MA 均线多重过滤 → 用学术框架确定方向，用趋势滤波提高精度"],
     accent_hex=BLUE_HEX)


# ═══════════════════════════════════════════════════════════
# SLIDE 3 · 论文出处
# ═══════════════════════════════════════════════════════════
s = add_slide()
add_title(s, "论文出处")

card(s, Inches(0.7), Inches(1.45), Inches(11.9), Inches(1.45),
     "N. Jegadeesh, J. Luo, A. Subrahmanyam, S. Titman (2025)", FG,
     ['"Short-Term Reversals and Longer-Term Momentum around the World"',
      "Review of Financial Studies (RFS) — 金融学三大顶刊之一"],
     accent_hex=BLUE_HEX)

card(s, Inches(0.7), Inches(3.15), Inches(5.65), Inches(3.4),
     "论文做了什么？", BLUE,
     ["研究了 39 个国家股票市场",
      "用同一模型解释\"短期反转\"与\"中长期动量\"如何同时存在",
      "发现噪声交易是两者联系的关键纽带"],
     accent_hex=BLUE_HEX)

card(s, Inches(6.75), Inches(3.15), Inches(5.85), Inches(3.4),
     "我们做了什么？", GREEN,
     ["将论文理论应用到 BTC 市场",
      "构建实时仪表盘，每日更新",
      "叠加 MA 多信号过滤，将理论收益变为实操策略",
      "开放可配置滤波条件，用户自行调优"],
     accent_hex=GREEN_HEX)


# ═══════════════════════════════════════════════════════════
# SECTION 01 · 核心概念
# ═══════════════════════════════════════════════════════════
section_divider("01", "核心概念", "短期反转 & 中长期动量")

# -- 反转 vs 动量
s = add_slide()
add_title(s, "短期反转 vs 中长期动量")

card(s, Inches(0.7), Inches(1.5), Inches(5.65), Inches(2.5),
     "短期反转", ORANGE,
     ["\"涨多了会回、跌多了会弹\"",
      "短期价格含大量情绪化过度反应，几天内被纠正",
      "策略：短期逆着最近几天的极端走"],
     accent_hex=ORANGE_HEX)

card(s, Inches(6.75), Inches(1.5), Inches(5.85), Inches(2.5),
     "中长期动量", GREEN,
     ["\"强者恒强、弱者恒弱\"",
      "过去几个月上涨的往往继续涨，趋势有惯性",
      "策略：中长期顺着大趋势走"],
     accent_hex=GREEN_HEX)

card(s, Inches(0.7), Inches(4.2), Inches(11.9), Inches(2.6),
     "如何共存？——噪声交易者是关键", PURPLE,
     ["两者作用在不同时间尺度上，并不冲突",
      "噪声交易者制造短期过度波动（反转来源），其行为在中期形成趋势（动量来源）",
      "噪声越大 → 反转机会越明显；噪声越小、趋势越干净 → 动量越可靠",
      "模型用噪声大小动态调整反转/动量权重 —— 这是它比固定参数指标聪明的地方"],
     accent_hex=PURPLE_HEX)

# -- 综合评分公式
s = add_slide()
add_title(s, "综合评分公式")

# 公式卡片
formula_box = s.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(1.2), Inches(2.0), Inches(10.9), Inches(1.5))
_set_rounded_corners(formula_box, 8000)
_set_shape_gradient(formula_box, "1A2332", "1E2940")
formula_box.line.color.rgb = RGBColor(0x30, 0x3E, 0x55)
formula_box.line.width = Pt(1)
_add_shadow(formula_box, blur=63500, dist=38100, alpha=40000)

tf = formula_box.text_frame
tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "Composite = w₁·z(Rev) + w₂·z(Mom) + w₃·z(FundingRate)"
_set_font(r, 24, BLUE, bold=True, font=FONT_NUM)

tb, tf = textbox(s, Inches(0.8), Inches(3.9), Inches(11.7), Inches(2.8))
add_paragraphs(tf, [
    ("三个分量先做标准化（z-score），再按动态权重加权求和。", FG, False, 16),
    ("权重不是固定的，而是根据当前市场噪声大小实时计算。", FG, False, 16),
    ("BTC 适配版额外引入资金费率、波动率分档、减半周期等加密市场特有因子。", FG2, False, 14),
], first=True)


# ═══════════════════════════════════════════════════════════
# SECTION 02 · 信号系统
# ═══════════════════════════════════════════════════════════
section_divider("02", "信号系统", "从评分到买卖信号")

s = add_slide()
add_title(s, "信号产生规则")

card(s, Inches(0.7), Inches(1.5), Inches(11.9), Inches(2.2),
     "评分如何变成信号", FG,
     ["综合评分向上突破阈值 → 做多信号（绿色↑）",
      "综合评分向下跌破阈值 → 做空信号（红色↓）",
      "多个条件同时触发 → 级联预警（紫色⚡，强信号）"],
     accent_hex=BLUE_HEX)

card(s, Inches(0.7), Inches(3.95), Inches(11.9), Inches(2.7),
     "JLST 原始信号前瞻回报（无 MA 滤波）", GREEN,
     ["做多信号 30 天平均 +10.71%，胜率 58%",
      "问题：部分做多信号出现在下跌趋势里（逆势抄底），容易吃套",
      "→ 这正是叠加均线滤波的动机"],
     accent_hex=GREEN_HEX)


# ═══════════════════════════════════════════════════════════
# SECTION 03 · 多信号策略
# ═══════════════════════════════════════════════════════════
section_divider("03", "多信号策略", "JLST + MA 滤波")

s = add_slide()
add_title(s, "多信号滤波逻辑")

card(s, Inches(0.7), Inches(1.5), Inches(5.65), Inches(3.2),
     "做多滤波（三条件全满足）", GREEN,
     ["① 价格 > MA110（大趋势向上）",
      "② MA6 > MA103（短期走强）",
      "③ 价格 > EMA50（中期向上）",
      "只有 JLST 做多信号 + 三条件全满足 = 高质量信号"],
     accent_hex=GREEN_HEX)

card(s, Inches(6.75), Inches(1.5), Inches(5.85), Inches(3.2),
     "做空滤波（价格在均线下方）", RED,
     ["价格 < MA110",
      "价格 < EMA50",
      "要求价格在均线下方才保留做空信号"],
     accent_hex=RED_HEX)

card(s, Inches(0.7), Inches(4.9), Inches(11.9), Inches(1.8),
     "为什么 MA 滤波有效 + 用户可配置", BLUE,
     ["均线是趋势的过滤网：筛掉逆势信号（亏损主要来源），只留顺势信号",
      "三个开关用户自主控制：全勾最严格、可只留一两个、可全取消看原始信号，实时重新统计前瞻收益"],
     accent_hex=BLUE_HEX)


# ═══════════════════════════════════════════════════════════
# SECTION 04 · 超额收益
# ═══════════════════════════════════════════════════════════
section_divider("04", "超额收益",
                "回测截止 2026-09-07 · 信号 200 条（做多 85 / 做空 109）")

# -- 对比表
s = add_slide()
add_title(s, "多信号 vs 原始 JLST（做多）")

add_table(s, Inches(0.9), Inches(1.55), Inches(11.5),
          ["持有期", "JLST 原始", "多信号滤波", "超额提升"],
          [
              ["1 天", "+1.12%", ("+1.74%", GREEN), ("+55%", GREEN)],
              ["3 天", "+2.01%", ("+3.50%", GREEN), ("+74%", GREEN)],
              ["7 天", "+3.25%", ("+6.03%", GREEN), ("+86%", GREEN)],
              ["30 天", "+10.71%", ("+20.52%", GREEN), ("+92%", GREEN)],
          ],
          col_widths=[Inches(2.5), Inches(3.0), Inches(3.0), Inches(3.0)])

card(s, Inches(0.9), Inches(5.05), Inches(11.5), Inches(1.7),
     "核心结论", GREEN,
     ["做多三重滤波后 30 天平均收益 +10.71% → +20.52%（超额 +9.82 个百分点，相对提升 92%）",
      "胜率 58% → 62%。过滤逆势信号是获取超额回报的关键"],
     accent_hex=GREEN_HEX)

# -- 做多最强阿尔法
s = add_slide()
add_title(s, "做多策略 — 最强阿尔法")

kpi_card(s, Inches(0.7), Inches(1.5), Inches(3.6), Inches(1.7),
         "+20.52%", "多信号做多 30 日均收益",
         GREEN, GREEN_HEX)
kpi_card(s, Inches(4.65), Inches(1.5), Inches(3.6), Inches(1.7),
         "67%", "3 天胜率",
         GREEN, GREEN_HEX)
kpi_card(s, Inches(8.6), Inches(1.5), Inches(3.95), Inches(1.7),
         "85 → 42", "信号次数（少而精）",
         BLUE, BLUE_HEX)

card(s, Inches(0.7), Inches(3.5), Inches(11.9), Inches(3.2),
     "解读", FG,
     ["三重滤波把做多信号从 85 次精简到 42 次 —— 宁可少做，也要做对",
      "各持有期收益全面提升：1天 +55%、3天 +74%、7天 +86%、30天 +92%",
      "胜率区间从 58-65% 抬升到 62-67%",
      "这是整套策略里最稳定、最值得依赖的部分"],
     accent_hex=GREEN_HEX)

# -- 做空
s = add_slide()
add_title(s, "做空策略 — 诚实面对局限")

add_table(s, Inches(0.9), Inches(1.55), Inches(11.5),
          ["持有期", "JLST 做空", "多信号做空", "变化"],
          [
              ["1 天", "+0.44% (53%)", ("-0.26% (44%)", FG2), "走弱"],
              ["3 天", "+0.29% (50%)", ("-1.04% (44%)", FG2), "走弱"],
              ["7 天", "+0.76% (52%)", ("-1.19% (44%)", FG2), "走弱"],
              ["30 天", ("-3.61% (42%)", RED), ("-2.87% (38%)", RED), "略改善"],
          ],
          col_widths=[Inches(2.3), Inches(3.2), Inches(3.2), Inches(2.8)])

card(s, Inches(0.9), Inches(5.05), Inches(11.5), Inches(1.7),
     "解读", ORANGE,
     ["BTC 长期上行，做空天然逆风：做空 30 天平均 -3.61%、胜率仅 42%",
      "均线下方常是超跌区，加滤波后短期反而更弱、易遇反弹",
      "做空的定位是风险预警与对冲，不宜作为主进攻策略"],
     accent_hex=ORANGE_HEX)

# -- 核心数据总结
s = add_slide()
add_title(s, "核心数据总结")

kpi_card(s, Inches(0.7), Inches(1.5), Inches(3.6), Inches(1.5),
         "+20.52%", "做多 30 日均收益",
         GREEN, GREEN_HEX)
kpi_card(s, Inches(4.65), Inches(1.5), Inches(3.6), Inches(1.5),
         "+92%", "30 日超额提升",
         BLUE, BLUE_HEX)
kpi_card(s, Inches(8.6), Inches(1.5), Inches(3.95), Inches(1.5),
         "67%", "3 日胜率（做多）",
         ORANGE, ORANGE_HEX)

card(s, Inches(0.7), Inches(3.3), Inches(5.65), Inches(3.3),
     "做多策略结论", GREEN,
     ["MA 滤波后 30 日收益近乎翻倍",
      "胜率 58-65% → 62-67%",
      "信号数 85 → 42（少而精）",
      "每次信号平均 30 日赚 20.52%"],
     accent_hex=GREEN_HEX)

card(s, Inches(6.75), Inches(3.3), Inches(5.85), Inches(3.3),
     "做空策略结论", RED,
     ["做空收益整体偏弱（BTC 长期上行）",
      "无滤波 30 日 -3.61%，滤波改善有限",
      "信号数 109 → 45（过滤假信号）",
      "更适合防御/对冲，不做主策略"],
     accent_hex=RED_HEX)


# ═══════════════════════════════════════════════════════════
# SECTION 05 · 仪表盘功能
# ═══════════════════════════════════════════════════════════
section_divider("05", "仪表盘功能", "看得懂、用得上")

# -- K 线图
s = add_slide()
add_title(s, "交互式 K 线图")

card(s, Inches(0.7), Inches(1.5), Inches(5.65), Inches(2.4),
     "价格同轴（趋势）", BLUE,
     ["MA6 · EMA50 · EMA110",
      "MA103 · MA110 · MA200",
      "已实现价格"],
     accent_hex=BLUE_HEX)

card(s, Inches(6.75), Inches(1.5), Inches(5.85), Inches(2.4),
     "独立轴（估值/链上/情绪）", PURPLE,
     ["成交量 · RSI · Mayer · MVRV · NUPL",
      "SMM · 卖方衰竭 · 风险回报",
      "ETF · USDT.D · BTC.D"],
     accent_hex=PURPLE_HEX)

card(s, Inches(0.7), Inches(4.1), Inches(11.9), Inches(2.6),
     "特色功能", GREEN,
     ["📌 信号标记：一键叠加 JLST 买卖信号（绿↑做多 / 红↓做空 / 紫⚡级联）",
      "对数坐标：长周期看 BTC 更合理",
      "坐标翻转：USDT.D / BTC.D 可翻转，直观看与价格的反向关系",
      "日线/周线切换：指标与信号标记自动适配周线聚合"],
     accent_hex=GREEN_HEX)

# -- 实操流程
s = add_slide()
add_title(s, "实操流程")

steps = [
    ("①", "看综合评分和信号方向", FG),
    ("②", "看三个 MA 滤波灯是否都绿", FG),
    ("③", "全绿 = 最高质量做多信号", GREEN),
    ("④", "结合估值指标（MVRV、NUPL、Mayer）做二次确认", FG),
    ("⑤", "自己决定仓位", FG),
]
for idx, (num, text, color) in enumerate(steps):
    y = Inches(1.55) + Inches(0.72) * idx
    # 数字圆
    circ = s.shapes.add_shape(MSO_SHAPE.OVAL,
                              Inches(0.85), y + Inches(0.03), Inches(0.38), Inches(0.38))
    _set_shape_gradient(circ, BLUE_HEX, PURPLE_HEX)
    circ.line.fill.background()
    _add_shadow(circ, blur=25400, dist=12700, alpha=25000)
    ctf = circ.text_frame
    ctf.vertical_anchor = MSO_ANCHOR.MIDDLE
    cp = ctf.paragraphs[0]
    cp.alignment = PP_ALIGN.CENTER
    cr = cp.add_run()
    cr.text = num
    _set_font(cr, 13, FG, bold=True, font=FONT_NUM)
    # 文字
    tb, tf = textbox(s, Inches(1.4), y, Inches(10.5), Inches(0.5))
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    _set_font(r, 18, color, bold=True, font=FONT_BODY)

card(s, Inches(0.7), Inches(5.2), Inches(11.9), Inches(1.6),
     "最佳实践", BLUE,
     ["做多 + 三 MA 全满足 = 最高质量信号（历史 30 日 +20.52%）",
      "做空在牛市中谨慎，作减仓/对冲依据；信号约 30 天一次，属低频高质量策略"],
     accent_hex=BLUE_HEX)


# ═══════════════════════════════════════════════════════════
# SECTION 06 · 指标撰写规则
# ═══════════════════════════════════════════════════════════
section_divider("06", "K 线指标撰写规则", "每个指标怎么算、怎么读")

# -- 趋势均线
s = add_slide()
add_title(s, "趋势均线：MA 与 EMA")

add_table(s, Inches(0.6), Inches(1.45), Inches(12.1),
          ["指标", "计算规则", "解读"],
          [
              ["MA6", "近 6 日收盘价简单平均 SMA(close,6)", "超短期趋势；与 MA103 交叉用作多滤波"],
              ["EMA50", "50 日指数移动平均，α=2/51，近期权重更高", "中期趋势线；站上=多头，跌破=空头"],
              ["EMA110", "110 日指数移动平均", "牛熊分界参考，比 SMA 反应更快"],
              ["MA103 / MA110", "103 / 110 日简单平均", "周期性中长期支撑/压力（BTC 半年级别）"],
              ["MA200", "200 日简单平均", "长期牛熊线；也是 Mayer 倍数的分母"],
          ],
          col_widths=[Inches(2.3), Inches(5.3), Inches(4.5)],
          cell_size=10.5, row_h=Inches(0.55))

card(s, Inches(0.6), Inches(5.6), Inches(12.1), Inches(1.15),
     "SMA vs EMA", BLUE,
     ["SMA 窗口内等权、平滑但滞后；EMA 用衰减系数 α 给近期更高权重、转向更快但更易受噪声影响。"
      "趋势策略里两者互补。"],
     accent_hex=BLUE_HEX)

# -- 估值类
s = add_slide()
add_title(s, "估值类：Mayer / 已实现价格 / MVRV")

add_table(s, Inches(0.6), Inches(1.45), Inches(12.1),
          ["指标", "计算规则", "解读"],
          [
              ["Mayer 倍数", "价格 ÷ MA200", ">2.4 顶部区；<1 低估区。衡量偏离长期均线程度"],
              ["已实现价格", "每枚 BTC「最后移动时价格」加权平均（全网成本）",
               "市场平均持仓成本线；跌破=多数人浮亏，历史大底常在此"],
              ["MVRV", "市值 ÷ 已实现市值", ">3.7 顶部风险；<1 深度低估。全网未实现盈亏比"],
          ],
          col_widths=[Inches(2.3), Inches(5.3), Inches(4.5)],
          cell_size=10.5, row_h=Inches(0.7))

card(s, Inches(0.6), Inches(5.15), Inches(12.1), Inches(1.3),
     "为什么重要", GREEN,
     ["估值类提供「贵不贵」的锚，与趋势类「涨不涨」互补——趋势告诉你方向，估值告诉你位置。"],
     accent_hex=GREEN_HEX)

# -- 链上情绪
s = add_slide()
add_title(s, "链上情绪：NUPL / SMM / 卖方衰竭 / 风险回报")

add_table(s, Inches(0.6), Inches(1.45), Inches(12.1),
          ["指标", "计算规则", "解读"],
          [
              ["NUPL 净未实现盈亏", "(市值 − 已实现市值) ÷ 市值",
               ">0.75 极度贪婪；<0 全网浮亏（投降/大底）"],
              ["SMM", "链上花费产出的动量", "反映抛压动能变化，捕捉筹码换手节奏"],
              ["卖方衰竭", "波动率 × 长期持有者未实现亏损占比", "走高=抛售动能耗尽，常见于底部"],
              ["风险回报", "历史价格分布的下行风险 vs 上行空间比值", "越低=当前介入性价比越高"],
          ],
          col_widths=[Inches(2.6), Inches(5.0), Inches(4.5)],
          cell_size=10, row_h=Inches(0.62))

tb, tf = textbox(s, Inches(0.6), Inches(6.15), Inches(12.1), Inches(0.6))
p = tf.paragraphs[0]
r = p.add_run()
r.text = "链上数据来自 CryptoQuant，按日期对齐到 K 线；周线视图取每周最后一个值。"
_set_font(r, 11, FG3, font=FONT_LIGHT)

# -- 资金流与情绪
s = add_slide()
add_title(s, "资金流与情绪：ETF / USDT.D / BTC.D / 资金费率")

add_table(s, Inches(0.6), Inches(1.45), Inches(12.1),
          ["指标", "计算规则", "解读"],
          [
              ["ETF 净流入", "现货 BTC ETF 每日净申赎（百万美元），周线求和",
               "正=机构增持进场；持续净流出=需求转弱"],
              ["USDT.D", "USDT 市值 ÷ 加密总市值（%）",
               "升=资金避险（利空）；降=资金进场（利多）。与 BTC 常反向"],
              ["BTC.D", "BTC 市值 ÷ 加密总市值（%）", "升=资金集中 BTC；降=流向山寨"],
              ["资金费率", "永续合约多空平衡费率，每 8h 结算，取日均",
               "正=多头拥挤；极端正值常预示回调"],
          ],
          col_widths=[Inches(2.3), Inches(5.3), Inches(4.5)],
          cell_size=10, row_h=Inches(0.62))

card(s, Inches(0.6), Inches(5.35), Inches(12.1), Inches(1.35),
     "资金费率的特殊地位", ORANGE,
     ["它是 JLST 动态权重中噪声代理的组成部分（占 30% 权重）——"
      "费率越极端，模型越调低动量权重、调高反转权重。"],
     accent_hex=ORANGE_HEX)

# -- 适用场景
s = add_slide()
add_title(s, "什么情况下最适合用多策略？")

card(s, Inches(0.5), Inches(1.5), Inches(3.9), Inches(4.0),
     "✅ 最适合", GREEN,
     ["明确趋势市：价格站上 MA110/EMA50，做多胜率 62-67%",
      "中长线持有（1-4 周）：30 日收益差距最大",
      "右侧交易者：愿放弃最低点换确定性"],
     accent_hex=GREEN_HEX)

card(s, Inches(4.65), Inches(1.5), Inches(3.9), Inches(4.0),
     "⚠️ 谨慎", ORANGE,
     ["震荡/横盘市：均线频繁穿越，滤波反复触发又失效",
      "做空：BTC 长期上行，整体为负，仅作预警",
      "短持有期（1-3 日）：超额小，成本占比高"],
     accent_hex=ORANGE_HEX)

card(s, Inches(8.8), Inches(1.5), Inches(3.9), Inches(4.0),
     "❌ 不适合", RED,
     ["日内/高频：信号约 30 天一次，频率不匹配",
      "黑天鹅急跌：任何趋势滤波都滞后",
      "当自动交易机器人：它是决策辅助，非全自动"],
     accent_hex=RED_HEX)

card(s, Inches(0.5), Inches(5.7), Inches(12.2), Inches(1.15),
     "一句话", BLUE,
     ["多策略的超额收益来自「在趋势里做多、并用均线把逆势信号筛掉」。"
      "趋势越明确、持有期越长，价值越大；越震荡、越短线，价值越小。"],
     accent_hex=BLUE_HEX)


# ═══════════════════════════════════════════════════════════
# SECTION 07 · 局限与总结
# ═══════════════════════════════════════════════════════════
section_divider("07", "局限与总结", "诚实面对模型的边界")

# -- 注意事项
s = add_slide()
add_title(s, "需要注意的事项")

caveats = [
    "历史回测 ≠ 未来表现 —— 市场结构会变化",
    "论文参数来自股票市场 —— 虽已做 BTC 适配，仍是近似而非精确",
    "低频策略 —— 约 30 天触发一次，不适合日内或高频",
    "做空信号效果一般 —— BTC 长期向上，做空天然逆风、胜率偏低",
    "MA 参数有优化空间 —— 当前 MA110/MA103/EMA50 基于经验，用户可自调",
    "极端行情下信号会滞后 —— 黑天鹅事件中任何趋势指标都滞后",
]
for idx, text in enumerate(caveats):
    y = Inches(1.5) + Inches(0.62) * idx
    # 小三角警示图标
    tri = s.shapes.add_shape(MSO_SHAPE.OVAL,
                             Inches(0.85), y + Inches(0.06), Pt(8), Pt(8))
    tri.fill.solid()
    tri.fill.fore_color.rgb = ORANGE
    tri.line.fill.background()
    tb, tf = textbox(s, Inches(1.15), y, Inches(11.2), Inches(0.5))
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    _set_font(r, 15, FG)

card(s, Inches(0.7), Inches(5.4), Inches(11.9), Inches(1.35),
     "正确定位", ORANGE,
     ["BTC Momentum 是辅助决策工具，帮你过滤噪声、提供学术框架支撑的方向判断，"
      "但最终决策权在你——不是自动赚钱机器。"],
     accent_hex=ORANGE_HEX)

# -- 总结
s = add_slide(decor=False)
add_title(s, "总结")
# 装饰
glow_circle(s, Inches(9), Inches(4.5), Inches(5), BLUE_HEX, alpha=5000)
glow_circle(s, Inches(-1.5), Inches(5), Inches(4), PURPLE_HEX, alpha=4000)

card(s, Inches(0.7), Inches(1.5), Inches(5.65), Inches(2.5),
     "🔬  理论基础", BLUE,
     ["JLST (2025) RFS 论文 — 39 国验证",
      "短期反转 + 中长期动量双因子",
      "噪声交易者是关键中介变量"],
     accent_hex=BLUE_HEX)

card(s, Inches(6.75), Inches(1.5), Inches(5.85), Inches(2.5),
     "🎯  多信号策略", GREEN,
     ["JLST 信号 + MA110/MA6>MA103/EMA50",
      "做多 30 日 +20.52%（胜率 62%）",
      "超额比原始信号提升 92%"],
     accent_hex=GREEN_HEX)

card(s, Inches(0.7), Inches(4.2), Inches(11.9), Inches(2.5),
     "🖥️  工具特色", PURPLE,
     ["每日自动更新数据 · 近 20 个技术/链上指标 + 信号标记",
      "5 个可配置 MA 滤波条件 · 日线/周线 · 暗/亮主题 · 三图时间轴联动",
      "把严谨的学术框架，变成你每天都能打开、看得懂、用得上的东西"],
     accent_hex=PURPLE_HEX)

# ── 保存 ──
out = Path(__file__).resolve().parent.parent / "BTC_Momentum_演示.pptx"
prs.save(str(out))
print(f"OK  已生成 {out}  （共 {slide_counter[0]} 页）")
