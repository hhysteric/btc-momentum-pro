#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 BTC Momentum 演示 PPT。

设计语言：单一深色背景（微渐变，无跨元素装饰线）+ 圆角卡片（纯色填充、
弱阴影）+ 一个主导色（蓝）+ 语义色仅用于真实的正负数据（绿/红）+
全篇 Microsoft YaHei（通过 bold 标记区分粗细，不依赖具体字重变体）+
不在标题下加装饰线、不在卡片上加色条（这两者是明显的 AI 生成痕迹）。

面向零基础第三方观众，把概念讲透；回测参数优化截止 2026-09-02，
信号数据更新至 2026-09-10。所有数字均经真实数据核实。

用法：python scripts/build_pptx.py
输出：BTC_Momentum_演示.pptx（项目根目录）
"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 配色 —— 蓝色为主导色（约 65% 视觉权重），绿/红仅用于真实的正负数据，
# 不再使用橙/紫/青作为装饰性强调色。
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
BG_TOP   = "0B0F16"
BG_BOT   = "121826"
CARD_BG  = "1A2130"
CARD_LN  = "323B4E"

FG   = RGBColor(0xF0, 0xF3, 0xF8)   # 主文字：接近纯白，最高对比度
FG2  = RGBColor(0xC7, 0xCF, 0xDC)   # 次要文字：仍保持强对比度（不用暗灰）
MUTE = RGBColor(0x8B, 0x96, 0xAC)   # 仅用于说明性小字（≥11pt 才用）

GREEN  = RGBColor(0x4C, 0xC9, 0x64)   # 正值 / 利好
RED    = RGBColor(0xF0, 0x66, 0x5E)   # 负值 / 风险
BLUE   = RGBColor(0x6E, 0xB4, 0xFF)   # 主导色
BLUE_DARK_HEX = "1A2332"

GREEN_HEX = "4CC964"
RED_HEX   = "F0665E"
BLUE_HEX  = "6EB4FF"

# 字体 —— 全部使用系统内置常规字体族，粗细用 bold 标记区分，
# 不指定具体字重变体名称（避免用户系统缺少该变体导致回退渲染）。
FONT_CJK = "Microsoft YaHei"
FONT_NUM = "Segoe UI"

EMU_W = Inches(13.333)
EMU_H = Inches(7.5)

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]

slide_counter = [0]
TOTAL_SLIDES = 30


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 底层 XML 工具
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _set_gradient_bg(slide, top=BG_TOP, bot=BG_BOT):
    """幻灯片背景：两段微渐变（纯粹用于避免大片纯色的单调感，非装饰线/条）。

    注意：<p:bg> 是 <p:cSld> 的子元素（且必须是其第一个子元素，位于
    <p:spTree> 之前），不是 <p:sld> 的子元素。之前误插到了 <p:sld> 顶层，
    导致背景渐变被静默忽略、整页回退成默认白色背景——这也是文字大面积
    "看不清"的根本原因（白底配了为深色背景设计的浅色文字）。
    """
    cSld = slide._element.find(qn('p:cSld'))
    for old in cSld.findall(qn('p:bg')):
        cSld.remove(old)
    bg_xml = f'''
    <p:bg xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:bgPr>
        <a:gradFill>
          <a:gsLst>
            <a:gs pos="0"><a:srgbClr val="{top}"/></a:gs>
            <a:gs pos="100000"><a:srgbClr val="{bot}"/></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="1"/>
        </a:gradFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>'''
    bg_el = parse_xml(bg_xml)
    cSld.insert(0, bg_el)


def _add_shadow(shape, blur=40000, dist=20000, alpha=30000):
    """弱外阴影，仅提供轻微景深，不影响文字对比度（阴影只在卡片外沿）。"""
    spPr = shape._element.spPr
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


def _remove_table_borders(cell):
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


def _hex_to_rgb(hex_str):
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 高级组件
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _set_font(run, size, color, bold=False, italic=False, font=None):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = font or FONT_CJK
    # 显式设置东亚字体，避免仅设置西文字体名导致 CJK 字形回退不一致。
    rPr = run._r.find(qn('a:rPr'))
    if rPr is not None:
        ea = rPr.find(qn('a:ea'))
        if ea is None:
            ea = parse_xml(
                '<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
            rPr.append(ea)
        ea.set('typeface', font or FONT_CJK)


def glow_circle(slide, left, top, size, color_hex, alpha=6000):
    """极弱的径向光晕，仅作背景角落装饰，不与文字重叠，不影响可读性。"""
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


def add_slide(decor=None):
    """新建幻灯片。decor: 可选 (left, top, size, color_hex) 加一个角落光晕。"""
    slide_counter[0] += 1
    s = prs.slides.add_slide(BLANK)
    _set_gradient_bg(s)
    if decor:
        glow_circle(s, *decor)
    return s


def textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    return tb, tf


def add_title(slide, text, page_num=None):
    """页面标题：大字加粗，仅靠留白与正文区分——不加下划线、不加色条。"""
    tb, tf = textbox(slide, Inches(0.7), Inches(0.35), Inches(10.3), Inches(0.85))
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    _set_font(r, 32, FG, bold=True)

    if page_num is None:
        page_num = slide_counter[0]
    tb2, tf2 = textbox(slide, Inches(11.5), Inches(0.42), Inches(1.5), Inches(0.5))
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    r2 = p2.add_run()
    r2.text = f"{page_num:02d} / {TOTAL_SLIDES:02d}"
    _set_font(r2, 11, MUTE, font=FONT_NUM)
    return tb


def card(slide, left, top, width, height, title, title_color, bullets,
         body_size=15):
    """圆角卡片：纯色深底 + 细边框 + 弱阴影。不加顶部/侧边色条。"""
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    _set_rounded_corners(box, 6000)
    box.fill.solid()
    box.fill.fore_color.rgb = _hex_to_rgb(CARD_BG)
    box.line.color.rgb = _hex_to_rgb(CARD_LN)
    box.line.width = Pt(1)
    _add_shadow(box)

    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.2)
    tf.margin_bottom = Inches(0.15)

    p0 = tf.paragraphs[0]
    p0.space_after = Pt(10)
    r = p0.add_run()
    r.text = title
    _set_font(r, 17, title_color, bold=True)

    for b in bullets:
        p = tf.add_paragraph()
        p.space_after = Pt(7)
        p.space_before = Pt(1)
        r = p.add_run()
        r.text = "•  " + b
        _set_font(r, body_size, FG)
    return box


def kpi_card(slide, left, top, width, height, number, label, color=GREEN):
    """KPI 数字卡片：大号数字居中 + 标签。不加侧边色条。"""
    box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    _set_rounded_corners(box, 7000)
    box.fill.solid()
    box.fill.fore_color.rgb = _hex_to_rgb(CARD_BG)
    box.line.color.rgb = _hex_to_rgb(CARD_LN)
    box.line.width = Pt(1)
    _add_shadow(box)

    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    r = p0.add_run()
    r.text = number
    _set_font(r, 38, color, bold=True, font=FONT_NUM)

    p1 = tf.add_paragraph()
    p1.alignment = PP_ALIGN.CENTER
    p1.space_before = Pt(5)
    r1 = p1.add_run()
    r1.text = label
    _set_font(r1, 13, FG2)
    return box


def add_table(slide, left, top, width, headers, rows, col_widths=None,
              header_size=14, cell_size=13, row_h=Inches(0.48)):
    """表格：无边框细线，表头深底、数据行弱交替底色，全部保持高对比度文字。"""
    nrows = len(rows) + 1
    ncols = len(headers)
    height = row_h * nrows
    tbl_shape = slide.shapes.add_table(nrows, ncols, left, top, width, height)
    gtbl = tbl_shape.table

    if col_widths:
        for i, w in enumerate(col_widths):
            gtbl.columns[i].width = w

    tbl_el = gtbl._tbl
    tbl_pr = tbl_el.find(qn('a:tblPr'))
    if tbl_pr is not None:
        tbl_pr.attrib.pop('bandRow', None)
        tbl_pr.attrib.pop('firstRow', None)
        tbl_pr.attrib.pop('lastRow', None)

    for c, h in enumerate(headers):
        cell = gtbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = _hex_to_rgb(BLUE_DARK_HEX)
        _remove_table_borders(cell)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = cell.text_frame.paragraphs[0]
        para.alignment = PP_ALIGN.CENTER
        run = para.add_run()
        run.text = h
        _set_font(run, header_size, FG, bold=True)

    for r_i, row in enumerate(rows, start=1):
        for c_i, val in enumerate(row):
            if isinstance(val, tuple):
                text, color = val
            else:
                text, color = val, FG
            cell = gtbl.cell(r_i, c_i)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _hex_to_rgb(CARD_BG if r_i % 2 == 1 else "141B29")
            _remove_table_borders(cell)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.CENTER if c_i > 0 else PP_ALIGN.LEFT
            cell.text_frame.margin_left = Inches(0.15)
            run = para.add_run()
            run.text = text
            _set_font(run, cell_size, color)
    return gtbl


def section_divider(num, title, subtitle=None):
    """分节页：大号纯色数字 + 角落光晕 + 标题。"""
    s = add_slide(decor=(Inches(8.8), Inches(0.8), Inches(5), BLUE_HEX))
    glow_circle(s, Inches(-1.2), Inches(4.2), Inches(4), BLUE_HEX, alpha=4000)

    tb, tf = textbox(s, Inches(0), Inches(1.7), EMU_W, Inches(2.0), MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = num
    _set_font(r, 96, BLUE, bold=True, font=FONT_NUM)

    tb2, tf2 = textbox(s, Inches(0), Inches(3.9), EMU_W, Inches(1.0), MSO_ANCHOR.MIDDLE)
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run()
    r2.text = title
    _set_font(r2, 36, FG, bold=True)

    if subtitle:
        tb3, tf3 = textbox(s, Inches(0), Inches(4.95), EMU_W, Inches(0.7), MSO_ANCHOR.MIDDLE)
        p3 = tf3.paragraphs[0]
        p3.alignment = PP_ALIGN.CENTER
        r3 = p3.add_run()
        r3.text = subtitle
        _set_font(r3, 18, FG2)
    return s


def add_paragraphs(tf, items, first=False):
    """items: [(text, color, bold, size), ...]"""
    for i, it in enumerate(items):
        text = it[0]
        color = it[1] if len(it) > 1 else FG
        bold = it[2] if len(it) > 2 else False
        size = it[3] if len(it) > 3 and it[3] else 16
        p = tf.paragraphs[0] if (first and i == 0) else tf.add_paragraph()
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = text
        _set_font(r, size, color, bold=bold)


# ═══════════════════════════════════════════════════════════
# 开场
# ═══════════════════════════════════════════════════════════

# SLIDE 1 · 封面
s = add_slide(decor=(Inches(-2), Inches(-1.5), Inches(7), BLUE_HEX))
glow_circle(s, Inches(9.5), Inches(4.5), Inches(6), BLUE_HEX, alpha=4000)

tb, tf = textbox(s, Inches(0), Inches(2.05), EMU_W, Inches(1.3), MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "BTC Momentum"
_set_font(r, 54, FG, bold=True, font=FONT_NUM)

tb, tf = textbox(s, Inches(0), Inches(3.25), EMU_W, Inches(0.7), MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "把一篇顶刊论文，变成能用的 BTC 择时工具"
_set_font(r, 22, FG2)

tb, tf = textbox(s, Inches(0), Inches(4.6), EMU_W, Inches(1.7), MSO_ANCHOR.MIDDLE)
for i, (line, sz, it) in enumerate([
    ("基于 Jegadeesh, Luo, Subrahmanyam & Titman (2025)", 15, False),
    ("Review of Financial Studies — 全球顶级金融学术期刊", 14, True),
    ("回测参数优化截止 2026-09-02 · 信号数据更新至 2026-09-10", 13, False),
]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(7)
    r = p.add_run()
    r.text = line
    _set_font(r, sz, FG2, italic=it)


# SLIDE 2 · 你可能遇到过这些问题
s = add_slide()
add_title(s, "你可能遇到过这些问题")

problems = [
    "BTC 跌了 10%，朋友说\"抄底\"，你不确定——还会反弹吗？",
    "BTC 涨了一个月，你想追——趋势还能持续吗？",
    "看了一堆指标（RSI、MACD、布林带），信号互相矛盾，不知听谁的",
    "想要一个有学术论文支撑、而不是靠\"感觉画线\"的分析工具",
    "找到了信号，却不知道怎么叠加多重确认来提高胜率",
]
for idx, prob in enumerate(problems):
    y = Inches(1.5) + Inches(0.66) * idx
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL,
                             Inches(0.85), y + Inches(0.1), Pt(9), Pt(9))
    dot.fill.solid()
    dot.fill.fore_color.rgb = BLUE
    dot.line.fill.background()
    tb, tf = textbox(s, Inches(1.15), y, Inches(11.0), Inches(0.55))
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = prob
    _set_font(r, 17, FG)

card(s, Inches(0.7), Inches(5.35), Inches(11.9), Inches(1.5),
     "这套工具想解决的核心问题", BLUE,
     ["用一个有学术论文支撑、逻辑自洽的框架先判断方向，再用均线做多重确认把胜率提上去"])


# SLIDE 3 · 一句话是什么 + 路线图
s = add_slide()
add_title(s, "这套工具是什么 · 今天讲什么")

card(s, Inches(0.7), Inches(1.4), Inches(11.9), Inches(1.5),
     "一句话", BLUE,
     ["用一篇顶刊论文的模型判断方向（该做多还是做空），再用均线滤波把\"逆势\"的买卖点筛掉，只留顺势的高质量信号。"])

card(s, Inches(0.7), Inches(3.1), Inches(5.65), Inches(3.5),
     "接下来会依次讲清 6 件事", FG,
     ["① 两个核心直觉：短期反转 & 中长期动量",
      "② 论文做了什么、怎么搬到 BTC",
      "③ 信号和买卖点是怎么产生的",
      "④ 有多少超额收益、为什么 30 天最好",
      "⑤ 样本外验证与适用边界",
      "⑥ 局限与总结"], body_size=14)

card(s, Inches(6.75), Inches(3.1), Inches(5.85), Inches(3.5),
     "给完全不懂的你一个承诺", GREEN,
     ["不需要任何量化或编程背景",
      "每个术语都会用生活化的例子讲清楚",
      "每个数字都来自真实回测、可溯源",
      "该说的局限一个都不藏——包括它什么时候不灵"], body_size=14)


# ═══════════════════════════════════════════════════════════
# SECTION 01 · 两个核心直觉
# ═══════════════════════════════════════════════════════════
section_divider("01", "两个核心直觉", "短期反转 & 中长期动量")

# SLIDE · 什么是短期反转
s = add_slide()
add_title(s, "什么是\"短期反转\"")

card(s, Inches(0.7), Inches(1.4), Inches(5.65), Inches(2.7),
     "用一句大白话", BLUE,
     ["\"涨多了会回、跌多了会弹\"",
      "价格短期内被情绪推得过头（追涨杀跌），过几天往往被纠正、往回走",
      "就像橡皮筋拉太远会弹回来"])

card(s, Inches(6.75), Inches(1.4), Inches(5.85), Inches(2.7),
     "在本工具里具体怎么算", GREEN,
     ["看最近 30 天：拿今天价格和 30 天前比",
      "涨得越猛 → 反转分越偏\"看空\"；跌得越狠 → 越偏\"看多\"",
      "所以短期策略是\"逆着最近的极端走\""])

card(s, Inches(0.7), Inches(4.35), Inches(11.9), Inches(2.3),
     "关键点", FG,
     ["\"短期\"在这里有明确定义 = 最近 30 天的价格变化（不是模糊的\"最近\"）",
      "它天生带\"逆势\"性质：最近涨太多，它反而提示谨慎；最近跌太狠，它反而提示机会",
      "为什么是 30 天：论文验证 + BTC 回测都用这个窗口，能较好捕捉情绪化过度反应"])

# SLIDE · 什么是中长期动量
s = add_slide()
add_title(s, "什么是\"中长期动量\"")

card(s, Inches(0.7), Inches(1.4), Inches(5.65), Inches(2.7),
     "用一句大白话", BLUE,
     ["\"强者恒强、弱者恒弱\"",
      "过去几个月一直涨的，往往还会接着涨；一直跌的，还会接着跌",
      "就像重物有惯性，趋势一旦形成不会马上停"])

card(s, Inches(6.75), Inches(1.4), Inches(5.85), Inches(2.7),
     "在本工具里具体怎么算", GREEN,
     ["看约 270 天的大趋势（BTC 适配后的窗口）",
      "但故意跳过最近 21 天——避免和\"短期反转\"抢同一段行情",
      "所以中长期策略是\"顺着大趋势走\""])

card(s, Inches(0.7), Inches(4.35), Inches(11.9), Inches(2.3),
     "关键点", FG,
     ["\"中长期\"也有明确定义 = 约 270 天的价格趋势（跳过最近 21 天）",
      "它天生带\"顺势\"性质：过去半年多在涨，它倾向继续看多",
      "\"跳过最近 21 天\"是论文的巧思：让动量只看大趋势，把最近的短期噪声留给反转因子去处理"])

# SLIDE · 两者为何不打架
s = add_slide()
add_title(s, "短期要逆、中长期要顺——不打架吗？")

card(s, Inches(0.7), Inches(1.45), Inches(11.9), Inches(1.5),
     "答案：不打架，因为它们看的是不同的时间尺度", BLUE,
     ["一个管\"最近几周的极端\"，一个管\"过去大半年的趋势\"，各司其职、互相补充。"])

card(s, Inches(0.7), Inches(3.15), Inches(5.65), Inches(3.4),
     "打个比方", FG,
     ["大方向（中长期动量）说：现在是上升趋势，倾向做多",
      "小择时（短期反转）说：最近几天暴跌超跌了，这是更好的进场点",
      "两者合起来 = \"在上升趋势里，等一个短期回调再进\"",
      "这比单看任何一个都聪明"])

card(s, Inches(6.75), Inches(3.15), Inches(5.85), Inches(3.4),
     "所以模型同时用两个因子", GREEN,
     ["顺势的动量分：判断大方向该不该做多",
      "逆势的反转分：判断此刻是不是好的进出点",
      "两个分数加权合成一个\"综合评分\"",
      "下一节讲：这个权重不是固定的，会根据市场情况自动调"])

# SLIDE · 噪声交易者是关键
s = add_slide()
add_title(s, "关键洞察：\"噪声\"决定谁说了算")

card(s, Inches(0.7), Inches(1.4), Inches(11.9), Inches(1.7),
     "什么是\"噪声交易者\"", BLUE,
     ["就是那些追涨杀跌、被情绪驱动、不看基本面乱买乱卖的人。市场里这种人的活跃程度，就是\"噪声\"大小。"])

card(s, Inches(0.7), Inches(3.25), Inches(5.65), Inches(3.3),
     "噪声大的时候", RED,
     ["情绪化交易多 → 价格容易被推过头",
      "\"过度反应\"频繁出现 → 短期反转的机会更明显",
      "此刻应该更相信\"反转分\""])

card(s, Inches(6.75), Inches(3.25), Inches(5.85), Inches(3.3),
     "噪声小的时候", GREEN,
     ["市场理性、趋势干净",
      "价格沿趋势平稳推进 → 动量更可靠",
      "此刻应该更相信\"动量分\""])

tb, tf = textbox(s, Inches(0.7), Inches(6.75), Inches(11.9), Inches(0.55))
p = tf.paragraphs[0]
r = p.add_run()
r.text = "模型正是用\"噪声大小\"实时调整反转/动量的权重——这就是它比固定参数指标聪明的地方。"
_set_font(r, 14, FG2)


# ═══════════════════════════════════════════════════════════
# SECTION 02 · 从论文到 BTC
# ═══════════════════════════════════════════════════════════
section_divider("02", "从论文到 BTC", "论文讲了什么 · 我如何把它搬过来")

# SLIDE · 论文做了什么
s = add_slide()
add_title(s, "论文做了什么")

card(s, Inches(0.7), Inches(1.45), Inches(11.9), Inches(1.5),
     "N. Jegadeesh, J. Luo, A. Subrahmanyam, S. Titman (2025)", FG,
     ['"Short-Term Reversals and Longer-Term Momentum around the World" · 发表于 Review of Financial Studies（金融学三大顶刊之一）'])

card(s, Inches(0.7), Inches(3.15), Inches(5.65), Inches(3.4),
     "三个核心发现", BLUE,
     ["研究了全球 39 个国家的股票市场",
      "用同一个模型，同时解释两个看似矛盾的现象：短期会反转、中长期会延续",
      "把两者联系起来的关键，是市场里的\"噪声交易者\""])

card(s, Inches(6.75), Inches(3.15), Inches(5.85), Inches(3.4),
     "为什么值得信", GREEN,
     ["顶刊论文，经过严格同行评审",
      "跨 39 国验证，不是只在一个市场碰巧成立",
      "有理论解释（噪声），不是纯粹的数据挖掘",
      "这给了我们一个可靠的\"方向判断\"底座"])

# SLIDE · 论文原始设定 vs 难题
s = add_slide()
add_title(s, "难题：论文是给\"一堆股票\"设计的")

card(s, Inches(0.7), Inches(1.45), Inches(5.65), Inches(3.0),
     "论文原始设定（股票）", FG,
     ["研究对象：39 国、成千上万只股票",
      "反转/动量：把所有股票横向排名",
      "做多排名靠前的、做空靠后的",
      "赚的是\"不同股票之间\"的价差"])

card(s, Inches(6.75), Inches(1.45), Inches(5.85), Inches(3.0),
     "搬到 BTC 的难题", RED,
     ["BTC 只有一个标的",
      "你没有\"39 个 BTC\"可以互相排名比较",
      "论文的\"横截面排序\"直接用不了",
      "→ 必须换一种思路"])

card(s, Inches(0.7), Inches(4.65), Inches(11.9), Inches(1.9),
     "一句话说清难题", BLUE,
     ["论文靠\"多只股票横向比较\"来定强弱；而 BTC 只有一个，无法横向比较。",
      "所以核心问题是：怎么在\"只有一个资产\"的情况下，保留论文的反转+动量思想？"])

# SLIDE · 解法：跨时间尺度
s = add_slide()
add_title(s, "解法：把\"跨资产\"改成\"跨时间尺度\"")

card(s, Inches(0.7), Inches(1.45), Inches(11.9), Inches(1.7),
     "核心改造", GREEN,
     ["不再比较\"不同股票\"，而是比较\"同一个 BTC 的不同时间尺度\"：",
      "用最近 30 天算\"反转分\"，用约 270 天算\"动量分\"——同一资产，两个尺度。"])

card(s, Inches(0.7), Inches(3.35), Inches(5.65), Inches(3.2),
     "论文思想被完整保留", BLUE,
     ["反转、动量、噪声三个核心因子都在",
      "只是输入从\"横截面\"变成\"时间序列\"",
      "理论骨架不变，换了适配加密市场的血肉"])

card(s, Inches(6.75), Inches(3.35), Inches(5.85), Inches(3.2),
     "噪声代理也换了", FG,
     ["股票用换手率、买卖价差衡量噪声",
      "BTC 换成加密市场能实时看到的：",
      "资金费率 + 波动率分档",
      "（资金费率反映合约多空拥挤程度，是很好的情绪/噪声信号）"])

# SLIDE · 参数改造对照表
s = add_slide()
add_title(s, "如何针对 BTC 改造论文的周期参数")

add_table(s, Inches(0.7), Inches(1.5), Inches(11.9),
          ["参数", "含义", "论文原始", "BTC 适配"],
          [
              ["反转窗口", "短期反转看多少天", "30 天", ("30 天（不变）", FG2)],
              ["动量窗口", "中长期动量看多少天", "330 天", ("270 天", BLUE)],
              ["动量跳过", "动量跳过最近多少天", "30 天", ("21 天", BLUE)],
              ["入场阈值", "评分超过多少才发信号", "1.0", ("0.5（更灵敏）", BLUE)],
              ["噪声敏感度", "权重随噪声调整的力度", "0.5", ("0.3（更稳）", BLUE)],
              ["减半增强", "减半后增强动量（BTC 独有）", "无", ("540 天 / ×1.2", GREEN)],
              ["资金费率", "作为噪声因子（BTC 独有）", "无", ("占噪声 30%", GREEN)],
          ],
          col_widths=[Inches(2.2), Inches(4.3), Inches(2.5), Inches(2.9)],
          cell_size=13, header_size=14, row_h=Inches(0.6))

tb, tf = textbox(s, Inches(0.7), Inches(6.45), Inches(11.9), Inches(0.7))
p = tf.paragraphs[0]
r = p.add_run()
r.text = "改造逻辑：BTC 波动更快、周期更短，所以动量窗口缩短、阈值调灵敏；并加入减半、资金费率等加密市场独有因子。"
_set_font(r, 13, FG2)

# SLIDE · 综合评分怎么算
s = add_slide()
add_title(s, "综合评分是怎么算出来的")

formula_box = s.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(1.2), Inches(1.5), Inches(10.9), Inches(1.2))
_set_rounded_corners(formula_box, 7000)
formula_box.fill.solid()
formula_box.fill.fore_color.rgb = _hex_to_rgb(BLUE_DARK_HEX)
formula_box.line.color.rgb = _hex_to_rgb("2E4160")
formula_box.line.width = Pt(1)
_add_shadow(formula_box, blur=50000, dist=25000, alpha=35000)
tf = formula_box.text_frame
tf.word_wrap = True
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "综合评分 = 动量权重 × 动量分 + 反转权重 × 反转分"
_set_font(r, 22, BLUE, bold=True)

card(s, Inches(0.7), Inches(2.95), Inches(5.65), Inches(3.6),
     "三步走", FG,
     ["① 两个分各自标准化（z-score），拉到可比较的尺度",
      "② 权重按当前\"噪声大小\"实时调整：噪声大→反转权重↑、动量权重↓",
      "③ 资金费率不是独立一项，而是作为噪声因子（占噪声 30%）来影响这个权重",
      "④ 再乘一个\"机制置信度\"做微调"], body_size=14)

card(s, Inches(6.75), Inches(2.95), Inches(5.85), Inches(3.6),
     "两个容易误会的点", BLUE,
     ["评分不是\"越高越好\"，而是穿过阈值才发信号（见下一节）",
      "\"资金费率占 30% 权重\"指的是它在噪声合成里的比重，不是直接进评分公式",
      "所以公式只有两项（动量+反转），资金费率藏在\"权重怎么定\"这一步里"], body_size=14)


# ═══════════════════════════════════════════════════════════
# SECTION 03 · 信号怎么来
# ═══════════════════════════════════════════════════════════
section_divider("03", "信号怎么来", "从评分到买卖点")

# SLIDE · 从评分到信号
s = add_slide()
add_title(s, "评分怎么变成买卖信号")

card(s, Inches(0.7), Inches(1.4), Inches(11.9), Inches(2.4),
     "三种信号", FG,
     ["综合评分\"向上穿过\"阈值 → 做多信号（绿色 ↑）",
      "综合评分\"向下穿过\"阈值 → 做空信号（红色 ↓）",
      "多个极端条件同时触发 → 级联预警（强信号）"])

card(s, Inches(0.7), Inches(4.05), Inches(5.65), Inches(2.5),
     "为什么是\"穿过\"而不是\"高于\"", BLUE,
     ["信号在评分\"上穿/下穿\"阈值的那一刻触发（边沿触发）",
      "不是\"只要评分高就一直发信号\"",
      "这样避免同一波行情反复重复报信号"])

card(s, Inches(6.75), Inches(4.05), Inches(5.85), Inches(2.5),
     "\"死区\"设计", GREEN,
     ["评分绝对值很小时（接近 0）不发任何信号",
      "因为方向不明确，宁可不动",
      "只在信号足够清晰时才出手"])

# SLIDE · 只靠原始信号够不够
s = add_slide()
add_title(s, "只靠论文信号，够好吗？")

kpi_card(s, Inches(0.9), Inches(1.5), Inches(3.5), Inches(1.7),
         "+10.71%", "做多信号 30 天平均收益", GREEN)
kpi_card(s, Inches(4.75), Inches(1.5), Inches(3.5), Inches(1.7),
         "58%", "做多胜率", BLUE)
kpi_card(s, Inches(8.6), Inches(1.5), Inches(3.85), Inches(1.7),
         "85 次", "做多信号有效样本", FG)

card(s, Inches(0.7), Inches(3.5), Inches(11.9), Inches(3.1),
     "还不错，但有个明显问题", RED,
     ["原始信号（不加任何过滤）做多 30 天平均 +10.71%、胜率约 58%——已经能赚，但不够稳。",
      "问题：有些做多信号出现在\"下跌趋势里\"，属于逆势抄底，容易被套。",
      "打个比方：论文告诉你\"该买了\"，但它不知道你是在牛市里买、还是在熊市半山腰接飞刀。",
      "→ 这正是下一步要叠加\"均线滤波\"的动机：把逆势的信号筛掉。"])

# SLIDE · 三重 MA 滤波
s = add_slide()
add_title(s, "买卖点升级：三重均线滤波")

card(s, Inches(0.7), Inches(1.45), Inches(5.65), Inches(3.2),
     "做多滤波（三条件全满足才保留）", GREEN,
     ["① 价格 > MA110（大趋势向上）",
      "② MA6 > MA103（短期也在走强）",
      "③ 价格 > EMA50（中期向上）",
      "只有论文做多信号 + 三条件全绿 = 高质量买点"])

card(s, Inches(6.75), Inches(1.45), Inches(5.85), Inches(3.2),
     "做空滤波（价格在均线下方）", RED,
     ["要求价格 < MA110 或 价格 < EMA50",
      "即价格确实在关键均线下方",
      "才保留做空信号"])

card(s, Inches(0.7), Inches(4.85), Inches(11.9), Inches(1.8),
     "为什么有效 + 用户可自己调", BLUE,
     ["均线就是\"趋势的过滤网\"：把逆着大趋势的信号筛掉，只留顺势的。逆势单是亏损的主要来源。",
      "这些滤波开关在网站上可自由勾选：全勾最严格、也可只留一两个、或全关看原始信号，实时重算收益。"])


# ═══════════════════════════════════════════════════════════
# SECTION 04 · 超额收益
# ═══════════════════════════════════════════════════════════
section_divider("04", "超额收益", "买卖点能捕获多少 · 为什么 30 天最好")

# SLIDE · 对比表
s = add_slide()
add_title(s, "加了滤波，收益提升多少（做多）")

add_table(s, Inches(0.9), Inches(1.55), Inches(11.5),
          ["持有期", "论文原始信号", "三重滤波后", "相对提升"],
          [
              ["1 天", "+1.12%", ("+1.74%", GREEN), ("+55%", GREEN)],
              ["3 天", "+2.01%", ("+3.50%", GREEN), ("+74%", GREEN)],
              ["7 天", "+3.25%", ("+6.03%", GREEN), ("+86%", GREEN)],
              ["30 天", "+10.71%", ("+20.52%", GREEN), ("+92%", GREEN)],
          ],
          col_widths=[Inches(2.5), Inches(3.0), Inches(3.0), Inches(3.0)])

card(s, Inches(0.9), Inches(5.05), Inches(11.5), Inches(1.7),
     "核心结论", GREEN,
     ["每个持有期，滤波后都明显更高；30 天做多从 +10.71% 提升到 +20.52%（超额 +9.82 个百分点、相对 +92%）。",
      "胜率也从约 58% 提到约 62%。过滤逆势信号，是拿到超额回报的关键。"])

# SLIDE · 为什么 30 天做多最佳
s = add_slide()
add_title(s, "为什么\"30 天做多\"表现最佳")

card(s, Inches(0.7), Inches(1.4), Inches(11.9), Inches(1.6),
     "现象：持有越久，滤波带来的超额越大", BLUE,
     ["滤波后各持有期收益：1 天 +1.74% → 3 天 +3.50% → 7 天 +6.03% → 30 天 +20.52%，一路递增。"])

card(s, Inches(0.7), Inches(3.15), Inches(5.65), Inches(3.4),
     "为什么会这样", GREEN,
     ["做多信号本质是在\"顺势\"下注",
      "动量（趋势惯性）需要时间才能充分兑现",
      "持有太短（1-3 天）：趋势还没走出来，收益小、还容易被手续费吃掉",
      "持有到 30 天：趋势充分展开，动量红利最大化"])

card(s, Inches(6.75), Inches(3.15), Inches(5.85), Inches(3.4),
     "数据也印证", FG,
     ["30 天档：超额绝对值 +9.82 个百分点，全期最大",
      "30 天档：相对提升 +92%，也是全期最大",
      "→ 这套策略的\"最佳持有期\"就落在约 30 天（1 个月级别）",
      "适合中长线、而非日内高频"])

# SLIDE · 做多最强阿尔法
s = add_slide()
add_title(s, "做多 —— 整套策略最强的部分")

kpi_card(s, Inches(0.7), Inches(1.5), Inches(3.6), Inches(1.7),
         "+20.52%", "滤波后做多 30 日均收益", GREEN)
kpi_card(s, Inches(4.65), Inches(1.5), Inches(3.6), Inches(1.7),
         "67%", "3 天胜率", GREEN)
kpi_card(s, Inches(8.6), Inches(1.5), Inches(3.95), Inches(1.7),
         "85 → 42", "信号次数（少而精）", BLUE)

card(s, Inches(0.7), Inches(3.5), Inches(11.9), Inches(3.1),
     "怎么解读", FG,
     ["三重滤波把做多信号从 85 次精简到 42 次 —— 宁可少做，也要做对",
      "各持有期收益全面提升：1 天 +55%、3 天 +74%、7 天 +86%、30 天 +92%",
      "胜率区间从约 58-65% 抬升到约 62-67%",
      "这是整套策略里最稳定、最值得依赖的部分（下一节会讲它的边界）"])

# SLIDE · 做空诚实局限
s = add_slide()
add_title(s, "做空 —— 诚实面对它的局限")

add_table(s, Inches(0.9), Inches(1.5), Inches(11.5),
          ["持有期", "论文做空信号", "滤波后做空", "变化"],
          [
              ["1 天", "+0.44% (53%)", ("-0.26% (44%)", FG2), "走弱"],
              ["7 天", "+0.76% (52%)", ("-1.19% (44%)", FG2), "走弱"],
              ["30 天", ("-3.47% (43%)", RED), ("-2.87% (38%)", RED), "略改善"],
          ],
          col_widths=[Inches(2.3), Inches(3.2), Inches(3.2), Inches(2.8)],
          row_h=Inches(0.55))

card(s, Inches(0.9), Inches(4.05), Inches(11.5), Inches(2.5),
     "为什么做空效果一般", RED,
     ["BTC 长期是上行的，做空天然逆风：做空 30 天平均约 -3.47%",
      "均线下方常是\"超跌区\"，加滤波后短期反而更弱、容易遇到反弹",
      "\"胜率 43%\"这里指\"价格继续下跌（做空获利）的比例\"——不到一半",
      "→ 做空的定位是风险预警和对冲，不适合当主进攻手段"])

# SLIDE · 核心数据总结
s = add_slide()
add_title(s, "超额收益 · 核心数据总结")

kpi_card(s, Inches(0.7), Inches(1.5), Inches(3.6), Inches(1.5),
         "+20.52%", "做多 30 日均收益", GREEN)
kpi_card(s, Inches(4.65), Inches(1.5), Inches(3.6), Inches(1.5),
         "+92%", "30 日超额提升", BLUE)
kpi_card(s, Inches(8.6), Inches(1.5), Inches(3.95), Inches(1.5),
         "67%", "3 日胜率（做多）", GREEN)

card(s, Inches(0.7), Inches(3.3), Inches(5.65), Inches(3.3),
     "做多结论", GREEN,
     ["滤波后 30 日收益近乎翻倍（+10.71% → +20.52%）",
      "胜率约 58-65% → 62-67%",
      "信号 85 → 42（少而精）",
      "策略最强、最值得依赖的部分"])

card(s, Inches(6.75), Inches(3.3), Inches(5.85), Inches(3.3),
     "做空结论", RED,
     ["做空整体偏弱（BTC 长期上行）",
      "无滤波 30 日约 -3.47%，滤波改善有限",
      "定位：防御/对冲，不做主策略"])

tb, tf = textbox(s, Inches(0.7), Inches(6.7), Inches(11.9), Inches(0.5))
p = tf.paragraphs[0]
r = p.add_run()
r.text = "信号库共 200 条（做多 86 / 做空 108 / 级联 6）；上述收益取有 30 日完整数据的样本统计。"
_set_font(r, 12, MUTE)


# ═══════════════════════════════════════════════════════════
# SECTION 05 · 验证与边界
# ═══════════════════════════════════════════════════════════
section_divider("05", "验证与边界", "样本外表现 · 什么时候该用")

# SLIDE · 样本外验证（诚实页）
s = add_slide()
add_title(s, "最诚实的一页：样本外验证")

card(s, Inches(0.7), Inches(1.4), Inches(11.9), Inches(1.35),
     "什么是样本外", BLUE,
     ["参数是用 2017-2023 的数据\"训练\"出来的；再拿 2024 年后它没见过的数据检验——这才算数。"])

add_table(s, Inches(0.9), Inches(2.95), Inches(11.5),
          ["30 天做多", "训练期(2017-2023)", "样本外(2024 起)"],
          [
              ["平均收益", ("+7.16%", GREEN), ("-3.32%", RED)],
              ["做多胜率", ("64.6%", GREEN), ("20.0%", RED)],
              ["夏普比率", ("1.06", GREEN), ("-1.07", RED)],
          ],
          col_widths=[Inches(3.5), Inches(4.0), Inches(4.0)],
          row_h=Inches(0.55))

card(s, Inches(0.9), Inches(5.0), Inches(11.5), Inches(1.65),
     "这说明什么（必须讲清）", RED,
     ["训练期很漂亮，但样本外 30 天明显衰减、甚至转负——存在过拟合风险。",
      "所以：全样本的 +20.52% 是历史统计的\"上限感受\"，不代表未来能稳定复现。把它当\"辅助判断\"，不是\"稳赚承诺\"。"])

# SLIDE · 什么情况最适合
s = add_slide()
add_title(s, "什么情况下最适合用它？")

card(s, Inches(0.5), Inches(1.5), Inches(3.9), Inches(4.0),
     "最适合", GREEN,
     ["明确趋势市：价格站上 MA110/EMA50，做多胜率 62-67%",
      "中长线持有（约 1-4 周）：30 日收益差距最大",
      "右侧交易者：愿放弃最低点、换更高确定性"], body_size=14)

card(s, Inches(4.65), Inches(1.5), Inches(3.9), Inches(4.0),
     "要谨慎", FG,
     ["震荡/横盘市：均线频繁穿越，滤波反复触发又失效",
      "做空：BTC 长期上行，整体为负，仅作预警",
      "短持有期（1-3 日）：超额小，成本占比高"], body_size=14)

card(s, Inches(8.8), Inches(1.5), Inches(3.9), Inches(4.0),
     "不适合", RED,
     ["日内/高频：信号约 30 天一次，频率不匹配",
      "黑天鹅急跌：任何趋势滤波都滞后",
      "当全自动交易机器人：它是决策辅助"], body_size=14)

card(s, Inches(0.5), Inches(5.7), Inches(12.2), Inches(1.15),
     "一句话", BLUE,
     ["超额收益来自\"在趋势里做多、并用均线把逆势信号筛掉\"。趋势越明确、持有期越长，价值越大；越震荡、越短线，价值越小。"])


# ═══════════════════════════════════════════════════════════
# SECTION 06 · 局限与总结
# ═══════════════════════════════════════════════════════════
section_divider("06", "局限与总结", "诚实面对模型的边界")

# SLIDE · 需要注意的事项
s = add_slide()
add_title(s, "需要注意的事项")

caveats = [
    "历史回测 ≠ 未来表现 —— 市场结构会变化",
    "样本外衰减 —— 训练期漂亮，但 2024 年后 30 天表现明显走弱甚至转负",
    "论文参数来自股票市场 —— 虽已做 BTC 适配，仍是近似而非精确",
    "低频策略 —— 约 30 天触发一次，不适合日内或高频",
    "做空效果一般 —— BTC 长期向上，做空天然逆风、胜率偏低",
    "极端行情下信号会滞后 —— 黑天鹅事件中任何趋势指标都滞后",
]
for idx, text in enumerate(caveats):
    y = Inches(1.5) + Inches(0.62) * idx
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL,
                             Inches(0.85), y + Inches(0.08), Pt(9), Pt(9))
    dot.fill.solid()
    dot.fill.fore_color.rgb = RED
    dot.line.fill.background()
    tb, tf = textbox(s, Inches(1.15), y, Inches(11.4), Inches(0.5))
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    _set_font(r, 15, FG)

card(s, Inches(0.7), Inches(5.4), Inches(11.9), Inches(1.35),
     "正确定位", FG,
     ["BTC Momentum 是辅助决策工具，帮你过滤噪声、提供学术框架支撑的方向判断，但最终决策权在你——不是自动赚钱机器。"])

# SLIDE · 总结
s = add_slide(decor=(Inches(9.2), Inches(4.8), Inches(5), BLUE_HEX))
add_title(s, "总结")
glow_circle(s, Inches(-1.5), Inches(5), Inches(4), BLUE_HEX, alpha=4000)

card(s, Inches(0.7), Inches(1.5), Inches(5.65), Inches(2.5),
     "理论基础", BLUE,
     ["JLST (2025) RFS 论文 — 39 国验证",
      "短期反转(30天) + 中长期动量(270天) 双因子",
      "噪声交易者是关键中介变量"])

card(s, Inches(6.75), Inches(1.5), Inches(5.85), Inches(2.5),
     "落地效果", GREEN,
     ["论文信号 + 三重 MA 滤波",
      "做多 30 日 +20.52%（胜率约 62%）",
      "超额比原始信号相对提升 92%"])

card(s, Inches(0.7), Inches(4.2), Inches(11.9), Inches(2.5),
     "从论文到实操的完整链条", BLUE,
     ["论文 → BTC 映射：把\"跨股票横截面\"改造成\"跨时间尺度\"，噪声代理换成资金费率+波动率",
      "买卖点作用：三重 MA 滤波把逆势信号筛掉，做多超额 +9.82 个百分点、相对提升 92%",
      "诚实边界：样本外会衰减、做空逆风、低频、参数近似——它是有学术支撑的方向判断，不是自动赚钱机器"])

# ── 保存 ──
out = Path(__file__).resolve().parent.parent / "BTC_Momentum_演示.pptx"
prs.save(str(out))
print(f"OK  已生成 {out}  （共 {slide_counter[0]} 页）")
