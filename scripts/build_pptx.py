#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 BTC Momentum 演示 PPT。

设计语言：单一深色背景（微渐变，无跨元素装饰线）+ 圆角卡片（纯色填充、
弱阴影）+ 一个主导色（蓝）+ 语义色仅用于真实的正负数据（绿/红）+
全篇 Microsoft YaHei（通过 bold 标记区分粗细，不依赖具体字重变体）+
不在标题下加装饰线、不在卡片上加色条（这两者是明显的 AI 生成痕迹）。

内容与 slides.html / TALK_SCRIPT.md 一一对应；回测数据冻结于 2026-09-07。

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
TOTAL_SLIDES = 21


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
# SLIDE 1 · 封面
# ═══════════════════════════════════════════════════════════
s = add_slide(decor=(Inches(-2), Inches(-1.5), Inches(7), BLUE_HEX))
glow_circle(s, Inches(9.5), Inches(4.5), Inches(6), BLUE_HEX, alpha=4000)

tb, tf = textbox(s, Inches(0), Inches(2.15), EMU_W, Inches(1.3), MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "BTC Momentum"
_set_font(r, 54, FG, bold=True, font=FONT_NUM)

tb, tf = textbox(s, Inches(0), Inches(3.35), EMU_W, Inches(0.7), MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "JLST 反转-动量多信号策略系统"
_set_font(r, 22, FG2)

tb, tf = textbox(s, Inches(0), Inches(4.7), EMU_W, Inches(1.6), MSO_ANCHOR.MIDDLE)
for i, (line, sz, it) in enumerate([
    ("基于 Jegadeesh, Luo, Subrahmanyam & Titman (2025)", 15, False),
    ("Review of Financial Studies — 全球顶级金融学术期刊", 14, True),
    ("回测数据截止 2026-09-07", 14, False),
]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(7)
    r = p.add_run()
    r.text = line
    _set_font(r, sz, FG2, italic=it)


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
     "BTC Momentum 的价值", BLUE,
     ["JLST 理论信号 + MA 均线多重过滤 → 用学术框架确定方向，用趋势滤波提高精度"])


# ═══════════════════════════════════════════════════════════
# SLIDE 3 · 论文出处
# ═══════════════════════════════════════════════════════════
s = add_slide()
add_title(s, "论文出处")

card(s, Inches(0.7), Inches(1.45), Inches(11.9), Inches(1.45),
     "N. Jegadeesh, J. Luo, A. Subrahmanyam, S. Titman (2025)", FG,
     ['"Short-Term Reversals and Longer-Term Momentum around the World"',
      "Review of Financial Studies (RFS) — 金融学三大顶刊之一"])

card(s, Inches(0.7), Inches(3.15), Inches(5.65), Inches(3.4),
     "论文做了什么？", BLUE,
     ["研究了 39 个国家股票市场",
      "用同一模型解释\"短期反转\"与\"中长期动量\"如何同时存在",
      "发现噪声交易是两者联系的关键纽带"])

card(s, Inches(6.75), Inches(3.15), Inches(5.85), Inches(3.4),
     "我们做了什么？", GREEN,
     ["将论文理论应用到 BTC 市场",
      "构建实时仪表盘，每日更新",
      "叠加 MA 多信号过滤，将理论收益变为实操策略",
      "开放可配置滤波条件，用户自行调优"])


# ═══════════════════════════════════════════════════════════
# SECTION 01 · 核心概念
# ═══════════════════════════════════════════════════════════
section_divider("01", "核心概念", "短期反转 & 中长期动量")

# -- 反转 vs 动量
s = add_slide()
add_title(s, "短期反转 vs 中长期动量")

card(s, Inches(0.7), Inches(1.5), Inches(5.65), Inches(2.5),
     "短期反转", BLUE,
     ["\"涨多了会回、跌多了会弹\"",
      "短期价格含大量情绪化过度反应，几天内被纠正",
      "策略：短期逆着最近几天的极端走"])

card(s, Inches(6.75), Inches(1.5), Inches(5.85), Inches(2.5),
     "中长期动量", GREEN,
     ["\"强者恒强、弱者恒弱\"",
      "过去几个月上涨的往往继续涨，趋势有惯性",
      "策略：中长期顺着大趋势走"])

card(s, Inches(0.7), Inches(4.2), Inches(11.9), Inches(2.6),
     "如何共存？——噪声交易者是关键", FG,
     ["两者作用在不同时间尺度上，并不冲突",
      "噪声交易者制造短期过度波动（反转来源），其行为在中期形成趋势（动量来源）",
      "噪声越大 → 反转机会越明显；噪声越小、趋势越干净 → 动量越可靠",
      "模型用噪声大小动态调整反转/动量权重 —— 这是它比固定参数指标聪明的地方"])

# -- 综合评分公式
s = add_slide()
add_title(s, "综合评分公式")

formula_box = s.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Inches(1.0), Inches(2.0), Inches(11.3), Inches(1.5))
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
r.text = "Composite = w₁·z(Rev) + w₂·z(Mom) + w₃·z(FundingRate)"
_set_font(r, 22, BLUE, bold=True, font=FONT_NUM)

tb, tf = textbox(s, Inches(0.7), Inches(3.9), Inches(11.9), Inches(2.8))
add_paragraphs(tf, [
    ("三个分量先做标准化（z-score），再按动态权重加权求和。", FG, False, 16),
    ("权重不是固定的，而是根据当前市场噪声大小实时计算。", FG, False, 16),
    ("BTC 适配版额外引入资金费率、波动率分档、减半周期等加密市场特有因子。", FG2, False, 15),
], first=True)

# -- 如何映射到 BTC
s = add_slide()
add_title(s, "如何把股票论文映射到 BTC")

card(s, Inches(0.7), Inches(1.45), Inches(5.65), Inches(3.3),
     "论文原始设定（股票）", FG,
     ["研究对象：39 国股票市场的横截面",
      "反转 / 动量：跨个股排序、多空组合",
      "噪声代理：换手率、买卖价差等微观结构变量",
      "因子在个股之间比较，赚的是横截面价差"])

card(s, Inches(6.75), Inches(1.45), Inches(5.85), Inches(3.3),
     "BTC 适配版（单一资产·时序）", GREEN,
     ["研究对象：BTC 单一资产的时间序列",
      "反转 / 动量：同一资产不同时间尺度的自相关",
      "噪声代理：资金费率 + 波动率分档（占 30% 权重）",
      "额外引入减半周期等加密特有因子"])

card(s, Inches(0.7), Inches(4.95), Inches(11.9), Inches(1.75),
     "映射的核心难点与解法", BLUE,
     ["难点：论文靠「多只股票横向比较」，BTC 只有一个标的，无法做横截面排序",
      "解法：把「跨资产比较」改成「跨时间尺度比较」——短期 z 分做反转、中长期 z 分做动量，噪声代理换成加密市场能实时观测的资金费率与波动率"])


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
      "多个条件同时触发 → 级联预警（强信号）"])

card(s, Inches(0.7), Inches(3.95), Inches(11.9), Inches(2.7),
     "JLST 原始信号前瞻回报（无 MA 滤波）", GREEN,
     ["做多信号 30 天平均 +10.71%，胜率 58%",
      "问题：部分做多信号出现在下跌趋势里（逆势抄底），容易吃套",
      "→ 这正是叠加均线滤波的动机"])


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
      "只有 JLST 做多信号 + 三条件全满足 = 高质量信号"])

card(s, Inches(6.75), Inches(1.5), Inches(5.85), Inches(3.2),
     "做空滤波（价格在均线下方）", RED,
     ["价格 < MA110",
      "价格 < EMA50",
      "要求价格在均线下方才保留做空信号"])

card(s, Inches(0.7), Inches(4.9), Inches(11.9), Inches(1.8),
     "为什么 MA 滤波有效 + 用户可配置", BLUE,
     ["均线是趋势的过滤网：筛掉逆势信号（亏损主要来源），只留顺势信号",
      "三个开关用户自主控制：全勾最严格、可只留一两个、可全取消看原始信号，实时重新统计前瞻收益"])


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
      "胜率 58% → 62%。过滤逆势信号是获取超额回报的关键"])

# -- 做多最强阿尔法
s = add_slide()
add_title(s, "做多策略 — 最强阿尔法")

kpi_card(s, Inches(0.7), Inches(1.5), Inches(3.6), Inches(1.7),
         "+20.52%", "多信号做多 30 日均收益", GREEN)
kpi_card(s, Inches(4.65), Inches(1.5), Inches(3.6), Inches(1.7),
         "67%", "3 天胜率", GREEN)
kpi_card(s, Inches(8.6), Inches(1.5), Inches(3.95), Inches(1.7),
         "85 → 42", "信号次数（少而精）", BLUE)

card(s, Inches(0.7), Inches(3.5), Inches(11.9), Inches(3.2),
     "解读", FG,
     ["三重滤波把做多信号从 85 次精简到 42 次 —— 宁可少做，也要做对",
      "各持有期收益全面提升：1天 +55%、3天 +74%、7天 +86%、30天 +92%",
      "胜率区间从 58-65% 抬升到 62-67%",
      "这是整套策略里最稳定、最值得依赖的部分"])

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
     "解读", RED,
     ["BTC 长期上行，做空天然逆风：做空 30 天平均 -3.61%、胜率仅 42%",
      "均线下方常是超跌区，加滤波后短期反而更弱、易遇反弹",
      "做空的定位是风险预警与对冲，不宜作为主进攻策略"])

# -- 核心数据总结
s = add_slide()
add_title(s, "核心数据总结")

kpi_card(s, Inches(0.7), Inches(1.5), Inches(3.6), Inches(1.5),
         "+20.52%", "做多 30 日均收益", GREEN)
kpi_card(s, Inches(4.65), Inches(1.5), Inches(3.6), Inches(1.5),
         "+92%", "30 日超额提升", BLUE)
kpi_card(s, Inches(8.6), Inches(1.5), Inches(3.95), Inches(1.5),
         "67%", "3 日胜率（做多）", GREEN)

card(s, Inches(0.7), Inches(3.3), Inches(5.65), Inches(3.3),
     "做多策略结论", GREEN,
     ["MA 滤波后 30 日收益近乎翻倍",
      "胜率 58-65% → 62-67%",
      "信号数 85 → 42（少而精）",
      "每次信号平均 30 日赚 20.52%"])

card(s, Inches(6.75), Inches(3.3), Inches(5.85), Inches(3.3),
     "做空策略结论", RED,
     ["做空收益整体偏弱（BTC 长期上行）",
      "无滤波 30 日 -3.61%，滤波改善有限",
      "信号数 109 → 45（过滤假信号）",
      "更适合防御/对冲，不做主策略"])


# ═══════════════════════════════════════════════════════════
# SECTION 05 · 适用边界
# ═══════════════════════════════════════════════════════════
section_divider("05", "适用边界", "什么时候该用它、什么时候别用")

# -- 适用场景
s = add_slide()
add_title(s, "什么情况下最适合用多策略？")

card(s, Inches(0.5), Inches(1.5), Inches(3.9), Inches(4.0),
     "最适合", GREEN,
     ["明确趋势市：价格站上 MA110/EMA50，做多胜率 62-67%",
      "中长线持有（1-4 周）：30 日收益差距最大",
      "右侧交易者：愿放弃最低点换确定性"], body_size=14)

card(s, Inches(4.65), Inches(1.5), Inches(3.9), Inches(4.0),
     "谨慎", FG,
     ["震荡/横盘市：均线频繁穿越，滤波反复触发又失效",
      "做空：BTC 长期上行，整体为负，仅作预警",
      "短持有期（1-3 日）：超额小，成本占比高"], body_size=14)

card(s, Inches(8.8), Inches(1.5), Inches(3.9), Inches(4.0),
     "不适合", RED,
     ["日内/高频：信号约 30 天一次，频率不匹配",
      "黑天鹅急跌：任何趋势滤波都滞后",
      "当自动交易机器人：它是决策辅助，非全自动"], body_size=14)

card(s, Inches(0.5), Inches(5.7), Inches(12.2), Inches(1.15),
     "一句话", BLUE,
     ["多策略的超额收益来自「在趋势里做多、并用均线把逆势信号筛掉」。趋势越明确、持有期越长，价值越大；越震荡、越短线，价值越小。"])


# ═══════════════════════════════════════════════════════════
# SECTION 06 · 局限与总结
# ═══════════════════════════════════════════════════════════
section_divider("06", "局限与总结", "诚实面对模型的边界")

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
    dot = s.shapes.add_shape(MSO_SHAPE.OVAL,
                             Inches(0.85), y + Inches(0.08), Pt(9), Pt(9))
    dot.fill.solid()
    dot.fill.fore_color.rgb = RED
    dot.line.fill.background()
    tb, tf = textbox(s, Inches(1.15), y, Inches(11.2), Inches(0.5))
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    _set_font(r, 15, FG)

card(s, Inches(0.7), Inches(5.4), Inches(11.9), Inches(1.35),
     "正确定位", FG,
     ["BTC Momentum 是辅助决策工具，帮你过滤噪声、提供学术框架支撑的方向判断，但最终决策权在你——不是自动赚钱机器。"])

# -- 总结
s = add_slide(decor=(Inches(9.2), Inches(4.8), Inches(5), BLUE_HEX))
add_title(s, "总结")
glow_circle(s, Inches(-1.5), Inches(5), Inches(4), BLUE_HEX, alpha=4000)

card(s, Inches(0.7), Inches(1.5), Inches(5.65), Inches(2.5),
     "理论基础", BLUE,
     ["JLST (2025) RFS 论文 — 39 国验证",
      "短期反转 + 中长期动量双因子",
      "噪声交易者是关键中介变量"])

card(s, Inches(6.75), Inches(1.5), Inches(5.85), Inches(2.5),
     "多信号策略", GREEN,
     ["JLST 信号 + MA110/MA6>MA103/EMA50",
      "做多 30 日 +20.52%（胜率 62%）",
      "超额比原始信号提升 92%"])

card(s, Inches(0.7), Inches(4.2), Inches(11.9), Inches(2.5),
     "从论文到实操的完整链条", BLUE,
     ["论文 → BTC 映射：把「跨股票横截面」改造成「跨时间尺度」，噪声代理换成资金费率",
      "买卖点作用：三重 MA 滤波把逆势信号筛掉，做多超额 +9.82 个百分点、相对提升 92%",
      "诚实边界：做空逆风、低频、参数近似——它是有学术支撑的方向判断，不是自动赚钱机器"])

# ── 保存 ──
out = Path(__file__).resolve().parent.parent / "BTC_Momentum_演示.pptx"
prs.save(str(out))
print(f"OK  已生成 {out}  （共 {slide_counter[0]} 页）")
