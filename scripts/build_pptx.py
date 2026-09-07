#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从 slides.html 的内容生成真正的 PowerPoint .pptx 文件。

内容与 slides.html 一一对应（暗色主题、封面、分节、指标规则表、回测数字等）。
回测数据已冻结（截止 2026-09-07），此脚本只负责排版，不重新计算任何东西。

用法：python scripts/build_pptx.py
输出：BTC_Momentum_演示.pptx（项目根目录）
"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ── 配色（对齐 slides.html） ──
BG      = RGBColor(0x0D, 0x11, 0x17)
CARD    = RGBColor(0x16, 0x1B, 0x22)
FG      = RGBColor(0xE6, 0xED, 0xF3)
FG2     = RGBColor(0x8B, 0x94, 0x9E)
GREEN   = RGBColor(0x3F, 0xB9, 0x50)
RED     = RGBColor(0xF8, 0x51, 0x49)
BLUE    = RGBColor(0x58, 0xA6, 0xFF)
ORANGE  = RGBColor(0xD2, 0x99, 0x22)
PURPLE  = RGBColor(0xBC, 0x8C, 0xFF)
ACCENT  = BLUE

FONT = "Microsoft YaHei"

EMU_W, EMU_H = Inches(13.333), Inches(7.5)  # 16:9

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]


def add_slide():
    s = prs.slides.add_slide(BLANK)
    # 背景
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    return s


def _set_font(run, size, color, bold=False, italic=False):
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = FONT


def textbox(slide, left, top, width, height, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    return tb, tf


def add_title(slide, text, color=FG):
    tb, tf = textbox(slide, Inches(0.6), Inches(0.35), Inches(12.1), Inches(0.9))
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = text
    _set_font(r, 30, color, bold=True)
    # 下划线条
    line = slide.shapes.add_shape(1, Inches(0.6), Inches(1.15), Inches(3.2), Pt(3))
    line.fill.solid(); line.fill.fore_color.rgb = ACCENT
    line.line.fill.background()
    return tb


def add_paragraphs(tf, items, base_size=16, first=False):
    """items: list of (text, color, bold, size_override or None)"""
    for i, it in enumerate(items):
        text, color, bold = it[0], it[1], it[2]
        size = it[3] if len(it) > 3 and it[3] else base_size
        p = tf.paragraphs[0] if (first and i == 0) else tf.add_paragraph()
        p.space_after = Pt(6)
        r = p.add_run(); r.text = text
        _set_font(r, size, color, bold=bold)


def bullet_box(slide, left, top, width, height, title, title_color, bullets,
               fill=CARD, accent=None):
    box = slide.shapes.add_shape(1, left, top, width, height)
    box.fill.solid(); box.fill.fore_color.rgb = fill
    box.line.color.rgb = accent if accent else RGBColor(0x30, 0x36, 0x3D)
    box.line.width = Pt(1)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.18); tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.12)
    p0 = tf.paragraphs[0]
    r = p0.add_run(); r.text = title
    _set_font(r, 15, title_color, bold=True)
    p0.space_after = Pt(6)
    for b in bullets:
        p = tf.add_paragraph(); p.space_after = Pt(3)
        r = p.add_run(); r.text = "• " + b
        _set_font(r, 12.5, FG)
    return box


def add_table(slide, left, top, width, headers, rows, col_widths=None,
              header_size=12, cell_size=11, row_h=Inches(0.42)):
    nrows = len(rows) + 1
    ncols = len(headers)
    height = row_h * nrows
    gtbl = slide.shapes.add_table(nrows, ncols, left, top, width, height).table
    if col_widths:
        for i, w in enumerate(col_widths):
            gtbl.columns[i].width = w
    # header
    for c, h in enumerate(headers):
        cell = gtbl.cell(0, c)
        cell.fill.solid(); cell.fill.fore_color.rgb = CARD
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = cell.text_frame.paragraphs[0]; para.alignment = PP_ALIGN.CENTER
        run = para.add_run(); run.text = h
        _set_font(run, header_size, FG2, bold=True)
    # rows
    for r_i, row in enumerate(rows, start=1):
        for c_i, val in enumerate(row):
            if isinstance(val, tuple):
                text, color = val
            else:
                text, color = val, FG
            cell = gtbl.cell(r_i, c_i)
            cell.fill.solid(); cell.fill.fore_color.rgb = BG if r_i % 2 else CARD
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.CENTER if c_i > 0 else PP_ALIGN.LEFT
            run = para.add_run(); run.text = text
            _set_font(run, cell_size, color)
    return gtbl


def section_divider(num, title, subtitle=None):
    s = add_slide()
    tb, tf = textbox(s, Inches(0), Inches(2.2), EMU_W, Inches(1.4), MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = num
    _set_font(r, 96, ACCENT, bold=True)
    tb2, tf2 = textbox(s, Inches(0), Inches(3.9), EMU_W, Inches(1.0), MSO_ANCHOR.MIDDLE)
    p = tf2.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = title
    _set_font(r, 34, FG, bold=True)
    if subtitle:
        tb3, tf3 = textbox(s, Inches(0), Inches(4.9), EMU_W, Inches(0.7), MSO_ANCHOR.MIDDLE)
        p = tf3.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = subtitle
        _set_font(r, 18, FG2)
    return s


# ═══════════════════════════════════════════════════════════
# SLIDE 1 · 封面
# ═══════════════════════════════════════════════════════════
s = add_slide()
tb, tf = textbox(s, Inches(0), Inches(2.4), EMU_W, Inches(1.2), MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "BTC Momentum"
_set_font(r, 54, FG, bold=True)
tb, tf = textbox(s, Inches(0), Inches(3.7), EMU_W, Inches(0.7), MSO_ANCHOR.MIDDLE)
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "JLST 反转-动量多信号策略系统"
_set_font(r, 22, FG2)
tb, tf = textbox(s, Inches(0), Inches(4.7), EMU_W, Inches(1.4), MSO_ANCHOR.MIDDLE)
for i, line in enumerate([
    "基于 Jegadeesh, Luo, Subrahmanyam & Titman (2025)",
    "Review of Financial Studies — 全球顶级金融学术期刊",
    "回测数据截止 2026-09-07",
]):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = line
    _set_font(r, 14, FG2, italic=(i == 1))

# ═══════════════════════════════════════════════════════════
# SLIDE 2 · 你可能遇到过这些问题
# ═══════════════════════════════════════════════════════════
s = add_slide(); add_title(s, "你可能遇到过这些问题")
tb, tf = textbox(s, Inches(0.7), Inches(1.5), Inches(11.9), Inches(5.2))
add_paragraphs(tf, [
    ("BTC 跌了 10%，朋友说\"抄底\"，你不确定——还会反弹吗？", FG, False, 18),
    ("BTC 涨了一个月，你想追——趋势还能持续吗？", FG, False, 18),
    ("看了一堆指标（RSI、MACD、布林带），信号互相矛盾", FG, False, 18),
    ("想要一个有学术论文支撑、而不是靠\"画线\"的分析工具", FG, False, 18),
    ("找到了好信号，但不知道如何叠加多重确认来提高胜率", FG, False, 18),
], first=True)
bullet_box(s, Inches(0.7), Inches(5.4), Inches(11.9), Inches(1.4),
           "BTC Momentum 的价值", ACCENT,
           ["JLST 理论信号 + MA 均线多重过滤 → 用学术框架确定方向，用趋势滤波提高精度"],
           accent=ACCENT)

# ═══════════════════════════════════════════════════════════
# SLIDE 3 · 论文出处
# ═══════════════════════════════════════════════════════════
s = add_slide(); add_title(s, "论文出处")
bullet_box(s, Inches(0.7), Inches(1.4), Inches(11.9), Inches(1.5),
           "N. Jegadeesh, J. Luo, A. Subrahmanyam, S. Titman (2025)", FG,
           ['"Short-Term Reversals and Longer-Term Momentum around the World"',
            "Review of Financial Studies (RFS) — 金融学三大顶刊之一"], accent=ACCENT)
bullet_box(s, Inches(0.7), Inches(3.1), Inches(5.75), Inches(3.4),
           "论文做了什么？", BLUE,
           ["研究了 39 个国家股票市场",
            "用同一模型解释\"短期反转\"与\"中长期动量\"如何同时存在",
            "发现噪声交易是两者联系的关键纽带"])
bullet_box(s, Inches(6.85), Inches(3.1), Inches(5.75), Inches(3.4),
           "我们做了什么？", GREEN,
           ["将论文理论应用到 BTC 市场",
            "构建实时仪表盘，每日更新",
            "叠加 MA 多信号过滤，将理论收益变为实操策略",
            "开放可配置滤波条件，用户自行调优"])

# ═══════════════════════════════════════════════════════════
# SECTION 01 · 核心概念
# ═══════════════════════════════════════════════════════════
section_divider("01", "核心概念", "短期反转 & 中长期动量")

s = add_slide(); add_title(s, "短期反转 vs 中长期动量")
bullet_box(s, Inches(0.7), Inches(1.5), Inches(5.75), Inches(2.6),
           "短期反转", ORANGE,
           ["\"涨多了会回、跌多了会弹\"",
            "短期价格含大量情绪化过度反应，几天内被纠正",
            "策略：短期逆着最近几天的极端走"])
bullet_box(s, Inches(6.85), Inches(1.5), Inches(5.75), Inches(2.6),
           "中长期动量", GREEN,
           ["\"强者恒强、弱者恒弱\"",
            "过去几个月上涨的往往继续涨，趋势有惯性",
            "策略：中长期顺着大趋势走"])
bullet_box(s, Inches(0.7), Inches(4.3), Inches(11.9), Inches(2.4),
           "如何共存？——噪声交易者是关键", PURPLE,
           ["两者作用在不同时间尺度上，并不冲突",
            "噪声交易者制造短期过度波动（反转来源），其行为在中期形成趋势（动量来源）",
            "噪声越大 → 反转机会越明显；噪声越小、趋势越干净 → 动量越可靠",
            "模型用噪声大小动态调整反转/动量权重 —— 这是它比固定参数指标聪明的地方"],
           accent=PURPLE)

s = add_slide(); add_title(s, "综合评分公式")
box = slide = s
b = s.shapes.add_shape(1, Inches(1.5), Inches(2.3), Inches(10.3), Inches(1.3))
b.fill.solid(); b.fill.fore_color.rgb = CARD; b.line.fill.background()
tf = b.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
r = p.add_run(); r.text = "Composite = w₁·z(Rev) + w₂·z(Mom) + w₃·z(FundingRate)"
_set_font(r, 22, BLUE, bold=True)
tb, tf = textbox(s, Inches(0.7), Inches(4.0), Inches(11.9), Inches(2.6))
add_paragraphs(tf, [
    ("三个分量先做标准化（z-score），再按动态权重加权求和。", FG, False, 16),
    ("权重不是固定的，而是根据当前市场噪声大小实时计算。", FG, False, 16),
    ("BTC 适配版额外引入资金费率、波动率分档、减半周期等加密市场特有因子。", FG2, False, 15),
], first=True)

# ═══════════════════════════════════════════════════════════
# SECTION 02 · 信号系统
# ═══════════════════════════════════════════════════════════
section_divider("02", "信号系统", "从评分到买卖信号")

s = add_slide(); add_title(s, "信号产生规则")
bullet_box(s, Inches(0.7), Inches(1.5), Inches(11.9), Inches(2.2),
           "评分如何变成信号", FG,
           ["综合评分向上突破阈值 → 做多信号（绿色↑）",
            "综合评分向下跌破阈值 → 做空信号（红色↓）",
            "多个条件同时触发 → 级联预警（紫色⚡，强信号）"], accent=ACCENT)
bullet_box(s, Inches(0.7), Inches(3.9), Inches(11.9), Inches(2.6),
           "JLST 原始信号前瞻回报（无 MA 滤波）", GREEN,
           ["做多信号 30 天平均 +10.71%，胜率 58%",
            "问题：部分做多信号出现在下跌趋势里（逆势抄底），容易吃套",
            "→ 这正是叠加均线滤波的动机"], accent=GREEN)

# ═══════════════════════════════════════════════════════════
# SECTION 03 · 多信号策略
# ═══════════════════════════════════════════════════════════
section_divider("03", "多信号策略", "JLST + MA 滤波")

s = add_slide(); add_title(s, "多信号滤波逻辑")
bullet_box(s, Inches(0.7), Inches(1.5), Inches(5.75), Inches(3.2),
           "做多滤波（三条件全满足）", GREEN,
           ["① 价格 > MA110（大趋势向上）",
            "② MA6 > MA103（短期走强）",
            "③ 价格 > EMA50（中期向上）",
            "只有 JLST 做多信号 + 三条件全满足 = 高质量信号"])
bullet_box(s, Inches(6.85), Inches(1.5), Inches(5.75), Inches(3.2),
           "做空滤波（价格在均线下方）", RED,
           ["价格 < MA110",
            "价格 < EMA50",
            "要求价格在均线下方才保留做空信号"])
bullet_box(s, Inches(0.7), Inches(4.9), Inches(11.9), Inches(1.7),
           "为什么 MA 滤波有效 + 用户可配置", ACCENT,
           ["均线是趋势的过滤网：筛掉逆势信号（亏损主要来源），只留顺势信号",
            "三个开关用户自主控制：全勾最严格、可只留一两个、可全取消看原始信号，实时重新统计前瞻收益"],
           accent=ACCENT)

# ═══════════════════════════════════════════════════════════
# SECTION 04 · 超额收益
# ═══════════════════════════════════════════════════════════
section_divider("04", "超额收益",
                "回测截止 2026-09-07 · 信号 200 条（做多 85 / 做空 109）")

s = add_slide(); add_title(s, "多信号 vs 原始 JLST（做多）")
add_table(s, Inches(0.9), Inches(1.5), Inches(11.5),
          ["持有期", "JLST 原始", "多信号滤波", "超额提升"],
          [
              ["1 天", "+1.12%", ("+1.74%", GREEN), ("+55%", GREEN)],
              ["3 天", "+2.01%", ("+3.50%", GREEN), ("+74%", GREEN)],
              ["7 天", "+3.25%", ("+6.03%", GREEN), ("+86%", GREEN)],
              ["30 天", "+10.71%", ("+20.52%", GREEN), ("+92%", GREEN)],
          ],
          col_widths=[Inches(2.5), Inches(3.0), Inches(3.0), Inches(3.0)])
bullet_box(s, Inches(0.9), Inches(5.0), Inches(11.5), Inches(1.6),
           "核心结论", GREEN,
           ["做多三重滤波后 30 天平均收益 +10.71% → +20.52%（超额 +9.82 个百分点，相对提升 92%）",
            "胜率 58% → 62%。过滤逆势信号是获取超额回报的关键"], accent=GREEN)

s = add_slide(); add_title(s, "做多策略 — 最强阿尔法")
bullet_box(s, Inches(0.7), Inches(1.5), Inches(3.7), Inches(1.6),
           "+20.52%", GREEN, ["多信号做多 30 日均收益"], accent=GREEN)
bullet_box(s, Inches(4.75), Inches(1.5), Inches(3.7), Inches(1.6),
           "67%", GREEN, ["3 天胜率"], accent=GREEN)
bullet_box(s, Inches(8.8), Inches(1.5), Inches(3.8), Inches(1.6),
           "85 → 42", BLUE, ["信号次数（少而精）"], accent=BLUE)
bullet_box(s, Inches(0.7), Inches(3.4), Inches(11.9), Inches(3.0),
           "解读", FG,
           ["三重滤波把做多信号从 85 次精简到 42 次 —— 宁可少做，也要做对",
            "各持有期收益全面提升：1天 +55%、3天 +74%、7天 +86%、30天 +92%",
            "胜率区间从 58-65% 抬升到 62-67%",
            "这是整套策略里最稳定、最值得依赖的部分"], accent=GREEN)

s = add_slide(); add_title(s, "做空策略 — 诚实面对局限")
add_table(s, Inches(0.9), Inches(1.5), Inches(11.5),
          ["持有期", "JLST 做空", "多信号做空", "变化"],
          [
              ["1 天", "+0.44% (53%)", ("-0.26% (44%)", FG2), "走弱"],
              ["3 天", "+0.29% (50%)", ("-1.04% (44%)", FG2), "走弱"],
              ["7 天", "+0.76% (52%)", ("-1.19% (44%)", FG2), "走弱"],
              ["30 天", ("-3.61% (42%)", RED), ("-2.87% (38%)", RED), "略改善"],
          ],
          col_widths=[Inches(2.3), Inches(3.2), Inches(3.2), Inches(2.8)])
bullet_box(s, Inches(0.9), Inches(5.0), Inches(11.5), Inches(1.7),
           "解读", ORANGE,
           ["BTC 长期上行，做空天然逆风：做空 30 天平均 -3.61%、胜率仅 42%",
            "均线下方常是超跌区，加滤波后短期反而更弱、易遇反弹",
            "做空的定位是风险预警与对冲，不宜作为主进攻策略"], accent=ORANGE)

s = add_slide(); add_title(s, "核心数据总结")
bullet_box(s, Inches(0.7), Inches(1.5), Inches(3.7), Inches(1.5),
           "+20.52%", GREEN, ["做多 30 日均收益"], accent=GREEN)
bullet_box(s, Inches(4.75), Inches(1.5), Inches(3.7), Inches(1.5),
           "+92%", ACCENT, ["30 日超额提升"], accent=ACCENT)
bullet_box(s, Inches(8.8), Inches(1.5), Inches(3.8), Inches(1.5),
           "67%", ORANGE, ["3 日胜率（做多）"], accent=ORANGE)
bullet_box(s, Inches(0.7), Inches(3.3), Inches(5.75), Inches(3.2),
           "做多策略结论", GREEN,
           ["MA 滤波后 30 日收益近乎翻倍",
            "胜率 58-65% → 62-67%",
            "信号数 85 → 42（少而精）",
            "每次信号平均 30 日赚 20.52%"])
bullet_box(s, Inches(6.85), Inches(3.3), Inches(5.75), Inches(3.2),
           "做空策略结论", RED,
           ["做空收益整体偏弱（BTC 长期上行）",
            "无滤波 30 日 -3.61%，滤波改善有限",
            "信号数 109 → 45（过滤假信号）",
            "更适合防御/对冲，不做主策略"])

# ═══════════════════════════════════════════════════════════
# SECTION 05 · 仪表盘功能
# ═══════════════════════════════════════════════════════════
section_divider("05", "仪表盘功能", "看得懂、用得上")

s = add_slide(); add_title(s, "交互式 K 线图")
bullet_box(s, Inches(0.7), Inches(1.5), Inches(5.75), Inches(2.4),
           "价格同轴（趋势）", BLUE,
           ["MA6 · EMA50 · EMA110",
            "MA103 · MA110 · MA200",
            "已实现价格"])
bullet_box(s, Inches(6.85), Inches(1.5), Inches(5.75), Inches(2.4),
           "独立轴（估值/链上/情绪）", PURPLE,
           ["成交量 · RSI · Mayer · MVRV · NUPL",
            "SMM · 卖方衰竭 · 风险回报",
            "ETF · USDT.D · BTC.D"])
bullet_box(s, Inches(0.7), Inches(4.1), Inches(11.9), Inches(2.4),
           "特色功能", GREEN,
           ["📌 信号标记：一键叠加 JLST 买卖信号（绿↑做多 / 红↓做空 / 紫⚡级联）",
            "对数坐标：长周期看 BTC 更合理",
            "坐标翻转：USDT.D / BTC.D 可翻转，直观看与价格的反向关系",
            "日线/周线切换：指标与信号标记自动适配周线聚合"], accent=GREEN)

s = add_slide(); add_title(s, "实操流程")
tb, tf = textbox(s, Inches(0.7), Inches(1.6), Inches(11.9), Inches(3.0))
add_paragraphs(tf, [
    ("① 看综合评分和信号方向", FG, True, 18),
    ("② 看三个 MA 滤波灯是否都绿", FG, True, 18),
    ("③ 全绿 = 最高质量做多信号", GREEN, True, 18),
    ("④ 结合估值指标（MVRV、NUPL、Mayer）做二次确认", FG, True, 18),
    ("⑤ 自己决定仓位", FG, True, 18),
], first=True)
bullet_box(s, Inches(0.7), Inches(5.0), Inches(11.9), Inches(1.6),
           "最佳实践", ACCENT,
           ["做多 + 三 MA 全满足 = 最高质量信号（历史 30 日 +20.52%）",
            "做空在牛市中谨慎，作减仓/对冲依据；信号约 30 天一次，属低频高质量策略"],
           accent=ACCENT)

# ═══════════════════════════════════════════════════════════
# SECTION 06 · 指标撰写规则
# ═══════════════════════════════════════════════════════════
section_divider("06", "K 线指标撰写规则", "每个指标怎么算、怎么读")

s = add_slide(); add_title(s, "趋势均线：MA 与 EMA")
add_table(s, Inches(0.6), Inches(1.4), Inches(12.1),
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
bullet_box(s, Inches(0.6), Inches(5.6), Inches(12.1), Inches(1.1),
           "SMA vs EMA", ACCENT,
           ["SMA 窗口内等权、平滑但滞后；EMA 用衰减系数 α 给近期更高权重、转向更快但更易受噪声影响。趋势策略里两者互补。"],
           accent=ACCENT)

s = add_slide(); add_title(s, "估值类：Mayer / 已实现价格 / MVRV")
add_table(s, Inches(0.6), Inches(1.4), Inches(12.1),
          ["指标", "计算规则", "解读"],
          [
              ["Mayer 倍数", "价格 ÷ MA200", ">2.4 顶部区；<1 低估区。衡量偏离长期均线程度"],
              ["已实现价格", "每枚 BTC「最后移动时价格」加权平均（全网成本）", "市场平均持仓成本线；跌破=多数人浮亏，历史大底常在此"],
              ["MVRV", "市值 ÷ 已实现市值", ">3.7 顶部风险；<1 深度低估。全网未实现盈亏比"],
          ],
          col_widths=[Inches(2.3), Inches(5.3), Inches(4.5)],
          cell_size=10.5, row_h=Inches(0.7))
bullet_box(s, Inches(0.6), Inches(5.5), Inches(12.1), Inches(1.2),
           "为什么重要", GREEN,
           ["估值类提供「贵不贵」的锚，与趋势类「涨不涨」互补——趋势告诉你方向，估值告诉你位置。"],
           accent=GREEN)

s = add_slide(); add_title(s, "链上情绪：NUPL / SMM / 卖方衰竭 / 风险回报")
add_table(s, Inches(0.6), Inches(1.4), Inches(12.1),
          ["指标", "计算规则", "解读"],
          [
              ["NUPL 净未实现盈亏", "(市值 − 已实现市值) ÷ 市值", ">0.75 极度贪婪；<0 全网浮亏（投降/大底）"],
              ["SMM", "链上花费产出的动量", "反映抛压动能变化，捕捉筹码换手节奏"],
              ["卖方衰竭", "波动率 × 长期持有者未实现亏损占比", "走高=抛售动能耗尽，常见于底部"],
              ["风险回报", "历史价格分布的下行风险 vs 上行空间比值", "越低=当前介入性价比越高"],
          ],
          col_widths=[Inches(2.6), Inches(5.0), Inches(4.5)],
          cell_size=10, row_h=Inches(0.62))
tb, tf = textbox(s, Inches(0.6), Inches(6.3), Inches(12.1), Inches(0.6))
p = tf.paragraphs[0]
r = p.add_run(); r.text = "链上数据来自 CryptoQuant，按日期对齐到 K 线；周线视图取每周最后一个值。"
_set_font(r, 11, FG2)

s = add_slide(); add_title(s, "资金流与情绪：ETF / USDT.D / BTC.D / 资金费率")
add_table(s, Inches(0.6), Inches(1.4), Inches(12.1),
          ["指标", "计算规则", "解读"],
          [
              ["ETF 净流入", "现货 BTC ETF 每日净申赎（百万美元），周线求和", "正=机构增持进场；持续净流出=需求转弱"],
              ["USDT.D", "USDT 市值 ÷ 加密总市值（%）", "升=资金避险（利空）；降=资金进场（利多）。与 BTC 常反向"],
              ["BTC.D", "BTC 市值 ÷ 加密总市值（%）", "升=资金集中 BTC；降=流向山寨"],
              ["资金费率", "永续合约多空平衡费率，每 8h 结算，取日均", "正=多头拥挤；极端正值常预示回调"],
          ],
          col_widths=[Inches(2.3), Inches(5.3), Inches(4.5)],
          cell_size=10, row_h=Inches(0.62))
bullet_box(s, Inches(0.6), Inches(5.5), Inches(12.1), Inches(1.2),
           "资金费率的特殊地位", ORANGE,
           ["它是 JLST 动态权重中噪声代理的组成部分（占 30% 权重）——费率越极端，模型越调低动量权重、调高反转权重。"],
           accent=ORANGE)

s = add_slide(); add_title(s, "什么情况下最适合用多策略？")
bullet_box(s, Inches(0.55), Inches(1.5), Inches(3.95), Inches(4.0),
           "✅ 最适合", GREEN,
           ["明确趋势市：价格站上 MA110/EMA50，做多胜率 62-67%",
            "中长线持有（1-4 周）：30 日收益差距最大",
            "右侧交易者：愿放弃最低点换确定性"])
bullet_box(s, Inches(4.7), Inches(1.5), Inches(3.95), Inches(4.0),
           "⚠️ 谨慎", ORANGE,
           ["震荡/横盘市：均线频繁穿越，滤波反复触发又失效",
            "做空：BTC 长期上行，整体为负，仅作预警",
            "短持有期（1-3 日）：超额小，成本占比高"])
bullet_box(s, Inches(8.85), Inches(1.5), Inches(3.95), Inches(4.0),
           "❌ 不适合", RED,
           ["日内/高频：信号约 30 天一次，频率不匹配",
            "黑天鹅急跌：任何趋势滤波都滞后",
            "当自动交易机器人：它是决策辅助，非全自动"])
bullet_box(s, Inches(0.55), Inches(5.7), Inches(12.25), Inches(1.1),
           "一句话", ACCENT,
           ["多策略的超额收益来自「在趋势里做多、并用均线把逆势信号筛掉」。趋势越明确、持有期越长，价值越大；越震荡、越短线，价值越小。"],
           accent=ACCENT)

# ═══════════════════════════════════════════════════════════
# SECTION 07 · 局限与总结
# ═══════════════════════════════════════════════════════════
section_divider("07", "局限与总结", "诚实面对模型的边界")

s = add_slide(); add_title(s, "需要注意的事项")
tb, tf = textbox(s, Inches(0.7), Inches(1.5), Inches(11.9), Inches(4.2))
add_paragraphs(tf, [
    ("历史回测 ≠ 未来表现 —— 市场结构会变化", FG, False, 16),
    ("论文参数来自股票市场 —— 虽已做 BTC 适配，仍是近似而非精确", FG, False, 16),
    ("低频策略 —— 约 30 天触发一次，不适合日内或高频", FG, False, 16),
    ("做空信号效果一般 —— BTC 长期向上，做空天然逆风、胜率偏低", FG, False, 16),
    ("MA 参数有优化空间 —— 当前 MA110/MA103/EMA50 基于经验，用户可自调", FG, False, 16),
    ("极端行情下信号会滞后 —— 黑天鹅事件中任何趋势指标都滞后", FG, False, 16),
], first=True)
bullet_box(s, Inches(0.7), Inches(5.7), Inches(11.9), Inches(1.1),
           "正确定位", ORANGE,
           ["BTC Momentum 是辅助决策工具，帮你过滤噪声、提供学术框架支撑的方向判断，但最终决策权在你——不是自动赚钱机器。"],
           accent=ORANGE)

s = add_slide(); add_title(s, "总结")
bullet_box(s, Inches(0.7), Inches(1.5), Inches(5.75), Inches(2.4),
           "🔬 理论基础", ACCENT,
           ["JLST (2025) RFS 论文 — 39 国验证",
            "短期反转 + 中长期动量双因子",
            "噪声交易者是关键中介变量"])
bullet_box(s, Inches(6.85), Inches(1.5), Inches(5.75), Inches(2.4),
           "🎯 多信号策略", GREEN,
           ["JLST 信号 + MA110/MA6>MA103/EMA50",
            "做多 30 日 +20.52%（胜率 62%）",
            "超额比原始信号提升 92%"])
bullet_box(s, Inches(0.7), Inches(4.1), Inches(11.9), Inches(2.4),
           "🖥️ 工具特色", PURPLE,
           ["每日自动更新数据 · 近 20 个技术/链上指标 + 信号标记",
            "5 个可配置 MA 滤波条件 · 日线/周线 · 暗/亮主题 · 三图时间轴联动",
            "把严谨的学术框架，变成你每天都能打开、看得懂、用得上的东西"], accent=PURPLE)

# ── 保存 ──
out = Path(__file__).resolve().parent.parent / "BTC_Momentum_演示.pptx"
prs.save(str(out))
print(f"OK 已生成 {out}  （共 {len(prs.slides.__iter__.__self__._sldIdLst)} 页）")
