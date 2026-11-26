"""生成 KVCM「创新先锋」idea 单页 PPT（16:9）。

用法：python3 kvcm/report/gen_innovation_idea.py [输出路径]
依赖：python-pptx。数字均取自 KVCM 仿真结论（kvcm/sim/results），【注入】为人为设定场景。
配色、字体、idea 名称与同心圆结构沿用 KVCM-汇报-含PD混部-20261124.pptx。
"""

import math
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

RED = RGBColor(0xC8, 0x10, 0x2E)
NAVY = RGBColor(0x0F, 0x25, 0x4A)
INK = RGBColor(0x1B, 0x24, 0x32)
GRAY = RGBColor(0x74, 0x82, 0x96)
SLATE = RGBColor(0x47, 0x55, 0x69)
LINE = RGBColor(0xD8, 0xE1, 0xEC)
BG = RGBColor(0xF8, 0xFA, 0xFC)
PINK = RGBColor(0xFD, 0xF2, 0xF4)
ROSE = RGBColor(0xF5, 0xB8, 0xC2)
TEAL = RGBColor(0x0D, 0x76, 0x80)
TEAL_BG = RGBColor(0xE6, 0xF3, 0xF4)
GREEN = RGBColor(0x15, 0x80, 0x3D)
GREEN_BG = RGBColor(0xEA, 0xF6, 0xEE)
ORANGE = RGBColor(0xB4, 0x53, 0x09)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Noto Sans CJK SC"

OUT = Path(__file__).with_name("KVCM-创新先锋-idea-20261126.pptx")


def set_font(run, size, color=INK, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = FONT
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set("typeface", FONT)


def shadow(shape, blur=60000, dist=22000, alpha=20000):
    sppr = shape._element.spPr
    eff = sppr.find(qn("a:effectLst"))
    if eff is None:
        eff = sppr.makeelement(qn("a:effectLst"), {})
        sppr.append(eff)
    sh = eff.makeelement(qn("a:outerShdw"), {
        "blurRad": str(blur), "dist": str(dist), "dir": "5400000", "algn": "t",
        "rotWithShape": "0"})
    clr = sh.makeelement(qn("a:srgbClr"), {"val": "0F254A"})
    clr.append(clr.makeelement(qn("a:alpha"), {"val": str(alpha)}))
    sh.append(clr)
    eff.append(sh)


def shape(slide, x, y, w, h, fill=None, line=None, kind=MSO_SHAPE.RECTANGLE, lw=0.75,
          dash=None, grad=None, shade=False, radius=0.06):
    s = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if grad:
        s.fill.gradient()
        s.fill.gradient_angle = grad[2] if len(grad) > 2 else 0
        s.fill.gradient_stops[0].color.rgb = grad[0]
        s.fill.gradient_stops[1].color.rgb = grad[1]
    elif fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(lw)
        if dash:
            s.line.dash_style = dash
    s.shadow.inherit = False
    if shade:
        shadow(s)
    if kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    return s


def text(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT,
         margin=0.03, spacing=1.08, space_after=2, target=None):
    """paras: 段落列表；每段是 [(文本, 字号, 颜色, 粗体), ...]。"""
    if target is None:
        target = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = target.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(margin))
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_after = Pt(space_after)
        for t, size, color, bold in runs:
            r = p.add_run()
            r.text = t
            set_font(r, size, color, bold)
    return target


def label(slide, x, y, w, h, paras, fill=None, line=None, kind=MSO_SHAPE.ROUNDED_RECTANGLE,
          grad=None, shade=False, radius=0.2, **kw):
    s = shape(slide, x, y, w, h, fill, line, kind, grad=grad, shade=shade, radius=radius)
    kw.setdefault("anchor", MSO_ANCHOR.MIDDLE)
    kw.setdefault("align", PP_ALIGN.CENTER)
    kw.setdefault("space_after", 0)
    kw.setdefault("spacing", 1.0)
    return text(slide, 0, 0, 0, 0, paras, target=s, **kw)


def circle(slide, cx, cy, d, **kw):
    return shape(slide, cx - d / 2, cy - d / 2, d, d, kind=MSO_SHAPE.OVAL, **kw)


def badge(slide, cx, cy, d, num, color, size=10):
    return label(slide, cx - d / 2, cy - d / 2, d, d, [[(num, size, WHITE, True)]], color,
                 WHITE, MSO_SHAPE.OVAL, shade=True, margin=0)


def line(slide, x1, y1, x2, y2, color, width=1.0, dash=None):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1),
                                   Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    if dash:
        c.line.dash_style = dash
    return c


def card(slide, x, y, w, h, no, title, accent, light):
    shape(slide, x, y, w, h, WHITE, LINE, MSO_SHAPE.ROUNDED_RECTANGLE, shade=True, radius=0.04)
    shape(slide, x + 0.02, y + 0.02, w - 0.04, 0.07, grad=(accent, light, 0))
    badge(slide, x + 0.3, y + 0.33, 0.36, no, accent, 9.5)
    text(slide, x + 0.55, y + 0.15, w - 0.7, 0.36, [[(title, 13, NAVY, True)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02)


def bullets(slide, x, y, w, h, items, accent, size=9.5):
    paras = []
    for head, body in items:
        paras.append([("▍", size, accent, True), (head, size, NAVY, True),
                      ("  " + body, size, SLATE, False)])
    return text(slide, x, y, w, h, paras, spacing=1.05, space_after=3.5)


def ring_diagram(s, px, py, pw, ph):
    cx, cy = px + pw / 2, py + 2.98
    r_dram, r_hbm, r_ctl, r_kv = 1.8, 1.35, 0.95, 0.7

    label(s, cx - 1.6, py + 0.5, 3.2, 0.34,
          [[("以 KV 为中心的统一资源调度中枢", 10.5, WHITE, True)]],
          grad=(RED, RGBColor(0x8E, 0x0B, 0x21), 0), shade=True, radius=0.5)

    circle(s, cx, cy, 2 * r_dram + 0.22, fill=None, line=ROSE, lw=0.75, dash=MSO_LINE.DASH)
    circle(s, cx, cy, 2 * r_dram, grad=(GREEN_BG, RGBColor(0xCF, 0xEB, 0xD8), 90), line=GREEN,
           lw=1.25, shade=True)
    circle(s, cx, cy, 2 * r_hbm, grad=(TEAL_BG, RGBColor(0xC9, 0xE4, 0xE7), 90), line=TEAL, lw=1.25)
    circle(s, cx, cy, 2 * r_ctl, fill=WHITE, line=RED, lw=1.75, dash=MSO_LINE.DASH)
    circle(s, cx, cy, 2 * r_kv, grad=(RGBColor(0x1E, 0x3A, 0x6B), NAVY, 90), shade=True)

    text(s, cx - r_kv, cy - 0.52, 2 * r_kv, 1.04,
         [[("KV", 22, WHITE, True)], [("集群第一公民", 10, RGBColor(0xFA, 0xCC, 0x15), True)],
          [("可寻址 · 可保活", 7.5, LINE, False)]],
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, space_after=0, spacing=1.0, margin=0)
    band = [(r_ctl, r_kv, "KVCM 控制面", "R0–R2", RED, 8),
            (r_hbm, r_ctl, "HBM · 引擎 LRU", "P / D / P⇄D 实例", TEAL, 8.5),
            (r_dram, r_hbm, "DRAM 共享池 · KVCM 管理", "跨实例命中 · 保活写回", GREEN, 9)]
    for ro, ri, top_t, bot_t, color, size in band:
        hgt = ro - ri
        w = 2 * math.sqrt(ro ** 2 - (ri + hgt * 0.35) ** 2) * 0.92
        text(s, cx - w / 2, cy - ro + 0.01, w, hgt - 0.02, [[(top_t, size, color, True)]],
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0, space_after=0)
        text(s, cx - w / 2, cy + ri + 0.01, w, hgt - 0.02, [[(bot_t, size - 0.5, color, False)]],
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, margin=0, space_after=0)

    cw, ch = 1.92, 1.15
    lx, rx = px + 0.15, px + pw - 0.15 - cw
    ty, by = cy - 1.0 - ch, cy + 1.0
    cards = [
        (lx, ty, "1", "会话感知保留", "R1 会话与 KV 环", RED, PINK,
         "会话 hook 标记停顿，停顿期 pin / 保活；按“未来复用”驱逐，告别 LRU 错杀", 225),
        (rx, ty, "2", "PD 动态转换", "R2 弹性与容量环", ORANGE, RGBColor(0xFF, 0xF7, 0xED),
         "按 P/D 配比在线切换实例角色，KV 写回池里保活，切换不丢、不重算", 315),
        (lx, by, "3", "局部性感知路由", "R0 请求选路环", TEAL, TEAL_BG,
         "选点联合计入 KV 数据位置 + 网卡出口队列 + 命中，网卡紧张不失稳", 135),
        (rx, by, "4", "旁挂零 Patch", "池简单 · 策略复杂", GREEN, GREEN_BG,
         "不改引擎内核，KVCM 故障时回退原生路径；池用哈希条带 + 单副本", 45),
    ]
    for x, y, no, title, ring, accent, light, body, ang in cards:
        a = math.radians(ang)
        r_b = r_ctl if no != "4" else r_dram
        bx, by_ = cx + r_b * math.cos(a), cy - r_b * math.sin(a) * -1
        ex = x + cw if x == lx else x
        ey = y + ch / 2
        line(s, bx, by_, ex, ey, accent, 1.0, MSO_LINE.ROUND_DOT)
        circle(s, ex, ey, 0.09, fill=accent)
        shape(s, x, y, cw, ch, WHITE, LINE, MSO_SHAPE.ROUNDED_RECTANGLE, shade=True, radius=0.06)
        shape(s, x, y + 0.08, 0.06, ch - 0.16, fill=accent)
        text(s, x + 0.12, y + 0.06, cw - 0.18, ch - 0.1,
             [[(title, 10.5, accent, True)], [(ring, 8, GRAY, True)], [(body, 8.5, INK, False)]],
             spacing=1.03, space_after=1.5, margin=0.02)
        badge(s, bx, by_, 0.3, no, accent, 9)

    fy = py + ph - 0.98
    steps = [("意图下行", RED), ("视图归一", NAVY), ("原语执行", TEAL), ("事实闭环", GREEN)]
    sw = (pw - 0.3 + 0.12 * 3) / 4
    for i, (t, c) in enumerate(steps):
        kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        label(s, px + 0.15 + i * (sw - 0.12), fy, sw, 0.38, [[(t, 10, WHITE, True)]], c, kind=kind,
              radius=None, margin=0)
    text(s, px + 0.15, fy + 0.45, pw - 0.3, 0.42,
         [[("三层打通：", 9, NAVY, True),
           ("应用意图层 × KVCM 控制调度层 × 引擎存储网络底座；HBM（引擎 LRU）+ DRAM 池（KVCM）两层，"
            "复杂度投在保留策略和路由", 9, SLATE, False)]],
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spacing=1.05, space_after=0)


def kpi(s, x, y, w, h, big, unit, caption, accent, light):
    shape(s, x, y, w, h, grad=(light, WHITE, 90), line=LINE, kind=MSO_SHAPE.ROUNDED_RECTANGLE,
          shade=True, radius=0.08)
    shape(s, x, y + 0.1, 0.05, h - 0.2, fill=accent)
    text(s, x + 0.08, y + 0.05, w - 0.12, 0.42, [[(big, 15, accent, True), (unit, 8, accent, True)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02, space_after=0)
    text(s, x + 0.08, y + 0.47, w - 0.12, h - 0.5, [[(caption, 7.5, INK, False)]],
         spacing=1.02, space_after=0, margin=0.02)


def build(out: Path):
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    shape(s, 0, 0, 13.333, 7.5, grad=(BG, RGBColor(0xEE, 0xF2, 0xF7), 90))
    shape(s, 0, 0, 13.333, 0.07, grad=(RED, NAVY, 0))

    label(s, 0.35, 0.22, 1.5, 0.36, [[("创新先锋 · idea", 10.5, WHITE, True)]],
          grad=(RED, RGBColor(0x8E, 0x0B, 0x21), 0), shade=True, radius=0.5)
    text(s, 1.98, 0.13, 11.0, 0.52,
         [[("KVCM", 22, RED, True), ("  以 KV 为中心的统一资源调度中枢", 22, NAVY, True)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    text(s, 0.35, 0.66, 12.6, 0.28,
         [[("将 KV 升级为推理集群「第一公民」：", 10.5, NAVY, True),
           ("HBM + DRAM 共享池 · 会话感知保留 · PD 动态转换 · 局部性感知路由 · 旁挂零 Patch 解耦",
            10.5, SLATE, False)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02)

    # 方案介绍（左侧大块）
    px, py, pw, ph = 0.35, 1.05, 7.55, 6.2
    card(s, px, py, pw, ph, "02", "方案介绍", TEAL, TEAL_BG)
    ring_diagram(s, px, py, pw, ph)

    # 右侧三块
    rx, rw = 8.07, 4.91
    y1, h1 = 1.05, 1.9
    card(s, rx, y1, rw, h1, "01", "问题背景", RED, ROSE)
    bullets(s, rx + 0.14, y1 + 0.55, rw - 0.24, h1 - 0.6, [
        ("停顿即错杀", "Agent 工具调用停顿期，引擎 LRU 驱逐热前缀，回来整段重算，白耗 prefill"),
        ("HBM 小且私有", "每实例 KV HBM 约 296 GB（DeepSeek V4.1-Flash，昇腾 A3），无法跨实例复用"),
        ("路由与卸载割裂", "选点不看数据位置与网卡出口，网卡紧张时 KV 搬运排队失稳"),
        ("PD 配比静态", "配比随时间变化，固定 P:D 跟不上；切换时 KV 随实例丢失"),
    ], RED)

    y2, h2 = 3.07, 1.88
    card(s, rx, y2, rw, h2, "03", "创新点", NAVY, RGBColor(0x9F, 0xB3, 0xD1))
    bullets(s, rx + 0.14, y2 + 0.55, rw - 0.24, h2 - 0.6, [
        ("KV 第一公民", "KV 从实例私有缓存升为 HBM + DRAM 集群共享、可编排资源"),
        ("会话感知保留", "用“未来复用”hook 决定保留 / 驱逐，替代 LRU"),
        ("PD 动态转换 + 池保活", "P/D 角色在线切换，KV 留在池里不重算"),
        ("局部性感知路由", "数据位置 + 网卡队列 + 命中联合选路"),
        ("池简单、策略复杂", "哈希条带 + 单副本即可避免热点"),
    ], NAVY)

    y3, h3 = 5.07, 2.18
    card(s, rx, y3, rw, h3, "04", "商业价值", GREEN, RGBColor(0x86, 0xEF, 0xAC))
    kw, kh, kg = 1.55, 1.1, 0.07
    kx = rx + (rw - 3 * kw - 2 * kg) / 2
    ky = y3 + 0.56
    kpi(s, kx, ky, kw, kh, "251.9→210.4", "", "prefill NPU 省卡｜CC 3.5 TB 池 · 理想 hook，"
        "TTFT P95 2.38→1.64 s", RED, PINK)
    kpi(s, kx + kw + kg, ky, kw, kh, "+18.6%", "  goodput", "【注入】PD 动态转换 S3"
        "（滞后控制器 + 池 +13.5%）", ORANGE, RGBColor(0xFF, 0xF7, 0xED))
    kpi(s, kx + 2 * (kw + kg), ky, kw, kh, "唯一稳定", "", "【注入】网卡紧张（出口负载 158–758%）"
        "仅 KVCM + 局部性稳住，P95 1.56 / 0.57 s", TEAL, TEAL_BG)
    text(s, rx + 0.14, ky + kh + 0.05, rw - 0.24, 0.5,
         [[("现实 hook 仅 250.9：hook 质量决定收益；旁挂零 Patch，引擎与存储底座两侧可替换", 8.5, NAVY,
            False)],
          [("口径：trace 回放仿真（DeepSeek V4.1-Flash，昇腾 A3 14 节点 × 16 卡），无真机校准；"
            "【注入】为人为设定场景；P95 为 CC / Qwen", 7, GRAY, False)]],
         spacing=1.0, space_after=1, margin=0.02)

    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return out


if __name__ == "__main__":
    print(build(Path(sys.argv[1]) if len(sys.argv) > 1 else OUT))
