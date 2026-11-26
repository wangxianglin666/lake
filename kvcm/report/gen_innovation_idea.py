"""生成 KVCM「创新先锋」idea 单页 PPT（16:9）。

用法：python3 kvcm/report/gen_innovation_idea.py [输出路径]
依赖：python-pptx。数字均取自 KVCM 仿真结论（kvcm/sim/results），【注入】为人为设定场景。
配色、字体、架构分层与原语名称沿用 KVCM-汇报-含PD混部-20261124.pptx。
"""

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


def arrow(slide, x1, y1, x2, y2, color, width=1.25, both=False):
    c = line(slide, x1, y1, x2, y2, color, width)
    ln = c.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if both:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return c


def chip(s, x, y, w, h, t, color, fill=WHITE, size=8, bold=False, tc=None):
    return label(s, x, y, w, h, [[(t, size, tc or color, bold)]], fill, color, radius=0.18, margin=0.02)


def layer(s, x, y, w, h, title, color, fill):
    shape(s, x, y, w, h, grad=(fill, WHITE, 90), line=color, kind=MSO_SHAPE.ROUNDED_RECTANGLE,
          radius=0.04, shade=True)
    text(s, x + 0.08, y + 0.02, w - 0.16, 0.24, [[(title, 9, color, True)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02)


def tag(s, x, y, w, t, color, up=False):
    y1, y2 = (y + 0.24, y) if up else (y, y + 0.24)
    arrow(s, x, y1, x, y2, color, 1.5)
    text(s, x + 0.05, y, 1.4, 0.24, [[(t, 7.5, color, True)]], anchor=MSO_ANCHOR.MIDDLE, margin=0)


def arch_diagram(s, px, py, pw, ph):
    x0, W = px + 0.15, pw - 0.3
    text(s, px + 1.75, py + 0.1, pw - 1.9, 0.36,
         [[("三层打通：", 9, NAVY, True), ("意图下行 → 视图归一 → 原语执行 → 事实闭环", 9, RED, True)]],
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT, margin=0.02)

    # 上层
    uy, uh = py + 0.5, 0.64
    layer(s, x0, uy, W, uh, "上层｜应用 · Agent · 网关（Hook 意图注入）", ORANGE, RGBColor(0xFF, 0xF7, 0xED))
    ups = ["多轮会话上下文与前缀标识", "Agent 工具等待停顿与下一轮 ETA", "业务优先级 / SLA 与保留预算",
           "Gateway 准入与流量过载控制"]
    cw = (W - 0.16 - 3 * 0.08) / 4
    for i, t in enumerate(ups):
        chip(s, x0 + 0.08 + i * (cw + 0.08), uy + 0.28, cw, 0.28, t, ORANGE, tc=INK)
    tag(s, x0 + W / 2, uy + uh, W, "请求 + Hook 意图下行", ORANGE)

    # 中层
    my, mh = uy + uh + 0.24, 2.02
    layer(s, x0, my, W, mh, "中层｜KVCM 联合控制中枢（管理面决策 × 控制面执行解耦）", RED, PINK)
    iy, ih = my + 0.3, mh - 0.38
    lw_ = 3.0
    lx = x0 + 0.08
    shape(s, lx, iy, lw_, ih, WHITE, NAVY, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    text(s, lx + 0.06, iy + 0.02, lw_ - 0.12, 0.22,
         [[("管理面｜R0–R2 三环", 8.5, NAVY, True), ("  只管全局收益", 7.5, GRAY, False)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    rings = [("R0 毫秒选路", "结合 KVT 命中、队列负载与网络成本，逐请求决定去哪个实例", NAVY),
             ("R1 秒级预热", "响应 Agent Hook 意图与热点变化，驱动会话保活、热点复制与升降层", RED),
             ("R2 分钟扩缩", "统筹计算池与 KV 池容量，带热缓存扩容、无损排空缩容", TEAL)]
    ry = iy + 0.27
    rh = 0.33
    for t, d, c in rings:
        label(s, lx + 0.06, ry, 0.82, rh, [[(t, 8, WHITE, True)]], c, radius=0.15, margin=0.01)
        text(s, lx + 0.92, ry - 0.01, lw_ - 0.97, rh + 0.02, [[(d, 7.5, INK, False)]],
             anchor=MSO_ANCHOR.MIDDLE, spacing=1.0, space_after=0, margin=0.01)
        ry += rh + 0.05
    chip(s, lx + 0.06, ry + 0.01, lw_ - 0.12, 0.25, "KVT 统一视图：块 → 实例 × 介质 × 状态", NAVY,
         fill=RGBColor(0xEE, 0xF2, 0xF7), size=8, bold=True)

    gx = lx + lw_ + 0.04
    gw = 0.5
    arrow(s, gx + 0.02, iy + ih * 0.38, gx + gw - 0.02, iy + ih * 0.38, RED, 1.5)
    text(s, gx - 0.02, iy + ih * 0.38 - 0.22, gw + 0.04, 0.2, [[("Intent", 7.5, RED, True)]],
         align=PP_ALIGN.CENTER, margin=0)
    arrow(s, gx + gw - 0.02, iy + ih * 0.62, gx + 0.02, iy + ih * 0.62, TEAL, 1.5)
    text(s, gx - 0.02, iy + ih * 0.62 + 0.02, gw + 0.04, 0.2, [[("READY", 7.5, TEAL, True)]],
         align=PP_ALIGN.CENTER, margin=0)

    cx = gx + gw + 0.04
    cw2 = x0 + W - 0.08 - cx
    shape(s, cx, iy, cw2, ih, WHITE, TEAL, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
    text(s, cx + 0.06, iy + 0.02, cw2 - 0.12, 0.22,
         [[("控制面｜8 元组件", 8.5, TEAL, True), ("  只管把 KV 管住：跨介质操作", 7.5, GRAY, False)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    groups = [("搬迁与复制", ("copy", "move"), NAVY), ("保活与锁留", ("keepalive", "pin"), ORANGE),
              ("冷热升降层", ("promote", "demote"), TEAL), ("存续与回收", ("retain", "remove"), GREEN)]
    gwid = (cw2 - 0.12 - 3 * 0.06) / 4
    for i, (g, prims, c) in enumerate(groups):
        bx = cx + 0.06 + i * (gwid + 0.06)
        text(s, bx, iy + 0.26, gwid, 0.2, [[(g, 7.5, c, True)]], align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, margin=0)
        for j, p in enumerate(prims):
            label(s, bx, iy + 0.48 + j * 0.4, gwid, 0.35, [[(p, 8.5, c, True)]],
                  RGBColor(0xF8, 0xFA, 0xFC), c, radius=0.1, margin=0.01)
    chip(s, cx + 0.06, iy + ih - 0.33, cw2 - 0.12, 0.25,
         "核验与台账：确定性 READY 核验 · 对账 · 生命周期台账", TEAL, fill=TEAL_BG, size=8, bold=True)

    # 中下层之间
    ty = my + mh
    q = W / 4
    tag(s, x0 + q * 0.5 - 0.4, ty, W, "选路分发", NAVY)
    tag(s, x0 + q * 1.5 - 0.4, ty, W, "KV 事件上行", TEAL, up=True)
    tag(s, x0 + q * 2.5 - 0.4, ty, W, "8 原语指令", RED)
    tag(s, x0 + q * 3.5 - 0.4, ty, W, "副本事实 / 淘汰", TEAL, up=True)

    # 下层
    by, bh = ty + 0.24, 1.24
    layer(s, x0, by, W, bh, "下层｜异构推理引擎 × 分布式 KV 存储池 × 三平面高速网络（数据面零拷贝直达）",
          NAVY, RGBColor(0xEE, 0xF2, 0xF7))
    ry = by + 0.29
    ew = 3.3
    shape(s, x0 + 0.08, ry, ew, 0.5, WHITE, NAVY, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, x0 + 0.12, ry + 0.01, ew - 0.08, 0.2, [[("推理引擎集群（vLLM / SGLang / 昇腾引擎）", 8, NAVY, True)]],
         margin=0.01)
    chip(s, x0 + 0.14, ry + 0.23, 1.5, 0.22, "Prefill / Decode / 混部", NAVY, fill=BG, size=7.5)
    chip(s, x0 + 1.72, ry + 0.23, 1.6, 0.22, "HBM 本地工作集", GREEN, fill=GREEN_BG, size=7.5, bold=True)
    ax = x0 + 0.08 + ew
    arrow(s, ax + 0.05, ry + 0.3, ax + 0.55, ry + 0.3, GREEN, 1.5, both=True)
    text(s, ax, ry + 0.02, 0.6, 0.2, [[("旁路直传", 7, GREEN, True)]], align=PP_ALIGN.CENTER, margin=0)
    sx = ax + 0.6
    sw = x0 + W - 0.08 - sx
    shape(s, sx, ry, sw, 0.5, WHITE, GREEN, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    text(s, sx + 0.04, ry + 0.01, sw - 0.08, 0.2, [[("分布式 KV 存储池（全局副本权威）", 8, GREEN, True)]],
         margin=0.01)
    chip(s, sx + 0.06, ry + 0.23, (sw - 0.18) / 2, 0.22, "本机 DRAM", GREEN, fill=GREEN_BG, size=7.5)
    chip(s, sx + 0.12 + (sw - 0.18) / 2, ry + 0.23, (sw - 0.18) / 2, 0.22, "远端 DRAM", GREEN,
         fill=GREEN_BG, size=7.5)
    ny = ry + 0.56
    shape(s, x0 + 0.08, ny, W - 0.16, 0.32, grad=(NAVY, RGBColor(0x1E, 0x3A, 0x6B), 0),
          kind=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.15)
    text(s, x0 + 0.14, ny, 1.0, 0.32, [[("三平面网络", 8, WHITE, True)]], anchor=MSO_ANCHOR.MIDDLE,
         margin=0)
    nets = [("H2D 专用 UB（本机首选）", RGBColor(0x93, 0xC5, 0xFD)),
            ("存储面 RoCE（跨机 / 升降层）", RGBColor(0x86, 0xEF, 0xAC)),
            ("计算面 HCCS（算子隔离）", RGBColor(0xFD, 0xBA, 0x74))]
    nw = (W - 0.16 - 0.95 - 3 * 0.08) / 3
    for i, (t, c) in enumerate(nets):
        chip(s, x0 + 1.0 + i * (nw + 0.08), ny + 0.05, nw, 0.22, t, c, fill=NAVY, size=7.5, tc=WHITE)

    # 生命周期带
    ly = by + bh + 0.08
    label(s, x0, ly, 1.05, 0.36, [[("KV 全生命周期", 8.5, WHITE, True)]], RED, radius=0.2, margin=0.01)
    steps = [("预热", RED), ("命中保活", ORANGE), ("均衡迁徙", NAVY), ("冷热升降", TEAL), ("存续回收", GREEN)]
    stw = (W - 1.12 + 0.08 * 4) / 5
    for i, (t, c) in enumerate(steps):
        kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        label(s, x0 + 1.12 + i * (stw - 0.08), ly, stw, 0.36, [[(t, 8.5, WHITE, True)]], c, kind=kind,
              radius=None, margin=0)

    # 底注
    fy = ly + 0.44
    label(s, x0, fy, W, py + ph - 0.1 - fy,
          [[("旁挂零 Patch：", 8.5, RGBColor(0xFA, 0xCC, 0x15), True),
            ("不改 vLLM / SGLang 调度内核；KVCM 异常时推理引擎自动回退原生路径；"
             "引擎 / 存储双向可替", 9, WHITE, False)],
           [("KV 第一公民：", 8.5, RGBColor(0xFA, 0xCC, 0x15), True),
            ("KV 从「算完即弃的副产物」升级为「可编排的算力资产」，统筹算力、显存与网络", 9, WHITE, False)]],
          grad=(NAVY, RGBColor(0x1E, 0x3A, 0x6B), 0), radius=0.12, shade=True, spacing=1.05,
          align=PP_ALIGN.LEFT, margin=0.1)



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
    arch_diagram(s, px, py, pw, ph)

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
