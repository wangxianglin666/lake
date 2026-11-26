"""生成 KVCM「创新先锋」idea 单页 PPT（16:9）。

用法：python3 kvcm/report/gen_innovation_idea.py [输出路径]
依赖：python-pptx。数字均取自 KVCM 仿真结论（kvcm/sim/results），【注入】为人为设定场景。
配色与字体沿用 KVCM-汇报-含PD混部-20261124.pptx。
"""

import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

RED = RGBColor(0xC8, 0x10, 0x2E)
NAVY = RGBColor(0x0F, 0x25, 0x4A)
INK = RGBColor(0x1B, 0x24, 0x32)
GRAY = RGBColor(0x74, 0x82, 0x96)
LINE = RGBColor(0xD8, 0xE1, 0xEC)
BG = RGBColor(0xF8, 0xFA, 0xFC)
PINK = RGBColor(0xFD, 0xF2, 0xF4)
TEAL = RGBColor(0x0D, 0x76, 0x80)
GREEN = RGBColor(0x15, 0x80, 0x3D)
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


def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, lw=0.75):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
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
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    return s


def text(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT,
         margin=0.04, spacing=1.12, space_after=2, shape=None):
    """paras: list of paragraphs; 每段是 [(文本, 字号, 颜色, 粗体), ...]。"""
    if shape is None:
        shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
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
            set_font(p.add_run(), size, color, bold)
            p.runs[-1].text = t
    return shape


def box_text(slide, x, y, w, h, paras, fill, line=None, **kw):
    s = rect(slide, x, y, w, h, fill, line, MSO_SHAPE.ROUNDED_RECTANGLE)
    kw.setdefault("anchor", MSO_ANCHOR.MIDDLE)
    kw.setdefault("align", PP_ALIGN.CENTER)
    kw.setdefault("margin", 0.05)
    kw.setdefault("space_after", 0)
    return text(slide, 0, 0, 0, 0, paras, shape=s, **kw)


def arrow(slide, x1, y1, x2, y2, color=GRAY, both=False, width=1.25):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1),
                                   Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    ln = c.line._get_or_add_ln()
    ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if both:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    return c


def panel(slide, x, y, w, h, no, title, accent):
    rect(slide, x, y, w, h, WHITE, LINE)
    rect(slide, x, y, 0.06, h, accent)
    text(slide, x + 0.14, y + 0.06, w - 0.2, 0.32,
         [[(f"{no}  ", 13, accent, True), (title, 13, NAVY, True)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    rect(slide, x + 0.16, y + 0.4, w - 0.32, 0.012, LINE)


def bullets(slide, x, y, w, h, items, size=10.5, accent=RED):
    paras = []
    for head, body in items:
        runs = [("■ ", size - 3, accent, False), (head, size, INK, True)]
        if body:
            runs.append(("：" + body, size, INK, False))
        paras.append(runs)
    return text(slide, x, y, w, h, paras, spacing=1.1, space_after=4)


def build(out: Path):
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, BG)
    rect(s, 0, 0, 13.333, 0.06, RED)

    # 标题栏
    box_text(s, 0.4, 0.22, 1.45, 0.36, [[("创新先锋 · idea", 10.5, WHITE, True)]], RED)
    text(s, 1.98, 0.14, 11.0, 0.52,
         [[("KVCM：HBM + DRAM 共享池上的会话感知 KV 调度", 22, NAVY, True)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02)
    text(s, 0.4, 0.66, 12.5, 0.3,
         [[("把 KV 从“实例私有缓存”升为集群级第一公民：会话感知保留 + PD 动态转换 + 局部性感知路由，"
            "旁挂零 Patch，不改引擎内核", 11, GRAY, False)]],
         anchor=MSO_ANCHOR.MIDDLE, margin=0.02)

    top, top_h = 1.05, 3.1
    bot, bot_h = 4.27, 2.83

    # ① 问题背景
    x, w = 0.4, 4.15
    panel(s, x, top, w, top_h, "01", "问题背景", RED)
    bullets(s, x + 0.16, top + 0.5, w - 0.3, top_h - 0.55, [
        ("Agent 多轮会话停顿长",
         "工具调用造成会话停顿，引擎本地 LRU 在停顿期把热前缀驱逐，回来后整段重算，白耗 prefill 算力、拉高 TTFT"),
        ("HBM 小且实例私有",
         "每实例可用 KV HBM 约 296 GB（DeepSeek V4.1-Flash，昇腾 A3），热前缀装不下，也不能跨实例复用"),
        ("路由与卸载割裂",
         "选点只看负载 / 命中，不看数据位置与网卡出口；网卡紧张时 KV 搬运排队、尾延迟失稳"),
        ("PD 配比静态",
         "P/D 负载配比随时间变化，固定 P:D 跟不上；角色切换时 KV 随实例丢失，只能重算"),
    ], size=11)

    # ② 方案介绍（示意图）
    x, w = 4.7, 8.23
    panel(s, x, top, w, top_h, "02", "方案介绍", TEAL)
    ix = x + 0.2
    box_text(s, ix, top + 0.55, 1.5, 0.62,
             [[("Agent / 网关", 10.5, NAVY, True)], [("会话 hook：停顿、复用预测", 8.5, GRAY, False)]],
             PINK, RED)
    # KVCM 控制面
    cx, cw = ix + 1.85, 6.0
    rect(s, cx, top + 0.5, cw, 0.98, RGBColor(0xEE, 0xF6, 0xF7), TEAL, MSO_SHAPE.ROUNDED_RECTANGLE)
    text(s, cx + 0.1, top + 0.52, 3.0, 0.24, [[("KVCM 控制面（旁挂）", 10, TEAL, True)]], margin=0.02)
    subs = [("① 会话感知保留", "停顿期 pin / 预取，按“未来复用”驱逐"),
            ("② PD 动态转换", "按配比切换 P/D，KV 经池保活"),
            ("③ 局部性感知路由", "计入数据位置 + 网卡队列 + 命中")]
    sw, gap = 1.88, 0.1
    for i, (h1, h2) in enumerate(subs):
        box_text(s, cx + 0.1 + i * (sw + gap), top + 0.78, sw, 0.62,
                 [[(h1, 9.5, NAVY, True)], [(h2, 8, INK, False)]], WHITE, TEAL, margin=0.03)
    arrow(s, ix + 1.5, top + 0.86, cx, top + 0.86, RED)

    # 引擎实例（HBM）
    ey, eh = top + 1.7, 0.55
    engines = [("P 实例", "HBM · 引擎 LRU"), ("D 实例", "HBM · 引擎 LRU"),
               ("P⇄D 可切换实例", "HBM · 引擎 LRU")]
    ew, eg = 1.95, 0.2
    ex0 = cx + (cw - (3 * ew + 2 * eg)) / 2
    for i, (h1, h2) in enumerate(engines):
        exi = ex0 + i * (ew + eg)
        box_text(s, exi, ey, ew, eh, [[(h1, 9.5, WHITE, True)], [(h2, 8, RGBColor(0xD8, 0xE1, 0xEC), False)]],
                 NAVY)
        arrow(s, exi + ew / 2, top + 1.48, exi + ew / 2, ey, TEAL, width=1)
        arrow(s, exi + ew / 2, ey + eh, exi + ew / 2, ey + eh + 0.3, GREEN, both=True, width=1.25)
    arrow(s, ix + 0.75, top + 1.17, ix + 0.75, ey + eh / 2, RED)
    arrow(s, ix + 0.75, ey + eh / 2, ex0, ey + eh / 2, RED)
    text(s, ix + 0.8, ey + eh / 2 - 0.26, 0.9, 0.24, [[("请求选路", 8.5, RED, True)]],
         align=PP_ALIGN.CENTER, margin=0)
    text(s, cx - 0.02, top + 1.5, 1.2, 0.22, [[("调度 / 视图", 8, TEAL, False)]], margin=0)

    # DRAM 共享池
    py = ey + eh + 0.3
    box_text(s, cx, py, cw, 0.42,
             [[("DRAM 共享池（集群级，KVCM 管理）", 10, WHITE, True),
               ("  跨实例命中 · 哈希条带单副本 · 保活写回", 8.5, WHITE, False)]], GREEN)
    text(s, ix - 0.05, py - 0.02, 1.85, 0.5,
         [[("KV 读写 / 写回", 8.5, GREEN, True)], [("故障时引擎回退原生路径", 8, GRAY, False)]],
         align=PP_ALIGN.CENTER, margin=0, spacing=1.0, space_after=0)

    # ③ 创新点
    x, w = 0.4, 6.2
    panel(s, x, bot, w, bot_h, "03", "创新点", NAVY)
    bullets(s, x + 0.16, bot + 0.5, w - 0.3, bot_h - 0.55, [
        ("会话感知保留", "用“预测未来复用”的 hook 替代 LRU 做保留 / 驱逐，停顿期保活而不是错杀"),
        ("HBM + DRAM 共享池", "KV 从实例私有升为集群共享资源，跨实例命中，池容量独立于算力扩展"),
        ("PD 动态转换 + 池保活", "P/D 角色随配比在线切换，KV 留在池里，切换不丢、不重算"),
        ("局部性感知路由", "选路同时计数据位置、网卡出口队列和命中，网卡紧张时防止排队失稳"),
        ("池简单、策略复杂", "池用哈希条带 + 单副本即可避免热点，复杂度投在保留策略和路由"),
    ], size=11, accent=NAVY)

    # ④ 商业价值
    x, w = 6.73, 6.2
    panel(s, x, bot, w, bot_h, "04", "商业价值", GREEN)
    bullets(s, x + 0.16, bot + 0.5, w - 0.3, bot_h - 0.86, [
        ("省卡", "Claude Code trace、3.5 TB DRAM 池、理想 hook：prefill NPU 251.9 → 210.4，"
               "TTFT P95 2.38 → 1.64 s；现实 hook 仅 250.9，hook 质量决定收益"),
        ("增吞吐", "【注入】PD 动态转换 S3 goodput +18.6%（滞后控制器 + 池 +13.5%）"),
        ("稳尾延迟", "【注入】网卡紧张（最坏组合出口负载 158–758%）时，只有 KVCM + 局部性稳定，"
                 "P95 1.56 s / 0.57 s（CC / Qwen），其余变体失稳"),
        ("低接入成本", "旁挂零 Patch，引擎与存储底座两侧可替换"),
    ], size=11, accent=GREEN)
    text(s, x + 0.16, bot + bot_h - 0.36, w - 0.3, 0.32,
         [[("口径：trace 回放仿真（DeepSeek V4.1-Flash，昇腾 A3 14 节点 × 16 卡，HBM + DRAM 两层），"
            "无真机校准；【注入】为人为设定场景", 8, GRAY, False)]],
         margin=0.02, spacing=1.0, space_after=0)

    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return out


if __name__ == "__main__":
    print(build(Path(sys.argv[1]) if len(sys.argv) > 1 else OUT))
