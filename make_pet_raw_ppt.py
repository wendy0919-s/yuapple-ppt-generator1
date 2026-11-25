from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ===== 视觉规范 =====
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GRAY = RGBColor(0xF5, 0xF5, 0xF7)
GREEN = RGBColor(0x2F, 0x5D, 0x3A)
SAGE = RGBColor(0xA8, 0xC3, 0xA0)
INK = RGBColor(0x1D, 0x1D, 0x1F)
MUTED = RGBColor(0x6E, 0x6E, 0x73)
DARK_MUTED = RGBColor(0xC7, 0xD6, 0xC3)
PLACE = RGBColor(0xE8, 0xEC, 0xE6)

FONT_TITLE = "PingFang SC"   # 没有苹方可改成 "微软雅黑" 或 "思源黑体"
FONT_BODY = "PingFang SC"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def set_bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def add_textbox(slide, left, top, width, height, lines, size=22, color=INK,
                bold=False, align=PP_ALIGN.LEFT, font=FONT_BODY,
                line_spacing=1.15, space_after=6, anchor=MSO_ANCHOR.TOP):
    if isinstance(lines, str):
        lines = [lines]
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor

    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_before = Pt(0)
        p.space_after = Pt(space_after)
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        r.font.name = font
    return box


def add_dot(slide, left, top, color):
    dot = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(left), Inches(top), Inches(0.10), Inches(0.10)
    )
    dot.fill.solid()
    dot.fill.fore_color.rgb = color
    dot.line.fill.background()
    try:
        dot.shadow.inherit = False
    except Exception:
        pass


def add_image_placeholder(slide, left, top, width, height, hint):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = PLACE
    shape.line.fill.background()
    try:
        shape.shadow.inherit = False
    except Exception:
        pass

    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)

    p1 = tf.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    r1 = p1.add_run()
    r1.text = "图片占位"
    r1.font.size = Pt(12)
    r1.font.bold = True
    r1.font.color.rgb = MUTED
    r1.font.name = FONT_BODY

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(4)
    r2 = p2.add_run()
    r2.text = hint
    r2.font.size = Pt(10)
    r2.font.color.rgb = MUTED
    r2.font.name = FONT_BODY


def add_page_number(slide, idx, dark=False):
    add_textbox(
        slide, 12.0, 7.12, 0.75, 0.2, [str(idx)],
        size=9, color=DARK_MUTED if dark else MUTED, align=PP_ALIGN.RIGHT
    )


def add_footer(slide, idx, dark=False):
    add_textbox(
        slide, 0.55, 7.12, 6.0, 0.2, ["宇苹果 · 宠物生食模板"],
        size=9, color=DARK_MUTED if dark else MUTED
    )
    add_page_number(slide, idx, dark)


def add_title(slide, title, dark=False, top=0.9, width=10.5):
    size = 40
    if len(title) > 22:
        size = 31
    elif len(title) > 18:
        size = 34
    elif len(title) > 14:
        size = 37
    add_textbox(
        slide, 0.85, top, width, 1.5, [title],
        size=size, color=WHITE if dark else INK,
        bold=True, font=FONT_TITLE, line_spacing=1.0
    )


def add_note(slide, note, dark=False):
    if note:
        add_textbox(
            slide, 0.85, 6.42, 8.5, 0.3, [note],
            size=11, color=DARK_MUTED if dark else GREEN
        )


def layout_cover(slide, page, idx):
    set_bg(slide, GREEN)
    add_textbox(
        slide, 0.85, 1.55, 10.8, 1.8, [page["title"]],
        size=52, color=WHITE, bold=True, font=FONT_TITLE, line_spacing=1.0
    )
    add_textbox(
        slide, 0.85, 3.55, 9.0, 1.4, page["body"],
        size=24, color=WHITE, line_spacing=1.35, space_after=8
    )
    add_textbox(
        slide, 0.85, 6.45, 8.0, 0.3, [page.get("brand", "宠物生食介绍 · 模板")],
        size=11, color=DARK_MUTED
    )
    add_page_number(slide, idx, dark=True)


def layout_statement(slide, page, idx):
    set_bg(slide, WHITE)
    add_title(slide, page["title"], top=1.0)
    add_textbox(
        slide, 0.85, 3.15, 8.8, 2.0, page["body"],
        size=24, color=INK, line_spacing=1.35, space_after=10
    )
    add_note(slide, page.get("note", ""), dark=False)
    add_footer(slide, idx, dark=False)


def layout_three(slide, page, idx, dark=False):
    set_bg(slide, GREEN if dark else page.get("bg", WHITE))
    add_title(slide, page["title"], dark=dark, top=0.9)
    lines = page["body"]
    n = len(lines)
    start = 2.55 if n >= 4 else 2.75
    step = 0.66 if n >= 4 else 0.86
    size = 26 if n >= 4 else 30
    for i, line in enumerate(lines):
        top = start + i * step
        add_dot(slide, 0.88, top + 0.13, SAGE if dark else GREEN)
        add_textbox(
            slide, 1.13, top, 10.8, 0.55, [line],
            size=size, color=WHITE if dark else INK,
            line_spacing=1.0
        )
    add_note(slide, page.get("note", ""), dark=dark)
    add_footer(slide, idx, dark=dark)


def layout_split(slide, page, idx):
    set_bg(slide, page.get("bg", WHITE))
    add_title(slide, page["title"], top=0.9, width=6.4)
    add_textbox(
        slide, 0.85, 2.75, 6.0, 3.2, page["body"],
        size=21, color=INK, line_spacing=1.3, space_after=9
    )
    add_image_placeholder(slide, 7.45, 1.35, 5.05, 4.95, page.get("img", "真实素材照片"))
    add_note(slide, page.get("note", ""), dark=False)
    add_footer(slide, idx, dark=False)


def layout_closing(slide, page, idx):
    set_bg(slide, GREEN)
    add_textbox(
        slide, 0.9, 1.75, 10.8, 1.4, [page["title"]],
        size=50, color=WHITE, bold=True, font=FONT_TITLE, line_spacing=1.0
    )
    add_textbox(
        slide, 0.9, 3.55, 10.2, 1.8, page["body"],
        size=23, color=WHITE, line_spacing=1.35, space_after=9
    )
    add_textbox(
        slide, 0.9, 6.35, 9.0, 0.3, [page.get("brand", "谢谢")],
        size=11, color=DARK_MUTED
    )
    add_page_number(slide, idx, dark=True)


pages = [
    {"type": "cover", "title": "生食，回到食物本身", "body": ["一份看得懂的宠物生食。", "少一点解释。多一点信任。"], "brand": "宠物生食介绍 · 模板"},
    {"type": "statement", "title": "一餐，应该被看懂", "body": ["不是把生食说得很神。", "是把一餐讲清楚。"], "note": "产品概念"},
    {"type": "split", "title": "配料表越来越长，信任越来越短", "body": ["主人想知道：它吃什么。", "也想知为什么。"], "img": "真实食材近景，低饱和自然光", "bg": GRAY},
    {"type": "statement", "title": "一份可被解释的生食", "body": ["生食不是一种风格。", "是一套可解释的喂养方案。"]},
    {"type": "three", "title": "三条原则", "body": ["少加工。", "可追溯。", "有指导。"], "bg": WHITE},
    {"type": "split", "title": "肉、骨、内脏与配方，结构清楚", "body": ["食材各司其职。", "比例由配方决定。"], "img": "3—4 个真实食材小图，等距排列", "bg": GRAY},
    {"type": "three", "title": "从第一口，到每一天", "body": ["日常主餐。", "换粮过渡。", "挑食与多宠管理。"], "bg": WHITE, "note": "应用场景"},
    {"type": "split", "title": "每天，稳定地吃", "body": ["稳定，比惊喜重要。", "分量清楚，节奏固定。"], "img": "宠物进食侧影或手部配餐", "bg": WHITE},
    {"type": "statement", "title": "慢一点，稳一点", "body": ["换粮不是切换。", "是渐进。"], "note": "少量 · 观察 · 调整"},
    {"type": "split", "title": "适口，也要好管理", "body": ["挑食，先看愿不愿意尝试。", "多宠，先看能不能分开管理。"], "img": "多宠或分餐照片", "bg": WHITE},
    {"type": "three_dark", "title": "不是更多，是更少", "body": ["少不必要添加。", "少过度加工。", "少看不懂的术语。"], "note": "生食优势"},
    {"type": "split", "title": "配料表短一点，餐盘清楚一点", "body": ["食材可见。", "来源可讲。", "一餐可解释。"], "img": "食材微距大图", "bg": WHITE},
    {"type": "three", "title": "少一道加工，多一分可解释", "body": ["保留食材质感。", "减少不必要添加。", "营养以配方与检测为准。"], "bg": GRAY},
    {"type": "three", "title": "一餐的满足感，来自质感", "body": ["有水分。", "有咀嚼。", "有食物本来的样子。"], "bg": WHITE},
    {"type": "three", "title": "先安全，再营养，再适口", "body": ["安全是底线。", "营养要完整。", "适口决定坚持。"], "bg": WHITE, "note": "宠物食品逻辑"},
    {"type": "split", "title": "配料表，先做减法", "body": ["先问：为什么加。", "再问：能不能少。"], "img": "配料表局部，模糊处理，不展示虚构数据", "bg": GRAY},
    {"type": "statement", "title": "完整，比丰富更重要", "body": ["完整均衡来自配方与验证。", "不来自食材堆叠。"], "note": "配方需由兽医/宠物营养师审阅"},
    {"type": "three_dark", "title": "安全不是卖点，是底线", "body": ["冷链。", "分装。", "清洁。", "指导。"], "note": "生肉有微生物风险，需规范处理"},
    {"type": "split", "title": "案例一：换粮记录（示例）", "body": ["背景：成年犬，挑食。", "做法：7 天渐进换粮。", "观察：记录进食、饮水、便便。"], "img": "宠物侧影照片", "bg": WHITE, "note": "虚构练习案例，个体差异，需指导"},
    {"type": "split", "title": "案例二：适口记录（示例）", "body": ["背景：成年猫，对原粮兴趣低。", "做法：少量试吃，固定时间。", "观察：记录愿意尝试的次数。"], "img": "餐盘局部", "bg": GRAY, "note": "虚构练习案例，不构成医疗建议"},
    {"type": "split", "title": "案例三：多宠管理（示例）", "body": ["背景：两犬一猫。", "做法：分餐、分储、清洁。", "观察：流程更顺，误食更少。"], "img": "多宠侧影或储物场景", "bg": WHITE, "note": "示例案例，不构成医疗建议"},
    {"type": "closing", "title": "把每一餐，讲清楚", "body": ["不承诺奇迹。", "只把食物、配方与喂养方式说明白。", "生食有边界，安全与营养需专业指导。"], "brand": "谢谢 · 宠物生食介绍模板"},
]

for idx, page in enumerate(pages, start=1):
    slide = prs.slides.add_slide(BLANK)
    t = page["type"]
    if t == "cover":
        layout_cover(slide, page, idx)
    elif t == "statement":
        layout_statement(slide, page, idx)
    elif t == "three":
        layout_three(slide, page, idx, dark=False)
    elif t == "three_dark":
        layout_three(slide, page, idx, dark=True)
    elif t == "split":
        layout_split(slide, page, idx)
    elif t == "closing":
        layout_closing(slide, page, idx)

out = "宠物生食_宇苹果风格_22页模板.pptx"
prs.save(out)
print(f"已生成：{out}")
