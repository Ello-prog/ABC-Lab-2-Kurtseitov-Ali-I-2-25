#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Презентация: Современные дезинфицирующие средства (совместимо с телефоном)."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

NAVY = RGBColor(0x0B, 0x3D, 0x5C)
TEAL = RGBColor(0x1A, 0x6B, 0x6B)
ACCENT = RGBColor(0x0D, 0x8A, 0x7A)
LIGHT = RGBColor(0xF2, 0xF7, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1A, 0x2A, 0x2E)
MUTED = RGBColor(0x3A, 0x4A, 0x4E)
GOLD = RGBColor(0xB8, 0x8A, 0x2E)
ROW_ALT = RGBColor(0xE8, 0xF2, 0xF2)


def set_run(run, size=16, bold=False, color=DARK, name="Arial"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name
    # Explicit latin/cs for mobile renderers
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:cs", "a:ea"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", name)


def fill_shape(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_shape(slide, left, top, width, height, fill, rounded=False):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(kind, left, top, width, height)
    fill_shape(sh, fill)
    return sh


def shape_text(shape, paragraphs, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, margin=0.12):
    """Write all text INTO the shape (no overlapping textboxes)."""
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = None
    try:
        shape.text_frame.paragraphs[0].alignment = align
    except Exception:
        pass
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin * 0.7)
    tf.margin_bottom = Inches(margin * 0.5)
    bodyPr = tf._txBody.bodyPr
    anchor = {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[valign]
    bodyPr.set("anchor", anchor)

    first = True
    for item in paragraphs:
        if isinstance(item, str):
            text, size, bold, color, palign = item, 15, False, DARK, align
        else:
            text = item.get("text", "")
            size = item.get("size", 15)
            bold = item.get("bold", False)
            color = item.get("color", DARK)
            palign = item.get("align", align)
            sp_before = item.get("before", 0)
            sp_after = item.get("after", 2)
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = palign
        if not isinstance(item, str):
            p.space_before = Pt(item.get("before", 0))
            p.space_after = Pt(item.get("after", 2))
        else:
            p.space_before = Pt(0)
            p.space_after = Pt(2)
        run = p.add_run()
        run.text = text
        set_run(run, size=size, bold=bold, color=color)
    return shape


def plain_text(slide, left, top, width, height, paragraphs, align=PP_ALIGN.LEFT):
    """Standalone textbox only — no white shape behind it."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    first = True
    for item in paragraphs:
        if isinstance(item, str):
            text, size, bold, color, palign = item, 15, False, DARK, align
            before, after = 0, 4
        else:
            text = item["text"]
            size = item.get("size", 15)
            bold = item.get("bold", False)
            color = item.get("color", DARK)
            palign = item.get("align", align)
            before = item.get("before", 0)
            after = item.get("after", 4)
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = palign
        p.space_before = Pt(before)
        p.space_after = Pt(after)
        run = p.add_run()
        run.text = text
        set_run(run, size=size, bold=bold, color=color)
    return box


def slide_bg(slide, color=LIGHT):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def header(slide, title):
    bar = add_shape(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.9), NAVY)
    shape_text(
        bar,
        [{"text": title, "size": 24, "bold": True, "color": WHITE, "align": PP_ALIGN.LEFT}],
        align=PP_ALIGN.LEFT,
        valign=MSO_ANCHOR.MIDDLE,
        margin=0.4,
    )
    add_shape(slide, Inches(0), Inches(0.9), Inches(13.333), Inches(0.07), ACCENT)


def footer(slide, page, total=11):
    bar = add_shape(slide, Inches(0), Inches(7.15), Inches(13.333), Inches(0.35), NAVY)
    shape_text(
        bar,
        [{
            "text": f"Медицинский колледж КФУ им. В.И. Вернадского  ·  Современные дезсредства          {page} / {total}",
            "size": 11,
            "color": RGBColor(0xC5, 0xD8, 0xDE),
            "align": PP_ALIGN.LEFT,
        }],
        valign=MSO_ANCHOR.MIDDLE,
        margin=0.35,
    )


def content_card(slide, left, top, width, height, title, body, header_color=TEAL):
    """One card = one shape with title+body inside (mobile-safe)."""
    sh = add_shape(slide, left, top, width, height, WHITE, rounded=True)
    # left accent via nested thin shape is OK if it doesn't cover text area —
    # put accent as thin bar LEFT, then text shape to the RIGHT of it.
    accent = add_shape(slide, left, top, Inches(0.14), height, header_color)
    # Main text area: separate shape to the right of accent (not overlapping accent)
    text_sh = add_shape(
        slide,
        left + Inches(0.14),
        top,
        width - Inches(0.14),
        height,
        WHITE,
        rounded=False,
    )
    paras = [
        {"text": title, "size": 15, "bold": True, "color": NAVY, "after": 8},
    ]
    if isinstance(body, str):
        paras.append({"text": body, "size": 13, "color": MUTED, "after": 2})
    else:
        for line in body:
            paras.append({"text": "•  " + line, "size": 13, "color": DARK, "before": 3, "after": 2})
    shape_text(text_sh, paras, valign=MSO_ANCHOR.TOP, margin=0.16)
    return text_sh


def simple_block(slide, left, top, width, height, title, body, title_bg=TEAL):
    """Header strip + body in ONE combined approach using two non-overlapping shapes."""
    head = add_shape(slide, left, top, width, Inches(0.5), title_bg)
    shape_text(
        head,
        [{"text": title, "size": 13, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}],
        align=PP_ALIGN.CENTER,
        valign=MSO_ANCHOR.MIDDLE,
        margin=0.08,
    )
    body_sh = add_shape(slide, left, top + Inches(0.5), width, height - Inches(0.5), WHITE)
    if isinstance(body, str):
        paras = [{"text": body, "size": 13, "color": DARK, "align": PP_ALIGN.CENTER}]
        shape_text(body_sh, paras, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, margin=0.12)
    else:
        paras = [{"text": "•  " + x, "size": 12, "color": DARK, "before": 4, "after": 2} for x in body]
        shape_text(body_sh, paras, valign=MSO_ANCHOR.TOP, margin=0.14)


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    TOTAL = 11

    # ===== 1. Title =====
    s = prs.slides.add_slide(blank)
    slide_bg(s, WHITE)
    add_shape(s, Inches(0), Inches(0), Inches(13.333), Inches(0.12), NAVY)
    add_shape(s, Inches(0), Inches(7.38), Inches(13.333), Inches(0.12), NAVY)

    header_paras = [
        {"text": "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
         "size": 11, "bold": True, "color": NAVY, "align": PP_ALIGN.CENTER, "after": 4},
        {"text": "Федеральное государственное автономное образовательное учреждение высшего образования",
         "size": 11, "color": NAVY, "align": PP_ALIGN.CENTER, "after": 3},
        {"text": "«КРЫМСКИЙ ФЕДЕРАЛЬНЫЙ УНИВЕРСИТЕТ имени В.И. ВЕРНАДСКОГО»",
         "size": 12, "bold": True, "color": NAVY, "align": PP_ALIGN.CENTER, "after": 4},
        {"text": "Медицинский колледж",
         "size": 12, "bold": True, "color": NAVY, "align": PP_ALIGN.CENTER, "after": 2},
        {"text": "Ордена Трудового Красного Знамени",
         "size": 11, "color": NAVY, "align": PP_ALIGN.CENTER, "after": 2},
        {"text": "Медицинского института им. С.И. Георгиевского",
         "size": 11, "color": NAVY, "align": PP_ALIGN.CENTER, "after": 2},
    ]
    plain_text(s, Inches(0.5), Inches(0.35), Inches(12.3), Inches(2.1), header_paras, PP_ALIGN.CENTER)

    add_shape(s, Inches(4.8), Inches(2.5), Inches(3.7), Inches(0.04), GOLD)

    mid = add_shape(s, Inches(1), Inches(2.8), Inches(11.3), Inches(1.6), LIGHT)
    shape_text(mid, [
        {"text": "ПРЕЗЕНТАЦИЯ", "size": 14, "bold": True, "color": ACCENT, "align": PP_ALIGN.CENTER, "after": 8},
        {"text": "Современные дезинфицирующие средства", "size": 28, "bold": True, "color": NAVY, "align": PP_ALIGN.CENTER, "after": 8},
        {"text": "по дисциплине: Инфекционная безопасность / Сестринское дело",
         "size": 13, "color": MUTED, "align": PP_ALIGN.CENTER},
    ], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, margin=0.2)

    author = add_shape(s, Inches(7.2), Inches(4.8), Inches(5.3), Inches(1.5), LIGHT)
    shape_text(author, [
        {"text": "Выполнила работу:", "size": 13, "bold": True, "color": TEAL, "after": 6},
        {"text": "студентка 261/1 группы", "size": 14, "color": DARK, "after": 4},
        {"text": "Аметова Эмине Суюновна", "size": 16, "bold": True, "color": NAVY, "after": 2},
    ], valign=MSO_ANCHOR.MIDDLE, margin=0.25)

    plain_text(s, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4), [
        {"text": "Симферополь, 2026", "size": 13, "color": MUTED, "align": PP_ALIGN.CENTER},
    ], PP_ALIGN.CENTER)

    # ===== 2. Plan =====
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header(s, "План презентации")
    footer(s, 2, TOTAL)
    plan = [
        ("01", "Как действует", "Механизм действия дезинфицирующих средств на микроорганизмы"),
        ("02", "Какие есть", "Основные виды современных дезсредств"),
        ("03", "Что дают", "Значение и эффект применения в медицине"),
        ("04", "Кто занимается", "Специалисты и организации, отвечающие за дезинфекцию"),
        ("05", "Группы", "Классификация по химическому составу и назначению"),
    ]
    for i, (num, title, desc) in enumerate(plan):
        top = Inches(1.2 + i * 1.1)
        num_sh = add_shape(s, Inches(0.7), top, Inches(1.1), Inches(0.9), TEAL)
        shape_text(num_sh, [{"text": num, "size": 22, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}],
                   align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        body = add_shape(s, Inches(1.8), top, Inches(10.8), Inches(0.9), WHITE)
        shape_text(body, [
            {"text": title, "size": 18, "bold": True, "color": NAVY, "after": 4},
            {"text": desc, "size": 13, "color": MUTED},
        ], valign=MSO_ANCHOR.MIDDLE, margin=0.25)

    # ===== 3. How it works =====
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header(s, "Как действует дезинфицирующее средство")
    footer(s, 3, TOTAL)
    intro = add_shape(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.55), WHITE)
    shape_text(intro, [{
        "text": "Дезинфекция — уничтожение патогенных и условно-патогенных микроорганизмов на объектах внешней среды.",
        "size": 14, "color": DARK,
    }], valign=MSO_ANCHOR.MIDDLE, margin=0.2)

    mechs = [
        ("Денатурация белков", [
            "Разрушение структуры белков клетки",
            "Нарушение жизненно важных ферментов",
            "Спирты, альдегиды, фенолы",
        ]),
        ("Повреждение мембраны", [
            "Нарушение клеточной стенки / оболочки",
            "Утечка цитоплазмы, гибель клетки",
            "ЧАС, хлорсодержащие средства",
        ]),
        ("Окисление", [
            "Окисление жизненно важных структур",
            "Повреждение ДНК и ферментов",
            "Перекись водорода, хлор, озон",
        ]),
        ("Коагуляция / фиксация", [
            "Свёртывание белков протоплазмы",
            "Блокировка обмена веществ",
            "Спирты, кислоты",
        ]),
    ]
    for i, (title, lines) in enumerate(mechs):
        col, row = i % 2, i // 2
        content_card(
            s,
            Inches(0.5 + col * 6.4),
            Inches(1.9 + row * 2.45),
            Inches(6.15),
            Inches(2.25),
            title,
            lines,
        )

    # ===== 4. Types =====
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header(s, "Какие есть современные дезсредства")
    footer(s, 4, TOTAL)
    types = [
        ("Хлорсодержащие", "Гипохлорит натрия, хлорамин Б, жавель и др."),
        ("Кислородсодержащие", "Перекись водорода, пергидроль, озон, ПВК"),
        ("Спирты", "Этиловый 70%, изопропиловый спирт, антисептики"),
        ("Альдегиды", "Глутаровый, ортофталевый, формальдегид"),
        ("ЧАС (катионные)", "Бензалкония хлорид, цетилпиридиния хлорид"),
        ("Гуанидины", "Хлоргексидин биглюконат, полигексанид"),
        ("Фенолы", "Триклозан, ортофенилфенол (ограниченно)"),
        ("Комбинированные", "Многокомпонентные составы (ЧАС + спирт и др.)"),
    ]
    for i, (t, d) in enumerate(types):
        col, row = i % 4, i // 4
        left = Inches(0.35 + col * 3.25)
        top = Inches(1.25 + row * 2.8)
        simple_block(s, left, top, Inches(3.1), Inches(2.5), t, d, TEAL if row == 0 else NAVY)

    # ===== 5. Benefits =====
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header(s, "Что дают дезинфицирующие средства")
    footer(s, 5, TOTAL)
    benefits = [
        ("Инфекционная безопасность", "Снижают риск внутрибольничных инфекций (ВБИ) для пациентов и персонала"),
        ("Прерывание путей передачи", "Разрывают цепь заражения: руки, поверхности, инструменты, воздух"),
        ("Стерильность процедур", "Готовят изделия медицинского назначения к стерилизации и использованию"),
        ("Эпидблагополучие", "Предотвращают вспышки инфекционных заболеваний в ЛПУ и обществе"),
        ("Защита персонала", "Снижают профессиональный риск заражения медработников"),
        ("Качество помощи", "Обеспечивают стандарты санитарно-противоэпидемического режима"),
    ]
    for i, (t, d) in enumerate(benefits):
        col, row = i % 3, i // 3
        content_card(
            s,
            Inches(0.4 + col * 4.3),
            Inches(1.25 + row * 2.75),
            Inches(4.1),
            Inches(2.5),
            t,
            d,
        )

    # ===== 6. Who =====
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header(s, "Кто занимается дезинфекцией")
    footer(s, 6, TOTAL)
    who = [
        ("В лечебном учреждении", TEAL, [
            "Медицинские сёстры и младший медперсонал",
            "Врач-эпидемиолог / помощник эпидемиолога",
            "Старшая медицинская сестра отделения",
            "Дезинфектор (при наличии ставки)",
            "Главная медсестра — контроль режима",
        ]),
        ("Государственный контроль", NAVY, [
            "Роспотребнадзор (санэпиднадзор)",
            "Территориальные органы здравоохранения",
            "Лаборатории контроля качества дезинфекции",
            "Центры гигиены и эпидемиологии",
        ]),
        ("Наука и производство", ACCENT, [
            "НИИ дезинфектологии",
            "Производители дезсредств (регистрация в РФ)",
            "Кафедры эпидемиологии и гигиены вузов",
            "Методические центры Минздрава России",
        ]),
    ]
    for i, (title, color, items) in enumerate(who):
        left = Inches(0.35 + i * 4.3)
        simple_block(s, left, Inches(1.2), Inches(4.1), Inches(5.6), title, items, color)

    # ===== 7. Groups =====
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header(s, "Группы дезсредств и к каким относятся")
    footer(s, 7, TOTAL)
    plain_text(s, Inches(0.5), Inches(1.1), Inches(12), Inches(0.35), [
        {"text": "Классификация по химическому составу (основные группы)", "size": 15, "bold": True, "color": TEAL},
    ])
    groups = [
        ("1. Галогенсодержащие", "Хлор, бром, йод и их соединения — жавель, хлорамин, йодофоры"),
        ("2. Кислородсодержащие", "Перекись водорода, перкарбонаты, озон, надуксусная кислота"),
        ("3. ПАВ / ЧАС", "Четвертичные аммониевые соединения — поверхности, инвентарь"),
        ("4. Спирты", "Этанол, пропанол — кожные антисептики, мелкий инструментарий"),
        ("5. Альдегиды", "Высокий уровень дезинфекции, эндоскопы, термолабильные изделия"),
        ("6. Гуанидины", "Хлоргексидин и аналоги — кожа, слизистые, поверхности"),
        ("7. Кислоты / щёлочи", "Специальные режимы, отдельные отрасли и ЛПУ"),
        ("8. Комбинированные", "Сочетание групп для расширения спектра и снижения резистентности"),
    ]
    for i, (t, d) in enumerate(groups):
        top = Inches(1.5 + i * 0.65)
        bg = WHITE if i % 2 == 0 else ROW_ALT
        row = add_shape(s, Inches(0.5), top, Inches(12.3), Inches(0.58), bg)
        shape_text(row, [
            {"text": f"{t}   —   {d}", "size": 13, "color": DARK},
        ], valign=MSO_ANCHOR.MIDDLE, margin=0.2)

    # ===== 8. By purpose =====
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header(s, "Классификация по назначению и уровню")
    footer(s, 8, TOTAL)
    levels = [
        ("По объекту обработки", [
            "Кожные антисептики (руки, операционное поле)",
            "Для поверхностей помещений и мебели",
            "Для изделий медицинского назначения (ИМН)",
            "Для воздуха и аэрозольной дезинфекции",
            "Для белья, посуды, выделений больных",
        ]),
        ("По уровню дезинфекции", [
            "Высокий уровень — почти стерилизация",
            "Средний — бактерии, вирусы, грибы, микобактерии",
            "Низкий — вегетативные бактерии, некоторые вирусы",
            "Выбор зависит от риска и вида ИМН",
            "Режимы указаны в инструкции к средству",
        ]),
        ("По режиму применения", [
            "Профилактическая дезинфекция",
            "Текущая дезинфекция (у постели больного)",
            "Заключительная дезинфекция",
            "Очаговая (при инфекционном заболевании)",
            "Генеральная уборка по графику ЛПУ",
        ]),
    ]
    for i, (title, items) in enumerate(levels):
        simple_block(s, Inches(0.35 + i * 4.3), Inches(1.2), Inches(4.1), Inches(5.6), title, items, ACCENT)

    # ===== 9. Requirements =====
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header(s, "Требования к современным дезсредствам")
    footer(s, 9, TOTAL)
    reqs = [
        ("Широкий спектр", "Бактерицидное, вирулицидное, фунгицидное, при необходимости туберкулоцидное действие"),
        ("Безопасность", "Малая токсичность в рабочих концентрациях, контроль экспозиции для персонала"),
        ("Совместимость", "Не повреждают материалы ИМН, мебели, поверхностей при правильном применении"),
        ("Удобство", "Готовые растворы / таблетки, понятные режимы, стабильность рабочего раствора"),
        ("Регистрация", "Разрешены к применению в РФ, есть инструкция и свидетельство о госрегистрации"),
        ("Контроль", "Возможность химического и бактериологического контроля качества обработки"),
    ]
    for i, (t, d) in enumerate(reqs):
        col, row = i % 3, i // 3
        content_card(
            s,
            Inches(0.4 + col * 4.3),
            Inches(1.25 + row * 2.75),
            Inches(4.1),
            Inches(2.5),
            t,
            d,
        )

    # ===== 10. Conclusions =====
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header(s, "Выводы")
    footer(s, 10, TOTAL)
    conclusions = [
        "Современные дезсредства действуют через денатурацию белков, окисление и повреждение клеточных мембран.",
        "Основные группы: хлор- и кислородсодержащие, спирты, альдегиды, ЧАС, гуанидины и комбинированные препараты.",
        "Обеспечивают инфекционную безопасность пациентов и персонала, снижают риск ВБИ.",
        "Организацию и контроль осуществляют медперсонал ЛПУ, эпидемиологи и органы Роспотребнадзора.",
        "Выбор средства зависит от объекта, уровня дезинфекции и режима; обязательно соблюдение инструкции.",
    ]
    for i, text in enumerate(conclusions):
        top = Inches(1.2 + i * 1.1)
        num = add_shape(s, Inches(0.6), top + Inches(0.15), Inches(0.55), Inches(0.55), TEAL)
        shape_text(num, [{"text": str(i + 1), "size": 16, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER}],
                   align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, margin=0.05)
        body = add_shape(s, Inches(1.35), top, Inches(11.3), Inches(0.9), WHITE)
        shape_text(body, [{"text": text, "size": 14, "color": DARK}],
                   valign=MSO_ANCHOR.MIDDLE, margin=0.22)

    # ===== 11. Thanks =====
    s = prs.slides.add_slide(blank)
    slide_bg(s, WHITE)
    add_shape(s, Inches(0), Inches(0), Inches(13.333), Inches(0.12), NAVY)
    add_shape(s, Inches(0), Inches(7.38), Inches(13.333), Inches(0.12), NAVY)
    mid = add_shape(s, Inches(0), Inches(2.3), Inches(13.333), Inches(2.4), NAVY)
    shape_text(mid, [
        {"text": "Спасибо за внимание!", "size": 36, "bold": True, "color": WHITE, "align": PP_ALIGN.CENTER, "after": 12},
        {"text": "Готова ответить на ваши вопросы", "size": 18, "color": RGBColor(0xA8, 0xD5, 0xD0), "align": PP_ALIGN.CENTER},
    ], align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    plain_text(s, Inches(0.5), Inches(5.2), Inches(12.3), Inches(0.9), [
        {"text": "Аметова Эмине Суюновна  ·  группа 261/1", "size": 15, "bold": True, "color": NAVY, "align": PP_ALIGN.CENTER, "after": 6},
        {"text": "Медицинский колледж КФУ им. В.И. Вернадского", "size": 13, "color": MUTED, "align": PP_ALIGN.CENTER},
    ], PP_ALIGN.CENTER)

    out = "/workspace/Sovremennye_dezsredstva_Ametova.pptx"
    prs.save(out)
    print("Saved:", out)


if __name__ == "__main__":
    main()
