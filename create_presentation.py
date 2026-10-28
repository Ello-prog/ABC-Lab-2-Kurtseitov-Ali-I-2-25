#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Презентация: Современные дезинфицирующие средства"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Colors — medical / academic (teal + deep navy, not purple)
NAVY = RGBColor(0x0B, 0x3D, 0x5C)
TEAL = RGBColor(0x1A, 0x6B, 0x6B)
ACCENT = RGBColor(0x0D, 0x8A, 0x7A)
LIGHT_BG = RGBColor(0xF2, 0xF7, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1A, 0x2A, 0x2E)
MUTED = RGBColor(0x4A, 0x5C, 0x60)
CARD_BG = RGBColor(0xE6, 0xF0, 0xF0)
GOLD = RGBColor(0xB8, 0x8A, 0x2E)


def set_run_font(run, size=18, bold=False, color=DARK, name="Arial"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name


def add_textbox(slide, left, top, width, height, text, size=18, bold=False,
                color=DARK, align=PP_ALIGN.LEFT, font="Arial", anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run_font(run, size=size, bold=bold, color=color, name=font)
    return box


def add_para(tf, text, size=16, bold=False, color=DARK, align=PP_ALIGN.LEFT,
             space_before=6, space_after=4, font="Arial", level=0):
    p = tf.add_paragraph()
    p.alignment = align
    p.level = level
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    run = p.add_run()
    run.text = text
    set_run_font(run, size=size, bold=bold, color=color, name=font)
    return p


def add_rect(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def add_round_rect(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def slide_bg(slide, color=LIGHT_BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def header_bar(slide, title):
    add_rect(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.95), NAVY)
    add_rect(slide, Inches(0), Inches(0.95), Inches(13.333), Inches(0.08), ACCENT)
    add_textbox(slide, Inches(0.5), Inches(0.22), Inches(12.3), Inches(0.6),
                title, size=26, bold=True, color=WHITE, align=PP_ALIGN.LEFT)


def footer(slide, page, total=10):
    add_rect(slide, Inches(0), Inches(7.2), Inches(13.333), Inches(0.3), NAVY)
    add_textbox(slide, Inches(0.4), Inches(7.22), Inches(10), Inches(0.25),
                "Медицинский колледж КФУ им. В.И. Вернадского  ·  Современные дезсредства",
                size=10, color=RGBColor(0xB0, 0xC8, 0xD0), align=PP_ALIGN.LEFT)
    add_textbox(slide, Inches(11.5), Inches(7.22), Inches(1.5), Inches(0.25),
                f"{page} / {total}", size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def bullet_block(slide, left, top, width, height, items, size=16):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(8)
        p.space_after = Pt(4)
        run = p.add_run()
        run.text = "●  " + item
        set_run_font(run, size=size, color=DARK)
    return box


def card(slide, left, top, width, height, title, body_lines, title_size=15, body_size=13):
    add_round_rect(slide, left, top, width, height, WHITE)
    # accent strip
    add_rect(slide, left, top, Inches(0.12), height, ACCENT)
    add_textbox(slide, left + Inches(0.28), top + Inches(0.15), width - Inches(0.4), Inches(0.4),
                title, size=title_size, bold=True, color=TEAL)
    box = slide.shapes.add_textbox(left + Inches(0.28), top + Inches(0.55),
                                   width - Inches(0.4), height - Inches(0.7))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in body_lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(3)
        run = p.add_run()
        run.text = line
        set_run_font(run, size=body_size, color=MUTED)


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    TOTAL = 11

    # ========== SLIDE 1: Title ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s, WHITE)
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.15), NAVY)
    add_rect(s, Inches(0), Inches(7.35), Inches(13.333), Inches(0.15), NAVY)

    header_lines = [
        "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
        "Федеральное государственное автономное образовательное учреждение высшего образования",
        "«КРЫМСКИЙ ФЕДЕРАЛЬНЫЙ УНИВЕРСИТЕТ имени В.И. ВЕРНАДСКОГО»",
        "Медицинский колледж",
        "Ордена Трудового Красного Знамени",
        "Медицинского института им. С.И. Георгиевского",
    ]
    y = 0.35
    for i, line in enumerate(header_lines):
        size = 11 if i == 0 else (12 if i == 2 else 11)
        bold = i in (0, 2, 3)
        add_textbox(s, Inches(0.6), Inches(y), Inches(12.1), Inches(0.32),
                    line, size=size, bold=bold, color=NAVY, align=PP_ALIGN.CENTER)
        y += 0.28 if i != 1 else 0.26

    add_rect(s, Inches(4.5), Inches(y + 0.15), Inches(4.3), Inches(0.04), GOLD)

    add_textbox(s, Inches(0.8), Inches(2.55), Inches(11.7), Inches(0.55),
                "ПРЕЗЕНТАЦИЯ", size=16, bold=True, color=ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0.8), Inches(3.05), Inches(11.7), Inches(0.7),
                "Современные дезинфицирующие средства",
                size=32, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0.8), Inches(3.75), Inches(11.7), Inches(0.4),
                "по дисциплине: Инфекционная безопасность / Сестринское дело",
                size=14, color=MUTED, align=PP_ALIGN.CENTER)

    add_textbox(s, Inches(7.5), Inches(5.0), Inches(5.2), Inches(0.35),
                "Выполнила работу:", size=14, bold=True, color=TEAL, align=PP_ALIGN.LEFT)
    add_textbox(s, Inches(7.5), Inches(5.35), Inches(5.2), Inches(0.35),
                "студентка 261/1 группы", size=14, color=DARK, align=PP_ALIGN.LEFT)
    add_textbox(s, Inches(7.5), Inches(5.7), Inches(5.2), Inches(0.4),
                "Аметова Эмине Суюновна", size=16, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
    add_textbox(s, Inches(0.8), Inches(6.5), Inches(11.7), Inches(0.35),
                "Симферополь, 2026", size=13, color=MUTED, align=PP_ALIGN.CENTER)

    # ========== SLIDE 2: Plan ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header_bar(s, "План презентации")
    footer(s, 2, TOTAL)
    plan = [
        ("01", "Как действует", "Механизм действия дезинфицирующих средств на микроорганизмы"),
        ("02", "Какие есть", "Основные виды современных дезсредств"),
        ("03", "Что дают", "Значение и эффект применения в медицине"),
        ("04", "Кто занимается", "Специалисты и организации, отвечающие за дезинфекцию"),
        ("05", "Группы", "Классификация по химическому составу и назначению"),
    ]
    for i, (num, title, desc) in enumerate(plan):
        top = Inches(1.3 + i * 1.05)
        add_round_rect(s, Inches(0.7), top, Inches(12), Inches(0.9), WHITE)
        add_rect(s, Inches(0.7), top, Inches(1.1), Inches(0.9), TEAL)
        add_textbox(s, Inches(0.7), top + Inches(0.22), Inches(1.1), Inches(0.5),
                    num, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(2.1), top + Inches(0.12), Inches(10), Inches(0.35),
                    title, size=18, bold=True, color=NAVY)
        add_textbox(s, Inches(2.1), top + Inches(0.48), Inches(10), Inches(0.35),
                    desc, size=13, color=MUTED)

    # ========== SLIDE 3: How it works ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header_bar(s, "Как действует дезинфицирующее средство")
    footer(s, 3, TOTAL)

    add_textbox(s, Inches(0.6), Inches(1.25), Inches(12), Inches(0.45),
                "Дезинфекция — уничтожение патогенных и условно-патогенных микроорганизмов "
                "на объектах внешней среды.",
                size=15, color=DARK)

    mechanisms = [
        ("Денатурация белков", [
            "Разрушение структуры белков клетки",
            "Нарушение жизненно важных ферментов",
            "Характерно для спиртов, альдегидов, фенолов",
        ]),
        ("Повреждение мембраны", [
            "Нарушение клеточной стенки / оболочки",
            "Утечка цитоплазмы, гибель клетки",
            "Характерно для ЧАС, хлорсодержащих",
        ]),
        ("Окисление", [
            "Окисление жизненно важных структур",
            "Повреждение ДНК и ферментов",
            "Характерно для перекиси, хлора, озона",
        ]),
        ("Коагуляция / фиксация", [
            "Свёртывание белков протоплазмы",
            "Блокировка обмена веществ",
            "Характерно для спиртов, кислот",
        ]),
    ]
    for i, (title, lines) in enumerate(mechanisms):
        col = i % 2
        row = i // 2
        left = Inches(0.55 + col * 6.35)
        top = Inches(1.85 + row * 2.45)
        card(s, left, top, Inches(6.05), Inches(2.25), title, ["• " + x for x in lines],
             title_size=16, body_size=14)

    # ========== SLIDE 4: Types ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header_bar(s, "Какие есть современные дезсредства")
    footer(s, 4, TOTAL)

    types = [
        ("Хлорсодержащие", "Гипохлорит натрия, хлорамин Б, жавель, хлоргексидин-хлор"),
        ("Кислородсодержащие", "Перекись водорода, пергидроль, ПВК, озон"),
        ("Спирты", "Этиловый 70%, изопропиловый спирт, спиртовые антисептики"),
        ("Альдегиды", "Глутаровый альдегид, ортофталевый альдегид, формальдегид"),
        ("ЧАС (катионные)", "Бензалкония хлорид, цетилпиридиния хлорид"),
        ("Гуанидины", "Хлоргексидин биглюконат, полигексанид"),
        ("Фенолы", "Триклозан, ортофенилфенол (ограниченно)"),
        ("Комбинированные", "Современные многокомпонентные составы (ЧАС + спирт и др.)"),
    ]
    for i, (t, d) in enumerate(types):
        col = i % 4
        row = i // 4
        left = Inches(0.4 + col * 3.2)
        top = Inches(1.35 + row * 2.7)
        add_round_rect(s, left, top, Inches(3.0), Inches(2.4), WHITE)
        add_rect(s, left, top, Inches(3.0), Inches(0.55), TEAL if row == 0 else NAVY)
        add_textbox(s, left + Inches(0.1), top + Inches(0.12), Inches(2.8), Inches(0.4),
                    t, size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, left + Inches(0.15), top + Inches(0.75), Inches(2.7), Inches(1.5),
                    d, size=13, color=DARK, align=PP_ALIGN.CENTER)

    # ========== SLIDE 5: What they give ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header_bar(s, "Что дают дезинфицирующие средства")
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
        col = i % 3
        row = i // 3
        left = Inches(0.45 + col * 4.25)
        top = Inches(1.35 + row * 2.7)
        add_round_rect(s, left, top, Inches(4.0), Inches(2.4), WHITE)
        add_rect(s, left, top, Inches(0.15), Inches(2.4), ACCENT)
        add_textbox(s, left + Inches(0.35), top + Inches(0.35), Inches(3.4), Inches(0.7),
                    t, size=16, bold=True, color=NAVY)
        add_textbox(s, left + Inches(0.35), top + Inches(1.15), Inches(3.4), Inches(1.0),
                    d, size=14, color=MUTED)

    # ========== SLIDE 6: Who deals ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header_bar(s, "Кто занимается дезинфекцией")
    footer(s, 6, TOTAL)

    who = [
        ("В лечебном учреждении", [
            "Медицинские сёстры и младший медперсонал",
            "Врач-эпидемиолог / помощник эпидемиолога",
            "Старшая медицинская сестра отделения",
            "Дезинфектор (при наличии ставки)",
            "Главная медицинская сестра — контроль режима",
        ]),
        ("Государственный контроль", [
            "Роспотребнадзор (санитарно-эпидемиологический надзор)",
            "Территориальные органы здравоохранения",
            "Лаборатории контроля качества дезинфекции",
            "Центры гигиены и эпидемиологии",
        ]),
        ("Наука и производство", [
            "НИИ дезинфектологии",
            "Производители дезсредств (регистрация в РФ)",
            "Кафедры эпидемиологии и гигиены вузов",
            "Методические центры Минздрава России",
        ]),
    ]
    for i, (title, items) in enumerate(who):
        left = Inches(0.4 + i * 4.25)
        add_round_rect(s, left, Inches(1.3), Inches(4.05), Inches(5.5), WHITE)
        add_rect(s, left, Inches(1.3), Inches(4.05), Inches(0.7), NAVY if i != 1 else TEAL)
        add_textbox(s, left + Inches(0.15), Inches(1.42), Inches(3.75), Inches(0.5),
                    title, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        bullet_block(s, left + Inches(0.25), Inches(2.2), Inches(3.6), Inches(4.3), items, size=13)

    # ========== SLIDE 7: Groups ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header_bar(s, "Группы дезсредств и к каким относятся")
    footer(s, 7, TOTAL)

    add_textbox(s, Inches(0.55), Inches(1.2), Inches(12), Inches(0.4),
                "Классификация по химическому составу (основные группы)",
                size=16, bold=True, color=TEAL)

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
        top = Inches(1.65 + i * 0.65)
        add_rect(s, Inches(0.55), top, Inches(12.2), Inches(0.55), WHITE if i % 2 == 0 else CARD_BG)
        add_textbox(s, Inches(0.7), top + Inches(0.1), Inches(3.8), Inches(0.4),
                    t, size=14, bold=True, color=NAVY)
        add_textbox(s, Inches(4.6), top + Inches(0.1), Inches(7.9), Inches(0.4),
                    d, size=13, color=DARK)

    # ========== SLIDE 8: By purpose ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header_bar(s, "Классификация по назначению и уровню")
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
            "Высокий уровень — почти стерилизация (альдегиды, НУК)",
            "Средний уровень — бактерии, вирусы, грибы, микобактерии",
            "Низкий уровень — вегетативные бактерии, некоторые вирусы",
            "Выбор режима зависит от риска и вида ИМН",
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
        left = Inches(0.4 + i * 4.25)
        add_round_rect(s, left, Inches(1.3), Inches(4.05), Inches(5.5), WHITE)
        add_rect(s, left, Inches(1.3), Inches(4.05), Inches(0.7), ACCENT)
        add_textbox(s, left + Inches(0.1), Inches(1.42), Inches(3.85), Inches(0.5),
                    title, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        bullet_block(s, left + Inches(0.25), Inches(2.2), Inches(3.6), Inches(4.3), items, size=13)

    # ========== SLIDE 9: Requirements ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header_bar(s, "Требования к современным дезсредствам")
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
        col = i % 3
        row = i // 3
        left = Inches(0.45 + col * 4.25)
        top = Inches(1.35 + row * 2.7)
        card(s, left, top, Inches(4.05), Inches(2.45), t, [d], title_size=16, body_size=14)

    # ========== SLIDE 10: Conclusion ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s)
    header_bar(s, "Выводы")
    footer(s, 10, TOTAL)

    conclusions = [
        "Современные дезсредства действуют через денатурацию белков, окисление и повреждение клеточных мембран микроорганизмов.",
        "Основные группы: хлор- и кислородсодержащие, спирты, альдегиды, ЧАС, гуанидины и комбинированные препараты.",
        "Они обеспечивают инфекционную безопасность пациентов и персонала, снижают риск ВБИ и поддерживают санэпидрежим.",
        "Организацию и контроль осуществляют медперсонал ЛПУ, эпидемиологи и органы Роспотребнадзора.",
        "Выбор средства зависит от объекта, уровня дезинфекции и режима; обязательно соблюдение инструкции производителя.",
    ]
    for i, text in enumerate(conclusions):
        top = Inches(1.3 + i * 1.05)
        add_round_rect(s, Inches(0.6), top, Inches(12.1), Inches(0.9), WHITE)
        circle_left = Inches(0.85)
        shape = s.shapes.add_shape(MSO_SHAPE.OVAL, circle_left, top + Inches(0.2),
                                   Inches(0.5), Inches(0.5))
        shape.fill.solid()
        shape.fill.fore_color.rgb = TEAL
        shape.line.fill.background()
        add_textbox(s, circle_left, top + Inches(0.28), Inches(0.5), Inches(0.4),
                    str(i + 1), size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(s, Inches(1.6), top + Inches(0.2), Inches(10.8), Inches(0.6),
                    text, size=14, color=DARK)

    # ========== SLIDE 11: Thank you ==========
    s = prs.slides.add_slide(blank)
    slide_bg(s, WHITE)
    add_rect(s, Inches(0), Inches(0), Inches(13.333), Inches(0.15), NAVY)
    add_rect(s, Inches(0), Inches(7.35), Inches(13.333), Inches(0.15), NAVY)
    add_rect(s, Inches(0), Inches(2.4), Inches(13.333), Inches(2.5), NAVY)
    add_textbox(s, Inches(0.5), Inches(2.7), Inches(12.3), Inches(0.8),
                "Спасибо за внимание!", size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0.5), Inches(3.6), Inches(12.3), Inches(0.5),
                "Готова ответить на ваши вопросы", size=18, color=RGBColor(0xA8, 0xD5, 0xD0),
                align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0.5), Inches(5.3), Inches(12.3), Inches(0.4),
                "Аметова Эмине Суюновна  ·  группа 261/1", size=14, bold=True,
                color=NAVY, align=PP_ALIGN.CENTER)
    add_textbox(s, Inches(0.5), Inches(5.75), Inches(12.3), Inches(0.35),
                "Медицинский колледж КФУ им. В.И. Вернадского", size=12,
                color=MUTED, align=PP_ALIGN.CENTER)

    out = "/workspace/Sovremennye_dezsredstva_Ametova.pptx"
    prs.save(out)
    print("Saved:", out)


if __name__ == "__main__":
    main()
