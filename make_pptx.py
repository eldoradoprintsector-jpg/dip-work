#!/usr/bin/env python3
"""
Генератор .pptx презентации для дипломной работы
Уктамова Мухаммаджона — профессиональный дизайн без сторонних библиотек.
"""

import zipfile, os

# ─── Цветовая палитра ────────────────────────────────────────────────────────
C_NAVY      = "0D1B2A"   # тёмно-синий фон
C_BLUE      = "1565C0"   # основной синий
C_CYAN      = "00B4D8"   # голубой акцент / линии
C_CYAN_LT   = "90E0EF"   # светлый голубой
C_WHITE     = "FFFFFF"
C_OFFWHITE  = "F0F4F8"   # фон контентных слайдов
C_GOLD      = "FFC300"   # золотой (для важных пунктов)
C_GRAY      = "546E7A"   # серый текст
C_DARK      = "1A237E"   # заголовки
C_RED       = "E53935"   # предупреждения

# Размер слайда 33.87 × 19.05 см (16:9, 1280×720 pt в EMU)
W = 33.87
H = 19.05

def emu(cm): return int(cm * 360000)

def _esc(s):
    return (str(s).replace("&","&amp;").replace("<","&lt;")
                  .replace(">","&gt;").replace('"',"&quot;"))

# ─── Низкоуровневые XML-примитивы ────────────────────────────────────────────

def para(text, sz, bold=False, color=C_WHITE, align="l", space_before=0, italic=False):
    b  = "1" if bold   else "0"
    it = "1" if italic else "0"
    al = {"l":"l","c":"ctr","r":"r"}.get(align,"l")
    sb = f'<a:spcBef><a:spcPts val="{space_before*100}"/></a:spcBef>' if space_before else '<a:spcBef><a:spcPts val="0"/></a:spcBef>'
    return (
        f'<a:p>'
        f'<a:pPr algn="{al}">{sb}</a:pPr>'
        f'<a:r>'
        f'<a:rPr lang="ru-RU" sz="{sz*100}" b="{b}" i="{it}" dirty="0" smtClean="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="+mj-lt"/>'
        f'</a:rPr>'
        f'<a:t>{_esc(text)}</a:t>'
        f'</a:r>'
        f'</a:p>'
    )

def para_empty(space_before=4):
    return f'<a:p><a:pPr><a:spcBef><a:spcPts val="{space_before*100}"/></a:spcBef></a:pPr><a:endParaRPr lang="ru-RU" dirty="0"/></a:p>'

def solid_fill(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'

def grad_fill(c1, c2, angle_deg=5400000):
    """Линейный градиент от c1 (0%) до c2 (100%)."""
    return (
        f'<a:gradFill>'
        f'<a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="{c1}"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="{c2}"/></a:gs>'
        f'</a:gsLst>'
        f'<a:lin ang="{angle_deg}" scaled="0"/>'
        f'</a:gradFill>'
    )

def shape(sid, name, x, y, w, h, fill_xml, txBody_xml="", line_xml="<a:ln><a:noFill/></a:ln>", geom="rect"):
    return (
        f'<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sid}" name="{_esc(name)}"/>'
        f'<p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>'
        + fill_xml + line_xml +
        f'</p:spPr>'
        + txBody_xml +
        f'</p:sp>'
    )

def txBody(paragraphs, anchor="t", wrap="square", autofit=True):
    af = '<a:normAutofit/>' if autofit else ''
    return (
        f'<p:txBody>'
        f'<a:bodyPr wrap="{wrap}" anchor="{anchor}">{af}</a:bodyPr>'
        f'<a:lstStyle/>'
        + "".join(paragraphs) +
        f'</p:txBody>'
    )

def no_fill():  return '<a:noFill/>'
def no_line():  return '<a:ln><a:noFill/></a:ln>'

def roundrect(sid, name, x, y, w, h, fill_xml, txBody_xml="", radius_pct=10):
    adj = f'<a:avLst><a:gd name="adj" fmla="val {radius_pct*1000}"/></a:avLst>'
    return (
        f'<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sid}" name="{_esc(name)}"/>'
        f'<p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
        f'<a:prstGeom prst="roundRect">{adj}</a:prstGeom>'
        + fill_xml +
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        + txBody_xml +
        f'</p:sp>'
    )

def line_shape(sid, x1, y1, x2, y2, color, width_pt=1):
    """Горизонтальная или вертикальная линия через <p:sp> с прозрачным текстом."""
    w = abs(x2-x1) if abs(x2-x1) > 0.01 else 0.02
    h = abs(y2-y1) if abs(y2-y1) > 0.01 else 0.02
    x = min(x1,x2); y = min(y1,y2)
    lw = width_pt * 12700
    fill = f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
    return shape(sid, f"Line{sid}", x, y, w, h,
                 fill_xml=fill,
                 txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>',
                 line_xml=no_line())

# ─── Номер слайда (нижний правый угол) ──────────────────────────────────────

def slide_number(sid, num, total, color=C_CYAN_LT):
    txt = txBody([para(f"{num} / {total}", 9, color=color, align="r")], anchor="ctr")
    return shape(sid, "SlideNum", W-3.5, H-0.65, 3.2, 0.5,
                 fill_xml=no_fill(), txBody_xml=txt, line_xml=no_line())

# ─── ШАБЛОНЫ СЛАЙДОВ ────────────────────────────────────────────────────────

def make_title_slide(data, num, total):
    parts = []

    # --- Фон: тёмно-синий градиент ---
    parts.append(shape(10,"BgGrad",0,0,W,H,
        fill_xml=grad_fill(C_NAVY,"1A3A5C",16200000), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Декоративный круг (большой, прозрачный — как "watermark")
    parts.append(shape(11,"Circle",18.5,-4.0,22,22,
        fill_xml=f'<a:solidFill><a:srgbClr val="{C_BLUE}"><a:alpha val="15000"/></a:srgbClr></a:solidFill>',
        line_xml=no_line(), geom="ellipse",
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Второй круг
    parts.append(shape(12,"Circle2",20.0,8.0,14,14,
        fill_xml=f'<a:solidFill><a:srgbClr val="{C_CYAN}"><a:alpha val="12000"/></a:srgbClr></a:solidFill>',
        line_xml=no_line(), geom="ellipse",
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Вертикальная цветная полоса слева
    parts.append(shape(13,"Bar",0,0,0.5,H,
        fill_xml=grad_fill(C_CYAN,C_BLUE,5400000), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Горизонтальная линия-разделитель
    parts.append(line_shape(14, 0.8, 12.5, 22.0, 12.5, C_CYAN))

    # Метка "ДИПЛОМНАЯ РАБОТА"
    txt_label = txBody([para("ДИПЛОМНАЯ РАБОТА", 9, bold=True,
                              color=C_CYAN, align="l")], anchor="ctr")
    parts.append(shape(15,"Label",0.9,11.4,12,0.6,
        fill_xml=no_fill(), txBody_xml=txt_label, line_xml=no_line()))

    # Главный заголовок
    title_lines = data["title"].split("\n")
    title_paras = []
    for i, line in enumerate(title_lines):
        sz = 30 if i == 0 else 26
        title_paras.append(para(line, sz, bold=True, color=C_WHITE, align="l"))
    txt_title = txBody(title_paras, anchor="t")
    parts.append(shape(16,"Title",0.9,3.0,21,9.0,
        fill_xml=no_fill(), txBody_xml=txt_title, line_xml=no_line()))

    # Подзаголовок / детали
    sub_lines = data["subtitle"].split("\n")
    sub_paras = []
    for line in sub_lines:
        if not line.strip():
            sub_paras.append(para_empty(3))
        else:
            sub_paras.append(para(line, 13, color=C_CYAN_LT, align="l"))
    txt_sub = txBody(sub_paras, anchor="t")
    parts.append(shape(17,"Sub",0.9,12.8,22,5.5,
        fill_xml=no_fill(), txBody_xml=txt_sub, line_xml=no_line()))

    # Нижняя полоска с университетом
    parts.append(shape(18,"Footer",0,H-1.1,W,1.1,
        fill_xml=f'<a:solidFill><a:srgbClr val="{C_BLUE}"><a:alpha val="80000"/></a:srgbClr></a:solidFill>',
        line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    parts.append(slide_number(19, num, total, C_CYAN_LT))
    return parts


def make_section_slide(data, num, total):
    parts = []

    # Фон — тёмный градиент по горизонтали
    parts.append(shape(10,"Bg",0,0,W,H,
        fill_xml=grad_fill("0D1B2A","1A3A5C",0), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Яркий прямоугольник-акцент слева
    parts.append(shape(11,"Accent",0,0,0.7,H,
        fill_xml=grad_fill(C_CYAN,C_BLUE,5400000), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Крупная надпись главы
    txt_ch = txBody([para(data["title"], 56, bold=True, color=C_CYAN, align="c")],
                    anchor="ctr")
    parts.append(shape(12,"Chapter",1.2,1.5,W-1.7,5.0,
        fill_xml=no_fill(), txBody_xml=txt_ch, line_xml=no_line()))

    # Горизонтальная линия под главой
    parts.append(line_shape(13, 3.0, 7.2, W-3.0, 7.2, C_CYAN))

    # Подзаголовок
    sub_lines = data["subtitle"].split("\n")
    sub_paras = [para(l, 20, color=C_WHITE, align="c") for l in sub_lines]
    txt_sub = txBody(sub_paras, anchor="t")
    parts.append(shape(14,"Sub",1.2,7.6,W-1.7,8.5,
        fill_xml=no_fill(), txBody_xml=txt_sub, line_xml=no_line()))

    parts.append(slide_number(15, num, total, C_CYAN_LT))
    return parts


def make_content_slide(data, num, total):
    parts = []

    # --- Светлый фон ---
    parts.append(shape(10,"Bg",0,0,W,H,
        fill_xml=solid_fill(C_OFFWHITE), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Шапка: тёмная + градиент
    parts.append(shape(11,"Header",0,0,W,3.3,
        fill_xml=grad_fill(C_NAVY,C_BLUE,0), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Голубая полоска под шапкой
    parts.append(line_shape(12, 0,3.25, W,3.25, C_CYAN))

    # Вертикальный акцент слева (в области контента)
    parts.append(shape(13,"Vbar",0,3.3,0.18,H-3.3,
        fill_xml=grad_fill(C_CYAN,C_OFFWHITE,5400000), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Заголовок
    txt_title = txBody([para(data["title"], 21, bold=True, color=C_WHITE, align="l")],
                       anchor="ctr")
    parts.append(shape(14,"Title",0.7,0.2,W-1.2,3.0,
        fill_xml=no_fill(), txBody_xml=txt_title, line_xml=no_line()))

    # Буллеты
    bullet_paras = []
    for i, bullet in enumerate(data["bullets"]):
        is_sub   = bullet.startswith("  ")
        is_warn  = "⚠" in bullet
        is_check = bullet.startswith("✓")

        raw = bullet.strip()

        if is_warn:
            color = C_RED
            prefix = "⚠  "
            sz = 14
        elif is_check:
            color = C_BLUE
            prefix = "✓  "
            sz = 15
        elif is_sub:
            color = C_GRAY
            prefix = "     "
            sz = 13
        else:
            color = C_DARK
            prefix = "▶  "
            sz = 15

        sp_before = 6 if (not is_sub and i > 0) else 1
        bullet_paras.append(
            para(prefix + raw, sz, color=color, align="l", space_before=sp_before)
        )

    txt_bullets = txBody(bullet_paras, anchor="t")
    parts.append(shape(15,"Bullets",0.6,3.5,W-1.0,H-3.9,
        fill_xml=no_fill(), txBody_xml=txt_bullets, line_xml=no_line()))

    # Номер слайда
    parts.append(slide_number(16, num, total))
    return parts


def make_two_col_slide(data, num, total):
    parts = []

    # Фон
    parts.append(shape(10,"Bg",0,0,W,H,
        fill_xml=solid_fill(C_OFFWHITE), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

    # Шапка
    parts.append(shape(11,"Header",0,0,W,3.3,
        fill_xml=grad_fill(C_NAVY,C_BLUE,0), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))
    parts.append(line_shape(12, 0,3.25, W,3.25, C_CYAN))

    # Заголовок
    txt_title = txBody([para(data["title"], 21, bold=True, color=C_WHITE, align="l")],
                       anchor="ctr")
    parts.append(shape(13,"Title",0.7,0.2,W-1.2,3.0,
        fill_xml=no_fill(), txBody_xml=txt_title, line_xml=no_line()))

    # Карточки-колонки
    COL_W = 15.1
    GAP   = 0.5
    COL_H = H - 3.8
    Y     = 3.6

    for col_i, (side, col_x) in enumerate([("left",0.4),("right",17.4)]):
        title_key   = f"{side}_title"
        bullets_key = f"{side}_bullets"

        # Белая карточка
        card_fill = f'<a:solidFill><a:srgbClr val="{C_WHITE}"/></a:solidFill>'
        card_line = f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{C_CYAN_LT}"/></a:solidFill></a:ln>'
        parts.append(shape(20+col_i*10,"Card",col_x,Y,COL_W,COL_H,
            fill_xml=card_fill, txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>',
            line_xml=card_line))

        # Цветная шапка карточки
        parts.append(shape(21+col_i*10,"CardHdr",col_x,Y,COL_W,1.1,
            fill_xml=grad_fill(C_BLUE,C_CYAN,0) if col_i==0 else grad_fill(C_CYAN,C_BLUE,0),
            line_xml=no_line(),
            txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))

        # Заголовок карточки
        ct = txBody([para(data[title_key], 14, bold=True, color=C_WHITE, align="c")],
                    anchor="ctr")
        parts.append(shape(22+col_i*10,"CardTitle",col_x,Y,COL_W,1.1,
            fill_xml=no_fill(), txBody_xml=ct, line_xml=no_line()))

        # Буллеты карточки
        bp = []
        for b in data[bullets_key]:
            bp.append(para("▶  " + b, 13, color=C_DARK, space_before=5))
        bt = txBody(bp, anchor="t")
        parts.append(shape(23+col_i*10,"CardBullets",col_x+0.2,Y+1.3,COL_W-0.4,COL_H-1.5,
            fill_xml=no_fill(), txBody_xml=bt, line_xml=no_line()))

    parts.append(slide_number(50, num, total))
    return parts


def make_stats_slide(data, num, total):
    """Слайд с крупными цифрами / ключевыми показателями."""
    parts = []

    parts.append(shape(10,"Bg",0,0,W,H,
        fill_xml=solid_fill(C_OFFWHITE), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))
    parts.append(shape(11,"Header",0,0,W,3.3,
        fill_xml=grad_fill(C_NAVY,C_BLUE,0), line_xml=no_line(),
        txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'))
    parts.append(line_shape(12, 0,3.25, W,3.25, C_CYAN))

    txt_title = txBody([para(data["title"], 21, bold=True, color=C_WHITE, align="l")],
                       anchor="ctr")
    parts.append(shape(13,"Title",0.7,0.2,W-1.2,3.0,
        fill_xml=no_fill(), txBody_xml=txt_title, line_xml=no_line()))

    stats = data["stats"]  # list of {"value":..., "label":...}
    n = len(stats)
    card_w = (W - 0.8 - (n-1)*0.4) / n
    y_card = 3.9

    for i, st in enumerate(stats):
        cx = 0.4 + i*(card_w + 0.4)
        # Карточка с градиентом
        cf = grad_fill(C_BLUE, C_NAVY, 5400000) if i % 2 == 0 else grad_fill(C_NAVY, C_BLUE, 5400000)
        parts.append(roundrect(20+i,"StatCard",cx,y_card,card_w,6.5,
            fill_xml=cf,
            txBody_xml='<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>',
            radius_pct=8))

        # Значение
        vt = txBody([para(st["value"], 36, bold=True, color=C_GOLD, align="c")],
                    anchor="ctr")
        parts.append(shape(30+i,"Val",cx,y_card+0.6,card_w,3.0,
            fill_xml=no_fill(), txBody_xml=vt, line_xml=no_line()))

        # Подпись
        lt_paras = []
        for line in st["label"].split("\n"):
            lt_paras.append(para(line, 11, color=C_WHITE, align="c"))
        lt = txBody(lt_paras, anchor="t")
        parts.append(shape(40+i,"Label",cx+0.15,y_card+3.5,card_w-0.3,3.0,
            fill_xml=no_fill(), txBody_xml=lt, line_xml=no_line()))

    # Подпись-источник
    if "source" in data:
        st = txBody([para(data["source"], 9, color=C_GRAY, italic=True, align="r")],
                    anchor="ctr")
        parts.append(shape(60,"Src",0.5,H-0.8,W-1.0,0.6,
            fill_xml=no_fill(), txBody_xml=st, line_xml=no_line()))

    parts.append(slide_number(61, num, total))
    return parts


# ─── Данные слайдов ─────────────────────────────────────────────────────────

slides_data = [
    # 0 — Титульный
    {
        "type": "title",
        "title": "Цифровизация экономики\nи устойчивость стран\nк глобальным кризисам",
        "subtitle": (
            "Выпускная квалификационная работа\n"
            "Уктамов Мухаммаджон\n\n"
            "Направление: Экономика\n"
            "Год защиты: 2025"
        ),
    },
    # 1 — Актуальность
    {
        "type": "content",
        "title": "Актуальность исследования",
        "bullets": [
            "Три волны глобальных кризисов: 2008–2009, 2020–2021, 2022 — одинаковый удар, разная реакция",
            "Цифровизация как структурный фактор устойчивости экономики",
            "COVID-19: цифровая инфраструктура обеспечила непрерывность государства и бизнеса",
            "Энергетический шок 2022: выявил пределы цифровой зрелости",
            "Узбекистан реализует Стратегию «Цифровой Узбекистан – 2030»",
            "⚠ Вопрос: влияет ли уровень цифровизации на устойчивость? Как? Через какие каналы?",
        ],
    },
    # 2 — Объект / предмет / цель
    {
        "type": "two_col",
        "title": "Объект, предмет и цель исследования",
        "left_title": "Объект и предмет",
        "left_bullets": [
            "Объект: цифровая трансформация национальных экономик",
            "Предмет: механизмы влияния цифровых технологий на макроэкономическую устойчивость",
            "Период: глобальные кризисы 2020–2022 годов",
        ],
        "right_title": "Цель и задачи",
        "right_bullets": [
            "Цель: выявить и эмпирически обосновать каналы влияния цифровизации на устойчивость",
            "Систематизировать теоретические подходы",
            "Сравнительный анализ 4 стран-лидеров",
            "Эконометрическая оценка на 49 странах",
            "Рекомендации для Узбекистана",
        ],
    },
    # 3 — Пять каналов
    {
        "type": "content",
        "title": "Теоретическая рамка: 5 каналов влияния цифровизации на устойчивость",
        "bullets": [
            "Канал 1 — Непрерывность государственных функций в условиях шоков",
            "  e-Government, электронная идентификация, межведомственный обмен данными",
            "Канал 2 — Эффективность фискальной трансмиссии",
            "  Цифровые платёжные системы → адресные выплаты в реальном времени",
            "Канал 3 — Адаптация частного сектора",
            "  Фирмы с высоким уровнем ИКТ показывают меньшее падение выручки в кризис",
            "Канал 4 — Структурный сдвиг спроса к цифровым услугам и e-commerce",
            "Канал 5 — Устойчивость ИКТ-сектора как отдельного компонента ВВП",
            "  ИКТ-сектор ОЭСР: рост даже в 2020 году (+ХХ% против спада в экономике)",
        ],
    },
    # 4 — Ключевые показатели сравнения (stats-слайд)
    {
        "type": "stats",
        "title": "Четыре модели цифровой трансформации: ключевые показатели",
        "stats": [
            {"value": "🇪🇪 #3",  "label": "Эстония\nEGDI ООН 2024\nX-Road, e-Residency"},
            {"value": "🇰🇷 #2",  "label": "Республика Корея\nNRI 2024\nDigital New Deal"},
            {"value": "🇸🇬 #1",  "label": "Сингапур\nIMD WDCR 2024\nSmart Nation 2.0"},
            {"value": "🇵🇱 #24", "label": "Польша\nDESI 2024\nДогоняющая модель"},
        ],
        "source": "Источник: UN DESA EGDI 2024; NRI 2024; IMD WDCR 2024; EC Digital Decade 2024",
    },
    # 5 — Реакция на кризисы 2020
    {
        "type": "stats",
        "title": "Реакция экономик на пандемийный шок 2020 года (% изменения ВВП)",
        "stats": [
            {"value": "−0.7%", "label": "Республика Корея\nЛучший результат"},
            {"value": "−2.0%", "label": "Польша\nМягкое падение"},
            {"value": "−3.0%", "label": "Эстония\nНиже среднего ЕС"},
            {"value": "−3.9%", "label": "Сингапур\nТоргово-зависимая"},
            {"value": "−6.1%", "label": "Среднее по ЕС\nЭталон сравнения"},
        ],
        "source": "Источник: World Bank WDI; Eurostat; Bank of Korea, 2024",
    },
    # 6 — Эконометрика (section)
    {
        "type": "section",
        "title": "Глава II",
        "subtitle": "Эконометрический анализ\n49 стран · 707 наблюдений · 2010–2024",
    },
    # 7 — Методология
    {
        "type": "content",
        "title": "Эконометрическая модель: методология",
        "bullets": [
            "Выборка: 49 стран, несбалансированная панель 2010–2024 гг., 707 наблюдений",
            "  Развитые ОЭСР (22) · Переходные (16) · Развивающиеся вкл. Узбекистан (11)",
            "Зависимая переменная: ln(ВВП на душу населения, постоянные цены 2015)",
            "Основной регрессор: ln(доля интернет-пользователей, %)",
            "Контрольные переменные: торговая открытость, качество институтов (WGI), норма накопления капитала",
            "Метод: FE-модели + робастные стандартные ошибки (кластеризация по стране)",
            "Тест Чоу: F = 5,34; p = 1,15×10⁻¹¹ → структурная неоднородность подтверждена",
        ],
    },
    # 8 — Результаты эконометрики (stats)
    {
        "type": "stats",
        "title": "Ключевые результаты: коэффициенты FE-моделей по группам стран",
        "stats": [
            {"value": "β=0.197",  "label": "Развитые экономики\n(p<0.001)\n+COVID: β=0.394***\n+Геополит.: β=0.289**"},
            {"value": "β=0.309",  "label": "Переходные экономики\n(p<0.001)\nНаибольший\nбазовый эффект"},
            {"value": "β=0.198",  "label": "Развивающиеся\nвкл. Узбекистан\n(p<0.001)\nКризисные члены\nнезначимы"},
        ],
        "source": "Гипотеза H₁ подтверждена во всех трёх группах стран. Источник: расчёты автора, R Studio.",
    },
    # 9 — Узбекистан (section)
    {
        "type": "section",
        "title": "Глава III",
        "subtitle": "Цифровая трансформация Узбекистана\nи практические рекомендации",
    },
    # 10 — Состояние цифровизации Узбекистана
    {
        "type": "content",
        "title": "Цифровая трансформация Узбекистана: текущее состояние",
        "bullets": [
            "Стратегия «Цифровой Узбекистан – 2030» (Указ №УП-6079, октябрь 2020)",
            "E-Government индекс ООН 2024: значительный рост в рейтинге",
            "IT Park Узбекистана: 2 500+ компаний-резидентов к концу 2024 г.",
            "Рост ИКТ-экспорта, развитие электронной коммерции и цифровых платежей",
            "⚠ Асимметрия: цифровое государство опережает цифровой бизнес и население",
            "⚠ Региональный цифровой разрыв: городские vs. сельские районы",
            "⚠ Уровень интернет-проникновения ~80% — необходимо достижение 90%+",
        ],
    },
    # 11 — Рекомендации
    {
        "type": "content",
        "title": "Практические рекомендации: 5 приоритетов",
        "bullets": [
            "1. Расширить интернет-проникновение до 90%+ → прирост ВВП: +84–124 USD на чел. в год",
            "2. Устранить асимметрию цифровой зрелости: цифровой бизнес и цифровые навыки населения",
            "3. Инфраструктура фискальной трансмиссии: интеграция MyID с системой соцвыплат",
            "4. Учитывать тип кризиса: пандемийный шок vs. энергетический/инфляционный шок",
            "5. Накапливать качественные данные → достичь статистической значимости кризисных эффектов",
        ],
    },
    # 12 — Дорожная карта (two_col)
    {
        "type": "two_col",
        "title": "Дорожная карта цифровой устойчивости Узбекистана до 2030",
        "left_title": "Этап 1: 2025–2027",
        "left_bullets": [
            "Инфраструктура фискальной трансмиссии",
            "Интеграция MyID с адресными выплатами",
            "Расширение широкополосного доступа в регионах",
            "Цифровые навыки: программы для населения и МСП",
        ],
        "right_title": "Этап 2: 2027–2030+",
        "right_bullets": [
            "Устранение разрыва: бизнес-сектор ≈ государственный",
            "Достижение «порога» антикризисной зрелости",
            "Значимые interaction terms в будущих моделях",
            "Адаптация опыта X-Road (Эстония) к масштабу Узбекистана",
        ],
    },
    # 13 — Научная новизна (two_col)
    {
        "type": "two_col",
        "title": "Научная новизна и практическая значимость",
        "left_title": "Научная новизна",
        "left_bullets": [
            "Систематизирована пятиканальная модель влияния цифровизации на устойчивость",
            "Эмпирически подтверждена структурная неоднородность эффекта (тест Чоу)",
            "Выявлен «порог» цифровой зрелости для значимого антикризисного эффекта",
        ],
        "right_title": "Практическая значимость",
        "right_bullets": [
            "Рекомендации для уточнения Стратегии «Цифровой Узбекистан – 2030»",
            "Методика адаптации опыта Эстонии, Кореи, Польши к условиям Узбекистана",
            "Прогнозные сценарии ВВП на душу населения до 2030 года",
        ],
    },
    # 14 — Заключение
    {
        "type": "title",
        "title": "Основные выводы",
        "subtitle": (
            "✓ H₁ подтверждена: цифровизация значимо связана с ВВП во всех трёх группах (β=0.20–0.31)\n\n"
            "✓ Антикризисный эффект значим только при высокой цифровой зрелости\n\n"
            "✓ Для Узбекистана: последовательная реализация Стратегии «Цифровой Узбекистан – 2030» —\n"
            "   необходимое условие перехода от роста к устойчивости\n\n"
            "Спасибо за внимание!"
        ),
    },
]

# ─── Служебные XML-файлы .pptx ──────────────────────────────────────────────

NS_SLD = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

def slide_xml(shapes_list):
    shapes = "".join(shapes_list)
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld {NS_SLD}>'
        f'<p:cSld><p:spTree>'
        f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + shapes +
        f'</p:spTree></p:cSld>'
        f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        f'</p:sld>'
    )

def content_types_xml(n):
    parts = "".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, n+1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml"  ContentType="application/xml"/>'
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        + parts + '</Types>'
    )

def root_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
        '</Relationships>'
    )

def presentation_xml(n):
    ids = "".join(f'<p:sldId id="{256+i}" r:id="rId{i}"/>' for i in range(1,n+1))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{ids}</p:sldIdLst>'
        '<p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        '</p:presentation>'
    )

def presentation_rels_xml(n):
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>'
        for i in range(1,n+1)
    )
    rels += '<Relationship Id="rIdM1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + rels + '</Relationships>'
    )

def slide_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>'
    )

def slide_layout_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'type="blank" preserve="1">'
        '<p:cSld name="Blank"><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sldLayout>'
    )

def slide_layout_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        '</Relationships>'
    )

def slide_master_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="acc1" '
        'accent2="acc2" accent3="acc3" accent4="acc4" accent5="acc5" accent6="acc6" '
        'hlink="hlink" folHlink="folHlink"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        '<p:txStyles>'
        '<p:titleStyle><a:lstStyle/></p:titleStyle>'
        '<p:bodyStyle><a:lstStyle/></p:bodyStyle>'
        '<p:otherStyle><a:lstStyle/></p:otherStyle>'
        '</p:txStyles>'
        '</p:sldMaster>'
    )

def slide_master_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>'
    )

# ─── Диспетчер слайдов ──────────────────────────────────────────────────────

def build_slide(data, num, total):
    t = data["type"]
    if   t == "title":   parts = make_title_slide(data, num, total)
    elif t == "section": parts = make_section_slide(data, num, total)
    elif t == "content": parts = make_content_slide(data, num, total)
    elif t == "two_col": parts = make_two_col_slide(data, num, total)
    elif t == "stats":   parts = make_stats_slide(data, num, total)
    else:                parts = make_content_slide(data, num, total)
    return slide_xml(parts)

# ─── Сборка .pptx ───────────────────────────────────────────────────────────

def create_pptx(filename):
    n = len(slides_data)
    with zipfile.ZipFile(filename, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",             content_types_xml(n))
        z.writestr("_rels/.rels",                     root_rels_xml())
        z.writestr("ppt/presentation.xml",            presentation_xml(n))
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels_xml(n))
        for i, data in enumerate(slides_data, start=1):
            z.writestr(f"ppt/slides/slide{i}.xml",           build_slide(data, i, n))
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", slide_rels_xml())
        z.writestr("ppt/slideLayouts/slideLayout1.xml",              slide_layout_xml())
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",   slide_layout_rels_xml())
        z.writestr("ppt/slideMasters/slideMaster1.xml",              slide_master_xml())
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels",   slide_master_rels_xml())
    size_kb = os.path.getsize(filename) // 1024
    print(f"✅  {filename}")
    print(f"    {n} слайдов · {size_kb} КБ")

if __name__ == "__main__":
    create_pptx("Презентация_Уктамов_Мухаммаджон.pptx")
