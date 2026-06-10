#!/usr/bin/env python3
"""
Генератор .pptx — дипломная работа Уктамова Мухаммаджона.
Использует только встроенные библиотеки Python (zipfile).
Протестирован на совместимость с PowerPoint и LibreOffice.
"""
import zipfile, os

# ══════════════════════════════════════════════════════════════════
#  КОНСТАНТЫ
# ══════════════════════════════════════════════════════════════════

# Размер слайда 16:9 — 12 192 000 × 6 858 000 EMU
CX = 12192000
CY = 6858000

# 1 см = 360 000 EMU
def cm(v): return int(v * 360000)

# Цвета (без #)
NAVY   = "0D1B2A"
BLUE   = "1565C0"
CYAN   = "00B4D8"
LTCYAN = "90E0EF"
WHITE  = "FFFFFF"
CREAM  = "F0F4F8"
GOLD   = "FFC300"
GRAY   = "546E7A"
DARK   = "1A237E"
RED    = "E53935"

def esc(s):
    return (str(s)
            .replace("&","&amp;")
            .replace("<","&lt;")
            .replace(">","&gt;")
            .replace('"',"&quot;"))

# ══════════════════════════════════════════════════════════════════
#  XML-ПРИМИТИВЫ
# ══════════════════════════════════════════════════════════════════

def solidFill(clr):
    return f'<a:solidFill><a:srgbClr val="{clr}"/></a:solidFill>'

def gradFill(c1, c2, ang=5400000):
    """Линейный градиент (ang=0 → слева направо, 5400000 → сверху вниз)."""
    return (
        f'<a:gradFill rotWithShape="1">'
        f'<a:gsLst>'
        f'<a:gs pos="0">{solidFill(c1)}</a:gs>'
        f'<a:gs pos="100000">{solidFill(c2)}</a:gs>'
        f'</a:gsLst>'
        f'<a:lin ang="{ang}" scaled="0"/>'
        f'</a:gradFill>'
    )

def noFill():  return '<a:noFill/>'
def noLine():  return '<a:ln><a:noFill/></a:ln>'

def xfrm(x, y, cx, cy, flipH=False, flipV=False):
    fh = ' flipH="1"' if flipH else ''
    fv = ' flipV="1"' if flipV else ''
    return (
        f'<a:xfrm{fh}{fv}>'
        f'<a:off x="{x}" y="{y}"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/>'
        f'</a:xfrm>'
    )

def prstGeom(prst="rect"):
    return f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'

def roundRectGeom(adj_pct=8):
    return (
        f'<a:prstGeom prst="roundRect">'
        f'<a:avLst><a:gd name="adj" fmla="val {adj_pct*1000}"/></a:avLst>'
        f'</a:prstGeom>'
    )

# ── Абзац ─────────────────────────────────────────────────────────

def run(text, sz_pt, bold=False, color=WHITE, italic=False, typeface="Calibri"):
    b  = "1" if bold   else "0"
    it = "1" if italic else "0"
    return (
        f'<a:r>'
        f'<a:rPr lang="ru-RU" altLang="en-US" sz="{sz_pt*100}" b="{b}" i="{it}"'
        f' dirty="0" smtClean="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{typeface}"/>'
        f'</a:rPr>'
        f'<a:t>{esc(text)}</a:t>'
        f'</a:r>'
    )

def para(text, sz_pt, bold=False, color=WHITE, align="l",
         space_before=0, italic=False, indent=0):
    al = {"l":"l","c":"ctr","r":"r"}.get(align,"l")
    spc = f'<a:spcBef><a:spcPts val="{space_before*100}"/></a:spcBef>' if space_before else ''
    ind = f' indent="-{cm(0.3)}" marL="{cm(0.5+indent*0.3)}"' if indent else ''
    return (
        f'<a:p>'
        f'<a:pPr algn="{al}"{ind}>{spc}</a:pPr>'
        + run(text, sz_pt, bold, color, italic) +
        f'</a:p>'
    )

def paraEmpty(space_before=4):
    return (
        f'<a:p><a:pPr>'
        f'<a:spcBef><a:spcPts val="{space_before*100}"/></a:spcBef>'
        f'</a:pPr>'
        f'<a:endParaRPr lang="ru-RU" dirty="0"/>'
        f'</a:p>'
    )

def txBody(paras, anchor="t", wrap="square"):
    return (
        f'<p:txBody>'
        f'<a:bodyPr wrap="{wrap}" anchor="{anchor}">'
        f'<a:normAutofit/>'
        f'</a:bodyPr>'
        f'<a:lstStyle/>'
        + "".join(paras) +
        f'</p:txBody>'
    )

# ── Фигура <p:sp> ─────────────────────────────────────────────────

_ID = [1]
def _nid():
    _ID[0] += 1
    return _ID[0]

def sp(x, y, cx, cy, fill, line=None, geom="rect", tx=""):
    """Универсальная фигура."""
    sid = _nid()
    ln  = line if line else noLine()
    return (
        f'<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sid}" name="sp{sid}"/>'
        f'<p:cNvSpPr txBox="1"/>'
        f'<p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr bwMode="auto">'
        + xfrm(x, y, cx, cy) +
        (prstGeom(geom) if geom != "roundRect" else roundRectGeom()) +
        fill + ln +
        f'</p:spPr>'
        + tx +
        f'</p:sp>'
    )

def rect(x, y, cx, cy, fill, line=None):
    return sp(x, y, cx, cy, fill, line, geom="rect")

def ellipse(x, y, cx, cy, fill):
    return sp(x, y, cx, cy, fill, line=noLine(), geom="ellipse")

def textbox(x, y, cx, cy, paras, anchor="t", color_fill=None):
    fill = solidFill(color_fill) if color_fill else noFill()
    tx   = txBody(paras, anchor)
    return sp(x, y, cx, cy, fill, line=noLine(), tx=tx)

def roundbox(x, y, cx, cy, fill, paras=None, anchor="t"):
    """Закруглённый прямоугольник."""
    tx = txBody(paras or [], anchor) if paras is not None else '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
    return sp(x, y, cx, cy, fill, line=noLine(), geom="roundRect", tx=tx)

# ── Слайд-обёртка ─────────────────────────────────────────────────

NS = (
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

def makeSlide(shapes):
    _ID[0] = 1  # сбрасываем счётчик для каждого нового слайда
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<p:sld {NS} show="1">'
        f'<p:cSld>'
        f'<p:spTree>'
        f'<p:nvGrpSpPr>'
        f'<p:cNvPr id="1" name=""/>'
        f'<p:cNvGrpSpPr/>'
        f'<p:nvPr/>'
        f'</p:nvGrpSpPr>'
        f'<p:grpSpPr>'
        f'<a:xfrm>'
        f'<a:off x="0" y="0"/><a:ext cx="{CX}" cy="{CY}"/>'
        f'<a:chOff x="0" y="0"/><a:chExt cx="{CX}" cy="{CY}"/>'
        f'</a:xfrm>'
        f'</p:grpSpPr>'
        + "".join(shapes) +
        f'</p:spTree>'
        f'</p:cSld>'
        f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        f'</p:sld>'
    )

# ══════════════════════════════════════════════════════════════════
#  ПЕРЕИСПОЛЬЗУЕМЫЕ ЭЛЕМЕНТЫ ДИЗАЙНА
# ══════════════════════════════════════════════════════════════════

# Пустой txBody-блок (для фигур только как фон)
EMPTY_TX = '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'

def bgRect(fill):
    """Фон на весь слайд."""
    return sp(0, 0, CX, CY, fill, line=noLine(),
              tx=EMPTY_TX)

def header(fill=None):
    """Тёмная шапка высотой 3.2 см."""
    f = fill or gradFill(NAVY, BLUE, ang=0)
    return sp(0, 0, CX, cm(3.2), f, line=noLine(), tx=EMPTY_TX)

def accentLine(y_cm=3.15):
    """Голубая тонкая линия."""
    return sp(0, cm(y_cm), CX, cm(0.12), solidFill(CYAN), line=noLine(), tx=EMPTY_TX)

def leftBar():
    """Вертикальная акцент-полоса слева."""
    return sp(0, cm(3.2), cm(0.22), CY - cm(3.2),
              gradFill(CYAN, CREAM, ang=5400000), line=noLine(), tx=EMPTY_TX)

def slideNum(num, total):
    """Номер слайда в правом нижнем углу."""
    return textbox(
        CX - cm(3.5), CY - cm(0.7), cm(3.2), cm(0.55),
        [para(f"{num} / {total}", 9, color=GRAY, align="r")]
    )

def titleBox(title_text, x=cm(0.7), y=cm(0.15), cx=CX - cm(1.2), cy=cm(2.95)):
    """Заголовок слайда в шапке."""
    return textbox(x, y, cx, cy,
                   [para(title_text, 22, bold=True, color=WHITE, align="l")],
                   anchor="ctr")

# ══════════════════════════════════════════════════════════════════
#  ПОСТРОИТЕЛИ СЛАЙДОВ
# ══════════════════════════════════════════════════════════════════

# ── Титульный слайд ───────────────────────────────────────────────
def slideTITLE(d, num, total):
    s = []
    # Фон-градиент
    s.append(bgRect(gradFill(NAVY, "1A3A5C", ang=0)))
    # Большие декоративные круги
    s.append(sp(cm(20), cm(-4), cm(22), cm(22),
                f'<a:solidFill><a:srgbClr val="{BLUE}"><a:alpha val="18000"/></a:srgbClr></a:solidFill>',
                line=noLine(), geom="ellipse", tx=EMPTY_TX))
    s.append(sp(cm(22), cm(9), cm(13), cm(13),
                f'<a:solidFill><a:srgbClr val="{CYAN}"><a:alpha val="14000"/></a:srgbClr></a:solidFill>',
                line=noLine(), geom="ellipse", tx=EMPTY_TX))
    # Левая цветная полоса
    s.append(sp(0, 0, cm(0.6), CY, gradFill(CYAN, BLUE, ang=5400000),
                line=noLine(), tx=EMPTY_TX))
    # Метка
    s.append(textbox(cm(0.9), cm(11.0), cm(16), cm(0.7),
                     [para("ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА", 9, bold=True,
                           color=CYAN, align="l")]))
    # Горизонтальная линия
    s.append(sp(cm(0.9), cm(11.85), cm(20), cm(0.1),
                solidFill(CYAN), line=noLine(), tx=EMPTY_TX))
    # Заголовок
    title_paras = []
    for i, line in enumerate(d["title"].split("\n")):
        sz = 30 if i == 0 else 25
        sp_before = 0 if i == 0 else 4
        title_paras.append(para(line, sz, bold=True, color=WHITE,
                                align="l", space_before=sp_before))
    s.append(textbox(cm(0.9), cm(2.5), cm(20), cm(9), title_paras, anchor="t"))
    # Подзаголовок
    sub_paras = []
    for line in d["subtitle"].split("\n"):
        if line.strip() == "":
            sub_paras.append(paraEmpty(3))
        else:
            sub_paras.append(para(line, 13, color=LTCYAN, align="l"))
    s.append(textbox(cm(0.9), cm(12.1), cm(22), cm(5.5), sub_paras, anchor="t"))
    # Нижняя плашка
    s.append(sp(0, CY - cm(1.15), CX, cm(1.15),
                f'<a:solidFill><a:srgbClr val="{BLUE}"><a:alpha val="90000"/></a:srgbClr></a:solidFill>',
                line=noLine(), tx=EMPTY_TX))
    s.append(slideNum(num, total))
    return makeSlide(s)


# ── Слайд-разделитель главы ───────────────────────────────────────
def slideSECTION(d, num, total):
    s = []
    s.append(bgRect(gradFill("0D1B2A", "1A3A5C", ang=0)))
    # Акцентная полоса слева
    s.append(sp(0, 0, cm(0.8), CY, gradFill(CYAN, BLUE, ang=5400000),
                line=noLine(), tx=EMPTY_TX))
    # Заголовок главы (крупно)
    s.append(textbox(cm(1.2), cm(1.5), CX - cm(1.7), cm(5),
                     [para(d["title"], 52, bold=True, color=CYAN, align="c")],
                     anchor="ctr"))
    # Линия
    s.append(sp(cm(3), cm(7.0), CX - cm(6), cm(0.1),
                solidFill(CYAN), line=noLine(), tx=EMPTY_TX))
    # Подзаголовок
    sub_paras = [para(l, 20, color=WHITE, align="c")
                 for l in d["subtitle"].split("\n")]
    s.append(textbox(cm(1.2), cm(7.4), CX - cm(1.7), CY - cm(8.0),
                     sub_paras, anchor="t"))
    s.append(slideNum(num, total))
    return makeSlide(s)


# ── Контентный слайд (буллеты) ────────────────────────────────────
def slideCONTENT(d, num, total):
    s = []
    s.append(bgRect(solidFill(CREAM)))
    s.append(header())
    s.append(accentLine())
    s.append(leftBar())
    s.append(titleBox(d["title"]))

    bullet_paras = []
    for i, b in enumerate(d["bullets"]):
        is_sub  = b.startswith("  ")
        is_warn = "⚠" in b
        raw = b.strip()

        if is_warn:
            clr = RED; sz = 14; prefix = "⚠  "
            sp_b = 5 if i > 0 else 0
        elif is_sub:
            clr = GRAY; sz = 13; prefix = "      "
            sp_b = 1
        else:
            clr = DARK; sz = 15; prefix = "▶  "
            sp_b = 7 if i > 0 else 0

        bullet_paras.append(
            para(prefix + raw, sz, color=clr,
                 align="l", space_before=sp_b)
        )

    s.append(textbox(cm(0.55), cm(3.45), CX - cm(0.9),
                     CY - cm(3.9), bullet_paras, anchor="t"))
    s.append(slideNum(num, total))
    return makeSlide(s)


# ── Двухколоночный слайд ──────────────────────────────────────────
def slideTWOCOL(d, num, total):
    s = []
    s.append(bgRect(solidFill(CREAM)))
    s.append(header())
    s.append(accentLine())
    s.append(titleBox(d["title"]))

    COL_W = cm(15.1)
    COL_H = CY - cm(3.9)
    COL_Y = cm(3.55)

    for side, col_x in [("left", cm(0.4)), ("right", cm(17.5))]:
        tk = f"{side}_title"
        bk = f"{side}_bullets"

        # Белая карточка
        card_fill = solidFill(WHITE)
        card_line = f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{LTCYAN}"/></a:solidFill></a:ln>'
        s.append(sp(col_x, COL_Y, COL_W, COL_H, card_fill, line=card_line,
                    geom="roundRect", tx=EMPTY_TX))

        # Шапка карточки
        card_hdr_fill = gradFill(BLUE, CYAN, ang=0) if side == "left" else gradFill(CYAN, BLUE, ang=0)
        s.append(sp(col_x, COL_Y, COL_W, cm(1.1), card_hdr_fill,
                    line=noLine(), geom="roundRect", tx=EMPTY_TX))

        # Заголовок карточки
        s.append(textbox(col_x, COL_Y, COL_W, cm(1.1),
                         [para(d[tk], 14, bold=True, color=WHITE, align="c")],
                         anchor="ctr"))

        # Буллеты карточки
        bp = [para("▶  " + b, 13, color=DARK, space_before=5)
              for b in d[bk]]
        s.append(textbox(col_x + cm(0.25), COL_Y + cm(1.25),
                         COL_W - cm(0.5), COL_H - cm(1.5), bp, anchor="t"))

    s.append(slideNum(num, total))
    return makeSlide(s)


# ── Stats-слайд (карточки с цифрами) ─────────────────────────────
def slideSTATS(d, num, total):
    s = []
    s.append(bgRect(solidFill(CREAM)))
    s.append(header())
    s.append(accentLine())
    s.append(titleBox(d["title"]))

    stats = d["stats"]
    n     = len(stats)
    gap   = cm(0.4)
    w_all = CX - cm(0.8) - gap * (n - 1)
    cw    = w_all // n
    cy_card = cm(3.65)
    ch_card = CY - cm(4.2)

    fills = [
        gradFill(BLUE, NAVY, ang=5400000),
        gradFill(NAVY, BLUE, ang=5400000),
        gradFill("0E4D92", BLUE, ang=5400000),
        gradFill(BLUE, "0E4D92", ang=5400000),
        gradFill(NAVY, "1A3A5C", ang=5400000),
    ]

    for i, st in enumerate(stats):
        cx_card = cm(0.4) + i * (cw + gap)
        # Карточка
        s.append(sp(cx_card, cy_card, cw, ch_card,
                    fills[i % len(fills)], line=noLine(),
                    geom="roundRect", tx=EMPTY_TX))
        # Значение
        s.append(textbox(cx_card, cy_card + cm(0.5), cw, cm(3.0),
                         [para(st["value"], 32, bold=True,
                               color=GOLD, align="c")],
                         anchor="ctr"))
        # Подпись
        lp = [para(l, 11, color=WHITE, align="c")
              for l in st["label"].split("\n")]
        s.append(textbox(cx_card + cm(0.15), cy_card + cm(3.3),
                         cw - cm(0.3), ch_card - cm(3.6), lp, anchor="t"))

    if "source" in d:
        s.append(textbox(cm(0.5), CY - cm(0.65), CX - cm(1.0), cm(0.5),
                         [para(d["source"], 8, color=GRAY,
                               italic=True, align="r")]))
    s.append(slideNum(num, total))
    return makeSlide(s)


# ══════════════════════════════════════════════════════════════════
#  ДАННЫЕ СЛАЙДОВ
# ══════════════════════════════════════════════════════════════════

SLIDES = [
    # 1 — Титул
    {"type": "title",
     "title": "Цифровизация экономики\nи устойчивость стран\nк глобальным кризисам",
     "subtitle": (
         "Выпускная квалификационная работа\n"
         "Уктамов Мухаммаджон\n\n"
         "Направление: Экономика\n"
         "Год защиты: 2025"
     )},

    # 2 — Актуальность
    {"type": "content",
     "title": "Актуальность исследования",
     "bullets": [
         "Три волны глобальных кризисов: 2008-2009, 2020-2021, 2022",
         "  Одинаковый внешний удар — но разная глубина спада в разных странах",
         "Цифровизация как структурный фактор устойчивости экономики",
         "COVID-19: цифровая инфраструктура обеспечила непрерывность государства и бизнеса",
         "Энергетический шок 2022: обнажил пределы цифровой зрелости",
         "Узбекистан реализует Стратегию «Цифровой Узбекистан - 2030»",
         "⚠ Центральный вопрос: через какие каналы цифровизация влияет на устойчивость?",
     ]},

    # 3 — Цель и задачи
    {"type": "two_col",
     "title": "Объект, предмет и цель исследования",
     "left_title": "Объект и предмет",
     "left_bullets": [
         "Объект: процесс цифровой трансформации национальных экономик",
         "Предмет: механизмы влияния цифровых технологий на макроэкономическую устойчивость",
         "Период: глобальные кризисы 2020-2022 годов",
     ],
     "right_title": "Цель и задачи",
     "right_bullets": [
         "Цель: выявить и эмпирически обосновать каналы влияния цифровизации на устойчивость",
         "Систематизировать теоретические подходы",
         "Сравнительный анализ 4 стран-лидеров",
         "Эконометрическая оценка на 49 странах",
         "Разработать рекомендации для Узбекистана",
     ]},

    # 4 — 5 каналов
    {"type": "content",
     "title": "Теоретическая рамка: 5 каналов влияния цифровизации на устойчивость",
     "bullets": [
         "Канал 1 - Непрерывность государственных функций в условиях шоков",
         "  e-Government, электронная идентификация, межведомственный обмен данными",
         "Канал 2 - Эффективность фискальной трансмиссии",
         "  Цифровые платежи -> адресные выплаты населению в реальном времени",
         "Канал 3 - Адаптация частного сектора через цифровые технологии",
         "  Цифровые фирмы показывают меньшее падение выручки в кризис",
         "Канал 4 - Структурный сдвиг спроса к цифровым услугам и e-commerce",
         "Канал 5 - Устойчивость ИКТ-сектора как отдельного компонента ВВП",
         "  ИКТ-сектор ОЭСР: рост даже в кризисном 2020 году",
     ]},

    # 5 — Сравнение стран (stats)
    {"type": "stats",
     "title": "Четыре модели цифровой трансформации: позиции в рейтингах 2024",
     "stats": [
         {"value": "#3 EGDI",   "label": "Эстония\nX-Road\ne-Residency"},
         {"value": "#2 NRI",    "label": "Республика\nКорея\nDigital New Deal"},
         {"value": "#1 IMD",    "label": "Сингапур\nSmart Nation 2.0\nЦЭ = 17.3% ВВП"},
         {"value": "#24 DESI",  "label": "Польша\nДогоняющая\nмодель"},
     ],
     "source": "Источники: UN DESA EGDI 2024; NRI 2024; IMD WDCR 2024; EC Digital Decade 2024"},

    # 6 — Реакция на кризис 2020 (stats)
    {"type": "stats",
     "title": "Реакция экономик на COVID-шок 2020 года (% изменения реального ВВП)",
     "stats": [
         {"value": "-0.7%", "label": "Республика Корея\nЛучший\nрезультат"},
         {"value": "-2.0%", "label": "Польша\nМягкое\nпадение"},
         {"value": "-3.0%", "label": "Эстония\nНиже среднего\nпо ЕС"},
         {"value": "-3.9%", "label": "Сингапур\nЗависимость\nот торговли"},
         {"value": "-6.1%", "label": "Среднее по ЕС\n(эталон\nсравнения)"},
     ],
     "source": "Источники: World Bank WDI; Eurostat; Bank of Korea, 2024"},

    # 7 — Разделитель Гл.II
    {"type": "section",
     "title": "Глава II",
     "subtitle": "Эконометрический анализ\n49 стран  |  707 наблюдений  |  2010-2024"},

    # 8 — Методология
    {"type": "content",
     "title": "Эконометрическая модель: методология",
     "bullets": [
         "Выборка: 49 стран, несбалансированная панель 2010-2024 гг., 707 наблюдений",
         "  Развитые ОЭСР (22 страны)  |  Переходные (16)  |  Развивающиеся (11, вкл. Узбекистан)",
         "Зависимая переменная: ln(ВВП на душу населения, постоянные цены 2015 г.)",
         "Основной регрессор: ln(доля интернет-пользователей, %)",
         "Контрольные: торговая открытость, качество институтов (WGI), норма накопления капитала",
         "Метод: модели с фиксированными эффектами (FE) + робастные S.E. по кластерам",
         "Тест Чоу: F=5.34, p=1.15e-11 - структурная неоднородность групп подтверждена",
     ]},

    # 9 — Результаты (stats)
    {"type": "stats",
     "title": "Результаты FE-моделей: коэффициент при ln(интернет-проникновение)",
     "stats": [
         {"value": "b=0.197",  "label": "Развитые\nэкономики\n(p<0.001)\n+COVID: b=0.394***\n+Геополит.: b=0.289**"},
         {"value": "b=0.309",  "label": "Переходные\nэкономики\n(p<0.001)\nНаибольший\nэффект"},
         {"value": "b=0.198",  "label": "Развивающиеся\n(вкл. Узбекистан)\n(p<0.001)\nКризисные члены\nнезначимы"},
     ],
     "source": "Гипотеза H1 подтверждена во всех трёх группах. Источник: расчёты автора, R Studio."},

    # 10 — Разделитель Гл.III
    {"type": "section",
     "title": "Глава III",
     "subtitle": "Цифровая трансформация Узбекистана\nи практические рекомендации"},

    # 11 — Узбекистан
    {"type": "content",
     "title": "Цифровая трансформация Узбекистана: текущее состояние",
     "bullets": [
         "Стратегия «Цифровой Узбекистан - 2030» (Указ No УП-6079, октябрь 2020 г.)",
         "E-Government индекс ООН 2024: значительный рост позиций страны",
         "IT Park Узбекистана: более 2 500 компаний-резидентов к концу 2024 г.",
         "Быстрый рост безналичных платежей, e-commerce и цифровых госсервисов",
         "⚠ Асимметрия: цифровое государство опережает цифровой бизнес и население",
         "⚠ Региональный цифровой разрыв: столица vs. сельские районы",
         "⚠ Интернет-проникновение ~80% — необходим рост до 90%+",
     ]},

    # 12 — Рекомендации
    {"type": "content",
     "title": "Практические рекомендации: 5 приоритетов",
     "bullets": [
         "1. Расширить интернет-проникновение до 90%+ -> прирост ВВП: +84-124 USD на чел./год",
         "2. Устранить асимметрию цифровой зрелости: подтянуть бизнес и цифровые навыки населения",
         "3. Развить инфраструктуру фискальной трансмиссии: интеграция MyID с адресными выплатами",
         "4. Учитывать тип кризиса: пандемийный шок vs. энергетический/инфляционный шок",
         "5. Накапливать данные для достижения статистической значимости кризисных эффектов",
     ]},

    # 13 — Дорожная карта
    {"type": "two_col",
     "title": "Дорожная карта цифровой устойчивости Узбекистана до 2030 года",
     "left_title": "Этап 1: 2025-2027",
     "left_bullets": [
         "Инфраструктура фискальной трансмиссии",
         "Интеграция MyID с системой адресных выплат",
         "Расширение широкополосного доступа в регионах",
         "Программы цифровых навыков для населения и МСП",
     ],
     "right_title": "Этап 2: 2027-2030+",
     "right_bullets": [
         "Устранение разрыва: цифровой бизнес = цифровое государство",
         "Достижение «порога» антикризисной цифровой зрелости",
         "Значимые кризисные эффекты в будущих эконометрических моделях",
         "Адаптация элементов X-Road к масштабу Узбекистана",
     ]},

    # 14 — Новизна
    {"type": "two_col",
     "title": "Научная новизна и практическая значимость",
     "left_title": "Научная новизна",
     "left_bullets": [
         "Систематизирована пятиканальная модель влияния цифровизации на устойчивость",
         "Эмпирически подтверждена структурная неоднородность эффекта (тест Чоу)",
         "Выявлен «порог» цифровой зрелости для значимого антикризисного эффекта",
     ],
     "right_title": "Практическая значимость",
     "right_bullets": [
         "Рекомендации для уточнения Стратегии «Цифровой Узбекистан - 2030»",
         "Методика адаптации опыта Эстонии, Кореи, Польши к условиям Узбекистана",
         "Прогнозные сценарии ВВП на душу населения до 2030 года",
     ]},

    # 15 — Итог
    {"type": "title",
     "title": "Основные выводы",
     "subtitle": (
         "H1 подтверждена: цифровизация значимо связана с ВВП во всех трёх группах (b=0.20-0.31)\n\n"
         "Антикризисный эффект значим только при высокой цифровой зрелости\n\n"
         "Для Узбекистана: последовательная реализация Стратегии «Цифровой Узбекистан - 2030»\n"
         "является необходимым условием перехода от роста к устойчивости\n\n"
         "Спасибо за внимание!"
     )},
]

# ══════════════════════════════════════════════════════════════════
#  СЛУЖЕБНЫЕ XML-ФАЙЛЫ .pptx
# ══════════════════════════════════════════════════════════════════

def contentTypesXML(n):
    overrides = "".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.'
        f'presentationml.slide+xml"/>'
        for i in range(1, n + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        + overrides +
        '</Types>'
    )

def rootRels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="ppt/presentation.xml"/>'
        '</Relationships>'
    )

def presentationXML(n):
    sldIds = "".join(
        f'<p:sldId id="{256 + i}" r:id="rId{i}"/>'
        for i in range(1, n + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'saveSubsetFonts="1">'
        '<p:sldMasterIdLst>'
        '<p:sldMasterId id="2147483648" r:id="rIdM"/>'
        '</p:sldMasterIdLst>'
        f'<p:sldIdLst>{sldIds}</p:sldIdLst>'
        '<p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        '</p:presentation>'
    )

def presentationRels(n):
    rels = "".join(
        f'<Relationship Id="rId{i}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
        f'Target="slides/slide{i}.xml"/>'
        for i in range(1, n + 1)
    )
    rels += (
        '<Relationship Id="rIdM" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        'Target="slideMasters/slideMaster1.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + rels +
        '</Relationships>'
    )

def slideRels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>'
    )

def slideLayoutXML():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldLayout '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'type="blank" preserve="1">'
        '<p:cSld name="Blank">'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm>'
        '<a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>'
        '</a:xfrm></p:grpSpPr>'
        '</p:spTree>'
        '</p:cSld>'
        '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        '</p:sldLayout>'
    )

def slideLayoutRels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        'Target="../slideMasters/slideMaster1.xml"/>'
        '</Relationships>'
    )

def slideMasterXML():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldMaster '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld>'
        '<p:bg><p:bgPr>'
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:effectLst/>'
        '</p:bgPr></p:bg>'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm>'
        '<a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>'
        '</a:xfrm></p:grpSpPr>'
        '</p:spTree>'
        '</p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" '
        'accent1="acc1" accent2="acc2" accent3="acc3" accent4="acc4" '
        'accent5="acc5" accent6="acc6" hlink="hlink" folHlink="folHlink"/>'
        '<p:sldLayoutIdLst>'
        '<p:sldLayoutId id="2147483649" r:id="rId1"/>'
        '</p:sldLayoutIdLst>'
        '<p:txStyles>'
        '<p:titleStyle><a:lstStyle/></p:titleStyle>'
        '<p:bodyStyle><a:lstStyle/></p:bodyStyle>'
        '<p:otherStyle><a:lstStyle/></p:otherStyle>'
        '</p:txStyles>'
        '</p:sldMaster>'
    )

def slideMasterRels():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>'
    )

# ══════════════════════════════════════════════════════════════════
#  СБОРКА .pptx
# ══════════════════════════════════════════════════════════════════

BUILDERS = {
    "title":   slideTITLE,
    "section": slideSECTION,
    "content": slideCONTENT,
    "two_col": slideTWOCOL,
    "stats":   slideSTATS,
}

def buildSlide(d, num, total):
    return BUILDERS.get(d["type"], slideCONTENT)(d, num, total)

def createPPTX(filename):
    n = len(SLIDES)
    with zipfile.ZipFile(filename, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",              contentTypesXML(n))
        z.writestr("_rels/.rels",                      rootRels())
        z.writestr("ppt/presentation.xml",             presentationXML(n))
        z.writestr("ppt/_rels/presentation.xml.rels",  presentationRels(n))
        z.writestr("ppt/slideMasters/slideMaster1.xml",           slideMasterXML())
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRels())
        z.writestr("ppt/slideLayouts/slideLayout1.xml",            slideLayoutXML())
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels",  slideLayoutRels())
        for i, d in enumerate(SLIDES, start=1):
            z.writestr(f"ppt/slides/slide{i}.xml",            buildSlide(d, i, n))
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", slideRels())
    kb = os.path.getsize(filename) // 1024
    print(f"OK  {filename}  ({n} слайдов, {kb} КБ)")

if __name__ == "__main__":
    createPPTX("Презентация_Уктамов_Мухаммаджон.pptx")
