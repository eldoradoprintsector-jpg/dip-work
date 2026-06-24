import zipfile, shutil, re, os

SRC = "Презентация Мухаммаджона (Защита).pptx"
DST = "Презентация Мухаммаджона (Защита).pptx"
TMP = "_tmp_pptx"

# ── вспомогательные функции ──────────────────────────────────────────────────

def rpr_bold(color, size=1000, font="Calibri"):
    return (
        f'<a:rPr lang="ru-RU" sz="{size}" b="1" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{font}" pitchFamily="34" charset="0"/>'
        f'<a:ea typeface="{font}" pitchFamily="34" charset="-122"/>'
        f'<a:cs typeface="{font}" pitchFamily="34" charset="-120"/>'
        f'</a:rPr>'
    )

def rpr_normal(color, size=1000, font="Calibri"):
    return (
        f'<a:rPr lang="ru-RU" sz="{size}" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{font}" pitchFamily="34" charset="0"/>'
        f'<a:ea typeface="{font}" pitchFamily="34" charset="-122"/>'
        f'<a:cs typeface="{font}" pitchFamily="34" charset="-120"/>'
        f'</a:rPr>'
    )

def para(pPr, *runs):
    """Собирает параграф из готовых <a:r>...</a:r> строк."""
    return f'<a:p>{pPr}{"".join(runs)}<a:endParaRPr lang="ru-RU" sz="1000" dirty="0"/></a:p>'

def run(rpr_str, text):
    return f'<a:r>{rpr_str}<a:t>{text}</a:t></a:r>'

PPR = '<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>'

# ── новые тексты для 4 стран ─────────────────────────────────────────────────
# Каждый блок: одна строка «Флаг Страна: » жирным + основной текст обычным

def make_country_txBody(flag_country, lines):
    """
    flag_country — напр. '🇪🇪 Эстония: '
    lines        — список строк обычного текста
    Возвращает полный <p:txBody>...</p:txBody>
    """
    paras = []
    # Первая строка: жирный заголовок + первая строка текста вместе
    first_line = lines[0] if lines else ""
    p0 = para(PPR,
              run(rpr_bold("6D2E46"), flag_country),
              run(rpr_normal("2D2D2D"), first_line))
    paras.append(p0)
    # Остальные строки — обычным текстом с небольшим отступом
    for line in lines[1:]:
        paras.append(para(PPR, run(rpr_normal("2D2D2D"), line)))

    body = (
        '<p:txBody>'
        '<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t"/>'
        '<a:lstStyle/>'
        + "".join(paras) +
        '</p:txBody>'
    )
    return body

# ── Тексты по странам ────────────────────────────────────────────────────────

ESTONIA_LINES = [
    "Tiger Leap (1996) - компьютеры и интернет в каждую школу.",
    "X-Road (2001) - платформа межведомственного обмена данными,",
    "  2,2 млрд транзакций/год, 450+ организаций.",
    "e-Residency (2014) - 128 тыс. чел. из 185 стран, 342 млн евро в бюджет.",
    "Вывод: EGDI №2 - цифровое государство с нуля за 25 лет.",
]

KOREA_LINES = [
    "Digital New Deal (2020) - 160 трлн вон (5,9% ВВП): данные, 5G, ИИ.",
    "DNA-концепция: Data + Network + AI - единая цифровая экосистема.",
    "18 больниц с 5G и IoT, 310 000 классов с быстрым интернетом.",
    "ОЭСР Digital Government 2023: №1 - государство-дирижёр.",
    "Вывод: привязать цифровизацию к антикризисным бюджетным программам.",
]

SINGAPORE_LINES = [
    "Smart Nation (2014) - 99% госуслуг онлайн, 99% домохозяйств в сети.",
    "Singpass - единый ID для всех услуг; PayNow - адресные выплаты за дни.",
    "Доля цифровой экономики в ВВП: 13,3% (2017) → 17,7% (2023).",
    "Smart Nation 2.0 (2024): акцент на генеративный ИИ, 120 млн SGD на науку.",
    "Вывод: IMD №1 - единый ID + быстрые платежи = антикризисный буфер.",
]

POLAND_LINES = [
    "«Цифровая Польша» 2014-2020 - 2,17 млрд евро из фондов ЕС.",
    "mObywatel (2017) - цифровой паспорт + 40+ госуслуг, 10 млн польз.",
    "Разрыв: е-правительство развито, но ИИ в бизнесе - лишь 3,7% (ЕС: 8%).",
    "Цифровые навыки населения: 44% (ниже среднего по ЕС - 56%).",
    "Вывод: не останавливаться на госсекторе - цифровизировать бизнес и людей.",
]

# ── Шаблон текстового блока (sp) ─────────────────────────────────────────────

def make_sp(sp_id, name, x, y, cx, cy, flag_country, lines):
    txBody = make_country_txBody(flag_country, lines)
    return (
        f'<p:sp>'
        f'<p:nvSpPr>'
        f'<p:cNvPr id="{sp_id}" name="{name}"/>'
        f'<p:cNvSpPr/><p:nvPr/>'
        f'</p:nvSpPr>'
        f'<p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/><a:ln/>'
        f'</p:spPr>'
        + txBody +
        f'</p:sp>'
    )

# ── Позиции блоков (оригинальные координаты из XML) ───────────────────────────
# id=11 (Польша), id=12 (Эстония), id=13 (Корея), id=14 (Сингапур)
# y-координаты подбираем под 4 блока равномерно под таблицей

BLOCKS = [
    # (sp_id, name,   x,       y,        cx,      cy,      flag+страна,              lines)
    (11, "Text 4a", 457200, 3870000, 8229600, 228600, "🇪🇪 Эстония: ",   ESTONIA_LINES),
    (12, "Text 4b", 457200, 4070000, 8229600, 228600, "🇰🇷 Корея: ",     KOREA_LINES),
    (13, "Text 4c", 457200, 4270000, 8229600, 228600, "🇸🇬 Сингапур: ",  SINGAPORE_LINES),
    (14, "Text 4d", 457200, 4470000, 8229600, 228600, "🇵🇱 Польша: ",    POLAND_LINES),
]

# ── Патч XML слайда 6 ─────────────────────────────────────────────────────────

def patch_slide6(xml: str) -> str:
    # 1. Удаляем старые текстовые блоки id=11..14 (они содержат описания стран)
    #    Ищем <p:sp>...<p:cNvPr id="11"...>...</p:sp> и т.д.
    for old_id in [11, 12, 13, 14]:
        pattern = (
            r'<p:sp>\s*<p:nvSpPr>\s*<p:cNvPr\s+id="' + str(old_id) + r'"[^/]*/>'
            r'.*?</p:sp>'
        )
        xml = re.sub(pattern, '', xml, flags=re.DOTALL)

    # 2. Вставляем 4 новых блока перед закрывающим </p:spTree>
    new_blocks = "\n".join(
        make_sp(sp_id, name, x, y, cx, cy, flag, lines)
        for sp_id, name, x, y, cx, cy, flag, lines in BLOCKS
    )
    xml = xml.replace('</p:spTree>', new_blocks + '\n</p:spTree>')

    return xml

# ── Пересборка PPTX ───────────────────────────────────────────────────────────

def rebuild_pptx(src, dst):
    tmp = src + ".tmp.pptx"
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'ppt/slides/slide6.xml':
                    xml = data.decode('utf-8')
                    xml = patch_slide6(xml)
                    data = xml.encode('utf-8')
                zout.writestr(item, data)
    os.replace(tmp, dst)
    print(f"✅ Готово: {dst}")

if __name__ == "__main__":
    rebuild_pptx(SRC, DST)
