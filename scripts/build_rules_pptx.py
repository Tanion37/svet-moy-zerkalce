"""Rules PPTX in Hobby World (МХ) booklet tone, with schematic examples.

A4 portrait slides. Illustrations = shapes + labels (no external art).
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Mm, Pt

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output"

SLIDE_W = 210.0
SLIDE_H = 297.0
MARGIN = 12.0

INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x55, 0x55, 0x55)
LINE = RGBColor(0x90, 0x90, 0x90)
ACCENT = RGBColor(0x2E, 0x5A, 0x88)
LIGHT = RGBColor(0xEE, 0xF2, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def _no_shadow(shape) -> None:
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def _blank(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _run(p, text, size, *, bold=False, color=INK):
    r = p.add_run()
    r.text = text
    r.font.name = "Arial"
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color


def _textbox(slide, left, top, width, height, *, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(Mm(left), Mm(top), Mm(width), Mm(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Mm(0)
    tf.margin_right = Mm(0)
    tf.margin_top = Mm(0)
    tf.margin_bottom = Mm(0)
    return tf


def add_title(slide, text: str, top: float = MARGIN) -> float:
    tf = _textbox(slide, MARGIN, top, SLIDE_W - 2 * MARGIN, 12)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    _run(p, text, 18, bold=True)
    return top + 14


def add_heading(slide, text: str, top: float | None = None) -> float:
    if top is None:
        top = MARGIN
    tf = _textbox(slide, MARGIN, top, SLIDE_W - 2 * MARGIN, 10)
    p = tf.paragraphs[0]
    _run(p, text.upper(), 13, bold=True, color=ACCENT)
    return top + 9


def add_body(slide, lines: list[str], top: float, *, size: float = 11) -> float:
    h = max(8.0, len(lines) * (size * 0.45 + 2.2))
    tf = _textbox(slide, MARGIN, top, SLIDE_W - 2 * MARGIN, h)
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(1)
        p.space_after = Pt(2)
        _run(p, line, size)
    return top + h + 2


def add_rect(slide, left, top, w, h, *, fill=None, line=LINE):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Mm(left), Mm(top), Mm(w), Mm(h)
    )
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    sh.line.color.rgb = line
    sh.line.width = Pt(0.75)
    _no_shadow(sh)
    return sh


def label_in_shape(shape, text: str, size: float = 9, *, bold=False, color=INK):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Mm(1.5)
    tf.margin_right = Mm(1.5)
    tf.margin_top = Mm(1)
    tf.margin_bottom = Mm(1)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _run(p, text, size, bold=bold, color=color)


def add_arrow_right(slide, left, top, w=8, h=4):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.RIGHT_ARROW, Mm(left), Mm(top), Mm(w), Mm(h)
    )
    sh.fill.solid()
    sh.fill.fore_color.rgb = ACCENT
    sh.line.fill.background()
    _no_shadow(sh)


def slide_cover(prs):
    s = _blank(prs)
    y = add_title(s, "Свет мой зеркальце")
    tf = _textbox(s, MARGIN, y, SLIDE_W - 2 * MARGIN, 8)
    p = tf.paragraphs[0]
    _run(p, "ПРАВИЛА ИГРЫ", 14, bold=True, color=ACCENT)
    y = add_body(
        s,
        [
            "3–6 игроков  ·  от 16 лет  ·  около 30 минут",
            "",
            "Контакты автора: yury.yamshchikov@gmail.com",
            "Версия правил 0.5  ·  стиль редакции МХ (прототип)",
        ],
        y + 10,
        size=12,
    )
    # meta strip
    box = add_rect(s, MARGIN, 80, SLIDE_W - 2 * MARGIN, 28, fill=LIGHT)
    label_in_shape(
        box,
        "Пати-игра: служба поддержки волшебного зеркала,\nбуквы и голосование «Ты восхитителен»",
        12,
        bold=True,
    )


def slide_about(prs):
    s = _blank(prs)
    y = add_heading(s, "Об игре")
    y = add_body(
        s,
        [
            "Вы – гремлины службы поддержки волшебного сервиса «Зеркальце».",
            "Клиенты задают через зеркало странные вопросы, а вы на лету",
            "собираете ответы из карт букв.",
            "",
            "Чем смешнее и точнее реакция зала – тем выше шанс забрать",
            "карту запроса себе. Побеждает тот, у кого к концу колоды",
            "больше всего таких карт.",
            "",
            "Содержание: взрослый и чёрный юмор (в т.ч. намёки на отношения,",
            "алкоголь и табак). Политика и религия не используются.",
        ],
        y,
    )
    y = add_heading(s, "Состав игры", y)
    add_body(
        s,
        [
            "• 14 карт запросов (двусторонние: начало + окончание фразы)",
            "• карты букв – по PnP-набору (частотность русского языка)",
            "• 6 карт «Ты восхитителен» (по 1 каждого цвета)",
            "• правила, которые вы читаете",
        ],
        y,
    )


def slide_setup(prs):
    s = _blank(prs)
    y = add_heading(s, "Подготовка к игре")
    y = add_body(
        s,
        [
            "1. Перемешайте колоду запросов и положите её рубашками вверх.",
            "   Верхняя карта показывает вторую половину будущего запроса.",
            "2. Каждый игрок выбирает цвет и берёт карту «Ты восхитителен»",
            "   своего цвета. Лишние уберите в коробку.",
            "3. Перемешайте карты букв и раздайте поровну. Остаток колоды,",
            "   если не делится без остатка, уберите в коробку закрытым.",
        ],
        y,
    )
    y = add_heading(s, "Схема стола", y)
    # table schematic
    cx = SLIDE_W / 2
    deck = add_rect(s, cx - 22, y + 8, 44, 28, fill=LIGHT)
    label_in_shape(deck, "Колода запросов\nрубашками вверх", 9, bold=True)

    # players around
    positions = [
        (cx - 55, y + 45, "Игрок A\nбуквы + голос"),
        (cx + 5, y + 45, "Игрок B\nбуквы + голос"),
        (cx - 55, y + 80, "Игрок C\nбуквы + голос"),
        (cx + 5, y + 80, "Игрок D\nбуквы + голос"),
    ]
    for left, top, text in positions:
        box = add_rect(s, left, top, 50, 22)
        label_in_shape(box, text, 8)


def slide_anatomy(prs):
    s = _blank(prs)
    y = add_heading(s, "Как устроена карта запроса")
    y = add_body(
        s,
        [
            "1. Лицевая сторона – начало фразы-запроса.",
            "2. Рубашка – окончание фразы-запроса.",
            "",
            "Рамка зеркала на всех картах одинаковая: лицевая одной карты",
            "и рубашка следующей сверху в колоде складываются в целый запрос.",
        ],
        y,
    )
    y = add_heading(s, "Пример зеркала", y)
    face = add_rect(s, MARGIN, y + 4, 80, 36, fill=LIGHT)
    label_in_shape(face, "ЛИЦЕВАЯ\n«Расскажи чем\nиспугать всех»", 10, bold=True)
    add_arrow_right(s, MARGIN + 84, y + 18, 10, 6)
    back = add_rect(s, MARGIN + 98, y + 4, 80, 36, fill=LIGHT)
    label_in_shape(back, "РУБАШКА колоды\n«соседей сверху»", 10, bold=True)

    y = y + 48
    result = add_rect(s, MARGIN, y, SLIDE_W - 2 * MARGIN, 22, fill=WHITE)
    label_in_shape(
        result,
        "Читаете вслух: «Расскажи чем испугать всех соседей сверху»",
        11,
        bold=True,
        color=ACCENT,
    )


def slide_round(prs):
    s = _blank(prs)
    y = add_heading(s, "Ход игры")
    y = add_body(
        s,
        [
            "Игра идёт раундами. Все действия в раунде, кроме передачи карт",
            "букв, выполняются одновременно, если не сказано иное.",
        ],
        y,
    )
    y = add_heading(s, "Раунд – 4 шага", y)
    steps = [
        ("1", "Открытие\nзапроса"),
        ("2", "Составление\nответа"),
        ("3", "Голосование\n«Ты восхитителен»"),
        ("4", "Передача\nсыгранных букв"),
    ]
    x = MARGIN
    for num, title in steps:
        circ = add_rect(s, x, y, 40, 28, fill=LIGHT)
        label_in_shape(circ, f"{num}\n{title}", 8, bold=True)
        if num != "4":
            add_arrow_right(s, x + 41, y + 11, 7, 5)
        x += 48
    y += 36
    add_body(
        s,
        [
            "1. Снимите верхнюю карту, переверните на лицевую рядом с колодой.",
            "   Прочитайте: лицевая + рубашка новой верхней карты.",
            "   Если пары нет – конец игры.",
            "2. Составьте ответ из букв на руке. Допустимы несуществующие слова.",
            "   Карта буквы рубашкой вверх = джокер (замену не озвучиваете).",
            "3. Все одновременно кидают «Ты восхитителен» к другому игроку.",
            "   Победитель раунда забирает лицевую карту запроса.",
            "4. Сыгранные буквы – соседу слева. Если букв < 10 – добор справа",
            "   до равенства рук (лишняя при нечёте остаётся справа).",
        ],
        y,
        size=10,
    )


def slide_vote_example(prs):
    s = _blank(prs)
    y = add_heading(s, "Пример голосования")
    y = add_body(
        s,
        [
            "Запрос: «Чем соблазнить милф».",
            "Игроки выложили ответы. Затем все кидают «Ты восхитителен».",
        ],
        y,
    )
    # three players
    players = [
        (MARGIN, "Артём\nответ: ШОКОЛАД", 1),
        (MARGIN + 62, "Настя\nответ: ТИШИНА", 3),
        (MARGIN + 124, "Кирилл\nответ: КОТЫ", 0),
    ]
    for left, text, votes in players:
        box = add_rect(s, left, y + 4, 56, 32, fill=LIGHT)
        label_in_shape(box, text, 9, bold=True)
        vbox = add_rect(s, left + 8, y + 40, 40, 14)
        label_in_shape(vbox, f"голосов: {votes}", 9)

    y = y + 62
    add_body(
        s,
        [
            "Победитель – Настя (3 голоса). Она берёт лицевую карту запроса",
            "как победное очко. Карты «Ты восхитителен» возвращаются владельцам.",
            "",
            "Ничья: все лидеры побеждают. Один берёт лицевую карту, остальные",
            "по часовой стрелке – по карте с верха колоды запросов.",
            "",
            "Опоздал кинуть голос: его карта не учитывается, а каждый другой",
            "игрок получает по 2 голоса.",
        ],
        y,
        size=10,
    )


def slide_letters(prs):
    s = _blank(prs)
    y = add_heading(s, "Карты букв")
    y = add_body(
        s,
        [
            "Ответ собирается из карт на руке. Несыгранные карты остаются.",
            "Любую карту можно сыграть рубашкой вверх как джокер.",
        ],
        y,
    )
    # letter row schematic
    for i, ch in enumerate(["С", "О", "С", "Е", "?"]):
        box = add_rect(s, MARGIN + i * 28, y + 2, 24, 24, fill=LIGHT if ch != "?" else WHITE)
        label_in_shape(box, ch, 14, bold=True, color=ACCENT if ch == "?" else INK)
    y += 32
    tip = add_rect(s, MARGIN, y, SLIDE_W - 2 * MARGIN, 18)
    label_in_shape(tip, "? = джокер (рубашка вверх). Замену вслух не называете.", 10)

    y += 26
    y = add_heading(s, "Передача после раунда", y)
    # left/right diagram
    a = add_rect(s, MARGIN, y + 4, 50, 24, fill=LIGHT)
    label_in_shape(a, "Вы", 11, bold=True)
    add_arrow_right(s, MARGIN + 52, y + 12, 12, 6)
    b = add_rect(s, MARGIN + 68, y + 4, 50, 24, fill=LIGHT)
    label_in_shape(b, "Сосед\nслева", 10, bold=True)

    tf = _textbox(s, MARGIN + 125, y + 2, 60, 30)
    p = tf.paragraphs[0]
    _run(p, "Сыгранные\nбуквы →", 10)

    y += 36
    add_body(
        s,
        [
            "Добор: если у вас меньше 10 карт, берите случайные у соседа справа,",
            "пока руки не станут равны. При нечётной сумме лишняя остаётся справа.",
        ],
        y,
        size=10,
    )


def slide_end_faq(prs):
    s = _blank(prs)
    y = add_heading(s, "Конец игры")
    y = add_body(
        s,
        [
            "Игра заканчивается, когда колода запросов больше не позволяет",
            "начать раунд (закончилась или осталась одна карта без пары).",
            "",
            "1. Каждый считает карты запросов перед собой.",
            "2. Побеждает игрок с наибольшим числом таких карт.",
            "При ничьей – у кого больше карт букв на руках;",
            "если и это одинаково – совместная победа.",
        ],
        y,
    )
    y = add_heading(s, "Частые вопросы", y)
    add_body(
        s,
        [
            "? Можно ли кинуть «Ты восхитителен» себе?",
            "Нет. Только другому игроку.",
            "",
            "? Карта упала между двумя игроками – чей голос?",
            "Решает владелец карты словами.",
            "",
            "? Обязательно составлять существующее слово?",
            "Нет. Допустимы неполные и выдуманные слова.",
            "",
            "? Сколько джокеров можно сыграть?",
            "Сколько угодно. Но зал реже голосует за нечитаемый ответ.",
        ],
        y,
        size=10,
    )


def slide_credits(prs):
    s = _blank(prs)
    y = add_heading(s, "Создатели")
    y = add_body(
        s,
        [
            "Автор: Юрий Ямщиков",
            "Редакция / развитие: [плейсхолдер]",
            "Иллюстрации схемы: прототип (фигуры PPTX)",
            "",
            "Стиль текста правил ориентирован на редакцию ООО «Мир Хобби».",
            "Это прототип, не коммерческий буклет МХ.",
            "",
            "Перепечатка и публикация без разрешения автора запрещены.",
            "© 2026 Юрий Ямщиков. Версия правил 0.5",
            "yury.yamshchikov@gmail.com",
        ],
        y,
    )
    # round memo
    y = add_heading(s, "Памятка раунда", y)
    add_body(
        s,
        [
            "1. Открыть зеркало-запрос",
            "2. Собрать ответ из букв",
            "3. Кинуть «Ты восхитителен»",
            "4. Передать сыгранные буквы влево → добрать при < 10",
        ],
        y,
    )


def build() -> Path:
    prs = Presentation()
    prs.slide_width = Mm(SLIDE_W)
    prs.slide_height = Mm(SLIDE_H)

    slide_cover(prs)
    slide_about(prs)
    slide_setup(prs)
    slide_anatomy(prs)
    slide_round(prs)
    slide_vote_example(prs)
    slide_letters(prs)
    slide_end_faq(prs)
    slide_credits(prs)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "svet-moy-zerkalce-rules.pptx"
    prs.save(str(out))
    return out


if __name__ == "__main__":
    path = build()
    print(f"Wrote {path} ({len(list(path.parent.glob('*.pptx')))} pptx in output)")
