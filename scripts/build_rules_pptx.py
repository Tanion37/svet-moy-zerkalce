"""Rules PPTX in Hobby World (МХ) booklet tone.

A4 portrait. Component art from assets/rules (PnP look, not named placeholders).
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
ART = ROOT / "assets" / "rules"

SLIDE_W = 210.0
SLIDE_H = 297.0
MARGIN = 12.0
CONTENT_W = SLIDE_W - 2 * MARGIN
MAX_GAP = 29.0  # ~5 lines

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


def add_title(slide, text: str, top: float) -> float:
    tf = _textbox(slide, MARGIN, top, CONTENT_W, 11)
    p = tf.paragraphs[0]
    _run(p, text, 18, bold=True)
    return top + 12


def add_heading(slide, text: str, top: float) -> float:
    tf = _textbox(slide, MARGIN, top, CONTENT_W, 8)
    p = tf.paragraphs[0]
    _run(p, text.upper(), 12, bold=True, color=ACCENT)
    return top + 8


def add_body(slide, lines: list[str], top: float, *, size: float = 11) -> float:
    h = max(7.0, len(lines) * (size * 0.42 + 2.0) + 2)
    tf = _textbox(slide, MARGIN, top, CONTENT_W, h)
    first = True
    for line in lines:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(0)
        p.space_after = Pt(2)
        _run(p, line, size)
    return top + h + 1


def add_picture(slide, name: str, left: float, top: float, width: float) -> float:
    path = ART / name
    pic = slide.shapes.add_picture(str(path), Mm(left), Mm(top), Mm(width))
    return pic.height.mm


class Book:
    def __init__(self, prs: Presentation):
        self.prs = prs
        self.slide = None
        self.y = MARGIN

    def leftover(self) -> float:
        return SLIDE_H - MARGIN - self.y

    def new_slide(self) -> None:
        self.slide = _blank(self.prs)
        self.y = MARGIN

    def need(self, h: float) -> None:
        if self.slide is None:
            self.new_slide()
            return
        if self.leftover() >= h:
            return
        self.new_slide()

    def heading(self, text: str) -> None:
        self.need(10)
        self.y = add_heading(self.slide, text, self.y)

    def title(self, text: str) -> None:
        self.need(14)
        self.y = add_title(self.slide, text, self.y)

    def body(self, lines: list[str], *, size: float = 11) -> None:
        h = max(7.0, len(lines) * (size * 0.42 + 2.0) + 3)
        self.need(h)
        self.y = add_body(self.slide, lines, self.y, size=size)

    def picture_row(
        self,
        items: list[tuple[str, str]],
        *,
        card_w: float = 42.0,
        gap: float = 4.0,
        caption_h: float = 8.0,
    ) -> None:
        n = len(items)
        total = n * card_w + (n - 1) * gap
        left0 = MARGIN + max(0.0, (CONTENT_W - total) / 2)
        # aspect 57 x 44.1
        card_h = card_w * (44.1 / 57.0)
        block = card_h + caption_h + 2
        self.need(block)
        for i, (name, caption) in enumerate(items):
            left = left0 + i * (card_w + gap)
            add_picture(self.slide, name, left, self.y, card_w)
            tf = _textbox(self.slide, left, self.y + card_h + 0.5, card_w, caption_h)
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            _run(p, caption, 8, color=MUTED)
        self.y += block

    def setup_image(self, max_h: float | None = None) -> None:
        name = "setup-table.png"
        # native ~180x115 mm at 200 dpi; fit width
        w = CONTENT_W
        h = w * (115.0 / 180.0)
        if max_h is not None and h > max_h:
            h = max_h
            w = h * (180.0 / 115.0)
        self.need(min(h, self.leftover() if self.leftover() > MAX_GAP else h))
        if self.leftover() < h:
            scale = self.leftover() / h
            w *= scale
            h *= scale
        left = MARGIN + (CONTENT_W - w) / 2
        add_picture(self.slide, name, left, self.y, w)
        self.y += h + 2

    def step_row(self, steps: list[tuple[str, str]]) -> None:
        n = len(steps)
        box_w = (CONTENT_W - (n - 1) * 4) / n
        h = 22.0
        self.need(h + 2)
        x = MARGIN
        for i, (num, title) in enumerate(steps):
            sh = self.slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Mm(x), Mm(self.y), Mm(box_w), Mm(h)
            )
            sh.fill.solid()
            sh.fill.fore_color.rgb = LIGHT
            sh.line.color.rgb = LINE
            sh.line.width = Pt(0.75)
            _no_shadow(sh)
            tf = sh.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            _run(p, f"{num}. {title}", 9, bold=True)
            x += box_w + 4
        self.y += h + 3


def build_book(book: Book) -> None:
    book.title("Свет мой зеркальце")
    book.body(
        [
            "ПРАВИЛА ИГРЫ",
            "3–6 игроков  ·  от 16 лет  ·  около 30 минут",
            "Контакты автора: yury.yamshchikov@gmail.com",
            "Версия правил 0.6",
        ],
        size=12,
    )
    book.heading("Об игре")
    book.body(
        [
            "Вы – гремлины службы поддержки волшебного сервиса «Зеркальце».",
            "Клиенты задают через зеркало странные вопросы, а вы на лету",
            "собираете ответы из карт букв, которые берёте из общей кучи.",
            "Чем смешнее и точнее реакция зала – тем выше шанс забрать",
            "карту запроса себе. Побеждает тот, у кого к концу колоды",
            "больше всего таких карт.",
            "Содержание: взрослый и чёрный юмор (в т.ч. намёки на отношения,",
            "алкоголь и табак). Политика и религия не используются.",
        ]
    )
    book.heading("Состав игры")
    book.body(
        [
            "• 14 карт запросов (двусторонние: начало + окончание фразы)",
            "• 72 карты букв (частотность русского языка; рубашка = джокер)",
            "• 6 карт «Ты восхитителен» (по 1 каждого цвета)",
            "• правила, которые вы читаете",
        ]
    )
    book.picture_row(
        [
            ("card-request-face.png", "запрос · лицо"),
            ("card-request-back.png", "запрос · оборот"),
            ("card-letter.png", "буква"),
            ("card-letter-joker.png", "джокер"),
            ("card-vote.png", "«Ты восхитителен»"),
        ],
        card_w=34.0,
        gap=3.0,
        caption_h=10.0,
    )

    book.heading("Подготовка к игре")
    book.body(
        [
            "1. Перемешайте колоду запросов и положите её рубашками вверх.",
            "   Верхняя карта показывает вторую половину будущего запроса.",
            "2. Каждый игрок выбирает цвет и берёт карту «Ты восхитителен»",
            "   своего цвета. Лишние уберите в коробку.",
            "3. Раскиньте все карты букв лицом вверх в случайном порядке",
            "   в центре стола. Руки игроков в начале раунда пустые.",
        ]
    )
    book.setup_image()

    book.heading("Цель игры")
    book.body(
        [
            "Набирайте карты запросов как победные очки. Когда колода",
            "запросов больше не позволяет начать раунд, побеждает игрок",
            "с наибольшим числом таких карт.",
        ]
    )
    book.heading("Как устроена карта запроса")
    book.body(
        [
            "1. Лицевая сторона – начало фразы-запроса.",
            "2. Рубашка – окончание фразы-запроса.",
            "Рамка зеркала на всех картах одинаковая: лицевая одной карты",
            "и рубашка следующей сверху в колоде складываются в целый запрос.",
        ]
    )
    book.picture_row(
        [
            ("card-request-face.png", "«Расскажи чем испугать всех»"),
            ("card-request-back.png", "«соседей сверху»"),
        ],
        card_w=70.0,
        gap=8.0,
        caption_h=10.0,
    )
    book.body(
        [
            "Читаете вслух: «Расскажи чем испугать всех соседей сверху».",
        ],
        size=11,
    )

    book.heading("Ход игры")
    book.body(
        [
            "Игра идёт раундами. Все действия в раунде, кроме возврата карт",
            "букв в центр, выполняются одновременно, если не сказано иное.",
        ]
    )
    book.step_row(
        [
            ("1", "Открытие запроса"),
            ("2", "Составление ответа"),
            ("3", "Голосование"),
            ("4", "Возврат букв"),
        ]
    )
    book.body(
        [
            "1. Снимите верхнюю карту, переверните на лицевую рядом с колодой.",
            "   Прочитайте: лицевая + рубашка новой верхней карты.",
            "   Если пары нет – конец игры.",
            "2. Игроки одновременно берут буквы из центра одной рукой по 1 карте.",
            "   Второй рукой можно держать взятое; копаться в куче нельзя.",
            "   Берите, пока хватит на ответ. Лишнее остаётся в центре.",
            "   Рубашка буквы вверх = джокер (замену вслух не называете).",
            "3. Все одновременно кидают «Ты восхитителен» к другому игроку.",
            "   Победитель раунда забирает лицевую карту запроса.",
            "4. Все карты букв возвращаются в центр, перемешиваются",
            "   и снова раскидываются лицом вверх.",
        ],
        size=10,
    )

    book.heading("Составление ответа")
    book.body(
        [
            "Ответ – слово или фраза на усмотрение игрока. Допустимы",
            "несуществующие и неполные слова. Несыгранные взятые буквы",
            "остаются у игрока до конца раунда.",
        ]
    )
    book.picture_row(
        [
            ("card-letter.png", "буква лицом"),
            ("card-letter-joker.png", "рубашка = джокер"),
        ],
        card_w=48.0,
        gap=10.0,
    )

    book.heading("Голосование «Ты восхитителен»")
    book.body(
        [
            "Карта должна быть в руке. Все кидают одновременно к другому",
            "игроку: себе оставлять нельзя. Неясно, к кому упала – решает",
            "владелец карты словами. Опоздал кинуть: его карта не учитывается,",
            "а каждый другой игрок получает по 2 голоса.",
            "Ничья: все лидеры побеждают. Один берёт лицевую карту, остальные",
            "по часовой стрелке – по карте с верха колоды запросов.",
        ],
        size=10,
    )
    book.heading("Пример")
    book.body(
        [
            "Запрос: «Чем соблазнить милф». Настя собрала ТИШИНА (3 голоса),",
            "Артём – ШОКОЛАД (1), Кирилл – КОТЫ (0). Победитель – Настя:",
            "она берёт лицевую карту запроса. Карты «Ты восхитителен»",
            "возвращаются владельцам, все буквы – в центр стола.",
        ],
        size=10,
    )
    book.picture_row(
        [
            ("card-vote.png", "голос (кидают к другому)"),
        ],
        card_w=50.0,
    )

    book.heading("Конец игры")
    book.body(
        [
            "Игра заканчивается, когда колода запросов больше не позволяет",
            "начать раунд (закончилась или осталась одна карта без пары).",
            "1. Каждый считает карты запросов перед собой.",
            "2. Побеждает игрок с наибольшим числом таких карт.",
            "При ничьей – совместная победа.",
        ]
    )
    book.heading("Частые вопросы")
    book.body(
        [
            "? Можно ли кинуть «Ты восхитителен» себе?  Нет. Только другому.",
            "? Карта упала между двумя – чей голос?  Решает владелец карты.",
            "? Копаться в разложенных буквах?  Нет. Только брать сверху/с края,",
            "  одной рукой по 1 карте. Вторая рука – только держать взятое.",
            "? Когда переставать брать буквы?  Когда хватит на ваш ответ.",
            "? Сколько джокеров можно сыграть?  Сколько угодно. Зал реже",
            "  голосует за нечитаемый ответ.",
        ],
        size=10,
    )
    book.heading("Создатели")
    book.body(
        [
            "Автор: Юрий Ямщиков",
            "Редакция / развитие: [плейсхолдер]",
            "Версия правил 0.6",
            "yury.yamshchikov@gmail.com",
        ],
        size=10,
    )
    book.heading("Памятка раунда")
    book.body(
        [
            "1. Открыть зеркало-запрос",
            "2. Взять буквы из центра (одна рука, по 1) и собрать ответ",
            "3. Кинуть «Ты восхитителен»",
            "4. Все буквы – в центр, перемешать, раскидать лицом вверх",
        ]
    )
    book.body(
        [
            "После раунда все карты букв снова лежат в центре лицом вверх.",
        ]
    )
    book.setup_image()


def build() -> Path:
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from generate_rules_art import build as build_art

    build_art()

    prs = Presentation()
    prs.slide_width = Mm(SLIDE_W)
    prs.slide_height = Mm(SLIDE_H)

    book = Book(prs)
    build_book(book)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "svet-moy-zerkalce-rules.pptx"
    prs.save(str(out))
    return out


if __name__ == "__main__":
    path = build()
    print(f"Wrote {path}")
