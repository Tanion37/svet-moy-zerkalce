"""A4-landscape PnP sheets for «Свет мой зеркальце».

Card design size: 57 × 44.1 mm (landscape). Grid 5×4 = 20 on slides 1–2.
Follows TanionAgentSetting/prototype-presentation.md.
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

DESIGN_CARD_W_MM = 57.0
DESIGN_CARD_H_MM = 44.1
COLS = 5
ROWS = 4
PER_SLIDE = COLS * ROWS

SLIDE_W_MM = 297.0
SLIDE_H_MM = 210.0
PRINTER_MARGIN_MM = 5.0

BORDER_EMU = 9525
LINE_MM = BORDER_EMU / 36000.0
SAFE_INSET_MM = PRINTER_MARGIN_MM + LINE_MM / 2
USABLE_W_MM = SLIDE_W_MM - 2 * SAFE_INSET_MM
USABLE_H_MM = SLIDE_H_MM - 2 * SAFE_INSET_MM

DESIGN_GRID_W = COLS * DESIGN_CARD_W_MM
DESIGN_GRID_H = ROWS * DESIGN_CARD_H_MM
SCALE = min(USABLE_W_MM / DESIGN_GRID_W, USABLE_H_MM / DESIGN_GRID_H)
CARD_W_MM = DESIGN_CARD_W_MM * SCALE
CARD_H_MM = DESIGN_CARD_H_MM * SCALE
CONTENT_W = COLS * CARD_W_MM
CONTENT_H = ROWS * CARD_H_MM
MARGIN_X_MM = (SLIDE_W_MM - CONTENT_W) / 2
MARGIN_Y_MM = (SLIDE_H_MM - CONTENT_H) / 2

BORDER_COLOR = RGBColor(0xC0, 0xC0, 0xC0)
INK = RGBColor(0x1A, 0x1A, 0x1A)

MIN_PT = 8.0
PREF_PT = 10.0

FACES: list[str] = [
    "Предложи ингредиенты любовного зелья для",
    "Произнеси заклинание усыпления для",
    "Скажи кто самый вонючий из",
    "Расскажи чем испугать всех",
    "Посоветуй чем подкупать",
    "Назови тайную слабость",
    "Подбери прозвище для любого из",
    "Объясни как быстро «уложить»",
    "Придумай страшное проклятие для",
    "Скажи чем эффективнее будить",
    "Составь комплимент, который вгонит в краску",
    "Назови причину, по которой стоит избегать",
    "Посоветуй чем задобрить наутро одного из",
    "Придумай алиби для кого-то из",
    "Скажи чем эффективно очаровывать",
    "Опиши худшее свидание для",
    "Подскажи как отшить",
    "Сочини смс «мы должны серьёзно поговорить» для",
    "Расскажи чем заменить утренний кофе для",
    "Придумай челлендж на слабо для",
]

BACKS: list[str] = [
    "бородатых мужиков",
    "животных на ферме",
    "героев тупых комедий",
    "конченных интровертов",
    "здесь присутствующих",
    "дровосеков",
    "типичных героев мемов",
    "твоих бывших",
    "твоих коллег после корпоратива",
    "соседей сверху",
    "вечных опоздунов",
    "гламурных девиц",
    "жертв пластической хирургии",
    "фанатов комиксов",
    "любителей настолок",
    "ручных питомцев",
    "криптоинвестеров",
    "самокатчиков",
    "маменьких сынков",
    "величайших злодеев в истории",
]

# Distinct readable ink colors (no card fill)
PLAYER_COLORS: list[RGBColor] = [
    RGBColor(0xC6, 0x28, 0x28),  # 1 red
    RGBColor(0x15, 0x65, 0xC0),  # 2 blue
    RGBColor(0x2E, 0x7D, 0x32),  # 3 green
    RGBColor(0xEF, 0x6C, 0x00),  # 4 orange
    RGBColor(0x6A, 0x1B, 0x9A),  # 5 purple
    RGBColor(0x00, 0x83, 0x8F),  # 6 teal
]

MEMO_TITLE = "Памятка"
MEMO_BODY = (
    "1. Открой зеркало-запрос\n"
    "2. Ответ из букв (рубашка = джокер)\n"
    "3. Кинь «Ты лучший!» другому\n"
    "4. Больше голосов — лицевая карта\n"
    "5. Буквы влево; если <10 — добор справа"
)

BEZ_TITLE = "Без тормозов"
BEZ_BODY = (
    "• Первый накрывает «Я первый»\n"
    "• Остальным: «Не тормози!» + 10 пальцев\n"
    "• Не успел — пропуск (не цель голоса, сам кидает)\n"
    "• Ничья голосов: «Я первый» побеждает один"
)


def _pt_for(text: str, *, long_if_over: int) -> float:
    return MIN_PT if len(text) > long_if_over else PREF_PT


def _no_shadow(shape) -> None:
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def _add_run(
    paragraph,
    text: str,
    size: float,
    *,
    bold: bool = False,
    color: RGBColor = INK,
    font_name: str = "Arial",
) -> None:
    run = paragraph.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def _cut_hline(slide, left_mm: float, top_mm: float, width_mm: float) -> None:
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Mm(left_mm),
        Mm(top_mm - LINE_MM / 2),
        Mm(width_mm),
        Mm(LINE_MM),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = BORDER_COLOR
    shape.line.fill.background()
    _no_shadow(shape)


def _cut_vline(slide, left_mm: float, top_mm: float, height_mm: float) -> None:
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Mm(left_mm - LINE_MM / 2),
        Mm(top_mm),
        Mm(LINE_MM),
        Mm(height_mm),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = BORDER_COLOR
    shape.line.fill.background()
    _no_shadow(shape)


def add_cut_grid(
    slide, left0: float, top0: float, occupied: set[tuple[int, int]] | None = None
) -> None:
    if occupied is None:
        occupied = {(r, c) for r in range(ROWS) for c in range(COLS)}

    for row_edge in range(ROWS + 1):
        for col in range(COLS):
            above = (row_edge - 1, col) in occupied if row_edge > 0 else False
            below = (row_edge, col) in occupied if row_edge < ROWS else False
            if above or below:
                _cut_hline(
                    slide,
                    left0 + col * CARD_W_MM,
                    top0 + row_edge * CARD_H_MM,
                    CARD_W_MM,
                )
    for col_edge in range(COLS + 1):
        for row in range(ROWS):
            left = (row, col_edge - 1) in occupied if col_edge > 0 else False
            right = (row, col_edge) in occupied if col_edge < COLS else False
            if left or right:
                _cut_vline(
                    slide,
                    left0 + col_edge * CARD_W_MM,
                    top0 + row * CARD_H_MM,
                    CARD_H_MM,
                )


def _textbox(slide, left_mm: float, top_mm: float, pad_scale: float = 1.0):
    pad_x = 2.0 * SCALE * pad_scale
    pad_y = 1.6 * SCALE * pad_scale
    box = slide.shapes.add_textbox(
        Mm(left_mm + pad_x),
        Mm(top_mm + pad_y),
        Mm(CARD_W_MM - 2 * pad_x),
        Mm(CARD_H_MM - 2 * pad_y),
    )
    _no_shadow(box)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    return tf


def add_text_card(
    slide,
    left_mm: float,
    top_mm: float,
    text: str,
    *,
    bold: bool = True,
    color: RGBColor = INK,
    long_if_over: int = 36,
) -> None:
    tf = _textbox(slide, left_mm, top_mm)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(0)
    _add_run(p0, text, _pt_for(text, long_if_over=long_if_over), bold=bold, color=color)


def add_ty_luchshiy_card(
    slide, left_mm: float, top_mm: float, player: int, color: RGBColor
) -> None:
    tf = _textbox(slide, left_mm, top_mm)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(0)
    _add_run(p0, "Ты лучший!", PREF_PT, bold=True, color=color)

    p1 = tf.add_paragraph()
    p1.alignment = PP_ALIGN.CENTER
    p1.space_before = Pt(4)
    p1.space_after = Pt(0)
    _add_run(p1, f"Игрок {player}", MIN_PT, bold=False, color=color)


def add_memo_card(
    slide, left_mm: float, top_mm: float, title: str, body: str
) -> None:
    tf = _textbox(slide, left_mm, top_mm, pad_scale=0.9)
    tf.vertical_anchor = MSO_ANCHOR.TOP

    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(2)
    _add_run(p0, title, PREF_PT, bold=True)

    for line in body.split("\n"):
        p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        _add_run(p, line, MIN_PT, bold=False)


def add_face_sheet(slide, left0: float, top0: float) -> None:
    assert len(FACES) == PER_SLIDE
    for idx, text in enumerate(FACES):
        col = idx % COLS
        row = idx // COLS
        add_text_card(
            slide,
            left0 + col * CARD_W_MM,
            top0 + row * CARD_H_MM,
            text,
            long_if_over=32,
        )
    add_cut_grid(slide, left0, top0)


def add_back_sheet(slide, left0: float, top0: float) -> None:
    """Column-mirrored backs for wide-side duplex."""
    assert len(BACKS) == PER_SLIDE
    for row in range(ROWS):
        for col in range(COLS):
            face_col = COLS - 1 - col
            face_idx = row * COLS + face_col
            add_text_card(
                slide,
                left0 + col * CARD_W_MM,
                top0 + row * CARD_H_MM,
                BACKS[face_idx],
                long_if_over=28,
            )
    add_cut_grid(slide, left0, top0)


def add_other_sheet(slide, left0: float, top0: float) -> None:
    """Я первый, 6× Ты лучший!, памятка, Без тормозов — same card size."""
    cards: list[tuple] = [("ya", None)]
    for i in range(1, 7):
        cards.append(("ty", i))
    cards.append(("memo", None))
    cards.append(("bez", None))

    occupied: set[tuple[int, int]] = set()
    for idx, (kind, player) in enumerate(cards):
        col = idx % COLS
        row = idx // COLS
        occupied.add((row, col))
        left = left0 + col * CARD_W_MM
        top = top0 + row * CARD_H_MM
        if kind == "ya":
            add_text_card(slide, left, top, "Я первый!", bold=True, long_if_over=20)
        elif kind == "ty":
            assert player is not None
            add_ty_luchshiy_card(slide, left, top, player, PLAYER_COLORS[player - 1])
        elif kind == "memo":
            add_memo_card(slide, left, top, MEMO_TITLE, MEMO_BODY)
        elif kind == "bez":
            add_memo_card(slide, left, top, BEZ_TITLE, BEZ_BODY)

    add_cut_grid(slide, left0, top0, occupied)


def _blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build() -> Path:
    assert len(FACES) == 20 and len(BACKS) == 20
    assert CARD_W_MM > CARD_H_MM

    prs = Presentation()
    prs.slide_width = Mm(SLIDE_W_MM)
    prs.slide_height = Mm(SLIDE_H_MM)

    left0, top0 = MARGIN_X_MM, MARGIN_Y_MM

    add_face_sheet(_blank_slide(prs), left0, top0)
    add_back_sheet(_blank_slide(prs), left0, top0)
    add_other_sheet(_blank_slide(prs), left0, top0)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "svet-moy-zerkalce-cards.pptx"
    prs.save(str(out))
    return out


if __name__ == "__main__":
    path = build()
    print(f"Wrote {path}")
    print(f"Slide {SLIDE_W_MM:.2f}x{SLIDE_H_MM:.2f} mm (exact A4 landscape)")
    print(
        f"Margins >= {PRINTER_MARGIN_MM} mm "
        f"(actual x={MARGIN_X_MM:.3f}, y={MARGIN_Y_MM:.3f})"
    )
    print(
        f"Cards on slide {CARD_W_MM:.3f}x{CARD_H_MM:.3f} mm "
        f"(design {DESIGN_CARD_W_MM}x{DESIGN_CARD_H_MM}, scale {SCALE:.4f})"
    )
    print(f"Font: prefer {PREF_PT} pt, min {MIN_PT} pt; no card fills")
    print("Slides: 1 faces, 2 backs (column-mirrored), 3 other + memos")
