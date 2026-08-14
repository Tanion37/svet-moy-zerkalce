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

MAIN_PT = 20.0
MEMO_TITLE_PT = 10.0
MEMO_BODY_PT = 8.0
PLAYER_LABEL_PT = 8.0

FACES: list[str] = [
    "Предложи ингредиенты любовного зелья для",
    "Произнеси заклинание превращения в*",
    "Скажи кто самый вонючий из*",
    "Расскажи чем испугать всех*",
    "Посоветуй чем подкупать*",
    "Назови тайную слабость*",
    "Подбери прозвище для любого из*",
    "Объясни как быстро «уложить»",
    "Придумай страшное проклятие для*",
    "Скажи чем эффективнее будить*",
    "Составь комплимент, который вгонит в краску*",
    "Назови причину, по которой стоит избегать",
    "Посоветуй чем задобрить наутро одного из",
    "Какое преступление типично для",
    "Скажи чем эффективно очаровывать",
    "Опиши худшее свидание для",
    "Подскажи как отшить",
    "Сочини сообщение «мы должны серьёзно поговорить» для",
    "Расскажи чем заменить утренний кофе для*",
    "Придумай чем взять на слабо*",
    "Чем оскорбить любого из",
    "Без чего нельзя представить*",
    "За что ненавидят",
    "Для кого сняли кино про*",
    "Что пригодится для пранка над любым из*",
    "Какую пользу можно ожидать от*",
    "Какой подлянки можно ожидать от*",
    "Какое наказание ждёт в аду для",
    "Любимая игра*",
    "Какое лекарство от головной боли лучше всего подойдёт для*",
    "Лучший подарок*",
]

BACKS: list[str] = [
    "*бородатых мужиков",
    "*животных на ферме",
    "*героев тупых комедий",
    "конченных интровертов",
    "*здесь присутствующих",
    "*ухоженных дровосеков",
    "твоих бывших",
    "твоих коллег после корпоратива",
    "*соседей сверху",
    "*вечных опоздунов",
    "гламурных девиц",
    "жертв пластической хирургии",
    "*фанатов комиксов",
    "*любителей настолок",
    "*ручных питомцев",
    "*самокатчиков",
    "*маменькиных сынков",
    "*бабок у подъезда",
    "алкашей",
    "чемпионов френдзоны",
    "*собачников",
    "*кошатниц",
    "пузатых скуфов",
    "милф",
    "*душнил",
    "вахтёров женской общаги",
    "*качков",
    "яжмамок",
    "*нейросетей",
    "*ботаников",
    "*очкариков",
]

PLAYER_COLORS: list[RGBColor] = [
    RGBColor(0xC6, 0x28, 0x28),
    RGBColor(0x15, 0x65, 0xC0),
    RGBColor(0x2E, 0x7D, 0x32),
    RGBColor(0xEF, 0x6C, 0x00),
    RGBColor(0x6A, 0x1B, 0x9A),
    RGBColor(0x00, 0x83, 0x8F),
]

BEZ_TITLE = "Без тормозов"
BEZ_BODY = (
    "• Первый накрывает «Я первый»\n"
    "• Остальным: «Не тормози!» + 10 пальцев\n"
    "• Не успел — пропуск (не цель голоса, сам кидает)\n"
    "• Ничья голосов: «Я первый» побеждает один"
)


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


def _card_shape(
    slide,
    left_mm: float,
    top_mm: float,
    *,
    anchor: MSO_ANCHOR,
    pad_scale: float = 1.0,
):
    """Card as a shape; label lives in its text_frame (not a separate textbox)."""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Mm(left_mm),
        Mm(top_mm),
        Mm(CARD_W_MM),
        Mm(CARD_H_MM),
    )
    shape.fill.background()
    shape.line.fill.background()
    _no_shadow(shape)

    pad_x = Mm(2.0 * SCALE * pad_scale)
    pad_y = Mm(1.6 * SCALE * pad_scale)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = pad_x
    tf.margin_right = pad_x
    tf.margin_top = pad_y
    tf.margin_bottom = pad_y
    return tf


def add_text_card(
    slide,
    left_mm: float,
    top_mm: float,
    text: str,
    *,
    anchor: MSO_ANCHOR,
    bold: bool = True,
    color: RGBColor = INK,
    size: float = MAIN_PT,
) -> None:
    tf = _card_shape(slide, left_mm, top_mm, anchor=anchor)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(0)
    _add_run(p0, text, size, bold=bold, color=color)


def add_ty_luchshiy_card(
    slide, left_mm: float, top_mm: float, player: int, color: RGBColor
) -> None:
    tf = _card_shape(slide, left_mm, top_mm, anchor=MSO_ANCHOR.MIDDLE)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(0)
    _add_run(p0, "Ты лучший!", MAIN_PT, bold=True, color=color)

    p1 = tf.add_paragraph()
    p1.alignment = PP_ALIGN.CENTER
    p1.space_before = Pt(4)
    p1.space_after = Pt(0)
    _add_run(p1, f"Игрок {player}", PLAYER_LABEL_PT, bold=False, color=color)


def add_bez_card(slide, left_mm: float, top_mm: float) -> None:
    tf = _card_shape(
        slide, left_mm, top_mm, anchor=MSO_ANCHOR.TOP, pad_scale=0.9
    )

    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(2)
    _add_run(p0, BEZ_TITLE, MEMO_TITLE_PT, bold=True)

    for line in BEZ_BODY.split("\n"):
        p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        _add_run(p, line, MEMO_BODY_PT, bold=False)


def add_face_sheet(slide, left0: float, top0: float, chunk: list[str]) -> None:
    occupied: set[tuple[int, int]] = set()
    for idx, text in enumerate(chunk):
        col = idx % COLS
        row = idx // COLS
        occupied.add((row, col))
        add_text_card(
            slide,
            left0 + col * CARD_W_MM,
            top0 + row * CARD_H_MM,
            text,
            anchor=MSO_ANCHOR.BOTTOM,
        )
    add_cut_grid(slide, left0, top0, occupied)


def add_back_sheet(slide, left0: float, top0: float, chunk: list[str]) -> None:
    """Column-mirrored backs for wide-side duplex within this sheet's chunk."""
    occupied: set[tuple[int, int]] = set()
    n = len(chunk)
    rows_used = (n + COLS - 1) // COLS
    for row in range(rows_used):
        row_count = min(COLS, n - row * COLS)
        for col in range(row_count):
            occupied.add((row, col))
            face_col = row_count - 1 - col
            face_idx = row * COLS + face_col
            add_text_card(
                slide,
                left0 + col * CARD_W_MM,
                top0 + row * CARD_H_MM,
                chunk[face_idx],
                anchor=MSO_ANCHOR.TOP,
            )
    add_cut_grid(slide, left0, top0, occupied)


def _chunks(items: list[str], size: int) -> list[list[str]]:
    return [items[i : i + size] for i in range(0, len(items), size)]


def add_other_sheet(slide, left0: float, top0: float) -> None:
    """Я первый, 6× Ты лучший!, Без тормозов — same card size."""
    cards: list[tuple] = [("ya", None)]
    for i in range(1, 7):
        cards.append(("ty", i))
    cards.append(("bez", None))

    occupied: set[tuple[int, int]] = set()
    for idx, (kind, player) in enumerate(cards):
        col = idx % COLS
        row = idx // COLS
        occupied.add((row, col))
        left = left0 + col * CARD_W_MM
        top = top0 + row * CARD_H_MM
        if kind == "ya":
            add_text_card(
                slide,
                left,
                top,
                "Я первый!",
                anchor=MSO_ANCHOR.MIDDLE,
                bold=True,
            )
        elif kind == "ty":
            assert player is not None
            add_ty_luchshiy_card(slide, left, top, player, PLAYER_COLORS[player - 1])
        elif kind == "bez":
            add_bez_card(slide, left, top)

    add_cut_grid(slide, left0, top0, occupied)


def _blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build() -> Path:
    assert CARD_W_MM > CARD_H_MM

    prs = Presentation()
    prs.slide_width = Mm(SLIDE_W_MM)
    prs.slide_height = Mm(SLIDE_H_MM)

    left0, top0 = MARGIN_X_MM, MARGIN_Y_MM

    # Interleave face/back sheets for duplex: 1 face, 2 back, 3 face, 4 back...
    face_chunks = _chunks(FACES, PER_SLIDE)
    back_chunks = _chunks(BACKS, PER_SLIDE)
    sheet_pairs = max(len(face_chunks), len(back_chunks))
    for i in range(sheet_pairs):
        if i < len(face_chunks):
            add_face_sheet(_blank_slide(prs), left0, top0, face_chunks[i])
        else:
            _blank_slide(prs)
        if i < len(back_chunks):
            add_back_sheet(_blank_slide(prs), left0, top0, back_chunks[i])
        else:
            _blank_slide(prs)

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
    print(
        f"Font: main {MAIN_PT} pt; memo/player label smaller; "
        "faces bottom, backs top; text on card shape"
    )
    print(
        f"Faces {len(FACES)}, backs {len(BACKS)}; "
        "duplex face/back sheets then other + Без тормозов"
    )
