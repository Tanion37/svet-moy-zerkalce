"""A4-landscape PnP: request cards + «Ты восхитителен» on one face/back pair.

Violates prototype-presentation §8 on purpose (mixed duplex + single-sided
on one sheet) per explicit user choice; do not change that guide.

Card design size: 57 × 44.1 mm (landscape). Grid 5×4 = 20.
Follows TanionAgentSetting/prototype-presentation.md otherwise.
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

MAIN_PT = 18.0
PLAYER_LABEL_PT = 8.0

FACES: list[str] = [
    "Произнеси заклинание превращения в",
    "Расскажи чем испугать всех",
    "Подбери прозвище для любого из",
    "Придумай страшное проклятие для",
    "Составь комплимент, который вгонит в краску",
    "Подскажи как отшить",
    "Что пригодится для пранка над любым из",
    "Какую пользу можно ожидать от",
    "Назови тайную слабость",
    "Какое наказание ждёт в аду для",
    "Чем оскорбить любого из",
    "Какой подлянки можно ожидать от",
    "Лучший подарок",
    "Чем соблазнить",
]

BACKS_UNIQUE: list[str] = [
    "конченных интровертов",
    "здесь присутствующих",
    "ухоженных дровосеков",
    "твоих бывших",
    "соседей сверху",
    "любителей настолок",
    "самокатчиков",
    "бабок у подъезда",
    "чемпионов френдзоны",
    "пузатых скуфов",
    "милф",
    "душнил",
    "нейросетей",
]

# 14th back repeats last unique ending so face/back counts match for duplex.
BACKS: list[str] = BACKS_UNIQUE + [BACKS_UNIQUE[-1]]

PLAYER_COLORS: list[RGBColor] = [
    RGBColor(0xC6, 0x28, 0x28),
    RGBColor(0x15, 0x65, 0xC0),
    RGBColor(0x2E, 0x7D, 0x32),
    RGBColor(0xEF, 0x6C, 0x00),
    RGBColor(0x6A, 0x1B, 0x9A),
    RGBColor(0x00, 0x83, 0x8F),
]


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


def add_ty_voskhititelen_card(
    slide, left_mm: float, top_mm: float, player: int, color: RGBColor
) -> None:
    tf = _card_shape(slide, left_mm, top_mm, anchor=MSO_ANCHOR.MIDDLE)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(0)
    _add_run(p0, "Ты восхитителен", MAIN_PT, bold=True, color=color)

    p1 = tf.add_paragraph()
    p1.alignment = PP_ALIGN.CENTER
    p1.space_before = Pt(4)
    p1.space_after = Pt(0)
    _add_run(p1, f"Игрок {player}", PLAYER_LABEL_PT, bold=False, color=color)


def _cell(left0: float, top0: float, idx: int) -> tuple[float, float, int, int]:
    col = idx % COLS
    row = idx // COLS
    return left0 + col * CARD_W_MM, top0 + row * CARD_H_MM, row, col


def add_face_sheet(slide, left0: float, top0: float) -> None:
    """14 request faces + 6 «Ты восхитителен» = 20 cells."""
    assert len(FACES) == 14
    assert len(PLAYER_COLORS) == 6
    occupied: set[tuple[int, int]] = set()

    for idx, text in enumerate(FACES):
        left, top, row, col = _cell(left0, top0, idx)
        occupied.add((row, col))
        add_text_card(slide, left, top, text, anchor=MSO_ANCHOR.BOTTOM)

    for i, color in enumerate(PLAYER_COLORS):
        idx = 14 + i
        left, top, row, col = _cell(left0, top0, idx)
        occupied.add((row, col))
        add_ty_voskhititelen_card(slide, left, top, i + 1, color)

    add_cut_grid(slide, left0, top0, occupied)


def add_back_sheet(slide, left0: float, top0: float) -> None:
    """Column-mirrored backs; vote cards get empty backs. No cut lines."""
    assert len(BACKS) == 14
    n_req = 14
    rows_used = (n_req + COLS - 1) // COLS
    for row in range(rows_used):
        row_count = min(COLS, n_req - row * COLS)
        for face_col in range(row_count):
            back_col = COLS - 1 - face_col
            face_idx = row * COLS + face_col
            add_text_card(
                slide,
                left0 + back_col * CARD_W_MM,
                top0 + row * CARD_H_MM,
                BACKS[face_idx],
                anchor=MSO_ANCHOR.TOP,
            )
    # Remaining 6 cells (row 3, mirrored): leave empty for «Ты восхитителен»


def _blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build() -> Path:
    assert CARD_W_MM > CARD_H_MM
    assert len(FACES) + len(PLAYER_COLORS) == PER_SLIDE

    prs = Presentation()
    prs.slide_width = Mm(SLIDE_W_MM)
    prs.slide_height = Mm(SLIDE_H_MM)

    left0, top0 = MARGIN_X_MM, MARGIN_Y_MM
    add_face_sheet(_blank_slide(prs), left0, top0)
    add_back_sheet(_blank_slide(prs), left0, top0)

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
        f"Faces {len(FACES)}, backs printed {len(BACKS)} "
        f"(unique {len(BACKS_UNIQUE)}); + {len(PLAYER_COLORS)} «Ты восхитителен»"
    )
