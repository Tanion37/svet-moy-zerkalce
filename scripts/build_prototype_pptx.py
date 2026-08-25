"""A4-landscape PnP: request cards (duplex) + «Ты восхитителен» (single-sided).

Follows TanionAgentSetting/prototype-presentation.md:
- word-fit font shrink; sheet labels; no back page for single-sided.
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Mm, Pt

try:
    from PIL import ImageFont
except ImportError:  # pragma: no cover
    ImageFont = None  # type: ignore

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output"

DESIGN_CARD_W_MM = 57.0
DESIGN_CARD_H_MM = 44.1
COLS = 5
ROWS = 4

SLIDE_W_MM = 297.0
SLIDE_H_MM = 210.0
PRINTER_MARGIN_MM = 5.0
LABEL_BAND_MM = 6.0

BORDER_EMU = 9525
LINE_MM = BORDER_EMU / 36000.0
SAFE_INSET_MM = PRINTER_MARGIN_MM + LINE_MM / 2
USABLE_W_MM = SLIDE_W_MM - 2 * SAFE_INSET_MM
# Leave a top band inside the safe area for the sheet label.
USABLE_H_MM = SLIDE_H_MM - 2 * SAFE_INSET_MM - LABEL_BAND_MM

DESIGN_GRID_W = COLS * DESIGN_CARD_W_MM
DESIGN_GRID_H = ROWS * DESIGN_CARD_H_MM
SCALE = min(USABLE_W_MM / DESIGN_GRID_W, USABLE_H_MM / DESIGN_GRID_H)
CARD_W_MM = DESIGN_CARD_W_MM * SCALE
CARD_H_MM = DESIGN_CARD_H_MM * SCALE
CONTENT_W = COLS * CARD_W_MM
CONTENT_H = ROWS * CARD_H_MM
MARGIN_X_MM = (SLIDE_W_MM - CONTENT_W) / 2
MARGIN_Y_MM = SAFE_INSET_MM + LABEL_BAND_MM + (
    (SLIDE_H_MM - SAFE_INSET_MM - LABEL_BAND_MM - SAFE_INSET_MM - CONTENT_H) / 2
)

BORDER_COLOR = RGBColor(0xC0, 0xC0, 0xC0)
INK = RGBColor(0x1A, 0x1A, 0x1A)
LABEL_COLOR = RGBColor(0x1A, 0x1A, 0x1A)

MAIN_PT = 18.0
MIN_PT = 8.0
PLAYER_LABEL_PT = 8.0
LABEL_PT = 11.0
PAD_X_MM = 2.0
PAD_Y_MM = 1.6

FACES: list[str] = [
    "Заклинание превращения в одного из",
    "Расскажи чем испугать всех",
    "Подбери прозвище для кого-то из",
    "Придумай страшное проклятие для",
    "Комплимент, который вгонит в краску",
    "Подскажи как отшить",
    "Почему стоит избегать",
    "Какая польза от",
    "Назови тайную слабость",
    "Какое наказание ждёт в аду для",
    "Чем оскорбить любого из",
    "Какой подлянки ждать от",
    "Лучший подарок для",
    "Чем соблазнить",
]

BACKS: list[str] = [
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
    "очкариков",
]

PLAYER_COLORS: list[RGBColor] = [
    RGBColor(0xC6, 0x28, 0x28),
    RGBColor(0x15, 0x65, 0xC0),
    RGBColor(0x2E, 0x7D, 0x32),
    RGBColor(0xEF, 0x6C, 0x00),
    RGBColor(0x6A, 0x1B, 0x9A),
    RGBColor(0x00, 0x83, 0x8F),
]


def _no_shadow(shape) -> None:
    """Drop theme effectRef (shadow) — inherit=False alone is not enough."""
    try:
        shape.shadow.inherit = False
    except Exception:
        pass
    el = shape._element
    style = el.find(
        "{http://schemas.openxmlformats.org/presentationml/2006/main}style"
    )
    if style is not None:
        el.remove(style)
    sp_pr = getattr(el, "spPr", None)
    if sp_pr is None:
        sp_pr = el.find(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}spPr"
        )
    if sp_pr is not None:
        for child in list(sp_pr):
            if child.tag.endswith("}effectLst"):
                sp_pr.remove(child)


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


def _word_width_pt(word: str, size_pt: float, *, bold: bool) -> float:
    if ImageFont is not None:
        try:
            name = "arialbd.ttf" if bold else "arial.ttf"
            font = ImageFont.truetype(name, int(round(size_pt)))
            return float(font.getlength(word))
        except OSError:
            pass
    # Fallback: conservative bold Cyrillic estimate
    return len(word) * size_pt * (0.72 if bold else 0.58)


def fit_font_for_words(
    text: str,
    max_width_mm: float,
    max_pt: float,
    *,
    bold: bool = True,
    min_pt: float = MIN_PT,
) -> float:
    """Shrink until every word fits on one line within max_width_mm."""
    words = [w for w in text.replace("\n", " ").split(" ") if w]
    if not words:
        return max_pt
    max_width_pt = max_width_mm * 72.0 / 25.4
    size = max_pt
    while size > min_pt + 1e-6:
        widest = max(_word_width_pt(w, size, bold=bold) for w in words)
        if widest <= max_width_pt:
            return round(size, 1)
        size -= 0.5
    return min_pt


def add_sheet_label(slide, text: str) -> None:
    """Label inside printer-safe area, above the card grid."""
    box = slide.shapes.add_textbox(
        Mm(SAFE_INSET_MM),
        Mm(SAFE_INSET_MM),
        Mm(SLIDE_W_MM - 2 * SAFE_INSET_MM),
        Mm(LABEL_BAND_MM - 0.5),
    )
    _no_shadow(box)
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = Mm(0)
    tf.margin_right = Mm(0)
    tf.margin_top = Mm(0)
    tf.margin_bottom = Mm(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    _add_run(p, text, LABEL_PT, bold=True, color=LABEL_COLOR)


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

    pad_x = Mm(PAD_X_MM * SCALE * pad_scale)
    pad_y = Mm(PAD_Y_MM * SCALE * pad_scale)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = pad_x
    tf.margin_right = pad_x
    tf.margin_top = pad_y
    tf.margin_bottom = pad_y
    return tf


def _text_width_mm(pad_scale: float = 1.0) -> float:
    return CARD_W_MM - 2 * PAD_X_MM * SCALE * pad_scale


def add_text_card(
    slide,
    left_mm: float,
    top_mm: float,
    text: str,
    *,
    anchor: MSO_ANCHOR,
    bold: bool = True,
    color: RGBColor = INK,
    size: float | None = None,
) -> None:
    if size is None:
        size = fit_font_for_words(text, _text_width_mm(), MAIN_PT, bold=bold)
    tf = _card_shape(slide, left_mm, top_mm, anchor=anchor)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(0)
    _add_run(p0, text, size, bold=bold, color=color)


def add_ty_voskhititelen_card(
    slide, left_mm: float, top_mm: float, player: int, color: RGBColor
) -> None:
    title = "Ты восхитителен"
    size = fit_font_for_words(title, _text_width_mm(), MAIN_PT, bold=True)
    tf = _card_shape(slide, left_mm, top_mm, anchor=MSO_ANCHOR.MIDDLE)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(0)
    _add_run(p0, title, size, bold=True, color=color)

    p1 = tf.add_paragraph()
    p1.alignment = PP_ALIGN.CENTER
    p1.space_before = Pt(4)
    p1.space_after = Pt(0)
    _add_run(p1, f"Игрок {player}", PLAYER_LABEL_PT, bold=False, color=color)


def _cell(left0: float, top0: float, idx: int) -> tuple[float, float, int, int]:
    col = idx % COLS
    row = idx // COLS
    return left0 + col * CARD_W_MM, top0 + row * CARD_H_MM, row, col


def add_request_face_sheet(slide, left0: float, top0: float) -> None:
    occupied: set[tuple[int, int]] = set()
    for idx, text in enumerate(FACES):
        left, top, row, col = _cell(left0, top0, idx)
        occupied.add((row, col))
        add_text_card(slide, left, top, text, anchor=MSO_ANCHOR.BOTTOM)
    add_cut_grid(slide, left0, top0, occupied)
    add_sheet_label(slide, "Карты запросов · ЛИЦО")


def add_request_back_sheet(slide, left0: float, top0: float) -> None:
    """Column-mirrored backs; no cut lines."""
    n = len(BACKS)
    rows_used = (n + COLS - 1) // COLS
    for row in range(rows_used):
        row_count = min(COLS, n - row * COLS)
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
    add_sheet_label(slide, "Карты запросов · ОБОРОТ")


def add_vote_sheet(slide, left0: float, top0: float) -> None:
    """Single-sided «Ты восхитителен» — no back slide."""
    occupied: set[tuple[int, int]] = set()
    for i, color in enumerate(PLAYER_COLORS):
        left, top, row, col = _cell(left0, top0, i)
        occupied.add((row, col))
        add_ty_voskhititelen_card(slide, left, top, i + 1, color)
    add_cut_grid(slide, left0, top0, occupied)
    add_sheet_label(slide, "Карты «Ты восхитителен»")


def _blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build() -> Path:
    assert CARD_W_MM > CARD_H_MM
    assert len(FACES) == len(BACKS) == 14

    prs = Presentation()
    prs.slide_width = Mm(SLIDE_W_MM)
    prs.slide_height = Mm(SLIDE_H_MM)

    left0, top0 = MARGIN_X_MM, MARGIN_Y_MM
    add_request_face_sheet(_blank_slide(prs), left0, top0)
    add_request_back_sheet(_blank_slide(prs), left0, top0)
    add_vote_sheet(_blank_slide(prs), left0, top0)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "svet-moy-zerkalce-cards.pptx"
    prs.save(str(out))
    return out


if __name__ == "__main__":
    path = build()
    sample = "здесь присутствующих"
    fitted = fit_font_for_words(sample, _text_width_mm(), MAIN_PT)
    print(f"Wrote {path}")
    print(f"Slide {SLIDE_W_MM:.2f}x{SLIDE_H_MM:.2f} mm (exact A4 landscape)")
    print(
        f"Margins >= {PRINTER_MARGIN_MM} mm "
        f"(actual x={MARGIN_X_MM:.3f}, y={MARGIN_Y_MM:.3f})"
    )
    print(
        f"Cards {CARD_W_MM:.3f}x{CARD_H_MM:.3f} mm "
        f"(design {DESIGN_CARD_W_MM}x{DESIGN_CARD_H_MM}, scale {SCALE:.4f})"
    )
    print(f"Word-fit sample «{sample}»: {fitted} pt (max {MAIN_PT})")
    print(
        f"Faces/backs {len(FACES)}; votes {len(PLAYER_COLORS)} "
        "(votes: face-only sheet, no back)"
    )
