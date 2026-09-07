"""A4-portrait PnP for letter cards (Russian frequency deck).

Design card 44.1×57 mm (portrait). Grid 4×5 = 20 per page; 72 cards.
Empty backs ⇒ single-sided sheets only (prototype-presentation §11).
"""
from __future__ import annotations

from collections.abc import Callable
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

FREQ_WEIGHTS: dict[str, int] = {
    "о": 1097,
    "е": 848,
    "а": 806,
    "и": 737,
    "н": 670,
    "т": 631,
    "с": 547,
    "р": 473,
    "в": 454,
    "л": 440,
    "к": 349,
    "м": 321,
    "д": 298,
    "п": 281,
    "у": 262,
    "я": 201,
    "ы": 190,
    "ь": 174,
    "г": 169,
    "з": 165,
    "б": 159,
    "ч": 147,
    "й": 121,
    "х": 97,
    "ж": 94,
    "ш": 73,
    "ю": 64,
    "ц": 48,
    "щ": 36,
    "э": 32,
    "ф": 26,
    "ё": 20,
    "ъ": 2,
}

DESIGN_CARD_W_MM = 44.1
DESIGN_CARD_H_MM = 57.0
COLS = 4
ROWS = 5
PER_SLIDE = COLS * ROWS

SLIDE_W_MM = 210.0
SLIDE_H_MM = 297.0
PRINTER_MARGIN_MM = 5.0
LABEL_BAND_MM = 6.0

BORDER_EMU = 9525
LINE_MM = BORDER_EMU / 36000.0
SAFE_INSET_MM = PRINTER_MARGIN_MM + LINE_MM / 2
USABLE_W_MM = SLIDE_W_MM - 2 * SAFE_INSET_MM
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
LETTER_PT_CAP = 200.0
MIN_PT = 8.0
LABEL_PT = 11.0
PAD_X_MM = 2.0
PAD_Y_MM = 2.0
# PPT vs PIL metrics: keep a small inset so glyphs stay inside the cut line.
FIT_SAFETY = 0.96

TOTAL_LETTERS = 72


def build_weighted_deck(weights: dict[str, int], total: int) -> list[str]:
    letters = sorted(weights.keys(), key=lambda ch: (ch.replace("ё", "е\uffff"), ch))
    if total < len(letters):
        raise ValueError(f"total {total} < alphabet {len(letters)}")

    assigned = {k: 1 for k in letters}
    remain = total - len(letters)

    entries = [(k, weights[k]) for k in letters if weights[k] > 0]
    weight_sum = sum(w for _, w in entries)
    raw = []
    for k, w in entries:
        exact = (w / weight_sum) * remain
        raw.append({"k": k, "floor": int(exact), "frac": exact - int(exact)})
    used = sum(x["floor"] for x in raw)
    raw.sort(key=lambda x: x["frac"], reverse=True)
    i = 0
    while used < remain:
        raw[i % len(raw)]["floor"] += 1
        used += 1
        i += 1
    for x in raw:
        assigned[x["k"]] += x["floor"]

    out: list[str] = []
    for k in letters:
        out.extend([k] * assigned[k])
    return out


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
    if sp_pr is not None:
        for child in list(sp_pr):
            if child.tag.endswith("}effectLst"):
                sp_pr.remove(child)


def _add_run(
    paragraph,
    text: str,
    size: float,
    *,
    bold: bool = True,
    color: RGBColor = INK,
) -> None:
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


_FONT_DIR = Path(r"C:\Windows\Fonts")


def _load_font(size_pt: float, *, bold: bool):
    if ImageFont is None:
        return None
    name = "arialbd.ttf" if bold else "arial.ttf"
    for candidate in (_FONT_DIR / name, name):
        try:
            return ImageFont.truetype(str(candidate), int(round(size_pt)))
        except OSError:
            continue
    return None


def _glyph_box_pt(ch: str, size_pt: float, *, bold: bool) -> tuple[float, float, float]:
    """Ink width, ink height, and em height in pt."""
    font = _load_font(size_pt, bold=bold)
    if font is not None:
        x0, y0, x1, y1 = font.getbbox(ch)
        ascent, descent = font.getmetrics()
        return float(x1 - x0), float(y1 - y0), float(ascent + descent)
    # Conservative Cyrillic fallback if Arial is missing.
    return size_pt * 1.05, size_pt * 0.95, size_pt * 1.15


def fit_font_for_glyph(
    ch: str,
    max_width_mm: float,
    max_height_mm: float,
    max_pt: float,
    *,
    bold: bool = True,
    min_pt: float = MIN_PT,
) -> float:
    max_width_pt = max_width_mm * 72.0 / 25.4 * FIT_SAFETY
    max_height_pt = max_height_mm * 72.0 / 25.4 * FIT_SAFETY
    size = min(max_pt, max_height_pt)
    while size > min_pt + 1e-6:
        ink_w, ink_h, em_h = _glyph_box_pt(ch, size, bold=bold)
        if ink_w <= max_width_pt and ink_h <= max_height_pt and em_h <= max_height_pt:
            return round(size, 1)
        size -= 0.5
    return min_pt


def _text_box_mm() -> tuple[float, float]:
    return (
        CARD_W_MM - 2 * PAD_X_MM * SCALE,
        CARD_H_MM - 2 * PAD_Y_MM * SCALE,
    )


def uniform_letter_pt() -> float:
    text_w, text_h = _text_box_mm()
    return min(
        fit_font_for_glyph(ch.upper(), text_w, text_h, LETTER_PT_CAP, bold=True)
        for ch in FREQ_WEIGHTS
    )


def add_sheet_label(slide, text: str) -> None:
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


def add_letter_card(
    slide,
    left_mm: float,
    top_mm: float,
    letter: str,
    size: float,
    *,
    color: RGBColor = INK,
) -> None:
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
    tf = shape.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Mm(PAD_X_MM * SCALE)
    tf.margin_right = Mm(PAD_X_MM * SCALE)
    tf.margin_top = Mm(PAD_Y_MM * SCALE)
    tf.margin_bottom = Mm(PAD_Y_MM * SCALE)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.CENTER
    p0.space_before = Pt(0)
    p0.space_after = Pt(0)
    p0.line_spacing = 1.0
    _add_run(p0, letter.upper(), size, bold=True, color=color)


def add_face_sheet(
    slide,
    left0: float,
    top0: float,
    chunk: list[str],
    size: float,
    *,
    color_for: Callable[[str], RGBColor] | None = None,
    sheet_label: str = "Карты букв",
) -> None:
    occupied: set[tuple[int, int]] = set()
    for idx, letter in enumerate(chunk):
        col = idx % COLS
        row = idx // COLS
        occupied.add((row, col))
        add_letter_card(
            slide,
            left0 + col * CARD_W_MM,
            top0 + row * CARD_H_MM,
            letter,
            size,
            color=color_for(letter) if color_for else INK,
        )
    add_cut_grid(slide, left0, top0, occupied)
    # Empty backs ⇒ single-sided (§11): no ЛИЦО/ОБОРОТ marker
    add_sheet_label(slide, sheet_label)


def _chunks(items: list[str], size: int) -> list[list[str]]:
    return [items[i : i + size] for i in range(0, len(items), size)]


def _blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build(
    *,
    letters: list[str] | None = None,
    out_name: str = "svet-moy-zerkalce-letters.pptx",
    sheet_label: str = "Карты букв",
    color_for: Callable[[str], RGBColor] | None = None,
) -> tuple[Path, list[str], float]:
    assert CARD_H_MM > CARD_W_MM
    if letters is None:
        letters = build_weighted_deck(FREQ_WEIGHTS, TOTAL_LETTERS)
        assert len(letters) == TOTAL_LETTERS
    size = uniform_letter_pt()

    prs = Presentation()
    prs.slide_width = Mm(SLIDE_W_MM)
    prs.slide_height = Mm(SLIDE_H_MM)
    left0, top0 = MARGIN_X_MM, MARGIN_Y_MM

    # Empty joker backs ⇒ no back sheets (prototype-presentation §11)
    for chunk in _chunks(letters, PER_SLIDE):
        add_face_sheet(
            _blank_slide(prs),
            left0,
            top0,
            chunk,
            size,
            color_for=color_for,
            sheet_label=sheet_label,
        )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / out_name
    prs.save(str(out))
    return out, letters, size


if __name__ == "__main__":
    path, letters, size = build()
    from collections import Counter

    counts = Counter(letters)
    n_full, n_last = divmod(TOTAL_LETTERS, PER_SLIDE)
    n_sheets = n_full + (1 if n_last else 0)
    print(f"Wrote {path}")
    print(f"Slide {SLIDE_W_MM:.2f}x{SLIDE_H_MM:.2f} mm (A4 portrait)")
    print(
        f"Grid {COLS}x{ROWS}={PER_SLIDE}; total {TOTAL_LETTERS} "
        f"({n_sheets} face-only sheets, last has {n_last or PER_SLIDE}, no backs)"
    )
    print(
        f"Cards {CARD_W_MM:.3f}x{CARD_H_MM:.3f} mm "
        f"(design {DESIGN_CARD_W_MM}x{DESIGN_CARD_H_MM}, scale {SCALE:.4f})"
    )
    print(f"Letter size {size} pt (uniform, limited by widest glyph)")
    print("Counts:", dict(sorted(counts.items(), key=lambda x: (-x[1], x[0]))))
