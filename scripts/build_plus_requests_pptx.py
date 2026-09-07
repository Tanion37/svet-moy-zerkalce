"""Separate request-card prototype: canon 14 + extra 10 (24 duplex cards).

Votes / letters / rules stay in the main 14-card prototype files.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_prototype_pptx import (
    BACKS,
    FACE_WRAP_LINES,
    FACES,
    MAIN_PT,
    PER_SLIDE,
    _frame_margins,
    _text_width_mm,
    _wrap_words,
    build,
    extra_side_for_n_lines,
    fit_font_for_words,
)

PLUS_FACES: list[str] = [
    "Предложи ингредиенты любовного зелья для",
    "Скажи кто самый вонючий из",
    "Посоветуй чем подкупать",
    "Объясни как быстро «уложить»",
    "Скажи чем эффективнее будить",
    "Назови причину, по которой стоит избегать",
    "Посоветуй чем задобрить наутро одного из",
    "Какое преступление типично для",
    "Скажи чем эффективно очаровывать",
    "Опиши худшее свидание для",
]

PLUS_BACKS: list[str] = [
    "бородатых мужиков",
    "животных на ферме",
    "героев тупых комедий",
    "твоих коллег после корпоратива",
    "вечных опоздунов",
    "гламурных девиц",
    "жертв пластической хирургии",
    "фанатов комиксов",
    "ручных питомцев",
    "маменькиных сынков",
]

# Keep MAIN_PT; narrower column, N lines (same trick as canon nickname).
PLUS_WRAP_LINES: dict[str, int] = {
    "Назови причину, по которой стоит избегать": 3,
    "Посоветуй чем задобрить наутро одного из": 3,
}

FACES_OUT = "svet-moy-zerkalce-cards-plus.pptx"
BACKS_OUT = "svet-moy-zerkalce-card-backs-plus.pptx"


def _report_wrap(texts: list[str], wrap_lines: dict[str, int]) -> None:
    margins = _frame_margins(face=True)
    width = _text_width_mm(margins)
    for text in texts:
        n_lines = wrap_lines.get(text)
        if n_lines:
            extra = extra_side_for_n_lines(text, width, n_lines, MAIN_PT)
            inner = width - 2 * extra
            lines = _wrap_words(text, inner, MAIN_PT)
            print(
                f"«{text}»: {len(lines)} lines at {MAIN_PT} pt, "
                f"width {inner:.2f} mm, {lines}"
            )
        else:
            fitted = fit_font_for_words(text, width, MAIN_PT)
            lines = _wrap_words(text, width, fitted)
            print(
                f"«{text}»: {len(lines)} lines at {fitted} pt "
                f"(max {MAIN_PT}), {lines}"
            )


if __name__ == "__main__":
    faces = FACES + PLUS_FACES
    backs = BACKS + PLUS_BACKS
    wrap = {**FACE_WRAP_LINES, **PLUS_WRAP_LINES}
    faces_path, backs_path, _votes = build(
        faces,
        backs,
        wrap_lines=wrap,
        faces_name=FACES_OUT,
        backs_name=BACKS_OUT,
        include_votes=False,
    )
    print(f"Wrote {faces_path}")
    print(f"Wrote {backs_path}")
    n_sheets = (len(faces) + PER_SLIDE - 1) // PER_SLIDE
    print(f"Faces/backs {len(faces)} on {n_sheets} duplex sheet(s); no votes file")
    _report_wrap(PLUS_FACES, wrap)
    _report_wrap(PLUS_BACKS, wrap)
