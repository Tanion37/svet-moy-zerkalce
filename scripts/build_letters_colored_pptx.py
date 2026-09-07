"""Separate letter-card variant: doubled vowels; vowels red, consonants blue.

Canon 72-card black deck stays in svet-moy-zerkalce-letters.pptx.
"""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.dml.color import RGBColor

from build_letters_pptx import (
    FREQ_WEIGHTS,
    INK,
    PER_SLIDE,
    TOTAL_LETTERS,
    build,
    build_weighted_deck,
)

VOWELS = frozenset("аеёиоуыэюя")
CONSONANTS = frozenset("бвгджзйклмнпрстфхцчшщ")
VOWEL_COLOR = RGBColor(0xFF, 0x00, 0x00)
CONSONANT_COLOR = RGBColor(0x00, 0x00, 0xFF)
OUT_NAME = "svet-moy-zerkalce-letters-colored.pptx"


def double_vowels(letters: list[str]) -> list[str]:
    out: list[str] = []
    for ch in letters:
        out.append(ch)
        if ch.lower() in VOWELS:
            out.append(ch)
    return out


def color_for(letter: str) -> RGBColor:
    ch = letter.lower()
    if ch in VOWELS:
        return VOWEL_COLOR
    if ch in CONSONANTS:
        return CONSONANT_COLOR
    return INK


def build_colored() -> tuple[Path, list[str], float]:
    letters = double_vowels(build_weighted_deck(FREQ_WEIGHTS, TOTAL_LETTERS))
    return build(
        letters=letters,
        out_name=OUT_NAME,
        sheet_label="Карты букв",
        color_for=color_for,
    )


if __name__ == "__main__":
    path, letters, size = build_colored()
    counts = Counter(letters)
    n_full, n_last = divmod(len(letters), PER_SLIDE)
    n_sheets = n_full + (1 if n_last else 0)
    print(f"Wrote {path}")
    print(
        f"Grid {PER_SLIDE} per sheet; total {len(letters)} "
        f"({n_sheets} face-only sheets, last has {n_last or PER_SLIDE}, no backs)"
    )
    print(f"Letter size {size} pt (uniform, limited by widest glyph)")
    print("Counts:", dict(sorted(counts.items(), key=lambda x: (-x[1], x[0]))))
