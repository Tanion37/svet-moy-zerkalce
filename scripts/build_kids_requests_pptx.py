"""Kids-only request cards PnP (faces + backs). No utility cards.

Follows TanionAgentSetting/prototype-presentation.md.
Reuses layout helpers from build_prototype_pptx.py.
"""
from __future__ import annotations

import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Mm

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_prototype_pptx import (  # noqa: E402
    BACKS,
    CARD_H_MM,
    CARD_W_MM,
    DESIGN_CARD_H_MM,
    DESIGN_CARD_W_MM,
    FACES,
    MAIN_PT,
    MARGIN_X_MM,
    MARGIN_Y_MM,
    OUT_DIR,
    PER_SLIDE,
    PRINTER_MARGIN_MM,
    SCALE,
    SLIDE_H_MM,
    SLIDE_W_MM,
    _blank_slide,
    _chunks,
    add_back_sheet,
    add_face_sheet,
)


def _strip_mark(text: str) -> str:
    if text.startswith("*"):
        return text[1:]
    if text.endswith("*"):
        return text[:-1]
    return text


def kids_faces() -> list[str]:
    return [_strip_mark(t) for t in FACES if t.endswith("*")]


def kids_backs() -> list[str]:
    return [_strip_mark(t) for t in BACKS if t.startswith("*")]


def _pad_pair(faces: list[str], backs: list[str]) -> tuple[list[str], list[str]]:
    """Same length for duplex grid; cycle the shorter pool."""
    if not faces or not backs:
        raise ValueError("Kids face/back pools must be non-empty")
    n = max(len(faces), len(backs))
    return (
        [faces[i % len(faces)] for i in range(n)],
        [backs[i % len(backs)] for i in range(n)],
    )


def build() -> Path:
    assert CARD_W_MM > CARD_H_MM
    faces, backs = _pad_pair(kids_faces(), kids_backs())

    prs = Presentation()
    prs.slide_width = Mm(SLIDE_W_MM)
    prs.slide_height = Mm(SLIDE_H_MM)
    left0, top0 = MARGIN_X_MM, MARGIN_Y_MM

    face_chunks = _chunks(faces, PER_SLIDE)
    back_chunks = _chunks(backs, PER_SLIDE)
    assert len(face_chunks) == len(back_chunks)
    for fc, bc in zip(face_chunks, back_chunks, strict=True):
        add_face_sheet(_blank_slide(prs), left0, top0, fc)
        add_back_sheet(_blank_slide(prs), left0, top0, bc)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "svet-moy-zerkalce-kids-requests.pptx"
    prs.save(str(out))
    return out


if __name__ == "__main__":
    raw_f, raw_b = kids_faces(), kids_backs()
    faces, backs = _pad_pair(raw_f, raw_b)
    path = build()
    print(f"Wrote {path}")
    print(f"Kids raw: {len(raw_f)} faces, {len(raw_b)} backs")
    print(f"Printed cards (padded for duplex): {len(faces)}")
    print(f"Slide {SLIDE_W_MM:.2f}x{SLIDE_H_MM:.2f} mm; margins >= {PRINTER_MARGIN_MM} mm")
    print(
        f"Cards {CARD_W_MM:.3f}x{CARD_H_MM:.3f} mm "
        f"(design {DESIGN_CARD_W_MM}x{DESIGN_CARD_H_MM}, scale {SCALE:.4f})"
    )
    print(f"Font {MAIN_PT} pt; faces bottom, backs top; requests only")
    print(f"Slides: {len(_chunks(faces, PER_SLIDE)) * 2} (face/back pairs, no utilities)")
