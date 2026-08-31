"""PNG samples of PnP components for the rules booklet."""
from __future__ import annotations

import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets" / "rules"
ASSETS = ROOT / "assets"

DPI = 200
MM = DPI / 25.4

REQ_W = int(round(88.0 * MM))
REQ_H = int(round(63.0 * MM))
LET_W = int(round(44.1 * MM))
LET_H = int(round(57.0 * MM))
VOTE_W = REQ_W
VOTE_H = REQ_H

FRAME_SIDE_FR = 0.11
FRAME_EDGE_MM = 3.5
LETTER_PAD_MM = 2.0
VOTE_PAD_MM = 2.4
STACK_GAP_MM = 3.0

BORDER = (0xC0, 0xC0, 0xC0, 255)
INK = (0x1A, 0x1A, 0x1A, 255)
TABLE = (0xF4, 0xF1, 0xEA, 255)
WHITE = (255, 255, 255, 255)

VOTE_RED = (0xC6, 0x28, 0x28, 255)
VOTE_BLUE = (0x15, 0x65, 0xC0, 255)
VOTE_GREEN = (0x2E, 0x7D, 0x32, 255)
VOTE_ORANGE = (0xEF, 0x6C, 0x00, 255)

FACE_SAMPLE = "Расскажи чем испугать всех"
BACK_SAMPLE = "соседей сверху"
LETTER_SAMPLE = "С"
VOTE_TITLE = "Ты восхитителен"

FACE_FRAME = ASSETS / "request-face-frame.png"
BACK_FRAME = ASSETS / "request-back-frame.png"
FONT_DIR = Path(r"C:\Windows\Fonts")


def _font(name: str, size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    path = FONT_DIR / name
    try:
        return ImageFont.truetype(str(path), size)
    except OSError:
        return ImageFont.load_default()


def _word_width(word: str, font: ImageFont.ImageFont) -> float:
    return float(font.getlength(word))


def fit_size(text: str, max_w: int, max_px: int, *, bold: bool = True, min_px: int = 11) -> int:
    words = [w for w in text.replace("\n", " ").split(" ") if w]
    if not words:
        return max_px
    name = "arialbd.ttf" if bold else "arial.ttf"
    size = max_px
    while size > min_px:
        font = _font(name, size)
        if max(_word_width(w, font) for w in words) <= max_w:
            return size
        size -= 1
    return min_px


def wrap_lines(text: str, font: ImageFont.ImageFont, max_w: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    cur = ""
    for w in words:
        trial = w if not cur else f"{cur} {w}"
        if _word_width(trial, font) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [text]


def _stroke(draw: ImageDraw.ImageDraw, w: int, h: int) -> None:
    draw.rectangle((0, 0, w - 1, h - 1), outline=BORDER, width=2)


def _blank(w: int, h: int) -> Image.Image:
    img = Image.new("RGBA", (w, h), WHITE)
    _stroke(ImageDraw.Draw(img), w, h)
    return img


def request_card(text: str, *, face: bool) -> Image.Image:
    src = FACE_FRAME if face else BACK_FRAME
    base = Image.open(src).convert("RGBA").resize((REQ_W, REQ_H), Image.Resampling.LANCZOS)
    img = Image.new("RGBA", (REQ_W, REQ_H), WHITE)
    img.alpha_composite(base)
    draw = ImageDraw.Draw(img)
    _stroke(draw, REQ_W, REQ_H)

    side = int(round(REQ_W * FRAME_SIDE_FR))
    edge = int(round(FRAME_EDGE_MM * MM))
    inner_w = REQ_W - 2 * side
    px = fit_size(text, inner_w, 56, bold=True)
    font = _font("arialbd.ttf", px)
    lines = wrap_lines(text, font, inner_w)
    line_h = px + 4
    block_h = line_h * len(lines)
    if face:
        y = REQ_H - edge - block_h
    else:
        y = edge
    for line in lines:
        tw = _word_width(line, font)
        x = (REQ_W - tw) / 2
        draw.text((x, y), line, font=font, fill=INK)
        y += line_h
    return img


def letter_card(ch: str) -> Image.Image:
    img = _blank(LET_W, LET_H)
    draw = ImageDraw.Draw(img)
    text = ch.upper()
    pad = int(round(LETTER_PAD_MM * MM))
    inner_w = LET_W - 2 * pad
    inner_h = LET_H - 2 * pad
    # pt→px at this DPI; start from a high cap and shrink to ink bbox.
    max_px = int(round(inner_h * 0.92))
    size = max_px
    while size > 12:
        font = _font("arialbd.ttf", size)
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        if tw <= inner_w and th <= inner_h:
            break
        size -= 1
    font = _font("arialbd.ttf", size)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (LET_W - tw) / 2 - bbox[0]
    y = (LET_H - th) / 2 - bbox[1]
    draw.text((x, y), text, font=font, fill=INK)
    return img


def vote_card(color: tuple[int, int, int, int], player: int) -> Image.Image:
    img = _blank(VOTE_W, VOTE_H)
    draw = ImageDraw.Draw(img)
    pad = int(round(VOTE_PAD_MM * MM))
    inner = VOTE_W - 2 * pad
    title_px = fit_size(VOTE_TITLE, inner, 56, bold=True)
    title_font = _font("arialbd.ttf", title_px)
    sub_font = _font("arial.ttf", 20)
    lines = wrap_lines(VOTE_TITLE, title_font, inner)
    sub = f"Игрок {player}"
    line_h = title_px + 2
    sub_h = 24
    block_h = line_h * len(lines) + 10 + sub_h
    y = (VOTE_H - block_h) / 2
    for line in lines:
        w = _word_width(line, title_font)
        draw.text(((VOTE_W - w) / 2, y), line, font=title_font, fill=color)
        y += line_h
    y += 10
    sw = _word_width(sub, sub_font)
    draw.text(((VOTE_W - sw) / 2, y), sub, font=sub_font, fill=color)
    return img


def request_pair(face: Image.Image, back: Image.Image) -> Image.Image:
    gap = int(round(STACK_GAP_MM * MM))
    img = Image.new("RGBA", (REQ_W, REQ_H * 2 + gap), (255, 255, 255, 0))
    img.paste(face, (0, 0))
    img.paste(back, (0, REQ_H + gap))
    return img


def card_stack(card: Image.Image, n: int = 4, dx: int = 7, dy: int = -5) -> Image.Image:
    """Offset pile, 3–5 cards. Front card on top."""
    n = max(3, min(5, n))
    w, h = card.size
    extra_x = abs(dx) * (n - 1)
    extra_y = abs(dy) * (n - 1)
    canvas = Image.new("RGBA", (w + extra_x, h + extra_y), (0, 0, 0, 0))
    for i in range(n - 1, -1, -1):
        x = i * dx if dx >= 0 else extra_x + i * dx
        y = (n - 1 - i) * abs(dy) if dy < 0 else i * dy
        canvas.alpha_composite(card, (x, y))
    return canvas


def duplex_pair(left: Image.Image, right: Image.Image) -> Image.Image:
    """Face (or letter) | ↔ | back, aligned to the bottom of the taller side."""
    gap = int(round(4 * MM))
    arrow_w = int(round(14 * MM))
    h = max(left.height, right.height)
    w = left.width + gap + arrow_w + gap + right.width
    img = Image.new("RGBA", (w, h), WHITE)
    ly = h - left.height
    ry = h - right.height
    img.alpha_composite(left.convert("RGBA"), (0, ly))
    img.alpha_composite(right.convert("RGBA"), (left.width + gap + arrow_w + gap, ry))
    draw = ImageDraw.Draw(img)
    font = _font("arialbd.ttf", int(round(18 * MM)))
    arrow = "↔"
    tw = _word_width(arrow, font)
    ax = left.width + gap + (arrow_w - tw) / 2
    ay = (h - 18 * MM) / 2
    draw.text((ax, ay), arrow, font=font, fill=INK)
    return img


def letter_answer_row() -> Image.Image:
    """Blank joker + О П А in one row (illustration for «Составление ответа»)."""
    cards = [_blank(LET_W, LET_H), letter_card("О"), letter_card("П"), letter_card("А")]
    gap = int(round(3 * MM))
    w = len(cards) * LET_W + (len(cards) - 1) * gap
    img = Image.new("RGBA", (w, LET_H), WHITE)
    x = 0
    for card in cards:
        img.alpha_composite(card.convert("RGBA"), (x, 0))
        x += LET_W + gap
    return img


def _rotated_aabb(
    cx: float, cy: float, w: float, h: float, angle_deg: float, pad: float = 0.0
) -> tuple[float, float, float, float]:
    import math

    a = math.radians(angle_deg)
    cos_a, sin_a = math.cos(a), math.sin(a)
    hw, hh = w / 2, h / 2
    xs, ys = [], []
    for x, y in ((-hw, -hh), (hw, -hh), (hw, hh), (-hw, hh)):
        xs.append(cx + x * cos_a - y * sin_a)
        ys.append(cy + x * sin_a + y * cos_a)
    return (min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad)


def _aabb_overlap(
    a: tuple[float, float, float, float], b: tuple[float, float, float, float]
) -> bool:
    return not (a[2] < b[0] or a[0] > b[2] or a[3] < b[1] or a[1] > b[3])


def _paste_rotated(scene: Image.Image, card: Image.Image, cx: int, cy: int, angle: float) -> None:
    rot = card.rotate(angle, expand=True, resample=Image.Resampling.BICUBIC)
    x = int(cx - rot.width / 2)
    y = int(cy - rot.height / 2)
    scene.alpha_composite(rot, (x, y))


def setup_table() -> Image.Image:
    rng = random.Random(42)
    w, h = int(200 * MM), int(130 * MM)
    scene = Image.new("RGBA", (w, h), TABLE)
    letters = ["С", "В", "Е", "Т", "М", "О", "Й", "З", "Р", "К", "А", "Л"]
    face = request_card(FACE_SAMPLE, face=True)
    back = request_card(BACK_SAMPLE, face=False)
    req_scale = 0.42
    req_size = (int(REQ_W * req_scale), int(REQ_H * req_scale))
    small_face = face.resize(req_size, Image.Resampling.LANCZOS)
    small_back = back.resize(req_size, Image.Resampling.LANCZOS)
    # Open request: face above remaining deck (backs), as in the booklet diagrams.
    deck_x, deck_y = int(w * 0.50), int(h * 0.28)
    for i in range(5):
        _paste_rotated(scene, small_back, deck_x + i * 4, deck_y - i * 3, 0)
    _paste_rotated(scene, small_face, deck_x, deck_y - req_size[1] - 6, 0)

    vote_scale = 0.34
    votes = [
        (vote_card(VOTE_RED, 1), 0.09, 0.11, -8),
        (vote_card(VOTE_BLUE, 2), 0.91, 0.11, 9),
        (vote_card(VOTE_GREEN, 3), 0.09, 0.90, 7),
        (vote_card(VOTE_ORANGE, 4), 0.91, 0.90, -11),
    ]
    vote_pad = int(round(6 * MM))
    vote_boxes = []
    for card, px, py, ang in votes:
        small = card.resize(
            (int(VOTE_W * vote_scale), int(VOTE_H * vote_scale)),
            Image.Resampling.LANCZOS,
        )
        cx, cy = int(w * px), int(h * py)
        _paste_rotated(scene, small, cx, cy, ang)
        vote_boxes.append(
            _rotated_aabb(cx, cy, small.width, small.height, ang, pad=vote_pad)
        )

    positions = [
        (0.42, 0.52, -18),
        (0.48, 0.58, 12),
        (0.54, 0.52, -8),
        (0.60, 0.58, 22),
        (0.51, 0.64, -14),
        (0.45, 0.64, 6),
        (0.57, 0.64, -22),
        (0.50, 0.58, 10),
        (0.55, 0.55, -6),
        (0.47, 0.55, 16),
        (0.43, 0.59, 8),
        (0.59, 0.61, -12),
    ]
    let_scale = 0.38
    small_letters = [
        letter_card(ch).resize(
            (int(LET_W * let_scale), int(LET_H * let_scale)), Image.Resampling.LANCZOS
        )
        for ch in letters
    ]
    for (px, py, ang), card in zip(positions, small_letters):
        jitter = rng.randint(-8, 8)
        cx = int(w * px) + jitter
        cy = int(h * py)
        box = _rotated_aabb(cx, cy, card.width, card.height, ang, pad=vote_pad)
        if any(_aabb_overlap(box, vb) for vb in vote_boxes):
            cx, cy = int(w * 0.50), int(h * 0.58)
            box = _rotated_aabb(cx, cy, card.width, card.height, ang, pad=vote_pad)
            if any(_aabb_overlap(box, vb) for vb in vote_boxes):
                continue
        _paste_rotated(scene, card, cx, cy, ang)

    rgb = Image.new("RGB", scene.size, (255, 255, 255))
    rgb.paste(scene, mask=scene.split()[3])
    return rgb


def save_rgba(img: Image.Image, path: Path) -> None:
    rgb = Image.new("RGB", img.size, (255, 255, 255))
    rgb.paste(img, mask=img.split()[3] if img.mode == "RGBA" else None)
    rgb.save(path, "PNG")


def build() -> Path:
    OUT.mkdir(parents=True, exist_ok=True)
    face = request_card(FACE_SAMPLE, face=True)
    back = request_card(BACK_SAMPLE, face=False)
    save_rgba(face, OUT / "card-request-face.png")
    save_rgba(back, OUT / "card-request-back.png")
    save_rgba(request_pair(face, back), OUT / "card-request-pair.png")
    face_stack = card_stack(face, 4)
    back_stack = card_stack(back, 4)
    save_rgba(duplex_pair(face_stack, back_stack), OUT / "card-request-duplex.png")
    letter = letter_card(LETTER_SAMPLE)
    joker = _blank(LET_W, LET_H)
    save_rgba(letter, OUT / "card-letter.png")
    save_rgba(joker, OUT / "card-letter-joker.png")
    save_rgba(
        duplex_pair(card_stack(letter, 4), card_stack(joker, 4)),
        OUT / "card-letter-duplex.png",
    )
    save_rgba(letter_answer_row(), OUT / "card-letter-answer.png")
    vote = vote_card(VOTE_RED, 1)
    save_rgba(vote, OUT / "card-vote.png")
    save_rgba(card_stack(vote, 4), OUT / "card-vote-stack.png")
    setup_table().save(OUT / "setup-table.png", "PNG")
    return OUT


if __name__ == "__main__":
    path = build()
    print(f"Wrote PNGs in {path}")
