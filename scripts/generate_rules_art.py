"""PNG samples of PnP components for the rules booklet."""
from __future__ import annotations

import random
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "assets" / "rules"

DPI = 200
MM = DPI / 25.4
CARD_W = int(round(57.0 * MM))
CARD_H = int(round(44.1 * MM))
PAD = int(round(2.0 * MM))

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


def blank_card() -> Image.Image:
    img = Image.new("RGBA", (CARD_W, CARD_H), WHITE)
    draw = ImageDraw.Draw(img)
    draw.rectangle((0, 0, CARD_W - 1, CARD_H - 1), outline=BORDER, width=2)
    return img


def request_card(text: str, *, top: bool) -> Image.Image:
    img = blank_card()
    draw = ImageDraw.Draw(img)
    inner = CARD_W - 2 * PAD
    px = fit_size(text, inner, 50, bold=True)
    font = _font("arialbd.ttf", px)
    lines = wrap_lines(text, font, inner)
    line_h = px + 4
    block_h = line_h * len(lines)
    y = PAD if top else CARD_H - PAD - block_h
    for line in lines:
        w = _word_width(line, font)
        x = (CARD_W - w) / 2
        draw.text((x, y), line, font=font, fill=INK)
        y += line_h
    return img


def letter_card(ch: str) -> Image.Image:
    img = blank_card()
    draw = ImageDraw.Draw(img)
    text = ch.upper()
    inner = CARD_W - 2 * PAD
    px = fit_size(text, inner, 96, bold=True)
    font = _font("arialbd.ttf", px)
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (CARD_W - tw) / 2 - bbox[0]
    y = (CARD_H - th) / 2 - bbox[1]
    draw.text((x, y), text, font=font, fill=INK)
    return img


def vote_card(color: tuple[int, int, int, int], player: int) -> Image.Image:
    img = blank_card()
    draw = ImageDraw.Draw(img)
    inner = CARD_W - 2 * PAD
    title_px = fit_size(VOTE_TITLE, inner, 44, bold=True)
    title_font = _font("arialbd.ttf", title_px)
    sub_font = _font("arial.ttf", 16)
    lines = wrap_lines(VOTE_TITLE, title_font, inner)
    sub = f"Игрок {player}"
    line_h = title_px + 2
    sub_h = 20
    block_h = line_h * len(lines) + 8 + sub_h
    y = (CARD_H - block_h) / 2
    for line in lines:
        w = _word_width(line, title_font)
        draw.text(((CARD_W - w) / 2, y), line, font=title_font, fill=color)
        y += line_h
    y += 8
    sw = _word_width(sub, sub_font)
    draw.text(((CARD_W - sw) / 2, y), sub, font=sub_font, fill=color)
    return img


def _paste_rotated(scene: Image.Image, card: Image.Image, cx: int, cy: int, angle: float) -> None:
    rot = card.rotate(angle, expand=True, resample=Image.Resampling.BICUBIC)
    x = int(cx - rot.width / 2)
    y = int(cy - rot.height / 2)
    scene.alpha_composite(rot, (x, y))


def setup_table() -> Image.Image:
    rng = random.Random(42)
    w, h = int(180 * MM), int(115 * MM)
    scene = Image.new("RGBA", (w, h), TABLE)
    letters = ["С", "В", "Е", "Т", "М", "О", "Й", "З", "Р", "К", "А", "Л"]
    back = request_card(BACK_SAMPLE, top=True)
    # request deck (backs visible, slight stack)
    deck_x, deck_y = int(w * 0.50), int(h * 0.22)
    for i in range(3):
        _paste_rotated(scene, back, deck_x + i * 4, deck_y - i * 3, 0)
    # open request face to the left of deck (after first round would look like this;
    # for setup only the deck is there — keep just the deck as start)
    # scattered letters
    positions = [
        (0.22, 0.48, -18),
        (0.34, 0.62, 12),
        (0.46, 0.55, -8),
        (0.58, 0.68, 22),
        (0.70, 0.52, -14),
        (0.28, 0.78, 6),
        (0.42, 0.82, -22),
        (0.56, 0.86, 10),
        (0.68, 0.76, -6),
        (0.78, 0.64, 16),
        (0.18, 0.64, 8),
        (0.82, 0.80, -12),
    ]
    scale = 0.72
    small_letters = [
        letter_card(ch).resize(
            (int(CARD_W * scale), int(CARD_H * scale)), Image.Resampling.LANCZOS
        )
        for ch in letters
    ]
    for (px, py, ang), card in zip(positions, small_letters):
        jitter = rng.randint(-8, 8)
        _paste_rotated(scene, card, int(w * px) + jitter, int(h * py), ang)

    vote_scale = 0.62
    votes = [
        (vote_card(VOTE_RED, 1), 0.12, 0.18, -8),
        (vote_card(VOTE_BLUE, 2), 0.88, 0.18, 9),
        (vote_card(VOTE_GREEN, 3), 0.12, 0.90, 7),
        (vote_card(VOTE_ORANGE, 4), 0.88, 0.90, -11),
    ]
    for card, px, py, ang in votes:
        small = card.resize(
            (int(CARD_W * vote_scale), int(CARD_H * vote_scale)),
            Image.Resampling.LANCZOS,
        )
        _paste_rotated(scene, small, int(w * px), int(h * py), ang)

    # tiny caption-free: that's the point
    rgb = Image.new("RGB", scene.size, (255, 255, 255))
    rgb.paste(scene, mask=scene.split()[3])
    return rgb


def save_rgba(img: Image.Image, path: Path) -> None:
    rgb = Image.new("RGB", img.size, (255, 255, 255))
    rgb.paste(img, mask=img.split()[3] if img.mode == "RGBA" else None)
    rgb.save(path, "PNG")


def build() -> Path:
    OUT.mkdir(parents=True, exist_ok=True)
    save_rgba(request_card(FACE_SAMPLE, top=False), OUT / "card-request-face.png")
    save_rgba(request_card(BACK_SAMPLE, top=True), OUT / "card-request-back.png")
    save_rgba(letter_card(LETTER_SAMPLE), OUT / "card-letter.png")
    save_rgba(blank_card(), OUT / "card-letter-joker.png")
    save_rgba(vote_card(VOTE_RED, 1), OUT / "card-vote.png")
    setup_table().save(OUT / "setup-table.png", "PNG")
    return OUT


if __name__ == "__main__":
    path = build()
    print(f"Wrote PNGs in {path}")
