"""Generate assets/icon.ico and docs/images/icon*.png (requires Pillow).

    python tools/make_icon.py
"""

from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
S = 1024  # master resolution


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def master() -> Image.Image:
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    # Rounded-square background with a diagonal violet gradient.
    grad = Image.new("RGBA", (S, S))
    top, bottom = (155, 140, 255), (88, 70, 214)
    px = grad.load()
    for y in range(S):
        for x in range(S):
            px[x, y] = (*lerp(top, bottom, (x + y) / (2 * S)), 255)
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((40, 40, S - 40, S - 40), radius=220, fill=255)
    img.paste(grad, (0, 0), mask)
    d = ImageDraw.Draw(img)

    # Form sheet with a folded corner.
    sx1, sy1, sx2, sy2, fold = 230, 180, 700, 830, 120
    d.rounded_rectangle((sx1, sy1, sx2, sy2), radius=48, fill=(255, 255, 255, 255))
    d.polygon([(sx2 - fold, sy1), (sx2 + 2, sy1), (sx2 + 2, sy1 + fold)], fill=(0, 0, 0, 0))
    d.polygon([(sx2 - fold, sy1), (sx2 - fold, sy1 + fold), (sx2, sy1 + fold)], fill=(214, 208, 255, 255))

    # Form rows: check boxes + lines.
    for i, y in enumerate((330, 470, 610)):
        d.rounded_rectangle((300, y - 34, 368, y + 34), radius=14, fill=(139, 124, 255, 255))
        d.line([(316, y), (332, y + 16), (354, y - 14)], fill=(255, 255, 255, 255), width=14, joint="curve")
        d.rounded_rectangle((400, y - 14, 620 - (i % 2) * 70, y + 14), radius=14, fill=(196, 202, 216, 255))

    # Green download badge.
    cx, cy, r = 720, 720, 200
    d.ellipse((cx - r - 22, cy - r - 22, cx + r + 22, cy + r + 22), fill=(88, 70, 214, 255))
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(69, 212, 131, 255))
    w = 26
    d.line([(cx, cy - 110), (cx, cy + 50)], fill="white", width=w * 2)
    d.polygon([(cx - 92, cy - 5), (cx + 92, cy - 5), (cx, cy + 90)], fill="white")
    d.rounded_rectangle((cx - 110, cy + 112, cx + 110, cy + 140), radius=14, fill="white")
    return img


def main():
    img = master()
    sizes = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256]
    img.resize((256, 256), Image.LANCZOS).save(ROOT / "assets/icon.ico",
                                               sizes=[(s, s) for s in sizes])
    out = ROOT / "docs/images"
    out.mkdir(parents=True, exist_ok=True)
    img.resize((512, 512), Image.LANCZOS).save(out / "icon.png")
    img.resize((128, 128), Image.LANCZOS).save(out / "icon-128.png")
    print("Icons written.")


if __name__ == "__main__":
    main()
