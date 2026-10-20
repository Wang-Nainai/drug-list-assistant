"""生成应用图标与空状态插画（无需外部素材）。"""

from __future__ import annotations

import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

ASSETS = Path(__file__).resolve().parent / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)


def rounded_linear_gradient(
    size: tuple[int, int],
    radius: int,
    top: tuple[int, int, int],
    bottom: tuple[int, int, int],
) -> Image.Image:
    width, height = size
    gradient = Image.new("RGBA", size)
    pixels = gradient.load()
    for y in range(height):
        t = y / max(height - 1, 1)
        colour = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3)) + (255,)
        for x in range(width):
            pixels[x, y] = colour

    mask = Image.new("L", size, 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, width - 1, height - 1), radius=radius, fill=255)
    gradient.putalpha(mask)
    return gradient


def draw_shadow(base: Image.Image, box: tuple[int, int, int, int], radius: int, alpha: int = 45, blur: int = 18) -> None:
    shadow = Image.new("RGBA", base.size, (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle(box, radius=radius, fill=(12, 40, 70, alpha))
    shadow = shadow.filter(ImageFilter.GaussianBlur(blur))
    base.alpha_composite(shadow)


def make_logo() -> None:
    scale = 3
    size = 512 * scale
    icon = rounded_linear_gradient((size, size), 112 * scale, (18, 118, 196), (10, 90, 160))
    draw = ImageDraw.Draw(icon)

    # Subtle highlight in the upper-left corner.
    glow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse(
        (-size * 0.25, -size * 0.45, size * 0.65, size * 0.42),
        fill=(255, 255, 255, 28),
    )
    icon.alpha_composite(glow)

    # Document card.
    doc_box = (126 * scale, 92 * scale, 386 * scale, 414 * scale)
    draw_shadow(icon, (doc_box[0] + 6 * scale, doc_box[1] + 14 * scale, doc_box[2] + 6 * scale, doc_box[3] + 14 * scale), 30 * scale, 70, 14 * scale)
    draw.rounded_rectangle(doc_box, radius=30 * scale, fill=(255, 255, 255, 255))

    line_colour = (15, 108, 189, 255)
    soft_colour = (15, 108, 189, 95)
    draw.rounded_rectangle((166 * scale, 162 * scale, 346 * scale, 180 * scale), radius=9 * scale, fill=line_colour)
    draw.rounded_rectangle((166 * scale, 218 * scale, 320 * scale, 236 * scale), radius=9 * scale, fill=line_colour)
    draw.rounded_rectangle((166 * scale, 274 * scale, 346 * scale, 292 * scale), radius=9 * scale, fill=soft_colour)
    draw.rounded_rectangle((166 * scale, 330 * scale, 286 * scale, 348 * scale), radius=9 * scale, fill=soft_colour)

    # Check badge.
    badge_center = (372 * scale, 372 * scale)
    badge_radius = 79 * scale
    draw.ellipse(
        (
            badge_center[0] - badge_radius,
            badge_center[1] - badge_radius,
            badge_center[0] + badge_radius,
            badge_center[1] + badge_radius,
        ),
        fill=(76, 194, 255, 255),
        outline=(255, 255, 255, 255),
        width=9 * scale,
    )
    check_points = [
        (334 * scale, 373 * scale),
        (362 * scale, 401 * scale),
        (414 * scale, 341 * scale),
    ]
    draw.line(check_points, fill=(6, 78, 140, 255), width=19 * scale, joint="curve")

    icon = icon.resize((512, 512), Image.LANCZOS)
    icon.save(ASSETS / "logo.png")

    icon.save(
        ASSETS / "app.ico",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )


def draw_document(
    layer: Image.Image,
    box: tuple[int, int, int, int],
    fill: tuple[int, int, int, int],
    line: tuple[int, int, int, int],
    accent: tuple[int, int, int, int],
    row_highlight: tuple[int, int, int, int],
    radius: int = 20,
) -> None:
    draw = ImageDraw.Draw(layer)
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=None)
    x1, y1, x2, _ = box
    inner_left = x1 + 28
    inner_right = x2 - 28
    widths = [1.0, 0.72, 0.9, 0.58]
    for index in range(4):
        y = y1 + 70 + index * 38
        width = int((inner_right - inner_left) * widths[index])
        draw.rounded_rectangle(
            (inner_left, y, inner_left + width, y + 10),
            radius=5,
            fill=line,
        )
    draw.rounded_rectangle(
        (inner_left, y1 + 60, inner_right, y1 + 72),
        radius=6,
        fill=accent,
    )


def make_illustration(name: str, dark: bool) -> None:
    scale = 2
    width, height = 520 * scale, 340 * scale
    image = Image.new("RGBA", (width, height), (0, 0, 0, 0))

    if dark:
        circle_fill = (29, 51, 68, 255)
        doc_fill = (46, 46, 46, 255)
        doc_fill_back = (38, 38, 38, 255)
        line = (91, 108, 122, 255)
        accent = (76, 194, 255, 255)
        soft = (35, 52, 66, 255)
        highlight = (41, 74, 96, 255)
        badge = (76, 194, 255, 255)
        check = (11, 48, 76, 255)
    else:
        circle_fill = (234, 243, 252, 255)
        doc_fill = (255, 255, 255, 255)
        doc_fill_back = (248, 251, 254, 255)
        line = (205, 220, 234, 255)
        accent = (15, 108, 189, 255)
        soft = (228, 240, 250, 255)
        highlight = (234, 243, 252, 255)
        badge = (15, 108, 189, 255)
        check = (255, 255, 255, 255)

    draw = ImageDraw.Draw(image)
    draw.ellipse((40 * scale, 34 * scale, 360 * scale, 314 * scale), fill=circle_fill)

    # Decorative dots.
    draw.ellipse((438 * scale, 58 * scale, 462 * scale, 82 * scale), fill=accent)
    draw.ellipse((60 * scale, 252 * scale, 76 * scale, 268 * scale), fill=accent)
    draw.ellipse((408 * scale, 278 * scale, 424 * scale, 294 * scale), fill=highlight)

    # Back document, slightly offset.
    back = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw_document(
        back,
        (128 * scale, 50 * scale, 338 * scale, 284 * scale),
        doc_fill_back,
        line,
        soft,
        highlight,
    )
    back = back.rotate(-5, resample=Image.BICUBIC, center=(233 * scale, 167 * scale))
    image.alpha_composite(back)

    # Front document with shadow.
    shadow = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle(
        (194 * scale, 86 * scale, 406 * scale, 322 * scale),
        radius=20 * scale,
        fill=(15, 60, 100, 38 if not dark else 80),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(12 * scale))
    image.alpha_composite(shadow)

    front = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    draw_document(
        front,
        (184 * scale, 76 * scale, 394 * scale, 312 * scale),
        doc_fill,
        line,
        accent,
        highlight,
    )
    front = front.rotate(3, resample=Image.BICUBIC, center=(289 * scale, 194 * scale))
    image.alpha_composite(front)

    # Check badge.
    cx, cy, radius = 386 * scale, 282 * scale, 54 * scale
    draw.ellipse(
        (cx - radius, cy - radius, cx + radius, cy + radius),
        fill=badge,
        outline=(255, 255, 255, 255) if not dark else (38, 38, 38, 255),
        width=6 * scale,
    )
    draw.line(
        [
            (361 * scale, 282 * scale),
            (379 * scale, 300 * scale),
            (411 * scale, 264 * scale),
        ],
        fill=check,
        width=12 * scale,
        joint="curve",
    )

    image = image.resize((260, 170), Image.LANCZOS)
    image.save(ASSETS / name)


def main() -> None:
    make_logo()
    make_illustration("empty_light.png", dark=False)
    make_illustration("empty_dark.png", dark=True)
    print("assets written to", ASSETS)


if __name__ == "__main__":
    main()
