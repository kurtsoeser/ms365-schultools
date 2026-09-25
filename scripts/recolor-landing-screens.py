"""Recolor landing screenshots: shift purple/violet brand hues toward teal."""
from __future__ import annotations

import colorsys
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
SCREENS = ROOT / "landing" / "assets" / "screens"

# Hue ranges in [0,1] for purple/violet/indigo family to remap
# Purple ~ 0.72–0.85, blue-violet ~ 0.65–0.72
PURPLE_H_MIN = 0.62
PURPLE_H_MAX = 0.92

# Target teal/cyan hue (~0.48–0.52 for #0d9f8a)
TEAL_H = 0.48


def remap_pixel(r: int, g: int, b: int, a: int) -> tuple[int, int, int, int]:
    if a == 0:
        return r, g, b, a
    # skip near-gray / near-white / near-black (UI chrome text)
    mx, mn = max(r, g, b), min(r, g, b)
    if mx < 28 or mn > 245:
        return r, g, b, a
    sat_probe = (mx - mn) / (mx + 1e-6)
    if sat_probe < 0.12:
        return r, g, b, a

    h, s, v = colorsys.rgb_to_hsv(r / 255.0, g / 255.0, b / 255.0)
    if not (PURPLE_H_MIN <= h <= PURPLE_H_MAX):
        return r, g, b, a
    if s < 0.12:
        return r, g, b, a

    # Map purple band into teal, keep relative position in band lightly
    t = (h - PURPLE_H_MIN) / (PURPLE_H_MAX - PURPLE_H_MIN)
    # slight variation around teal
    new_h = (TEAL_H - 0.04) + t * 0.08
    new_h %= 1.0
    # slightly boost sat for brand punch on headers
    new_s = min(1.0, s * 1.05 + 0.02)
    nr, ng, nb = colorsys.hsv_to_rgb(new_h, new_s, v)
    return int(nr * 255), int(ng * 255), int(nb * 255), a


def recolor(path: Path) -> None:
    im = Image.open(path).convert("RGBA")
    px = im.load()
    w, h = im.size
    for y in range(h):
        for x in range(w):
            px[x, y] = remap_pixel(*px[x, y])
    im.save(path, optimize=True)
    print(f"recolored {path.name} ({w}x{h})")


def main() -> None:
    files = sorted(SCREENS.glob("*.png"))
    if not files:
        raise SystemExit(f"No screens in {SCREENS}")
    for f in files:
        recolor(f)
    print(f"done, {len(files)} files")


if __name__ == "__main__":
    main()
