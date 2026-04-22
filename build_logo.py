"""
Build amazon_logo_header.png using Chrome headless to rasterize the official
Amazon wordmark at its native SVG dimensions, then compose onto Squid Ink.

This avoids the CSS-flex-scaling bug that was dropping path segments.

Output:
  - amazon_logo_white.svg         (white-tinted version — for HTML inline use)
  - amazon_logo_header.png        (1024 x 320, dark bg + centered wordmark)
"""
import os
import re
import subprocess
import tempfile
from PIL import Image

CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
SVG_SRC = "amazon_logo.svg"
SVG_WHITE = "amazon_logo_white.svg"
OUT = "amazon_logo_header.png"
BG = (0x16, 0x1D, 0x26)

CANVAS_W, CANVAS_H = 1024, 320
LOGO_W_RATIO = 0.55

# Native SVG dimensions (from amazon_logo.svg attributes)
SVG_W, SVG_H = 603, 182
RENDER_SCALE = 4  # render at 4x then downsample for crispness


def make_white_svg():
    with open(SVG_SRC, "r", encoding="utf-8") as f:
        svg = f.read()
    svg = re.sub(r'fill="[^"]*"', 'fill="#FFFFFF"', svg)
    svg = re.sub(r'fill:#[0-9A-Fa-f]{3,6}', 'fill:#FFFFFF', svg)
    svg = re.sub(r'style="fill:[^"]*"', 'style="fill:#FFFFFF"', svg)
    with open(SVG_WHITE, "w", encoding="utf-8") as f:
        f.write(svg)
    print(f"Wrote white SVG: {SVG_WHITE}")


def rasterize_svg_native():
    """Render the white SVG at N-x-multiplied intrinsic size via Chrome
    headless. Returns a PIL Image with transparent background."""
    tmp = tempfile.mkdtemp(prefix="svg2png_")
    png_path = os.path.join(tmp, "raw.png")

    w = SVG_W * RENDER_SCALE
    h = SVG_H * RENDER_SCALE

    # Use a tiny HTML wrapper that forces the svg to render at w x h in pixels
    with open(SVG_WHITE, "r", encoding="utf-8") as f:
        svg_markup = f.read()
    svg_markup = re.sub(r"<\?xml[^>]+\?>\s*", "", svg_markup)
    # Force the outer <svg> element's width/height so Chrome rasterizes at Nx
    svg_markup = re.sub(
        r'(<svg\b[^>]*?)width="\d+"', f'\\1width="{w}"', svg_markup, count=1
    )
    svg_markup = re.sub(
        r'(<svg\b[^>]*?)height="\d+"', f'\\1height="{h}"', svg_markup, count=1
    )

    html = f"""<!doctype html><html><head><meta charset="utf-8">
<style>html,body {{ margin:0; padding:0; background:transparent; }}</style>
</head><body>{svg_markup}</body></html>"""
    html_path = os.path.join(tmp, "logo.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)

    cmd = [
        CHROME,
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        f"--window-size={w},{h}",
        f"--screenshot={png_path}",
        "--default-background-color=00000000",
        f"file:///{html_path.replace(os.sep, '/')}",
    ]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    if r.returncode:
        print("Chrome stderr:", r.stderr)
    if not os.path.exists(png_path):
        raise RuntimeError("Chrome produced no PNG")

    img = Image.open(png_path).convert("RGBA")
    # Crop to non-transparent content bbox to remove any empty border
    bbox = img.getbbox()
    if bbox:
        img = img.crop(bbox)
    return img


def main():
    if not os.path.exists(SVG_SRC):
        raise FileNotFoundError(f"{SVG_SRC} not found")
    if not os.path.exists(CHROME):
        raise FileNotFoundError(f"Chrome not found at {CHROME}")

    make_white_svg()
    logo = rasterize_svg_native()
    print(f"Rasterized logo: {logo.size}")

    # Resize to target width preserving aspect
    target_w = int(CANVAS_W * LOGO_W_RATIO)
    scale = target_w / logo.size[0]
    target_h = int(logo.size[1] * scale)
    logo = logo.resize((target_w, target_h), Image.LANCZOS)

    # Compose onto Squid Ink canvas
    canvas = Image.new("RGB", (CANVAS_W, CANVAS_H), BG)
    x = (CANVAS_W - target_w) // 2
    y = (CANVAS_H - target_h) // 2
    canvas.paste(logo, (x, y), logo)
    canvas.save(OUT, "PNG", optimize=True)
    print(f"Saved {OUT} size={canvas.size} logo={logo.size} at=({x},{y})")


if __name__ == "__main__":
    main()
