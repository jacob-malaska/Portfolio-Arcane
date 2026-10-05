"""Generate web copies without changing originals. Run with Python + Pillow.

Re-run after adding images to index.html. Adjust MAX_EDGE and QUALITY to taste.
"""
import html
import io
from pathlib import Path
import re

from PIL import Image, ImageOps

ROOT = Path(__file__).resolve().parents[1]
MAX_EDGE = 2400
QUALITY = 84


def main():
    page = ROOT / "index.html"
    source = page.read_text(encoding="utf-8")
    before = after = 0
    seen = set()

    def optimize(match):
        nonlocal before, after
        tag = match.group(0)
        src = re.search(r'\bsrc="([^"]+)"', tag)
        if not src or not src[1].startswith("images/"):
            return tag
        original = re.search(r'\bdata-original="([^"]+)"', tag)
        path = html.unescape(original[1] if original else src[1])
        input_path = ROOT / path
        with Image.open(input_path) as image:
            image = ImageOps.exif_transpose(image)
            image.thumbnail((MAX_EDGE, MAX_EDGE), Image.Resampling.LANCZOS)
            if image.mode not in ("RGB", "RGBA"):
                image = image.convert("RGBA" if "transparency" in image.info else "RGB")
            buffer = io.BytesIO()
            image.save(buffer, format="WEBP", quality=QUALITY, method=6)
            data = buffer.getvalue()
            # Include the source extension to prevent same-name PNG/JPG collisions.
            output = Path("images/optimized") / (input_path.name + ".webp")
            if len(data) < input_path.stat().st_size:
                (ROOT / output).parent.mkdir(parents=True, exist_ok=True)
                (ROOT / output).write_bytes(data)
                new_src = output.as_posix()
                width, height = image.size
                size = len(data)
            else:
                new_src = path
                with Image.open(input_path) as original_image:
                    width, height = ImageOps.exif_transpose(original_image).size
                size = input_path.stat().st_size
        if path not in seen:
            seen.add(path)
            before += input_path.stat().st_size
            after += size
        tag = re.sub(r'\s+(?:src|data-original|width|height|loading|decoding)="[^"]*"', '', tag)
        eager = input_path.name == "JRM-02.png"
        attrs = (f' src="{html.escape(new_src, quote=True)}"'
                 f' data-original="{html.escape(path, quote=True)}"'
                 f' width="{width}" height="{height}"'
                 f' loading="{"eager" if eager else "lazy"}" decoding="async"')
        return tag.replace("<img", "<img" + attrs, 1)

    source = re.sub(r'<img\b[^>]*>', optimize, source)
    page.write_text(source, encoding="utf-8")
    print(f"{len(seen)} unique images: {before / 1024**2:.2f} MiB -> {after / 1024**2:.2f} MiB "
          f"({(1 - after / before):.1%} smaller)")


if __name__ == "__main__":
    main()
