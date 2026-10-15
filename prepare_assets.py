"""Generate Rainy Teacher's PNG and Windows ICO from its checked-in PNG payload."""
from __future__ import annotations

import base64
import io
from pathlib import Path

from PIL import Image

ASSETS = Path(__file__).resolve().parent / "assets"
source = ASSETS / "rainy_teacher_mark.png.b64"
png = ASSETS / "rainy_teacher_mark.png"
ico = ASSETS / "rainy_teacher.ico"

image_bytes = base64.b64decode(source.read_text(encoding="ascii"))
with Image.open(io.BytesIO(image_bytes)) as image:
    image = image.convert("RGBA")
    image.save(png, format="PNG")
    image.save(ico, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])

print(f"Prepared {png.name} and {ico.name}")
