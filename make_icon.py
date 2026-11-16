"""Generates icon.ico (two overlapping documents with a conversion arrow)."""
from pathlib import Path

from PIL import Image, ImageDraw

S = 512
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)
d.rounded_rectangle((16, 16, S - 16, S - 16), radius=110, fill=(37, 99, 235))
# back document
d.rounded_rectangle((96, 92, 296, 352), radius=26, fill=(191, 212, 254))
# front document
d.rounded_rectangle((216, 160, 416, 420), radius=26, fill=(255, 255, 255))
for y in (230, 280, 330):
    d.rounded_rectangle((254, y, 378, y + 18), radius=9, fill=(147, 180, 248))
# arrow
d.line((140, 400, 196, 400), fill=(255, 255, 255), width=26)
d.polygon([(190, 366), (236, 400), (190, 434)], fill=(255, 255, 255))

out = Path(__file__).parent / "icon.ico"
img.save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print(out)
