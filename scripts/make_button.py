# -*- coding: utf-8 -*-
"""
สร้าง PNG ปุ่ม "ดึงข้อมูลล่าสุด" แล้วพิมพ์เป็น base64 สำหรับฝังใน apps_script/Code.gs

Apps Script วางปุ่มกดจริงได้ด้วย Sheet.insertImage(blob).assignScript('syncFromButton')
แต่รูปต้องมาจาก blob จึงต้อง hardcode base64 ไว้ในซอร์ส

ผลลัพธ์:
  build/sync_button.png       ไว้ดูด้วยตา
  build/sync_button.gs.txt    ก๊อปไปแทนค่า SYNC_BUTTON_PNG_B64 ใน Code.gs

  python scripts/make_button.py
"""
import base64
import io
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "build"

W, H = 520, 88                  # เรนเดอร์ 2x แล้วย่อแสดงที่ 260x44 ในชีต
BG = (31, 56, 100)              # #1f3864 สีเดียวกับหัวตาราง
FG = (255, 255, 255)
RADIUS = 14
TEXT = "ดึงข้อมูลล่าสุด"

# Leelawadee UI มี glyph ไทยครบและวางสระ/วรรณยุกต์ถูกโดยไม่ต้องพึ่ง HarfBuzz
FONT = r"C:\Windows\Fonts\LeelawUI.ttf"


def refresh_icon(d, cx, cy, r, width=6):
    """ลูกศรวงกลม วาดเอง — Leelawadee ไม่มี glyph ↻ จะขึ้นเป็นกล่องว่าง"""
    d.arc([cx - r, cy - r, cx + r, cy + r], start=40, end=330, fill=FG, width=width)
    a = math.radians(330)
    tx, ty = cx + r * math.cos(a), cy + r * math.sin(a)
    s = r * 0.55
    d.polygon(
        [(tx + s * 0.1, ty - s * 0.9), (tx + s * 0.95, ty + s * 0.15), (tx - s * 0.7, ty + s * 0.35)],
        fill=FG,
    )


def main():
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, W - 1, H - 1], radius=RADIUS, fill=BG)

    font = ImageFont.truetype(FONT, 40)
    box = d.textbbox((0, 0), TEXT, font=font)
    tw = box[2] - box[0]

    icon_r, gap = 16, 22
    x0 = (W - (icon_r * 2 + gap + tw)) / 2
    refresh_icon(d, x0 + icon_r, H / 2, icon_r)
    d.text((x0 + icon_r * 2 + gap - box[0], (H - (box[3] - box[1])) / 2 - box[1]), TEXT, font=font, fill=FG)

    OUT.mkdir(exist_ok=True)
    img.save(OUT / "sync_button.png")

    # RGBA quantize ได้เฉพาะ FASTOCTREE — ลดจาก ~6KB เหลือ ~1.8KB
    buf = io.BytesIO()
    img.quantize(colors=16, method=Image.Quantize.FASTOCTREE).save(buf, format="PNG", optimize=True)
    raw = buf.getvalue()

    plain = io.BytesIO()
    img.save(plain, format="PNG", optimize=True)
    if len(plain.getvalue()) < len(raw):
        raw = plain.getvalue()

    b64 = base64.b64encode(raw).decode("ascii")
    lines = [b64[i:i + 96] for i in range(0, len(b64), 96)]
    gs = "const SYNC_BUTTON_PNG_B64 = [\n" + "\n".join(f"  '{ln}'," for ln in lines) + "\n].join('');\n"
    (OUT / "sync_button.gs.txt").write_text(gs, encoding="utf-8")

    print(f"PNG {len(raw)} bytes -> base64 {len(b64)} chars ({len(lines)} บรรทัด)")
    print(f"เขียน {OUT / 'sync_button.png'}")
    print(f"เขียน {OUT / 'sync_button.gs.txt'}")


if __name__ == "__main__":
    main()
