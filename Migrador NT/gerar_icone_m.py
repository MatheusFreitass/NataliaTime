"""
Gera M_icone.ico com o novo design:
  fundo #161b22, borda verde #3fb950, letra M branca bold (Segoe UI).
"""

from PIL import Image, ImageDraw, ImageFont
import os

LETTER       = "M"
BG           = (22, 27, 34, 255)       # #161b22
BORDER       = (63, 185, 80, 255)      # #3fb950
TEXT         = (255, 255, 255, 255)
FONT_PATH    = "C:/Windows/Fonts/segoeuib.ttf"
SIZES        = [256, 64, 48, 32, 16]
OUTPUT       = os.path.join(os.path.dirname(__file__), "M_icone.ico")


def make_frame(size):
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    radius   = max(3, int(size * 0.165))
    border_w = max(2, int(size * 0.025))

    # Fundo arredondado
    draw.rounded_rectangle([0, 0, size - 1, size - 1],
                           radius=radius, fill=BG)

    # Borda
    half = border_w / 2
    draw.rounded_rectangle([half, half, size - 1 - half, size - 1 - half],
                           radius=radius, outline=BORDER, width=border_w)

    # Letra
    font_size = int(size * 0.60)
    try:
        font = ImageFont.truetype(FONT_PATH, font_size)
    except Exception:
        font = ImageFont.load_default()

    bbox   = draw.textbbox((0, 0), LETTER, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = (size - text_w) // 2 - bbox[0]
    y = (size - text_h) // 2 - bbox[1]
    draw.text((x, y), LETTER, fill=TEXT, font=font)

    return img


frames = [make_frame(s) for s in SIZES]
frames[0].save(OUTPUT, format="ICO", sizes=[(s, s) for s in SIZES],
               append_images=frames[1:])
print(f"Ícone salvo em: {OUTPUT}")
