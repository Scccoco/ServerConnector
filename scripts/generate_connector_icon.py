from PIL import Image, ImageDraw, ImageFilter


def build_icon(path: str) -> None:
    size = 1024
    gradient = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw_gradient = ImageDraw.Draw(gradient)
    for y in range(size):
        t = y / (size - 1)
        r = int(8 * (1 - t) + 20 * t)
        g = int(24 * (1 - t) + 78 * t)
        b = int(42 * (1 - t) + 116 * t)
        draw_gradient.line([(0, y), (size, y)], fill=(r, g, b, 255))

    mask = Image.new("L", (size, size), 0)
    mask_draw = ImageDraw.Draw(mask)
    mask_draw.rounded_rectangle((48, 48, size - 48, size - 48), radius=220, fill=255)
    icon = Image.composite(gradient, Image.new("RGBA", (size, size), (0, 0, 0, 0)), mask)

    ring_glow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ring_draw = ImageDraw.Draw(ring_glow)
    ring_draw.ellipse((180, 180, 844, 844), outline=(106, 224, 255, 210), width=42)
    ring_glow = ring_glow.filter(ImageFilter.GaussianBlur(3))
    icon = Image.alpha_composite(icon, ring_glow)

    draw = ImageDraw.Draw(icon)
    draw.ellipse((220, 220, 804, 804), outline=(188, 240, 255, 235), width=26)

    nodes = [(300, 512), (512, 300), (724, 512), (512, 724)]
    for x, y in nodes:
        draw.ellipse((x - 52, y - 52, x + 52, y + 52), fill=(15, 35, 55, 255), outline=(146, 234, 255, 255), width=14)

    links = [
        (352, 512, 460, 352),
        (564, 352, 672, 512),
        (672, 512, 564, 672),
        (460, 672, 352, 512),
    ]
    for line in links:
        draw.line(line, fill=(146, 234, 255, 235), width=26)

    draw.ellipse((426, 426, 598, 598), fill=(10, 50, 70, 255), outline=(186, 248, 255, 255), width=16)

    reduced = icon.resize((512, 512), Image.Resampling.LANCZOS)
    reduced.save(
        path,
        format="ICO",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )


if __name__ == "__main__":
    build_icon(r"E:\00_Cursor\18_Server\connector-desktop\Connector.Desktop\Assets\structura_connector.ico")
