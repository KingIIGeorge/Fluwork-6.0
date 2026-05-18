from pathlib import Path

from PIL import Image


BASE = Path(__file__).resolve().parent
BRAND_CANVAS = (220, 80)
BRAND_TARGET_HEIGHT = 56
BRAND_MAX_WIDTH = 200
WHATSAPP_CANVAS = (32, 32)


BRANDS = {
    "hp.logo.gif": {"target": "hp_print.png", "height": 76, "top": 2},
    "epson.logo.gif": {"target": "epson_print.png", "height": 56, "top": 2},
    "Brother.jpg": {"target": "brother_print.png", "height": 56},
    "Pantum.png": {"target": "pantum_print.png", "height": 80, "top": 0},
}


def trim_image(image):
    rgba = image.convert("RGBA")
    pixels = rgba.load()
    width, height = rgba.size
    xs = []
    ys = []

    for y in range(height):
        for x in range(width):
            r, g, b, a = pixels[x, y]
            if a > 10 and not (r > 245 and g > 245 and b > 245):
                xs.append(x)
                ys.append(y)

    if not xs:
        return rgba

    return rgba.crop((min(xs), min(ys), max(xs) + 1, max(ys) + 1))


def fit_brand(image, target_height):
    width, height = image.size
    scale = target_height / float(height)
    new_width = int(round(width * scale))
    new_height = target_height

    if new_width > BRAND_MAX_WIDTH:
        scale = BRAND_MAX_WIDTH / float(width)
        new_width = BRAND_MAX_WIDTH
        new_height = int(round(height * scale))

    return image.resize((new_width, new_height), Image.LANCZOS)


def fit_inside(image, max_size):
    image = image.copy()
    image.thumbnail(max_size, Image.LANCZOS)
    return image


def center_on_canvas(image, canvas_size, transparent=True, top=None):
    background = (255, 255, 255, 0) if transparent else (255, 255, 255, 255)
    canvas = Image.new("RGBA", canvas_size, background)
    x = (canvas_size[0] - image.size[0]) // 2
    y = top if top is not None else (canvas_size[1] - image.size[1]) // 2
    canvas.alpha_composite(image, (x, y))
    return canvas


def normalize_brand(source_name, spec):
    source = Image.open(BASE / source_name)
    trimmed = trim_image(source)
    fitted = fit_brand(trimmed, spec.get("height", BRAND_TARGET_HEIGHT))
    output = center_on_canvas(fitted, BRAND_CANVAS, transparent=True, top=spec.get("top"))
    output.save(BASE / spec["target"])


def normalize_whatsapp():
    source = Image.open(BASE / "WhatsApp.jpg")
    trimmed = trim_image(source)
    fitted = fit_inside(trimmed, WHATSAPP_CANVAS)
    output = center_on_canvas(fitted, WHATSAPP_CANVAS, transparent=True)
    output.save(BASE / "whatsapp_print.png")


def main():
    for source_name, spec in BRANDS.items():
        normalize_brand(source_name, spec)
    normalize_whatsapp()


if __name__ == "__main__":
    main()


