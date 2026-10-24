import argparse
from PIL import Image, ImageDraw, ImageFont


def generate_font(font_path, char_list, output_path, font_size=16, shadow=False):
    try:
        font = ImageFont.truetype(font_path, font_size, layout_engine=ImageFont.Layout.BASIC)
    except IOError:
        print(f"Warning: could not load font '{font_path}'. Using default.")
        font = ImageFont.load_default()

    ascent, descent = font.getmetrics()
    font_height_needed = ascent + descent

    dummy_img = Image.new("RGBA", (1, 1))
    draw = ImageDraw.Draw(dummy_img)

    char_widths = []

    for char in char_list:
        bbox = draw.textbbox((0, 0), char, font=font, anchor="lt")
        w = bbox[2] - bbox[0]
        char_widths.append(w)

    shadow_vertical_padding = 4 if shadow else 0
    font_height_final = font_height_needed + shadow_vertical_padding

    total_width = sum(char_widths) + len(char_list) + 1

    img = Image.new("RGBA", (total_width, font_height_final), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    current_x = 0
    draw.line([(current_x, 0), (current_x, font_height_final)], fill=(255, 0, 0, 255), width=1)
    draw.line([(current_x, font_size), (current_x, font_size)], fill=(255, 0, 255, 255), width=1)
    current_x += 1

    for i, char in enumerate(char_list):
        w = char_widths[i]

        bbox = draw.textbbox((0, 0), char, font=font, anchor="lt")
        x_offset = bbox[0]

        base_y = ascent

        if shadow:
            draw.text((current_x + 1 - x_offset, base_y + 1), char, font=font, fill=(0, 0, 0, 200), anchor="ls")
            draw.text((current_x - x_offset, base_y), char, font=font, fill=(255, 255, 255, 255), anchor="ls")
        else:
            draw.text((current_x - x_offset, base_y), char, font=font, fill=(255, 255, 255, 255), anchor="ls")

        current_x += w

        draw.line([(current_x, 0), (current_x, font_height_final)], fill=(255, 0, 0, 255), width=1)
        current_x += 1

    img.save(output_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Transparent image font generator with unified baseline.")

    parser.add_argument("--font", required=True, help="Path to the .ttf or .otf font file.")
    parser.add_argument("--out", default="font_output.png", help="Output PNG path.")
    parser.add_argument("--size", type=int, default=16, help="Font size in pixels.")
    parser.add_argument("--shadow", action="store_true", help="Add shadow to characters.")

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--chars", help="Continuous string of characters.")
    group.add_argument("--chars-file", help="Path to a .txt file containing the character string.")

    args = parser.parse_args()

    if args.chars_file:
        with open(args.chars_file, "r", encoding="utf-8") as f:
            char_list = f.read().rstrip("\r\n")
    else:
        char_list = args.chars

    generate_font(
        font_path=args.font,
        char_list=char_list,
        output_path=args.out,
        font_size=args.size,
        shadow=args.shadow
    )
