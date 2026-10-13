import os
from PIL import Image, ImageDraw, ImageFont

def create_sample_assets():
    assets_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(assets_dir, exist_ok=True)

    # 1. Sample Logo (200x100)
    logo_img = Image.new("RGBA", (200, 100), (41, 128, 185, 255))
    draw = ImageDraw.Draw(logo_img)
    draw.rectangle([10, 10, 190, 90], outline=(255, 255, 255, 255), width=3)
    draw.text((30, 40), "LOGO BU", fill=(255, 255, 255, 255))
    logo_img.save(os.path.join(assets_dir, "sample_logo.png"))

    # 2. Sample Stamp (150x150)
    stamp_img = Image.new("RGBA", (150, 150), (0, 0, 0, 0))
    draw = ImageDraw.Draw(stamp_img)
    draw.ellipse([10, 10, 140, 140], outline=(192, 57, 43, 220), width=4)
    draw.ellipse([20, 20, 130, 130], outline=(192, 57, 43, 220), width=2)
    draw.text((40, 65), "STEMPEL BU", fill=(192, 57, 43, 220))
    stamp_img.save(os.path.join(assets_dir, "sample_stamp.png"))

    # 3. Sample Signature (200x100)
    sig_img = Image.new("RGBA", (200, 100), (0, 0, 0, 0))
    draw = ImageDraw.Draw(sig_img)
    draw.line([(20, 70), (50, 20), (90, 80), (140, 30), (180, 70)], fill=(44, 62, 80, 255), width=4)
    draw.line([(20, 80), (180, 80)], fill=(44, 62, 80, 255), width=2)
    sig_img.save(os.path.join(assets_dir, "sample_signature.png"))

    print("Sample placeholder assets created successfully in", assets_dir)

if __name__ == "__main__":
    create_sample_assets()
