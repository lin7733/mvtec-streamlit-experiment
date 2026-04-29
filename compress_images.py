from pathlib import Path

from PIL import Image, ImageOps

SOURCE_DIR = Path("00_raw")

OUTPUT_DIR = Path("00_raw_compressed")

MAX_WIDTH = 600

JPEG_QUALITY = 75

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}

for src_path in SOURCE_DIR.rglob("*"):

    if src_path.suffix.lower() not in IMAGE_EXTS:

        continue

    rel_path = src_path.relative_to(SOURCE_DIR)

    # 输出路径保持原来的文件夹结构，但统一转成 jpg

    out_path = OUTPUT_DIR / rel_path

    out_path = out_path.with_suffix(".jpg")

    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:

        img = Image.open(src_path)

        img = ImageOps.exif_transpose(img).convert("RGB")

        w, h = img.size

        if w > MAX_WIDTH:

            new_h = int(h * MAX_WIDTH / w)

            img = img.resize((MAX_WIDTH, new_h), Image.LANCZOS)

        img.save(out_path, "JPEG", quality=JPEG_QUALITY, optimize=True)

        print(f"OK: {src_path} -> {out_path}")

    except Exception as e:

        print(f"FAILED: {src_path}, reason: {e}")