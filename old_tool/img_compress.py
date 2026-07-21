"""图片压缩 - 遍历文件夹递归压缩 >1000×1000 的大图，保持分辨率"""

from pathlib import Path
from PIL import Image

BASE = Path(__file__).resolve().parent
THRESHOLD_W, THRESHOLD_H = 1000, 1000
JPEG_QUALITY = 80


def compress_image(filepath):
    old_size = filepath.stat().st_size
    ext = filepath.suffix.lower()
    img = Image.open(filepath)

    w, h = img.size
    if w <= THRESHOLD_W and h <= THRESHOLD_H:
        return None, None

    tmp = filepath.with_suffix(".tmp_compress")

    if ext in (".jpg", ".jpeg"):
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        img.save(tmp, "JPEG", quality=JPEG_QUALITY, optimize=True, progressive=True)
    elif ext == ".png":
        img.save(tmp, "PNG", optimize=True)
    else:
        return None, None

    new_size = tmp.stat().st_size
    if new_size < old_size:
        tmp.replace(filepath)
        return old_size, new_size
    else:
        tmp.unlink()
        return old_size, old_size


def main():
    files = [f for f in BASE.rglob("*") if f.is_file()
             and f.suffix.lower() in (".jpg", ".jpeg", ".png")]
    if not files:
        print("没有找到图片文件")
        return

    print(f"扫描到 {len(files)} 张图片\n")
    total_before = total_after = compressed = 0

    for f in sorted(files):
        try:
            old, new = compress_image(f)
        except Exception as e:
            print(f"  ✗ {f.name}: {e}")
            continue

        if old is None:
            img = Image.open(f)
            print(f"  - {f.name}  {img.size[0]}×{img.size[1]}  "
                  f"≤{THRESHOLD_W}×{THRESHOLD_H}  跳过")
            total_before += f.stat().st_size
            total_after += f.stat().st_size
            continue

        ratio = (1 - new / old) * 100
        total_before += old
        total_after += new
        compressed += 1

        if new < old:
            print(f"  ✓ {f.name}: {old//1024}KB → {new//1024}KB  (-{ratio:.0f}%)")
        else:
            print(f"  = {f.name}: {old//1024}KB  已最优  未改动")

    print(f"\n完成: {compressed}/{len(files)} 张已处理")
    if total_before:
        print(f"{total_before/1024/1024:.1f}MB → {total_after/1024/1024:.1f}MB  "
              f"(-{(1-total_after/total_before)*100:.0f}%)")


if __name__ == "__main__":
    main()
