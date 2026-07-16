#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
增量转换：扫描原始 JPG，缺少 _thumb 的才处理

## 使用说明

### 功能
扫描 `classified/` 下所有猫咪文件夹，对没有 `_thumb.jpg` 的原图自动生成 300x300 缩略图。
已存在的缩略图不会重复生成（增量模式）。

### 运行
```bash
python auto_thumb.py
```

### 需要的基础文件
```
SUAT-cats/
├── classified/          ← 猫咪图片（必须有，按 {id} {name}/ 组织）
│   ├── 01 丑橘/
│   │   ├── ugly_orange_01.jpg
│   │   └── ...
│   └── ...
├── auto_thumb.py        ← 就是这个文件
└── (无其他依赖)
```

### 依赖
```bash
pip install Pillow
```
"""

import os
from PIL import Image

INPUT_DIR = "classified"
SUFFIX = "_thumb"
OUTPUT_SIZE = 300
CROP_RATIO = 0.6
MIN_CROP = 100
JPEG_QUALITY = 90

def get_output_path(input_path, suffix=SUFFIX):
    dir_name = os.path.dirname(input_path)
    base = os.path.basename(input_path)
    name, ext = os.path.splitext(base)
    return os.path.join(dir_name, f"{name}{suffix}{ext}")

def is_original_jpg(filename, suffix=SUFFIX):
    name, ext = os.path.splitext(filename)
    return ext.lower() in ('.jpg', '.jpeg') and suffix not in name

def main():
    # 收集原始 JPG
    jpg_files = []
    for root, dirs, files in os.walk(INPUT_DIR):
        for f in files:
            if is_original_jpg(f):
                jpg_files.append(os.path.join(root, f))

    # 分类：已处理（有 _thumb） vs 未处理
    to_process = []
    skipped = 0
    for fp in jpg_files:
        out = get_output_path(fp)
        if os.path.exists(out):
            skipped += 1
        else:
            to_process.append(fp)

    print(f"📸 原始 JPG 总数: {len(jpg_files)}")
    print(f"✅ 已有 _thumb (跳过): {skipped}")
    print(f"🆕 需要处理: {len(to_process)}")

    if not to_process:
        print("✨ 所有图片均已处理，无需转换。")
        return

    success = 0
    failed = []
    for idx, fp in enumerate(to_process, 1):
        out = get_output_path(fp)
        print(f"[{idx}/{len(to_process)}] {fp}")
        try:
            img = Image.open(fp).convert("RGB")
            w, h = img.size
            crop = max(MIN_CROP, int(min(w, h) * CROP_RATIO))
            left = (w - crop) // 2
            top = (h - crop) // 2
            img = img.crop((left, top, left + crop, top + crop))
            img = img.resize((OUTPUT_SIZE, OUTPUT_SIZE), Image.LANCZOS)
            img.save(out, "JPEG", quality=JPEG_QUALITY, optimize=True)
            print(f"   ✅ {os.path.basename(out)}")
            success += 1
        except Exception as e:
            print(f"   ❌ {e}")
            failed.append((fp, str(e)))

    print(f"\n✅ 成功: {success}, ❌ 失败: {len(failed)}")
    if failed:
        for fp, err in failed:
            print(f"  • {fp}\n    {err}")

if __name__ == "__main__":
    main()