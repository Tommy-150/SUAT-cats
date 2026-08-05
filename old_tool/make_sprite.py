# -*- coding: utf-8 -*-
"""
生成猫咪缩略图雪碧图

## 使用说明

### 功能
将 classified/ 下所有猫咪的主缩略图合并为一张大图（sprite sheet），
减少前端页面 HTTP 请求数（35 张图 → 1 张）。

### 运行
```bash
python make_sprite.py
```

### 需要的基础文件
```
SUAT-cats/
├── data/
│   └── cats.json         ← 猫咪数据
├── classified/           ← 猫咪图片（读取 _01_thumb.jpg）
├── static/               ← 输出目录（自动创建）
│   ├── sprite_thumb.jpg  ← 输出雪碧图
│   └── sprite_map.js     ← 输出映射文件
└── make_sprite.py        ← 就是这个文件
```

### 依赖
```bash
pip install Pillow
```
"""
import json
from pathlib import Path
from PIL import Image

BASE_DIR = Path(__file__).resolve().parent.parent
CATS_JSON = BASE_DIR / "data" / "cats.json"
CLASSIFIED = BASE_DIR / "classified"
STATIC = BASE_DIR / "static"
OUTPUT_IMG = STATIC / "sprite_thumb.jpg"
OUTPUT_JS = STATIC / "sprite_map.js"

THUMB_SIZE = 300
COLS = 7

def main():
    STATIC.mkdir(parents=True, exist_ok=True)
    cats = json.load(open(CATS_JSON, "r", encoding="utf-8"))
    n = len(cats)
    rows = (n + COLS - 1) // COLS
    canvas_w = COLS * THUMB_SIZE
    canvas_h = rows * THUMB_SIZE
    canvas = Image.new("RGB", (canvas_w, canvas_h), "#F7F5F0")

    positions = {}
    for i, cat in enumerate(cats):
        col = i % COLS
        row = i // COLS
        x, y = col * THUMB_SIZE, row * THUMB_SIZE
        positions[cat["id"]] = {"x": x, "y": y}
        # 从文件系统推导主图，不依赖 JSON 静态 avatar 字段
        pic = cat.get("pic_name", "")
        folder = CLASSIFIED / f"{cat['id']} {cat.get('name','')}"
        local_path = None
        if pic:
            cand = folder / f"{pic}_01_thumb.jpg"
            if cand.is_file():
                local_path = cand
        if local_path and local_path.is_file():
            try:
                img = Image.open(local_path).convert("RGB")
                img = img.resize((THUMB_SIZE, THUMB_SIZE), Image.LANCZOS)
                canvas.paste(img, (x, y))
            except Exception:
                pass
        print(f"[{cat['id']}] {cat['name']} → ({x}, {y})")

    canvas.save(OUTPUT_IMG, "JPEG", quality=85, optimize=True)
    print(f"\n雪碧图: {OUTPUT_IMG} ({canvas_w}×{canvas_h})")

    # 生成 JS 映射
    js = f"// 雪碧图位置映射  COL={COLS}  ROWS={rows}  SIZE={THUMB_SIZE}\n"
    js += f"const SPRITE_COLS = {COLS};\n"
    js += f"const SPRITE_ROWS = {rows};\n"
    js += "const SPRITE_MAP = {\n"
    for cat_id, pos in positions.items():
        js += f'  "{cat_id}": {{ x: {pos["x"]}, y: {pos["y"]} }},\n'
    js += "};\n"
    OUTPUT_JS.write_text(js, encoding="utf-8")
    print(f"映射文件: {OUTPUT_JS}")

if __name__ == "__main__":
    main()
