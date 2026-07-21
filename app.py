import os
import sys
import json
import shutil
import re
import glob
import atexit
import threading
from pathlib import Path
from http.server import HTTPServer, SimpleHTTPRequestHandler

import webview
from PIL import Image

# ================= 路径处理 =================
def get_base_dir():
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
        # onedir 模式：exe 在 SUATCatManager/ 子文件夹，数据在上级
        if (exe_dir / "_internal").is_dir() or exe_dir.name == "SUATCatManager":
            return exe_dir.parent
        return exe_dir
    return Path(__file__).resolve().parent

def get_resource_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

BASE_DIR = get_base_dir()
DATA_DIR = BASE_DIR / "data"
TAG_COLOR_PATH = DATA_DIR / "tag_color.json"
BAN_WORDS_PATH = DATA_DIR / "ban_words.json"
CATS_JSON_PATH = DATA_DIR / "cats.json"
CLASSIFIED_DIR = BASE_DIR / "classified"

# ================= 图片处理 =================
def auto_thumbnail(src_path, dst_path, size=300):
    img = Image.open(src_path).convert("RGB")
    w, h = img.size
    crop = max(100, int(min(w, h) * 0.6))
    left = (w - crop) // 2
    top = (h - crop) // 2
    img = img.crop((left, top, left + crop, top + crop))
    img = img.resize((size, size), Image.LANCZOS)
    img.save(dst_path, "JPEG", quality=90, optimize=True)

def next_seq_for(folder, pic_name):
    if not folder.is_dir():
        return 1
    used = set()
    for f in folder.iterdir():
        m = re.match(rf"^{re.escape(pic_name)}_(\d{{2}})(?:_thumb)?\.jpe?g$", f.name, re.IGNORECASE)
        if m:
            used.add(int(m.group(1)))
    n = 1
    while n in used:
        n += 1
    return n

# ================= API 类 =================
class ManagerAPI:
    def __init__(self):
        self._tag_colors = self._load_json(TAG_COLOR_PATH)
        self._ban_words = self._load_json(BAN_WORDS_PATH).get("words", [])
        self._ensure_dir()
        self._cats = self._load_cats()
        atexit.register(self._cleanup_all)

    def _ensure_dir(self):
        CLASSIFIED_DIR.mkdir(parents=True, exist_ok=True)
        DATA_DIR.mkdir(parents=True, exist_ok=True)

    def _load_json(self, path):
        if not path.exists():
            return {}
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _save_json(self, path, data):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _load_cats(self):
        if not CATS_JSON_PATH.exists():
            return []
        with open(CATS_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)

    def _save_cats(self):
        with open(CATS_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(self._cats, f, ensure_ascii=False, indent=2)

    def _cleanup_temp(self):
        """清理临时文件目录（保留目录本身）"""
        temp_dir = BASE_DIR / ".tmp_thumbs"
        if temp_dir.is_dir():
            for f in temp_dir.iterdir():
                if f.is_file():
                    try:
                        f.unlink()
                    except Exception:
                        pass

    def _update_sprite(self):
        """照片变更后自动更新雪碧图"""
        try:
            from PIL import Image as PILImage
            sd = BASE_DIR / "static"
            sd.mkdir(parents=True, exist_ok=True)
            cats = self._cats
            if not cats:
                return
            TS, COLS = 300, 7
            rows = (len(cats) + COLS - 1) // COLS
            canvas = PILImage.new("RGB", (COLS * TS, rows * TS), "#F7F5F0")
            positions = {}
            for i, cat in enumerate(cats):
                col = i % COLS
                r = i // COLS
                x, y = col * TS, r * TS
                positions[cat["id"]] = {"x": x, "y": y}
                av = cat.get("avatar", "")
                if av:
                    lp = BASE_DIR / av
                    if lp.is_file():
                        im = PILImage.open(lp).convert("RGB")
                        im = im.resize((TS, TS), PILImage.LANCZOS)
                        canvas.paste(im, (x, y))
            canvas.save(sd / "sprite_thumb.jpg", "JPEG", quality=85, optimize=True)
            js = "const SPRITE_MAP = {\n"
            for cid, pos in positions.items():
                js += f'  "{cid}": {{ x: {pos["x"]}, y: {pos["y"]} }},\n'
            js += "};\n"
            (sd / "sprite_map.js").write_text(js, encoding="utf-8")
        except Exception:
            pass

    def _cleanup_all(self):
        """完全清理：删除整个临时文件夹和 __pycache__"""
        for d in [BASE_DIR / ".tmp_thumbs", BASE_DIR / "__pycache__"]:
            if d.is_dir():
                try:
                    shutil.rmtree(d, ignore_errors=True)
                except Exception:
                    pass

    def _find_cat(self, cat_id):
        for c in self._cats:
            if c["id"] == cat_id:
                return c
        return None

    def _next_id(self):
        max_id = 0
        for c in self._cats:
            try:
                max_id = max(max_id, int(c["id"]))
            except:
                pass
        return f"{max_id + 1:02d}"

    def _build_cat_data(self, cat):
        """从 cats.json 条目 + 文件系统构建完整的猫数据"""
        cat_id = cat["id"]
        name = cat["name"]
        pic = cat.get("pic_name", "")
        folder = CLASSIFIED_DIR / f"{cat_id} {name}"
        folder_json = f"classified/{cat_id} {name}"

        avatar = ""
        avatar_hd = ""
        if pic and (folder / f"{pic}_01_thumb.jpg").is_file():
            avatar = f"{folder_json}/{pic}_01_thumb.jpg"
        if pic and (folder / f"{pic}_01.jpg").is_file():
            avatar_hd = f"{folder_json}/{pic}_01.jpg"

        other_photos = []
        if pic and folder.is_dir():
            pattern = str(folder / f"{pic}_[0-9][0-9].jpg")
            for fpath in sorted(glob.glob(pattern)):
                m = re.match(rf"^{re.escape(pic)}_(\d{{2}})\.jpg$", os.path.basename(fpath))
                if not m:
                    continue
                seq = int(m.group(1))
                if seq < 2:
                    continue
                seq_str = f"{seq:02d}"
                other_photos.append({
                    "thumb": f"{folder_json}/{pic}_{seq_str}_thumb.jpg",
                    "hd": f"{folder_json}/{pic}_{seq_str}.jpg",
                    "seq": seq,
                })

        return {
            "id": cat_id,
            "name": name,
            "gender": cat.get("gender", "unknown"),
            "avatar": avatar,
            "avatar_hd": avatar_hd,
            "affection": cat.get("affection", 1),
            "status": cat.get("status", "normal"),
            "desc": cat.get("desc", ""),
            "story": cat.get("story", ""),
            "tags": cat.get("tags", []),
            "otherPhotos": other_photos,
            "pic_name": pic,
        }

    # ---------- 公开 API ----------
    def get_cats(self):
        return [self._build_cat_data(c) for c in self._cats]

    def save_cat(self, form):
        cat_id = str(form.get("id", "")).strip().zfill(2)
        name = form.get("name", "").strip()
        pic_name = form.get("pic_name", "").strip()
        if not name or not pic_name:
            return {"success": False, "message": "姓名和图名不能为空"}
        if not re.match(r"^[A-Za-z0-9_]+$", pic_name):
            return {"success": False, "message": "图名只能含英文/数字/下划线"}

        existing = self._find_cat(cat_id)
        if existing:
            old_name = existing.get("name", "")
            old_pic = existing.get("pic_name", "")
            # 图名变化 → 重命名文件
            if old_pic and old_pic != pic_name:
                old_folder = CLASSIFIED_DIR / f"{cat_id} {old_name}"
                if old_folder.is_dir():
                    rename_ops = []
                    try:
                        for f in old_folder.iterdir():
                            m = re.match(rf"^{re.escape(old_pic)}_(\d{{2}})(?:_thumb)?\.jpg$", f.name, re.IGNORECASE)
                            if m:
                                seq = m.group(1)
                                is_thumb = "_thumb" in f.name
                                new_file = f"{pic_name}_{seq}{'_thumb' if is_thumb else ''}.jpg"
                                rename_ops.append((f, old_folder / new_file))
                        for old_path, new_path in rename_ops:
                            old_path.rename(new_path)
                    except Exception as e:
                        for old_path, new_path in reversed(rename_ops):
                            if new_path.exists():
                                try:
                                    new_path.rename(old_path)
                                except Exception:
                                    pass
                        return {"success": False, "message": f"图名重命名失败: {e}"}
            # 名字变化 → 重命名文件夹
            if old_name and old_name != name:
                old_folder = CLASSIFIED_DIR / f"{cat_id} {old_name}"
                new_folder = CLASSIFIED_DIR / f"{cat_id} {name}"
                if new_folder.exists():
                    return {"success": False, "message": f"目标文件夹已存在: {new_folder.name}"}
                if old_folder.is_dir():
                    old_folder.rename(new_folder)
            # 同步 JSON 中所有图片路径（avatar / avatar_hd / otherPhotos）
            if (old_name and old_name != name) or (old_pic and old_pic != pic_name):
                old_dir = f"{cat_id} {old_name}"
                new_dir = f"{cat_id} {name}"
                def _fix_path(p):
                    if not p:
                        return p
                    s = p
                    if old_name and old_name != name:
                        s = s.replace(f"classified/{old_dir}/", f"classified/{new_dir}/")
                    if old_pic and old_pic != pic_name:
                        s = re.sub(rf"{re.escape(old_pic)}_(\d{{2}})", rf"{pic_name}_\1", s)
                    return s
                existing["avatar"] = _fix_path(existing.get("avatar", ""))
                existing["avatar_hd"] = _fix_path(existing.get("avatar_hd", ""))
                for photo in existing.get("otherPhotos", []):
                    photo["thumb"] = _fix_path(photo.get("thumb", ""))
                    photo["hd"] = _fix_path(photo.get("hd", ""))
            # 更新已有条目
            existing.update({
                "name": name,
                "gender": form.get("gender", "unknown"),
                "affection": int(form.get("affection", 1)),
                "status": form.get("status", "normal"),
                "desc": form.get("desc", ""),
                "story": form.get("story", ""),
                "pic_name": pic_name,
                "tags": form.get("tags", []),
            })
        else:
            if not cat_id or cat_id == '00':
                cat_id = self._next_id()
            new_cat = {
                "id": cat_id,
                "name": name,
                "gender": form.get("gender", "unknown"),
                "affection": int(form.get("affection", 1)),
                "status": form.get("status", "normal"),
                "desc": form.get("desc", ""),
                "story": form.get("story", ""),
                "pic_name": pic_name,
                "tags": form.get("tags", []),
            }
            self._cats.append(new_cat)
            folder = CLASSIFIED_DIR / f"{cat_id} {name}"
            folder.mkdir(parents=True, exist_ok=True)
        self._save_cats()
        self._cleanup_temp()
        return {"success": True}

    def delete_cat(self, cat_id):
        cat_id = str(cat_id).strip().zfill(2)
        self._cats = [c for c in self._cats if c["id"] != cat_id]
        self._save_cats()
        for folder in CLASSIFIED_DIR.glob(f"{cat_id} *"):
            shutil.rmtree(folder, ignore_errors=True)
        self._cleanup_temp()
        return {"success": True}

    def upload_image(self, source_path, cat_id, name, pic_name, seq=None):
        cat_id = str(cat_id).strip().zfill(2)
        folder = CLASSIFIED_DIR / f"{cat_id} {name}"
        folder.mkdir(parents=True, exist_ok=True)
        if seq is None:
            seq = next_seq_for(folder, pic_name)
        dst = folder / f"{pic_name}_{seq:02d}.jpg"
        Image.open(source_path).convert("RGB").save(dst, "JPEG", quality=95)
        thumb_dst = folder / f"{pic_name}_{seq:02d}_thumb.jpg"
        auto_thumbnail(dst, thumb_dst)
        self._update_sprite()
        self._cleanup_temp()
        return {"success": True, "seq": seq, "thumb": f"classified/{cat_id} {name}/{pic_name}_{seq:02d}_thumb.jpg"}

    def open_file_dialog(self):
        result = webview.windows[0].create_file_dialog(webview.OPEN_DIALOG, allow_multiple=False, file_types=['Image Files (*.jpg;*.jpeg;*.png)'])
        return result[0] if result else None

    def get_tag_colors(self):
        return self._tag_colors

    def save_tag_colors(self, data):
        old_keys = set(self._tag_colors.keys())
        new_keys = set(data.keys())
        removed = old_keys - new_keys
        added = new_keys - old_keys
        if len(removed) == 1 and len(added) == 1:
            old_tag = removed.pop()
            new_tag = added.pop()
            for cat in self._cats:
                if old_tag in cat.get("tags", []):
                    cat["tags"] = [new_tag if t == old_tag else t for t in cat["tags"]]
            self._save_cats()
        self._tag_colors = data
        self._save_json(TAG_COLOR_PATH, data)
        return {"success": True}

    def get_ban_words(self):
        return {"words": self._ban_words}

    def save_ban_words(self, data):
        self._ban_words = data.get("words", [])
        self._save_json(BAN_WORDS_PATH, {"words": self._ban_words})
        return {"success": True}

    def get_readme(self):
        readme_path = BASE_DIR / "README.md"
        if readme_path.exists():
            return readme_path.read_text(encoding="utf-8")
        return ""

    def delete_photo(self, cat_id, pic_name, seq):
        cat_id = str(cat_id).strip().zfill(2)
        for folder in CLASSIFIED_DIR.glob(f"{cat_id} *"):
            base_name = f"{pic_name}_{int(seq):02d}"
            for ext in ['.jpg', '_thumb.jpg']:
                file_path = folder / f"{base_name}{ext}"
                if file_path.exists():
                    file_path.unlink()
            self._update_sprite()
            self._cleanup_temp()
            return {"success": True}
        return {"success": False, "message": "未找到照片"}

    def crop_thumbnail(self, cat_id, pic_name, seq, crop_box):
        cat_id = str(cat_id).strip().zfill(2)
        for folder in CLASSIFIED_DIR.glob(f"{cat_id} *"):
            original_path = folder / f"{pic_name}_{int(seq):02d}.jpg"
            thumb_path = folder / f"{pic_name}_{int(seq):02d}_thumb.jpg"
            if not original_path.is_file():
                return {"success": False, "message": "原图不存在"}
            try:
                left, top, right, bottom = map(int, crop_box)
                img = Image.open(original_path).convert("RGB")
                left = max(0, left); top = max(0, top)
                right = min(img.width, right); bottom = min(img.height, bottom)
                cropped = img.crop((left, top, right, bottom))
                cropped = cropped.resize((300, 300), Image.LANCZOS)
                cropped.save(thumb_path, "JPEG", quality=90, optimize=True)
                self._update_sprite()
                self._cleanup_temp()
                return {"success": True}
            except Exception as e:
                return {"success": False, "message": str(e)}
        return {"success": False, "message": "未找到文件夹"}

    def apply_sort_order(self, ordered_ids):
        if len(ordered_ids) != len(self._cats):
            return {"success": False, "message": "排序列表与数据不匹配"}
        id_to_cat = {c["id"]: c for c in self._cats}
        ordered = []
        rename_ops = []
        for idx, cid in enumerate(ordered_ids, start=1):
            cat = id_to_cat.get(cid)
            if not cat:
                return {"success": False, "message": f"找不到猫咪: {cid}"}
            old_id = cat["id"]
            new_id = f"{idx:02d}"
            cat["id"] = new_id
            ordered.append(cat)
            if old_id != new_id:
                old_folder = CLASSIFIED_DIR / f"{old_id} {cat['name']}"
                new_folder = CLASSIFIED_DIR / f"{new_id} {cat['name']}"
                if old_folder.is_dir() and not new_folder.exists():
                    rename_ops.append((old_folder, new_folder))
        for old_f, new_f in rename_ops:
            old_f.rename(new_f)
        self._cats = ordered
        self._save_cats()
        self._cleanup_temp()
        return {"success": True, "count": len(ordered)}

    def apply_tags(self, cat_id, tags):
        cat_id = str(cat_id).strip().zfill(2)
        cat = self._find_cat(cat_id)
        if cat:
            cat["tags"] = tags
            self._save_cats()
        return {"success": True}

    def save_readme(self, content):
        readme_path = BASE_DIR / "README.md"
        try:
            readme_path.write_text(content, encoding="utf-8")
            return {"success": True}
        except Exception as e:
            return {"success": False, "message": str(e)}

    def swap_to_main(self, cat_id, pic_name, seq):
        cat_id = str(cat_id).strip().zfill(2)
        target_seq = int(seq)
        if target_seq == 1:
            return {"success": False, "message": "已经是主图"}
        for folder in CLASSIFIED_DIR.glob(f"{cat_id} *"):
            try:
                def swap_pair(a, b):
                    for suffix in [".jpg", "_thumb.jpg"]:
                        a_file = folder / f"{pic_name}_{a:02d}{suffix}"
                        b_file = folder / f"{pic_name}_{b:02d}{suffix}"
                        tmp_file = folder / f"{pic_name}_{a:02d}{suffix}.tmpswap"
                        if a_file.is_file():
                            a_file.rename(tmp_file)
                        if b_file.is_file():
                            b_file.rename(a_file)
                        if tmp_file.is_file():
                            tmp_file.rename(b_file)
                swap_pair(1, target_seq)
                self._update_sprite()
                self._cleanup_temp()
                return {"success": True}
            except Exception as e:
                return {"success": False, "message": str(e)}
        return {"success": False, "message": "未找到文件夹"}

    def open_folder(self, cat_id, name):
        cat_id = str(cat_id).strip().zfill(2)
        import subprocess
        folder = CLASSIFIED_DIR / f"{cat_id} {name}"
        if folder.is_dir():
            subprocess.Popen(['explorer', str(folder)])
            return {"success": True}
        return {"success": False, "message": "文件夹不存在"}

    def move_row(self, cat_id, direction, step=1):
        cat_id = str(cat_id).strip().zfill(2)
        step = int(step)
        idx = next((i for i, c in enumerate(self._cats) if c["id"] == cat_id), None)
        if idx is None:
            return {"success": False, "message": "未找到该猫咪"}
        new_idx = max(0, idx - step) if direction == "up" else min(len(self._cats) - 1, idx + step)
        if new_idx == idx:
            return {"success": True, "message": "已在边界"}
        self._cats.insert(new_idx, self._cats.pop(idx))
        self._save_cats()
        self._cleanup_temp()
        return {"success": True}

    def open_file_dialog_multi(self):
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG, allow_multiple=True,
            file_types=['Image Files (*.jpg;*.jpeg;*.png)']
        )
        return list(result) if result else []

    def upload_images(self, source_paths, cat_id, name, pic_name):
        results = []
        for src in source_paths:
            r = self.upload_image(src, cat_id, name, pic_name)
            results.append(r)
        return {"success": True, "results": results}

    def rename_pic_name(self, cat_id, name, old_pic, new_pic):
        cat_id = str(cat_id).strip().zfill(2)
        if not re.match(r"^[A-Za-z0-9_]+$", new_pic):
            return {"success": False, "message": "图名只能含英文/数字/下划线"}
        for folder in CLASSIFIED_DIR.glob(f"{cat_id} *"):
            if not folder.is_dir():
                continue
            try:
                for f in sorted(folder.iterdir()):
                    m = re.match(rf"^{re.escape(old_pic)}_(\d{{2}})(_thumb)?\.(jpe?g)$", f.name, re.IGNORECASE)
                    if m:
                        seq = m.group(1)
                        suffix = m.group(2) or ""
                        ext = m.group(3)
                        new_name = f"{new_pic}_{seq}{suffix}.{ext}"
                        f.rename(folder / new_name)
                return {"success": True}
            except Exception as e:
                return {"success": False, "message": str(e)}
        return {"success": False, "message": "未找到文件夹"}

    def get_photo_health(self, cat_id, pic_name):
        cat_id = str(cat_id).strip().zfill(2)
        for folder in CLASSIFIED_DIR.glob(f"{cat_id} *"):
            issues = []
            for f in sorted(folder.iterdir()):
                m = re.match(rf"^{re.escape(pic_name)}_(\d{{2}})\.jpg$", f.name, re.IGNORECASE)
                if not m:
                    continue
                seq = int(m.group(1))
                thumb_file = folder / f"{pic_name}_{seq:02d}_thumb.jpg"
                if not thumb_file.is_file():
                    issues.append({"seq": seq, "issue": "missing_thumb", "file": f"{pic_name}_{seq:02d}_thumb.jpg"})
            return {"success": True, "issues": issues}
        return {"success": False, "message": "未找到文件夹"}

    def cleanup_temp_files(self):
        temp_dir = BASE_DIR / ".tmp_thumbs"
        deleted = []
        size_freed = 0
        if temp_dir.is_dir():
            for f in temp_dir.iterdir():
                if f.is_file():
                    try:
                        size_freed += f.stat().st_size
                        f.unlink()
                        deleted.append(f.name)
                    except Exception:
                        pass
        return {"success": True, "deleted": len(deleted), "size_freed": size_freed, "files": deleted[:50]}

    def scan_orphan_files(self):
        orphans = []
        if not CLASSIFIED_DIR.is_dir():
            return {"success": True, "orphans": []}
        valid_pics = set(c.get("pic_name", "") for c in self._cats if c.get("pic_name"))
        for folder in sorted(CLASSIFIED_DIR.iterdir()):
            if not folder.is_dir():
                continue
            for f in sorted(folder.iterdir()):
                if not f.is_file():
                    continue
                name = f.name
                matched = name.endswith('.tmpswap')
                if not matched:
                    for pic in valid_pics:
                        if re.match(rf"^{re.escape(pic)}_\d{{2}}(_thumb)?\.jpe?g$", name, re.IGNORECASE):
                            matched = True
                            break
                if not matched:
                    orphans.append(str(f.relative_to(CLASSIFIED_DIR)))
        return {"success": True, "orphans": orphans}

    def delete_orphan_files(self, paths):
        deleted = []
        for rel_path in paths:
            full_path = CLASSIFIED_DIR / rel_path
            if full_path.is_file():
                try:
                    full_path.unlink()
                    deleted.append(rel_path)
                except Exception:
                    pass
        return {"success": True, "deleted": deleted}

    def compress_all_images(self):
        """一键压缩 classified/ 下所有 >1000×1000 的大图，跳过缩略图"""
        THRESHOLD = 1000
        QUALITY = 80
        before_total = 0
        after_total = 0
        compressed = []
        errors = []

        if not CLASSIFIED_DIR.is_dir():
            return {"success": False, "message": "classified 文件夹不存在"}

        for folder in sorted(CLASSIFIED_DIR.iterdir()):
            if not folder.is_dir():
                continue
            for f in sorted(folder.iterdir()):
                if not f.is_file() or f.suffix.lower() not in (".jpg", ".jpeg"):
                    continue
                if "_thumb" in f.name or f.name.endswith(".tmpswap"):
                    continue
                try:
                    from PIL import Image as PILImage
                    img = PILImage.open(f)
                    w, h = img.size
                    old_size = f.stat().st_size
                    if w <= THRESHOLD and h <= THRESHOLD:
                        continue
                    before_total += old_size
                    tmp = f.with_suffix(".tmp_compress")
                    if img.mode in ("RGBA", "P"):
                        img = img.convert("RGB")
                    img.save(tmp, "JPEG", quality=QUALITY, optimize=True, progressive=True)
                    new_size = tmp.stat().st_size
                    if new_size < old_size:
                        tmp.replace(f)
                        after_total += new_size
                        compressed.append({
                            "file": str(f.relative_to(CLASSIFIED_DIR)),
                            "before": old_size,
                            "after": new_size,
                            "ratio": round((1 - new_size / old_size) * 100),
                        })
                    else:
                        tmp.unlink()
                        after_total += old_size
                except Exception as e:
                    errors.append(str(f.relative_to(CLASSIFIED_DIR)) + ": " + str(e))

        return {
            "success": True,
            "compressed": compressed,
            "errors": errors,
            "before_total": before_total,
            "after_total": after_total,
            "saved_pct": round((1 - after_total / before_total) * 100) if before_total else 0,
        }

# ================= HTTP 服务器 =================
def start_http_server():
    base_dir = get_base_dir()
    res_dir = get_resource_dir()
    (base_dir / 'classified').mkdir(parents=True, exist_ok=True)

    # 前端静态文件已被打包进 exe，从资源目录(_MEIPASS)读取，
    # 保证即使 exe 同目录缺少这些文件也能正常打开；
    # 其余内容(classified/、data/ 等可编辑数据)仍从磁盘 base_dir 读写。
    FRONTEND_FILES = {"app_manager.html", "marked.min.js"}

    class Handler(SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=str(base_dir), **kwargs)
        def translate_path(self, path):
            rel = path.split("?", 1)[0].split("#", 1)[0].lstrip("/")
            if rel in FRONTEND_FILES:
                cand = res_dir / rel
                if cand.is_file():
                    return str(cand)
            return super().translate_path(path)
        def log_message(self, format, *args):
            pass
    server = HTTPServer(('127.0.0.1', 0), Handler)
    port = server.server_address[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return f'http://127.0.0.1:{port}/app_manager.html', server

# ================= 主入口 =================
if __name__ == '__main__':
    import traceback, ctypes
    try:
        api = ManagerAPI()
        url, server = start_http_server()
        window = webview.create_window(
            title='🐱 SUAT 猫咪管理器',
            url=url,
            js_api=api,
            width=1200,
            height=800,
            min_size=(900, 600)
        )
        webview.start()
    except Exception as e:
        err = f"{e}\n{traceback.format_exc()}"
        ctypes.windll.user32.MessageBoxW(0, err, "SUAT启动失败", 0x10)
        (BASE_DIR / "_crash.log").write_text(err, encoding="utf-8")
        raise
