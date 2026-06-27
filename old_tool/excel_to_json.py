import pandas as pd
import os
import json
import glob
import re
import sys

# ---------- 配置 ----------
EXCEL_FILENAME = "统计信息.xlsx"
JSON_FILENAME = "cats.json"
CLASSIFIED_DIR = "classified"   # 图片总目录

# 列名关键字映射
COLUMN_KEYWORDS = {
    "编号": "id_col",
    "姓名": "name_col",
    "性别": "gender_col",
    "亲人指数": "affection_col",
    "状态": "status_col",
    "概要": "desc_col",
    "故事": "story_col",
    "图名": "pic_col"
}


def find_column_map(df_columns, keywords_map):
    col_map = {}
    missing = []
    for keyword, internal_key in keywords_map.items():
        matched = [col for col in df_columns if keyword in col]
        if len(matched) == 0:
            missing.append(keyword)
        elif len(matched) > 1:
            raise ValueError(f"列名关键字 '{keyword}' 匹配到多个列: {matched}，请修改 Excel 表头使其唯一。")
        else:
            col_map[internal_key] = matched[0]
    if missing:
        raise KeyError(
            f"在 Excel 中找不到包含以下关键字的列: {missing}\n"
            f"当前列名: {df_columns.tolist()}"
        )
    return col_map


def check_thumbnails(df, col_map, classified_dir):
    missing = []
    for _, row in df.iterrows():
        raw_id = row[col_map['id_col']]
        cat_id = f"{int(raw_id):02d}" if isinstance(raw_id, (int, float)) else str(raw_id).strip().zfill(2)

        name = str(row[col_map['name_col']]).strip() if pd.notna(row[col_map['name_col']]) else ""
        pic_name = str(row[col_map['pic_col']]).strip()

        folder_name = f"{cat_id} {name}"
        cat_folder = os.path.join(classified_dir, folder_name)  # 这里用系统路径，用于检查文件

        if not os.path.isdir(cat_folder):
            missing.append((cat_id, name, f"文件夹不存在: {folder_name}"))
            continue

        pattern = os.path.join(cat_folder, f"{pic_name}_[0-9][0-9].jpg")
        original_files = glob.glob(pattern)

        if not original_files:
            missing.append((cat_id, name, f"文件夹内没有任何以 {pic_name}_数字.jpg 命名的原图"))
            continue

        seqs = set()
        for fpath in original_files:
            filename = os.path.basename(fpath)
            m = re.match(rf"^{re.escape(pic_name)}_(\d{{2}})\.jpg$", filename)
            if m:
                seqs.add(int(m.group(1)))

        if 1 not in seqs:
            missing.append((cat_id, name, f"缺少头像原图 {pic_name}_01.jpg"))
        else:
            seqs.discard(1)

        for seq in sorted(seqs):
            thumb_filename = f"{pic_name}_{seq:02d}_thumb.jpg"
            thumb_path = os.path.join(cat_folder, thumb_filename)
            if not os.path.isfile(thumb_path):
                missing.append((cat_id, name, f"缺少缩略图: {thumb_filename}"))

        avatar_thumb = f"{pic_name}_01_thumb.jpg"
        avatar_thumb_path = os.path.join(cat_folder, avatar_thumb)
        if not os.path.isfile(avatar_thumb_path):
            missing.append((cat_id, name, f"缺少头像缩略图: {avatar_thumb}"))

    return missing


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_dir, EXCEL_FILENAME)
    json_path = os.path.join(base_dir, JSON_FILENAME)
    classified_dir = os.path.join(base_dir, CLASSIFIED_DIR)

    if not os.path.isdir(classified_dir):
        print(f"错误：未找到 {CLASSIFIED_DIR} 目录。")
        sys.exit(1)

    df = pd.read_excel(excel_path)
    df.columns = df.columns.str.strip()
    print(f"Excel 列名: {df.columns.tolist()}")

    col_map = find_column_map(df.columns, COLUMN_KEYWORDS)
    print(f"列名映射: {col_map}")

    print("正在逐文件夹检查缩略图...")
    missing_thumbs = check_thumbnails(df, col_map, classified_dir)

    if missing_thumbs:
        print("\n❌ 发现缺失的缩略图或文件夹问题，请修正后重试：\n")
        for cat_id, name, desc in missing_thumbs:
            print(f"  - 编号 {cat_id}「{name}」: {desc}")
        print("\n请补全以上内容后重新运行脚本。")
        sys.exit(0)

    print("✅ 所有缩略图检查通过，开始生成 cat.json...\n")

    cats = []
    for _, row in df.iterrows():
        raw_id = row[col_map['id_col']]
        cat_id = f"{int(raw_id):02d}" if isinstance(raw_id, (int, float)) else str(raw_id).strip().zfill(2)

        name = str(row[col_map['name_col']]).strip() if pd.notna(row[col_map['name_col']]) else ""
        gender = str(row[col_map['gender_col']]).strip() if pd.notna(row[col_map['gender_col']]) else "unknown"
        affection = int(row[col_map['affection_col']]) if pd.notna(row[col_map['affection_col']]) else 1
        status = str(row[col_map['status_col']]).strip() if pd.notna(row[col_map['status_col']]) else "normal"
        desc = str(row[col_map['desc_col']]).strip() if pd.notna(row[col_map['desc_col']]) else ""
        story = str(row[col_map['story_col']]).strip() if pd.notna(row[col_map['story_col']]) else ""
        pic_name = str(row[col_map['pic_col']]).strip()

        folder_name = f"{cat_id} {name}"
        # ↓↓↓ 关键修改：用正斜杠拼接 JSON 中使用的相对路径
        cat_folder_json = f"{CLASSIFIED_DIR}/{folder_name}"
        # 文件系统路径仍然用 os.path.join
        cat_folder_real = os.path.join(base_dir, CLASSIFIED_DIR, folder_name)

        avatar = f"{cat_folder_json}/{pic_name}_01_thumb.jpg"
        avatar_hd = f"{cat_folder_json}/{pic_name}_01.jpg"

        other_photos = []
        pattern = os.path.join(cat_folder_real, f"{pic_name}_[0-9][0-9].jpg")
        for fpath in sorted(glob.glob(pattern)):
            filename = os.path.basename(fpath)
            m = re.match(rf"^{re.escape(pic_name)}_(\d{{2}})\.jpg$", filename)
            if m:
                seq = int(m.group(1))
                if seq >= 2:
                    seq_str = f"{seq:02d}"
                    thumb = f"{cat_folder_json}/{pic_name}_{seq_str}_thumb.jpg"
                    hd = f"{cat_folder_json}/{pic_name}_{seq_str}.jpg"
                    other_photos.append({"thumb": thumb, "hd": hd})

        cats.append({
            "id": cat_id,
            "name": name,
            "gender": gender,
            "avatar": avatar,
            "avatar_hd": avatar_hd,
            "affection": affection,
            "status": status,
            "desc": desc,
            "story": story,
            "otherPhotos": other_photos
        })

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(cats, f, ensure_ascii=False, indent=2)

    print(f"已成功生成 {len(cats)} 条数据到 {json_path}")


if __name__ == "__main__":
    main()