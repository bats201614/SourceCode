import os
import re
import shutil
import pandas as pd

def normalize_artwork_name(name):
    """
    将作品名称规范化为纯文字形式
    - 去掉书名号《》、引号""、''、括号等所有特殊符号
    - 保留中英文、数字
    """
    return re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '', name)

def ask_excel_path():
    """询问Excel路径"""
    path = input("请输入Excel文件路径: ").strip().strip('"')
    if not os.path.exists(path):
        print(f"文件不存在: {path}")
        return ask_excel_path()
    return path

def ask_folder_path(prompt):
    """询问文件夹路径"""
    path = input(f"{prompt}: ").strip().strip('"')
    if not os.path.exists(path):
        print(f"文件夹不存在: {path}")
        return ask_folder_path(prompt)
    return path

def load_mapping_from_excel(excel_path):
    """
    从Excel读取作品名称和ASIN的对应关系
    Excel列: 产品名称, ASIN
    """
    df = pd.read_excel(excel_path)
    # 找到正确的列
    # ASIN列：表头为'ASIN'
    # 名称列：用列索引3（产品名称）
    asin_col = df.columns[25]  # ASIN
    name_col = df.columns[3]  # 产品名称

    # 取有ASIN的行
    valid_rows = df[df[asin_col].notna()]
    mapping = {}
    for _, row in valid_rows.iterrows():
        name = str(row[name_col]).strip()
        name = normalize_artwork_name(name)
        asin = str(row[asin_col]).strip()
        if name not in mapping:
            mapping[name] = []
        mapping[name].append(asin)
    return mapping

# 分隔词列表（图片名中的通用描述，不是作品名的一部分）
SEPARATORS = ["无框帆布画", "款帆布画", "无框", "款", "横", "竖", "横板", "竖板"]

def is_separator(text, start_idx):
    """检查 text[start_idx:] 是否以分隔词开头"""
    substring = text[start_idx:]
    return any(substring.startswith(sep) for sep in SEPARATORS)

def match_images_to_artwork(image_folder, artwork_mapping):
    """
    将图片文件夹中的图片匹配到作品
    图片名称包含作品名称 -> 形成对应关系
    """
    all_images = os.listdir(image_folder)
    matched = []
    unmatched_images = []

    for img in all_images:
        # 规范化图片名称后再匹配
        img_normalized = normalize_artwork_name(img)
        matched_artwork = None
        for artwork_name in artwork_mapping.keys():
            if artwork_name in img_normalized:
                # 找到匹配后，检查是否构成完整词
                start_idx = img_normalized.find(artwork_name)
                end_idx = start_idx + len(artwork_name)
                # 前一个字符如果不是中文，是边界
                prev_char_ok = (start_idx == 0 or not('\u4e00' <= img_normalized[start_idx - 1] <= '\u9fff'))
                # 后一个字符如果不是中文，或者是分隔词，是边界
                next_char_ok = (end_idx >= len(img_normalized) or
                                not('\u4e00' <= img_normalized[end_idx] <= '\u9fff') or
                                is_separator(img_normalized, end_idx))
                if prev_char_ok and next_char_ok:
                    matched_artwork = artwork_name
                    break
        if matched_artwork:
            matched.append((img, artwork_mapping[matched_artwork]))
        else:
            unmatched_images.append(img)

    return matched, unmatched_images

def process():
    # 1. 询问路径
    excel_path = ask_excel_path()
    image_folder = ask_folder_path("请输入图片文件夹路径")
    output_folder = ask_folder_path("请输入输出文件夹路径")

    # 2. 创建输出文件夹
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 3. 从Excel加载映射
    artwork_mapping = load_mapping_from_excel(excel_path)
    print(f"\n从Excel加载了 {len(artwork_mapping)} 个作品")

    # 4. 匹配图片
    matched, unmatched = match_images_to_artwork(image_folder, artwork_mapping)
    print(f"匹配成功: {len(matched)} 张图片")
    if unmatched:
        print(f"未匹配的图片: {unmatched}")

    # 5. 生成文件
    all_generated = []
    for img_file, asins in matched:
        source_path = os.path.join(image_folder, img_file)
        ext = os.path.splitext(img_file)[1]
        print(f"处理: {img_file}，ASIN数量: {len(asins)}")

        for asin in asins:
            new_name = f"{asin}{ext}"
            shutil.copy(source_path, os.path.join(output_folder, new_name))
            print(f"  -> 已生成: {new_name}")
            all_generated.append(asin)

    print(f"\n总计生成: {len(all_generated)} 个文件")

if __name__ == "__main__":
    process()