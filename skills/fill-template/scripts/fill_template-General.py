"""
亚马逊批量上传模板自动化脚本 - 配置驱动版本
从SKU编码申请表和Listing文案表读取数据，填充到Amazon模板
支持多类目配置（当前: Wall Art - Paintings）
"""

import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import Dict, List, Any
from pathlib import Path
import logging
import re
import yaml # type: ignore

# ==================== 日志配置 ====================

def setup_logging() -> logging.Logger:
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    return logging.getLogger(__name__)


# ==================== 类目选择与配置加载 ====================

def select_category() -> str:
    """让用户选择模板类目"""
    categories = {
        '1': 'wall_art_paintings',
    }

    print("\n" + "=" * 50)
    print("请选择模板类目：")
    print("  1. Wall Art - Paintings")
    print("=" * 50)

    choice = input("请输入序号 (1): ").strip()

    if choice not in categories:
        raise ValueError(f"无效的选择: {choice}")

    return categories[choice]


def load_category_config(category_name: str) -> dict:
    """加载类目配置（base.yaml + 类目.yaml）"""
    skill_dir = Path(__file__).parent

    base_path = skill_dir / 'config' / 'base.yaml'
    category_path = skill_dir / 'config' / f'{category_name}.yaml'

    if not base_path.exists():
        raise FileNotFoundError(f"基础配置文件不存在: {base_path}")
    if not category_path.exists():
        raise FileNotFoundError(f"类目配置文件不存在: {category_path}")

    with open(base_path, 'r', encoding='utf-8') as f:
        config = yaml.safe_load(f)

    with open(category_path, 'r', encoding='utf-8') as f:
        category_config = yaml.safe_load(f)

    # 合并配置
    config.update(category_config)
    return config


def normalize_path(path: str) -> str:
    """
    验证并转换路径为标准格式
    1. 移除首尾双引号（Windows复制自带）
    2. 移除首尾空格
    3. 展开 ~ 为用户主目录
    4. 统一斜杠为正斜杠 /
    5. 移除多余斜杠
    """
    if not path:
        raise ValueError("路径不能为空")

    # 移除首尾双引号（Windows复制文件地址时自带）
    path = path.strip().strip('"').strip("'")

    # 移除首尾空格
    path = path.strip()

    # 展开 ~ 为用户主目录
    if path.startswith('~'):
        path = str(Path.home()) + path[1:]

    # 统一斜杠（Windows 反斜杠转正斜杠）
    path = path.replace('\\', '/')

    # 移除多余的斜杠（处理 // 等情况）
    while '//' in path:
        path = path.replace('//', '/')

    return path


def validate_file_path(path: str, file_type: str) -> str:
    """
    验证文件路径是否存在且为有效文件
    如果路径是相对路径，尝试转换为绝对路径
    """
    normalized = normalize_path(path)

    # 检查文件是否存在
    file_path = Path(normalized)
    if not file_path.exists():
        raise FileNotFoundError(f"{file_type}文件不存在: {normalized}")

    if not file_path.is_file():
        raise ValueError(f"{file_type}路径不是有效文件: {normalized}")

    return normalized


def ask_user_inputs(config: dict) -> dict:
    """询问用户输入文件路径等"""
    print("\n" + "=" * 50)
    print("=== 文件路径配置 ===")
    print("=" * 50)

    # 1. 询问产品信息表路径
    sku_path = input("请输入产品信息表路径（SKU编码申请表）: ").strip()
    config['sku_application_path'] = validate_file_path(sku_path, "产品信息表")
    print(f"  [OK] 产品信息表: {config['sku_application_path']}")

    # 2. 询问listing表路径
    listing_path = input("请输入Listing文案表路径: ").strip()
    config['listing_content_path'] = validate_file_path(listing_path, "Listing文案表")
    print(f"  [OK] Listing文案表: {config['listing_content_path']}")

    # 3. 询问listing表的sheet名称
    listing_sheet = input("请输入Listing文案表的工作表名称: ").strip()
    if not listing_sheet:
        raise ValueError("Listing文案表工作表名称不能为空")
    config['listing_content_sheet'] = listing_sheet

    # 4. 询问listing标题
    print("\n" + "=" * 50)
    print("=== Listing标题配置 ===")
    print("将在Listing文案表中查找您输入的标题，然后读取下方的文案内容")
    print("=" * 50)
    listing_title = input("请输入要查找的Listing标题: ").strip()
    if not listing_title:
        raise ValueError("Listing标题不能为空")
    config['listing_title'] = listing_title

    # 5. 询问需要修改的模板路径（带验证）
    print("\n" + "=" * 50)
    print("=== 需要修改的模板路径 ===")
    print("=" * 50)
    template_path = input("请输入需要修改的模板文件路径: ").strip()
    config['template_path'] = validate_file_path(template_path, "需要修改的模板")
    print(f"  [OK] 模板路径: {config['template_path']}")

    return config


# ==================== 模板列名匹配 ====================

def get_all_columns_by_name(logger: logging.Logger, ws: Worksheet) -> Dict[str, List[int]]:
    """
    读取Template表的Row 4列标题，建立列名到所有列号的映射
    返回: {列标题: [列号1, 列号2, ...]}
    """
    column_map = {}

    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=4, column=col).value
        if header:
            header_clean = str(header).strip()
            if header_clean not in column_map:
                column_map[header_clean] = []
            column_map[header_clean].append(col)

    logger.info(f"Found {len(column_map)} unique column headers in template")

    for key in ['SKU', 'Product Type', 'Item Name', 'Bullet Point', 'Generic Keyword',
                'Special Features', 'Material', 'Frame Material', 'Frame Type', 'Size']:
        if key in column_map:
            logger.info(f"  {key}: {column_map[key]}")

    return column_map


def find_columns(ws: Worksheet, keyword: str) -> List[int]:
    """查找所有包含关键词的列（不区分大小写）"""
    cols = []
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=4, column=col).value
        if header and keyword.lower() in str(header).lower():
            cols.append(col)
    return sorted(cols)


def find_exact_columns(ws: Worksheet, exact_name: str) -> List[int]:
    """查找精确匹配的列"""
    cols = []
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=4, column=col).value
        if header and str(header).strip() == exact_name:
            cols.append(col)
    return sorted(cols)


# ==================== 数据读取函数 ====================

def delete_template_copy(logger: logging.Logger, wb: openpyxl.Workbook) -> None:
    if 'Template_Copy' in wb.sheetnames:
        del wb['Template_Copy']
        logger.info("Deleted existing Template_Copy worksheet")


def read_sku申请表(logger: logging.Logger, config: dict) -> pd.DataFrame:
    """读取SKU申请表，使用配置中的列索引"""
    path = config['sku_application_path']
    sheet = 'Sheet1'  # SKU申请表默认使用Sheet1
    sku_columns = config['sku_columns']
    data_start = 2  # 0-based，从第2行开始是数据

    logger.info(f"Reading SKU application form: {path}")

    df = pd.read_excel(path, sheet_name=sheet, header=None)

    result_df = pd.DataFrame()
    for col_name, col_idx in sku_columns.items():
        result_df[col_name] = df.iloc[data_start:, col_idx].reset_index(drop=True)

    # 过滤掉没有SKU的行
    sku_col_name = '亚马逊SKU'
    if sku_col_name in result_df.columns:
        result_df = result_df[result_df[sku_col_name].notna()]

    logger.info(f"Loaded {len(result_df)} SKUs from application form")
    return result_df


def read_listing_table(logger: logging.Logger, config: dict) -> dict:
    """根据用户输入的标题，在listing表中精确查找并读取下方文案"""
    path = config['listing_content_path']
    sheet = config['listing_content_sheet']
    search_col = config['listing_search']['title_search']['search_column']
    read_col = config['listing_search']['title_search']['read_column']
    offsets = config['listing_search']['content_offsets']
    user_title = config['listing_title']

    logger.info(f"Reading listing table: {path}, sheet: {sheet}")

    df = pd.read_excel(path, sheet_name=sheet, header=None)

    # 精确查找标题
    title_row = None
    logger.info(f"正在搜索标题 '{user_title}'，搜索列: {search_col}")

    for row in range(df.shape[0]):
        cell_value = str(df.iloc[row, search_col]).strip() if pd.notna(df.iloc[row, search_col]) else ''
        if cell_value == user_title:
            title_row = row
            break

    if title_row is None:
        raise ValueError(f"未找到标题 '{user_title}'，请检查标题是否正确")

    logger.info(f"Found title '{user_title}' at row {title_row}")

    # 根据偏移量读取内容（在read_col列）
    result = {
        'item_name': str(df.iloc[title_row + offsets['item_name'], read_col]) if pd.notna(df.iloc[title_row + offsets['item_name'], read_col]) else '',
        'bullet_points': [],
        'product_description': str(df.iloc[title_row + offsets['product_description'], read_col]) if pd.notna(df.iloc[title_row + offsets['product_description'], read_col]) else '',
        'generic_keyword': str(df.iloc[title_row + offsets['generic_keyword'], read_col]) if pd.notna(df.iloc[title_row + offsets['generic_keyword'], read_col]) else '',
    }

    for offset in offsets['bullet_points']:
        bp = str(df.iloc[title_row + offset, read_col]) if pd.notna(df.iloc[title_row + offset, read_col]) else ''
        result['bullet_points'].append(bp)

    logger.info(f"Loaded item_name: {result['item_name'][:50] if result['item_name'] else 'None'}...")
    logger.info(f"Loaded {len(result['bullet_points'])} bullet points")

    return result


def process_item_name(item_name_template: str, size_str: str) -> str:
    """处理Item Name模板，替换【size】和【num】pcs"""
    if not item_name_template or not size_str:
        return item_name_template

    numbers = re.findall(r'[\d.]+', str(size_str))

    if len(numbers) >= 2:
        size_value = f'{numbers[0]}" x {numbers[1]}"'
        num_value = numbers[2] if len(numbers) >= 3 else None
    else:
        size_value = size_str
        num_value = None

    result = re.sub(r'【size】', size_value, item_name_template)

    if num_value:
        result = re.sub(r'【num】\s*pcs', num_value, result)
    else:
        result = re.sub(r'\s*[xX]\s*【num】\s*pcs', '', result)

    return result


def get_item_type_keyword_from_template(logger: logging.Logger, wb: openpyxl.Workbook) -> str:
    """从模板的Valid Values工作表获取Item Type Keyword的默认值"""
    logger.info("Reading Item Type Keyword from template")

    try:
        ws_valid = wb['Valid Values']
        for row in range(1, min(ws_valid.max_row + 1, 500)):
            field_name = ws_valid.cell(row=row, column=2).value
            if field_name and 'Item Type Keyword' in str(field_name):
                for col in range(3, min(ws_valid.max_column + 1, 50)):
                    val = ws_valid.cell(row=row, column=col).value
                    if val and str(val).strip():
                        logger.info(f"Found Item Type Keyword: {val}")
                        return str(val).strip()
    except Exception as e:
        logger.warning(f"Error reading Item Type Keyword from template: {e}")

    return 'Paintings'


# ==================== 模板处理函数 ====================

def load_and_prepare_template(logger: logging.Logger, config: dict) -> tuple:
    """加载模板并准备"""
    template_path = config['template_path']
    logger.info(f"Opening template: {template_path}")

    wb = load_workbook(template_path, keep_vba=True)
    delete_template_copy(logger, wb)

    if 'Template' not in wb.sheetnames:
        raise ValueError("Template sheet not found in workbook")

    template_ws = wb['Template']
    copy_ws = wb.copy_worksheet(template_ws)
    copy_ws.title = 'Template_Copy'

    logger.info("Created new copy: Template_Copy")

    column_map = get_all_columns_by_name(logger, copy_ws)

    return wb, column_map


def fill_template_with_data(
    logger: logging.Logger,
    wb: openpyxl.Workbook,
    column_map: Dict[str, List[int]],
    sku_df: pd.DataFrame,
    油画指标分析: Dict[str, Any],
    config: dict
) -> None:
    """向模板副本填充数据"""
    logger.info("Filling template with data")

    ws = wb['Template_Copy']

    item_type_keyword = get_item_type_keyword_from_template(logger, wb)
    start_row = config['template']['data_start_row']
    item_name_template = 油画指标分析['item_name']

    # 获取固定值配置
    fv = config['fixed_values']

    # 遍历SKU数据行
    for idx, row in sku_df.iterrows():
        target_row = start_row + idx

        sku = str(row['亚马逊SKU']) if pd.notna(row['亚马逊SKU']) else ''
        型号 = str(row['型号']) if pd.notna(row['型号']) else ''
        作品名称 = str(row['作品名称']) if pd.notna(row['作品名称']) else ''

        logger.info(f"Filling row {target_row}: SKU={sku}, 型号={型号}")

        # 处理Item Name
        item_name = process_item_name(item_name_template, 型号)
        color = re.sub(r'《|》', '', 作品名称).strip()

        # A. 核心必填列
        _set_cell(ws, column_map, 'SKU', target_row, sku, index=0)
        _set_cell(ws, column_map, 'Product Type', target_row, config['category']['product_type'], index=0)
        _set_cell(ws, column_map, 'Listing Action', target_row, fv['core']['listing_action'], index=0)
        _set_cell(ws, column_map, 'Item Name', target_row, item_name, index=0)
        _set_cell(ws, column_map, 'Brand Name', target_row, fv['core']['brand_name'], index=0)
        _set_cell(ws, column_map, 'Product Id Type', target_row, fv['core']['product_id_type'], index=0)
        _set_cell(ws, column_map, 'Item Type Keyword', target_row, item_type_keyword, index=0)
        _set_cell(ws, column_map, 'Manufacturer', target_row, fv['core']['manufacturer'], index=0)
        _set_cell(ws, column_map, 'Product Description', target_row, 油画指标分析['product_description'], index=0)

        # Bullet Points - 填满5列
        bullet_points = 油画指标分析['bullet_points']
        bp_cols = column_map.get('Bullet Point', [])
        for i in range(len(bp_cols)):
            value = bullet_points[i] if i < len(bullet_points) else ''
            ws.cell(row=target_row, column=bp_cols[i], value=value)

        # Generic Keyword - 只填第一列
        gk_cols = column_map.get('Generic Keyword', [])
        if gk_cols:
            ws.cell(row=target_row, column=gk_cols[0], value=油画指标分析['generic_keyword'])

        # B. 产品详情列（固定值）
        sf_cols = column_map.get('Special Features', [])
        special_features = fv['multi_value']['special_features']
        for i in range(len(sf_cols)):
            value = special_features[i] if i < len(special_features) else ''
            ws.cell(row=target_row, column=sf_cols[i], value=value)

        _set_cell(ws, column_map, 'Style', target_row, fv['frame']['style'], index=0)

        mat_cols = column_map.get('Material', [])
        material = fv['multi_value']['material']
        for i in range(len(mat_cols)):
            value = material[i] if i < len(material) else ''
            ws.cell(row=target_row, column=mat_cols[i], value=value)

        _set_cell(ws, column_map, 'Number of Items', target_row, fv['product_details']['number_of_items'], index=0)
        _set_cell(ws, column_map, 'Color', target_row, color, index=0)

        # C. 尺寸与规格列
        规格长 = row['规格长']
        规格宽 = row['规格宽']
        try:
            规格长_int = int(float(规格长)) if pd.notna(规格长) else 0
            规格宽_int = int(float(规格宽)) if pd.notna(规格宽) else 0
        except (ValueError, TypeError):
            规格长_int = 0
            规格宽_int = 0

        numbers = re.findall(r'[\d.]+', str(型号))
        if len(numbers) >= 3:
            size_str = f'{numbers[0]}" x {numbers[1]}" x {numbers[2]}'
        elif len(numbers) >= 2:
            size_str = f'{numbers[0]}" x {numbers[1]}"'
        else:
            size_str = f'{规格长_int}" x {规格宽_int}"'

        _set_cell(ws, column_map, 'Size', target_row, size_str, index=0)

        _set_cell(ws, column_map, 'Item Length Longer Edge', target_row, 规格长_int, index=0)
        _set_cell(ws, column_map, 'Item Width Shorter Edge', target_row, 规格宽_int, index=0)
        _set_cell(ws, column_map, 'Item Length Unit', target_row, fv['units']['item_length_unit'], index=0)
        _set_cell(ws, column_map, 'Item Width Unit', target_row, fv['units']['item_width_unit'], index=0)

        # D. 框架与表面处理列（固定值）
        _set_cell(ws, column_map, 'Item Shape', target_row, fv['product_details']['item_shape'], index=0)

        fm_cols = column_map.get('Frame Material', [])
        if fm_cols:
            ws.cell(row=target_row, column=fm_cols[0], value=fv['frame']['frame_material'])

        ft_cols = column_map.get('Frame Type', [])
        if ft_cols:
            ws.cell(row=target_row, column=ft_cols[0], value=fv['frame']['frame_type'])

        _set_cell(ws, column_map, 'Frame Color', target_row, fv['frame']['frame_color'], index=0)
        _set_cell(ws, column_map, 'Mounting Type', target_row, fv['frame']['mounting_type'], index=0)
        _set_cell(ws, column_map, 'Finish Type', target_row, fv['frame']['finish_type'], index=0)
        _set_cell(ws, column_map, 'Wall Art Form', target_row, fv['frame']['wall_art_form'], index=0)

        # E. 包装尺寸重量列
        _set_cell(ws, column_map, 'Item Package Length', target_row, 规格长, index=0)
        _set_cell(ws, column_map, 'Item Package Width', target_row, 规格宽, index=0)
        _set_cell(ws, column_map, 'Item Package Height', target_row, row['规格高'], index=0)
        _set_cell(ws, column_map, 'Package Length Unit', target_row, fv['units']['package_length_unit'], index=0)
        _set_cell(ws, column_map, 'Package Width Unit', target_row, fv['units']['package_width_unit'], index=0)
        _set_cell(ws, column_map, 'Package Height Unit', target_row, fv['units']['package_height_unit'], index=0)
        _set_cell(ws, column_map, 'Package Weight', target_row, row['重量LB_净重'], index=0)
        _set_cell(ws, column_map, 'Package Weight Unit', target_row, fv['units']['package_weight_unit'], index=0)
        _set_cell(ws, column_map, 'Item Weight', target_row, row['重量LB_净重'], index=0)
        _set_cell(ws, column_map, 'Item Weight Unit', target_row, fv['units']['item_weight_unit'], index=0)

        # F. 单位数量列
        _set_cell(ws, column_map, 'Unit Count', target_row, fv['compliance']['unit_count'], index=0)

        # G. 库存与物流列
        _set_cell(ws, column_map, 'Item Condition', target_row, fv['fulfillment']['item_condition'], index=0)
        _set_cell(ws, column_map, 'Quantity (US)', target_row, fv['fulfillment']['quantity_us'], index=0)
        _set_cell(ws, column_map, 'Merchant Shipping Group (US)', target_row, fv['fulfillment']['merchant_shipping_group'], index=0)
        _set_cell(ws, column_map, 'Fulfillment Channel Code (US)', target_row, fv['fulfillment']['fulfillment_channel_code_us'], index=0)

        # H. 合规性列
        _set_cell(ws, column_map, 'Country of Origin', target_row, fv['compliance']['country_of_origin'], index=0)

    logger.info(f"Filled {len(sku_df)} rows successfully")


def _set_cell(ws: Worksheet, column_map: Dict[str, List[int]], col_name: str, row: int, value: Any, index: int = 0) -> None:
    """通过列名安全地设置单元格值"""
    if col_name in column_map and index < len(column_map[col_name]):
        col = column_map[col_name][index]
        ws.cell(row=row, column=col, value=value)


def save_template(logger: logging.Logger, wb: openpyxl.Workbook, config: dict) -> None:
    """保存模板文件"""
    template_path = config['template_path']
    logger.info(f"Saving template: {template_path}")
    wb.save(template_path)
    logger.info("Template saved successfully")


def copy_values_to_template(logger: logging.Logger, config: dict, start_row: int = 9, num_rows: int = 12) -> None:
    """将Template_Copy中有内容的行复制到Template表中（只复制值）"""
    template_path = config['template_path']
    logger.info(f"Copying values from Template_Copy to Template (rows {start_row}-{start_row + num_rows - 1})")

    wb = load_workbook(template_path, keep_vba=True)

    copy_ws = wb['Template_Copy']
    template_ws = wb['Template']

    max_col = copy_ws.max_column

    for row_idx in range(num_rows):
        source_row = start_row + row_idx
        target_row = start_row + row_idx

        for col in range(1, max_col + 1):
            cell_value = copy_ws.cell(row=source_row, column=col).value
            template_ws.cell(row=target_row, column=col, value=cell_value)

        logger.info(f"Copied row {source_row} to row {target_row}")

    if 'Template_Copy' in wb.sheetnames:
        del wb['Template_Copy']

    wb.save(template_path)
    logger.info("Values copied to Template successfully")
    logger.info("Template_Copy worksheet deleted")


# ==================== 主函数 ====================

def main():
    """主函数"""
    logger = setup_logging()

    try:
        # 1. 选择类目
        category = select_category()
        logger.info(f"Selected category: {category}")

        # 2. 加载配置
        config = load_category_config(category)
        logger.info(f"Loaded config for category: {config['category']['name']}")

        # 3. 询问用户输入
        config = ask_user_inputs(config)

        # 4. 读取数据
        sku_df = read_sku申请表(logger, config)
        listing_data = read_listing_table(logger, config)

        # 5. 填充模板
        wb, column_map = load_and_prepare_template(logger, config)
        fill_template_with_data(logger, wb, column_map, sku_df, listing_data, config)
        save_template(logger, wb, config)

        logger.info("Data filled successfully in Template_Copy!")
        logger.info("Please verify the data in Template_Copy sheet.")

        # 6. 询问是否复制到Template
        confirm = input("\n是否将数据复制到Template表？(yes/no): ").strip().lower()
        if confirm in ['yes', 'y', '是']:
            copy_values_to_template(logger, config)
            logger.info("Process completed successfully!")
        else:
            logger.info("Cancelled. Data remains in Template_Copy sheet.")

    except Exception as e:
        logger.error(f"错误: {str(e)}")
        raise


if __name__ == '__main__':
    main()
