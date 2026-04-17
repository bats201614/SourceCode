"""
亚马逊批量上传模板自动化脚本
从产品信息表和listing表读取数据，填充到HighFashion.xlsm模板
"""

import pandas as pd
import openpyxl
import os
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import Dict, List, Any
import logging
import re

# ==================== 配置常量 ====================

# 源文件路径（运行时询问用户）
产品信息表_PATH = None
listing表_PATH = None
目标文件_PATH = None

# 源文件工作表
产品信息表_SHEET = None
listing表_SHEET = None

# 产品信息表列索引（0-based）
COL_序号 = 0
COL_材质 = 1
COL_作品名称 = 2
COL_型号 = 3
COL_PCS数 = 4
COL_颜色 = 5
COL_规格长 = 6
COL_规格宽 = 7
COL_规格高 = 8
COL_重量LB_毛重 = 9
COL_重量LB_净重 = 10
COL_亚马逊SKU = 19
COL_商品名称 = 23

# 固定值配置
FIXED_VALUES = {
    'product_type': 'WALL_ART',
    'brand_name': 'IUI',
    'product_id_type': 'GTIN Exempt',
    'listing_action': 'Create or Replace (Full Update)',
    'item_type_keyword': 'Paintings',
    'manufacturer': 'IUI',
    'parentage_level': 'Child',
    'variation_theme': 'Size/Color',
    'special_features': ['Water Resistant', 'Fade Resistant', 'Scratch Resistant', 'Durable'],
    'style': 'Abstract',
    'material': ['Cotton', 'Wood'],
    'number_of_items': 1,
    'item_shape': 'Rectangular',
    'frame_color': 'Brown',
    'frame_material': 'Wood',
    'frame_type': 'Unframed',
    'mounting_type': 'Wall Mount',
    'finish_type': 'Painted',
    'wall_art_form': 'Art Print',
    'unit_count': 1,
    'item_length_unit': 'Inches',
    'item_width_unit': 'Inches',
    'item_condition': 'New',
    'quantity_us': 100,
    'merchant_shipping_group': 'Migrated Template',
    'fulfillment_channel_code_us': 'DEFAULT',
    'package_length_unit': 'Centimeters',
    'package_width_unit': 'Centimeters',
    'package_height_unit': 'Centimeters',
    'package_weight_unit': 'Kilograms',
    'item_weight_unit': 'Pounds',
    'country_of_origin': 'United States',
}

# ==================== 日志配置 ====================

def setup_logging() -> logging.Logger:
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    return logging.getLogger(__name__)


# ==================== 模板列名匹配 ====================

def get_all_columns_by_name(logger: logging.Logger, ws: Worksheet) -> Dict[str, List[int]]:
    """
    读取Template表的Row 4列标题，建立列名到所有列号的映射
    返回: {列标题: [列号1, 列号2, ...]}
    """
    column_map = {}  # {列标题: [列号列表]}

    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=4, column=col).value
        if header:
            header_clean = str(header).strip()
            if header_clean not in column_map:
                column_map[header_clean] = []
            column_map[header_clean].append(col)

    logger.info(f"Found {len(column_map)} unique column headers in template")

    # 打印关键列的位置
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


# ==================== 自动检测列位置 ====================

def auto_detect_columns(logger: logging.Logger, df: pd.DataFrame, header_row: int = 0) -> Dict[str, int]:
    """
    读取表头行，通过关键词匹配找到对应列索引
    支持单行或双行表头结构（Row 0是主分类，Row 1是子列名）
    返回: {'亚马逊SKU': 19, '型号': 3, ...}
    对于每个列：先查Row 0，找不到再到Row 1找（适用于合并单元格情况）
    """
    headers_row0 = df.iloc[header_row]
    has_row1 = header_row + 1 < len(df)

    col_map = {}

    def get_header_value(row_idx: int, col_idx: int) -> str:
        """安全获取表头值"""
        try:
            val = df.iloc[row_idx, col_idx] if row_idx < len(df) else None
            return str(val).strip() if val is not None and pd.notna(val) else ''
        except Exception:
            return ''

    def try_match_Col(h: str, *keywords: str) -> bool:
        """检查表头值是否包含任意一个关键词"""
        for kw in keywords:
            if kw in h:
                return True
        return False

    # 用于追踪当前所属的主标题（处理合并单元格）
    current_parent_header = None

    for i in range(len(headers_row0)):
        h0 = get_header_value(header_row, i)
        h1 = get_header_value(header_row + 1, i) if has_row1 else ''

        # 如果h0不为空，可能是新的主标题（用于处理合并单元格子列）
        if h0:
            if '规格' in h0:
                current_parent_header = '规格'
            elif '重量' in h0 and 'LB' in h0:
                current_parent_header = '重量'
            elif '销售包装' in h0 and '重量' in h0:
                current_parent_header = '销售包装重量'
            elif '销售包装' in h0:
                current_parent_header = '销售包装'
            else:
                current_parent_header = None

        # 匹配亚马逊SKU列 - Row0优先，Row1备选
        if '亚马逊SKU' not in col_map:
            if try_match_Col(h0, 'SKU', '亚马逊SKU') and try_match_Col(h0, '亚马逊', 'Amazon', 'amazon'):
                col_map['亚马逊SKU'] = i
            elif try_match_Col(h1, 'SKU', '亚马逊SKU') and try_match_Col(h1, '亚马逊', 'Amazon', 'amazon'):
                col_map['亚马逊SKU'] = i

        # 匹配型号列 - Row0优先，Row1备选
        if '型号' not in col_map:
            if try_match_Col(h0, '型号') and not try_match_Col(h0, '图片'):
                col_map['型号'] = i
            elif try_match_Col(h1, '型号') and not try_match_Col(h1, '图片'):
                col_map['型号'] = i

        # 匹配作品名称列
        if '作品名称' not in col_map:
            if try_match_Col(h0, '作品名称', '商品名称'):
                col_map['作品名称'] = i
            elif try_match_Col(h1, '作品名称', '商品名称'):
                col_map['作品名称'] = i

        # 匹配规格长/宽/高 - 当current_parent_header是'规格'时（Row0有主标题，Row1是子列）
        if current_parent_header == '规格':
            if '规格长' not in col_map and try_match_Col(h1, '长') and not try_match_Col(h1, '宽') and not try_match_Col(h1, '高'):
                col_map['规格长'] = i
            elif '规格宽' not in col_map and try_match_Col(h1, '宽') and not try_match_Col(h1, '长'):
                col_map['规格宽'] = i
            elif '规格高' not in col_map and try_match_Col(h1, '高'):
                col_map['规格高'] = i

        # 匹配销售包装尺寸长/宽/高 - 当current_parent_header是'销售包装'时
        if current_parent_header == '销售包装':
            if '销售包装长' not in col_map and try_match_Col(h1, '长') and not try_match_Col(h1, '宽') and not try_match_Col(h1, '高'):
                col_map['销售包装长'] = i
            elif '销售包装宽' not in col_map and try_match_Col(h1, '宽') and not try_match_Col(h1, '长'):
                col_map['销售包装宽'] = i
            elif '销售包装高' not in col_map and try_match_Col(h1, '高'):
                col_map['销售包装高'] = i

        # 匹配重量LB毛重/净重 - 当current_parent_header是'重量'时
        if current_parent_header == '重量':
            if '重量LB_毛重' not in col_map and try_match_Col(h1, '毛重'):
                col_map['重量LB_毛重'] = i
            elif '重量LB_净重' not in col_map and try_match_Col(h1, '净重'):
                col_map['重量LB_净重'] = i

        # 匹配销售包装重量毛重/净重 - 当current_parent_header是'销售包装重量'时
        if current_parent_header == '销售包装重量':
            if '销售包装重量_毛重' not in col_map and try_match_Col(h1, '毛重'):
                col_map['销售包装重量_毛重'] = i
            elif '销售包装重量_净重' not in col_map and try_match_Col(h1, '净重'):
                col_map['销售包装重量_净重'] = i

        # 匹配商品名称列
        if '商品名称' not in col_map:
            if h0 == '商品名称' or h1 == '商品名称':
                col_map['商品名称'] = i

        # 匹配序号列
        if '序号' not in col_map:
            if h0 in ['序号', 'NO.', 'No.'] or h1 in ['序号', 'NO.', 'No.']:
                col_map['序号'] = i

        # 匹配材质列
        if '材质' not in col_map:
            if h0 == '材质' or h1 == '材质':
                col_map['材质'] = i

        # 匹配PCS数列
        if 'PCS数' not in col_map:
            if try_match_Col(h0, 'PCS', 'pcs', '数量') or try_match_Col(h1, 'PCS', 'pcs', '数量'):
                col_map['PCS数'] = i

        # 匹配颜色列
        if '颜色' not in col_map:
            if h0 == '颜色' or h1 == '颜色':
                col_map['颜色'] = i

    logger.info(f"Auto-detected {len(col_map)} columns from header row {header_row}:")
    for name, idx in sorted(col_map.items(), key=lambda x: x[1]):
        logger.info(f"  {name}: column {idx}")

    return col_map


# ==================== 数据读取函数 ====================

def delete_template_copy(logger: logging.Logger, wb: openpyxl.Workbook) -> None:
    if 'Template_Copy' in wb.sheetnames:
        del wb['Template_Copy']
        logger.info("Deleted existing Template_Copy worksheet")


def read_产品信息表(logger: logging.Logger) -> pd.DataFrame:
    logger.info(f"Reading SKU application form: {产品信息表_PATH}")

    df = pd.read_excel(产品信息表_PATH, sheet_name=产品信息表_SHEET, header=None)

    # 自动检测列位置
    col_map = auto_detect_columns(logger, df, header_row=0)

    # 检查必要的列是否存在
    required_cols = ['亚马逊SKU', '型号', '规格长', '规格宽']
    missing = [c for c in required_cols if c not in col_map]
    if missing:
        raise ValueError(f"Cannot find required columns: {missing}")

    data_start = 2  # 数据从第3行开始

    result_df = pd.DataFrame()
    result_df['序号'] = df.iloc[data_start:, col_map.get('序号', 0)].reset_index(drop=True) if '序号' in col_map else None
    result_df['材质'] = df.iloc[data_start:, col_map.get('材质', 1)].reset_index(drop=True) if '材质' in col_map else None
    result_df['作品名称'] = df.iloc[data_start:, col_map.get('作品名称', 2)].reset_index(drop=True) if '作品名称' in col_map else None
    result_df['型号'] = df.iloc[data_start:, col_map['型号']].reset_index(drop=True)
    result_df['PCS数'] = df.iloc[data_start:, col_map.get('PCS数', 4)].reset_index(drop=True) if 'PCS数' in col_map else None
    result_df['颜色'] = df.iloc[data_start:, col_map.get('颜色', 5)].reset_index(drop=True) if '颜色' in col_map else None
    result_df['规格长'] = df.iloc[data_start:, col_map['规格长']].reset_index(drop=True)
    result_df['规格宽'] = df.iloc[data_start:, col_map['规格宽']].reset_index(drop=True)
    result_df['规格高'] = df.iloc[data_start:, col_map.get('规格高', 8)].reset_index(drop=True) if '规格高' in col_map else None
    result_df['重量LB_毛重'] = df.iloc[data_start:, col_map.get('重量LB_毛重', 9)].reset_index(drop=True) if '重量LB_毛重' in col_map else None
    result_df['重量LB_净重'] = df.iloc[data_start:, col_map.get('重量LB_净重', 10)].reset_index(drop=True) if '重量LB_净重' in col_map else None
    result_df['亚马逊SKU'] = df.iloc[data_start:, col_map['亚马逊SKU']].reset_index(drop=True)
    result_df['商品名称'] = df.iloc[data_start:, col_map.get('商品名称', col_map.get('作品名称', 23))].reset_index(drop=True)
    # 销售包装尺寸(cm)长/宽/高
    result_df['销售包装长'] = df.iloc[data_start:, col_map['销售包装长']].reset_index(drop=True) if '销售包装长' in col_map else None
    result_df['销售包装宽'] = df.iloc[data_start:, col_map['销售包装宽']].reset_index(drop=True) if '销售包装宽' in col_map else None
    result_df['销售包装高'] = df.iloc[data_start:, col_map['销售包装高']].reset_index(drop=True) if '销售包装高' in col_map else None
    # 销售包装重量(kg)毛重/净重
    result_df['销售包装重量_毛重'] = df.iloc[data_start:, col_map.get('销售包装重量_毛重', 17)].reset_index(drop=True) if '销售包装重量_毛重' in col_map else None
    result_df['销售包装重量_净重'] = df.iloc[data_start:, col_map.get('销售包装重量_净重', 18)].reset_index(drop=True) if '销售包装重量_净重' in col_map else None

    result_df = result_df[result_df['亚马逊SKU'].notna()]
    logger.info(f"Loaded {len(result_df)} SKUs from application form")
    return result_df


def find_listing_categories(logger: logging.Logger, df: pd.DataFrame) -> List[tuple]:
    """
    查找listing表中的所有大标题（分类）
    返回: [(行号, 标题名称), ...]
    """
    categories = []
    for i in range(len(df)):
        v0 = df.iloc[i, 0]
        v1 = df.iloc[i, 1] if len(df.columns) > 1 else None
        if pd.notna(v0) and str(v0).strip():
            v0_str = str(v0).strip()
            # Category titles have col 1 empty or nan
            is_category = (v1 is None or pd.isna(v1) or str(v1).strip() == '')
            if is_category:
                categories.append((i, v0_str))
    return categories


def read_listing表(logger: logging.Logger) -> Dict[str, Any]:
    """
    读取listing表中的文案数据
    1. 查找所有大标题
    2. 让用户选择使用哪个大标题下的数据
    3. 读取选中分类下的标题、五点描述、产品描述、ST
    """
    logger.info(f"Reading listing indicator analysis: {listing表_PATH}")

    df = pd.read_excel(listing表_PATH, sheet_name=listing表_SHEET, header=None)

    # 查找所有分类
    categories = find_listing_categories(logger, df)

    if not categories:
        raise ValueError("No categories found in listing table")

    # 打印所有分类供用户选择
    print("\n" + "=" * 50)
    print("Listing表中的大标题：")
    print("=" * 50)
    for idx, (row_num, title) in enumerate(categories):
        print(f"  {idx + 1}. {title} (Row {row_num})")
    print("=" * 50)

    # 让用户选择分类
    while True:
        try:
            choice = input("请选择大标题编号 (1-{}): ".format(len(categories))).strip()
            if not choice:
                print("输入不能为空")
                continue
            choice_idx = int(choice) - 1
            if 0 <= choice_idx < len(categories):
                selected_row, selected_title = categories[choice_idx]
                break
            else:
                print(f"请输入 1 到 {len(categories)} 之间的数字")
        except ValueError:
            print("请输入有效的数字")

    logger.info(f"Selected category: {selected_title} at row {selected_row}")

    # 读取选中分类下的数据
    # 结构: Row n=分类标题, Row n+1=标题, Row n+2~n+6=5个五点描述, Row n+7=产品描述, Row n+8=ST
    item_name_row = selected_row + 1
    bullet_start_row = selected_row + 2
    product_desc_row = selected_row + 7
    generic_keyword_row = selected_row + 8

    result = {
        'item_name': None,
        'bullet_points': [],
        'product_description': None,
        'generic_keyword': None,
        'category_title': selected_title,
    }

    try:
        # 读取标题
        if item_name_row < len(df):
            result['item_name'] = str(df.iloc[item_name_row, 1]) if pd.notna(df.iloc[item_name_row, 1]) else ''

        # 读取5个五点描述
        for i in range(5):
            row_idx = bullet_start_row + i
            if row_idx < len(df):
                bp = str(df.iloc[row_idx, 1]) if pd.notna(df.iloc[row_idx, 1]) else ''
                result['bullet_points'].append(bp)

        # 读取产品描述
        if product_desc_row < len(df):
            result['product_description'] = str(df.iloc[product_desc_row, 1]) if pd.notna(df.iloc[product_desc_row, 1]) else ''

        # 读取ST
        if generic_keyword_row < len(df):
            result['generic_keyword'] = str(df.iloc[generic_keyword_row, 1]) if pd.notna(df.iloc[generic_keyword_row, 1]) else ''

        logger.info(f"Loaded item_name: {result['item_name'][:50] if result['item_name'] else 'None'}...")
        logger.info(f"Loaded {len(result['bullet_points'])} bullet points")
        logger.info(f"Loaded generic_keyword: {result['generic_keyword'][:50] if result['generic_keyword'] else 'None'}...")

    except Exception as e:
        logger.warning(f"Error reading listing analysis: {e}")

    return result

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

def load_and_prepare_template(logger: logging.Logger) -> tuple:
    """加载模板并准备"""
    logger.info(f"Opening template: {目标文件_PATH}")

    wb = load_workbook(目标文件_PATH, keep_vba=True)
    delete_template_copy(logger, wb)

    if 'Template' not in wb.sheetnames:
        raise ValueError("Template sheet not found in workbook")

    template_ws = wb['Template']
    copy_ws = wb.copy_worksheet(template_ws)
    copy_ws.title = 'Template_Copy'

    logger.info("Created new copy: Template_Copy")

    # 获取所有列名及其列号列表
    column_map = get_all_columns_by_name(logger, copy_ws)

    return wb, column_map


def fill_template_with_data(
    logger: logging.Logger,
    wb: openpyxl.Workbook,
    column_map: Dict[str, List[int]],
    sku_df: pd.DataFrame,
    listing表: Dict[str, Any]
) -> None:
    """向模板副本填充数据"""
    logger.info("Filling template with data")

    ws = wb['Template_Copy']

    item_type_keyword = get_item_type_keyword_from_template(logger, wb)
    start_row = 9  # 从第9行开始填写，第8行及之前是带颜色的示例单元格
    item_name_template = listing表['item_name']

    # 遍历SKU数据行
    for idx, row in sku_df.iterrows():
        target_row = start_row + idx # type: ignore

        sku = str(row['亚马逊SKU']) if pd.notna(row['亚马逊SKU']) else ''
        型号 = str(row['型号']) if pd.notna(row['型号']) else ''
        作品名称 = str(row['作品名称']) if pd.notna(row['作品名称']) else ''

        logger.info(f"Filling row {target_row}: SKU={sku}, 型号={型号}")

        # 处理Item Name
        item_name = process_item_name(item_name_template, 型号)
        color = re.sub(r'《|》', '', 作品名称).strip()

        # A. 核心必填列
        _set_cell(ws, column_map, 'SKU', target_row, sku, index=0)
        _set_cell(ws, column_map, 'Product Type', target_row, FIXED_VALUES['product_type'], index=0)
        _set_cell(ws, column_map, 'Listing Action', target_row, FIXED_VALUES['listing_action'], index=0)
        _set_cell(ws, column_map, 'Item Name', target_row, item_name, index=0)
        _set_cell(ws, column_map, 'Brand Name', target_row, FIXED_VALUES['brand_name'], index=0)
        _set_cell(ws, column_map, 'Product Id Type', target_row, FIXED_VALUES['product_id_type'], index=0)
        _set_cell(ws, column_map, 'Item Type Keyword', target_row, item_type_keyword, index=0)
        _set_cell(ws, column_map, 'Manufacturer', target_row, FIXED_VALUES['manufacturer'], index=0)
        _set_cell(ws, column_map, 'Product Description', target_row, listing表['product_description'], index=0)

        # Bullet Points - 填满5列
        bullet_points = listing表['bullet_points']
        bp_cols = column_map.get('Bullet Point', [])
        for i in range(len(bp_cols)):
            value = bullet_points[i] if i < len(bullet_points) else ''
            ws.cell(row=target_row, column=bp_cols[i], value=value)

        # Generic Keyword - 只填第一列
        gk_cols = column_map.get('Generic Keyword', [])
        if gk_cols:
            ws.cell(row=target_row, column=gk_cols[0], value=listing表['generic_keyword'])

        # B. 产品详情列（固定值）
        # Special Features - 填满5列（4个固定值+1个空或重复）
        sf_cols = column_map.get('Special Features', [])
        for i in range(len(sf_cols)):
            value = FIXED_VALUES['special_features'][i] if i < len(FIXED_VALUES['special_features']) else ''
            ws.cell(row=target_row, column=sf_cols[i], value=value)

        _set_cell(ws, column_map, 'Style', target_row, FIXED_VALUES['style'], index=0)

        # Material - 填满5列
        mat_cols = column_map.get('Material', [])
        for i in range(len(mat_cols)):
            value = FIXED_VALUES['material'][i] if i < len(FIXED_VALUES['material']) else ''
            ws.cell(row=target_row, column=mat_cols[i], value=value)

        _set_cell(ws, column_map, 'Number of Items', target_row, FIXED_VALUES['number_of_items'], index=0)
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

        # Size格式: 如果型号有3个数字填3个，否则填前两个
        numbers = re.findall(r'[\d.]+', str(型号))
        if len(numbers) >= 3:
            size_str = f'{numbers[0]}" x {numbers[1]}" x {numbers[2]}'
        elif len(numbers) >= 2:
            size_str = f'{numbers[0]}" x {numbers[1]}"'
        else:
            size_str = f'{规格长_int}" x {规格宽_int}"'

        _set_cell(ws, column_map, 'Size', target_row, size_str, index=0)

        # Item Length/Width 去掉小数点，转为整数
        规格长_int = int(float(规格长)) if pd.notna(规格长) else 0
        规格宽_int = int(float(规格宽)) if pd.notna(规格宽) else 0
        _set_cell(ws, column_map, 'Item Length Longer Edge', target_row, 规格长_int, index=0)
        _set_cell(ws, column_map, 'Item Width Shorter Edge', target_row, 规格宽_int, index=0)
        _set_cell(ws, column_map, 'Item Length Unit', target_row, FIXED_VALUES['item_length_unit'], index=0)
        _set_cell(ws, column_map, 'Item Width Unit', target_row, FIXED_VALUES['item_width_unit'], index=0)

        # D. 框架与表面处理列（固定值）
        _set_cell(ws, column_map, 'Item Shape', target_row, FIXED_VALUES['item_shape'], index=0)

        # Frame Material - 只填第一列
        fm_cols = column_map.get('Frame Material', [])
        if fm_cols:
            ws.cell(row=target_row, column=fm_cols[0], value=FIXED_VALUES['frame_material'])

        # Frame Type - 只填第一列
        ft_cols = column_map.get('Frame Type', [])
        if ft_cols:
            ws.cell(row=target_row, column=ft_cols[0], value=FIXED_VALUES['frame_type'])

        _set_cell(ws, column_map, 'Frame Color', target_row, FIXED_VALUES['frame_color'], index=0)
        _set_cell(ws, column_map, 'Mounting Type', target_row, FIXED_VALUES['mounting_type'], index=0)
        _set_cell(ws, column_map, 'Finish Type', target_row, FIXED_VALUES['finish_type'], index=0)
        _set_cell(ws, column_map, 'Wall Art Form', target_row, FIXED_VALUES['wall_art_form'], index=0)

        # E. 包装尺寸重量列 - 使用销售包装尺寸(cm)和销售包装重量(kg)
        _set_cell(ws, column_map, 'Item Package Length', target_row, row['销售包装长'], index=0)
        _set_cell(ws, column_map, 'Item Package Width', target_row, row['销售包装宽'], index=0)
        _set_cell(ws, column_map, 'Item Package Height', target_row, row['销售包装高'], index=0)
        _set_cell(ws, column_map, 'Package Length Unit', target_row, FIXED_VALUES['package_length_unit'], index=0)
        _set_cell(ws, column_map, 'Package Width Unit', target_row, FIXED_VALUES['package_width_unit'], index=0)
        _set_cell(ws, column_map, 'Package Height Unit', target_row, FIXED_VALUES['package_height_unit'], index=0)
        _set_cell(ws, column_map, 'Package Weight', target_row, round(float(row['销售包装重量_净重']), 2) if pd.notna(row['销售包装重量_净重']) else None, index=0)
        _set_cell(ws, column_map, 'Package Weight Unit', target_row, FIXED_VALUES['package_weight_unit'], index=0)
        _set_cell(ws, column_map, 'Item Weight', target_row, row['重量LB_净重'], index=0)
        _set_cell(ws, column_map, 'Item Weight Unit', target_row, FIXED_VALUES['item_weight_unit'], index=0)

        # F. 单位数量列
        _set_cell(ws, column_map, 'Unit Count', target_row, FIXED_VALUES['unit_count'], index=0)

        # G. 库存与物流列
        _set_cell(ws, column_map, 'Item Condition', target_row, FIXED_VALUES['item_condition'], index=0)
        _set_cell(ws, column_map, 'Quantity (US)', target_row, FIXED_VALUES['quantity_us'], index=0)
        _set_cell(ws, column_map, 'Merchant Shipping Group (US)', target_row, FIXED_VALUES['merchant_shipping_group'], index=0)
        _set_cell(ws, column_map, 'Fulfillment Channel Code (US)', target_row, FIXED_VALUES['fulfillment_channel_code_us'], index=0)

        # H. 合规性列
        _set_cell(ws, column_map, 'Country of Origin', target_row, FIXED_VALUES['country_of_origin'], index=0)

    logger.info(f"Filled {len(sku_df)} rows successfully")


def _set_cell(ws: Worksheet, column_map: Dict[str, List[int]], col_name: str, row: int, value: Any, index: int = 0) -> None:
    """通过列名安全地设置单元格值"""
    if col_name in column_map and index < len(column_map[col_name]):
        col = column_map[col_name][index]
        ws.cell(row=row, column=col, value=value)


def save_template(logger: logging.Logger, wb: openpyxl.Workbook) -> None:
    """保存模板文件"""
    logger.info(f"Saving template: {目标文件_PATH}")
    wb.save(目标文件_PATH)
    logger.info("Template saved successfully")


def copy_values_to_template(logger: logging.Logger, start_row: int = 9, num_rows: int = 12) -> None:
    """
    将Template_Copy中有内容的行复制到Template表中（只复制值）

    Args:
        start_row: 起始行（第9行）
        num_rows: 要复制的行数（12个SKU）
    """
    logger.info(f"Copying values from Template_Copy to Template (rows {start_row}-{start_row + num_rows - 1})")

    # 打开模板文件
    wb = load_workbook(目标文件_PATH, keep_vba=True)

    copy_ws = wb['Template_Copy']
    template_ws = wb['Template']

    # 获取最大列数
    max_col = copy_ws.max_column

    # 复制每一行的值到Template
    for row_idx in range(num_rows):
        source_row = start_row + row_idx
        target_row = start_row + row_idx

        for col in range(1, max_col + 1):
            # 只复制值，不复制公式
            cell_value = copy_ws.cell(row=source_row, column=col).value
            template_ws.cell(row=target_row, column=col, value=cell_value)

        logger.info(f"Copied row {source_row} to row {target_row}")

    # 删除Template_Copy表
    if 'Template_Copy' in wb.sheetnames:
        del wb['Template_Copy']

    # 保存
    wb.save(目标文件_PATH)
    logger.info("Values copied to Template successfully")
    logger.info("Template_Copy worksheet deleted")

    return wb # type: ignore


# ==================== 主函数 ====================

def main():
    """主函数"""
    global 产品信息表_PATH, listing表_PATH, 目标文件_PATH
    global 产品信息表_SHEET, listing表_SHEET

    logger = setup_logging()

    def clean_path(p: str) -> str:
        """去除路径首尾的双引号和空格，兼容复制粘贴的带引号路径"""
        return p.strip().strip('"').strip()

    def ask_path(prompt: str) -> str:
        """循环询问路径，直到文件存在或用户选择退出"""
        while True:
            path = clean_path(input(prompt))
            if not path:
                print("路径不能为空。")
                continue
            if not os.path.isfile(path):
                print(f"文件不存在：{path}")
                retry = input("按 Enter 重新输入，或输入 q 退出：").strip().lower()
                if retry == 'q':
                    print("已退出。")
                    raise SystemExit(0)
                continue
            return path

    def ask_sheet(prompt: str, default: str) -> str:
        """循环询问工作表名，直到工作表存在或用户选择退出"""
        sheet_name = input(f"{prompt}（默认：{default}）：").strip()
        return sheet_name if sheet_name else default

    def validate_sheet(path: str, sheet: str) -> bool:
        """验证工作表是否存在"""
        try:
            xl = pd.ExcelFile(path)
            return sheet in xl.sheet_names
        except Exception:
            return False

    print("=" * 50)
    print("请依次输入以下文件路径和信息：")
    print("=" * 50)

    产品信息表_PATH = ask_path("产品信息表路径：")
    产品信息表_SHEET = ask_sheet("产品信息表工作表名", "Sheet1")
    while not validate_sheet(产品信息表_PATH, 产品信息表_SHEET):
        print(f"工作表「{产品信息表_SHEET}」不存在。")
        retry = input("按 Enter 重新输入工作表名，或输入 q 退出：").strip().lower()
        if retry == 'q':
            print("已退出。")
            raise SystemExit(0)
        产品信息表_SHEET = ask_sheet("产品信息表工作表名", "Sheet1")

    listing表_PATH = ask_path("listing表路径：")
    listing表_SHEET = ask_sheet("listing表工作表名", "链接文案与尺寸定价")
    while not validate_sheet(listing表_PATH, listing表_SHEET):
        print(f"工作表「{listing表_SHEET}」不存在。")
        retry = input("按 Enter 重新输入工作表名，或输入 q 退出：").strip().lower()
        if retry == 'q':
            print("已退出。")
            raise SystemExit(0)
        listing表_SHEET = ask_sheet("listing表工作表名", "链接文案与尺寸定价")

    目标文件_PATH = ask_path("目标模板文件(.xlsm)路径：")
    print("=" * 50)

    try:
        sku_df = read_产品信息表(logger)
        listing表 = read_listing表(logger)
        wb, column_map = load_and_prepare_template(logger)
        fill_template_with_data(logger, wb, column_map, sku_df, listing表)
        save_template(logger, wb)

        logger.info("Data filled successfully in Template_Copy!")
        logger.info("Please verify the data in Template_Copy sheet.")

        # 询问用户是否确认复制到Template
        confirm = input("Do you want to copy values to Template sheet? (yes/no): ").strip().lower()
        if confirm in ['yes', 'y', '是', '确认']:
            copy_values_to_template(logger)
            logger.info("Process completed successfully!")
        else:
            logger.info("Cancelled. Data remains in Template_Copy sheet.")

    except Exception as e:
        logger.error(f"Error occurred: {str(e)}")
        raise


if __name__ == '__main__':
    main()
