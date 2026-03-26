"""
亚马逊批量上传模板自动化脚本
从SKU编码申请表和油画指标分析表读取数据，填充到HighFashion.xlsm模板
"""

import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import Dict, List, Any
import logging
import re

# ==================== 配置常量 ====================

# 源文件路径
SKU申请表_PATH = r'C:\Users\Administrator\Desktop\个人\分析\油画\3.产品信息与模板\SKU编码申请表\HighFashion油画SKU编码申请(1).xlsx'
油画指标分析_PATH = r'C:\Users\Administrator\Desktop\个人\分析\油画\3.产品信息与模板\油画指标分析.xlsx'
目标文件_PATH = r'C:\Users\Administrator\Desktop\个人\分析\油画\链接上传\HighFashion.xlsm'

# 源文件工作表
SKU申请表_SHEET = 'Sheet1'
油画指标分析_SHEET = '链接文案与尺寸定价'

# SKU编码申请表列索引（0-based）
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
    'listing_action': '(Default) Create or Replace',
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


# ==================== 数据读取函数 ====================

def delete_template_copy(logger: logging.Logger, wb: openpyxl.Workbook) -> None:
    if 'Template_Copy' in wb.sheetnames:
        del wb['Template_Copy']
        logger.info("Deleted existing Template_Copy worksheet")


def read_sku申请表(logger: logging.Logger) -> pd.DataFrame:
    logger.info(f"Reading SKU application form: {SKU申请表_PATH}")

    df = pd.read_excel(SKU申请表_PATH, sheet_name=SKU申请表_SHEET, header=None)
    data_start = 2

    result_df = pd.DataFrame()
    result_df['序号'] = df.iloc[data_start:, COL_序号].reset_index(drop=True)
    result_df['材质'] = df.iloc[data_start:, COL_材质].reset_index(drop=True)
    result_df['作品名称'] = df.iloc[data_start:, COL_作品名称].reset_index(drop=True)
    result_df['型号'] = df.iloc[data_start:, COL_型号].reset_index(drop=True)
    result_df['PCS数'] = df.iloc[data_start:, COL_PCS数].reset_index(drop=True)
    result_df['颜色'] = df.iloc[data_start:, COL_颜色].reset_index(drop=True)
    result_df['规格长'] = df.iloc[data_start:, COL_规格长].reset_index(drop=True)
    result_df['规格宽'] = df.iloc[data_start:, COL_规格宽].reset_index(drop=True)
    result_df['规格高'] = df.iloc[data_start:, COL_规格高].reset_index(drop=True)
    result_df['重量LB_毛重'] = df.iloc[data_start:, COL_重量LB_毛重].reset_index(drop=True)
    result_df['重量LB_净重'] = df.iloc[data_start:, COL_重量LB_净重].reset_index(drop=True)
    result_df['亚马逊SKU'] = df.iloc[data_start:, COL_亚马逊SKU].reset_index(drop=True)
    result_df['商品名称'] = df.iloc[data_start:, COL_商品名称].reset_index(drop=True)

    result_df = result_df[result_df['亚马逊SKU'].notna()]
    logger.info(f"Loaded {len(result_df)} SKUs from application form")
    return result_df


def read_油画指标分析(logger: logging.Logger) -> Dict[str, Any]:
    logger.info(f"Reading oil painting indicator analysis: {油画指标分析_PATH}")

    df = pd.read_excel(油画指标分析_PATH, sheet_name=油画指标分析_SHEET, header=None)

    result = {
        'item_name': None,
        'bullet_points': [],
        'product_description': None,
        'generic_keyword': None,
    }

    try:
        result['item_name'] = str(df.iloc[32, 1]) if pd.notna(df.iloc[32, 1]) else ''
        for i in range(5):
            bp = str(df.iloc[33 + i, 1]) if pd.notna(df.iloc[33 + i, 1]) else ''
            result['bullet_points'].append(bp)
        result['product_description'] = str(df.iloc[38, 1]) if pd.notna(df.iloc[38, 1]) else ''
        result['generic_keyword'] = str(df.iloc[39, 1]) if pd.notna(df.iloc[39, 1]) else ''

        logger.info(f"Loaded item_name: {result['item_name'][:50] if result['item_name'] else 'None'}...")
        logger.info(f"Loaded {len(result['bullet_points'])} bullet points")
        logger.info(f"Loaded generic_keyword: {result['generic_keyword'][:50] if result['generic_keyword'] else 'None'}...")

    except Exception as e:
        logger.warning(f"Error reading oil painting analysis: {e}")

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
    油画指标分析: Dict[str, Any]
) -> None:
    """向模板副本填充数据"""
    logger.info("Filling template with data")

    ws = wb['Template_Copy']

    item_type_keyword = get_item_type_keyword_from_template(logger, wb)
    start_row = 9  # 从第9行开始填写，第8行及之前是带颜色的示例单元格
    item_name_template = 油画指标分析['item_name']

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

        _set_cell(ws, column_map, 'Item Length Longer Edge', target_row, 规格长, index=0)
        _set_cell(ws, column_map, 'Item Width Shorter Edge', target_row, 规格宽, index=0)
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

        # E. 包装尺寸重量列
        _set_cell(ws, column_map, 'Item Package Length', target_row, 规格长, index=0)
        _set_cell(ws, column_map, 'Item Package Width', target_row, 规格宽, index=0)
        _set_cell(ws, column_map, 'Item Package Height', target_row, row['规格高'], index=0)
        _set_cell(ws, column_map, 'Package Length Unit', target_row, FIXED_VALUES['package_length_unit'], index=0)
        _set_cell(ws, column_map, 'Package Width Unit', target_row, FIXED_VALUES['package_width_unit'], index=0)
        _set_cell(ws, column_map, 'Package Height Unit', target_row, FIXED_VALUES['package_height_unit'], index=0)
        _set_cell(ws, column_map, 'Package Weight', target_row, row['重量LB_净重'], index=0)
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
    logger = setup_logging()

    try:
        sku_df = read_sku申请表(logger)
        油画指标分析 = read_油画指标分析(logger)
        wb, column_map = load_and_prepare_template(logger)
        fill_template_with_data(logger, wb, column_map, sku_df, 油画指标分析)
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
