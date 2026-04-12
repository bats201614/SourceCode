"""
物流尾程数据销量预测脚本
根据走货链接类型自动填入预测周销量
"""

from openpyxl import load_workbook
import os
import re
import pandas as pd
import logging
from datetime import datetime


def setup_logging(log_dir):
    """配置日志：同时输出到控制台和文件"""
    name_part = datetime.now().strftime("%m%d_%H%M%S")
    log_path = os.path.join(log_dir, f"forecast_{name_part}.log")
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return log_path


def parse_file_date(filename):
    """从文件名提取日期前缀，如'尾程-4.8.xlsx' -> (4, 8)"""
    basename = os.path.basename(filename)
    # 匹配 中文-数字.数字.扩展名 模式
    match = re.search(r'-(\d+)\.(\d+)', basename)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None, None


def unlock_columns(ws, col_indices):
    """解除指定列的保护锁定"""
    from openpyxl.styles import Protection
    unlocked = Protection(locked=False)
    # 使用iter_cols按列迭代，更高效
    min_col = min(col_indices)
    max_col = max(col_indices)
    for col in ws.iter_cols(min_col=min_col, max_col=max_col):
        if col[0].column in col_indices:
            for cell in col:
                cell.protection = unlocked


def find_col_by_keywords(headers_dict, keywords):
    """通过关键词模糊匹配找到列索引，未找到则抛出异常"""
    import re
    # 清理关键词中的非字母数字字符（非必要匹配字符）
    def normalize(s):
        return re.sub(r'[\s（）\(\)【】\[\]·�_]', '', s)

    keywords_norm = [normalize(kw) for kw in keywords]
    core_keywords = [kw for kw in keywords_norm if len(kw) >= 4]

    for header, col_idx in headers_dict.items():
        header_norm = normalize(header)
        if all(kw in header_norm for kw in keywords_norm):
            return col_idx

    for header, col_idx in headers_dict.items():
        header_norm = normalize(header)
        if all(kw in header_norm for kw in core_keywords):
            return col_idx

    raise ValueError(f"未找到列，关键词: {keywords}，请检查Excel模板是否变更")


def calculate_current_month_sales(monthly_30d_sales, current_day):
    """计算当月销量：当月1日至当前日期的销量"""
    return monthly_30d_sales / 30 * current_day


def calculate_weekly_forecast(row_data, weights=(0.4, 0.35, 0.25)):
    """
    计算周预测销量
    row_data: dict with keys: sales_7d, sales_30d, monthly_speed
    weights: (weight_7d, weight_30d, weight_speed)
    """
    w1, w2, w3 = weights
    sales_7d = row_data.get('sales_7d', 0) or 0
    sales_30d = row_data.get('sales_30d', 0) or 0
    monthly_speed = row_data.get('monthly_speed', 0) or 0

    # 转换为周销量
    weekly_from_7d = sales_7d * 7 / 30
    weekly_from_30d = sales_30d / 4  # 月化销量除以4得到周销量

    forecast = w1 * weekly_from_7d + w2 * weekly_from_30d + w3 * monthly_speed / 4
    return round(forecast, 2)


def analyze_trend(sales_7d, sales_14d, sales_30d, sales_60d, sales_90d):
    """
    分析趋势类型，基于各期销量比值
    返回: (trend_type, trend_factor)
    trend_type: 'growth', 'decline', 'stable', 'ramping_up', 'ramping_down', 'out_of_stock', 'recovery'
    trend_factor: 每递增一周的调整系数
    """
    # 计算各期与30天的比值
    r7 = sales_7d / sales_30d if sales_30d > 0 else 1.0
    r14 = sales_14d / sales_30d if sales_30d > 0 else 1.0
    r60 = sales_60d / sales_30d if sales_30d > 0 else 1.0
    r90 = sales_90d / sales_30d if sales_30d > 0 else 1.0

    # ========== 1. 断档识别 (Out of Stock) ==========
    # 特征：30天销量极低（可能是缺货），但60/90天销量正常或偏高
    # 说明：断档期间无法销售，导致30天数据失真
    if sales_30d > 0 and r7 < 0.3 and r14 < 0.5 and (r60 > 1.0 or r90 > 1.0):
        # 近7天和14天都极低，但长期数据正常 → 断档恢复期
        if r7 > r14 > 0.5:  # 正在恢复
            return 'recovery', 1.10  # 断档恢复中，销量回升
        elif r7 < 0.1 and r14 < 0.2:  # 仍在断档
            return 'out_of_stock', 1.0  # 仍处于断档状态
        else:
            return 'out_of_stock', 1.05  # 断档后过渡期

    # ========== 2. 持续增长 ==========
    if r7 > 1.1 and r14 > 1.0 and r60 > 1.0 and r90 > 1.0:
        return 'growth', 1.10  # 每周增长10%

    # ========== 3. 持续下降/衰退 ==========
    # 特征：各期比值都 < 1，且呈现持续下滑趋势
    if r7 < 0.9 and r14 < 1.0 and r60 < 1.0 and r90 < 1.0:
        if r7 < r14 < r60 < r90:  # 持续下滑
            return 'decline', 0.90  # 每周下降10%
        elif r90 < r60 < r14 < r7:  # 曾经高现在低，下降趋稳
            return 'ramping_down', 0.95  # 每周下降5%后趋稳

    # ========== 4. 衰退 vs 季节性区分 ==========
    # 关键判断：60/90天高，但7/14天低 → 可能是衰退而非季节性
    if r60 > 1.1 or r90 > 1.2:
        # 近7天和14天都在正常范围 → 季节性波动
        if 0.7 <= r7 <= 1.3 and 0.6 <= r14 <= 1.3:
            return 'seasonal', 1.06  # 季节性因素，每周6%
        # 近7天或14天极低 → 衰退趋势（长期数据可能是历史高峰）
        elif r7 < 0.7 or r14 < 0.6:
            return 'decline', 0.92  # 衰退，每周下降8%

    # ========== 5. 增长放缓 ==========
    # 特征：7天仍然高于30天，但14天已低于30天，且7天增速低于14天（r7 < r14）
    # 说明：增长势头在减缓
    if r7 > 1.0 and r7 < r14 and r14 < 1.0:
        return 'growth_slowing', 1.05  # 增长放缓，每周5%

    # ========== 6. 下降放缓（回升中） ==========
    # 特征：7天最低，但长期数据显示曾经更高
    if r7 < 0.9 and r90 > 1.0:
        return 'decline_slowing', 1.02  # 下降放缓，逐步回升2%

    # ========== 7. 新品爬坡 ==========
    # 特征：7天和14天接近但都低于长期均值
    if 0.8 <= r7 <= 1.1 and 0.7 <= r14 <= 1.0 and r90 < 0.8:
        return 'ramping_up', 1.08  # 爬坡中，每周增长8%

    # ========== 8. 波动持平 ==========
    if 0.8 <= r7 <= 1.2 and 0.7 <= r14 <= 1.2:
        return 'stable', 1.0  # 保持稳定

    # 默认：平稳
    return 'stable', 1.0


def generate_trend_weekly_forecast(base_weekly, sales_7d, sales_14d, sales_30d, sales_60d, sales_90d,
                                   fba_transit_inventory=0, fba_transit_days=0, replenishment_days=0, num_weeks=8):
    """
    生成8周趋势预测值
    基于各期销量综合分析趋势类型
    """
    trend_type, trend_factor = analyze_trend(sales_7d, sales_14d, sales_30d, sales_60d, sales_90d)

    weekly_forecasts = []
    for week in range(1, num_weeks + 1):
        if week == 1:
            weekly_forecasts.append(int(round(base_weekly)))
        else:
            # 趋势调整
            if trend_type in ['growth', 'ramping_up', 'seasonal']:
                adjusted = base_weekly * (trend_factor ** (week - 1))
            elif trend_type in ['decline', 'ramping_down']:
                # 下降趋势有下限，不能低于基准的50%
                adjusted = base_weekly * max(0.5, trend_factor ** (week - 1))
            elif trend_type == 'growth_slowing':
                # 增长放缓，前几周增长，后期稳定
                if week <= 4:
                    adjusted = base_weekly * (trend_factor ** (week - 1))
                else:
                    adjusted = base_weekly * (trend_factor ** 4)
            elif trend_type == 'decline_slowing':
                # 下降放缓，逐步回升
                adjusted = base_weekly * (1 + (trend_factor - 1) * (week - 1) / 4)
            elif trend_type == 'recovery':
                # 断档恢复期：前3周快速回升，之后回归稳定
                if week <= 3:
                    adjusted = base_weekly * (trend_factor ** (week - 1))
                else:
                    adjusted = base_weekly * (trend_factor ** 3)  # 第4周起保持第3周水平
            elif trend_type == 'out_of_stock':
                # 断档状态，预测为0或极低
                adjusted = 0
            else:  # stable
                adjusted = base_weekly

            weekly_forecasts.append(int(round(adjusted)))

    # ========== 库存分析建议（不修改预测值） ==========
    total_demand = sum(weekly_forecasts)
    avg_weekly = total_demand / num_weeks if num_weeks > 0 else 0
    weeks_supported = fba_transit_inventory / avg_weekly if avg_weekly > 0 else float('inf')

    inventory_suggestion = ''
    if fba_transit_days > 0 and replenishment_days > 0:
        diff = fba_transit_days - replenishment_days
        if diff < 0:
            inventory_suggestion = f'⚠️ 库存{weeks_supported:.1f}周 < 补货周期{replenishment_days}天，建议补货'
        elif diff < 7:
            inventory_suggestion = f'⚡ 库存{weeks_supported:.1f}周，快到补货点'
        else:
            inventory_suggestion = f'✓ 库存{weeks_supported:.1f}周充足'

    return weekly_forecasts, trend_type, inventory_suggestion


def apply_outage_correction(forecast, stockout_days, period_days=7):
    """断货校正：如果有断货，预测值可能偏低"""
    if stockout_days and stockout_days > 0:
        correction_factor = 1 + (stockout_days / period_days)
        forecast = forecast * min(correction_factor, 1.5)
    return forecast


def main():
    # 询问用户文件路径
    file_path = input("请输入尾程Excel文件完整路径：").strip().strip('"')

    if not os.path.exists(file_path):
        print(f"文件不存在: {file_path}")
        return

    # 初始化日志（输出目录/尾程处理结果/）
    name_part, ext_part = os.path.splitext(file_path)
    output_dir = name_part + "-已完成自动填入"
    result_dir = os.path.join(output_dir, "尾程处理结果")
    os.makedirs(result_dir, exist_ok=True)
    log_path = setup_logging(result_dir)
    logging.info(f"日志文件: {log_path}")
    logging.info("=" * 60)
    logging.info(f"开始处理文件: {file_path}")

    # 解析文件日期
    month, day = parse_file_date(file_path)
    if month is None:
        logging.error("无法从文件名解析日期，请确保文件名格式如'尾程-4.8.xlsx'")
        return
    logging.info(f"解析到日期：{month}月{day}日")

    # 加载工作簿（data_only=True读取计算值，False读取公式用于写入）
    wb_data = load_workbook(file_path, data_only=True)
    ws_data = wb_data.active

    wb = load_workbook(file_path)
    ws = wb.active

    # 定义需要解除锁定的列索引（1-based for openpyxl）
    # 根据之前分析: 列28,29月度规划; 列38-45预测; 列82,84销量; 列92库存; 列102当月销量
    locked_cols_to_unlock = [29, 30, 39, 40, 41, 42, 43, 44, 45, 46, 83, 85, 93, 103, 104]
    logging.info(f"解除{len(locked_cols_to_unlock)}列的保护锁定...")
    unlock_columns(ws, locked_cols_to_unlock)

    # 读取列名（第4行）建立索引映射
    headers = {}
    for col in range(1, ws.max_column + 1):
        header = ws.cell(row=4, column=col).value
        if header:
            # 清理换行符为空格，便于匹配
            header_clean = str(header).replace('\n', ' ').strip()
            headers[header_clean] = col

    logging.info(f"共识别{len(headers)}个列名")

    # 使用关键词模糊匹配关键列
    COL_LINK_TYPE = find_col_by_keywords(headers, ['走货链接类型']) or 26
    COL_SALES_7D = find_col_by_keywords(headers, ['近7天', '月化销量']) or 83
    COL_SALES_14D = find_col_by_keywords(headers, ['近14天', '月化销量']) or 84
    COL_SALES_30D = find_col_by_keywords(headers, ['近30天', '月化销量']) or 85
    COL_SALES_60D = find_col_by_keywords(headers, ['近60天', '月化销量']) or 86
    COL_SALES_90D = find_col_by_keywords(headers, ['近90天', '月化销量']) or 87
    COL_STOCKOUT_7D = find_col_by_keywords(headers, ['近7天', '断货']) or 88
    COL_STOCKOUT_14D = find_col_by_keywords(headers, ['近14天', '断货']) or 90
    COL_STOCKOUT_30D = find_col_by_keywords(headers, ['近30天', '断货']) or 92
    COL_MONTHLY_SPEED = find_col_by_keywords(headers, ['月综合销售速度', '剔除近期断货']) or 82
    COL_FBA_DAYS = find_col_by_keywords(headers, ['FBA', '在仓可售天数']) or 94
    COL_FBA_TRANSIT = find_col_by_keywords(headers, ['FBA', '在仓+在途库存']) or 95
    COL_FBA_TRANSIT_DAYS = find_col_by_keywords(headers, ['FBA在仓+FBA在途可售天数']) or 96
    COL_REPLENISHMENT_DAYS = find_col_by_keywords(headers, ['尾程补货天数']) or 77
    COL_CURRENT_MONTH_SALES = find_col_by_keywords(headers, ['当月销量']) or 103
    COL_FORECAST_WEEK1 = find_col_by_keywords(headers, ['店长预测', '第1周']) or 39
    COL_FORECAST_WEEK2 = find_col_by_keywords(headers, ['店长预测', '第2周']) or 40
    COL_FORECAST_WEEK3 = find_col_by_keywords(headers, ['店长预测', '第3周']) or 41
    COL_FORECAST_WEEK4 = find_col_by_keywords(headers, ['店长预测', '第4周']) or 42
    COL_FORECAST_WEEK5 = find_col_by_keywords(headers, ['店长预测', '第5周']) or 43
    COL_FORECAST_WEEK6 = find_col_by_keywords(headers, ['店长预测', '第6周']) or 44
    COL_FORECAST_WEEK7 = find_col_by_keywords(headers, ['店长预测', '第7周']) or 45
    COL_FORECAST_WEEK8 = find_col_by_keywords(headers, ['店长预测', '第8周']) or 46
    COL_FORECAST_SUM = find_col_by_keywords(headers, ['未来8周预测销量汇总']) or 47

    logging.info(f"计算当月销量（{month}月1日至{day}日）...")

    # 遍历数据行（第5行开始）
    processed_count = 0
    for row in range(5, ws.max_row + 1):
        asin = ws.cell(row=row, column=headers.get('ASIN', 22)).value
        link_type = ws.cell(row=row, column=COL_LINK_TYPE).value

        if not asin or asin == 'ASIN':
            continue

        logging.info(f"处理 ASIN: {asin}, 链接类型: {link_type}")

        # 获取关键数据（从data_only读取计算值）
        sales_7d = ws_data.cell(row=row, column=COL_SALES_7D).value or 0
        sales_14d = ws_data.cell(row=row, column=COL_SALES_14D).value or 0
        sales_30d = ws_data.cell(row=row, column=COL_SALES_30D).value or 0
        sales_60d = ws_data.cell(row=row, column=COL_SALES_60D).value or 0
        sales_90d = ws_data.cell(row=row, column=COL_SALES_90D).value or 0
        monthly_speed = ws_data.cell(row=row, column=COL_MONTHLY_SPEED).value or 0
        fba_days = ws_data.cell(row=row, column=COL_FBA_DAYS).value or 0
        fba_transit = ws_data.cell(row=row, column=COL_FBA_TRANSIT).value or 0
        fba_transit_days = ws_data.cell(row=row, column=COL_FBA_TRANSIT_DAYS).value or 0
        replenishment_days = ws_data.cell(row=row, column=COL_REPLENISHMENT_DAYS).value or 0

        stockout_7d = ws_data.cell(row=row, column=COL_STOCKOUT_7D).value or 0
        stockout_14d = ws_data.cell(row=row, column=COL_STOCKOUT_14D).value or 0
        stockout_30d = ws_data.cell(row=row, column=COL_STOCKOUT_30D).value or 0

        # Step 2a: 当月销量保留原值（不覆盖）
        original_current_month = ws_data.cell(row=row, column=COL_CURRENT_MONTH_SALES).value
        logging.info(f"  当月销量: 保留原值 {original_current_month}")

        # Step 3: 根据链接类型填入预测周销量
        forecast_cols = [
            COL_FORECAST_WEEK1, COL_FORECAST_WEEK2, COL_FORECAST_WEEK3,
            COL_FORECAST_WEEK4, COL_FORECAST_WEEK5, COL_FORECAST_WEEK6,
            COL_FORECAST_WEEK7, COL_FORECAST_WEEK8
        ]

        # 清货/停售类型：8周预测全填0
        zero_link_types = {'C', '可调拨链接', '停售链接'}
        if link_type in zero_link_types:
            logging.info(f"  [清货/停售产品] 填入8个0")
            for col in forecast_cols:
                ws.cell(row=row, column=col).value = 0
            forecast_sum = 0
        else:
            # 链接类型系数映射
            link_type_multipliers = {'A': 1.2, 'N': 0.8}
            multiplier = link_type_multipliers.get(link_type, 1.0)
            link_type_names = {'A': '拉升链接', 'N': '新品链接'}
            type_name = link_type_names.get(link_type, '其他类型')
            logging.info(f"  [{type_name}] 使用模型计算 x {multiplier}")

            # 计算基础预测
            row_data = {
                'sales_7d': sales_7d,
                'sales_30d': sales_30d,
                'monthly_speed': monthly_speed
            }
            max_stockout = max(stockout_7d, stockout_14d, stockout_30d)

            weekly_base = calculate_weekly_forecast(row_data)
            weekly_base = apply_outage_correction(weekly_base, max_stockout)
            weekly_base = int(round(weekly_base * multiplier))

            # 生成8周趋势预测
            weekly_forecasts, trend_type, inventory_suggestion = generate_trend_weekly_forecast(
                weekly_base, sales_7d, sales_14d, sales_30d, sales_60d, sales_90d,
                fba_transit, fba_transit_days, replenishment_days
            )
            logging.info(f"  基础预测: {weekly_base}, 趋势类型: {trend_type}")
            logging.info(f"  8周预测: {weekly_forecasts}")
            if inventory_suggestion:
                logging.info(f"  库存建议: {inventory_suggestion}")

            for i, col in enumerate(forecast_cols):
                ws.cell(row=row, column=col).value = weekly_forecasts[i]
            forecast_sum = sum(weekly_forecasts)

        # Step 4: 填入8周预测汇总
        ws.cell(row=row, column=COL_FORECAST_SUM).value = round(forecast_sum, 2)
        logging.info(f"  8周预测汇总: {forecast_sum:.2f}")

        processed_count += 1

    logging.info(f"处理完成！共处理 {processed_count} 个产品")

    # 保存到新文件（原文件名 + "-已完成自动填入"）
    output_path = output_dir + ext_part
    wb.save(output_path)
    logging.info(f"文件已保存: {output_path}")
    logging.info("=" * 60)
    print(f"\n日志文件: {log_path}")


if __name__ == "__main__":
    main()
