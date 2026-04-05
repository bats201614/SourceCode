import pandas as pd

cn = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二']

# 1. 售罄率与库容健康度 (Sell-through & IPI)
def audit_inventory_health(row):
    """计算当前总库存（含在途）是否能支撑销售规划"""
    try:
        # 提取并确保数值化（处理 NaN 情况）
        def get_val(key):
            val = row.get(key, 0)
            return val if pd.notna(val) and isinstance(val, (int, float)) else 0
        # 库存字段：现有库存 + 已下单未到货
        total_stock = (get_val('FBA库存') + get_val('FBA在途') + get_val('海外仓数量')
                        + get_val('集团库存数量') + get_val('海上在途数量') + get_val('中仓数量')
                        + get_val('PR未交量') + get_val('PO未交量'))

        five_month_plan = 0
        for i in range(5):
            col_name = f'第{cn[i]}个月销售规划'
            five_month_plan += get_val(col_name)

        if five_month_plan == 0:
            return "规划为0，无需审计"

        # 3. 覆盖月数计算
        avg_monthly_demand = five_month_plan / 5
        stock_coverage = total_stock / avg_monthly_demand

        # 4. 断货信息获取
        stockout_days = get_val('近7天断货天数')
        stockout_info = f"，近7天已断货{stockout_days}天" if stockout_days > 0 else ""

        # 5. 判定返回
        first_month_plan = get_val('第一个月销售规划')
        if total_stock < first_month_plan:
            return f"❌ 紧急：当前库存({total_stock:.0f})支撑不起第1月规划({first_month_plan:.0f})！{stockout_info}"
        
        if total_stock < five_month_plan:
            return f"⚠️ 远期断货：库存({total_stock:.0f})支撑不了5个月总量({five_month_plan:.0f}){stockout_info}"
        
        if stock_coverage > 12:
            return f"🧊 积压风险：库存过多，可撑 {stock_coverage:.1f} 个月{stockout_info}"
            
        return f"✅ 库存健康：可支撑 {stock_coverage:.1f} 个月{stockout_info}"

    except Exception as e:
        return f"库存审计异常：{str(e)}"

# 2. 去年同期比与季节性趋势 (Seasonality)
def audit_seasonality(row):
    """对比去年同期各月销量，判断规划是否过于激进"""
    try:
        # 取去年12个月销量
        last_year = [row.get(f'去年{cn[i]}月销量', 0) for i in range(12)]
        this_plan = [row.get(f'第{cn[i]}个月销售规划', 0) for i in range(12)]

        # 统计有数据的月份
        valid_months = [(i, ly, tp) for i, (ly, tp) in enumerate(zip(last_year, this_plan))
                        if not pd.isna(ly) and ly > 0]

        if not valid_months:
            return "新产品或去年无数据"

        # 计算各月增长率
        growth_rates = [(tp - ly) / ly for _, ly, tp in valid_months]
        avg_growth = sum(growth_rates) / len(growth_rates)
        max_growth = max(growth_rates)
        min_growth = min(growth_rates)

        # 判断：增长率超过50%为激进，低于-50%为保守
        if max_growth > 0.5:
            return f"增长激进：部分月份规划比去年同期增长超过50%（最高{max_growth:.0%}），需核实"
        elif min_growth < -0.5:
            return f"规划保守：部分月份规划比去年同期减少超过50%（最低{min_growth:.0%}），可适度上调"
        elif avg_growth > 0.3:
            return f"增长较快：平均比去年同期增长 {avg_growth:.0%}，需关注"
        elif avg_growth < -0.3:
            return f"整体下滑：平均比去年同期减少 {abs(avg_growth):.0%}，需分析原因"
        return f"符合季节性规律：平均同比增长 {avg_growth:.0%}"
    except Exception as e:
        return f"季节审计异常：{e}"

# 3. 价格与销量的联动关系 + 推广效率 (Price Elasticity & ROI)
def audit_price_logic(row):
    """检查提降价合理性及推广投入产出比"""
    try:
        price_diff = row.get('第一个月售价（修改前后差异比例）', 0)
        sales_plan = row.get('第一个月销售规划', 0)
        sales_speed = row.get('月化综合销售速度（剔除断货）', 0)
        promotion_fee = row.get('第一个月推广费(美元)', 0)
        planned_revenue = row.get('第一个月规划销售金额', 0)

        results = []

        # 价格-销量逻辑
        if pd.isna(price_diff) or pd.isna(sales_speed):
            results.append("数据缺失")
        elif price_diff > 0.05 and (sales_plan - sales_speed) > 0:
            results.append("逻辑矛盾：计划提价但销量规划却在增长")
        elif price_diff < -0.05 and (sales_plan - sales_speed) <= 0:
            results.append("建议：计划降价但销量规划未增长，可适度上调")

        # 推广效率：推广费 / 规划销售金额
        if not pd.isna(promotion_fee) and not pd.isna(planned_revenue) and planned_revenue > 0:
            roi = planned_revenue / promotion_fee if promotion_fee > 0 else float('inf')
            if promotion_fee > 0 and roi < 2.5:
                results.append(f"推广激进：${promotion_fee}推广费对应${planned_revenue:.0f}销售额，ROI={roi:.1f}")
            elif promotion_fee == 0 and planned_revenue > 0:
                results.append("零推广费高销售额，需确认是否自然流量")
            else:
                results.append(f"推广效率正常：ROI={roi:.1f}")

        return "; ".join(results) if results else "价格/推广逻辑正常"
    except Exception as e:
        return f"价格审计异常：{e}"

# 4. 链接定位与生命周期 (Life Cycle)
def audit_lifecycle_strategy(row):
    """根据链接定位、规划趋势、近期销量趋势综合判断"""
    try:
        tag = row.get('链接定位', '')
        plan_trend = [row.get(f'第{cn[i]}个月销售规划', 0) for i in range(5)]

        # 近期销量趋势
        recent_7d = row.get('近7天月化销售', 0)
        recent_30d = row.get('近30天月化销售', 0)
        recent_trend = ""
        if not pd.isna(recent_7d) and not pd.isna(recent_30d) and recent_30d > 0:
            if recent_7d > recent_30d * 1.2:
                recent_trend = "近期上升"
            elif recent_7d < recent_30d * 0.8:
                recent_trend = "近期下滑"

        plan_sum = sum(plan_trend)
        plan_start = plan_trend[0]
        plan_end = plan_trend[-1]

        # 规划趋势
        if plan_start > plan_end and plan_sum > 0:
            trend_desc = "下降"
        elif plan_start < plan_end:
            trend_desc = "上升"
        else:
            trend_desc = "持平"

        # 新链接判断
        if tag == "新链接":
            if plan_trend[0] >= plan_trend[-1]:
                return f"新链接警示：规划应呈阶梯上升，实际{trend_desc}，需核实"
            return f"新链接正常：规划趋势{trend_desc}{'，' + recent_trend if recent_trend else ''}"

        # 非保留链接（可调拨链接、C类 — 清货性质）
        if tag in ("可调拨链接", "C") or tag == "C":
            stock_fields = ['FBA库存', 'FBA在途', '海外仓数量', '集团库存数量', '海上在途数量', '中仓数量', 'PR未交量', 'PO未交量']
            total_stock = sum(row.get(f, 0) if pd.notna(row.get(f, 0)) else 0 for f in stock_fields)
            recent_7d_speed = row.get('近7天月化销售', 0) if pd.notna(row.get('近7天月化销售', 0)) else 0

            if plan_sum > total_stock:
                return f"清货异常：规划量({plan_sum}件)大于总库存({total_stock:.0f}件)，无法完成"
            if plan_sum < total_stock:
                # 流速判断：规划清货速度 vs 近期实际流速
                if recent_7d_speed > 0 and plan_sum < recent_7d_speed:
                    return f"清货过慢：规划月清{plan_sum}件，低于近期日均流速{recent_7d_speed}件，建议加大清货力度"
                elif recent_7d_speed > 0 and recent_7d_speed < 5:
                    return f"当前日流速过低({recent_7d_speed}件)，需核实是否需要降价刺激"
                return f"清货进行中：规划与库存匹配，库存{total_stock:.0f}件，规划{plan_sum}件"
            return f"清货进行中：规划与库存匹配"

        # 成熟链接判断
        if not pd.isna(recent_30d) and recent_30d > 0:
            # 规划 vs 近期销量对比
            if plan_start > recent_30d * 3:
                return f"异常：规划第1月销量远高于近期月均，需核实"
            elif plan_start < recent_30d * 0.3 and recent_30d > 10:
                return f"异常：规划第1月销量远低于近期月均，可能过于保守"

        return f"定位匹配：规划{trend_desc}{'，' + recent_trend if recent_trend else ''}"
    except Exception as e:
        return f"生命周期审计异常：{e}"

# --- 模拟数据运行测试 ---
if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding='utf-8')

    file_path = input("请输入测试数据文件路径：")
    sheet_name = input("请输入测试数据所在的Excel工作表名称：")
    output_path = input("请输入审计报告输出路径（Excel）：")
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # 收集审计结果
    results = []
    for idx, row in df.iterrows():
        results.append({
            "商品中文名称": row["商品中文名称"],
            "平台SKU": row.get("平台SKU", ""),
            "店铺": row.get("店铺", ""),
            "链接定位": row.get("链接定位", ""),
            "三级类目": row.get("三级类目", ""),
            "四级类目": row.get("四级类目", ""),
            "库存维度": audit_inventory_health(row),
            "季节维度": audit_seasonality(row),
            "价格维度": audit_price_logic(row),
            "生命周期": audit_lifecycle_strategy(row),
        })

    result_df = pd.DataFrame(results)
    result_df.to_excel(output_path, index=False, sheet_name="规划审计报告")
    print(f"审计报告已生成：{output_path}")
