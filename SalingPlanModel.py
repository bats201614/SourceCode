import numpy as np

class AmazonLaunchSim:
    def __init__(self, v7, v14, v30, v45, v60, v90, comp_data):
        self.v7 = v7
        self.v14 = v14
        self.v30 = v30
        self.v45 = v45
        self.v60 = v60
        self.v90 = v90
        self.comp_data = comp_data # 5个竞品前5个月的数据矩阵
        
    def calculate_internal_momentum(self):
        # 计算内部动能：近期速度对比中期速度的增幅
        # 1. 定义短期权重 (Short-term focus)
        short_term = (self.v7 * 0.7) + (self.v14 * 0.2) + (self.v30 * 0.1)
        # 2. 定义长期基准 (Long-term baseline)
        # 注意：如果是新品，v60/v90 可能为 0，需要做平滑处理 (+1)
        long_term = (self.v30 * 0.4) + (self.v45 * 0.3) + (self.v60 * 0.2) + (self.v90 * 0.1)
        if long_term < 1:
            long_term = 1.2
        # 3. 计算动能比率
        momentum_ratio = short_term / long_term
        # 4. 加速/衰减因子（互斥条件：加速 or 减速 or 平稳）
        acceleration = 1.0
        if self.v7 > self.v14 and self.v14 > self.v30:
            acceleration = 1.2  # 连续加速增长：奖励20%
        elif self.v7 < self.v14 and self.v14 < self.v30:
            # 连续减速：用v14到v7的跌幅为基准，更准确
            drop_rate = (self.v14 - self.v7) / self.v14 if self.v14 > 0 else 0
            acceleration = max(0.6, 1.0 - drop_rate)  # 销量下降衰减，最多衰减40%
        # 5. 最终动能 = 基础比率 × 加速度因子
        final_momentum = momentum_ratio * acceleration
        
        return final_momentum

    def predict_5_months(self):
        """
        重构公式：三个因子同一层级，动能主导，不指数叠加
        公式：effective_growth = w_m×(m-1) + w_c×(c-1) + w_k×(k-1)
        其中 m=动能，c=竞品增速，k=市场因子
        """
        results = []

        current_v = self.v30
        initial_momentum = self.calculate_internal_momentum()

        # 权重：动能70%，竞品20%，市场20%
        w_m, w_c, w_k = 0.7, 0.2, 0.2

        for m in range(1, 6):
            # --- 新增：动能随时间衰减逻辑 ---
            # m=1时，decay=1.0(全效)；m=5时，decay=0.4
            decay_factor = 1.0 - (m - 1) * 0.15 
            # 计算当前月份的实际动能，使用 max(0.2, ...) 防止衰减成负数
            # 这意味着动能会随着时间逐渐回归到 1.0 (不增不减的状态)
            current_momentum = 1 + (initial_momentum - 1) * max(0.2, decay_factor)
            # 1. 竞品该月基准：过滤零值，中位数抗异常
            month_data = [c[m - 1] for c in self.comp_data]
            valid_data = [d for d in month_data if d > 0]
            comp_benchmark = np.median(valid_data) if valid_data else current_v
            comp_rate = comp_benchmark / current_v if current_v > 0 else 1.0

            # 2. 市场环境因子（占位参数）
            keyword_search_trend = 500000 / 460000   # 占位
            Top100_sales_trend = 200000 / 180000    # 占位
            Growth_trend = 1 + (20000 - 18000) / 18000  # 占位
            market_rate = 0.5 * keyword_search_trend + 0.3 * Top100_sales_trend + 0.2 * Growth_trend

            # 3. 统一公式：三个因子同一层级，动能决定增长上限（无指数叠加）
            if current_momentum >= 1:
                # 动能为增长提供基础，上限为动能本身（防止竞品+市场过度放大）
                base_growth = w_m * (current_momentum - 1)
                comp_growth = w_c * (comp_rate - 1)
                market_growth = w_k * (market_rate - 1)
                # 限制本月增长率不超过当前动能上限
                effective_growth = min(base_growth + comp_growth + market_growth, current_momentum - 1)
            else:
                market_growth = w_k * 0.2 * (market_rate - 1)
                effective_growth = current_momentum * (1 + market_growth) - 1

            next_v = current_v * (1 + effective_growth)
            results.append(round(next_v, 2))
            current_v = next_v

        return results, initial_momentum

# 假设竞品数据 (5个竞品，前5个月的销量)
comp_samples = [
    [50, 150, 400, 800, 1200],
    [40, 120, 350, 700, 1100],
    [60, 200, 500, 900, 1500],
    [45, 130, 380, 750, 1000],
    [55, 180, 450, 850, 1300]
]

# 执行模拟
sim = AmazonLaunchSim(v7=39, v14=45, v30=46, v45=37, v60=24, v90=29, comp_data=comp_samples)
prediction, momentum = sim.predict_5_months()

print(f"当前内部动能（Momentum）: {round(momentum, 3)}")
print(f"基于当前 Momentum，未来5个月销量预估为: {prediction}")