import numpy as np
import pandas as pd

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
        # 4. 加入加速/衰减因子 (Acceleration/Deceleration factor)
        acceleration = 1.0
        if self.v7 > self.v14 > self.v30:
            acceleration = 1.2  # 增长加速奖励 20%
        elif self.v7 < self.v14 < self.v30:
            drop_rate = (self.v30 - self.v7) / self.v30 if self.v30 > 0 else 0
            acceleration = max(0.6, 1.0 - drop_rate)  # 销量下降衰减，最多衰减40%
        #5. 稳定性调整 (Stability adjustment)
        v_series = [self.v7, self.v14, self.v30, self.v45, self.v60, self.v90]
        std_dev = np.std(v_series)
        mean_v = np.mean(v_series)
        cv = std_dev / mean_v if mean_v > 0 else 0
        # 波动率越小，稳定性加成越高 (范围在 1.0 到 1.05 之间)
        stability_bonus = 1 + (0.05 * (1 - min(cv, 1)))
        # 最终动能 = 基础比率 * 加速度 * 稳定性
        final_momentum = momentum_ratio * acceleration * stability_bonus
        
        return final_momentum

    def predict_5_months(self):
        results = []
        # 初始权重分配
        weights = {'internal': 0.2, 'comp': 0.5, 'external': 0.3}
        
        current_v = self.v30
        for m in range(1, 6):
            # 1. 获取竞品该月基准
            comp_benchmark = np.mean([c[m-1] for c in self.comp_data])
            
            # 2. 模拟外部环境波动 (市场季节性，可以加入如3-5个核心大词关键词在亚马逊站内的月度搜索总量、类目 Top 100 总销量、周搜索趋势等数据来调整这个因子)
            keyword_search_trend = 500000/460000
            Top100_sales_trend = 200000/180000
            Growth_trend = 1 + (20000-18000)/18000
            market_factor = 0.5 * keyword_search_trend + 0.3 * Top100_sales_trend + 0.2 * Growth_trend 
            
            # 3. 核心公式：下个月预测 = (内部动能 * 权重 + 竞品基准 * 权重) * 市场因子
            momentum = self.calculate_internal_momentum()
            
            # 随着月份增加，内部权重提升，竞品参考权重下降（因为你有了自己的历史）
            if m > 3:
                weights['internal'] = 0.5
                weights['comp'] = 0.2
                weights['external'] = 0.3
            
            next_v = (current_v * momentum * weights['internal'] + 
                      comp_benchmark * weights['comp']) * market_factor
            
            results.append(round(next_v, 2))
            current_v = next_v
            
        return results

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
prediction = sim.predict_5_months()

print(f"基于当前 momentum，未来5个月销量预估为: {prediction}")