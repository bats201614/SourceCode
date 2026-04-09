import os
import time
import re
import pandas as pd
from DrissionPage import ChromiumPage, ChromiumOptions

def setup_browser():
    """初始化浏览器"""
    co = ChromiumOptions()
    # 如果你的 Chrome 路径不同，请修改这里
    browser_path = r"D:\software\dailysoftware\Chrome\chrome.exe" 
    co.set_browser_path(browser_path)
    return ChromiumPage(co)

def get_multiple_pages_asins(page, target_url, target_count):
    asins = set()
    page.get(target_url)
    
    # 针对亚马逊产品块的更精准定位（通常包含在 [data-asin] 属性中）
    # 或者针对排行榜的特定结构
    
    while len(asins) < target_count:
        print("🔍 正在扫描页面...")
        page.scroll.to_half() # 先滚一半触发懒加载
        time.sleep(1)
        page.scroll.to_bottom()
        time.sleep(2)

        # 优化方案：直接寻找带有 ASIN 特征的链接，减少无效遍历
        # 亚马逊产品链接通常包含 /dp/XXXXXXXXXX
        items = page.eles('xpath://a[contains(@href, "/dp/")]')
        
        for item in items:
            href = item.attr('href')
            match = re.search(r'/dp/([A-Z0-9]{10})', href)
            if match:
                asin = match.group(1)
                if asin not in asins:
                    asins.add(asin)
                    if len(asins) >= target_count:
                        return list(asins)
        
        print(f"📊 当前数量: {len(asins)}")

        # 改进的翻页逻辑
        next_btn = page.ele('@@tag:a@@text():Next', timeout=3) or \
                   page.ele('.s-pagination-next', timeout=1) or \
                   page.ele('.a-last', timeout=1) # 排行榜常用的翻页类
        
        if next_btn:
            # 检查类名中是否包含“disabled”，亚马逊通常用这个来禁用按钮
            is_disabled = 'disabled' in (next_btn.attr('class') or '')
            
            if not is_disabled:
                print("移动至下一页...")
                next_btn.click()
                page.wait.load_start()
                time.sleep(4)
            else:
                print("⚠️ 已到达最后一个有效页面（按钮已禁用）。")
                break
        else:
            print("⚠️ 未找到下一页按钮，停止采集。")
            break

if __name__ == '__main__':
    # --- 用户配置区 ---
    search_url = "https://www.amazon.com/gp/bestsellers/kitchen/3744211/ref=pd_zg_hrsr_kitchen" # 你想搜的关键词页面
    try:
        need_num = int(input("请输入你需要抓取的 ASIN 数量 (例如 50): "))
    except ValueError:
        print("请输入有效的数字！")
        exit()

    # --- 执行区 ---
    page = setup_browser()
    page.get("https://www.amazon.com")
    
    try:
        print("="*50)
        print("请在浏览器中确认邮编（如 10001）并处理验证码。")
        input("👉 确认无误后，请在此按【回车键】开始执行：")

        final_asins = get_multiple_pages_asins(page, search_url, need_num)

        # 保存结果
        if final_asins:
            df = pd.DataFrame({'ASIN': final_asins})
            output_file = "竞品ASIN清单.xlsx"
            df.to_excel(output_file, index=False)
            print("-" * 50)
            print(f"✅ 采集完成！实际抓取数量: {len(final_asins)}")
            print(f"📁 文件已保存至: {os.path.abspath(output_file)}")
        else:
            print("❌ 未提取到任何数据。")

    finally:
        page.quit()