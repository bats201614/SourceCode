import os
import time
import random
import pandas as pd
from DrissionPage import ChromiumPage, ChromiumOptions

def setup_edge_browser():
    """
    配置并启动 Chrome 浏览器
    """
    print("正在初始化 Chrome 浏览器配置...")
    co = ChromiumOptions()

    # 设置 Edge 浏览器的执行路径（这里以 Windows 默认路径为例）
    # 较新版本的 DrissionPage 直接写 'edge' 即可自动寻找
    # edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    co.set_browser_path('chrome')
    #配合隐藏自动化特征
    co.set_argument('--disable-blink-features=AutomationControlled')
    # 设置UA
    my_ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36"
    co.set_user_agent(my_ua)
    
    # 启动浏览器并返回页面对象
    page = ChromiumPage(co)
    return page

def scrape_amazon_item(page, url, asin):
    print(f"\n正在访问: {url}")
    page.get(url)
    # 等待页面开始加载（确保浏览器已经响应了跳转指令）
    page.wait.load_start()
    # 随机停顿 3 到 6 秒，模拟人类浏览，防止被亚马逊风控
    time.sleep(random.uniform(3, 6))
    
    # 准备一个字典来存放提取的数据
    data = {
        'ASIN': asin,
        '链接': url,
        '商品标题': '',
        '当前价格': '',
        '五点描述': ''
    }

    # 1. 抓取标题
    title_ele = page.ele('#productTitle')
    if title_ele:
        data['商品标题'] = title_ele.text.strip()
        print(f"成功获取标题")
    else:
        print("警告：未能定位到标题元素")

    # 2. 抓取价格 (亚马逊价格分整数和小数两部分)
    price_whole = page.ele('.a-price-whole')
    price_fraction = page.ele('.a-price-fraction')
    if price_whole and price_fraction:
        # 去掉整数部分可能带的换行符和逗号
        whole = price_whole.text.replace('\n', '').replace('.', '')
        data['当前价格'] = f"${whole}.{price_fraction.text}"
        print(f"成功获取价格")
    else:
        print("警告：未能定位到价格")

    # 3. 抓取五点描述 (Bullet Points)
    bullets_box = page.ele('#feature-bullets')
    if bullets_box:
        bullet_items = bullets_box.eles('t:li')
        bullets_list =[]
        for index, item in enumerate(bullet_items, start=1):
            text = item.text.strip()
            # 过滤掉亚马逊页面上一些杂乱的隐藏提示（如售后等）
            if text and not text.startswith("Make sure this fits"):
                bullets_list.append(f"{index}. {text}")
                
        # 将五点描述用“换行符”连接起来，这样在Excel里它们会待在同一个单元格且分行显示
        data['五点描述'] = '\n'.join(bullets_list)
        print(f"成功获取五点描述")
    else:
        print("警告：未能定位到五点描述")

    return data

if __name__ == '__main__':
    # ================= 1. 准备你的竞品链接 =================
    input_file = r"D:\文件\14 工作文件\广州市速鸽科技有限公司\竞品ASIN清单.xlsx"
    try:
        df_input = pd.read_excel(input_file)
        # 将 ASIN 列转为列表，并去除可能的空格
        asin_list = df_input['ASIN'].dropna().astype(str).str.strip().tolist()
        print(f"成功导入 {len(asin_list)} 个 ASIN。")
    except Exception as e:
        print(f"读取 Excel 出错，请检查路径或列名: {e}")
        asin_list = []
    product_urls =[f"https://www.amazon.com/dp/{asin}" for asin in asin_list]

    # ================= 2. 启动浏览器并手动预热 =================
    print("正在启动 Chrome 浏览器...")
    page = setup_edge_browser()

    # 先打开亚马逊首页
    page.get("https://www.amazon.com")

    print("\n" + "="*50)
    print("【重要提示】")
    print("浏览器已打开。请在浏览器中：")
    print("1. 如果有验证码，请手动点击通过。")
    print("2. 点击左上角将配送地址 (Deliver to) 的邮编修改为美国邮编（如 10001）。")
    print("修改完毕后，请回到这个黑色窗口。")
    print("="*50 + "\n")
    
    input("👉 修改完邮编后，请按【回车键 (Enter)】开始批量抓取并写入 Excel：")

    # ================= 3. 开始自动批量抓取 =================
    all_data =[] # 用来存放所有商品的数据

    for asin, url in zip(asin_list, product_urls):
        try:
            item_data = scrape_amazon_item(page, url, asin)
            all_data.append(item_data)
            time.sleep(random.uniform(2, 5))
            print("抓取成功")
        except Exception as e:
            print(f"【报错】抓取 {url} 时出错: {e}")

    # ================= 4. 保存到 Excel =================
    if all_data:
        print("\n正在生成 Excel 文件...")
        # 将数据列表转换为 pandas 数据框
        df = pd.DataFrame(all_data)
        
        # 定义保存的文件名
        # 1. 基础路径
        base_path = r"D:\文件\14 工作文件\广州市速鸽科技有限公司"
        # 2. 动态生成今天的日期文件夹名 (例如: 2026-03-01)
        today_folder = time.strftime('%Y-%m-%d')
        # 3. 拼接完整的存放路径
        target_path = os.path.join(base_path, today_folder)
        # 4. 自动创建文件夹 (exist_ok=True 表示如果文件夹已存在，直接跳过不报错)
        os.makedirs(target_path, exist_ok=True)
        # 5. 定义文件名
        file_name = f"竞品数据抓取_{time.strftime('%H时%M分')}.xlsx"
        excel_filename = os.path.join(target_path, file_name)
        
        # 写入 Excel，index=False 表示不写入行号 0, 1, 2...
        df.to_excel(excel_filename, index=False)
        print(f"✅ 任务大功告成！数据已保存到当前文件夹：{excel_filename}")
    else:
        print("未能抓取到任何数据，没有生成 Excel。")
        
    # 结束任务，关闭浏览器
    page.quit()