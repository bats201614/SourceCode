import os
import time
import random
import re
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from DrissionPage import ChromiumPage, ChromiumOptions

def setup_edge_browser():
    """
    配置并启动 Chrome 浏览器
    """
    print("正在初始化 Chrome 浏览器配置...")
    co = ChromiumOptions()
    # 设置UA
    my_ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36"
    co.set_user_agent(my_ua)

    # 设置浏览器的执行路径（这里以 Windows 默认路径为例）
    # --- 策略 1: 尝试让系统自己找 ---
    # 尝试顺序：Chrome -> Edge
    success = False
    for browser in ['chrome', 'edge']:
        try:
            co.set_browser_path(browser)
            # 尝试启动一个空页面测试路径是否有效
            test_page = ChromiumPage(co)
            test_page.quit()
            success = True
            print(f"✅ 已自动定位到浏览器: {browser}")
            break
        except Exception:
            continue

    # --- 策略 2: 如果还是找不到，弹窗让用户指路 ---
    if not success:
        print("⚠️ 未能在默认路径找到浏览器，请在弹出的窗口中手动选择浏览器执行文件...")
        # 隐藏 tkinter 主窗口
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True) # 让弹窗显示在最前面
        
        # 弹出文件选择对话框
        manual_path = filedialog.askopenfilename(
            title="请选择您的浏览器执行文件 (chrome.exe 或 msedge.exe)",
            filetypes=[("Executable files", "*.exe")],
            initialdir="C:/Program Files"
        )
        
        if manual_path:
            co.set_browser_path(manual_path)
            print(f"✅ 已手动加载浏览器路径: {manual_path}")
        else:
            raise FileNotFoundError("❌ 用户未选择浏览器路径，程序无法运行。")
            
    return ChromiumPage(co)

SIZE_PATTERN = re.compile(r'(\d+\.?\d*)\s*[xX*×-]\s*(\d+\.?\d*)(?:\s*(?:inch|inches|cm|mm|")\b)?')

def extract_size_from_text(text):
    if not text:
        return None
    match = SIZE_PATTERN.search(text)
    if match:
        return f"{match.group(1)} x {match.group(2)}"
    return None

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
        '尺寸/规格': '',
        '五点描述': ''
    }

    # 1. 抓取标题
    title_ele = page.ele('#productTitle')
    if title_ele:
        data['商品标题'] = title_ele.text.strip()
        print("成功获取标题")
    else:
        print("警告：未能定位到标题元素")

    # 2. 抓取价格 (亚马逊价格分整数和小数两部分)
    price_whole = page.ele('.a-price-whole')
    price_fraction = page.ele('.a-price-fraction')
    if price_whole and price_fraction:
        # 去掉整数部分可能带的换行符和逗号
        whole = price_whole.text.replace('\n', '').replace('.', '')
        data['当前价格'] = f"${whole}.{price_fraction.text}"
        print("成功获取价格")
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
        print("成功获取五点描述")
    else:
        print("警告：未能定位到五点描述")

    # 4. 抓取尺寸 (Size)
    selectors = [
        '#inline-twister-expanded-dimension-text-size_name',
        '#variation_size_name .selection',
        '.inline-twister-dim-title-value'
    ]

    # --- 第一阶段：尝试从页面特定元素获取 ---
    for selector in selectors:
        size_ele = page.ele(selector, timeout=1)
        if size_ele and size_ele.text.strip():
            val = size_ele.text.strip()
            data['尺寸/规格'] = val
            print(f"✅ 选择器捕获: {val}")
            return data

    # --- 第二阶段：模糊匹配（备选方案） ---
    print("🔍 进入保底方案：正则扫描文本...")

    title_ele = page.ele('#productTitle')
    bullets_ele = page.ele('#feature-bullets')
    details_ele = page.ele('#prodDetails')
    title = title_ele.text.strip() if title_ele else ""
    bullets = bullets_ele.text.strip() if bullets_ele else ""
    details = details_ele.text.strip() if details_ele else ""

    full_text = f"{title} {bullets} {details}"
    extracted = extract_size_from_text(full_text)

    if extracted:
        data['尺寸/规格'] = extracted
        print(f"🎯 正则匹配成功: {extracted}")
    else:
        data['尺寸/规格'] = "未发现"
        print("❌ 彻底未发现尺寸")

    return data

if __name__ == '__main__':
    # ================= 1. 准备你的竞品链接 =================
    # 获取当前脚本运行的绝对目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # 让用户输入 ASIN 清单 Excel 文件路径
    input_file = input("请输入 ASIN 清单 Excel 文件路径：").strip()
    if not input_file:
        raise FileNotFoundError("未输入文件路径，程序退出。")
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
        # 动态生成今天的日期文件夹名 (例如: 2026-03-01)
        today_folder = time.strftime('%Y-%m-%d')
        # 拼接完整的存放路径
        target_path = os.path.join(current_dir, today_folder)
        # 自动创建文件夹 (exist_ok=True 表示如果文件夹已存在，直接跳过不报错)
        os.makedirs(target_path, exist_ok=True)
        # 定义文件名
        file_name = f"竞品数据抓取_{time.strftime('%H时%M分')}.xlsx"
        excel_filename = os.path.join(target_path, file_name)
        
        # 写入 Excel，index=False 表示不写入行号 0, 1, 2...
        df.to_excel(excel_filename, index=False)
        print(f"✅ 任务大功告成！数据已保存到当前文件夹：{excel_filename}")
    else:
        print("未能抓取到任何数据，没有生成 Excel。")
        
    # 结束任务，关闭浏览器
    page.quit()