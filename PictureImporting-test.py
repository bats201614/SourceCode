import os
from DrissionPage import ChromiumPage, ChromiumOptions

def get_all_images(base_folder):
    """
    自动遍历 POD/orders 目录下所有子文件夹里的图片
    """
    supported_extensions = ('.jpg', '.jpeg', '.png', '.webp')
    all_file_paths = []
    
    for root, dirs, files in os.walk(base_folder):
        for file in files:
            if file.lower().endswith(supported_extensions):
                # 拼接成完整的绝对路径
                full_path = os.path.join(root, file)
                all_file_paths.append(full_path)

    return all_file_paths

def start_upload_task(target_folder):
    # --- 1. 配置接管模式 ---
    co = ChromiumOptions()
    # 指向手动开启的 9222 端口
    co.set_address('127.0.0.1:9222') 
    
    try:
        # 建立连接
        browser = ChromiumPage(co)
    except Exception:
        print("❌ 无法连接到浏览器！")
        print("💡 请先关闭所有 Edge，然后在 CMD 运行: start msedge.exe --remote-debugging-port=9222")
        return

    # --- 2. 定位标签页 ---
    # get_tab 会在当前所有打开的标签页里寻找网址包含 'lovart' 的那一个
    tab = browser.get_tab(url='lovart')
    
    if not tab:
        print("⚠️ 没找到 Lovart 标签页。请先在浏览器里打开工作台！")
        return
    
    print(f"🔗 已成功接管: {tab.title}")

    # --- 3. 扫描图片 ---
    image_list = get_all_images(target_folder)
    if not image_list:
        print("❌ 文件夹内没发现图片。")
        return
    print(f"📦 准备上传 {len(image_list)} 张图片...")

    # --- 4. 执行上传动作 ---
    # 注意：后续操作全部使用 tab，而不是 browser
    nav_icon = tab.ele('@data-testid=generate-menu-trigger')
    
    if nav_icon:
        print("🖱️ 点击上传图标...")
        nav_icon.click()
        
        # 等待弹出菜单中的“上传图片”选项出现，timeout=3 增加容错
        upload_option = tab.wait.ele_displayed('text:上传图片', timeout=3)
        
        if upload_option:
            print("🚀 正在注入路径，这可能需要几秒钟...")
            # 这一步会自动把文件列表塞进系统对话框
            upload_option.click.to_upload(image_list)
            print(f"✅ 任务完成！已成功提交 {len(image_list)} 张图片。")
        else:
            print("❌ 菜单已打开，但没看到“上传图片”选项。")

if __name__ == "__main__":
    # 这里请替换成你实际的图片文件夹路径
    target_folder = r"C:\Users\Administrator\Desktop\自动下载配置_勿删\POD\orders"
    start_upload_task(target_folder)