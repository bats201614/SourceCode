"""
SellerSprite Product Pipeline
合并爬虫 + 清洗 + SQLite 入库：
  1. DrissionPage 模拟点击导出 Excel
  2. 自动找到下载的 Excel 文件
  3. 读取 → 清洗 → 写入 SQLite（3张表）
"""

import math
import os
import time
import glob
import sqlite3
import pandas as pd
from datetime import date
from DrissionPage import ChromiumPage

# ===== 配置 =====
# 下载相关
BASE_URL = "https://www.sellersprite.com"
DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
# Excel 文件名模式（用于匹配下载的文件）
EXCEL_NAME_PATTERN = "Product-*.xlsx"

# SellerSprite 商品库页面配置
LIBRARY_ID = "701312"   # 商品库 ID
MARKET = "US"          # 市场（US/UK/DE 等）
CID = "460678"         # 类目 ID

# 页面交互配置
SCROLL_STEPS = 8
SCROLL_DISTANCE = 400
SCROLL_PAUSE = 0.5

# 数据库
DB_PATH = r"E:\38 SQLite\products_monitor.db"
EXPORT_DATE = date.today().isoformat()


# ===== 浏览器：导出 Excel =====
def wait_for_login():
    """等待用户在浏览器中手动完成登录"""
    input("请在浏览器中完成登录，然后按 Enter 继续...")


def click_element_safely(page, selectors):
    """尝试多个选择器，返回第一个成功的元素"""
    for selector in selectors:
        try:
            ele = page.ele(selector, timeout=3)
            if ele:
                return ele
        except Exception:
            continue
    return None


def get_latest_excel():
    """获取 Downloads 目录下最新的 Excel 文件"""
    files = glob.glob(os.path.join(DOWNLOAD_DIR, EXCEL_NAME_PATTERN))
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def find_export_checkbox(page):
    """向下滚动找到全选复选框"""
    checkbox = click_element_safely(page, [
        'xpath://div[@class="left"]/label[@data-v-15544ec6]'
    ])
    if checkbox:
        return checkbox

    for _ in range(SCROLL_STEPS):
        page.scroll.down(SCROLL_DISTANCE)
        time.sleep(SCROLL_PAUSE)
        checkbox = click_element_safely(page, [
            'xpath://div[@class="left"]/label[@data-v-15544ec6]'
        ])
        if checkbox:
            return checkbox
    return None


def export_excel():
    """使用 DrissionPage 导出 Excel，返回下载的文件路径"""
    # 记录导出前的最新文件
    latest_before = get_latest_excel()

    page = ChromiumPage()
    page.set.download_path(DOWNLOAD_DIR)

    try:
        page.get(BASE_URL)
        wait_for_login()

        detail_url = f"{BASE_URL}/v3/product-store/{LIBRARY_ID}?market={MARKET}&cid={CID}"
        print(f"打开商品库页面: {detail_url}")
        page.get(detail_url)
        time.sleep(3)

        # 勾选全选
        print("查找全选复选框...")
        checkbox = find_export_checkbox(page)
        if checkbox:
            print("找到全选复选框，点击...")
            checkbox.click()
        else:
            input("未找到全选复选框，请手动勾选后按 Enter 继续...")

        # 点击导出按钮
        print("查找导出按钮...")
        export_btn = click_element_safely(page, [
            'text:导出',
            '@class:export-btn',
            '@class:el-button--primary',
            'tag:button',
        ])
        if export_btn:
            print(f"找到导出按钮: '{export_btn.text.strip()}'，点击...")
            export_btn.click()
            mission = page.wait.download_begin(timeout=60)
            if mission:
                print(f"下载已开始: {mission.name}")
                mission.wait()
                print(f"下载完成: {mission.name}")
            else:
                input("下载未开始，请手动点击导出，等待完成后按 Enter 继续...")
        else:
            input("未找到导出按钮，请手动点击，等待完成后按 Enter 继续...")

    finally:
        page.quit()
        print("浏览器已关闭")

    # 找到最新的 Excel 文件（应该是刚下载的）
    latest_after = get_latest_excel()
    if latest_after and latest_after != latest_before:
        print(f"检测到新文件: {latest_after}")
        return latest_after
    elif latest_after:
        print(f"使用最新文件（可能未变化）: {latest_after}")
        return latest_after
    else:
        print("未找到 Excel 文件")
        return None


# ===== Excel 列名映射（原始 → 标准字段名）=====
COLUMN_MAP = {
    "ASIN": "asin",
    "SKU": "sku",
    "品牌": "brand",
    "品牌链接": "brand_url",
    "商品标题": "product_name",
    "商品详情页链接": "product_url",
    "商品主图": "main_image_url",
    "父ASIN": "parent_asin",
    "类目路径": "category_path",
    "大类目": "big_category",
    "标签": "tags",
    "大类BSR": "big_bsr",
    "大类BSR增长数": "big_bsr_growth_num",
    "大类BSR增长率": "big_bsr_growth_rate",
    "小类目": "small_category",
    "小类BSR": "small_bsr",
    "月销量": "monthly_sales",
    "月销量增长率": "sales_growth_rate",
    "月销售额($)": "monthly_revenue",
    "子体销量": "variant_sales",
    "子体销售额($)": "variant_revenue",
    "变体数": "variant_count",
    "价格($)": "price",
    "prime价格($)": "prime_price",
    "Coupon": "coupon",
    "Q&A": "qa_count",
    "评分数": "review_count",
    "月新增评分数": "monthly_new_reviews",
    "评分": "rating",
    "留评率": "review_rate",
    "FBA($)": "fba_fee",
    "毛利率": "gross_profit_rate",
    "评级": "rating_star",
    "上架时间": "listing_date",
    "上架天数": "listing_days",
    "配送方式": "delivery_type",
    "买家运费($)": "buyer_shipping",
    "LQS": "lqs",
    "卖家数": "seller_count",
    "BuyBox卖家": "buybox_seller",
    "BuyBox类型": "buybox_type",
    "卖家所属地": "seller_location",
    "卖家信息": "seller_info",
    "卖家首页": "seller_homepage",
    "Best Seller标识": "is_best_seller",
    "Amazon's Choice": "is_amazon_choice",
    "New Release标识": "is_new_release",
    "A+页面": "has_a_plus",
    "视频介绍": "has_video",
    "品牌故事": "has_brand_story",
    "品牌广告": "has_brand_ad",
    "7天促销": "has_7day_promo",
    "AC关键词": "ac_keywords",
    "商品重量": "product_weight_lbs",
    "商品重量（单位换算）": "product_weight_kg",
    "商品尺寸": "product_size_in",
    "商品尺寸（单位换算）": "product_size_cm",
    "包装重量": "package_weight_lbs",
    "包装重量（单位换算）": "package_weight_kg",
    "包装尺寸": "package_size_in",
    "包装尺寸（单位换算）": "package_size_cm",
    "包装尺寸分段": "package_size_bucket",
    "详细参数": "detail_params",
    "SP广告": "sp_advertising",
}


# ===== 清洗函数 =====
def to_float(val):
    if val is None:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def to_int(val):
    if val is None:
        return None
    try:
        return int(float(val))
    except (ValueError, TypeError):
        return None


def clean_text(val):
    if val is None:
        return None
    # 处理 pandas 读出的空单元格（np.nan / float('nan')）
    try:
        if isinstance(val, float) and math.isnan(val):
            return None
    except (TypeError, ValueError):
        pass
    val = str(val).strip()
    val = val.replace("&quot;", '"').replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    return val if val else None


def clean_bool_y(val):
    if val is None:
        return 0
    return 1 if str(val).strip().upper() == "Y" else 0


def clean_number_unit(val):
    """从含单位的字符串提取数值，如 '5.27 kg' → 5.27"""
    if val is None:
        return None
    s = str(val).strip()
    import re
    m = re.match(r"^([\d.]+)", s)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return None
    return None


def clean_percent(val):
    """百分比字符串转浮点数（直接存原值，如 -0.01 代表 -1%）"""
    if val is None:
        return None
    s = str(val).strip().replace("%", "")
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def clean_dataframe(df):
    """对 DataFrame 应用所有清洗规则"""
    df = df.copy()

    # 1. 重命名列为标准字段名
    df = df.rename(columns=COLUMN_MAP)

    # 2. 文本字段清洗
    text_cols = [
        "asin", "sku", "brand",
        "brand_url", "product_url", "main_image_url",
        "product_name", "parent_asin", "category_path",
        "big_category", "small_category", "tags",
        "delivery_type", "seller_location",
        "buybox_seller", "buybox_type", "coupon",
        "seller_info", "seller_homepage", "ac_keywords",
        "detail_params", "sp_advertising",
    ]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_text)

    # 3. 整数字段
    int_cols = [
        "big_bsr", "big_bsr_growth_num", "small_bsr",
        "monthly_sales", "variant_sales", "variant_count",
        "qa_count", "review_count", "monthly_new_reviews",
        "listing_days", "seller_count",
    ]
    for col in int_cols:
        if col in df.columns:
            df[col] = df[col].apply(to_int)

    # 4. 浮点数字段
    float_cols = [
        "price", "prime_price", "fba_fee", "rating",
        "buyer_shipping", "lqs", "monthly_revenue",
        "variant_revenue",
    ]
    for col in float_cols:
        if col in df.columns:
            df[col] = df[col].apply(to_float)

    # 4b. 带单位的重量字段
    weight_cols = [
        "product_weight_lbs", "product_weight_kg",
        "package_weight_lbs", "package_weight_kg",
    ]
    for col in weight_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_number_unit)

    # 5. 百分比字段
    percent_cols = ["review_rate", "gross_profit_rate", "big_bsr_growth_rate", "sales_growth_rate"]
    for col in percent_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_percent)

    # 6. 布尔字段（Y/N → 1/0）
    bool_cols = [
        "is_best_seller", "is_amazon_choice", "is_new_release",
        "has_a_plus", "has_video", "has_brand_story",
        "has_brand_ad", "has_7day_promo", "rating_star",
    ]
    for col in bool_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_bool_y)

    # 7. 日期字段
    if "listing_date" in df.columns:
        df["listing_date"] = pd.to_datetime(df["listing_date"], errors="coerce").dt.date
        df["listing_date"] = df["listing_date"].apply(lambda x: x.isoformat() if x else None)

    # 8. 添加导出日期
    df["export_date"] = EXPORT_DATE
    df["record_date"] = EXPORT_DATE

    return df


# ===== SQLite 操作 =====
def init_db(db_path):
    """创建数据库和 3 张表（含触发器）"""
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # ---- 全量表：保留每日快照 ----
    cur.execute("""
        CREATE TABLE IF NOT EXISTS products_full (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            asin TEXT NOT NULL,
            sku TEXT,
            product_name TEXT,
            brand TEXT,
            brand_url TEXT,
            product_url TEXT,
            main_image_url TEXT,
            parent_asin TEXT,
            category_path TEXT,
            big_category TEXT,
            small_category TEXT,
            tags TEXT,
            price REAL,
            prime_price REAL,
            fba_fee REAL,
            rating REAL,
            rating_star INTEGER DEFAULT 0,
            review_count INTEGER,
            monthly_new_reviews INTEGER,
            review_rate REAL,
            gross_profit_rate REAL,
            big_bsr INTEGER,
            big_bsr_growth_num INTEGER,
            big_bsr_growth_rate REAL,
            small_bsr INTEGER,
            monthly_sales INTEGER,
            sales_growth_rate REAL,
            monthly_revenue REAL,
            variant_sales INTEGER,
            variant_revenue REAL,
            variant_count INTEGER,
            qa_count INTEGER,
            coupon TEXT,
            listing_date TEXT,
            listing_days INTEGER,
            delivery_type TEXT,
            buyer_shipping REAL,
            lqs REAL,
            seller_count INTEGER,
            buybox_seller TEXT,
            buybox_type TEXT,
            seller_location TEXT,
            seller_info TEXT,
            seller_homepage TEXT,
            is_best_seller INTEGER DEFAULT 0,
            is_amazon_choice INTEGER DEFAULT 0,
            is_new_release INTEGER DEFAULT 0,
            has_a_plus INTEGER DEFAULT 0,
            has_video INTEGER DEFAULT 0,
            has_brand_story INTEGER DEFAULT 0,
            has_brand_ad INTEGER DEFAULT 0,
            has_7day_promo INTEGER DEFAULT 0,
            ac_keywords TEXT,
            product_weight_lbs REAL,
            product_weight_kg REAL,
            product_size_in TEXT,
            product_size_cm TEXT,
            package_weight_lbs REAL,
            package_weight_kg REAL,
            package_size_in TEXT,
            package_size_cm TEXT,
            package_size_bucket TEXT,
            detail_params TEXT,
            sp_advertising TEXT,
            export_date TEXT,
            record_date TEXT,
            UNIQUE(asin, export_date)
        )
    """)

    # ---- 静态表：ASIN 首次出现记录 ----
    cur.execute("""
        CREATE TABLE IF NOT EXISTS products_static (
            asin TEXT PRIMARY KEY,
            brand TEXT,
            brand_url TEXT,
            main_image_url TEXT,
            parent_asin TEXT,
            category_path TEXT,
            big_category TEXT,
            small_category TEXT,
            tags TEXT,
            delivery_type TEXT,
            listing_date TEXT,
            product_weight_lbs REAL,
            product_weight_kg REAL,
            product_size_in TEXT,
            product_size_cm TEXT,
            package_weight_lbs REAL,
            package_weight_kg REAL,
            package_size_in TEXT,
            package_size_cm TEXT,
            package_size_bucket TEXT,
            detail_params TEXT,
            sp_advertising TEXT,
            product_url TEXT,
            fba_fee REAL,
            gross_profit_rate REAL,
            buyer_shipping REAL,
            seller_location TEXT,
            seller_info TEXT,
            seller_homepage TEXT,
            has_a_plus INTEGER DEFAULT 0,
            has_video INTEGER DEFAULT 0,
            has_brand_story INTEGER DEFAULT 0,
            has_brand_ad INTEGER DEFAULT 0,
            updated_at TEXT,
            first_seen_date TEXT,
            last_seen_date TEXT
        )
    """)

    # 触发器：静态字段实际变化时才更新 updated_at
    cur.execute("""
        CREATE TRIGGER IF NOT EXISTS trg_static_updated_at
        AFTER UPDATE ON products_static
        FOR EACH ROW
        WHEN
            OLD.brand IS NOT NEW.brand OR
            OLD.brand_url IS NOT NEW.brand_url OR
            OLD.main_image_url IS NOT NEW.main_image_url OR
            OLD.parent_asin IS NOT NEW.parent_asin OR
            OLD.category_path IS NOT NEW.category_path OR
            OLD.big_category IS NOT NEW.big_category OR
            OLD.small_category IS NOT NEW.small_category OR
            OLD.tags IS NOT NEW.tags OR
            OLD.delivery_type IS NOT NEW.delivery_type OR
            OLD.listing_date IS NOT NEW.listing_date OR
            OLD.product_weight_lbs IS NOT NEW.product_weight_lbs OR
            OLD.product_weight_kg IS NOT NEW.product_weight_kg OR
            OLD.product_size_in IS NOT NEW.product_size_in OR
            OLD.product_size_cm IS NOT NEW.product_size_cm OR
            OLD.package_weight_lbs IS NOT NEW.package_weight_lbs OR
            OLD.package_weight_kg IS NOT NEW.package_weight_kg OR
            OLD.package_size_in IS NOT NEW.package_size_in OR
            OLD.package_size_cm IS NOT NEW.package_size_cm OR
            OLD.package_size_bucket IS NOT NEW.package_size_bucket OR
            OLD.detail_params IS NOT NEW.detail_params OR
            OLD.sp_advertising IS NOT NEW.sp_advertising OR
            OLD.product_url IS NOT NEW.product_url OR
            OLD.fba_fee IS NOT NEW.fba_fee OR
            OLD.gross_profit_rate IS NOT NEW.gross_profit_rate OR
            OLD.buyer_shipping IS NOT NEW.buyer_shipping OR
            OLD.seller_location IS NOT NEW.seller_location OR
            OLD.seller_info IS NOT NEW.seller_info OR
            OLD.seller_homepage IS NOT NEW.seller_homepage OR
            OLD.has_a_plus IS NOT NEW.has_a_plus OR
            OLD.has_video IS NOT NEW.has_video OR
            OLD.has_brand_story IS NOT NEW.has_brand_story OR
            OLD.has_brand_ad IS NOT NEW.has_brand_ad
        BEGIN
            UPDATE products_static
            SET updated_at = DATE('now')
            WHERE asin = NEW.asin;
        END
    """)

    # ---- 动态表：每日追加，asin + record_date 唯一 ----
    cur.execute("""
        CREATE TABLE IF NOT EXISTS products_dynamic (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            asin TEXT NOT NULL,
            record_date TEXT NOT NULL,
            price REAL,
            prime_price REAL,
            rating REAL,
            rating_star INTEGER DEFAULT 0,
            review_count INTEGER,
            monthly_new_reviews INTEGER,
            review_rate REAL,
            big_bsr INTEGER,
            big_bsr_growth_num INTEGER,
            big_bsr_growth_rate REAL,
            small_bsr INTEGER,
            monthly_sales INTEGER,
            sales_growth_rate REAL,
            monthly_revenue REAL,
            variant_sales INTEGER,
            variant_revenue REAL,
            variant_count INTEGER,
            qa_count INTEGER,
            coupon TEXT,
            listing_days INTEGER,
            lqs REAL,
            seller_count INTEGER,
            buybox_seller TEXT,
            buybox_type TEXT,
            is_best_seller INTEGER DEFAULT 0,
            is_amazon_choice INTEGER DEFAULT 0,
            is_new_release INTEGER DEFAULT 0,
            has_7day_promo INTEGER DEFAULT 0,
            ac_keywords TEXT,
            product_name TEXT,
            sku TEXT,
            UNIQUE(asin, record_date)
        )
    """)

    conn.commit()
    return conn


def _batch_insert(cur, sql, cols, df, extra_vals_fn=None):
    """通用批量插入"""
    cols_in_df = [c for c in cols if c in df.columns]
    placeholders = ", ".join(["?"] * len(cols_in_df))
    cols_sql = ", ".join(cols_in_df)

    vals_list = df[cols_in_df].values.tolist()
    if extra_vals_fn:
        records = df.to_dict("records")
        for i, rec in enumerate(records):
            vals_list[i].extend(extra_vals_fn(rec))

    cur.executemany(sql.format(cols=cols_sql, placeholders=placeholders), vals_list)
    return len(vals_list)


def insert_to_db(conn, df):
    """写入 3 张表，每行数据都经过清洗函数处理"""
    cur = conn.cursor()

    # ========== 1. 全量表 ==========
    full_sql = "INSERT OR REPLACE INTO products_full ({cols}) VALUES ({placeholders})"
    full_rows = _batch_insert(cur, full_sql, [
        "asin", "sku", "product_name", "brand",
        "brand_url", "product_url", "main_image_url",
        "parent_asin", "category_path", "big_category", "small_category", "tags",
        "price", "prime_price", "fba_fee", "rating", "rating_star",
        "review_count", "monthly_new_reviews", "review_rate", "gross_profit_rate",
        "big_bsr", "big_bsr_growth_num", "big_bsr_growth_rate", "small_bsr",
        "monthly_sales", "sales_growth_rate", "monthly_revenue",
        "variant_sales", "variant_revenue", "variant_count", "qa_count", "coupon",
        "listing_date", "listing_days", "delivery_type",
        "buyer_shipping", "lqs", "seller_count",
        "buybox_seller", "buybox_type", "seller_location",
        "seller_info", "seller_homepage",
        "is_best_seller", "is_amazon_choice", "is_new_release",
        "has_a_plus", "has_video", "has_brand_story",
        "has_brand_ad", "has_7day_promo", "ac_keywords",
        "product_weight_lbs", "product_weight_kg",
        "product_size_in", "product_size_cm",
        "package_weight_lbs", "package_weight_kg",
        "package_size_in", "package_size_cm", "package_size_bucket",
        "detail_params", "sp_advertising",
        "export_date", "record_date",
    ], df)
    print(f"  全量表: 写入 {full_rows} 条（export_date={EXPORT_DATE}）")

    # ========== 2. 静态表 ==========
    static_data_cols = [
        "asin", "brand", "brand_url", "main_image_url",
        "parent_asin", "category_path", "big_category", "small_category", "tags",
        "delivery_type", "listing_date",
        "product_weight_lbs", "product_weight_kg",
        "product_size_in", "product_size_cm",
        "package_weight_lbs", "package_weight_kg",
        "package_size_in", "package_size_cm", "package_size_bucket",
        "detail_params", "sp_advertising", "product_url",
        "fba_fee", "gross_profit_rate", "buyer_shipping",
        "seller_location", "seller_info", "seller_homepage",
        "has_a_plus", "has_video", "has_brand_story", "has_brand_ad",
    ]
    static_data_cols_in_df = [c for c in static_data_cols if c in df.columns]
    static_all_cols = static_data_cols_in_df + ["updated_at", "first_seen_date", "last_seen_date"]

    static_vals_list = df[static_data_cols_in_df].values.tolist()
    extra_dates = [[EXPORT_DATE, EXPORT_DATE, EXPORT_DATE]] * len(static_vals_list)
    for i, ed in enumerate(extra_dates):
        static_vals_list[i].extend(ed)

    placeholders = ", ".join(["?"] * len(static_all_cols))
    cur.executemany(
        f"INSERT INTO products_static ({', '.join(static_all_cols)}) "
        f"VALUES ({placeholders}) "
        f"ON CONFLICT(asin) DO UPDATE SET "
        f"brand=excluded.brand, "
        f"brand_url=excluded.brand_url, main_image_url=excluded.main_image_url, "
        f"parent_asin=excluded.parent_asin, category_path=excluded.category_path, "
        f"big_category=excluded.big_category, small_category=excluded.small_category, "
        f"tags=excluded.tags, delivery_type=excluded.delivery_type, "
        f"listing_date=excluded.listing_date, "
        f"product_weight_lbs=excluded.product_weight_lbs, product_weight_kg=excluded.product_weight_kg, "
        f"product_size_in=excluded.product_size_in, product_size_cm=excluded.product_size_cm, "
        f"package_weight_lbs=excluded.package_weight_lbs, package_weight_kg=excluded.package_weight_kg, "
        f"package_size_in=excluded.package_size_in, package_size_cm=excluded.package_size_cm, "
        f"package_size_bucket=excluded.package_size_bucket, "
        f"detail_params=excluded.detail_params, sp_advertising=excluded.sp_advertising, "
        f"product_url=excluded.product_url, "
        f"fba_fee=excluded.fba_fee, gross_profit_rate=excluded.gross_profit_rate, "
        f"buyer_shipping=excluded.buyer_shipping, "
        f"seller_location=excluded.seller_location, "
        f"seller_info=excluded.seller_info, seller_homepage=excluded.seller_homepage, "
        f"has_a_plus=excluded.has_a_plus, has_video=excluded.has_video, "
        f"has_brand_story=excluded.has_brand_story, has_brand_ad=excluded.has_brand_ad, "
        f"last_seen_date=excluded.last_seen_date",
        static_vals_list
    )
    cur.execute("SELECT COUNT(*) FROM products_static WHERE first_seen_date = ?", (EXPORT_DATE,))
    new_count = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM products_static")
    total_static = cur.fetchone()[0]
    print(f"  静态表: 共 {total_static} 条（当日新增 {new_count} 条）")

    # ========== 3. 动态表 ==========
    dyn_rows = _batch_insert(
        cur,
        "INSERT OR REPLACE INTO products_dynamic ({cols}) VALUES ({placeholders})",
        [
            "asin", "price", "prime_price", "rating", "rating_star",
            "review_count", "monthly_new_reviews", "review_rate",
            "big_bsr", "big_bsr_growth_num", "big_bsr_growth_rate", "small_bsr",
            "monthly_sales", "sales_growth_rate", "monthly_revenue",
            "variant_sales", "variant_revenue", "variant_count", "qa_count", "coupon",
            "listing_days", "lqs", "seller_count",
            "buybox_seller", "buybox_type",
            "is_best_seller", "is_amazon_choice", "is_new_release",
            "has_7day_promo", "ac_keywords",
            "product_name", "sku",
            "record_date",
        ],
        df,
    )
    print(f"  动态表: 写入 {dyn_rows} 条")

    conn.commit()


# ===== 主流程 =====
def main():
    print("=" * 50)
    print("SellerSprite Product Pipeline")
    print("=" * 50)

    # Step 1: 导出 Excel
    excel_path = export_excel()
    if not excel_path:
        print("未找到 Excel 文件，退出")
        return

    print(f"\n使用 Excel 文件: {excel_path}")

    # Step 2: 读取并清洗
    print("\n读取 Excel...")
    import pandas as pd
    df_raw = pd.read_excel(excel_path)
    print(f"读取成功: {len(df_raw)} 行, {len(df_raw.columns)} 列")

    print("\n清洗数据...")
    df_clean = clean_dataframe(df_raw)
    print(f"清洗完成: {len(df_clean)} 行, {len(df_clean.columns)} 列")

    # 预览关键字段
    pd.set_option("display.max_columns", 10)
    pd.set_option("display.width", 200)
    preview = ["asin", "brand", "price", "rating", "big_bsr", "monthly_sales"]
    preview = [c for c in preview if c in df_clean.columns]
    print(df_clean[preview].head())

    # Step 3: 入库
    print("\n写入数据库...")
    conn = init_db(DB_PATH)
    insert_to_db(conn, df_clean)

    # Step 4: 数据概览
    cur = conn.cursor()
    print("\n=== SQLite 数据概览 ===")
    for tbl in ["products_full", "products_static", "products_dynamic"]:
        cur.execute(f"SELECT COUNT(*) FROM {tbl}")
        print(f"  {tbl}: {cur.fetchone()[0]} 条")

    print("\n=== 百分比字段验证（抽样3条）===")
    cur.execute("SELECT asin, review_rate, gross_profit_rate, big_bsr_growth_rate FROM products_full LIMIT 3")
    for row in cur.fetchall():
        print(f"  {row}")

    conn.close()
    print(f"\n完成！数据库: {DB_PATH}")


if __name__ == "__main__":
    main()
