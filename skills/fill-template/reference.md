# Fill Template 技能 - 参考文档

## 概述

`fill-template` 技能可自动填写亚马逊批量上传模板。它从两个源文件（SKU编码申请表和Listing文案表）读取数据，填充到亚马逊卖家模板中。

## 文件结构

```
skills/fill-template/
├── SKILL.md                      # 主技能定义
├── config/
│   ├── base.yaml                 # 基础列映射配置（所有类目通用）
│   └── wall_art_paintings.yaml   # Wall Art 类目固定值
├── scripts/
│   ├── fill_template-General.py  # 核心自动化脚本
│   └── validate-config.py        # 验证配置文件
├── reference.md                  # 本文件
├── examples.md                  # 使用示例
└── scripts/
    ├── validate-config.py        # 验证配置文件
    └── quick-run.py              # 预配置路径快速运行脚本
```

## 配置文件

### base.yaml

包含读取SKU和Listing数据的通用列索引映射：

| 区块 | 键 | 说明 |
|------|-----|------|
| `sku_columns` | 列名到0基索引的映射 | 用于读取SKU申请表 |
| `listing_search` | `title_search` + `content_offsets` | 定义如何搜索和读取Listing内容 |
| `template` | 工作表名、表头行、数据起始行 | 模板结构 |

### wall_art_paintings.yaml

包含 Wall Art - Paintings 类目的固定值：

| 区块 | 字段 |
|------|------|
| `category` | name, product_type, item_type_keyword |
| `core` | brand_name, product_id_type, listing_action, manufacturer |
| `product_details` | parentage_level, variation_theme, number_of_items, item_shape |
| `multi_value` | special_features[], material[]（多列填充） |
| `frame` | frame_color, frame_material, frame_type, mounting_type, finish_type, wall_art_form, style |
| `units` | 所有单位字段（长度、宽度、重量） |
| `fulfillment` | item_condition, quantity_us, merchant_shipping_group, fulfillment_channel_code_us |
| `compliance` | country_of_origin, unit_count |

## 数据流程

```
1. 用户选择类目 → 加载类目配置（base.yaml + 类目.yaml）
2. 用户提供文件路径 → SKU编码申请表、Listing文案表、目标模板
3. 用户指定要搜索的Listing标题
4. 脚本从申请表中读取SKU数据（使用sku_columns索引）
5. 脚本通过查找标题并读取偏移行来获取Listing内容
6. 模板被填充合并后的数据
7. 用户确认将值从Template_Copy复制到Template工作表
```

## 核心函数

### `normalize_path(path: str) → str`

标准化文件路径以兼容Windows：
- 去除首尾引号（Windows复制自带）
- 将`~`展开为用户主目录
- 将反斜杠转换为正斜杠
- 去除重复的斜杠

### `read_sku申请表(logger, config) → DataFrame`

使用配置的列索引读取SKU申请表：
- 工作表：`Sheet1`
- 数据从第2行开始（0基索引）
- 过滤掉没有SKU的行

### `read_listing_table(logger, config) → dict`

通过精确标题匹配读取Listing内容：
- 在`search_column`中搜索精确匹配的标题
- 使用`content_offsets`从`read_column`读取内容
- 返回：item_name、bullet_points[]、product_description、generic_keyword

### `process_item_name(template: str, size_str: str) → str`

通过替换来处理Item Name模板：
- `【size】` → 格式化尺寸（如 `12" x 16"`）
- `【num】 pcs` → 件数

### `fill_template_with_data(...)`

向模板行填充：
- 申请表中的SKU数据
- Listing内容（item name、要点、描述、关键词）
- 类目配置中的固定值
- 从型号自动计算的尺寸字符串

## 模板处理流程

1. 以`keep_vba=True`保留设置打开模板
2. 创建`Template_Copy`工作表作为工作副本
3. 列标题从第4行读取
4. 数据行从第9行开始（可配置）
5. 用户确认后，将值复制到`Template`工作表
6. 复制完成后删除`Template_Copy`

## 添加新类目

要添加新的产品类目：

1. 创建`config/<category_name>.yaml`，包含：
   - `category`区块（name、product_type、item_type_keyword）
   - `fixed_values`区块，包含所有必需字段

2. 在`scripts/fill_template-General.py`中更新`select_category()`：
   ```python
   categories = {
       '1': 'wall_art_paintings',
       '2': 'your_new_category',  # 添加此项
   }
   ```

## 依赖项

- `pandas` - DataFrame操作
- `openpyxl` - Excel文件操作
- `pyyaml` - 配置文件解析
- `logging` - 内置日志模块

## 错误处理

脚本会验证：
- 处理前检查文件是否存在
- 标题匹配（未找到则报错）
- 尺寸的数值转换
- 必填字段是否为空

所有错误都带有描述性消息记录。
