# Fill Template 技能 - 使用示例

## 基本用法

### 标准交互流程

```
1. 运行脚本
   python skills/fill-template/scripts/fill_template-General.py

2. 选择类目（目前仅有 Wall Art - Paintings）
   请选择模板类目：
     1. Wall Art - Paintings
   请输入序号 (1): 1

3. 按提示输入文件路径
   请输入产品信息表路径（SKU编码申请表）: C:\data\sku申请.xlsx
   请输入Listing文案表路径: C:\data\listing表.xlsx
   请输入Listing文案表的工作表名称: Sheet1
   请输入需要修改的模板文件路径: C:\templates\wall_art_template.xlsx

4. 输入要搜索的Listing标题
   请输入要查找的Listing标题: Abstract Sunset Painting

5. 确认复制到Template工作表
   是否将数据复制到Template表？(yes/no): yes
```

## 示例文件结构

### SKU编码申请表结构

base.yaml中定义的列索引（0基索引）：

| 列名 | 索引 | 说明 |
|------|------|------|
| 序号 | 0 | 序号 |
| 材质 | 1 | 材质 |
| 作品名称 | 2 | 产品名称 |
| 型号 | 3 | 型号（包含尺寸信息） |
| PCS数 | 4 | 件数 |
| 颜色 | 5 | 颜色 |
| 规格长 | 6 | 长度 |
| 规格宽 | 7 | 宽度 |
| 规格高 | 8 | 高度 |
| 重量LB_毛重 | 9 | 毛重 |
| 重量LB_净重 | 10 | 净重 |
| 亚马逊SKU | 19 | 亚马逊SKU |

### Listing文案表结构

每个产品的预期布局（使用content_offsets）：

| 行偏移 | 内容 |
|--------|------|
| +1 | Item Name（商品名称） |
| +2 到 +6 | 5个Bullet Points（要点） |
| +7 | Product Description（产品描述） |
| +8 | Generic Keywords（通用关键词） |

### 型号格式

型号（型号）会被解析以提取尺寸。支持的格式：

- `12x16-1pcs` → 尺寸: `12" x 16"`，1件
- `24x36x0.8-1pcs` → 尺寸: `24" x 36" x 0.8"`，1件

## Item Name模板示例

在Listing表中，Item Name可以包含占位符：

```
Abstract Sunset Art - 【size】 - 【num】 pcs
```

如果 型号 = `24x36-1pcs`，结果为：
```
Abstract Sunset Art - 24" x 36" - 1 pcs
```

## 使用预配置路径运行

对于频繁使用的场景，可以创建带硬编码路径的脚本：

```python
# scripts/quick-run-example.py
import sys
sys.path.insert(0, 'skills/fill-template')
from fill_template-General import main

if __name__ == '__main__':
    import os
    os.environ['SKU_PATH'] = r'C:\data\sku申请.xlsx'
    os.environ['LISTING_PATH'] = r'C:\data\listing表.xlsx'
    os.environ['TEMPLATE_PATH'] = r'C:\templates\wall_art_template.xlsx'
    main()
```

## 运行前验证

处理前先验证配置：

```bash
python skills/fill-template/scripts/validate-config.py
```

验证内容：
- 所有必需的配置文件是否存在
- YAML语法是否正确
- 必需的配置区块是否完整

## 常见问题与解决方案

### "未找到标题"（Title not found）

**原因：** 精确匹配失败
**解决方法：** 检查Listing表中的标题拼写，确保没有多余的空格或特殊字符

### 尺寸数值转换错误

**原因：** 型号格式不符合预期
**解决方法：** 确保型号格式为 `12x16-1pcs` 或 `24x36x0.8-1pcs`

### Template_Copy工作表没有更新

**原因：** 模板保存后没有刷新
**解决方法：** 重新运行脚本，每次都会重新创建Template_Copy
