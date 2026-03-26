import os
import shutil

# 1. 配置路径
image_folder = r'C:\Users\Administrator\Desktop\个人\分析\油画\复活节画'
output_folder = r'C:\Users\Administrator\Desktop\个人\分析\油画\复活节画\映射表'

# 2. 定义对应关系 (图片文件, ASIN 列表)
# 基于 Excel 64个不重复 ASIN，按物理尺寸/比例对应
tasks = [
    # 春季郁金香与野兔
    # 横3:2 (12,16,20,24寸): B0GT9VZ4RY, B0GT9F48TD, B0GT9RYD7W, B0GT99DB38
    # 竖2:3 (12,16,20,24寸): B0GT9QVH9G, B0GT9JQ373, B0GT9RGHRV, B0GT9PMV4S
    ("春季郁金香与野兔-3比2.png", ["B0GT9VZ4RY", "B0GT9F48TD", "B0GT9RYD7W", "B0GT99DB38"]),
    ("春季郁金香与野兔-2比3.png", ["B0GT9QVH9G", "B0GT9JQ373", "B0GT9RGHRV", "B0GT9PMV4S"]),

    # 经典复活节花环与雏鸡
    # 横3:2 (12,16,20,24寸): B0GT9GNXXH, B0GT9CCHRR, B0GT9R1TBZ, B0GT9FX5T3
    # 竖2:3 (12,16,20,24寸): B0GT9TPNP3, B0GT9MPMF8, B0GT9NTBDJ, B0GT9G63BX
    # 横5:4 (16x20, 8x10寸): B0GT9FBWWD, B0GTLWPKBW
    ("经典复活节花环与雏鸡_3比2.png", ["B0GT9GNXXH", "B0GT9CCHRR", "B0GT9R1TBZ", "B0GT9FX5T3"]),
    ("经典复活节花环与雏鸡_2比3.png", ["B0GT9TPNP3", "B0GT9MPMF8", "B0GT9NTBDJ", "B0GT9G63BX"]),
    ("经典复活节花环与雏鸡_5比4.png", ["B0GT9FBWWD", "B0GTLWPKBW"]),

    # 花篮与彩蛋
    # 横3:2: B0GT9CRYJQ(12), B0GT9GPCXQ(16), B0GT9R8TDW(20), B0GT9PF4ZX(24)
    # 竖2:3: B0GT9RXFB5(12), B0GT9CH3XJ(16), B0GT9K5DSP(20), B0GT9HYHXC(24)
    ("花篮与彩蛋-3比2.jpg", ["B0GT9CRYJQ", "B0GT9GPCXQ", "B0GT9R8TDW", "B0GT9PF4ZX"]),
    ("花篮与彩蛋平面-2比3.png", ["B0GT9RXFB5", "B0GT9CH3XJ", "B0GT9K5DSP", "B0GT9HYHXC"]),

    # 复活节羊羔
    # 横3:2 (12,16,20,24寸): B0GT9HJDQ5, B0GT9JK12L, B0GT9L58T1, B0GT9M3WK4
    # 竖2:3 (12,16,20,24寸): B0GT9WQR3V, B0GT9QCQH8, B0GT9GMR17, B0GT9K884Z
    # 横5:4 (16x20, 8x10寸): B0GT9HF7B6, B0GTLTHB24
    ("复活节羊羔_3比2.jpg", ["B0GT9HJDQ5", "B0GT9JK12L", "B0GT9L58T1", "B0GT9M3WK4"]),
    ("复活节羊羔_2比3.jpg", ["B0GT9WQR3V", "B0GT9QCQH8", "B0GT9GMR17", "B0GT9K884Z"]),
    ("复活节羊羔_5比4.jpg", ["B0GT9HF7B6", "B0GTLTHB24"]),

    # 复活节的呼吸
    # 横3:2: B0GT9V3PQJ(12), B0GT9JK12K(16), B0GT9NMNV9(20), B0GT9KDJ4K(24)
    ("复活节的呼吸_3比2.png", ["B0GT9V3PQJ", "B0GT9JK12K", "B0GT9NMNV9", "B0GT9KDJ4K"]),

    # 复活节的聚会
    # 横3:2: B0GT9XD32B(12), B0GT9Q14LJ(16), B0GT9NB9H9(20), B0GT9K7MZM(24)
    ("复活节的聚会_3比2.png", ["B0GT9XD32B", "B0GT9Q14LJ", "B0GT9NB9H9", "B0GT9K7MZM"]),

    # 双兔对视
    # 横5:4: B0GT9LXSCX(16x20), B0GTLYTFQ7(8x10)
    ("双兔对视_5比4.jpg", ["B0GT9LXSCX", "B0GTLYTFQ7"]),

    # 野兔抱着彩蛋
    # 横5:4: B0GTB1H5S2(16x20), B0GTLSKB3D(8x10)
    ("野兔抱着彩蛋1_5比4.jpg", ["B0GTB1H5S2", "B0GTLSKB3D"]),

    # 带着希望归家
    # 竖2:3: B0GT9WY4YN(12), B0GT9XX1LK(16), B0GT9QN826(20), B0GT9N9KK4(24)
    ("带着希望归家_2比3.png", ["B0GT9WY4YN", "B0GT9XX1LK", "B0GT9QN826", "B0GT9N9KK4"]),

    # 手推花篮彩蛋
    # 竖2:3: B0GT9GT7QC(12), B0GT9WY4YM(16), B0GT9N8Z9S(20), B0GT9F13DL(24)
    ("手推花篮彩蛋_2比3.png", ["B0GT9GT7QC", "B0GT9WY4YM", "B0GT9N8Z9S", "B0GT9F13DL"]),

    # 林间秘密寻宝
    # 竖2:3: B0GT9VHV25(12), B0GT9RGXTS(16), B0GT9GGFCL(20), B0GT9RN6NR(24)
    ("林间秘密寻宝_2比3.png", ["B0GT9VHV25", "B0GT9RGXTS", "B0GT9GGFCL", "B0GT9RN6NR"]),

    # 午后时光
    # 竖2:3: B0GT9ZCN1P(12), B0GT9JS7BX(16), B0GT9TWZ7N(20), B0GT9NMKT8(24)
    ("午后时光_2比3.png", ["B0GT9ZCN1P", "B0GT9JS7BX", "B0GT9TWZ7N", "B0GT9NMKT8"]),
]

if not os.path.exists(output_folder):
    os.makedirs(output_folder)

def process():
    all_generated = []
    for img_file, asins in tasks:
        source_path = os.path.join(image_folder, img_file)
        if not os.path.exists(source_path):
            print(f"未找到原图: {img_file}")
            continue

        ext = os.path.splitext(img_file)[1]
        print(f"找到原图: {img_file}，正在生成副本...")

        for asin in asins:
            new_name = f"{asin}{ext}"
            shutil.copy(source_path, os.path.join(output_folder, new_name))
            print(f"  -> 已生成: {new_name}")
            all_generated.append(asin)

    print(f"\n总计生成: {len(all_generated)} 个文件")

if __name__ == "__main__":
    process()