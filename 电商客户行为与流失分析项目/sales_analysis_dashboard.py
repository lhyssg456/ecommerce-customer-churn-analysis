import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import warnings

warnings.filterwarnings('ignore')

# ---------------------- 1. 配置与加载数据 ----------------------
plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# 你的文件路径
file_path = r"E:\Project Dataset.xlsx"


# 关键修改：遍历所有Sheet，找到真实数据Sheet
def find_data_sheet(file_path):
    """自动找到包含真实数据的Sheet"""
    # 获取所有Sheet名称
    sheet_names = pd.ExcelFile(file_path).sheet_names
    print(f"📌 Excel中的所有Sheet：{sheet_names}")

    for sheet in sheet_names:
        # 读取每个Sheet的前几行
        temp_df = pd.read_excel(file_path, sheet_name=sheet, nrows=10)
        # 判断是否是数据Sheet（有数值列、列名不是Unnamed、行数>10）
        if len(temp_df.columns) > 0 and not all(col.startswith('Unnamed') for col in temp_df.columns):
            if len(temp_df) > 5 and temp_df.select_dtypes(include=[np.number]).shape[1] > 2:
                print(f"✅ 找到真实数据Sheet：{sheet}")
                return sheet, pd.read_excel(file_path, sheet_name=sheet)
    # 如果没找到，返回第一个Sheet（兜底）
    print("⚠️ 未自动识别数据Sheet，使用第一个Sheet")
    return sheet_names[0], pd.read_excel(file_path, sheet_name=sheet_names[0])


# 执行自动识别
sheet_name, df = find_data_sheet(file_path)
print(f"✅ 数据加载成功！数据形状为：{df.shape}")
print("\n📋 数据列名：")
print(df.columns.tolist())
print("\n📋 数据前5行预览：")
print(df.head())

# ---------------------- 2. 数据清洗 ----------------------
# 清理列名（去空格、特殊字符）
df.columns = [col.strip().replace(' ', '_').replace('-', '_') for col in df.columns]
# 删除全是NaN的行/列
df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
# 重置索引
df = df.reset_index(drop=True)
print(f"\n✅ 数据清洗完成，有效数据量：{len(df)} 条")


# ---------------------- 3. 可视化（适配真实数据列名） ----------------------
def generate_plots(df):
    # 3.1 流失客户占比（Churn/流失列）
    churn_cols = [col for col in df.columns if 'Churn' in col or '流失' in col]
    if churn_cols:
        churn_col = churn_cols[0]
        plt.figure(figsize=(8, 6))
        churn_counts = df[churn_col].value_counts()
        plt.pie(churn_counts, labels=['未流失', '流失'] if 0 in churn_counts.index else churn_counts.index,
                autopct='%1.1f%%', colors=['#66b3ff', '#ff9999'], startangle=90)
        plt.title('客户流失率分布', fontsize=14, fontweight='bold')
        plt.savefig('1_客户流失率.png', dpi=300, bbox_inches='tight')
        plt.close()
        print("✅ 生成：1_客户流失率.png")

    # 3.2 订单数量/消费金额分布
    order_cols = [col for col in df.columns if 'Order' in col or '订单' in col or 'Amount' in col or '金额' in col]
    if order_cols:
        order_col = order_cols[0]
        # 只保留数值型数据
        df_order = df[pd.to_numeric(df[order_col], errors='coerce').notna()]
        plt.figure(figsize=(10, 6))
        sns.histplot(pd.to_numeric(df_order[order_col]), bins=20, kde=True, color='green')
        plt.title(f'{order_col} 分布', fontsize=14, fontweight='bold')
        plt.xlabel(order_col)
        plt.ylabel('客户数')
        plt.grid(alpha=0.3)
        plt.savefig('2_订单金额分布.png', dpi=300)
        plt.close()
        print("✅ 生成：2_订单金额分布.png")

    # 3.3 满意度评分分布（Satisfaction/满意度列）
    sat_cols = [col for col in df.columns if 'Satisfaction' in col or '满意' in col]
    if sat_cols:
        sat_col = sat_cols[0]
        plt.figure(figsize=(8, 6))
        sns.countplot(x=sat_col, data=df, palette='Blues')
        plt.title('客户满意度评分分布', fontsize=14, fontweight='bold')
        plt.xlabel('满意度评分')
        plt.ylabel('客户数')
        plt.savefig('3_满意度分布.png', dpi=300)
        plt.close()
        print("✅ 生成：3_满意度分布.png")

    # 3.4 优惠券使用情况
    coupon_cols = [col for col in df.columns if 'Coupon' in col or '优惠' in col or '折扣' in col]
    if coupon_cols:
        coupon_col = coupon_cols[0]
        plt.figure(figsize=(8, 6))
        sns.countplot(x=coupon_col, data=df, palette='Oranges')
        plt.title('优惠券使用情况分布', fontsize=14, fontweight='bold')
        plt.xlabel('优惠券使用次数/是否使用')
        plt.ylabel('客户数')
        plt.savefig('4_优惠券使用.png', dpi=300)
        plt.close()
        print("✅ 生成：4_优惠券使用.png")

    print("\n🎉 所有可生成的图表已完成！")


# 运行可视化
generate_plots(df)