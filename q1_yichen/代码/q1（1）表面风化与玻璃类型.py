# ==========================================
# 2022年国赛C题
# 表面风化 与 玻璃类型 的关系分析
# ==========================================


# ==========================================
# 1. 导入需要的库
# ==========================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from pathlib import Path
from scipy.stats import chi2_contingency
from scipy.stats import fisher_exact


# ==========================================
# 2. 设置中文字体
# 解决 Mac 上 matplotlib 中文显示为方框的问题
# ==========================================

plt.rcParams["font.sans-serif"] = ["Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False


# ==========================================
# 3. 读取清洗后的 Excel 数据
# ==========================================

# 当前 q1.py 所在文件夹
script_dir = Path(__file__).resolve().parent

# Excel 文件位于 q1.py 上一级文件夹 q1_yichen 中
file_path = script_dir.parent / "2022Cdata_最终清洗版.xlsx"

print("Python正在读取：")
print(file_path)

print("\n文件是否存在：")
print(file_path.exists())


df = pd.read_excel(
    file_path,
    sheet_name="表单1_清洗"
)

print("\nExcel读取成功！")

print("\n数据前5行：")
print(df.head())

print("\n数据规模：")
print(df.shape)


# ==========================================
# 4. 提取本问题需要的两个变量
# ==========================================

data = df[
    ["类型", "表面风化"]
].copy()

print("\n========== 本题分析数据 ==========")

print(data.head())


# ==========================================
# 5. 检查缺失值
# ==========================================

print("\n========== 缺失值检查 ==========")

print(data.isnull().sum())


# ==========================================
# 6. 查看变量频数
# ==========================================

print("\n========== 玻璃类型频数 ==========")

print(
    data["类型"].value_counts()
)


print("\n========== 表面风化频数 ==========")

print(
    data["表面风化"].value_counts()
)


# ==========================================
# 7. 建立列联表
# ==========================================

cross_table = pd.crosstab(
    data["类型"],
    data["表面风化"]
)

print("\n========== 玻璃类型 × 表面风化 列联表 ==========")

print(cross_table)


# ==========================================
# 8. 计算各玻璃类型内部的风化比例
# ==========================================

proportion_table = pd.crosstab(
    data["类型"],
    data["表面风化"],
    normalize="index"
)

print("\n========== 各玻璃类型内部比例 ==========")

print(proportion_table)


print("\n========== 各玻璃类型内部百分比 ==========")

print(
    (proportion_table * 100).round(2)
)


# ==========================================
# 9. Pearson 卡方独立性检验
# ==========================================

chi2, p, dof, expected = chi2_contingency(
    cross_table,
    correction=False
)

print("\n========== Pearson 卡方检验 ==========")

print("卡方统计量 chi-square =", chi2)

print("p-value =", p)

print("自由度 degrees of freedom =", dof)


# ==========================================
# 10. 输出期望频数
# ==========================================

expected_df = pd.DataFrame(
    expected,
    index=cross_table.index,
    columns=cross_table.columns
)

print("\n========== 期望频数 ==========")

print(expected_df)


# ==========================================
# 11. 根据 p-value 判断显著性
# ==========================================

alpha = 0.05

print("\n========== 显著性判断 ==========")

if p < alpha:

    print("p < 0.05")

    print("拒绝原假设 H0。")

    print(
        "结论：玻璃类型与表面风化之间存在显著关联。"
    )

else:

    print("p >= 0.05")

    print("不能拒绝原假设 H0。")

    print(
        "结论：目前没有充分证据认为玻璃类型与表面风化存在显著关联。"
    )


# ==========================================
# 12. Fisher 精确检验
# ==========================================

odds_ratio, fisher_p = fisher_exact(
    cross_table
)

print("\n========== Fisher 精确检验 ==========")

print(
    "Odds Ratio =",
    odds_ratio
)

print(
    "Fisher exact p-value =",
    fisher_p
)


if fisher_p < alpha:

    print(
        "Fisher检验结果显著。"
    )

else:

    print(
        "Fisher检验结果不显著。"
    )


# ==========================================
# 13. 计算 Cramér's V
# 衡量两个分类变量之间关联强度
# ==========================================

n = cross_table.to_numpy().sum()

cramers_v = np.sqrt(

    chi2 /

    (
        n *

        min(
            cross_table.shape[0] - 1,
            cross_table.shape[1] - 1
        )
    )

)

print("\n========== Cramér's V ==========")

print(
    "Cramér's V =",
    cramers_v
)


# ==========================================
# 14. 单独计算各玻璃类型风化率
# ==========================================

weathering_rate = (

    data.assign(

        是否风化=
        data["表面风化"].eq("风化")

    )

    .groupby("类型")["是否风化"]

    .mean()

    * 100

)


print("\n========== 各玻璃类型风化率 ==========")

for glass_type, rate in weathering_rate.items():

    print(
        f"{glass_type}玻璃风化率：{rate:.2f}%"
    )


# ==========================================
# 15. 可视化一
# 100%堆积柱状图
# 最推荐用于论文正文
# ==========================================

ax = proportion_table.plot(

    kind="bar",

    stacked=True,

    figsize=(8, 6)

)


plt.title(
    "不同玻璃类型的表面风化比例",
    fontsize=16
)

plt.xlabel(
    "玻璃类型",
    fontsize=13
)

plt.ylabel(
    "比例",
    fontsize=13
)


# Y轴显示为百分比
plt.yticks(

    [
        0,
        0.2,
        0.4,
        0.6,
        0.8,
        1.0
    ],

    [
        "0%",
        "20%",
        "40%",
        "60%",
        "80%",
        "100%"
    ]

)


plt.xticks(
    rotation=0,
    fontsize=12
)


plt.legend(
    title="表面风化情况"
)


# ==========================================
# 在柱子中显示百分比
# ==========================================

for container in ax.containers:

    labels = [

        f"{value * 100:.1f}%"

        if value > 0

        else ""

        for value in container.datavalues

    ]

    ax.bar_label(

        container,

        labels=labels,

        label_type="center",

        fontsize=11

    )


plt.tight_layout()

plt.show()


# ==========================================
# 16. 可视化二
# 分组柱状图
# 比较风化/无风化数量
# ==========================================

ax2 = cross_table.plot(

    kind="bar",

    figsize=(8, 6)

)


plt.title(
    "不同玻璃类型的表面风化数量",
    fontsize=16
)

plt.xlabel(
    "玻璃类型",
    fontsize=13
)

plt.ylabel(
    "文物数量",
    fontsize=13
)


plt.xticks(
    rotation=0,
    fontsize=12
)


plt.legend(
    title="表面风化情况"
)


# 在柱子顶部显示数量
for container in ax2.containers:

    ax2.bar_label(
        container,
        fmt="%.0f",
        fontsize=11
    )


plt.tight_layout()

plt.show()


# ==========================================
# 17. 可视化三
# 各玻璃类型风化率柱状图
# ==========================================

ax3 = weathering_rate.plot(

    kind="bar",

    figsize=(7, 5)

)


plt.title(
    "不同玻璃类型的风化率",
    fontsize=16
)

plt.xlabel(
    "玻璃类型",
    fontsize=13
)

plt.ylabel(
    "风化率 (%)",
    fontsize=13
)


plt.xticks(
    rotation=0,
    fontsize=12
)


# 给风化率柱子添加数值标签
for container in ax3.containers:

    labels = [

        f"{value:.1f}%"

        for value in container.datavalues

    ]

    ax3.bar_label(

        container,

        labels=labels,

        padding=3,

        fontsize=11

    )


plt.tight_layout()

plt.show()