# ============================================================
# 2022年全国大学生数学建模竞赛 C题
# 表面风化与颜色的关系分析
#
# 分析目标：
# 1. 探究颜色与表面风化是否存在总体关联
# 2. 解决颜色类别样本量较小导致的稀疏列联表问题
# 3. 检查“玻璃类型”是否可能成为颜色与风化之间的混杂因素
# ============================================================


# ============================================================
# 1. 导入需要的库
# ============================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from pathlib import Path
from scipy.stats import chi2_contingency
from scipy.stats import MonteCarloMethod


# ============================================================
# 2. 设置 Matplotlib 中文字体
# ============================================================

plt.rcParams["font.sans-serif"] = ["Arial Unicode MS"]
plt.rcParams["axes.unicode_minus"] = False


# ============================================================
# 3. 读取清洗后的 Excel
# ============================================================

# 当前 Python 文件所在的位置
script_dir = Path(__file__).resolve().parent

# Excel 位于代码文件夹的上一级
file_path = script_dir.parent / "2022Cdata_最终清洗版.xlsx"

print("\n========== 文件读取 ==========")

print("Excel路径：")
print(file_path)

print("\n文件是否存在：")
print(file_path.exists())


# 读取清洗后的表单1
df = pd.read_excel(
    file_path,
    sheet_name="表单1_清洗"
)

print("\nExcel读取成功！")
print("原始样本数量：", len(df))


# ============================================================
# 4. 提取本次分析需要的三个变量
# ============================================================

data = df[
    [
        "清洗颜色",
        "类型",
        "表面风化"
    ]
].copy()


# 将“清洗颜色”改名为“颜色”
# 后面代码更容易阅读
data = data.rename(
    columns={
        "清洗颜色": "颜色"
    }
)


print("\n========== 原始分析数据 ==========")

print(data.head())


# ============================================================
# 5. 查看颜色分布
# ============================================================

print("\n========== 原始颜色样本数量 ==========")

print(
    data["颜色"].value_counts(
        dropna=False
    )
)


# ============================================================
# 6. 排除“未知”颜色
#
# 原因：
# “未知”不是一种真正的颜色，
# 而是原始数据中颜色信息缺失后的清洗标记。
#
# 因此不能把“未知”当成一个颜色类别参与分析。
# ============================================================

data = data[
    data["颜色"] != "未知"
].copy()


# 如果仍然存在真正的 NaN，也一并排除
data = data.dropna(
    subset=[
        "颜色",
        "类型",
        "表面风化"
    ]
)


print("\n========== 排除未知颜色后 ==========")

print(
    "有效样本数量：",
    len(data)
)


print("\n各颜色样本数量：")

color_counts = (
    data["颜色"]
    .value_counts()
)

print(color_counts)


# ============================================================
# 7. 建立“颜色 × 表面风化”列联表
# ============================================================

cross_table = pd.crosstab(
    data["颜色"],
    data["表面风化"]
)


print("\n========== 颜色 × 表面风化列联表 ==========")

print(cross_table)


# ============================================================
# 8. 计算每种颜色内部的风化比例
# ============================================================

proportion_table = pd.crosstab(
    data["颜色"],
    data["表面风化"],
    normalize="index"
)


print("\n========== 各颜色风化比例 ==========")

print(
    (
        proportion_table * 100
    ).round(2)
)


# ============================================================
# 9. 单独计算：
#
# 每种颜色的
# - 样本量
# - 风化数量
# - 风化率
#
# 这样可以避免只看100%或0%，
# 却忽略该颜色可能只有1～2个样本。
# ============================================================

color_summary = (

    data.assign(
        是否风化=
        data["表面风化"].eq("风化")
    )

    .groupby("颜色")

    .agg(
        样本量=("是否风化", "size"),
        风化数量=("是否风化", "sum"),
        风化率=("是否风化", "mean")
    )

)


# 转换成百分比
color_summary["风化率"] = (
    color_summary["风化率"] * 100
)


# 按风化率从高到低排列
color_summary = color_summary.sort_values(
    "风化率",
    ascending=False
)


print("\n========== 各颜色描述性统计 ==========")

print(
    color_summary.round(2)
)


# ============================================================
# 10. Pearson 卡方独立性检验
#
# 注意：
# 这里先计算普通Pearson卡方，
# 但暂时不能直接根据p-value下结论。
#
# 因为颜色类别多、部分样本量很少，
# 后面必须检查期望频数。
# ============================================================

chi2, p_asymptotic, dof, expected = (
    chi2_contingency(
        cross_table,
        correction=False
    )
)


print("\n========== Pearson 卡方检验 ==========")

print(
    "Chi-square =",
    chi2
)

print(
    "自由度 =",
    dof
)

print(
    "普通渐近 p-value =",
    p_asymptotic
)


# ============================================================
# 11. 输出期望频数
# ============================================================

expected_df = pd.DataFrame(
    expected,
    index=cross_table.index,
    columns=cross_table.columns
)


print("\n========== 卡方检验期望频数 ==========")

print(
    expected_df.round(2)
)


# ============================================================
# 12. 检查 Pearson 卡方检验的适用条件
# ============================================================

expected_array = np.asarray(
    expected
)


# 总单元格数量
total_cells = expected_array.size


# 期望频数 < 5
cells_less_than_5 = np.sum(
    expected_array < 5
)


# 期望频数 < 1
cells_less_than_1 = np.sum(
    expected_array < 1
)


# <5的比例
ratio_less_than_5 = (
    cells_less_than_5 /
    total_cells
)


print("\n========== Pearson卡方适用条件检查 ==========")

print(
    "总单元格数量：",
    total_cells
)

print(
    "期望频数 < 5 的单元格数量：",
    cells_less_than_5
)

print(
    "期望频数 < 5 的比例：",
    f"{ratio_less_than_5:.2%}"
)

print(
    "期望频数 < 1 的单元格数量：",
    cells_less_than_1
)

print(
    "最小期望频数：",
    expected_array.min()
)


# 自动判断
if cells_less_than_1 > 0:

    print(
        "\n结论：存在期望频数 < 1 的单元格。"
    )

    print(
        "Pearson卡方的渐近p值可靠性不足，"
        "因此不能仅依赖普通卡方检验。"
    )

elif ratio_less_than_5 > 0.20:

    print(
        "\n结论：超过20%的单元格期望频数 < 5。"
    )

    print(
        "Pearson卡方的渐近p值需要谨慎解释。"
    )

else:

    print(
        "\nPearson卡方检验的期望频数条件基本满足。"
    )


# ============================================================
# 13. Monte Carlo 卡方检验
#
# 这是这一部分更重要的检验。
#
# 原因：
# 当前是一个稀疏的 R × 2 列联表，
# 普通Pearson卡方依赖的大样本近似可能不可靠。
#
# Monte Carlo通过大量随机重抽样估计p-value。
# ============================================================

print("\n========== Monte Carlo 检验 ==========")


try:

    monte_carlo_method = MonteCarloMethod(
        n_resamples=100000
    )


    result_mc = chi2_contingency(
        cross_table,
        correction=False,
        method=monte_carlo_method
    )


    monte_carlo_p = result_mc.pvalue


    print(
        "Chi-square =",
        result_mc.statistic
    )

    print(
        "Monte Carlo p-value =",
        monte_carlo_p
    )


    # 显著性判断
    if monte_carlo_p < 0.05:

        print(
            "\nMonte Carlo p < 0.05"
        )

        print(
            "拒绝原假设 H0。"
        )

        print(
            "现有数据支持颜色与表面风化之间存在显著关联。"
        )

    else:

        print(
            "\nMonte Carlo p >= 0.05"
        )

        print(
            "不能拒绝原假设 H0。"
        )

        print(
            "现有数据尚无充分证据表明颜色与表面风化存在显著关联。"
        )


except Exception as e:

    print(
        "当前SciPy环境无法直接执行Monte Carlo方法。"
    )

    print(
        "具体错误信息："
    )

    print(e)

    monte_carlo_p = np.nan


# ============================================================
# 14. Cramér's V
#
# 用于衡量：
# 颜色与表面风化之间观察到的关联程度。
#
# 注意：
# Cramér's V表示效应大小，
# 不能代替显著性检验。
# ============================================================

n = cross_table.to_numpy().sum()

r, k = cross_table.shape


cramers_v = np.sqrt(

    chi2 /

    (
        n *
        min(
            r - 1,
            k - 1
        )
    )

)


print("\n========== Cramér's V ==========")

print(
    "Cramér's V =",
    cramers_v
)


# ============================================================
# 15. 检查“颜色 × 玻璃类型”
#
# 目的：
# 看不同颜色是不是集中在某一种玻璃类型中。
#
# 如果颜色与玻璃类型高度绑定，
# 那么颜色与风化之间观察到的关系，
# 可能受到玻璃类型的混杂影响。
# ============================================================

color_type_table = pd.crosstab(
    data["颜色"],
    data["类型"]
)


print("\n========== 颜色 × 玻璃类型 ==========")

print(
    color_type_table
)


# ============================================================
# 16. 计算每种颜色内部的玻璃类型比例
# ============================================================

color_type_prop = pd.crosstab(
    data["颜色"],
    data["类型"],
    normalize="index"
) * 100


print("\n========== 各颜色内部的玻璃类型比例 ==========")

print(
    color_type_prop.round(2)
)


# ============================================================
# 17. 建立三维交叉表
#
# 颜色 × 玻璃类型 × 表面风化
#
# 用于进一步观察：
#
# 某个颜色风化率很高，
# 到底可能是颜色本身的问题，
# 还是因为这个颜色恰好主要属于铅钡玻璃。
# ============================================================

three_way_table = pd.crosstab(
    [
        data["颜色"],
        data["类型"]
    ],
    data["表面风化"]
)


print(
    "\n========== 颜色 × 玻璃类型 × 表面风化 =========="
)

print(
    three_way_table
)


# ============================================================
# 18. 可视化一：
# 不同颜色的100%风化比例堆积柱状图
#
# 作用：
# 直观看每种颜色中
# 风化 / 无风化分别占多少。
# ============================================================

ax1 = proportion_table.plot(
    kind="bar",
    stacked=True,
    figsize=(10, 6)
)


plt.title(
    "不同颜色玻璃文物的表面风化比例",
    fontsize=16
)

plt.xlabel(
    "颜色",
    fontsize=13
)

plt.ylabel(
    "比例",
    fontsize=13
)


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
    rotation=45,
    ha="right"
)


plt.legend(
    title="表面风化情况"
)


# 给柱状图添加百分比
for container in ax1.containers:

    labels = []

    for value in container.datavalues:

        if value >= 0.08:

            labels.append(
                f"{value * 100:.1f}%"
            )

        else:

            labels.append("")


    ax1.bar_label(
        container,
        labels=labels,
        label_type="center",
        fontsize=9
    )


plt.tight_layout()

plt.show()


# ============================================================
# 19. 可视化二：
# 各颜色风化率 + 样本量
#
# 这一张图非常重要。
#
# 因为如果只看到：
# 黑色 = 100%
# 绿色 = 0%
#
# 很容易忽略：
# 黑色可能只有2件
# 绿色可能只有1件
#
# 所以标签同时显示：
# 风化率 + n
# ============================================================

plot_summary = color_summary.copy()


ax2 = plot_summary["风化率"].plot(
    kind="bar",
    figsize=(10, 6)
)


plt.title(
    "不同颜色玻璃文物的表面风化率",
    fontsize=16
)

plt.xlabel(
    "颜色",
    fontsize=13
)

plt.ylabel(
    "风化率 (%)",
    fontsize=13
)


plt.ylim(
    0,
    115
)


plt.xticks(
    rotation=45,
    ha="right"
)


# 给每根柱子标注：
# 风化率 + 样本量
for i, (color, row) in enumerate(
    plot_summary.iterrows()
):

    rate = row["风化率"]

    sample_n = int(
        row["样本量"]
    )


    plt.text(
        i,
        rate + 2,
        f"{rate:.1f}%\n(n={sample_n})",
        ha="center",
        va="bottom",
        fontsize=9
    )


plt.tight_layout()

plt.show()


# ============================================================
# 20. 可视化三：
# 每种颜色内部的玻璃类型构成
#
# 作用：
# 检查“颜色”和“玻璃类型”是否高度绑定。
#
# 如果某个颜色几乎全部属于铅钡，
# 而铅钡本身又更容易风化，
# 那么该颜色的高风化率可能受到玻璃类型影响。
# ============================================================

ax3 = (
    color_type_prop / 100
).plot(
    kind="bar",
    stacked=True,
    figsize=(10, 6)
)


plt.title(
    "不同颜色玻璃文物的玻璃类型构成",
    fontsize=16
)

plt.xlabel(
    "颜色",
    fontsize=13
)

plt.ylabel(
    "玻璃类型占比",
    fontsize=13
)


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
    rotation=45,
    ha="right"
)


plt.legend(
    title="玻璃类型"
)


# 添加百分比标签
for container in ax3.containers:

    labels = []

    for value in container.datavalues:

        if value >= 0.08:

            labels.append(
                f"{value * 100:.1f}%"
            )

        else:

            labels.append("")


    ax3.bar_label(
        container,
        labels=labels,
        label_type="center",
        fontsize=9
    )


plt.tight_layout()

plt.show()


# ============================================================
# 21. 输出最终分析提醒
# ============================================================

print("\n")
print("=" * 60)
print("分析完成")
print("=" * 60)

print(
    "\n请重点查看以下结果："
)

print(
    "1. Monte Carlo p-value：判断颜色与风化是否存在总体显著关联"
)

print(
    "2. Cramér's V：观察颜色与风化的关联程度"
)

print(
    "3. 各颜色样本量：判断0%/100%风化率是否由极小样本造成"
)

print(
    "4. 颜色 × 玻璃类型表：判断颜色是否与玻璃类型高度绑定"
)

print(
    "5. 颜色 × 类型 × 风化三维表：判断是否可能存在玻璃类型混杂"
)

print("=" * 60)