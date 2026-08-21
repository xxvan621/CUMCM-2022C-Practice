"""绘制风化/无风化箱线图，并进行 Mann–Whitney U 检验。

运行方式：
    python mann_whitney_boxplot.py

需要安装：
    pandas numpy scipy matplotlib openpyxl

默认输入：
    ../data_processing/2022Cdata_最终清洗版.xlsx

运行后输出：
    ../图表/高钾玻璃_风化前后箱线图.png
    ../图表/铅钡玻璃_风化前后箱线图.png
    ../results/Mann_Whitney_U检验结果.csv

统计说明：
1. 仅使用“表单2_清洗”中标记为“有效”的67个样本；
2. 在高钾、铅钡玻璃内部，分别比较风化组与无风化组；
3. 对14种归一化成分进行双侧 Mann–Whitney U 检验；
4. 对两类玻璃共28次检验进行 Benjamini–Hochberg 校正；
5. 同时报告秩二列相关系数，正值表示风化组通常更高；
6. 检验反映统计关联，不直接证明风化造成了成分变化。
"""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.font_manager import FontProperties
from scipy.stats import mannwhitneyu


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

DEFAULT_INPUT = PROJECT_DIR / "data_processing" / "2022Cdata_最终清洗版.xlsx"
DEFAULT_FIGURE_DIR = PROJECT_DIR / "图表"
DEFAULT_RESULT_DIR = PROJECT_DIR / "results"

SHEET_NAME = "表单2_清洗"
FONT_PATH = Path("/System/Library/Fonts/STHeiti Medium.ttc")
GROUP_COLUMN = "玻璃类型"
WEATHER_COLUMN = "采样点实际风化"
VALID_COLUMN = "数据有效性"
VALID_VALUE = "有效"
GLASS_TYPES = ("高钾", "铅钡")
WEATHER_STATES = ("无风化", "风化")
ALPHA = 0.05
RANDOM_SEED = 2022

COMPONENTS = {
    "二氧化硅(SiO2)_归一化": "二氧化硅 (SiO2)",
    "氧化钠(Na2O)_归一化": "氧化钠 (Na2O)",
    "氧化钾(K2O)_归一化": "氧化钾 (K2O)",
    "氧化钙(CaO)_归一化": "氧化钙 (CaO)",
    "氧化镁(MgO)_归一化": "氧化镁 (MgO)",
    "氧化铝(Al2O3)_归一化": "氧化铝 (Al2O3)",
    "氧化铁(Fe2O3)_归一化": "氧化铁 (Fe2O3)",
    "氧化铜(CuO)_归一化": "氧化铜 (CuO)",
    "氧化铅(PbO)_归一化": "氧化铅 (PbO)",
    "氧化钡(BaO)_归一化": "氧化钡 (BaO)",
    "五氧化二磷(P2O5)_归一化": "五氧化二磷 (P2O5)",
    "氧化锶(SrO)_归一化": "氧化锶 (SrO)",
    "氧化锡(SnO2)_归一化": "氧化锡 (SnO2)",
    "二氧化硫(SO2)_归一化": "二氧化硫 (SO2)",
}


def load_valid_data(input_path: Path) -> pd.DataFrame:
    """读取有效样本并检查分析所需字段。"""
    data = pd.read_excel(input_path, sheet_name=SHEET_NAME)
    required = {GROUP_COLUMN, WEATHER_COLUMN, VALID_COLUMN, *COMPONENTS.keys()}
    missing = sorted(required.difference(data.columns))
    if missing:
        raise ValueError(f"工作表缺少必需字段：{missing}")

    valid = data.loc[data[VALID_COLUMN].eq(VALID_VALUE)].copy()
    for column in COMPONENTS:
        valid[column] = pd.to_numeric(valid[column], errors="raise")

    expected_counts = {
        ("高钾", "无风化"): 12,
        ("高钾", "风化"): 6,
        ("铅钡", "无风化"): 23,
        ("铅钡", "风化"): 26,
    }
    actual_counts = valid.groupby([GROUP_COLUMN, WEATHER_COLUMN]).size().to_dict()
    if actual_counts != expected_counts:
        raise ValueError(
            "四组样本数量与清洗结果不一致。"
            f"\n预期：{expected_counts}\n实际：{actual_counts}"
        )
    return valid


def benjamini_hochberg(p_values: np.ndarray) -> np.ndarray:
    """对一组p值进行Benjamini–Hochberg FDR校正。"""
    p_values = np.asarray(p_values, dtype=float)
    count = len(p_values)
    order = np.argsort(p_values)
    ranked = p_values[order] * count / np.arange(1, count + 1)
    ranked = np.minimum.accumulate(ranked[::-1])[::-1]
    adjusted = np.empty(count, dtype=float)
    adjusted[order] = np.clip(ranked, 0.0, 1.0)
    return adjusted


def effect_magnitude(value: float) -> str:
    """按照Cliff's delta常用阈值描述效应量大小。"""
    absolute = abs(value)
    if absolute < 0.147:
        return "可忽略"
    if absolute < 0.330:
        return "小"
    if absolute < 0.474:
        return "中等"
    return "大"


def run_tests(data: pd.DataFrame) -> pd.DataFrame:
    """完成28次Mann–Whitney U检验，并进行统一的BH校正。"""
    rows: list[dict[str, object]] = []
    for glass_type in GLASS_TYPES:
        glass = data.loc[data[GROUP_COLUMN].eq(glass_type)]
        for column, display_name in COMPONENTS.items():
            unweathered = glass.loc[
                glass[WEATHER_COLUMN].eq("无风化"), column
            ].dropna().to_numpy(dtype=float)
            weathered = glass.loc[
                glass[WEATHER_COLUMN].eq("风化"), column
            ].dropna().to_numpy(dtype=float)

            # weathered放在第一个位置，因此正效应量表示风化组整体偏高。
            test = mannwhitneyu(
                weathered,
                unweathered,
                alternative="two-sided",
                method="auto",
            )
            u_statistic = float(test.statistic)
            rank_biserial = 2 * u_statistic / (len(weathered) * len(unweathered)) - 1

            rows.append(
                {
                    "玻璃类型": glass_type,
                    "化学成分": display_name,
                    "无风化样本数": len(unweathered),
                    "风化样本数": len(weathered),
                    "无风化中位数": float(np.median(unweathered)),
                    "风化中位数": float(np.median(weathered)),
                    "无风化平均值": float(np.mean(unweathered)),
                    "风化平均值": float(np.mean(weathered)),
                    "平均变化量_百分点": float(np.mean(weathered) - np.mean(unweathered)),
                    "U统计量": u_statistic,
                    "原始p值": float(test.pvalue),
                    "秩二列效应量": rank_biserial,
                }
            )

    result = pd.DataFrame(rows)
    result["BH校正p值_28项"] = benjamini_hochberg(result["原始p值"].to_numpy())
    result["校正后是否显著"] = np.where(
        result["BH校正p值_28项"] < ALPHA, "显著", "不显著"
    )
    result["效应方向"] = np.where(
        result["秩二列效应量"] > 0,
        "风化组较高",
        np.where(result["秩二列效应量"] < 0, "风化组较低", "无方向"),
    )
    result["效应大小"] = result["秩二列效应量"].map(effect_magnitude)
    result["解释"] = np.where(
        result["校正后是否显著"].eq("显著"),
        "差异不太可能仅由随机波动解释",
        "尚无充分证据排除随机波动",
    )
    return result


def format_probability(value: float) -> str:
    if value < 0.001:
        return "<0.001"
    return f"{value:.3f}"


def draw_boxplots(
    data: pd.DataFrame,
    test_result: pd.DataFrame,
    glass_type: str,
    output_dir: Path,
) -> Path:
    """绘制某种玻璃14种成分的箱线图和样本散点。"""
    colors = {"高钾": ("#83C5BE", "#E76F51"), "铅钡": ("#8ECAE6", "#E76F51")}
    no_weather_color, weather_color = colors[glass_type]
    glass = data.loc[data[GROUP_COLUMN].eq(glass_type)]
    glass_tests = test_result.loc[test_result["玻璃类型"].eq(glass_type)].set_index(
        "化学成分"
    )

    font = FontProperties(fname=str(FONT_PATH))
    rng = np.random.default_rng(RANDOM_SEED)
    figure, axes = plt.subplots(4, 4, figsize=(18, 15), dpi=180)
    axes_flat = axes.ravel()

    for index, (column, display_name) in enumerate(COMPONENTS.items()):
        axis = axes_flat[index]
        unweathered = glass.loc[
            glass[WEATHER_COLUMN].eq("无风化"), column
        ].dropna().to_numpy(dtype=float)
        weathered = glass.loc[
            glass[WEATHER_COLUMN].eq("风化"), column
        ].dropna().to_numpy(dtype=float)

        box = axis.boxplot(
            [unweathered, weathered],
            positions=[1, 2],
            widths=0.58,
            patch_artist=True,
            showfliers=False,
            medianprops={"color": "#111827", "linewidth": 2},
            whiskerprops={"color": "#4B5563", "linewidth": 1.2},
            capprops={"color": "#4B5563", "linewidth": 1.2},
        )
        box["boxes"][0].set(facecolor=no_weather_color, alpha=0.72)
        box["boxes"][1].set(facecolor=weather_color, alpha=0.72)

        jitter0 = rng.normal(1, 0.055, size=len(unweathered))
        jitter1 = rng.normal(2, 0.055, size=len(weathered))
        axis.scatter(jitter0, unweathered, s=22, color="#16697A", alpha=0.72, zorder=3)
        axis.scatter(jitter1, weathered, s=22, color="#B23A48", alpha=0.72, zorder=3)

        q_value = float(glass_tests.loc[display_name, "BH校正p值_28项"])
        effect = float(glass_tests.loc[display_name, "秩二列效应量"])
        significant = q_value < ALPHA
        title_color = "#B91C1C" if significant else "#111827"
        axis.set_title(
            f"{display_name}\nq={format_probability(q_value)}，r={effect:+.2f}",
            fontproperties=font,
            fontsize=12,
            color=title_color,
            fontweight="bold" if significant else "normal",
        )
        axis.set_xticks([1, 2], ["无风化", "风化"], fontproperties=font, fontsize=10)
        axis.set_ylabel("归一化含量（%）", fontproperties=font, fontsize=9)
        axis.grid(axis="y", color="#D1D5DB", linewidth=0.8, alpha=0.6)
        axis.spines["top"].set_visible(False)
        axis.spines["right"].set_visible(False)

    for index in range(len(COMPONENTS), len(axes_flat)):
        axes_flat[index].axis("off")

    n0 = int((glass[WEATHER_COLUMN] == "无风化").sum())
    n1 = int((glass[WEATHER_COLUMN] == "风化").sum())
    figure.suptitle(
        f"{glass_type}玻璃风化前后化学成分箱线图（无风化n={n0}，风化n={n1}）",
        fontproperties=font,
        fontsize=24,
        y=0.995,
    )
    figure.text(
        0.5,
        0.012,
        "箱体表示四分位范围，中线为中位数，散点为每个有效采样点；q为28项BH校正p值，r为秩二列效应量。",
        ha="center",
        fontproperties=font,
        fontsize=11,
        color="#4B5563",
    )
    figure.tight_layout(rect=(0.02, 0.035, 0.98, 0.965), h_pad=2.1, w_pad=1.5)

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{glass_type}玻璃_风化前后箱线图.png"
    figure.savefig(output_path, bbox_inches="tight", facecolor="#F8FAFC")
    plt.close(figure)
    return output_path


def save_results(result: pd.DataFrame, output_dir: Path) -> Path:
    """保存可直接用Excel打开的检验结果CSV。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "Mann_Whitney_U检验结果.csv"
    result.to_csv(output_path, index=False, encoding="utf-8-sig", float_format="%.6f")
    return output_path


def print_summary(result: pd.DataFrame, outputs: list[Path]) -> None:
    print("\nMann–Whitney U检验完成。判断标准：BH校正p值 < 0.05。")
    significant = result.loc[result["校正后是否显著"].eq("显著")]
    if significant.empty:
        print("校正后没有达到显著水平的成分。")
    else:
        print("\n校正后显著的成分：")
        for _, row in significant.iterrows():
            print(
                f"  {row['玻璃类型']} / {row['化学成分']}："
                f"q={format_probability(row['BH校正p值_28项'])}，"
                f"效应量={row['秩二列效应量']:+.3f}（{row['效应大小']}）"
            )
    print("\n已生成文件：")
    for output in outputs:
        print(f"  {output}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="绘制箱线图并执行Mann–Whitney U检验")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="输入Excel文件")
    parser.add_argument(
        "--figure-dir", type=Path, default=DEFAULT_FIGURE_DIR, help="箱线图输出目录"
    )
    parser.add_argument(
        "--result-dir", type=Path, default=DEFAULT_RESULT_DIR, help="检验结果输出目录"
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise FileNotFoundError(f"输入文件不存在：{args.input}")
    if not FONT_PATH.exists():
        raise FileNotFoundError(f"中文字体不存在：{FONT_PATH}")

    data = load_valid_data(args.input)
    test_result = run_tests(data)
    result_file = save_results(test_result, args.result_dir)
    figure_files = [
        draw_boxplots(data, test_result, glass_type, args.figure_dir)
        for glass_type in GLASS_TYPES
    ]
    print_summary(test_result, [*figure_files, result_file])


if __name__ == "__main__":
    main()
