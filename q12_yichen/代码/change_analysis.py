"""计算并绘制高钾、铅钡玻璃风化前后的化学成分变化表。

计算定义：
    变化量 = 风化组均值 - 无风化组均值（单位：百分点）
    相对变化率 = 变化量 / 无风化组均值 × 100%

说明：
1. 使用表单2_清洗中的有效样本和归一化成分；
2. 先用完整精度计算，最后仅在展示时四舍五入；
3. 无风化均值小于 0.10% 时，不展示相对变化率，避免小基数失真；
4. “升高/降低”是描述性方向，不代表已经通过显著性检验。
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

DEFAULT_INPUT = PROJECT_DIR / "data_processing" / "2022Cdata_最终清洗版.xlsx"
DEFAULT_OUTPUT = PROJECT_DIR / "图表"
SHEET_NAME = "表单2_清洗"
FONT_PATH = Path("/System/Library/Fonts/STHeiti Medium.ttc")

GROUP_COLUMN = "玻璃类型"
WEATHER_COLUMN = "采样点实际风化"
VALID_COLUMN = "数据有效性"
VALID_VALUE = "有效"
GLASS_TYPES = ("高钾", "铅钡")
SMALL_BASE = 0.10

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
    """读取清洗表，只保留有效样本，并保证成分列为数值。"""
    data = pd.read_excel(input_path, sheet_name=SHEET_NAME)
    required = {GROUP_COLUMN, WEATHER_COLUMN, VALID_COLUMN, *COMPONENTS.keys()}
    missing = sorted(required.difference(data.columns))
    if missing:
        raise ValueError(f"工作表缺少必需字段：{missing}")

    valid = data.loc[data[VALID_COLUMN].eq(VALID_VALUE)].copy()
    for column in COMPONENTS:
        valid[column] = pd.to_numeric(valid[column], errors="raise")
    return valid


def calculate_change_table(data: pd.DataFrame, glass_type: str) -> pd.DataFrame:
    """按完整精度计算某种玻璃风化前后的成分变化。"""
    glass = data.loc[data[GROUP_COLUMN].eq(glass_type)]
    unweathered = glass.loc[
        glass[WEATHER_COLUMN].eq("无风化"), list(COMPONENTS)
    ]
    weathered = glass.loc[
        glass[WEATHER_COLUMN].eq("风化"), list(COMPONENTS)
    ]
    if unweathered.empty or weathered.empty:
        raise ValueError(f"{glass_type}缺少风化组或无风化组数据")

    unweathered_mean = unweathered.mean()
    weathered_mean = weathered.mean()
    delta = weathered_mean - unweathered_mean
    relative_rate = delta.div(unweathered_mean).mul(100)

    result = pd.DataFrame(
        {
            "化学成分": [COMPONENTS[column] for column in COMPONENTS],
            "无风化均值": unweathered_mean.to_numpy(),
            "风化均值": weathered_mean.to_numpy(),
            "变化量": delta.to_numpy(),
            "相对变化率": relative_rate.to_numpy(),
        }
    )
    result["变化方向"] = np.where(
        result["变化量"] > 0,
        "相对升高",
        np.where(result["变化量"] < 0, "相对降低", "基本不变"),
    )
    result["解释提示"] = np.where(
        result["无风化均值"].abs() < SMALL_BASE,
        "基数过小，不解释变化率",
        np.where(
            result["无风化均值"].abs() < 0.50,
            "基数较小，变化率仅供参考",
            "—",
        ),
    )
    result["绝对变化排名"] = (
        result["变化量"].abs().rank(method="min", ascending=False).astype(int)
    )

    # 归一化数据的各成分均值之和应为100%，变化量之和应接近0。
    if not np.isclose(result["无风化均值"].sum(), 100.0, atol=1e-6):
        raise ValueError(f"{glass_type}无风化组均值合计不为100%")
    if not np.isclose(result["风化均值"].sum(), 100.0, atol=1e-6):
        raise ValueError(f"{glass_type}风化组均值合计不为100%")
    if not np.isclose(result["变化量"].sum(), 0.0, atol=1e-6):
        raise ValueError(f"{glass_type}变化量合计不为0")

    result.attrs["无风化样本数"] = len(unweathered)
    result.attrs["风化样本数"] = len(weathered)
    return result


def centered_text(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    text: str,
    font: ImageFont.FreeTypeFont,
    fill: str,
    left_align: bool = False,
) -> None:
    x0, y0, x1, y1 = box
    text_box = draw.textbbox((0, 0), text, font=font)
    width = text_box[2] - text_box[0]
    height = text_box[3] - text_box[1]
    x = x0 + 18 if left_align else x0 + (x1 - x0 - width) / 2
    y = y0 + (y1 - y0 - height) / 2 - text_box[1]
    draw.text((x, y), text, font=font, fill=fill)


def format_rate(row: pd.Series) -> str:
    if abs(row["无风化均值"]) < SMALL_BASE:
        return "—"
    return f"{row['相对变化率']:+.2f}%"


def draw_change_table(
    result: pd.DataFrame,
    glass_type: str,
    output_dir: Path,
) -> Path:
    palette = {
        "高钾": ("#176B5B", "#EAF5F1"),
        "铅钡": ("#315A8C", "#EDF3FA"),
    }
    header_color, stripe_color = palette[glass_type]
    increase_fill = "#FDECEC"
    decrease_fill = "#EAF2FD"
    highlight_fill = "#FFF3CD"
    grid_color = "#D5DCE5"
    text_color = "#172033"

    image = Image.new("RGB", (3300, 1510), "#F7F8FA")
    draw = ImageDraw.Draw(image)
    fonts = {
        "title": ImageFont.truetype(str(FONT_PATH), 54),
        "subtitle": ImageFont.truetype(str(FONT_PATH), 29),
        "header": ImageFont.truetype(str(FONT_PATH), 27),
        "body": ImageFont.truetype(str(FONT_PATH), 25),
        "component": ImageFont.truetype(str(FONT_PATH), 25),
        "note": ImageFont.truetype(str(FONT_PATH), 23),
    }

    centered_text(
        draw,
        (100, 35, 3200, 125),
        f"{glass_type}玻璃风化前后化学成分变化表",
        fonts["title"],
        "#111827",
    )
    n0 = result.attrs["无风化样本数"]
    n1 = result.attrs["风化样本数"]
    subtitle = f"无风化 n = {n0}；风化 n = {n1}；变化量 = 风化均值 − 无风化均值"
    centered_text(draw, (100, 125, 3200, 185), subtitle, fonts["subtitle"], "#4B5563")

    left = 80
    top = 215
    row_height = 70
    widths = [560, 300, 300, 330, 330, 270, 270, 660]
    headers = [
        "化学成分",
        "无风化均值(%)",
        "风化均值(%)",
        "变化量(百分点)",
        "相对变化率",
        "变化方向",
        "|Δ|排名",
        "解释提示",
    ]
    positions = [left]
    for width in widths:
        positions.append(positions[-1] + width)

    for col, header in enumerate(headers):
        cell = (positions[col], top, positions[col + 1], top + row_height)
        draw.rectangle(cell, fill=header_color, outline=grid_color, width=2)
        centered_text(draw, cell, header, fonts["header"], "white")

    for row_index, (_, row) in enumerate(result.iterrows(), start=1):
        y0 = top + row_index * row_height
        y1 = y0 + row_height
        base_fill = stripe_color if row_index % 2 == 0 else "white"
        values = [
            row["化学成分"],
            f"{row['无风化均值']:.2f}",
            f"{row['风化均值']:.2f}",
            f"{row['变化量']:+.2f}",
            format_rate(row),
            row["变化方向"],
            str(row["绝对变化排名"]),
            row["解释提示"],
        ]
        for col, value in enumerate(values):
            fill = base_fill
            if col in (3, 5):
                fill = increase_fill if row["变化量"] > 0 else decrease_fill
            if col == 6 and row["绝对变化排名"] <= 3:
                fill = highlight_fill
            cell = (positions[col], y0, positions[col + 1], y1)
            draw.rectangle(cell, fill=fill, outline=grid_color, width=2)
            centered_text(
                draw,
                cell,
                str(value),
                fonts["component"] if col == 0 else fonts["body"],
                text_color,
                left_align=(col in (0, 7)),
            )

    note1 = "注1：所有计算使用未四舍五入的组均值，表中数值仅为显示时保留两位小数。"
    note2 = "注2：无风化均值 < 0.10% 时不展示相对变化率；颜色只表示方向，是否显著需另做统计检验。"
    draw.text((100, 1335), note1, font=fonts["note"], fill="#4B5563")
    draw.text((100, 1380), note2, font=fonts["note"], fill="#4B5563")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{glass_type}玻璃_风化前后成分变化表.png"
    image.save(output_path, format="PNG", dpi=(220, 220), optimize=True)
    return output_path


def print_analysis(result: pd.DataFrame, glass_type: str) -> None:
    key = result.nsmallest(5, "绝对变化排名")
    print(f"\n{glass_type}玻璃变化幅度最大的5种成分：")
    for _, row in key.iterrows():
        print(
            f"  {row['化学成分']}：{row['无风化均值']:.2f}% → "
            f"{row['风化均值']:.2f}%，变化 {row['变化量']:+.2f} 个百分点"
        )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成风化前后化学成分变化表")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="输入Excel文件")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT, help="PNG输出目录")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise FileNotFoundError(f"输入文件不存在：{args.input}")
    if not FONT_PATH.exists():
        raise FileNotFoundError(f"中文字体不存在：{FONT_PATH}")

    data = load_valid_data(args.input)
    outputs: list[Path] = []
    for glass_type in GLASS_TYPES:
        result = calculate_change_table(data, glass_type)
        print_analysis(result, glass_type)
        outputs.append(draw_change_table(result, glass_type, args.output_dir))

    print("\n已生成变化表：")
    for output in outputs:
        print(output)


if __name__ == "__main__":
    main()
