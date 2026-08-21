"""
生成高钾玻璃、铅钡玻璃风化前后化学成分描述性统计图表。

统计量：样本数、平均值、标准差、中位数、第一四分位数、
第三四分位数、最小值、最大值。
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

DEFAULT_INPUT = PROJECT_DIR / "data_processing" / "2022Cdata_最终清洗版.xlsx"
DEFAULT_OUTPUT = PROJECT_DIR / "results"
SHEET_NAME = "表单2_清洗"
FONT_PATH = Path("/System/Library/Fonts/STHeiti Medium.ttc")

GROUP_COLUMN = "玻璃类型"
WEATHER_COLUMN = "采样点实际风化"
VALID_COLUMN = "数据有效性"
VALID_VALUE = "有效"
GLASS_TYPES = ("高钾", "铅钡")
WEATHER_STATES = ("无风化", "风化")

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

STAT_COLUMNS = ["样本数", "平均值", "标准差", "中位数", "Q1（25%）", "Q3（75%）", "最小值", "最大值"]


def load_valid_data(input_path: Path) -> pd.DataFrame:
    """读取清洗表，仅保留有效样本。"""
    data = pd.read_excel(input_path, sheet_name=SHEET_NAME)
    required = {GROUP_COLUMN, WEATHER_COLUMN, VALID_COLUMN, *COMPONENTS.keys()}
    missing = sorted(required.difference(data.columns))
    if missing:
        raise ValueError(f"工作表缺少必需字段: {missing}")

    valid = data.loc[data[VALID_COLUMN].eq(VALID_VALUE)].copy()
    for column in COMPONENTS:
        valid[column] = pd.to_numeric(valid[column], errors="coerce")
    return valid


def describe_group(data: pd.DataFrame, glass_type: str, weather: str) -> pd.DataFrame:
    """计算单一玻璃类型、单一风化状态下 14 种成分的统计量。"""
    group = data.loc[
        data[GROUP_COLUMN].eq(glass_type) & data[WEATHER_COLUMN].eq(weather),
        list(COMPONENTS),
    ]
    if group.empty:
        raise ValueError(f"未找到分组数据: {glass_type} / {weather}")

    result = pd.DataFrame(index=list(COMPONENTS))
    result["样本数"] = group.count()
    result["平均值"] = group.mean()
    result["标准差"] = group.std(ddof=1)
    result["中位数"] = group.median()
    result["Q1（25%）"] = group.quantile(0.25)
    result["Q3（75%）"] = group.quantile(0.75)
    result["最小值"] = group.min()
    result["最大值"] = group.max()
    result.index = result.index.map(COMPONENTS)
    result.index.name = "化学成分"
    return result[STAT_COLUMNS]


def format_for_table(stats: pd.DataFrame) -> list[list[str]]:
    """将数值格式化为图表单元格文本。"""
    rows: list[list[str]] = []
    for component, values in stats.iterrows():
        row = [component, f"{int(values['样本数'])}"]
        row.extend(f"{values[col]:.2f}" for col in STAT_COLUMNS[1:])
        rows.append(row)
    return rows


def centered_text(
    draw: ImageDraw.ImageDraw,
    xy: tuple[float, float, float, float],
    text: str,
    font: ImageFont.FreeTypeFont,
    fill: str,
    left_align: bool = False,
) -> None:
    """在指定矩形内垂直居中绘制文本。"""
    x0, y0, x1, y1 = xy
    box = draw.textbbox((0, 0), text, font=font)
    text_width = box[2] - box[0]
    text_height = box[3] - box[1]
    x = x0 + 20 if left_align else x0 + (x1 - x0 - text_width) / 2
    y = y0 + (y1 - y0 - text_height) / 2 - box[1]
    draw.text((x, y), text, font=font, fill=fill)


def draw_section(
    draw: ImageDraw.ImageDraw,
    stats: pd.DataFrame,
    weather: str,
    top: int,
    header_color: str,
    stripe_color: str,
    fonts: dict[str, ImageFont.FreeTypeFont],
) -> None:
    """绘制单一风化状态的统计表。"""
    left = 120
    table_width = 3360
    section_title_height = 74
    row_height = 64
    col_widths = [700, 265, 335, 335, 335, 350, 350, 335, 355]
    headers = ["化学成分", *STAT_COLUMNS]
    rows = format_for_table(stats)
    sample_count = int(stats.iloc[0]["样本数"])

    draw.rounded_rectangle(
        (left, top, left + 420, top + 54), radius=16, fill=header_color
    )
    centered_text(
        draw,
        (left, top, left + 420, top + 54),
        f"{weather}（n = {sample_count}）",
        fonts["section"],
        "white",
    )
    table_top = top + section_title_height

    x_positions = [left]
    for width in col_widths:
        x_positions.append(x_positions[-1] + width)
    if x_positions[-1] != left + table_width:
        raise ValueError("表格列宽总和与表格宽度不一致")

    for col, header in enumerate(headers):
        cell = (x_positions[col], table_top, x_positions[col + 1], table_top + row_height)
        draw.rectangle(cell, fill=header_color, outline="#D5DCE5", width=2)
        centered_text(draw, cell, header, fonts["header"], "white")

    for row_index, row in enumerate(rows, start=1):
        y0 = table_top + row_index * row_height
        y1 = y0 + row_height
        background = stripe_color if row_index % 2 == 0 else "white"
        for col, value in enumerate(row):
            cell = (x_positions[col], y0, x_positions[col + 1], y1)
            draw.rectangle(cell, fill=background, outline="#D5DCE5", width=2)
            centered_text(
                draw,
                cell,
                value,
                fonts["component"] if col == 0 else fonts["body"],
                "#172033",
                left_align=(col == 0),
            )


def draw_glass_table(
    data: pd.DataFrame,
    glass_type: str,
    output_dir: Path,
) -> Path:
    """将无风化和风化统计表绘制在同一张图内。"""
    palette = {
        "高钾": ("#176B5B", "#EAF5F1"),
        "铅钡": ("#315A8C", "#EDF3FA"),
    }
    header_color, stripe_color = palette[glass_type]
    image = Image.new("RGB", (3600, 2520), "#F7F8FA")
    draw = ImageDraw.Draw(image)
    fonts = {
        "title": ImageFont.truetype(str(FONT_PATH), 58),
        "section": ImageFont.truetype(str(FONT_PATH), 32),
        "header": ImageFont.truetype(str(FONT_PATH), 27),
        "component": ImageFont.truetype(str(FONT_PATH), 27),
        "body": ImageFont.truetype(str(FONT_PATH), 26),
        "note": ImageFont.truetype(str(FONT_PATH), 26),
    }

    centered_text(
        draw,
        (120, 45, 3480, 135),
        f"{glass_type}玻璃风化前后化学成分描述性统计表",
        fonts["title"],
        "#111827",
    )
    section_tops = (175, 1325)
    for weather, top in zip(WEATHER_STATES, section_tops):
        stats = describe_group(data, glass_type, weather)
        draw_section(draw, stats, weather, top, header_color, stripe_color, fonts)

    note = "注：成分含量为归一化质量百分比（%）；未检出成分按 0 计；标准差为样本标准差（ddof=1）。"
    centered_text(draw, (120, 2445, 3480, 2500), note, fonts["note"], "#4B5563")

    output_path = output_dir / f"{glass_type}玻璃_风化前后化学成分统计表.png"
    image.save(output_path, format="PNG", dpi=(220, 220), optimize=True)
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成玻璃化学成分描述性统计图表")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="输入 Excel 文件")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT, help="PNG 输出目录")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if not args.input.exists():
        raise FileNotFoundError(f"输入文件不存在: {args.input}")
    if not FONT_PATH.exists():
        raise FileNotFoundError(f"中文字体不存在: {FONT_PATH}")

    args.output_dir.mkdir(parents=True, exist_ok=True)
    data = load_valid_data(args.input)
    outputs = [draw_glass_table(data, glass_type, args.output_dir) for glass_type in GLASS_TYPES]

    print("已生成统计图表：")
    for output in outputs:
        print(output)


if __name__ == "__main__":
    main()
