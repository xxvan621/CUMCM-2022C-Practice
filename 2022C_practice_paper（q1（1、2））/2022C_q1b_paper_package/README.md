# 2022C 问题一第二小问 LaTeX 练习包

这个文件夹已经把正文和两张需要插入的箱线图整理到一起。

## 上传到 Overleaf

1. 把 `q1b_section.tex` 上传到 Overleaf 项目根目录。
2. 把整个 `figures` 文件夹上传到 Overleaf，保持文件夹名和图片名不变。
3. 在 `main.tex` 中问题一相应位置加入：

```latex
\input{q1b_section}
```

4. 确认导言区已经加载以下宏包；已有的不要重复添加：

```latex
\usepackage{amsmath}
\usepackage{graphicx}
\usepackage{booktabs}
\usepackage{float}
```

5. 点击 Recompile。若图片报找不到文件，先检查 Overleaf 左侧是否存在
   `figures/q1b_high_k_boxplot.png` 和 `figures/q1b_pbba_boxplot.png`。

## 正文取舍

- 正文保留两张箱线图和两张“仅列显著成分”的三线表。
- 两张完整描述统计表、两张均值变化表以及 28 行完整检验结果适合放附录或支撑材料，不建议全部塞进正文。
- 论文中优先写“风化组与无风化组”，避免把横截面比较直接表述成同一文物的“风化前后变化”。
