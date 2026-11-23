# 写作、脚本与 SPEC 规范参考

本文件补充 `SKILL.md` 的细则，供编写 `SPEC.md`、实验脚本和 `section*.md` 时查阅。

---

## 1. 标准目录结构

```text
<实验目录>/
├── <学号>-<姓名>-<实验标签>-《<实验题目>》实验报告.docx   # 最终 Word 交付物
├── <学号>-<姓名>-<实验标签>-《<实验题目>》实验报告.pdf    # 最终 PDF 交付物（同名）
├── <expN_xxx.ipynb>                                      # 按需：指导书要求时生成
├── SPEC.md                                               # 单一事实源（协作/编号/基准）
├── <原始指导书.docx / .pdf>                               # 原始下发，禁止修改
├── <原始报告母版.doc / .docx>                             # 原始下发，禁止修改
├── <原始数据.csv / .xlsx ...>                             # 原始下发，禁止修改
├── figs/
│   ├── figN_1.png ... figN_M.png                         # 实验图表 / 代码运行截图
│   └── tables/
│       └── tabN_1.csv ... tabN_K.csv                     # 实验数据表（UTF-8-SIG）
└── 源码/
    ├── report_config.json                                # 本次实验的报告装配配置
    ├── expN_<part1>.py                                   # 题目一脚本
    ├── expN_<part2>.py                                   # 题目二脚本（可按题增减）
    ├── section0.md                                       # 实验目的 + 实验环境与数据集
    ├── section1.md                                       # 题目一正文
    ├── section2.md                                       # 题目二正文（含课后拓展）
    ├── section3.md                                       # 实验小结
    └── _build/                                           # 装配中间产物（母版转换/自检PNG）
```

---

## 2. `SPEC.md` 必备章节模板

每个新实验在动工前先落 `<实验目录>/SPEC.md`，至少包含：

1. **交付物与独占写入范围**（若拆分给多名子代理并行，每人只写自己那一行，禁止交叉写入）：
   | 角色 | 独占写入范围 |
   | --- | --- |
   | Lead | `SPEC.md`、`源码/report_config.json`、`源码/section0.md`、`源码/section3.md`、`.ipynb`、最终 `.docx`/`.pdf`、`源码/_build/` |
   | A（题目一） | `源码/expN_part1.py`、`源码/section1.md`、`figs/figN_1..a.png`、`figs/tables/tabN_1..b.csv` |
   | B（题目二） | `源码/expN_part2.py`、`源码/section2.md`、`figs/figN_(a+1)..m.png`、`figs/tables/tabN_(b+1)..n.csv` |
2. **环境与数据注意事项**：Python 解释器、随机种子（默认 `random_state=0` / `np.random.default_rng(0)`）、数据编码、需跳过的元数据行、标签字符串映射、缺失值处理策略（严防泄漏）。
3. **固定图表编号清单**：逐一列出 `图 N-1..M`、`表 N-1..K` 的内容与列名规范。
4. **每题硬要求与思考题清单**：对照指导书逐条列出必须覆盖的指标、图形、手写算法约束及课后思考题。
5. **基准数字锁定区**（可选但强烈推荐）：试跑真实数据后记录关键基准值，任何环节发现偏差先核对，禁止私自改数。

---

## 3. 实验脚本规范

### 3.1 通用头模板（跨平台 + 兼容 CLI 与 Notebook/exec）

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""实验N 题目X：<题目名称>"""

from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# 跨平台中文字体回退链（macOS / Windows / Linux）
plt.rcParams["font.sans-serif"] = [
    "Hiragino Sans GB", "Heiti SC", "PingFang SC",
    "Microsoft YaHei", "SimHei", "SimSun",
    "Noto Sans CJK SC", "WenQuanYi Micro Hei", "DejaVu Sans",
]
plt.rcParams["axes.unicode_minus"] = False
pd.set_option("display.unicode.east_asian_width", True)

# 同时兼容命令行直接运行与 notebook/exec（__file__ 不存在时回退 cwd）
HERE = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
ROOT = HERE.parent if HERE.name == "源码" else HERE
DATA_DIR = ROOT
FIG_DIR = ROOT / "figs"
TAB_DIR = ROOT / "figs" / "tables"
FIG_DIR.mkdir(parents=True, exist_ok=True)
TAB_DIR.mkdir(parents=True, exist_ok=True)


def save_fig(fig, name: str):
    fig.savefig(FIG_DIR / name, dpi=200, bbox_inches="tight")
    plt.close(fig)  # 必须关闭，防止 make_notebooks 重复捕获同一图形


def save_csv(df: pd.DataFrame, name: str):
    df.to_csv(TAB_DIR / name, index=False, encoding="utf-8-sig", float_format="%.4f")
    return df
```

### 3.2 脚本与可视化细则

- **分节标记**：用 `# %% [N.1] 中文小标题` 划分步骤（这是 `render_notebook.py` 与 `make_notebooks.py` 切分 cell 的唯一正则依据）。
- **图内禁用缺字字符**：matplotlib 图标题、坐标轴、图例与标注中**禁止使用 `R²`（U+00B2）和全角减号 `−`（U+2212）**，统一写 `R2` 和 ASCII `-`。（Word 正文和 CSV 表格中可以写 `R²`，`build_report.py` 会自动将其转为真正的上标 `2`。）
- **画布尺寸**：单图 `figsize=(7.2, 4.4)`，双面板 `(11.5, 4.5)`，2×2 网格 `(11.0, 8.5)`，2×3 网格 `(12.0, 7.6)`；统一 `dpi=200`。
- **CSV 规范**：首行即中文表头，不加额外标题行；连续数值统一保留 **4 位小数**，计数类保持整数；长数字串不要无故拉宽到十几位小数，以免三线表被迫降字号。
- **输出精炼**：所有注释、`print` 一律中文；每个 cell 的 `stdout` 控制在 **30 行以内**，不要把几百行 DataFrame 整表打印出来。
- **严禁**：不要调用 `plt.show()`、不要联网下载数据、不要写 `<实验目录>` 之外的文件路径。

---

## 4. `section*.md` 格式与写作规范

### 4.1 `build_report.py` 支持的全部语法

```markdown
## 题目一 基于线性支持向量机的数据分类

### 1. 数据构造与分布观察

正文自然段：一整段写在同一行，不要在句中强行换行。支持 **行内加粗** 与 `行内代码`。
正文中先按「表 N-1」「图 N-1」引用（中间一个半角空格），切勿只写“见下图”“如下表”。

{{TAB:tables/tabN_1.csv|表 N-1 数据集基本统计量}}

{{FIG:figs/figN_1.png|图 N-1 样本散点分布}}

{{FIG:figs/figN_2a.png|figs/figN_2b.png|图 N-2 上下叠放双图共用一个图题}}

```text
回归方程：y = 0.9968 * x + 0.0110
测试集指标：MAE = 0.2140, RMSE = 0.3012, R2 = 0.8745
```
```

**注意事项**：
- `section0.md` **必须**包含 `## 实验目的` 与 `## 实验环境与数据集` 两个标题（脚本按此切分填入表格第 3 行「一、实验目的」与第 4 行「二、实验过程」开头）。
- `section3.md` 以 `## 实验小结` 开头，下分 `### 1. ...` 等小节。
- **不支持** Markdown 管道表格 `| a | b |`（全部改用 CSV + `{{TAB:...}}`）；**不支持** `{{OUT:...}}`（全部改用 ```` ```text ```` 围栏块）；**不支持**任意其他 `{{...}}` 标记（遇到未知标记会直接报错终止）。
- `{{FIG:...}}` 与 `{{TAB:...}}` 前后各空一行；表路径写 `tables/tabN_1.csv` 或 `figs/tables/tabN_1.csv` 均可自动解析。
- 段落末尾应使用正常中文标点结尾；避免小于 6 字的孤立碎片段落。

### 4.2 工程化写作风格

1. **三段式信息密度**：每个 `###` 小节 **400~900 字**，遵循「**先说做法与参数 → 再给真实数字与图表定位 → 最后解释为什么/指出边界与局限**」。
2. **只写实验本身**：正文只陈述算法原理、实验设计、真实结果与分析，**不要写**“本报告由脚本自动生成”“通过本次实验我学到了很多”等流水线元说明或套话。
3. **显式回答每一道拓展/思考题**：在对应章节末尾逐条正面作答，并给出脚本运行得到的量化证据。
4. **中文排版细节**：中文正文用全角标点，算法缩写与专有名词保持通篇统一（如 `K-Means`、`SVM`、`KNN`、`OLS`、`AUC`、`F1`、`IQR`）。
5. **诚实讨论局限**：若数据集较小、线性可分导致测试准确率 1.0000 或存在类别不均衡/上限截断，主动在分析与小结中指出局限（如样本量分辨率、隐式正则化、单次划分波动、后续交叉验证/网格搜索方向），报告质量会显著高于只夸“准确率 100% 很好”。
