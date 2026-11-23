---
name: lab-report
description: >-
  End-to-end workflow and tooling for completing university Python, data science,
  and machine learning lab assignments and producing publication-quality Word
  (.docx), PDF, and optional Jupyter Notebook (.ipynb) lab reports. Use this
  skill whenever the user asks to complete an experiment directory (e.g.
  "完成实验十", "写实验报告", "整理实验代码并生成 docx/pdf 报告") or assemble
  a lab report from an instruction document, dataset, and report template.
---

# 自动完成实验报告（Lab Report Autopilot）

本 Skill 封装了一套从**勘察指导书与数据 → 编写可复现实验脚本 → 真实运行产出图表与 CSV → 撰写工程化章节正文 → Word/PDF 排版装配与多遍视觉自检**的完整流水线，适用于各类 Python / 数据科学 / 机器学习课程实验报告（支持学校下发 `.doc`/`.docx` 母版，也支持无母版时自动生成标准 A4 版式）。

---

## 核心原则（硬红线）

1. **禁止篡改原始下发文件**：指导书、报告母版（`学号-*.doc/.docx`）、原始数据文件一律只读，所有代码、章节与中间产物写入 `<实验目录>/源码/`，图表写入 `<实验目录>/figs/`。
2. **零编造数字**：正文、`section0.md`、`section3.md` 中的每一个指标、统计量、坐标、样例行，必须来自脚本真实运行的 `stdout` 或 `figs/tables/tab*.csv`；严禁先写数字再凑代码，严禁四舍五入到与表格冲突的位数。
3. **严防数据泄漏与硬要求遗漏**：
   - 标准化、缺失值填充（均值/中位数）必须**只用训练集统计量**；
   - 时间序列（股票、气象、日志等）必须**按时间顺序切分，禁止 shuffle**；
   - 若指导书明确要求「手写算法（如 NumPy 梯度下降 / 自写 Jaccard）」或「提交 `.ipynb`」，必须严格落实，不得直接用调库蒙混。
4. **通用化与隐私隔离**：如果目录和用户环境中没有现成的 `profile.json`，**先向用户询问封面信息**（课程名称、专业、班级、学号、姓名、指导教师、实验组别、实验日期），切勿猜测或套用他人信息；可引导用户保存一份 `profile.json` 复用。

---

## 端到端执行流程

### 阶段 0：环境核验与实验勘察

1. 运行勘察脚本（提取指导书正文与**全部表格**、定位母版、检测 CSV 编码/元数据行/字符串标签与本地依赖）：
   ```bash
   python3 <SKILL_DIR>/scripts/inspect_lab.py \
     --dir <实验目录> \
     --dump-instruction <实验目录>/源码/_build/instruction.md
   ```
   > `<SKILL_DIR>` 即本 `SKILL.md` 所在目录（例如 `.agents/skills/lab-report`）。若环境缺库，先执行 `python3 -m pip install -r <SKILL_DIR>/requirements.txt`。
2. 完整阅读导出的 `instruction.md`，逐项列出：
   - 实验目的条目；
   - 每道题目的具体步骤、算法限制（是否要求手写）、固定指标/图表、**课后拓展与思考题**；
   - 是否要求额外交付物（如 `.ipynb`、结果 CSV）。
3. 确认封面信息：检查 `<实验目录>/profile.json`、`<实验目录>/../profile.json` 或 `~/.config/lab-report/profile.json`。若不存在，向用户确认封面 6 项字段及实验日期，再写入配置。

### 阶段 1：制定 `SPEC.md` 与 `report_config.json`

1. 参考 [写作与 SPEC 规范](./references/writing_and_spec.md)，在 `<实验目录>/SPEC.md` 中固化：
   - 交付物清单与目录结构；
   - 每道题目的脚本名、数据读取细节（编码、`skiprows`、标签映射）、随机种子（默认 `random_state=0` / `np.random.default_rng(0)`）；
   - **固定图表编号表**（`figN_1.png..`、`tabN_1.csv..`），禁止写代码时临时改号；
   - 如需多子代理/并行分工（例如题目一、题目二独立执行），明确**互不重叠的独占写入范围**。
2. 参考 [report_config.template.json](./resources/report_config.template.json) 创建 `<实验目录>/源码/report_config.json`，填好 `exp_num`、`exp_label`、`exp_title`、`date`、`body_sections` 与 `cover`（若已有 `profile.json` 可省略 `cover`）。

### 阶段 2：编写实验脚本并实跑产出图表

1. 按题目编写 `<实验目录>/源码/expN_<topic>.py`（多题可并行交给独立子代理，各自只写自己的脚本、`sectionK.md` 和已分配编号的图/表）。
2. 脚本必须遵守 [写作与 SPEC 规范](./references/writing_and_spec.md) 中的代码与绘图规范：
   - `# %% [N.x] 中文小标题` 分节；
   - 非交互 `Agg` 后端 + 跨平台中文字体列表 + `axes.unicode_minus = False`；
   - **图内禁用 `R²` 与 U+2212 负号 `−`**（容易缺字成方框，图内一律写 `R2` 与 ASCII `-`）；
   - `fig.savefig(..., dpi=200, bbox_inches='tight')` 后必须 `plt.close(fig)`；
   - CSV 一律 `encoding='utf-8-sig'`，首行表头，浮点保留 4 位小数。
3. 执行全部实验脚本并核验真实输出：
   - 检查无异常退出、总耗时合理、CSV 数值无 `NaN`/错列；
   - **用 `read_file` 目视检查关键 PNG 图表**：确认无方框缺字、坐标轴/图例不遮挡、文字清晰。
4. 若报告需要「代码 + 运行输出」Notebook 风格截图，运行 [render_notebook.py](./scripts/render_notebook.py)：
   ```bash
   python3 <SKILL_DIR>/scripts/render_notebook.py <脚本.py> --list
   python3 <SKILL_DIR>/scripts/render_notebook.py <脚本.py> --cells 1-3 \
     --out <实验目录>/figs/figN_1.png --title "核心代码与运行输出"
   ```
   > 控制单张截图不要过高（装配上限 19cm），过高时拆减 `--cells` 或精简 `print` 行数（单 cell ≤ 30 行）。

### 阶段 3：撰写章节 Markdown（`section0..3.md`）

1. 按真实运行结果撰写各章节（顺序要求：**先跑完实验脚本，再写 `section0.md` 与 `section3.md`**，确保目的阐述与实验小结也能引用真实数字）：
   - `section0.md`：必须包含 `## 实验目的` 与 `## 实验环境与数据集` 两个标题；
   - `section1.md` / `section2.md` / ...：对应每道题目与课后拓展；
   - `section3.md`：以 `## 实验小结` 开头（装配时会自动剥离该行，填入「三、实验小结」）。
2. 严守 [写作与 SPEC 规范](./references/writing_and_spec.md) 的语法约束：
   - 仅使用 `##`/`###`、单行自然段、`**粗体**`、`` `等宽` ``、```` ```text ```` 围栏块、`{{FIG:...}}`、`{{TAB:...}}`；
   - 禁止 Markdown 管道表 `|...|`、禁止 `{{OUT:...}}` 或其他未知 `{{...}}` 标记；
   - 每个 `###` 小节 400~900 字，遵循「先做法 → 再给数字 → 最后解释原因」，并显式回答指导书中的全部思考题/拓展题；
   - 每个图/表在正文中必须先按「图 N-k」「表 N-k」引用，再插入标记。

### 阶段 4：生成 `.ipynb`（按需）与排版装配自检

1. 若指导书要求提交 Jupyter Notebook，运行 [make_notebooks.py](./scripts/make_notebooks.py)（无需 `ipykernel`）：
   ```bash
   python3 <SKILL_DIR>/scripts/make_notebooks.py \
     --script <实验目录>/源码/expN_part1.py \
     --out <实验目录>/expN_part1.ipynb \
     --title "实验N 题目一：..." \
     --intro "..." --summary "..."
   ```
   （多个脚本可用 `--config notebooks.json --root <实验目录>` 批量生成。）
2. 运行 [build_report.py](./scripts/build_report.py) 装配 `.docx` 并执行多遍 PDF 几何自检：
   ```bash
   python3 <SKILL_DIR>/scripts/build_report.py --root <实验目录> --check
   ```
   > 若目录无学校母版，可自动降级或显式加 `--standalone`；若只需快速验证流水线可用 `--demo --check`。
3. 确认自检退出码为 `0`，且检查终端输出中无异常告警（交叉引用缺失、未知标记、空白页、图片越界、下划线不齐、跨页断表、数字折断、字体回退）。

### 阶段 5：逐页视觉复核与最终交付

1. 用 `read_file` 抽查 `<实验目录>/源码/_build/pdf/page-*.png`：
   - **必查**：第 1 页（封面下划线齐平）、第 2 页（确认无空白页、表头与「一、实验目的」衔接自然）、含密集宽表或高图的页面、末页（实验小结完整）；
   - 如发现视觉瑕疵，优先微调图宽高比或拆简长输出后重跑 `build_report.py --root <实验目录> --check`。
2. 核对最终交付物清单（`.docx`、同名 `.pdf`、源码、图表、CSV、按需 `.ipynb`），确认未误改原始文件，并向用户给出简明交付摘要。

---

## 参考文档与排错

- 详细写作规范、`SPEC.md` 模板与绘图/CSV 约定：[references/writing_and_spec.md](./references/writing_and_spec.md)
- 常见问题与跨平台环境配置（macOS/Windows/Linux 字体、LibreOffice、数据读取坑点）：[references/troubleshooting.md](./references/troubleshooting.md)
- 给同学分享与首次安装说明：[README.md](./README.md)
