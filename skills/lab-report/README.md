# lab-report — 自动完成课程实验报告 Skill（通用分享版）

把实验指导书、原始数据和报告母版放到一个文件夹里，让 AI 助手帮你完成：**数据探查 → 编写并运行 Python 实验代码 → 生成高清图表与 CSV → 撰写工程化实验分析 → 自动装配排版规范的 Word（`.docx`）与 PDF（及可选 `.ipynb`）实验报告，并逐页自检排版质量**。

---

## 1. 安装到你的 Antigravity 环境

将整个 `lab-report/` 文件夹复制到以下任一位置（二选一）：

- **仅在当前课程项目中使用（推荐）**：
  `<你的课程目录>/.agents/skills/lab-report/`
- **所有项目全局通用**：
  - macOS / Linux：`~/.gemini/config/skills/lab-report/`
  - Windows：`%USERPROFILE%\.gemini\config\skills\lab-report\`

## 2. 安装本地依赖（首次只需做一次）

1. 安装 Python 依赖：
   ```bash
   python3 -m pip install -r lab-report/requirements.txt
   ```
   （Windows 若没有 `python3` 命令，把 `python3` 换成 `python`。）

2. （强烈推荐）安装 LibreOffice，用于自动转换 `.doc` 旧版母版、导出 PDF 以及执行多遍排版自检：
   - **macOS**：`brew install --cask libreoffice`
   - **Windows**：`winget install TheDocumentFoundation.LibreOffice`
   - **Linux**：`sudo apt install libreoffice`

3. 检查环境是否就绪：
   ```bash
   python3 lab-report/scripts/inspect_lab.py --check-env
   ```

## 3. 配置你的个人信息（配置一次，整学期复用）

复制 `resources/profile.example.json` 到你的课程根目录（例如 `ML/profile.json`）或 `~/.config/lab-report/profile.json`，填好你自己的信息：

```json
{
  "cover": {
    "course": "机器学习原理与应用",
    "major": "计算机科学与技术",
    "class": "25计科1班",
    "student_id": "202500000001",
    "name": "你的姓名",
    "teacher": "任课教师姓名"
  },
  "group": "单人组",
  "output_pattern": "{student_id}-{name}-{exp_label}-《{exp_title}》实验报告.docx"
}
```

> 即使不提前创建 `profile.json`，第一次使用时助手也会主动询问你的封面信息。

## 4. 使用方法

1. 把老师下发的实验指导书（`.docx` / `.doc` / `.pdf`）、实验报告母版（`.doc` / `.docx`，没有也可以）和数据集（`.csv` / `.xlsx` 等）放进新文件夹，例如 `实验十/`。
2. 在 Antigravity 中直接对助手说：
   - `用 lab-report skill 完成 实验十/ 的实验报告`
   - 或直接：`帮我完成 实验十/ 的实验和实验报告`
3. 助手会自动完成：
   - 勘察指导书与数据集 → 制定 `SPEC.md`
   - 编写并真实运行实验脚本 → 生成图（`figs/fig*.png`）与表（`figs/tables/tab*.csv`）
   - 基于真实输出撰写各章节内容（零编造数字）
   - 如指导书要求，自动生成带真实运行结果和内嵌图片的 `.ipynb`
   - 自动装配 `.docx` + 导出同名 `.pdf` + 最多 5 遍几何自检（检查空白页、封面下划线齐平、三线表跨页/断字、图表题分离、字体回退）并目视抽查页面截图。

## 5. 自带工具脚本一览

| 脚本 | 作用 |
| --- | --- |
| `scripts/inspect_lab.py` | 环境检查（`--check-env`）、提取指导书正文与表格、定位母版、自动探测 CSV 编码与异常头行 |
| `scripts/build_report.py` | 通用 Word/PDF 排版装配与多遍自检（支持学校母版，也支持 `--standalone` 无母版标准版式与 `--demo` 流水线自测） |
| `scripts/render_notebook.py` | 把 `# %%` 分节脚本渲染成 Jupyter Notebook 风格的「代码 + 输出」高清截图 |
| `scripts/make_notebooks.py` | 无需 `ipykernel`，把 `# %%` 分节脚本真实执行并生成带 stdout 与内嵌 PNG 的 `.ipynb` |

## 6. 分享给同学前

本 Skill 目录内**不含任何个人隐私信息或本机绝对路径**，直接打包整个 `lab-report/` 文件夹发给同学即可开箱使用。
