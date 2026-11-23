# 常见问题与跨平台排错指南

本指南整理自多轮真实实验报告装配经验，覆盖 Windows / macOS / Linux 环境配置与高频数据/排版坑点。

---

## 1. 跨平台环境与依赖

### 1.1 Python 依赖安装
```bash
python3 -m pip install -r <SKILL_DIR>/requirements.txt
# Windows 若 python3 不可用，改用:
python -m pip install -r <SKILL_DIR>/requirements.txt
```
运行环境自检：
```bash
python3 <SKILL_DIR>/scripts/inspect_lab.py --check-env
```

### 1.2 LibreOffice（`.doc` 转 `.docx` 与 PDF 导出自检）
- **macOS**：`brew install --cask libreoffice`（自动探测 `/opt/homebrew/bin/soffice` 或 `/Applications/LibreOffice.app/Contents/MacOS/soffice`）
- **Windows**：`winget install TheDocumentFoundation.LibreOffice`（自动探测 `C:\Program Files\LibreOffice\program\soffice.exe`）
- **Linux**：`sudo apt install libreoffice`（自动探测 `/usr/bin/soffice`）

> 若暂时无法安装 LibreOffice：
> - 如果母版是 `.docx`（或在 Word/WPS 中先另存为 `.docx`），`build_report.py` **无需 LibreOffice** 即可直接生成最终 `.docx`；
> - 如果没有母版，加 `--standalone` 即可生成标准 A4 `.docx`。只是无法执行自动 PDF 转换与几何自检。

### 1.3 macOS / Linux 下的 LibreOffice 真宋体（SimSun）
Windows 自带 `SimSun`（宋体），LibreOffice 导出的 PDF 与 Word 几乎完全一致。
macOS / Linux 默认没有 `SimSun`，LibreOffice 会回退到其他字体导致 PDF 行距放大 ~30%：
- **macOS 解决**：把 `SimSun.ttf` / `SimSun-Bold.ttf`、`Times New Roman.ttf`、`Consolas.ttf` 放到 `~/Library/Application Support/LibreOffice/4/user/fonts/`；
- **Linux 解决**：放到 `~/.local/share/fonts/` 并执行 `fc-cache -fv`。
- 即使未安装，`.docx` 在装有宋体的 Word/WPS 上打开依然完全正常，仅本地 PDF 预览页数会有差异。

---

## 2. 高频数据读取与建模坑点

1. **指导书正文为空（`paragraphs` 几乎没有内容）**：
   高校实验指导书 `.docx` 经常把全部内容放进一张 9×2 大表格里。切勿只读 `doc.paragraphs`，务必使用 `scripts/inspect_lab.py` 同时提取段落与表格。
2. **CSV 编码报错 `UnicodeDecodeError`**：
   国内教学数据集常见 `GBK` / `GB18030` 编码（如课程学习记录、电商订单表）。`inspect_lab.py` 会自动探测编码并提示 `encoding='gbk'`。
3. **CSV 首行是元数据而非表头**：
   例如波士顿房价 CSV 首行常为 `506,13,,,,,,,,,,,,`，真正的列名在第二行。务必检查原始前 3 行，需要时传 `pd.read_csv(..., skiprows=1)`。
4. **标签列是字符串而非 0/1**：
   例如泰坦尼克数据集的 `Survived` 列可能是 `"survived"` / `"not survived"`（pandas 3.0 显示 dtype 为 `str`），建模前必须显式 `.map({"survived": 1, "not survived": 0})`。
5. **消融实验报 `KeyError`**：
   做多特征方案对比（如 3 特征 vs 6 特征 vs 全部特征）时，必须先用包含**全部候选列**的完整 DataFrame 做一次性 `train_test_split`，再在训练/测试集上按列切片，既保证划分完全一致又避免缺列报错。
6. **数据泄漏**：
   - 缺失值填充（如 `Age` 用中位数）与 `StandardScaler` 只能在训练集上 `fit`，再 `transform` 测试集；
   - 时间序列数据（股票、销量趋势等）必须按时间顺序切前 80% / 后 20%，不要随机 `shuffle`。
7. **无 `ipykernel` / `jupyter` CLI 却要交 `.ipynb`**：
   直接用 `scripts/make_notebooks.py`，它基于 `nbformat` + `exec` + `Figure.savefig` 钩子把 `.py` 脚本和真实输出、PNG 图片写入标准 v4 `.ipynb`，完全不依赖 Jupyter 内核服务。

---

## 3. Word / PDF 排版坑点（已内置修复）

1. **母版转换后出现空白第 2 页**：
   旧版 `.doc` 母版常在封面末尾堆叠 4~6 个空段落把正文表格顶到下一页；再插显式分页符就会多出一页空白。`build_report.py` 的 `normalize_for_word()` 会自动删除主表正前方的尾部空段落再插分页符，且自检会拦截任何空白页。
2. **Word 与 LibreOffice 行距不一致 / 行间压字**：
   母版表格内自带 `docGrid` 文档网格，1.5 倍 `auto` 行距在 Word 兼容模式下可能被压扁。`build_report.py` 统一关闭 `snapToGrid` 并将正文行距固化为 `24pt`（`480 twips exact`），同时跳过含图片的段落以免裁切插图。
3. **图表题跨页分离 / 短表被劈成两半**：
   LibreOffice 表格单元格内常忽略 `keepNext` 与 `pageBreakBefore`。`build_report.py` 对 ≤22 行的三线表使用 1×1 无框 `cantSplit` 容器锁住表题与表体，并在多遍自检中自动对分离的图题执行缩图（14.2 → 13.0 → 12.0cm）或显式分页重排。
4. **图内出现方框 `□`**：
   通常是用了 `R²`（上标 ²）或全角负号 `−`（U+2212），部分中文字体缺字。图内一律改为 `R2` 与 ASCII `-`。
5. **Notebook 截图被压缩得字太小**：
   当 `render_notebook.py` 生成的竖图超过 19cm 高度上限时，Word 会等比缩窄图片。解决方法：减少 `--cells` 数量、拆分为两张图或把单 cell 的 `print` 行数压到 30 行以内。
