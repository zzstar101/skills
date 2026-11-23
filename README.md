# Skills

可复用的 AI Agent Skills 集合（兼容 Antigravity、Claude Code、Codex、Cursor 等支持 `SKILL.md` 规范的 AI 编程助手）。

## 已收录 Skills

| Skill | 路径 | 简介 |
| --- | --- | --- |
| **lab-report** | [`skills/lab-report`](./skills/lab-report/) | 自动完成高校 Python / 数据科学 / 机器学习课程实验与报告：勘察指导书与数据 → 编写并真实运行实验脚本 → 生成高清图表与 CSV → 撰写零编造工程化正文 → 自动装配 Word (`.docx`) / PDF / 可选 `.ipynb` 并多遍排版自检（支持学校 `.doc`/`.docx` 母版与无母版独立模式，跨平台支持 macOS / Windows / Linux）。 |

---

## 快速安装

### 方式一：安装到当前项目（推荐）

```bash
git clone https://github.com/zzstar101/skills.git /tmp/zzstar101-skills
mkdir -p .agents/skills
cp -R /tmp/zzstar101-skills/skills/lab-report .agents/skills/
```

（若使用 Claude Code / Cursor，也可复制到对应工具的 `.claude/skills/` 或 `.cursor/skills/` 目录。）

### 方式二：全局安装（所有项目可用）

- **Antigravity（macOS / Linux）**：
  ```bash
  mkdir -p ~/.gemini/config/skills
  cp -R /tmp/zzstar101-skills/skills/lab-report ~/.gemini/config/skills/
  ```
- **Antigravity（Windows PowerShell）**：
  ```powershell
  git clone https://github.com/zzstar101/skills.git "$env:TEMP\zzstar101-skills"
  New-Item -ItemType Directory -Force "$env:USERPROFILE\.gemini\config\skills" | Out-Null
  Copy-Item -Recurse -Force "$env:TEMP\zzstar101-skills\skills\lab-report" "$env:USERPROFILE\.gemini\config\skills\"
  ```
- **通过 `skills` CLI**：
  ```bash
  npx skills add zzstar101/skills
  ```

---

## 使用 `lab-report`

1. 安装 Python 依赖（首次）：
   ```bash
   python3 -m pip install -r skills/lab-report/requirements.txt
   ```
2. （推荐）安装 LibreOffice 以支持旧版 `.doc` 母版转换与 PDF 多遍排版自检：
   - macOS：`brew install --cask libreoffice`
   - Windows：`winget install TheDocumentFoundation.LibreOffice`
   - Linux：`sudo apt install libreoffice`
3. 把老师下发的实验指导书、数据文件和报告母版放入新目录（如 `实验十/`），直接对 AI 助手说：
   > 用 `lab-report` skill 完成 `实验十/` 的实验和实验报告

详细配置（含整学期复用的个人封面档案 `profile.json`）见 [`skills/lab-report/README.md`](./skills/lab-report/README.md)。

## License

[MIT](./LICENSE)
