#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通用实验目录勘察工具：检查本地工具链、提取指导书全文（含表格）、定位母版、探查数据文件。

用法::
    # 1) 仅检查当前环境是否满足依赖
    python3 inspect_lab.py --check-env

    # 2) 勘察指定实验目录并把指导书导出为 Markdown
    python3 inspect_lab.py --dir 实验十 --dump-instruction 实验十/源码/_build/instruction.md

    # 3) 输出结构化 JSON 供脚本或 Agent 解析
    python3 inspect_lab.py --dir 实验十 --json
"""

from __future__ import annotations

import argparse
import csv
import importlib.util
import json
import os
import platform
import re
import shutil
import subprocess
import sys
from pathlib import Path

REQUIRED_MODULES = [
    "numpy", "pandas", "matplotlib", "sklearn", "scipy",
    "docx", "fitz", "nbformat", "PIL",
]


def detect_soffice() -> str | None:
    """跨平台定位 LibreOffice soffice 可执行文件。"""
    for name in ("soffice", "libreoffice"):
        hit = shutil.which(name)
        if hit:
            return hit
    candidates = [
        # macOS
        "/opt/homebrew/bin/soffice",
        "/usr/local/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        # Windows
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        # Linux
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None


def detect_simsun() -> bool:
    """跨平台检查系统或 LibreOffice 用户字体目录是否已有宋体（SimSun / SimSong）。"""
    sys_name = platform.system()
    dirs: list[Path] = []
    if sys_name == "Darwin":
        dirs = [
            Path.home() / "Library/Application Support/LibreOffice/4/user/fonts",
            Path.home() / "Library/Fonts",
            Path("/Library/Fonts"),
            Path("/System/Library/Fonts"),
        ]
    elif sys_name == "Windows":
        win = Path(os.environ.get("WINDIR", r"C:\Windows"))
        dirs = [win / "Fonts", Path.home() / "AppData/Local/Microsoft/Windows/Fonts"]
    else:
        dirs = [
            Path.home() / ".config/libreoffice/4/user/fonts",
            Path.home() / ".local/share/fonts",
            Path("/usr/share/fonts"),
        ]
    for d in dirs:
        if not d.exists():
            continue
        for p in d.rglob("*"):
            if p.is_file() and re.search(r"simsun|simsong|songti|noto.*serif.*cjk", p.name, re.I):
                return True
    return False


def inspect_toolchain() -> dict:
    mods = {m: importlib.util.find_spec(m) is not None for m in REQUIRED_MODULES}
    return {
        "os": f"{platform.system()} {platform.release()}",
        "python": sys.executable,
        "python_version": sys.version.split()[0],
        "soffice": detect_soffice(),
        "pandoc": shutil.which("pandoc"),
        "has_songti_font": detect_simsun(),
        "modules": mods,
    }


def extract_docx_markdown(path: Path) -> str:
    """提取 .docx 全部正文段落与表格（自动去重横向合并单元格）。"""
    from docx import Document

    doc = Document(str(path))
    lines: list[str] = [f"# 指导书原文提取：{path.name}", ""]

    para_texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    if para_texts:
        lines.append("## 正文段落")
        lines.append("")
        lines.extend(para_texts)
        lines.append("")

    for ti, tbl in enumerate(doc.tables):
        lines.append(f"## 表格 {ti + 1}（{len(tbl.rows)}×{len(tbl.columns)}）")
        lines.append("")
        seen_tr_text: set[str] = set()
        for ri, row in enumerate(tbl.rows):
            uniq_cells: list[str] = []
            seen_tc = set()
            for cell in row.cells:
                tc_id = id(cell._tc)
                if tc_id in seen_tc:
                    continue
                seen_tc.add(tc_id)
                txt = "\n".join(
                    ln.rstrip() for ln in cell.text.replace("\r\n", "\n").splitlines()
                ).strip()
                if txt:
                    uniq_cells.append(txt)
            if not uniq_cells:
                continue
            joined = " | ".join(uniq_cells)
            if joined in seen_tr_text:
                continue
            seen_tr_text.add(joined)
            if len(uniq_cells) == 1:
                lines.append(f"### Row {ri}")
                lines.append(uniq_cells[0])
                lines.append("")
            else:
                lines.append(f"### Row {ri}: {uniq_cells[0]}")
                for item in uniq_cells[1:]:
                    lines.append(item)
                lines.append("")

    return "\n".join(lines).strip() + "\n"


def extract_pdf_markdown(path: Path) -> str:
    """支持以 PDF 形式下发的实验指导书。"""
    import fitz

    doc = fitz.open(str(path))
    lines: list[str] = [f"# 指导书原文提取（PDF）：{path.name}", ""]
    for i, page in enumerate(doc, start=1):
        lines.append(f"## Page {i}")
        lines.append(page.get_text("text").strip())
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def convert_doc_to_docx(doc_path: Path, out_dir: Path, soffice: str | None) -> Path | None:
    if not soffice:
        return None
    out_dir.mkdir(parents=True, exist_ok=True)
    res = subprocess.run(
        [soffice, "--headless", "--convert-to", "docx", "--outdir", str(out_dir), str(doc_path)],
        capture_output=True, text=True,
    )
    target = out_dir / (doc_path.stem + ".docx")
    return target if target.exists() and res.returncode == 0 else None


def classify_files(exp_dir: Path) -> dict:
    instructions: list[Path] = []
    masters: list[Path] = []
    existing_reports: list[Path] = []
    data_files: list[Path] = []
    other_files: list[Path] = []

    for p in sorted(exp_dir.iterdir()):
        if p.name.startswith(".") or p.name.startswith("~$") or p.is_dir():
            continue
        suf = p.suffix.lower()
        if suf in (".doc", ".docx"):
            if re.match(r"^\d{6,}-", p.name):
                existing_reports.append(p)
            elif any(k in p.name for k in ("学号", "模板", "母版")) or (
                "实验报告" in p.name and ("姓名" in p.name or suf == ".doc")
            ):
                masters.append(p)
            else:
                instructions.append(p)
        elif suf == ".pdf":
            if re.match(r"^\d{6,}-", p.name):
                existing_reports.append(p)
            elif "报告" not in p.name or "指导" in p.name or "要求" in p.name:
                instructions.append(p)
        elif suf in (".csv", ".tsv", ".xlsx", ".xls", ".txt", ".json"):
            data_files.append(p)
        elif suf not in (".ipynb", ".md"):
            other_files.append(p)

    return {
        "instructions": instructions,
        "masters": masters,
        "existing_reports": existing_reports,
        "data_files": data_files,
        "other_files": other_files,
    }


def find_profile(exp_dir: Path) -> Path | None:
    candidates = [
        exp_dir / "profile.json",
        exp_dir.parent / "profile.json",
        Path.home() / ".config/lab-report/profile.json",
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


def detect_encoding(path: Path) -> tuple[str, list[str]]:
    raw = path.read_bytes()[:65536]
    # 截断到最后一个换行符，防止恰好劈断多字节中文字符导致 GBK/UTF-8 误判为 latin1
    if len(raw) == 65536 and b"\n" in raw:
        raw = raw[:raw.rfind(b"\n") + 1]
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk", "latin1"):
        try:
            text = raw.decode(enc)
            return enc, text.replace("\r\n", "\n").splitlines()[:5]
        except UnicodeDecodeError:
            continue
    return "unknown", []


def inspect_dataframe(df, info: dict) -> None:
    info["shape"] = list(df.shape)
    info["duplicates"] = int(df.duplicated().sum())
    cols_info = []
    for col in df.columns:
        ser = df[col]
        missing = int(ser.isna().sum())
        nunique = int(ser.nunique(dropna=True))
        col_rec = {
            "name": str(col),
            "dtype": str(ser.dtype),
            "missing": missing,
            "nunique": nunique,
        }
        if nunique <= 8:
            col_rec["sample_values"] = [str(v) for v in ser.dropna().unique()[:8]]
            if ser.dtype == object or str(ser.dtype) in ("str", "string"):
                if any(k in str(col).lower() for k in ("label", "target", "survived", "churn", "class", "y")):
                    info["hints"].append(
                        f"分类/标签列 `{col}` 为字符串取值 {col_rec['sample_values']}，建模前需显式映射为数值"
                    )
        else:
            col_rec["sample_values"] = [str(v) for v in ser.dropna().head(3).tolist()]
        cols_info.append(col_rec)
    info["columns"] = cols_info


def inspect_data_file(path: Path) -> dict:
    import pandas as pd

    suf = path.suffix.lower()
    info: dict = {
        "file": path.name,
        "size_kb": round(path.stat().st_size / 1024, 1),
        "hints": [],
    }
    if suf in (".xlsx", ".xls"):
        try:
            xl = pd.ExcelFile(path)
            info["sheets"] = xl.sheet_names
            df = xl.parse(xl.sheet_names[0])
            inspect_dataframe(df, info)
        except Exception as exc:
            info["read_error"] = f"{type(exc).__name__}: {exc}"
        return info

    if suf not in (".csv", ".tsv"):
        info["note"] = "非表格文件，请按格式单独探查"
        return info

    enc, preview_lines = detect_encoding(path)
    info["encoding"] = enc
    info["raw_first_3_lines"] = preview_lines[:3]
    sep = "\t" if suf == ".tsv" else ","

    # 检查首行是否为元数据行（如 `506,13,,,,,,,,,,,,`）
    skiprows = 0
    if len(preview_lines) >= 2 and sep == ",":
        row0 = next(csv.reader([preview_lines[0]]), [])
        row1 = next(csv.reader([preview_lines[1]]), [])
        non_empty0 = [c for c in row0 if c.strip()]
        non_empty1 = [c for c in row1 if c.strip()]
        if (
            len(non_empty0) < len(non_empty1)
            and non_empty0
            and all(re.fullmatch(r"[-+]?\d+(\.\d+)?", c.strip()) for c in non_empty0)
        ):
            skiprows = 1
            info["hints"].append("首行疑似元数据行（全为数字且列数少于第二行），建议 pd.read_csv(..., skiprows=1)")
    if enc in ("gbk", "gb18030"):
        info["hints"].append(f"非 UTF-8 编码（{enc}），读取时请显式传 encoding='{enc}'")

    try:
        df = pd.read_csv(path, encoding=enc, sep=sep, skiprows=skiprows)
        info["skiprows_used"] = skiprows
        inspect_dataframe(df, info)
    except Exception as exc:
        info["read_error"] = f"{type(exc).__name__}: {exc}"
    return info


def main() -> int:
    ap = argparse.ArgumentParser(description="通用实验目录勘察与环境自检工具")
    ap.add_argument("--dir", default=None, help="实验目录路径（如 实验十）")
    ap.add_argument("--check-env", action="store_true", help="仅检查本地 Python / LibreOffice / 字体环境")
    ap.add_argument("--dump-instruction", default=None, help="把指导书全文提取保存到指定 .md 文件")
    ap.add_argument("--json", action="store_true", help="输出纯 JSON")
    args = ap.parse_args()

    toolchain = inspect_toolchain()
    if args.check_env or not args.dir:
        if args.json:
            print(json.dumps(toolchain, ensure_ascii=False, indent=2))
        else:
            print("=" * 64)
            print("环境自检（lab-report skill）")
            print("=" * 64)
            print(f"OS:       {toolchain['os']}")
            print(f"Python:   {toolchain['python']} ({toolchain['python_version']})")
            print(f"soffice:  {toolchain['soffice'] or '未找到（无法转换 .doc / 导出 PDF 自检）'}")
            print(f"宋体字体: {'✔ 已检测到' if toolchain['has_songti_font'] else '✖ 未检测到（macOS/Linux PDF 行距可能偏宽）'}")
            missing = [k for k, ok in toolchain["modules"].items() if not ok]
            print(f"Python库: {'✔ 全部就绪' if not missing else '✖ 缺少: ' + ', '.join(missing)}")
            if missing:
                print("  安装命令: python3 -m pip install -r requirements.txt")
        return 0

    exp_dir = Path(args.dir).resolve()
    if not exp_dir.is_dir():
        sys.exit(f"目录不存在：{exp_dir}")

    soffice = toolchain["soffice"]
    files = classify_files(exp_dir)
    profile_path = find_profile(exp_dir)

    instruction_texts: dict[str, str] = {}
    for p in files["instructions"]:
        suf = p.suffix.lower()
        if suf == ".docx":
            instruction_texts[p.name] = extract_docx_markdown(p)
        elif suf == ".doc":
            tmp_dir = exp_dir / "源码" / "_build" / "inspect_tmp"
            converted = convert_doc_to_docx(p, tmp_dir, soffice)
            if converted:
                instruction_texts[p.name] = extract_docx_markdown(converted)
        elif suf == ".pdf":
            instruction_texts[p.name] = extract_pdf_markdown(p)

    if args.dump_instruction and instruction_texts:
        out_p = Path(args.dump_instruction).resolve()
        out_p.parent.mkdir(parents=True, exist_ok=True)
        out_p.write_text("\n\n---\n\n".join(instruction_texts.values()), encoding="utf-8")

    data_reports = [inspect_data_file(p) for p in files["data_files"]]

    combined_inst = "\n".join(instruction_texts.values())
    hard_flags = {
        "mentions_ipynb": bool(re.search(r"\.ipynb|jupyter\s*notebook", combined_inst, re.I)),
        "mentions_handcrafted": bool(re.search(r"手写|手动|自写|自行实现|不得直接调用|不要直接调用|不调用|numpy.*实现|原生.*实现", combined_inst, re.I)),
        "mentions_extension": bool(re.search(r"拓展|思考题|选做|课后", combined_inst)),
    }

    summary = {
        "exp_dir": str(exp_dir),
        "toolchain": toolchain,
        "profile_found": str(profile_path) if profile_path else None,
        "instructions": [p.name for p in files["instructions"]],
        "masters": [p.name for p in files["masters"]],
        "existing_reports": [p.name for p in files["existing_reports"]],
        "hard_flags": hard_flags,
        "data_files": data_reports,
        "dumped_instruction": args.dump_instruction if instruction_texts else None,
    }

    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return 0

    print("=" * 70)
    print(f"实验目录勘察：{exp_dir}")
    print("=" * 70)
    print(f"OS / Python: {toolchain['os']} | {toolchain['python']} ({toolchain['python_version']})")
    print(f"soffice:     {toolchain['soffice'] or '未找到'} | 宋体字体: {'✔' if toolchain['has_songti_font'] else '✖'}")
    missing_mods = [k for k, ok in toolchain["modules"].items() if not ok]
    print(f"Python 依赖: {'全部就绪' if not missing_mods else '缺少 ' + ', '.join(missing_mods)}")
    print(f"个人配置:    {summary['profile_found'] or '未找到（需向同学确认姓名/学号/班级等或创建 profile.json）'}")
    print(f"指导书:      {summary['instructions'] or '未找到'}")
    print(f"报告母版:    {summary['masters'] or '未找到（可使用 --standalone 自动生成标准版式）'}")
    print(f"已有成品:    {summary['existing_reports'] or '无'}")
    print(
        f"硬要求扫描:  ipynb={'是' if hard_flags['mentions_ipynb'] else '否'} | "
        f"手写算法={'是' if hard_flags['mentions_handcrafted'] else '否'} | "
        f"课后拓展/思考={'是' if hard_flags['mentions_extension'] else '否'}"
    )
    if args.dump_instruction and instruction_texts:
        print(f"指导书已提取至: {args.dump_instruction}")

    if data_reports:
        print("\n── 数据文件探查 ──")
        for d in data_reports:
            print(f"\n• {d['file']} ({d['size_kb']} KB, encoding={d.get('encoding', '-')})")
            if "shape" in d:
                print(f"  shape={d['shape']}, duplicates={d['duplicates']}, skiprows={d.get('skiprows_used', 0)}")
                for h in d.get("hints", []):
                    print(f"  [提示] {h}")
                for c in d.get("columns", []):
                    print(
                        f"    - {c['name']:<20} {c['dtype']:<10} "
                        f"missing={c['missing']:<4} nunique={c['nunique']:<5} "
                        f"samples={c['sample_values'][:4]}"
                    )
            elif "read_error" in d:
                print(f"  [读取失败] {d['read_error']}")
                for ln in d.get("raw_first_3_lines", []):
                    print(f"    raw: {ln[:100]}")

    if not args.dump_instruction and instruction_texts:
        print("\n── 指导书全文 ──\n")
        for txt in instruction_texts.values():
            print(txt)

    return 0


if __name__ == "__main__":
    sys.exit(main())
