#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通用实验报告 Word / PDF 排版装配与多遍自检脚本。

特性
----
1. 跨平台（macOS / Windows / Linux）：自动定位 LibreOffice（soffice）与当前 Python。
2. 零硬编码个人信息：从 `profile.json` + `report_config.json` + CLI 参数合并读取。
3. 双模式支持：
   - 模板模式（默认）：自动匹配实验目录下的 `.doc` / `.docx` 报告母版，保留学校原版式。
   - 独立模式（`--standalone` 或目录无母版时自动降级）：从零生成标准 A4 封面 + 5×4 实验报告表格。
4. 工程级排版修复：
   - 清理 LibreOffice `宋体;SimSun` 非法字体名；将 `²` 转真实上标避免全角空隙；
   - 封面下划线制表位右端齐平，并自动清理封面尾部多余空段以彻底消除「空白第 2 页」；
   - 正文 12pt 宋体 + 固定 24pt 行距（Word 与 LibreOffice 渲染一致，不压字、不裁图）；
   - 三线表自适应防断字列宽 + 不可跨页容器锁住表题与表体 + 宽表自动降字号与首列纵向合并；
   - 多遍 PDF 自检（最多 5 遍）：自动检查越界、下划线齐平、字体回退、空白页、图表题分离、
     图表齐全性、跨页断表、数字折断与大块留白，自动缩图（下限 12cm）或分页重排。

用法::
    # 1) 正式构建 + 多遍排版自检 + 导出同名 PDF
    python3 build_report.py --root 实验十 --check

    # 2) 指定配置文件与个人档案
    python3 build_report.py --root 实验十 --config 实验十/源码/report_config.json --profile profile.json --check

    # 3) 自测流水线（无需真实 section*.md，生成 demo_report.docx）
    python3 build_report.py --root 实验十 --demo --check
    python3 build_report.py --standalone --demo --check

    # 4) 仅对已有 docx 做几何与视觉自检
    python3 build_report.py --root 实验十 --check-only --docx path/to/report.docx
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import (
    WD_ALIGN_PARAGRAPH,
    WD_BREAK,
    WD_LINE_SPACING,
    WD_TAB_ALIGNMENT,
)
from docx.image.image import Image as DocxImage
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ---------------------------------------------------------------- 排版常量

FONT_CN = "宋体"
FONT_EN = "Times New Roman"
FONT_MONO = "Consolas"

SZ_BODY = Pt(12)          # 正文
SZ_H1 = Pt(14)            # 一、实验目的 / 二、实验过程 / 三、实验小结
SZ_H2 = Pt(12)            # ## 二级标题（题目一/环境与数据集）
SZ_H3 = Pt(12)            # ### 三级小标题
SZ_CAPTION = Pt(10.5)     # 图题 / 表题
SZ_TABLE = Pt(10.5)       # 表格正文
SZ_CODE = Pt(9)           # 代码/输出围栏块
SZ_COVER = Pt(16)         # 封面字段

FIG_TARGET_W_CM = 14.2    # 图片目标宽度
FIG_MAX_H_CM = 19.0       # 图片最大高度
CODE_SHADE = "F2F2F2"     # 围栏块浅灰底纹
GROUP_COLS = {"特征方案", "分组", "方案", "类别", "分群", "模型组", "组别"}
LINE_15 = 1.5             # XML 先写 360 auto，normalize_for_word() 统一固定为 480 twips (24pt)

BIG_TITLES = {2: "一、实验目的", 3: "二、实验过程", 4: "三、实验小结"}

DEFAULT_COVER_ORDER = [
    ("course", "课程名称"),
    ("major", "专　　业"),
    ("class", "班　　级"),
    ("student_id", "学　　号"),
    ("name", "姓　　名"),
    ("teacher", "指导教师"),
]

DEMO_COVER = {
    "course": "机器学习原理与应用",
    "major": "计算机科学与技术",
    "class": "计算机科学与技术 1 班",
    "student_id": "202500000001",
    "name": "张同学",
    "teacher": "李老师",
}

DEFAULT_TABLE_TOTAL_TW = 8522
DEFAULT_TABLE_COLS_TW = [2130, 2130, 1638, 2624]

# 运行态全局上下文（由 load_runtime_config 初始化）
CFG: dict = {}
BUILD_INFO: dict = {}
TABLE_INFO: list = []
report_anomalies_warnings: list[str] = []


# ---------------------------------------------------------------- 跨平台与配置工具

def log(msg: str = "") -> None:
    print(msg, flush=True)


def die(msg: str, code: int = 1):
    log("")
    log("!" * 70)
    log("构建失败：" + msg)
    log("!" * 70)
    sys.exit(code)


def detect_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        hit = shutil.which(name)
        if hit:
            return hit
    for c in (
        "/opt/homebrew/bin/soffice",
        "/usr/local/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
    ):
        if Path(c).exists():
            return c
    return None


def norm_key(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def relp(p: Path, root: Path | None = None) -> str:
    base = root or CFG.get("root")
    if base is not None:
        try:
            return str(Path(p).relative_to(base))
        except ValueError:
            pass
    return str(p)


def read_json_if_exists(path: Path | None) -> dict:
    if path is None or not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        die(f"无法解析 JSON 配置文件 {path}：{exc}")


def deep_merge(base: dict, override: dict) -> dict:
    out = dict(base)
    for k, v in override.items():
        if v is None:
            continue
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def find_master_doc(root: Path, configured: str | None) -> Path | None:
    if configured:
        p = Path(configured)
        if not p.is_absolute():
            p = root / p
        if not p.exists():
            die(f"配置指定的母版文件不存在：{p}")
        return p
    if not root.exists():
        return None
    candidates: list[Path] = []
    for p in sorted(root.iterdir()):
        if not p.is_file() or p.name.startswith((".", "~$")):
            continue
        if p.suffix.lower() not in (".doc", ".docx"):
            continue
        if re.match(r"^\d{6,}-", p.name):
            continue
        if any(k in p.name for k in ("学号", "模板", "母版")) or (
            "实验报告" in p.name and ("姓名" in p.name or p.suffix.lower() == ".doc")
        ):
            candidates.append(p)
    # 优先 .docx，其次 .doc
    candidates.sort(key=lambda x: (0 if x.suffix.lower() == ".docx" else 1, x.name))
    return candidates[0] if candidates else None


def resolve_section_paths(root: Path, cfg: dict) -> dict[str, list[Path] | Path]:
    src_dir = root / "源码"
    if not src_dir.is_dir():
        src_dir = root

    def _p(name: str) -> Path:
        p = Path(name)
        return p if p.is_absolute() else (src_dir / p)

    intro = _p(cfg.get("intro_section", "section0.md"))
    summary = _p(cfg.get("summary_section", "section3.md"))

    body_list = cfg.get("body_sections")
    if body_list:
        bodies = [_p(x) for x in body_list]
    elif cfg.get("sections"):
        all_secs = [_p(x) for x in cfg["sections"]]
        if len(all_secs) < 2:
            die("config.sections 至少需要包含 section0（目的与环境）和 1 个正文/小结文件")
        intro = all_secs[0]
        summary = all_secs[-1]
        bodies = all_secs[1:-1]
    else:
        # 自动发现 section1.md, section2.md, ...（排除 intro 与 summary）
        discovered = sorted(
            p for p in src_dir.glob("section*.md")
            if p.resolve() not in {intro.resolve(), summary.resolve()}
        )
        bodies = discovered if discovered else [_p("section1.md"), _p("section2.md")]

    return {"intro": intro, "bodies": bodies, "summary": summary}


def build_cover_fields(cover_cfg: dict) -> list[tuple[str, str]]:
    fields: list[tuple[str, str]] = []
    if isinstance(cover_cfg, list):
        for item in cover_cfg:
            if isinstance(item, (list, tuple)) and len(item) == 2:
                fields.append((str(item[0]), str(item[1])))
            elif isinstance(item, dict) and "label" in item and "value" in item:
                fields.append((str(item["label"]), str(item["value"])))
        return fields
    for key, label in DEFAULT_COVER_ORDER:
        if key in cover_cfg and cover_cfg[key] is not None:
            fields.append((label, str(cover_cfg[key])))
    return fields


def init_runtime(args) -> dict:
    if args.root:
        root = Path(args.root).resolve()
    elif args.config:
        cp = Path(args.config).resolve()
        root = cp.parent.parent if cp.parent.name == "源码" else cp.parent
    else:
        cwd = Path.cwd().resolve()
        root = cwd.parent if cwd.name == "源码" else cwd

    # 查找 profile.json（低优先级）
    profile_path = Path(args.profile).resolve() if args.profile else None
    if profile_path is None:
        for cand in (
            root / "profile.json",
            root.parent / "profile.json",
            Path.home() / ".config/lab-report/profile.json",
        ):
            if cand.exists():
                profile_path = cand
                break

    # 查找 report_config.json（高优先级）
    config_path = Path(args.config).resolve() if args.config else None
    if config_path is None:
        for cand in (root / "源码/report_config.json", root / "report_config.json"):
            if cand.exists():
                config_path = cand
                break

    merged = deep_merge(read_json_if_exists(profile_path), read_json_if_exists(config_path))

    # CLI 覆盖
    if args.exp_num is not None:
        merged["exp_num"] = args.exp_num
    if args.exp_label:
        merged["exp_label"] = args.exp_label
    if args.exp_title:
        merged["exp_title"] = args.exp_title
    if args.date:
        merged["date"] = args.date
    if args.group:
        merged["group"] = args.group
    if args.master:
        merged["master_doc"] = args.master
    if args.standalone:
        merged["template_mode"] = "standalone"

    today_cn = f"{dt.date.today().year}年{dt.date.today().month}月{dt.date.today().day}日"
    exp_num = merged.get("exp_num")
    exp_label = merged.get("exp_label") or (f"实验{exp_num}" if exp_num is not None else "实验")
    exp_title = merged.get("exp_title") or "实验报告"
    date_str = merged.get("date") or today_cn
    group_str = merged.get("group") or "单人组"

    cover_cfg = merged.get("cover") or {}
    if args.demo and not cover_cfg:
        cover_cfg = dict(DEMO_COVER)
    cover_fields = build_cover_fields(cover_cfg)
    if not cover_fields:
        if args.demo or args.check_only:
            cover_cfg = dict(DEMO_COVER)
            cover_fields = build_cover_fields(cover_cfg)
        else:
            die(
                "缺少封面个人信息（cover）。请在 report_config.json 或 profile.json 中填写 "
                "course / major / class / student_id / name / teacher。"
            )

    # 非 demo 模式下拒绝未替换的占位符
    if not args.demo and not args.check_only:
        bad_vals = [f"{lbl}={val}" for lbl, val in cover_fields if "请填写" in val or not val.strip()]
        if bad_vals:
            die("封面字段仍包含未填写的占位符，请先向同学确认后更新配置：\n  - " + "\n  - ".join(bad_vals))

    template_mode = merged.get("template_mode", "auto")
    master_doc = None if template_mode == "standalone" else find_master_doc(root, merged.get("master_doc"))
    if master_doc is None:
        template_mode = "standalone"

    src_dir = root / "源码"
    build_dir = (src_dir / "_build") if src_dir.is_dir() else (root / "_build")

    # 输出文件名
    if args.out:
        final_docx = Path(args.out).resolve()
    elif merged.get("output_docx"):
        out_p = Path(merged["output_docx"])
        final_docx = out_p if out_p.is_absolute() else (root / out_p)
    else:
        sid = cover_cfg.get("student_id", "学号") if isinstance(cover_cfg, dict) else "学号"
        sname = cover_cfg.get("name", "姓名") if isinstance(cover_cfg, dict) else "姓名"
        pattern = merged.get("output_pattern", "{student_id}-{name}-{exp_label}-《{exp_title}》实验报告.docx")
        fname = pattern.format(
            student_id=sid,
            name=sname,
            exp_label=exp_label,
            exp_title=exp_title,
            exp_num=exp_num if exp_num is not None else "",
        )
        final_docx = root / fname

    # 图表编号匹配前缀
    num_pat = rf"{re.escape(str(exp_num))}-(\d+)" if exp_num is not None else r"(?:\d+-)?(\d+)"

    runtime = {
        "root": root,
        "build": build_dir,
        "profile_path": profile_path,
        "config_path": config_path,
        "soffice": detect_soffice(),
        "exp_num": exp_num,
        "exp_label": exp_label,
        "exp_title": exp_title,
        "date": date_str,
        "group": group_str,
        "cover_fields": cover_fields,
        "template_mode": template_mode,
        "master_doc": master_doc,
        "final_docx": final_docx,
        "sections": resolve_section_paths(root, merged),
        "expect_fig": merged.get("expect_fig"),
        "expect_tab": merged.get("expect_tab"),
        "fig_num_re": re.compile(rf"^图\s*{num_pat}(?:\s|$)"),
        "tab_num_re": re.compile(rf"^表\s*{num_pat}(?:\s|$)"),
        "any_cap_re": re.compile(rf"^([图表])\s*{num_pat}"),
        "inline_ref_re": re.compile(rf"([图表])\s*{num_pat}"),
        "no_pdf": bool(args.no_pdf),
    }
    CFG.clear()
    CFG.update(runtime)
    return runtime


# ---------------------------------------------------------------- 文本与样式工具

def display_em(s: str) -> float:
    w = 0.0
    for ch in s:
        if ch == "\u3000":
            w += 1.0
        elif unicodedata.east_asian_width(ch) in ("W", "F"):
            w += 1.0
        elif ch.isascii():
            w += 0.5
        else:
            w += 1.0
    return w


def set_run(run, *, size=None, bold=None, mono=False, cn=FONT_CN, en=FONT_EN,
            underline=None, color=None):
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    if mono:
        rFonts.set(qn("w:ascii"), FONT_MONO)
        rFonts.set(qn("w:hAnsi"), FONT_MONO)
        rFonts.set(qn("w:cs"), FONT_MONO)
        rFonts.set(qn("w:eastAsia"), FONT_MONO)
    else:
        rFonts.set(qn("w:ascii"), en)
        rFonts.set(qn("w:hAnsi"), en)
        rFonts.set(qn("w:cs"), en)
        rFonts.set(qn("w:eastAsia"), cn)
    rFonts.set(qn("w:hint"), "eastAsia")
    if size is not None:
        run.font.size = size
        szCs = rPr.find(qn("w:szCs"))
        if szCs is None:
            szCs = OxmlElement("w:szCs")
            rPr.append(szCs)
        szCs.set(qn("w:val"), str(int(size.pt * 2)))
    if bold is not None:
        run.font.bold = bold
    if underline is not None:
        run.font.underline = underline
    if color is not None:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def add_text(par, text, *, size, bold=None, mono=False, cn=FONT_CN, en=FONT_EN,
             underline=None, color=None):
    """写一段文本；把 U+00B2（²）拆成真正的上标 2，避免 SimSun/Times 出现全角空隙。"""
    chunks = str(text).split("\u00b2")
    for i, chunk in enumerate(chunks):
        if chunk:
            set_run(par.add_run(chunk), size=size, bold=bold, mono=mono, cn=cn, en=en,
                    underline=underline, color=color)
        if i < len(chunks) - 1:
            run = par.add_run("2")
            set_run(run, size=size, bold=bold, mono=mono, cn=cn, en=en,
                    underline=underline, color=color)
            run.font.superscript = True


def set_indent(par, *, first_line_chars=None, first_line_pt=None,
               left_cm=None, right_cm=0.0):
    pPr = par._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    for attr in ("w:firstLine", "w:firstLineChars", "w:start", "w:startChars",
                 "w:left", "w:end", "w:endChars", "w:right"):
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]
    if left_cm is not None:
        ind.set(qn("w:start"), str(int(Cm(left_cm).twips)))
        ind.set(qn("w:startChars"), "0")
    if right_cm:
        ind.set(qn("w:end"), str(int(Cm(right_cm).twips)))
        ind.set(qn("w:endChars"), "0")
    if first_line_chars is not None:
        ind.set(qn("w:firstLineChars"), str(int(first_line_chars * 100)))
        ind.set(qn("w:firstLine"), str(int((first_line_chars * SZ_BODY.pt) * 20)))
    elif first_line_pt is not None:
        ind.set(qn("w:firstLine"), str(int(first_line_pt * 20)))
        ind.set(qn("w:firstLineChars"), "0")
    return ind


def disable_grid_snap(par):
    """关掉「对齐到文档网格」，避免母版 docGrid 把 1.5 倍行距拉散。"""
    pPr = par._p.get_or_add_pPr()
    el = pPr.find(qn("w:snapToGrid"))
    if el is None:
        el = OxmlElement("w:snapToGrid")
        succ = [qn("w:ind"), qn("w:jc"), qn("w:rPr"), qn("w:sectPr")]
        idx = len(pPr)
        for i, child in enumerate(pPr):
            if child.tag in succ:
                idx = i
                break
        pPr.insert(idx, el)
    el.set(qn("w:val"), "0")


def shade_paragraph(par, fill: str):
    pPr = par._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def _border_el(tag: str, sz: int, val="single"):
    el = OxmlElement(tag)
    if val == "none":
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
    else:
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
    return el


def set_table_three_line(table, header_rows: int = 1):
    """三线表：上下框 1.5pt，表头下细线 0.5pt，无竖线与内部横线。"""
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    borders.append(_border_el("w:top", 12))
    borders.append(_border_el("w:left", 0, "none"))
    borders.append(_border_el("w:bottom", 12))
    borders.append(_border_el("w:right", 0, "none"))
    borders.append(_border_el("w:insideH", 0, "none"))
    borders.append(_border_el("w:insideV", 0, "none"))
    tblPr.append(borders)
    for cell in table.rows[header_rows - 1].cells:
        tcPr = cell._tc.get_or_add_tcPr()
        old_tc = tcPr.find(qn("w:tcBorders"))
        if old_tc is not None:
            tcPr.remove(old_tc)
        tcB = OxmlElement("w:tcBorders")
        tcB.append(_border_el("w:bottom", 4))
        tcPr.append(tcB)


def set_table_outer_grid(table, sz: int = 4):
    """标准实线网格（供独立模式生成外层 5×4 实验报告框线）。"""
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for tag in ("w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"):
        borders.append(_border_el(tag, sz))
    tblPr.append(borders)


def set_row_cant_split(row, header: bool = False):
    trPr = row._tr.get_or_add_trPr()
    for tag in ("w:cantSplit", "w:tblHeader"):
        el = trPr.find(qn(tag))
        if el is not None:
            trPr.remove(el)
    trPr.append(OxmlElement("w:cantSplit"))
    if header:
        trPr.append(OxmlElement("w:tblHeader"))


def set_table_fixed(table, widths_twips: list[int]):
    tbl = table._tbl
    tblPr = tbl.tblPr
    old_layout = tblPr.find(qn("w:tblLayout"))
    if old_layout is not None:
        tblPr.remove(old_layout)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(sum(widths_twips)))
    ind = tblPr.find(qn("w:tblInd"))
    if ind is None:
        ind = OxmlElement("w:tblInd")
        tblPr.append(ind)
    ind.set(qn("w:type"), "dxa")
    ind.set(qn("w:w"), "0")

    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths_twips:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w)))
        grid.append(gc)
    tblPr.addnext(grid)

    for row in table.rows:
        for cell, w in zip(row.cells, widths_twips):
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.insert(0, tcW)
            tcW.set(qn("w:type"), "dxa")
            tcW.set(qn("w:w"), str(int(w)))


def text_units(s: str, size_pt: float) -> float:
    w = 0.0
    for ch in s:
        if ch == "\u3000" or unicodedata.east_asian_width(ch) in ("W", "F"):
            w += size_pt
        elif ch.isdigit():
            w += size_pt * 0.5
        elif ch in ".,:;'\"":
            w += size_pt * 0.28
        elif ch == "%":
            w += size_pt * 0.83
        elif ch in "-—–":
            w += size_pt * 0.35
        elif ch == "~":
            w += size_pt * 0.5
        elif ch == " ":
            w += size_pt * 0.25
        elif ch == "/":
            w += size_pt * 0.28
        elif ch in "()[]":
            w += size_pt * 0.33
        elif ch.isupper():
            w += size_pt * 0.63
        elif ch.isascii():
            w += size_pt * 0.5
        else:
            w += size_pt
    return w


def unbreakable_units(s: str, size_pt: float) -> float:
    """最长不可断片段宽度（pt）：避免长数字或拉丁词被从中间劈开。"""
    best = 0.0
    cur = ""
    for ch in s:
        if ch == "\u3000" or unicodedata.east_asian_width(ch) in ("W", "F"):
            best = max(best, text_units(cur, size_pt), text_units(ch, size_pt))
            cur = ""
        elif ch in "/ ":
            best = max(best, text_units(cur, size_pt))
            cur = ""
        else:
            cur += ch
    return max(best, text_units(cur, size_pt))


def table_layout(rows: list[list[str]], total_tw: int, size_pt: float,
                 margin_tw: int) -> tuple[list[int], list[float], bool]:
    ncol = max(len(r) for r in rows)
    text_total = total_tw - ncol * 2 * margin_tw
    avail_pt = text_total / 20.0
    need, want = [], []
    for c in range(ncol):
        cells = [r[c] for r in rows if c < len(r)]
        n = max(unbreakable_units(v, size_pt) for v in cells) + size_pt * 0.45
        w = max(text_units(v, size_pt) for v in cells) + size_pt * 0.6
        need.append(n)
        want.append(min(max(w, n), avail_pt * 0.60))
    sum_need = sum(need)
    overflow = sum_need > avail_pt
    if overflow:
        widths = [v * avail_pt / sum_need for v in need]
    else:
        widths = list(need)
        gap = avail_pt - sum(widths)
        rest = []
        for i in sorted(range(ncol), key=lambda j: want[j] - widths[j]):
            shortfall = want[i] - widths[i]
            if 0.05 < shortfall <= size_pt * 3.2 and shortfall <= gap + 0.05:
                widths[i] += shortfall
                gap -= shortfall
            elif shortfall > 0.05:
                rest.append(i)
        if gap > 0.05 and rest:
            span = sum(want[i] - widths[i] for i in rest) or 1.0
            k = min(1.0, gap / span)
            for i in rest:
                widths[i] += (want[i] - widths[i]) * k
            gap = avail_pt - sum(widths)
        if gap > 0.5:
            tot = sum(widths) or 1.0
            widths = [w + gap * w / tot for w in widths]
    min_tw = int(Cm(0.8).twips)
    tw = [max(min_tw, int(round(w / avail_pt * text_total))) for w in widths]
    diff = text_total - sum(tw)
    i = 0
    guard = 0
    while diff != 0 and guard < 400000:
        guard += 1
        j = i % ncol
        step = 1 if diff > 0 else -1
        if tw[j] + step >= min_tw:
            tw[j] += step
            diff -= step
        i += 1
    return [w + 2 * margin_tw for w in tw], need, overflow


def set_cell_margins(table, margin_tw: int):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblCellMar"))
    if old is not None:
        tblPr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for side in ("left", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(margin_tw))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def set_cell_nowrap(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    if tcPr.find(qn("w:noWrap")) is None:
        tcPr.append(OxmlElement("w:noWrap"))


def merge_first_col(table, col: int = 0) -> int:
    n = 0
    vals = [table.cell(k, col).text.strip() for k in range(len(table.rows))]
    i = 1
    while i < len(vals):
        if not vals[i] or vals[i] != vals[i - 1]:
            i += 1
            continue
        j = i
        while j < len(vals) and vals[j] == vals[i - 1]:
            j += 1
        for k in range(i, j):
            tc = table.cell(k, col)
            for par in tc.paragraphs[1:]:
                par._p.getparent().remove(par._p)
            p0 = tc.paragraphs[0]
            for r in list(p0.runs):
                p0._p.remove(r._r)
        top = table.cell(i - 1, col)
        top.merge(table.cell(j - 1, col))
        n += 1
        i = j
    return n


def keep_table_together(table):
    for row in table.rows[:-1]:
        for cell in row.cells:
            for par in cell.paragraphs:
                par.paragraph_format.keep_with_next = True


def set_table_borders_none(table):
    tblPr = table._tbl.tblPr
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for tag in ("w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"):
        borders.append(_border_el(tag, 0, "none"))
    tblPr.append(borders)
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            old_tc = tcPr.find(qn("w:tcBorders"))
            if old_tc is not None:
                tcPr.remove(old_tc)
            tcB = OxmlElement("w:tcBorders")
            for tag in ("w:top", "w:left", "w:bottom", "w:right"):
                tcB.append(_border_el(tag, 0, "none"))
            tcPr.append(tcB)


# ---------------------------------------------------------------- 母版准备与独立版式

def font_cleanup(doc):
    """清理 LibreOffice 转换生成的 `宋体;SimSun` 复合非法字体名。"""
    fixed = 0
    parts = [doc.element]
    for st in doc.styles:
        try:
            parts.append(st.element)
        except Exception:
            pass
    for root in parts:
        for el in root.iter(qn("w:rFonts")):
            for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
                v = el.get(qn(attr))
                if v and ";" in v:
                    el.set(qn(attr), v.split(";")[0].strip())
                    fixed += 1
    if fixed:
        log(f"  · 修正非法字体名 `xxx;yyy` → 单字体：{fixed} 处")


def create_standalone_doc() -> Document:
    """无母版时创建标准 A4 学术实验报告骨架（封面 + 5×4 主表）。"""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.175)
    sec.right_margin = Cm(3.175)

    # 封面标题与留白
    for _ in range(3):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(12)
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(12)
    title_p.paragraph_format.space_after = Pt(48)
    set_run(title_p.add_run("实 验 报 告"), size=Pt(32), bold=True)

    for lbl, _val in CFG["cover_fields"]:
        p = doc.add_paragraph()
        set_run(p.add_run(lbl), size=SZ_COVER)

    # 5×4 主表格（与高校常见实验报告模板结构一致）
    tbl = doc.add_table(rows=5, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_fixed(tbl, DEFAULT_TABLE_COLS_TW)
    set_table_outer_grid(tbl, sz=4)
    set_cell_margins(tbl, 108)

    # r0: 实验题目 | [合并 1..3]
    _set_cell_text(tbl.rows[0].cells[0], "实验题目", SZ_H1, center=True)
    tbl.rows[0].cells[1].merge(tbl.rows[0].cells[3])
    # r1: 实验组别 | 值 | 实验时间 | 值
    _set_cell_text(tbl.rows[1].cells[0], "实验组别", SZ_H1, center=True)
    _set_cell_text(tbl.rows[1].cells[2], "实验时间", SZ_H1, center=True)
    # r2..r4: 整行合并
    for ri in (2, 3, 4):
        tbl.rows[ri].cells[0].merge(tbl.rows[ri].cells[3])
    return doc


def load_base_document() -> Document:
    """加载母版（.docx 直接读取，.doc 通过 soffice 转换）或生成独立骨架。"""
    if CFG["template_mode"] == "standalone" or CFG["master_doc"] is None:
        log("  · 使用独立标准版式（standalone 模式）")
        return create_standalone_doc()

    master = CFG["master_doc"]
    if master.suffix.lower() == ".docx":
        log(f"  · 直接读取 .docx 母版：{relp(master)}")
        return Document(str(master))

    build_dir = CFG["build"]
    build_dir.mkdir(parents=True, exist_ok=True)
    master_docx = build_dir / (master.stem + ".docx")
    if master_docx.exists() and master_docx.stat().st_mtime >= master.stat().st_mtime:
        log(f"  · 复用已转换的母版：{relp(master_docx)}")
        return Document(str(master_docx))

    soffice = CFG["soffice"]
    if not soffice:
        die(
            f"母版是旧版 .doc 格式（{master.name}），需要 LibreOffice (soffice) 转换，但当前系统未检测到 soffice。\n"
            "  解决方式（二选一）：\n"
            "  1) 安装 LibreOffice（macOS: brew install --cask libreoffice；Windows: winget install TheDocumentFoundation.LibreOffice）\n"
            "  2) 在 Word/WPS 中把母版另存为 .docx 后重跑，或使用 --standalone 模式。"
        )
    log(f"  · soffice 转换母版 → {relp(master_docx)}")
    res = subprocess.run(
        [soffice, "--headless", "--convert-to", "docx", "--outdir", str(build_dir), str(master)],
        capture_output=True, text=True,
    )
    if not master_docx.exists():
        die("母版转换失败：\n" + res.stdout + res.stderr)
    return Document(str(master_docx))


# ---------------------------------------------------------------- 封面与表头

def _text_width_cm(doc) -> float:
    s = doc.sections[0]
    return (s.page_width - s.left_margin - s.right_margin) / 360000


def rebuild_cover(doc):
    """按配置重写封面字段段落：标签左对齐、值带下划线、下划线右端齐平。"""
    cover_fields: list[tuple[str, str]] = CFG["cover_fields"]
    picked: list[tuple] = []
    used_indices: set[int] = set()

    for lbl, val in cover_fields:
        clean_lbl = re.sub(r"[\s\u3000:：]+", "", lbl)
        for idx, p in enumerate(doc.paragraphs[:30]):
            if idx in used_indices:
                continue
            t = re.sub(r"[\s\u3000:：]+", "", p.text)
            if t.startswith(clean_lbl):
                picked.append((p, lbl, val))
                used_indices.add(idx)
                break

    if not picked:
        log("  ! 未在母版封面匹配到字段段落，跳过封面重写")
        return

    left_cm = 1.15
    gap_em = 0.5
    label_em = max(display_em(lbl) for _, lbl, _ in picked)
    target_em = max(display_em(val) for _, _, val in picked) + 1.5
    em_cm = SZ_COVER.pt * 2.54 / 72
    tab_cm = left_cm + (label_em + gap_em + target_em) * em_cm
    text_w_cm = _text_width_cm(doc)
    if tab_cm > text_w_cm - 0.2:
        tab_cm = text_w_cm - 0.2
        log(f"  ! 封面下划线右端受限，制表位收到 {tab_cm:.2f}cm")

    for par, lbl, value in picked:
        for r in list(par.runs):
            r._element.getparent().remove(r._element)
        pf = par.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.space_before = Pt(6)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        set_indent(par, left_cm=left_cm)
        pf.tab_stops.add_tab_stop(Cm(tab_cm), WD_TAB_ALIGNMENT.LEFT)

        set_run(par.add_run(lbl), size=SZ_COVER)
        set_run(par.add_run(" " * max(1, int(gap_em / 0.25))), size=SZ_COVER)
        pad_em = target_em - 1.5 - display_em(value)
        if pad_em < 0:
            log(f"  ! 警告：字段「{lbl}」的值过宽，下划线右端可能超出制表位")
            pad_em = 0
        set_run(par.add_run(value + "\u3000" * int(pad_em) + "\t"), size=SZ_COVER, underline=True)
    log(f"  · 封面 {len(picked)} 个字段已重写：起始 x={left_cm}cm，下划线右端制表位 {tab_cm:.2f}cm")


def fill_header_rows(tbl):
    """填充主表格前两行（实验题目 / 实验组别 / 实验时间），兼容单元格列数差异。"""
    if len(tbl.rows) < 2:
        return
    r0 = tbl.rows[0]
    if len(r0.cells) >= 2:
        _set_cell_text(r0.cells[1], CFG["exp_title"], SZ_H1)
    r1 = tbl.rows[1]
    if len(r1.cells) >= 4:
        _set_cell_text(r1.cells[1], CFG["group"], SZ_H1, center=True)
        _set_cell_text(r1.cells[3], CFG["date"], SZ_H1, center=True)
    elif len(r1.cells) >= 2:
        _set_cell_text(r1.cells[1], f"{CFG['group']} / {CFG['date']}", SZ_H1, center=True)
    log(f"  · 表格 r0/r1 已填：{CFG['exp_title']} | {CFG['group']} | {CFG['date']}")


# ---------------------------------------------------------------- 正文装配器

class ContentBuilder:
    """把解析出的内容块写入目标单元格。"""

    def __init__(self, doc, avail_tw: int, forced_breaks=None, rerun: int = 1,
                 shrink=None, keep_tables=None):
        self.doc = doc
        self.avail_tw = avail_tw
        self.n_para = self.n_fig = self.n_tab = self.n_code = 0
        self.warnings: list[str] = []
        self.forced_breaks = set(forced_breaks or ())
        self.shrink = dict(shrink or ())
        self.keep_tables = set(keep_tables or ())
        self.fig_nat: dict[str, tuple] = {}
        self.rerun = rerun

    def body_para(self, cell, text: str):
        p = cell.add_paragraph()
        disable_grid_snap(p)
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.line_spacing = LINE_15
        pf.space_before = Pt(0)
        pf.space_after = Pt(5)
        set_indent(p, first_line_chars=2)
        self._inline_runs(p, text, size=SZ_BODY)
        self.n_para += 1
        return p

    def heading(self, cell, text: str, level: int):
        p = cell.add_paragraph()
        disable_grid_snap(p)
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = LINE_15
        pf.space_before = Pt(16 if level <= 2 else 10)
        pf.space_after = Pt(6 if level <= 2 else 3)
        pf.keep_with_next = True
        set_indent(p, first_line_chars=0)
        add_text(p, text, size=SZ_H2 if level <= 2 else SZ_H3, bold=True)
        self.n_para += 1
        return p

    def _inline_runs(self, par, text: str, size):
        tokens = re.split(r"(\*\*[^*]+\*\*|`[^`]+`)", text)
        for tok in tokens:
            if not tok:
                continue
            if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
                add_text(par, tok[2:-2], size=size, bold=True)
            elif tok.startswith("`") and tok.endswith("`") and len(tok) > 2:
                add_text(par, tok[1:-1], size=size, mono=True)
            else:
                add_text(par, tok, size=size)

    def big_title(self, cell, text: str, keep_next: bool = False):
        p = cell.add_paragraph()
        disable_grid_snap(p)
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = LINE_15
        pf.space_before = Pt(10)
        pf.space_after = Pt(10)
        pf.keep_with_next = keep_next
        set_indent(p, first_line_chars=0)
        add_text(p, text, size=SZ_H1, bold=True)
        return p

    def figure(self, cell, paths, caption: str):
        if isinstance(paths, (str, Path)):
            paths = [paths]
        key = norm_key(caption)
        force = key in self.forced_breaks
        limit_w_cm = min(FIG_TARGET_W_CM, self.avail_tw / 1440 * 2.54)
        geo = []
        for path in paths:
            try:
                im = DocxImage.from_file(str(path))
            except Exception as exc:
                die(f"图片无法读取：{path}\n{exc}")
            nat_w, nat_h = im.width, im.height
            w_cm = limit_w_cm
            h_cm = nat_h / nat_w * w_cm
            if h_cm > FIG_MAX_H_CM:
                h_cm = FIG_MAX_H_CM
                w_cm = nat_w / nat_h * h_cm
            geo.append([path, w_cm, h_cm, nat_w, nat_h])
        self.fig_nat[key] = (geo[0][1], geo[0][2] / geo[0][1])

        target_w = self.shrink.get(key)
        if target_w and geo[0][1] > target_w + 0.05:
            f = target_w / geo[0][1]
            for g in geo:
                g[1] *= f
                g[2] *= f
        total_h = sum(g[2] for g in geo)
        if total_h > FIG_MAX_H_CM:
            f = FIG_MAX_H_CM / total_h
            for g in geo:
                g[1] *= f
                g[2] *= f

        for gi, (path, w_cm, h_cm, nat_w, nat_h) in enumerate(geo):
            p = cell.add_paragraph()
            disable_grid_snap(p)
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.line_spacing = 1.0
            pf.space_before = Pt(12) if gi == 0 else Pt(6)
            pf.space_after = Pt(4)
            pf.keep_with_next = True
            if force and gi == 0:
                pf.page_break_before = True
                br_run = p.add_run()
                br_run.add_break(WD_BREAK.PAGE)
                log(f"    * 第 {self.rerun} 遍：图「{caption[:18]}」强制另起一页")
            set_indent(p, first_line_chars=0)
            p.add_run().add_picture(str(path), width=Cm(w_cm), height=Cm(h_cm))
            log(f"    · 图 {caption[:26]:<28} {w_cm:.2f}×{h_cm:.2f}cm  {path.name}")
            if w_cm < 9.0:
                self.warnings.append(
                    f"图「{caption[:16]}」被 {FIG_MAX_H_CM:g}cm 高度上限压至 {w_cm:.1f}cm 宽"
                    f"，建议拆图或减少截图 cell 数：{path.name}"
                )
        self.caption(cell, caption, after=Pt(12), keep=False)
        self.n_fig += 1

    def caption(self, cell, text: str, before=Pt(10), after=Pt(6), keep=True):
        p = cell.add_paragraph()
        disable_grid_snap(p)
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.line_spacing = 1.0
        pf.space_before = before
        pf.space_after = after
        pf.keep_with_next = keep
        set_indent(p, first_line_chars=0)
        add_text(p, text, size=SZ_CAPTION)
        return p

    def table_container(self, cell):
        container = cell.add_table(rows=1, cols=1)
        set_table_fixed(container, [self.avail_tw])
        set_cell_margins(container, 0)
        set_table_borders_none(container)
        set_row_cant_split(container.rows[0])
        return container.cell(0, 0)

    def csv_table(self, cell, path: Path, caption: str):
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            rows = [[c.strip() for c in row] for row in csv.reader(f) if any(c.strip() for c in row)]
        if not rows:
            die(f"CSV 为空：{path}")
        ncol = max(len(r) for r in rows)
        rows = [r + [""] * (ncol - len(r)) for r in rows]
        n_data = len(rows) - 1
        key = norm_key(caption)
        # 短/中表整表放入 cantSplit 容器，防止表题孤悬页底；超长表（>22 行）允许跨页并重复表头
        keep_whole = (len(rows) <= 22) or (key in self.keep_tables)
        if key in self.forced_breaks:
            pb = cell.add_paragraph()
            disable_grid_snap(pb)
            pf = pb.paragraph_format
            pf.page_break_before = True
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            set_run(pb.add_run(""), size=Pt(1))
            log(f"    * 第 {self.rerun} 遍：表「{caption[:18]}」另起一页")
        host = self.table_container(cell) if keep_whole else cell
        cap_par = self.caption(host, caption, before=Pt(12), after=Pt(6), keep=True)
        if keep_whole:
            for p0 in list(host.paragraphs):
                if p0._p is cap_par._p:
                    break
                if not p0.text.strip():
                    p0._p.getparent().remove(p0._p)

        size_pt, margin_tw = (9.0, int(Cm(0.06).twips)) if ncol >= 7 else (10.5, 108)
        widths, _need_pt, overflow = table_layout(rows, self.avail_tw, size_pt, margin_tw)
        if overflow:
            size_pt, margin_tw = (8.5, int(Cm(0.05).twips))
            widths, _need_pt, overflow = table_layout(rows, self.avail_tw, size_pt, margin_tw)
            self.warnings.append(f"表「{caption[:20]}」{ncol} 列在 9pt 下仍放不下，已降到 8.5pt")

        table = host.add_table(rows=len(rows), cols=ncol)
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_table_fixed(table, widths)
        set_cell_margins(table, margin_tw)
        set_table_three_line(table, header_rows=1)

        numeric_cols = [
            c for c in range(ncol)
            if all(re.fullmatch(r"[-+]?[\d.,%~/]+", r[c]) for r in rows[1:] if r[c].strip())
        ]
        for ri, row in enumerate(rows):
            set_row_cant_split(table.rows[ri], header=(ri == 0))
            for ci, val in enumerate(row):
                tc = table.cell(ri, ci)
                tc.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                par = tc.paragraphs[0]
                disable_grid_snap(par)
                pf = par.paragraph_format
                pf.line_spacing = 1.15
                pf.space_before = Pt(3)
                pf.space_after = Pt(3)
                set_indent(par, first_line_chars=0)
                header = (ri == 0)
                col_max = max(display_em(r[ci]) for r in rows[1:]) if len(rows) > 1 else 0
                pf.alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER if header or col_max <= 8 else WD_ALIGN_PARAGRAPH.LEFT
                )
                add_text(par, val, size=Pt(size_pt), bold=header)
                if ri > 0 and ci in numeric_cols and val.strip():
                    set_cell_nowrap(tc)

        group_head = ncol >= 6 and (rows[0][0].strip() in GROUP_COLS)
        n_merged = merge_first_col(table, 0) if (n_data >= 4 and group_head) else 0
        if keep_whole:
            keep_table_together(table)
            compact_empty_paragraphs(host)
        self.n_tab += 1
        TABLE_INFO.append({
            "cap": caption, "key": key, "ncol": ncol, "nrows": len(rows),
            "size": size_pt, "width_pt": sum(widths) / 20.0, "margin_tw": margin_tw,
            "rows": rows, "keep": keep_whole,
        })
        extra = (f"，{size_pt:g}pt" + (f"，合并首列 {n_merged} 处" if n_merged else "")) \
            if (size_pt != 10.5 or n_merged) else ""
        log(f"    · 表 {caption[:26]:<28} {len(rows)}×{ncol}{extra}  {path.name}")
        return table

    def code_block(self, cell, lines: list[str]):
        p = cell.add_paragraph()
        disable_grid_snap(p)
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = 1.0
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.space_before = Pt(8)
        pf.space_after = Pt(10)
        pf.keep_together = len(lines) <= 10
        set_indent(p, first_line_chars=0, left_cm=0.2, right_cm=0.2)
        shade_paragraph(p, CODE_SHADE)
        for i, line in enumerate(lines):
            r = p.add_run(line)
            set_run(r, size=SZ_CODE, mono=True)
            if i != len(lines) - 1:
                r.add_break()
        self.n_code += 1
        for line in lines:
            if display_em(line) > 72:
                self.warnings.append(f"围栏块有超长行（{display_em(line):.0f} 字符宽），可能折行：{line[:40]}…")


# ---------------------------------------------------------------- Markdown 解析

FIG_RE = re.compile(r"^\{\{FIG:(?P<body>.+?)\}\}\s*$")
TAB_RE = re.compile(r"^\{\{TAB:(?P<path>[^|]+)\|(?P<cap>.*)\}\}\s*$")
OUT_RE = re.compile(r"^\{\{OUT:(?P<txt>.*)\}\}\s*$")
HEAD_RE = re.compile(r"^(?P<hashes>#{1,6})\s+(?P<text>.+?)\s*#*\s*$")
OTHER_MARK_RE = re.compile(r"^\{\{(?!FIG:|TAB:|OUT:)(?P<name>[A-Za-z_]+):")
SENT_END = "。！？：；…」』）\"”'’.!?:;)]}"


def resolve_asset(root: Path, rel: str, kind: str) -> Path | None:
    rel = rel.strip().strip("`").strip()
    p_rel = Path(rel)
    if p_rel.is_absolute() and p_rel.exists():
        return p_rel
    cands = [root / rel, root / "figs" / rel, root / "源码" / rel]
    if kind == "TAB":
        cands.append(root / "figs" / "tables" / p_rel.name)
    for c in cands:
        if c.exists():
            return c
    return None


def parse_md(text: str, src: str) -> list[dict]:
    blocks: list[dict] = []
    lines = text.replace("\r\n", "\n").split("\n")
    i = 0
    in_fence = False
    fence_lines: list[str] = []
    while i < len(lines):
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()
        if in_fence:
            if stripped.startswith("```"):
                blocks.append({"type": "code", "lines": fence_lines, "src": src})
                fence_lines = []
                in_fence = False
            else:
                fence_lines.append(raw)
            i += 1
            continue
        if stripped.startswith("```"):
            in_fence = True
            fence_lines = []
            i += 1
            continue
        if not stripped or stripped.startswith("<!--"):
            i += 1
            continue
        m = HEAD_RE.match(stripped)
        if m:
            blocks.append({"type": "heading", "level": len(m.group("hashes")),
                           "text": m.group("text").strip(), "src": src})
            i += 1
            continue
        m = FIG_RE.match(stripped)
        if m:
            parts = [x.strip() for x in m.group("body").split("|")]
            paths = [x for x in parts[:-1] if x]
            if not paths:
                die(f"{{{{FIG:}}}} 缺少图片路径，出现在 {Path(src).name}：{stripped}")
            if len(paths) > 2:
                die(f"{{{{FIG:}}}} 最多支持 2 张图共用一个图题：{stripped}")
            blocks.append({"type": "fig", "paths": paths,
                           "cap": parts[-1] if len(parts) > 1 else "", "src": src})
            i += 1
            continue
        m = TAB_RE.match(stripped)
        if m:
            blocks.append({"type": "tab", "path": m.group("path").strip(),
                           "cap": m.group("cap").strip(), "src": src})
            i += 1
            continue
        m = OTHER_MARK_RE.match(stripped)
        if m:
            blocks.append({"type": "bad_mark", "text": stripped, "name": m.group("name"), "src": src})
            i += 1
            continue
        m = OUT_RE.match(stripped)
        if m:
            blocks.append({"type": "out_legacy", "text": m.group("txt").strip(), "src": src})
            i += 1
            continue
        blocks.append({"type": "para", "text": stripped, "src": src})
        i += 1
    if in_fence:
        blocks.append({"type": "code", "lines": fence_lines, "src": src, "unclosed": True})
    return blocks


def split_section0(text: str) -> tuple[str, str]:
    lines = text.replace("\r\n", "\n").split("\n")
    idx_purpose = idx_env = None
    for i, ln in enumerate(lines):
        m = re.match(r"^(#{1,6})\s*(.+?)\s*$", ln.strip())
        if not m:
            continue
        title = m.group(2)
        if idx_purpose is None and "实验目的" in title:
            idx_purpose = i
        elif idx_env is None and ("环境" in title or "数据集" in title):
            idx_env = i
    if idx_purpose is None:
        die("section0.md 中找不到「## 实验目的」标题，无法填充「一、实验目的」")
    if idx_env is None:
        die("section0.md 中找不到「## 实验环境与数据集」标题，无法填充「二、实验过程」")
    return "\n".join(lines[idx_purpose + 1: idx_env]), "\n".join(lines[idx_env:])


def strip_leading_title(text: str, keywords: tuple[str, ...] = ("实验小结", "实验总结", "总结")) -> str:
    lines = text.replace("\r\n", "\n").split("\n")
    for i, ln in enumerate(lines):
        if not ln.strip():
            continue
        m = re.match(r"^(#{1,6})\s*(.+?)\s*$", ln.strip())
        if m and any(k in m.group(2) for k in keywords):
            return "\n".join(lines[i + 1:])
        return text
    return text


# ---------------------------------------------------------------- 组装流程

def load_content(demo: bool) -> dict:
    if demo:
        return make_demo_sections()
    secs = CFG["sections"]
    needed = [secs["intro"], *secs["bodies"], secs["summary"]]
    missing = [relp(p) for p in needed if not p.exists()]
    if missing:
        die("缺少章节文本文件（不会用假内容交付）：\n  - " + "\n  - ".join(missing))

    def _read(p: Path) -> str:
        txt = p.read_text(encoding="utf-8")
        if not txt.strip():
            die(f"章节文件为空：{relp(p)}")
        log(f"  · 读取 {relp(p)}（{len(txt)} 字符）")
        return txt

    return {
        "intro": (_read(secs["intro"]), str(secs["intro"])),
        "bodies": [(_read(p), str(p)) for p in secs["bodies"]],
        "summary": (_read(secs["summary"]), str(secs["summary"])),
    }


def collect_assets(blocks: list[dict]) -> None:
    root = CFG["root"]
    missing = []
    for b in blocks:
        if b["type"] == "fig":
            got = []
            for rel in b["paths"]:
                p = resolve_asset(root, rel, "FIG")
                if p is None:
                    missing.append(f"FIG {rel}（出现在 {Path(b['src']).name}）")
                got.append(p)
            if all(got):
                b["resolved"] = got
        elif b["type"] == "tab":
            p = resolve_asset(root, b["path"], "TAB")
            if p is None:
                missing.append(f"TAB {b['path']}（出现在 {Path(b['src']).name}）")
            else:
                b["resolved"] = p
    if missing:
        die("以下图片/表格文件不存在：\n  - " + "\n  - ".join(missing))


def table_geometry(doc) -> tuple[int, list[int]]:
    t = doc.tables[0]
    tblPr = t._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    grid = t._tbl.find(qn("w:tblGrid"))
    cols = [int(gc.get(qn("w:w"))) for gc in grid.findall(qn("w:gridCol"))] if grid is not None else list(DEFAULT_TABLE_COLS_TW)
    total = int(tblW.get(qn("w:w"))) if tblW is not None and tblW.get(qn("w:w")) else sum(cols)
    return total, cols


def _set_cell_text(cell, text, size, center=False):
    for p in cell.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p = cell.paragraphs[0]
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    set_indent(p, first_line_chars=0)
    set_run(p.add_run(text), size=size)


def clear_cell(cell):
    tc = cell._tc
    for child in list(tc):
        if child.tag in (qn("w:p"), qn("w:tbl")):
            tc.remove(child)


def shrink_paragraph(par):
    disable_grid_snap(par)
    pf = par.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(2)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.first_line_indent = Pt(0)
    pf.left_indent = Pt(0)
    pPr = par._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), "2")
        rPr.append(el)
    return par


def _spacer(cell):
    p = cell.add_paragraph()
    shrink_paragraph(p)
    set_run(p.add_run(""), size=Pt(1))


def compact_empty_paragraphs(cell):
    from docx.text.paragraph import Paragraph
    n = 0
    for ch in list(cell._tc):
        if ch.tag != qn("w:p"):
            continue
        if ch.find(qn("w:r")) is not None:
            continue
        shrink_paragraph(Paragraph(ch, cell))
        n += 1
    return n


def _ensure_trailing_paragraph(cell):
    tc = cell._tc
    last = [c for c in list(tc) if c.tag in (qn("w:p"), qn("w:tbl"))]
    if not last or last[-1].tag != qn("w:p"):
        _spacer(cell)


def _make_row_splittable(row):
    trPr = row._tr.get_or_add_trPr()
    el = trPr.find(qn("w:cantSplit"))
    if el is not None:
        trPr.remove(el)


def _fix_line_spacing(cell):
    """把表内 12pt / 1.5 倍行距统一写死为 24pt（480 twips），跳过含图段落。"""
    for par in cell.paragraphs:
        pel = par._p
        if pel.findall(".//" + qn("w:drawing")) or pel.findall(".//" + qn("w:pict")):
            continue
        pPr = pel.pPr
        if pPr is None:
            continue
        sp = pPr.find(qn("w:spacing"))
        if sp is None:
            continue
        if sp.get(qn("w:line")) == "360" and sp.get(qn("w:lineRule")) == "auto":
            sp.set(qn("w:line"), "480")
            sp.set(qn("w:lineRule"), "exact")
    for t in cell.tables:
        for row in t.rows:
            for c in row.cells:
                _fix_line_spacing(c)


def normalize_for_word(doc):
    """确保 Word 与 LibreOffice 渲染一致：固定正文行距 + 清理封面尾部空段后显式分页。"""
    tbl = doc.tables[0]
    for row in list(tbl.rows)[2:]:
        for cell in row.cells:
            _fix_line_spacing(cell)

    tbl_el = tbl._tbl
    # 先移除主表正前方连续的空段落（避免母版尾部多余空行把分页符顶到第 2 页造成整页空白）
    removed = 0
    prev = tbl_el.getprevious()
    while prev is not None and prev.tag == qn("w:p"):
        has_drawing = bool(prev.findall(".//" + qn("w:drawing")) or prev.findall(".//" + qn("w:pict")))
        txt = "".join(prev.itertext()).strip()
        if not txt and not has_drawing:
            earlier = prev.getprevious()
            prev.getparent().remove(prev)
            removed += 1
            prev = earlier
        else:
            break
    if removed:
        log(f"  · 清理封面尾部空段落 {removed} 个（防止出现空白第 2 页）")

    # 紧贴主表前插入显式分页符
    prev = tbl_el.getprevious()
    if prev is not None and prev.tag == qn("w:p") and prev.findall(".//" + qn("w:br")):
        return
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    tbl_el.addprevious(p)


def report_anomalies(blocks: list[dict]):
    """内容侧体检：格式异常、编号重复、正文引用与图表清单一致性。"""
    report_anomalies_warnings.clear()
    fig_caps, tab_caps = [], []
    cited: set[tuple[str, int]] = set()
    defined: set[tuple[str, int]] = set()
    any_cap_re = CFG["any_cap_re"]
    inline_ref_re = CFG["inline_ref_re"]

    for b in blocks:
        if b["type"] == "fig":
            fig_caps.append(b["cap"])
        elif b["type"] == "tab":
            tab_caps.append(b["cap"])
        elif b["type"] == "para":
            txt = b["text"]
            for m in inline_ref_re.finditer(txt):
                cited.add((m.group(1), int(m.group(2))))
            if txt.startswith("|"):
                report_anomalies_warnings.append(f"疑似 Markdown 表格行（请改用 CSV + {{{{TAB:...}}}}）：{txt[:40]}…")
            if (
                len(txt) > 2
                and txt[-1] not in SENT_END
                and not txt.startswith(("1", "2", "3", "4", "5", "6", "7", "8", "9"))
                and "：" not in txt[-3:]
            ):
                report_anomalies_warnings.append(f"段落未以标点结尾（可能是被硬换行的半句）：{txt[-24:]}")
            if len(txt) < 6 and not txt.startswith(("1", "2", "3", "4", "5")):
                report_anomalies_warnings.append(f"过短段落（可能是残留碎片）：{txt!r}")
        elif b["type"] == "heading":
            if len(b["text"]) > 42:
                report_anomalies_warnings.append(f"小标题过长：{b['text'][:30]}…")

    if fig_caps:
        log("  · 图：" + " / ".join(" ".join(c.split()[:2]) if " " in c else c for c in fig_caps))
    if tab_caps:
        log("  · 表：" + " / ".join(" ".join(c.split()[:2]) if " " in c else c for c in tab_caps))

    seen = []
    for c in fig_caps + tab_caps:
        num = any_cap_re.match(c)
        if num:
            key = (num.group(1), int(num.group(2)))
            defined.add(key)
            if key in seen:
                report_anomalies_warnings.append(f"图表编号重复：{c[:24]}")
            seen.append(key)

    # 交叉引用检查
    prefix = f"{CFG['exp_num']}-" if CFG["exp_num"] is not None else ""
    uncited = sorted(defined - cited)
    dangling = sorted(cited - defined)
    for kind, n in uncited:
        report_anomalies_warnings.append(f"正文中未引用已插入的「{kind} {prefix}{n}」")
    for kind, n in dangling:
        report_anomalies_warnings.append(f"正文引用了不存在的「{kind} {prefix}{n}」")


def build(demo: bool = False, forced_breaks=None, rerun: int = 1, shrink=None,
          keep_tables=None) -> Path:
    log("=" * 70)
    log(
        f"{CFG['exp_label']}《{CFG['exp_title']}》报告装配"
        + ("（DEMO 自测模式）" if demo else "")
        + (f"（第 {rerun} 遍排版修正）" if rerun > 1 else "")
    )
    log("=" * 70)

    doc = load_base_document()
    font_cleanup(doc)

    if not doc.tables:
        die("母版中未找到正文表格，请改用 --standalone 模式或检查母版文件。")
    tbl = doc.tables[0]
    if len(tbl.rows) < 5:
        die(f"母版主表格只有 {len(tbl.rows)} 行（期望 ≥5 行），请检查母版或改用 --standalone。")

    total_tw, cols = table_geometry(doc)
    cell_mar = 108 + 108
    avail_tw = total_tw - cell_mar - 60
    log(f"  · 正文表格宽 {total_tw} twips（列 {cols}），单元格可用宽 ≈ {avail_tw/1440*2.54:.2f}cm")

    rebuild_cover(doc)
    fill_header_rows(tbl)

    raw = load_content(demo)
    intro_txt, intro_src = raw["intro"]
    purpose_txt, env_txt = split_section0(intro_txt)
    summary_txt, summary_src = raw["summary"]
    section3_txt = strip_leading_title(summary_txt)

    r3_blocks = parse_md(env_txt, intro_src)
    for b_txt, b_src in raw["bodies"]:
        r3_blocks.extend(parse_md(b_txt, b_src))

    blocks_by_row = {
        2: parse_md(purpose_txt, intro_src),
        3: r3_blocks,
        4: parse_md(section3_txt, summary_src),
    }
    all_blocks = [b for bs in blocks_by_row.values() for b in bs]
    collect_assets(all_blocks)
    report_anomalies(all_blocks)

    TABLE_INFO.clear()
    builder = ContentBuilder(doc, avail_tw, forced_breaks, rerun, shrink, keep_tables)
    for ri in (2, 3, 4):
        cell = tbl.rows[ri].cells[0]
        _make_row_splittable(tbl.rows[ri])
        clear_cell(cell)
        builder.big_title(cell, BIG_TITLES[ri], keep_next=False)
        log(f"  · 填 r{ri}「{BIG_TITLES[ri]}」：{len(blocks_by_row[ri])} 个内容块")
        prev_was_table = False
        for b in blocks_by_row[ri]:
            t = b["type"]
            if t == "para":
                builder.body_para(cell, b["text"])
                prev_was_table = False
            elif t == "heading":
                builder.heading(cell, b["text"], b["level"])
                prev_was_table = False
            elif t == "fig":
                builder.figure(cell, b["resolved"], b["cap"])
                prev_was_table = False
            elif t == "tab":
                if prev_was_table:
                    _spacer(cell)
                builder.csv_table(cell, b["resolved"], b["cap"])
                prev_was_table = True
            elif t == "code":
                builder.code_block(cell, b["lines"])
                prev_was_table = False
            elif t == "out_legacy":
                builder.warnings.append(f"{{{{OUT:...}}}} 是旧写法，请改用 ```text 围栏块：{b['text'][:40]}…")
                builder.body_para(cell, b["text"])
                prev_was_table = False
            elif t == "bad_mark":
                die(f"未知标记 {{{{{b['name']}:...}}}}，出现在 {Path(b['src']).name}：\n  {b['text']}")
        _ensure_trailing_paragraph(cell)
        n_tiny = compact_empty_paragraphs(cell)
        if n_tiny:
            log(f"    · 压扁空段落 {n_tiny} 个")

    log("")
    log(
        f"  统计：正文段 {builder.n_para} ／ 图 {builder.n_fig} ／ "
        f"表 {builder.n_tab} ／ 围栏块 {builder.n_code}"
    )
    for w in report_anomalies_warnings + builder.warnings:
        log("  ! " + w)

    BUILD_INFO.clear()
    BUILD_INFO.update(
        fig_nat=builder.fig_nat,
        n_fig=builder.n_fig,
        n_tab=builder.n_tab,
        total_tw=total_tw,
        cover_count=len(CFG["cover_fields"]),
    )

    normalize_for_word(doc)
    out = (CFG["build"] / "demo_report.docx") if demo else CFG["final_docx"]
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    log("")
    log(f"已生成：{out}")
    log(f"  大小：{out.stat().st_size / 1024:.0f} KB")
    return out


# ---------------------------------------------------------------- DEMO 内容

def make_demo_sections() -> dict:
    demo_dir = CFG["build"] / "demo"
    demo_dir.mkdir(parents=True, exist_ok=True)
    try:
        from PIL import Image, ImageDraw
    except Exception:
        die("--demo 需要 Pillow 生成演示图片，请先安装 pillow")

    specs = [
        ("fig_a.png", (1440, 840), (33, 102, 172)),
        ("fig_b.png", (900, 1600), (178, 66, 22)),
        ("fig_c.png", (1440, 560), (46, 139, 87)),
    ]
    for name, (w, h), col in specs:
        p = demo_dir / name
        if p.exists():
            continue
        img = Image.new("RGB", (w, h), (255, 255, 255))
        dr = ImageDraw.Draw(img)
        dr.rectangle([2, 2, w - 3, h - 3], outline=col, width=3)
        for i in range(8):
            dr.line([(40 + i * (w - 80) // 8, h - 60), (40 + i * (w - 80) // 8, 60 + i * 20)],
                    fill=col, width=2)
        dr.text((30, 20), f"DEMO {name} {w}x{h}", fill=col)
        img.save(p)

    csv_path = demo_dir / "tab_demo.csv"
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        wr = csv.writer(f)
        wr.writerow(["序号", "区间", "样本数", "平均偏差", "R²"])
        wr.writerow(["1", "<15", "18", "-0.5134", "0.8120"])
        wr.writerow(["2", "15~20", "41", "0.2210", "0.8745"])

    n_pref = f"{CFG['exp_num']}-" if CFG["exp_num"] is not None else ""
    s0 = f"""## 实验目的

1、验证通用实验报告装配流水线：覆盖封面重建、三线表自适应列宽、图片等比缩放与多遍排版自检。

2、确认正文 12pt 宋体、固定 24pt 行距、首行缩进两字符与上标 R² 渲染正常。

## 实验环境与数据集

实验环境为 Python 3 + numpy / pandas / matplotlib / scikit-learn，所有随机过程固定 `random_state=0`。表 {n_pref}1 与图 {n_pref}1 用于演示三线表和宽幅插图效果。

{{{{TAB:{csv_path}|表 {n_pref}1 演示用三线表（含上标 R² 与负号）}}}}

{{{{FIG:{demo_dir / 'fig_a.png'}|图 {n_pref}1 演示用宽幅图片（14.2cm 宽）}}}}
"""
    s1 = f"""## 题目一 演示模块

### 1. 围栏块与竖长图测试

下列等宽代码输出块测试浅灰底纹与单倍行距，图 {n_pref}2 测试竖长图按 19cm 高度上限自动等比缩放，图 {n_pref}3 测试双图叠放共用一个图题。

```text
测试集指标：MAE = 0.2140, RMSE = 0.3012, R2 = 0.8745
```

{{{{FIG:{demo_dir / 'fig_b.png'}|图 {n_pref}2 演示用竖长图片（按 19cm 上限缩放）}}}}

{{{{FIG:{demo_dir / 'fig_a.png'}|{demo_dir / 'fig_c.png'}|图 {n_pref}3 演示用双图叠放共用图题}}}}
"""
    s3 = """## 实验小结

1、整条流水线（母版/独立版式、封面下划线齐平、去空白页、正文与图表装配、PDF 多遍自检）运行正常。

2、无硬编码个人信息与绝对路径，可直接分享给其他同学使用。
"""
    return {
        "intro": (s0, "demo_s0.md"),
        "bodies": [(s1, "demo_s1.md")],
        "summary": (s3, "demo_s3.md"),
    }


# ---------------------------------------------------------------- 多遍 PDF 自检

def run_selfcheck(docx_path: Path, expect_fig: int | None = None, expect_tab: int | None = None):
    soffice = CFG["soffice"]
    if not soffice:
        log("")
        log("! 未检测到 LibreOffice (soffice)，跳过 PDF 转换与几何自检（.docx 已生成完毕）。")
        log("  安装 LibreOffice 后重跑 `--check` 即可自动导出 PDF 并逐页自检。")
        return {}, [], set()

    pdf_dir = CFG["build"] / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    for old in pdf_dir.glob("*.png"):
        old.unlink()
    log("")
    log("── 自检：转 PDF ──")
    res = subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(docx_path)],
        capture_output=True, text=True,
    )
    pdf = pdf_dir / (docx_path.stem + ".pdf")
    if not pdf.exists():
        die("转 PDF 失败：\n" + res.stdout + res.stderr)
    log(f"  · PDF：{pdf}")

    script = r'''
import json, sys, fitz
pdf = sys.argv[1]; outdir = sys.argv[2]
d = fitz.open(pdf)
res = {"pages": d.page_count, "page": [d[0].rect.width, d[0].rect.height], "sheets": []}
for i, page in enumerate(d):
    pix = page.get_pixmap(dpi=110)
    fn = f"{outdir}/page-{i+1:02d}.png"
    pix.save(fn)
    info = {"n": i + 1, "images": [], "texts": [], "maxx": 0.0,
            "fonts": [], "hlines": [], "vlines": []}
    for b in page.get_text("dict")["blocks"]:
        if b["type"] == 1:
            info["images"].append([round(v, 1) for v in b["bbox"]])
        else:
            for l in b["lines"]:
                for s in l["spans"]:
                    info["fonts"].append(s["font"])
                txt = "".join(s["text"] for s in l["spans"])
                info["texts"].append({
                    "x0": round(l["bbox"][0], 1), "x1": round(l["bbox"][2], 1),
                    "y": round(l["bbox"][1], 1), "t": txt,
                    "sz": round(max(s["size"] for s in l["spans"]), 1),
                })
    for dr in page.get_drawings():
        r = dr["rect"]
        if r.height < 2.5 and r.width > 15:
            info["hlines"].append([round(r.x0, 1), round(r.x1, 1), round(r.y0, 1)])
        elif r.width < 2.5:
            info["vlines"].append([round(r.x0, 1), round(r.y0, 1), round(r.y1, 1)])
        info["maxx"] = max(info["maxx"], round(r.x1, 1))
    res["sheets"].append(info)
print(json.dumps(res))
'''
    r = subprocess.run([sys.executable, "-c", script, str(pdf), str(pdf_dir)], capture_output=True, text=True)
    if r.returncode != 0:
        die("PyMuPDF 逐页转 PNG 失败（请确认已安装 PyMuPDF）：\n" + r.stdout + r.stderr)

    data = json.loads(r.stdout.strip().splitlines()[-1])
    sheets = data["sheets"]
    pw, ph = data["page"]
    margin_l = margin_r = 3.175 / 2.54 * 72
    margin_t = margin_b = 2.54 / 2.54 * 72
    problems: list[str] = []
    fig_num_re = CFG["fig_num_re"]
    tab_num_re = CFG["tab_num_re"]

    log(f"  · 页数：{data['pages']}，页面 {pw:.1f}×{ph:.1f}pt（A4=595×842），PNG → {pdf_dir}")

    # 0) 空白页检查
    for s in sheets:
        non_empty_text = "".join(t["t"].strip() for t in s["texts"])
        if not non_empty_text and not s["images"]:
            problems.append(f"p{s['n']} 为空白页（无文字也无图片）")

    # 1) 图片是否越界
    log("── 自检：图片位置 ──")
    total_tw = BUILD_INFO.get("total_tw", DEFAULT_TABLE_TOTAL_TW)
    cell_r = margin_l + total_tw * 0.05 - 5.4
    for s in sheets:
        for bb in s["images"]:
            x0, y0, x1, y1 = bb
            flag = ""
            if x1 > cell_r + 4:
                problems.append(f"p{s['n']} 图片右边越界：x1={x1} > {cell_r:.0f}")
                flag = " ← 越界"
            if y1 > ph - margin_b + 3 or y0 < margin_t - 3:
                problems.append(f"p{s['n']} 图片纵向越界：{y0}~{y1}")
                flag += " ← 纵向越界"
            log(f"  p{s['n']:>2}  图 {x1-x0:6.1f}×{y1-y0:6.1f}pt  x[{x0:.0f},{x1:.0f}] y[{y0:.0f},{y1:.0f}]{flag}")

    # 2) 封面下划线
    log("── 自检：封面下划线 ──")
    if sheets:
        rows = {}
        for x0, x1, y in sheets[0]["hlines"]:
            if x0 > margin_l + 10 and (x1 - x0) < pw * 0.75:
                k = round(y)
                lo, hi = rows.get(k, (x0, x1))
                rows[k] = (min(lo, x0), max(hi, x1))
        for y in sorted(rows):
            lo, hi = rows[y]
            log(f"  下划线 x[{lo:.1f} → {hi:.1f}] y={y}  长度 {hi-lo:.1f}pt")
        ends = [hi for lo, hi in rows.values() if hi - lo > 30]
        expected_ul = BUILD_INFO.get("cover_count", len(CFG["cover_fields"]))
        if len(ends) >= 2:
            spread = max(ends) - min(ends)
            log(f"  下划线 {len(ends)} 条，右端散布 {spread:.1f}pt")
            if spread > 2.5:
                problems.append(f"封面下划线右端不齐（差 {spread:.1f}pt）")
        elif expected_ul and len(ends) < expected_ul:
            problems.append(f"封面下划线只识别到 {len(ends)} 条（应为 {expected_ul} 条）")

    # 3) 字体
    log("── 自检：字体 ──")
    fonts = sorted({f for s in sheets for f in s["fonts"]})
    log("  " + "、".join(fonts))
    for f in fonts:
        if re.search(r"DejaVu|Noto Sans CJK.*not|missing", f, re.I):
            problems.append(f"疑似字体回退：{f}")

    nat = dict(BUILD_INFO.get("fig_nat") or {})

    # 4) 图表题与内容同页 / 文本越界
    log("── 自检：标题-内容是否同页、文本框是否越界 ──")
    fixes: set[str] = set()
    right_limit = pw - margin_r + 14
    for s in sheets:
        for t in s["texts"]:
            if t["x1"] > right_limit:
                problems.append(f"p{s['n']} 文本越右边界：x1={t['x1']:.0f} 「{t['t'][:18]}」")

    for i, s in enumerate(sheets):
        txts = s["texts"]
        for t in txts:
            if t["t"].strip() in BIG_TITLES.values():
                body_below = [x for x in txts if x["y"] > t["y"] + 6 and len(x["t"].strip()) >= 15]
                if not body_below:
                    log(f"  ! p{s['n']} 大标题「{t['t'].strip()}」停在页底（仅提示）")
        for t in txts:
            if fig_num_re.match(t["t"].strip()) and t.get("sz", 12) <= 11.2:
                above = [b for b in s["images"] if b[3] <= t["y"] + 4 and t["y"] - b[3] < 90]
                if not above:
                    where = f"p{sheets[i-1]['n']} 底部有图" if i and sheets[i - 1]["images"] else "上方没有图片"
                    problems.append(f"p{s['n']} 图题「{t['t'][:16]}」与图分离（{where}）")
                    key = norm_key(t["t"])
                    fixed = False
                    if i and key in nat:
                        prev_img = sheets[i - 1]["images"]
                        if prev_img:
                            b = prev_img[-1]
                            ratio = nat[key][1]
                            max_h_pt = (ph - margin_b) - b[1] - 40
                            w_just = round((max_h_pt / 72 * 2.54) / ratio, 2)
                            if 10.5 <= w_just < nat[key][0]:
                                data.setdefault("shrink", {})[key] = w_just
                                log(f"    ↑ 「{t['t'][:16]}」缩到 {w_just:.2f}cm 可与图题留在 p{sheets[i-1]['n']}")
                                fixed = True
                    if not fixed:
                        fixes.add(key)
            if tab_num_re.match(t["t"].strip()) and t.get("sz", 12) <= 11.2:
                below = [x for x in txts if x["y"] > t["y"] + 2 and x["y"] - t["y"] < 40]
                if not below:
                    problems.append(f"p{s['n']} 表题「{t['t'][:16]}」下方没有表格内容（可能跨页分离）")
                    fixes.add(norm_key(t["t"]))

    # 4b) 图表齐全性
    log("── 自检：图表齐全性 ──")
    seen_fig, seen_tab = set(), set()
    for s_ in sheets:
        for t in s_["texts"]:
            m = fig_num_re.match(t["t"].strip())
            if m:
                seen_fig.add(int(m.group(1)))
            m = tab_num_re.match(t["t"].strip())
            if m:
                seen_tab.add(int(m.group(1)))

    ef = expect_fig if expect_fig is not None else BUILD_INFO.get("n_fig", 0)
    et = expect_tab if expect_tab is not None else BUILD_INFO.get("n_tab", 0)
    miss_fig = [i for i in range(1, ef + 1) if i not in seen_fig] if ef else []
    miss_tab = [i for i in range(1, et + 1) if i not in seen_tab] if et else []
    log(f"  图题命中 {len(seen_fig)}/{ef} {sorted(seen_fig)}，表题命中 {len(seen_tab)}/{et} {sorted(seen_tab)}")
    if miss_fig:
        problems.append(f"缺图题：编号 {miss_fig}（PDF 中未找到）")
    if miss_tab:
        problems.append(f"缺表题：编号 {miss_tab}（PDF 中未找到）")
        for tb in TABLE_INFO:
            m = tab_num_re.match(tb["cap"].strip())
            if m and int(m.group(1)) in miss_tab:
                fixes.add(tb["key"])
    if (ef or et) and not miss_fig and not miss_tab:
        log(f"  ✔ {ef} 图 {et} 表全部命中")

    # 4c) 三线表专项：跨页、表头重复、数字折断
    log("── 自检：表格跨页、表头重复、数字折断 ──")
    data.setdefault("keep", set())
    notes: list[str] = []
    if not any(re.search(r"SimSun|SimSong|Songti", f, re.I) for f in fonts):
        notes.append("PDF 中未出现 SimSun/宋体，LibreOffice 使用了替代字体（页数与留白仅供参照）")

    def table_rules(sheet_):
        return sorted(
            (h[2], h[0], h[1]) for h in sheet_["hlines"]
            if 280.0 <= h[1] - h[0] <= 430.0 and h[0] > 82.0
        )

    def cap_lines(sheet_):
        out = []
        for t in sheet_["texts"]:
            m = tab_num_re.match(t["t"].strip())
            if m and t.get("sz", 12) <= 11.2:
                out.append((t["y"], int(m.group(1)), t["t"]))
        return sorted(out)

    def row_bases(sheet_, y0, y1):
        buckets: dict[int, set] = {}
        for t in sheet_["texts"]:
            if y0 <= t["y"] <= y1:
                buckets.setdefault(round(t["y"] / 3.0), set()).add(round(t["x0"] / 3.0))
        ks = sorted(k for k, cols in buckets.items() if len(cols) >= 3)
        merged: list[int] = []
        for k in ks:
            if merged and k - merged[-1] <= 2:
                continue
            merged.append(k)
        return [k * 3.0 for k in merged]

    def base_text(sheet_, ybase):
        cells = [t for t in sheet_["texts"] if abs(t["y"] - ybase) <= 3.5]
        cells.sort(key=lambda t: t["x0"])
        return norm_key("".join(t["t"] for t in cells))

    page_bottom = ph - margin_b
    for idx, tb in enumerate(TABLE_INFO):
        rows_ = tb["rows"]
        ncol_ = tb["ncol"]
        short = tb["cap"][:18]
        size_ = tb["size"]
        m = tab_num_re.match(tb["cap"].strip())
        tno = int(m.group(1)) if m else idx + 1
        pno, ycap = None, None
        for i, sheet_ in enumerate(sheets):
            for y, n_, _t in cap_lines(sheet_):
                if n_ == tno:
                    pno, ycap = i, y
                    break
            if pno is not None:
                break
        if pno is None:
            problems.append(f"表「{short}」的题注在 PDF 中找不到，无法判定版式")
            continue
        rules = [r for r in table_rules(sheets[pno]) if r[0] > ycap + 1]
        if not rules:
            problems.append(f"p{pno + 1} 表「{short}」下方没有表格内容（题注与表分离）")
            fixes.add(tb["key"])
            continue
        next_cap = [y for y, _n, _t in cap_lines(sheets[pno]) if y > ycap + 1]
        lim = next_cap[0] if next_cap else page_bottom
        own_rules = [r for r in rules if r[0] < lim]
        split = len(own_rules) < 3
        top_y = own_rules[0][0] if own_rules else rules[0][0]
        last_y = own_rules[-1][0] if own_rules else rules[-1][0]

        band0, band1 = top_y - 8, (page_bottom if split else last_y + 6)
        lines = [t for t in sheets[pno]["texts"] if band0 <= t["y"] <= band1]
        by_x: dict[int, list] = {}
        for t in lines:
            by_x.setdefault(round(t["x0"] / 3.0), []).append(t)
        known = {norm_key(v) for r in rows_ for v in r if v.strip()}
        breaks = []
        for group in by_x.values():
            group.sort(key=lambda t: t["y"])
            for a, b2 in zip(group, group[1:]):
                ga, gb = a["t"].rstrip(), b2["t"].lstrip()
                if not ga or not gb:
                    continue
                if ga[-1] not in "0123456789.,%~" or gb[0] not in "0123456789.,":
                    continue
                if not 0 < b2["y"] - a["y"] < 16.0:
                    continue
                ka, kab = norm_key(ga), norm_key(ga + gb)
                if kab in known and ka not in known:
                    breaks.append((ga + "⏎" + gb)[:26])

        status = []
        if split:
            n_left = max(0, len(row_bases(sheets[pno], top_y + 4, last_y)) - 1)
            status.append(f"跨页 p{pno + 1}/{pno + 2}（本页数据行 {n_left}）")
            nxt = sheets[pno + 1] if pno + 1 < len(sheets) else None
            if nxt is not None:
                tops = row_bases(nxt, margin_t, margin_t + 120)
                hdr = norm_key("".join(rows_[0]))
                status.append("表头已重复" if (tops and base_text(nxt, tops[0]) == hdr) else "表头未重复")
            if tb["keep"]:
                fixes.add(tb["key"])
                problems.append(f"表「{short}」跨页（p{pno + 1}→p{pno + 2}）")
        else:
            status.append("未跨页")
        if breaks:
            problems.append(f"表「{short}」有 {len(breaks)} 处数字折断：" + "；".join(breaks[:4]))
            status.append(f"数字折断 {len(breaks)} 处")
        else:
            status.append("数字未折断")
        log(f"  表 {tb['cap'][:8]:<9} {tb['nrows']}×{ncol_:<2} {size_:.1f}pt  " + " / ".join(status))

    # 5) 大块空白
    log("── 自检：页内大块空白 ──")
    data.setdefault("shrink", {})
    data["gaps"] = []
    CAP_ALLOW, SLACK = 24.0, 2.0
    SOFT_CM = 12.0
    SOFT = SOFT_CM / 2.54 * 72
    WIDTH_LADDER = [14.2, 13.0, 12.0]
    W_MIN_CM = 12.0
    for idx, s_ in enumerate(sheets):
        ys = [t["y"] for t in s_["texts"]] + [b[3] for b in s_["images"]]
        if not ys or idx == len(sheets) - 1:
            continue
        gap = ph - margin_b - max(ys)
        if gap <= 120:
            continue
        if gap > SOFT:
            data["gaps"].append((s_["n"], gap))
            log(f"  p{s_['n']} 底部空白 {gap:.0f}pt = {gap/72*2.54:.1f}cm（> {SOFT_CM:g}cm）")
        else:
            log(f"  p{s_['n']} 底部空白 {gap:.0f}pt = {gap/72*2.54:.1f}cm（正常留白）")
        handled = False
        if gap > SOFT and idx + 1 < len(sheets):
            nxt = sheets[idx + 1]
            head = [b for b in nxt["images"] if b[1] <= margin_t + 8]
            if head:
                b = head[0]
                img_h = b[3] - b[1]
                cap = next(
                    (t["t"] for t in sorted(nxt["texts"], key=lambda v: v["y"])
                     if fig_num_re.match(t["t"].strip()) and t["y"] > b[3] - 6),
                    None,
                )
                if cap:
                    handled = True
                    key = norm_key(cap)
                    if img_h + CAP_ALLOW <= gap + 2:
                        log("    ↑ 图本身放得下，受 keep-with-next 影响，留待下遍观察")
                    elif key in nat:
                        _bw, ratio = nat[key]
                        fit_w = None
                        for w in WIDTH_LADDER[1:]:
                            h = min(ratio * w, FIG_MAX_H_CM)
                            if h / 2.54 * 72 + CAP_ALLOW + SLACK <= gap:
                                fit_w = w
                                break
                        if fit_w is not None:
                            data["shrink"][key] = fit_w
                            log(f"    ↑ 把「{cap[:18]}」缩到 {fit_w:.1f}cm 即可留在本页")
                        else:
                            notes.append(f"p{s_['n']} 底部空白 {gap/72*2.54:.1f}cm：「{cap[:18]}」缩到 {W_MIN_CM:g}cm 仍放不下")
        if not handled and gap > SOFT:
            notes.append(f"p{s_['n']} 底部空白 {gap/72*2.54:.1f}cm：表格行/大图整块下移到下一页")

    log("")
    log("=" * 70)
    if problems:
        log(f"自检发现 {len(problems)} 个问题：")
        for p in problems:
            log("  ! " + p)
    else:
        log("自检通过：无空白页、图片未越界、封面下划线齐平、图表题同页、无字体回退。")
    if notes:
        log(f"提示 {len(notes)} 处（通常不影响交付）：")
        for n in notes:
            log("  · " + n)
    log("=" * 70)
    return data, problems, fixes


# ---------------------------------------------------------------- CLI

def main(argv=None):
    ap = argparse.ArgumentParser(description="通用实验报告 Word / PDF 排版装配与自检脚本")
    ap.add_argument("--root", default=None, help="实验根目录（含 figs/、源码/、母版）")
    ap.add_argument("--config", default=None, help="实验报告配置文件 report_config.json")
    ap.add_argument("--profile", default=None, help="个人封面配置文件 profile.json")
    ap.add_argument("--master", default=None, help="手动指定母版 .doc/.docx 路径")
    ap.add_argument("--out", default=None, help="输出 docx 路径")
    ap.add_argument("--exp-num", type=int, default=None, help="实验编号（整数，用于图表号检查）")
    ap.add_argument("--exp-label", default=None, help="实验标签（如 实验十）")
    ap.add_argument("--exp-title", default=None, help="实验题目（如 决策树）")
    ap.add_argument("--date", default=None, help="实验日期（如 2026年11月30日）")
    ap.add_argument("--group", default=None, help="实验组别（默认 单人组）")
    ap.add_argument("--standalone", action="store_true", help="忽略母版，使用内置标准 A4 版式")
    ap.add_argument("--demo", action="store_true", help="用内置演示内容自测流水线")
    ap.add_argument("--check", action="store_true", help="构建后转 PDF/PNG 并多遍自检")
    ap.add_argument("--check-only", action="store_true", help="只对已有 docx 做自检")
    ap.add_argument("--docx", default=None, help="配合 --check-only 指定待检查 docx")
    ap.add_argument("--no-pdf", action="store_true", help="自检通过后不把 PDF 复制到输出目录")
    args = ap.parse_args(argv)

    init_runtime(args)

    if args.check_only:
        target = Path(args.docx).resolve() if args.docx else CFG["final_docx"]
        if not target.exists():
            die(f"找不到待检查的 docx：{target}")
        _data, problems, _fixes = run_selfcheck(target, expect_fig=CFG["expect_fig"], expect_tab=CFG["expect_tab"])
        return 0 if not problems else 1

    forced: set[str] = set()
    shrink: dict[str, float] = {}
    keep: set[str] = set()
    problems: list[str] = []
    out = None
    for it in range(1, 6):
        out = build(demo=args.demo, forced_breaks=forced, rerun=it, shrink=shrink, keep_tables=keep)
        if not args.check:
            return 0
        ef = 0 if args.demo else CFG["expect_fig"]
        et = 0 if args.demo else CFG["expect_tab"]
        data, problems, fixes = run_selfcheck(out, expect_fig=ef, expect_tab=et)
        new_forced = {f for f in fixes if f not in forced}
        new_shrink = {k: v for k, v in data.get("shrink", {}).items() if v < shrink.get(k, 1e9) - 0.05}
        new_keep = {k for k in data.get("keep", set()) if k not in keep}
        if not new_forced and not new_shrink and not new_keep:
            break
        log("")
        log(
            f"  ▶ 第 {it} 遍自检给出 {len(new_forced)} 处分页修正、"
            f"{len(new_shrink)} 处缩图修正、{len(new_keep)} 处整表不跨页修正，重排"
        )
        forced |= new_forced
        shrink.update(new_shrink)
        keep |= new_keep

    if args.check and not args.demo and not CFG["no_pdf"] and out is not None:
        src = CFG["build"] / "pdf" / (out.stem + ".pdf")
        dst = out.parent / (out.stem + ".pdf")
        if src.exists() and src.resolve() != dst.resolve():
            shutil.copyfile(src, dst)
            log("")
            log(f"已导出 PDF：{dst}")
            log(f"  大小：{dst.stat().st_size / 1024:.0f} KB")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(main())
