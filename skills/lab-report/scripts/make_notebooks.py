#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""将已跑通的 `# %%` 分节 Python 脚本转为带真实输出与内嵌图片的 Jupyter Notebook (.ipynb)。

不依赖 ipykernel / jupyter CLI：
  1. 按 `# %% [N.x] 标题` 切分 cell（首个 `# %%` 前的公共前导代码自动合并进第 1 个 cell）；
  2. 在共享 namespace 中顺序 exec 全部 cell，redirect_stdout 捕获真实标准输出；
  3. 临时挂钩 `matplotlib.figure.Figure.savefig`，每保存一张图即自动嵌入 PNG(base64)；
  4. 用 `nbformat.v4` 写出并校验。

用法::
    # 1) 单脚本模式
    python3 make_notebooks.py \
        --script 实验十/源码/exp10_dt.py \
        --out 实验十/exp10_dt.ipynb \
        --title "实验十 题目一：决策树分类" \
        --intro "本 Notebook 对应实验十题目一..." \
        --summary-file 实验十/源码/_build/summary1.md

    # 2) 批量配置模式（JSON 列表，每项含 script / out / title / intro / summary）
    python3 make_notebooks.py --config 实验十/源码/notebooks.json --root 实验十
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
import re
import sys
from contextlib import redirect_stdout
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import nbformat

CELL_RE = re.compile(r"^#\s*%%\s*(?:\[([^\]]*)\])?\s*(.*)$")


def split_cells(path: Path):
    lines = path.read_text(encoding="utf-8").splitlines()
    cells, title, buf, preamble = [], None, [], []
    for line in lines:
        m = CELL_RE.match(line.strip())
        if m:
            if title is None:
                preamble = buf[:]
            else:
                code = "\n".join(buf).strip("\n")
                if code.strip():
                    cells.append((len(cells) + 1, title or f"cell {len(cells) + 1}", code))
            title = (m.group(1) or "") + (" " + m.group(2).strip() if m.group(2).strip() else "")
            title = title.strip() or f"cell {len(cells) + 1}"
            buf = []
        else:
            buf.append(line)
    code = "\n".join(buf).strip("\n")
    if code.strip():
        cells.append((len(cells) + 1, title or f"cell {len(cells) + 1}", code))
    if preamble and cells:
        head = "\n".join(preamble).strip("\n")
        if head.strip():
            idx, first_title, first_code = cells[0]
            cells[0] = (idx, first_title, head + "\n\n" + first_code)
    return cells


class FigureCapture:
    """临时挂钩 Figure.savefig：脚本存图时同步捕获一份 PNG 供 Notebook 嵌入。"""

    def __init__(self):
        self._orig = Figure.savefig
        self.pending: list[bytes] = []

    def __enter__(self):
        orig = self._orig

        def hooked(self_fig, *args, **kwargs):
            result = orig(self_fig, *args, **kwargs)
            try:
                buf = io.BytesIO()
                orig(self_fig, buf, format="png", dpi=150, bbox_inches="tight")
                self.pending.append(buf.getvalue())
            except Exception as exc:
                print(f"[notebook 截图跳过] {type(exc).__name__}: {exc}")
            return result

        Figure.savefig = hooked
        return self

    def take(self) -> list[bytes]:
        out, self.pending = self.pending, []
        return out

    def __exit__(self, *exc):
        Figure.savefig = self._orig
        return False


def build_notebook(spec: dict, allow_no_images: bool = False) -> dict:
    path = Path(spec["script"]).resolve()
    out = Path(spec["out"]).resolve()
    if not path.exists():
        raise FileNotFoundError(f"找不到脚本：{path}")

    cells = split_cells(path)
    if not cells:
        raise RuntimeError(f"脚本中未找到 `# %%` 分节标记：{path}")

    ns = {"__name__": "__main__", "__file__": str(path), "__builtins__": __builtins__}
    old_argv, old_cwd = sys.argv, os.getcwd()
    sys.argv = [str(path)]
    os.chdir(path.parent)

    nb_cells = []
    title = spec.get("title") or path.stem
    intro = spec.get("intro") or ""
    header_md = f"# {title}\n\n{intro}".strip() if intro else f"# {title}"
    nb_cells.append(nbformat.v4.new_markdown_cell(source=header_md))

    img_count = 0
    try:
        with FigureCapture() as cap:
            for idx, cell_title, code in cells:
                cap.take()
                stdout = io.StringIO()
                with redirect_stdout(stdout):
                    try:
                        exec(compile(code, f"<cell {idx}>", "exec"), ns)
                    except Exception as exc:
                        raise RuntimeError(
                            f"{path.name} 第 {idx} 节「{cell_title}」执行失败：{type(exc).__name__}: {exc}"
                        ) from exc
                text = stdout.getvalue()

                nb_cells.append(nbformat.v4.new_markdown_cell(source=f"### {cell_title}"))
                outputs = []
                if text:
                    outputs.append(nbformat.v4.new_output("stream", name="stdout", text=text))
                for png in cap.take():
                    outputs.append(
                        nbformat.v4.new_output(
                            "display_data",
                            data={"image/png": base64.b64encode(png).decode("ascii")},
                            metadata={},
                        )
                    )
                    img_count += 1
                nb_cells.append(
                    nbformat.v4.new_code_cell(source=code, execution_count=idx, outputs=outputs)
                )
        plt.close("all")
    finally:
        sys.argv = old_argv
        os.chdir(old_cwd)

    summary = spec.get("summary")
    if summary:
        nb_cells.append(nbformat.v4.new_markdown_cell(source=summary))

    nb = nbformat.v4.new_notebook(cells=nb_cells)
    nb.metadata["kernelspec"] = {"display_name": "Python 3", "language": "python", "name": "python3"}
    nb.metadata["language_info"] = {
        "name": "python",
        "version": f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        "mimetype": "text/x-python",
        "file_extension": ".py",
        "pygments_lexer": "ipython3",
        "nbconvert_exporter": "python",
        "codemirror_mode": {"name": "ipython", "version": 3},
    }

    nbformat.validate(nb)
    if img_count == 0 and not allow_no_images:
        raise RuntimeError(f"{out.name} 未嵌入任何图片，请检查脚本是否调用 savefig（无图脚本可加 --allow-no-images）")

    out.parent.mkdir(parents=True, exist_ok=True)
    nbformat.write(nb, str(out))
    return {
        "out": str(out),
        "cells": len(cells),
        "images": img_count,
        "size_kb": round(out.stat().st_size / 1024, 1),
    }


def load_specs(args) -> list[dict]:
    if args.config:
        cfg_path = Path(args.config).resolve()
        base = Path(args.root).resolve() if args.root else cfg_path.parent
        raw = json.loads(cfg_path.read_text(encoding="utf-8"))
        items = raw.get("notebooks", raw) if isinstance(raw, dict) else raw
        specs = []
        for item in items:
            s_path = Path(item["script"])
            o_path = Path(item["out"])
            summary = item.get("summary", "")
            if not summary and item.get("summary_file"):
                sf = Path(item["summary_file"])
                sf = sf if sf.is_absolute() else (base / sf)
                summary = sf.read_text(encoding="utf-8")
            specs.append({
                "script": s_path if s_path.is_absolute() else (base / s_path),
                "out": o_path if o_path.is_absolute() else (base / o_path),
                "title": item.get("title"),
                "intro": item.get("intro", ""),
                "summary": summary,
            })
        return specs

    if not args.script:
        sys.exit("需要指定 --script 或 --config")
    s_path = Path(args.script).resolve()
    o_path = Path(args.out).resolve() if args.out else s_path.with_suffix(".ipynb")
    summary = args.summary or ""
    if not summary and args.summary_file:
        summary = Path(args.summary_file).read_text(encoding="utf-8")
    return [{
        "script": s_path,
        "out": o_path,
        "title": args.title or s_path.stem,
        "intro": args.intro or "",
        "summary": summary,
    }]


def main() -> int:
    ap = argparse.ArgumentParser(description="将 `# %%` 分节脚本转为带真实运行输出与图片的 .ipynb")
    ap.add_argument("--config", default=None, help="批量配置文件 JSON（含 notebooks 列表）")
    ap.add_argument("--root", default=None, help="配合 --config 解析相对路径的根目录")
    ap.add_argument("--script", default=None, help="单脚本模式：输入 .py 路径")
    ap.add_argument("--out", default=None, help="单脚本模式：输出 .ipynb 路径")
    ap.add_argument("--title", default=None, help="Notebook 顶部一级标题")
    ap.add_argument("--intro", default=None, help="顶部简介段落")
    ap.add_argument("--summary", default=None, help="末尾总结 Markdown 字符串")
    ap.add_argument("--summary-file", default=None, help="末尾总结 Markdown 文件路径")
    ap.add_argument("--allow-no-images", action="store_true", help="允许无图片的纯文本脚本")
    args = ap.parse_args()

    specs = load_specs(args)
    ok = True
    for spec in specs:
        try:
            info = build_notebook(spec, allow_no_images=args.allow_no_images)
            print(
                f"[OK] {info['out']}  code cells={info['cells']}  "
                f"images={info['images']}  size={info['size_kb']} KB"
            )
        except Exception as exc:
            ok = False
            print(f"[FAIL] {spec['out']}: {type(exc).__name__}: {exc}")
    print("=" * 66)
    print("全部 notebook 生成并校验通过" if ok else "存在失败项，请按上方信息修复")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
