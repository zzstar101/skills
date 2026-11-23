#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把带 `# %%` 分节标记的实验脚本渲染成 Jupyter Notebook 风格的「代码 + 运行输出」高清截图。

跨平台字体自动适配（macOS / Windows / Linux），支持中文与等宽 ASCII 2:1 对齐排版，
确保 `pandas` 表格输出在截图中依然整齐对齐。

用法::
    # 1) 列出脚本里的所有 cell
    python3 render_notebook.py path/to/exp.py --list

    # 2) 把第 1~3 个 cell 合成一张图
    python3 render_notebook.py path/to/exp.py --cells 1-3 \
        --out ../figs/fig_code.png --title "核心代码与运行输出"

    # 3) 每个 cell 单独出一张图
    python3 render_notebook.py path/to/exp.py --cells all --out-prefix ../figs/cell_
"""

from __future__ import annotations

import argparse
import io
import os
import re
import sys
import traceback
from contextlib import redirect_stdout
from pathlib import Path

os.environ.setdefault("MPLBACKEND", "Agg")

from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------- 跨平台字体发现

MONO_CANDIDATES = [
    # macOS
    "/System/Library/Fonts/Menlo.ttc",
    "/System/Library/Fonts/SFNSMono.ttf",
    "/Library/Fonts/Consolas.ttf",
    # Windows
    r"C:\Windows\Fonts\consola.ttf",
    r"C:\Windows\Fonts\CascadiaMono.ttf",
    r"C:\Windows\Fonts\cour.ttf",
    # Linux
    "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationMono-Regular.ttf",
    "/usr/share/fonts/truetype/ubuntu/UbuntuMono-R.ttf",
]

CJK_CANDIDATES = [
    # macOS
    "/System/Library/Fonts/Hiragino Sans GB.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/Supplemental/Songti.ttc",
    # Windows
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
    # Linux
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
]


def find_font(override: str | None, candidates: list[str], kind: str) -> str:
    if override:
        if Path(override).exists():
            return override
        sys.exit(f"指定的 {kind} 字体不存在：{override}")
    for c in candidates:
        if Path(c).exists():
            return c
    sys.exit(
        f"未能自动找到系统 {kind} 字体，请用命令行参数 "
        f"{'--mono-font' if kind == '等宽' else '--cjk-font'} 手动指定 .ttf/.ttc 字体路径。"
    )


# ---------------------------------------------------------------- 配色与排版参数

BG = (255, 255, 255)
CARD_BG = (247, 247, 250)
CARD_BORDER = (224, 224, 231)
TITLE_COLOR = (31, 35, 40)
BAR_COLOR = (48, 112, 179)
PROMPT_COLOR = (32, 119, 180)
OUT_COLOR = (33, 33, 33)
C_STR = (163, 21, 21)
C_NUM = (9, 134, 88)
C_KW = (0, 0, 255)
C_FUNC = (121, 94, 38)
C_COMMENT = (128, 128, 128)
C_DEFAULT = (31, 31, 31)

FONT_SIZE = 15
LINE_H = 26
PAD = 22
CARD_PAD_X = 18
CARD_PAD_Y = 14
MAX_W = 1660

KEYWORDS = {
    "def", "class", "for", "while", "if", "elif", "else", "return", "import",
    "from", "as", "in", "not", "and", "or", "is", "None", "True", "False",
    "try", "except", "finally", "with", "lambda", "print", "break", "continue",
    "pass", "global", "yield", "assert", "del", "raise", "range",
}

TOKEN_RE = re.compile(
    r"(?P<comment>\#.*$)"
    r'|(?P<string>[fr]?"(?:[^"\\]|\\.)*"|[fr]?\'(?:[^\'\\]|\\.)*\')'
    r"|(?P<number>\b\d+\.\d+\b|\b\d+\b)"
    r"|(?P<word>[A-Za-z_][A-Za-z0-9_]*)"
)

CELL_RE = re.compile(r"^#\s*%%\s*(?:\[([^\]]*)\])?\s*(.*)$")


def is_cjk(ch: str) -> bool:
    o = ord(ch)
    return (
        0x2E80 <= o <= 0x9FFF
        or 0xF900 <= o <= 0xFAFF
        or 0xFF00 <= o <= 0xFF60
        or 0xFFE0 <= o <= 0xFFE6
        or 0x3000 <= o <= 0x303F
        or o in (0x2018, 0x2019, 0x201C, 0x201D, 0x2026, 0x2014, 0x2013)
    )


class Painter:
    """汉字占 2 个等宽单元、ASCII 占 1 个单元的混排绘制器。"""

    def __init__(self, mono_path: str, cjk_path: str, size: int = FONT_SIZE):
        self.mono = ImageFont.truetype(mono_path, size, index=0)
        self.cjk = ImageFont.truetype(cjk_path, size, index=0)
        self.cell = self.mono.getlength("M")

    def cells(self, text: str) -> float:
        return sum(2 if is_cjk(c) else 1 for c in text) * self.cell

    def draw(self, img, x, y, text, color=C_DEFAULT, font=None):
        d = ImageDraw.Draw(img)
        font = font or self.mono
        for ch in text:
            if is_cjk(ch):
                w = 2 * self.cell
                try:
                    gw = d.textlength(ch, font=self.cjk)
                except Exception:
                    gw = w
                d.text((x + (w - gw) / 2, y), ch, font=self.cjk, fill=color)
                x += w
            else:
                d.text((x, y), ch, font=font, fill=color)
                x += self.cell
        return x


def colorize_code(line: str):
    out = []
    pos = 0
    for m in TOKEN_RE.finditer(line):
        if m.start() > pos:
            out.append((line[pos:m.start()], C_DEFAULT))
        kind = m.lastgroup
        text = m.group()
        if kind == "comment":
            color = C_COMMENT
        elif kind == "string":
            color = C_STR
        elif kind == "number":
            color = C_NUM
        else:
            if text in KEYWORDS:
                color = C_KW
            elif line[m.end():m.end() + 1] == "(":
                color = C_FUNC
            else:
                color = C_DEFAULT
        out.append((text, color))
        pos = m.end()
        if kind == "comment":
            break
    if pos < len(line):
        out.append((line[pos:], C_DEFAULT))
    return out


# ---------------------------------------------------------------- Cell 切分与执行

def split_cells(path: str):
    """按 `# %% [x.y] 标题` 切分，首个 `# %%` 前的前导代码并入第 1 个 cell。"""
    with open(path, encoding="utf-8") as f:
        lines = f.read().splitlines()
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


def run_cells(path: str, cells):
    """顺序执行全部 cell，返回 [(序号, 标题, 代码, 输出文本)]。"""
    script = os.path.abspath(path)
    script_dir = os.path.dirname(script)
    ns = {"__name__": "__main__", "__file__": script, "__builtins__": __builtins__}
    old_argv, old_cwd = sys.argv, os.getcwd()
    sys.argv = [script]
    os.chdir(script_dir)
    results = []
    try:
        for idx, title, code in cells:
            buf = io.StringIO()
            err = None
            with redirect_stdout(buf):
                try:
                    exec(compile(code, f"<cell {idx}>", "exec"), ns)
                except Exception:
                    err = traceback.format_exc()
            if err:
                buf.write(err)
            results.append((idx, title, code, buf.getvalue().rstrip("\n")))
        try:
            import matplotlib.pyplot as plt
            plt.close("all")
        except Exception:
            pass
    finally:
        sys.argv = old_argv
        os.chdir(old_cwd)
    return results


# ---------------------------------------------------------------- 渲染

def wrap(text: str, painter: Painter, max_cells: int):
    if painter.cells(text) / painter.cell <= max_cells:
        return [text]
    out, cur, cur_w = [], "", 0.0
    for ch in text:
        w = 2 if is_cjk(ch) else 1
        if cur_w + w > max_cells:
            out.append(cur)
            cur, cur_w = "  " + ch, 2 + w
        else:
            cur += ch
            cur_w += w
    out.append(cur)
    return out


def render(blocks, out_path: str, mono_path: str, cjk_path: str,
           title: str | None = None, show_prompt: bool = True):
    p = Painter(mono_path, cjk_path)
    prompt_col = max((p.cells(f"In [{n}]:") for n, _, _ in blocks), default=0) + 18 if show_prompt else 0
    max_cells = int((MAX_W - 2 * PAD - prompt_col - 2 * CARD_PAD_X) // p.cell)

    laid = []
    for n, code, output in blocks:
        code_lines = []
        for ln in code.splitlines() if code.strip() else [""]:
            code_lines.extend(wrap(ln.expandtabs(4), p, max_cells))
        out_lines = []
        for ln in output.splitlines() if output.strip() else []:
            out_lines.extend(wrap(ln.expandtabs(4), p, max_cells))
        laid.append((n, code_lines, out_lines))

    content_cells = max([p.cells(l) / p.cell for _, cl, ol in laid for l in cl + ol] + [10])
    card_w = content_cells * p.cell + 2 * CARD_PAD_X
    total_w = int(2 * PAD + prompt_col + card_w)

    h = 2 * PAD + (42 if title else 0)
    for _n, cl, ol in laid:
        h += 2 * CARD_PAD_Y + len(cl) * LINE_H
        if ol:
            h += 10 + len(ol) * LINE_H
        h += 18
    total_h = int(h)

    img = Image.new("RGB", (total_w, total_h), BG)
    d = ImageDraw.Draw(img)
    y = PAD
    if title:
        d.rectangle([PAD, y + 6, PAD + 5, y + 26], fill=BAR_COLOR)
        p_title = Painter(mono_path, cjk_path, 17)
        p_title.draw(img, PAD + 16, y + 4, title, color=TITLE_COLOR)
        y += 42

    x0 = PAD + prompt_col
    for n, cl, ol in laid:
        card_h = 2 * CARD_PAD_Y + len(cl) * LINE_H
        d.rounded_rectangle(
            [x0, y, x0 + card_w, y + card_h], radius=8,
            fill=CARD_BG, outline=CARD_BORDER, width=1,
        )
        if show_prompt:
            p.draw(img, PAD, y + CARD_PAD_Y + 2, f"In [{n}]:", color=PROMPT_COLOR)
        ty = y + CARD_PAD_Y
        for ln in cl:
            cx = x0 + CARD_PAD_X
            for text, color in colorize_code(ln):
                cx = p.draw(img, cx, ty, text, color=color)
            ty += LINE_H
        y += card_h
        if ol:
            y += 10
            for ln in ol:
                p.draw(img, x0 + 2, y, ln, color=OUT_COLOR)
                y += LINE_H
        y += 18

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    img.save(out_path)
    print(f"[render] {out_path}  {img.size[0]}x{img.size[1]}")
    return img.size


def parse_cell_spec(spec: str, total: int):
    if spec in ("all", "*"):
        return list(range(1, total + 1))
    picked = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            picked.extend(range(int(a), int(b) + 1))
        elif part:
            picked.append(int(part))
    return [i for i in picked if 1 <= i <= total]


def main():
    ap = argparse.ArgumentParser(description="将 `# %%` 分节脚本渲染为 Jupyter 风格代码+输出截图")
    ap.add_argument("script", help="Python 实验脚本路径")
    ap.add_argument("--list", action="store_true", help="仅列出所有 cell")
    ap.add_argument("--cells", default="all", help="要渲染的 cell 编号，如 1-3 或 2,4")
    ap.add_argument("--out", default=None, help="合并输出 PNG 路径")
    ap.add_argument("--out-prefix", default=None, help="逐 cell 输出前缀")
    ap.add_argument("--title", default=None, help="顶部标题")
    ap.add_argument("--mono-font", default=None, help="手动指定等宽字体路径")
    ap.add_argument("--cjk-font", default=None, help="手动指定中文字体路径")
    args = ap.parse_args()

    cells = split_cells(args.script)
    if args.list:
        for idx, title, code in cells:
            print(f"{idx:>3}  {title}  ({len(code.splitlines())} 行)")
        return
    if not cells:
        sys.exit(f"未在 {args.script} 中找到 `# %%` 分节标记")

    mono_path = find_font(args.mono_font, MONO_CANDIDATES, "等宽")
    cjk_path = find_font(args.cjk_font, CJK_CANDIDATES, "中文")

    if args.out_prefix:
        for idx, title, code, output in run_cells(args.script, cells):
            render([(idx, code, output)], f"{args.out_prefix}{idx}.png", mono_path, cjk_path, title=title)
        return
    if not args.out:
        sys.exit("需要指定 --out 或 --out-prefix")

    picked = set(parse_cell_spec(args.cells, len(cells)))
    blocks = [
        (idx, code, output)
        for idx, _title, code, output in run_cells(args.script, cells)
        if idx in picked
    ]
    if not blocks:
        sys.exit(f"未匹配到任何 cell（--cells {args.cells}，总计 {len(cells)} 个）")
    render(blocks, args.out, mono_path, cjk_path, title=args.title)


if __name__ == "__main__":
    main()
