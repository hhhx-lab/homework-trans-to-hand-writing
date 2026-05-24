from __future__ import annotations

import math
import random
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable

from docx import Document
from docx.shared import Inches
from markdown_math import normalize_latex_math, normalize_math_markdown
from PIL import Image, ImageDraw, ImageFont


GREEK = {
    "alpha": "α",
    "beta": "β",
    "gamma": "γ",
    "delta": "δ",
    "epsilon": "ε",
    "varepsilon": "ε",
    "theta": "θ",
    "vartheta": "ϑ",
    "lambda": "λ",
    "mu": "μ",
    "pi": "π",
    "sigma": "σ",
    "varphi": "φ",
    "phi": "φ",
    "omega": "ω",
    "Omega": "Ω",
    "Lambda": "Λ",
    "Delta": "Δ",
    "Sigma": "Σ",
}

SYMBOLS = {
    "le": "≤",
    "leq": "≤",
    "leqslant": "≤",
    "ge": "≥",
    "geq": "≥",
    "geqslant": "≥",
    "neq": "≠",
    "approx": "≈",
    "sim": "∼",
    "infty": "∞",
    "times": "×",
    "cdot": "·",
    "ldots": "…",
    "cdots": "⋯",
    "vdots": "⋮",
    "ddots": "⋱",
    "dots": "…",
    "pm": "±",
    "mp": "∓",
    "to": "→",
    "rightarrow": "→",
    "leftarrow": "←",
    "Rightarrow": "⇒",
    "cap": "∩",
    "cup": "∪",
    "subset": "⊂",
    "subseteq": "⊆",
    "supset": "⊃",
    "supseteq": "⊇",
    "emptyset": "∅",
    "partial": "∂",
    "nabla": "∇",
    "forall": "∀",
    "exists": "∃",
    "in": "∈",
    "notin": "∉",
    "mid": "|",
    "vert": "|",
    "lvert": "|",
    "rvert": "|",
    "langle": "〈",
    "rangle": "〉",
    "wedge": "∧",
    "land": "∧",
    "vee": "∨",
    "lor": "∨",
    "Pr": "P",
    "ln": "ln",
    "exp": "exp",
    "sin": "sin",
    "cos": "cos",
    "tan": "tan",
    "log": "log",
}

BIG_OPERATORS = {"sum": "∑", "int": "∫", "prod": "∏", "lim": "lim"}
MATRIX_ENVS = {"matrix", "pmatrix", "bmatrix", "vmatrix", "cases", "aligned"}
STYLE_COMMANDS = {"displaystyle", "textstyle", "scriptstyle", "scriptscriptstyle", "limits", "nolimits"}
GROUP_WRAPPERS = {
    "mathrm",
    "mathbf",
    "mathbb",
    "mathcal",
    "operatorname",
    "text",
    "textrm",
    "boldsymbol",
    "overline",
    "underline",
    "hat",
    "bar",
    "tilde",
}
SIZE_DELIMITERS = {"big", "Big", "bigg", "Bigg", "bigl", "bigr", "Bigl", "Bigr", "biggl", "biggr", "Biggl", "Biggr"}
LATEX_RESIDUAL_RE = re.compile(
    r"\\(?:frac|sqrt|sum|int|begin|end|left|right|mathbf|mathrm|mathbb|operatorname|textstyle|displaystyle|ldots|cdots|dots)"
)


@dataclass
class HandwritingRenderConfig:
    line_spacing: int
    font_size: int
    left_margin: int
    top_margin: int
    right_margin: int
    bottom_margin: int
    word_spacing: int = 0
    line_spacing_sigma: float = 0
    font_size_sigma: float = 0
    word_spacing_sigma: float = 0
    perturb_x_sigma: float = 0
    perturb_y_sigma: float = 0
    perturb_theta_sigma: float = 0.04
    ink_depth_sigma: float = 0
    fill: tuple[int, int, int] = (0, 0, 0)
    seed: int = 14790897


class FontCache:
    def __init__(self, base_font):
        self.base_font = base_font
        self._cache: dict[int, ImageFont.FreeTypeFont] = {}

    def get(self, size: int):
        size = max(8, int(size))
        if size not in self._cache:
            self._cache[size] = self.base_font.font_variant(size=size)
        return self._cache[size]


class DrawContext:
    def __init__(self, draw: ImageDraw.ImageDraw, fonts: FontCache, config: HandwritingRenderConfig, rand: random.Random):
        self.draw = draw
        self.fonts = fonts
        self.config = config
        self.rand = rand

    def color(self) -> tuple[int, int, int]:
        delta = self.rand.gauss(0, self.config.ink_depth_sigma) if self.config.ink_depth_sigma else 0
        return tuple(max(0, min(255, int(channel + delta))) for channel in self.config.fill)

    def jitter(self) -> tuple[int, int]:
        return (
            round(self.rand.gauss(0, self.config.perturb_x_sigma)) if self.config.perturb_x_sigma else 0,
            round(self.rand.gauss(0, self.config.perturb_y_sigma)) if self.config.perturb_y_sigma else 0,
        )


class Box:
    width: int
    height: int
    baseline: int

    def draw(self, ctx: DrawContext, x: int, y: int) -> None:
        raise NotImplementedError

    def debug_text(self) -> str:
        return ""


class TextBox(Box):
    def __init__(self, text: str, fonts: FontCache, size: int):
        self.text = text
        self.size = size
        self.font = fonts.get(size)
        left, top, right, bottom = self.font.getbbox(text or " ")
        self.width = max(1, right - left)
        ascent, descent = self.font.getmetrics()
        self.baseline = ascent
        self.height = max(1, ascent + descent, bottom - top)

    def draw(self, ctx: DrawContext, x: int, y: int) -> None:
        dx, dy = ctx.jitter()
        actual_size = self.size
        if ctx.config.font_size_sigma:
            actual_size = max(8, round(ctx.rand.gauss(self.size, ctx.config.font_size_sigma)))
        font = ctx.fonts.get(actual_size)
        ctx.draw.text((x + dx, y + dy), self.text, fill=ctx.color(), font=font)

    def debug_text(self) -> str:
        return self.text


class HBox(Box):
    def __init__(self, children: list[Box], gap: int = 0):
        self.children = children
        self.gap = gap
        self.width = sum(child.width for child in children) + gap * max(0, len(children) - 1)
        self.baseline = max((child.baseline for child in children), default=0)
        below = max((child.height - child.baseline for child in children), default=0)
        self.height = self.baseline + below

    def draw(self, ctx: DrawContext, x: int, y: int) -> None:
        cursor = x
        for child in self.children:
            child.draw(ctx, cursor, y + self.baseline - child.baseline)
            cursor += child.width + self.gap

    def debug_text(self) -> str:
        return "".join(child.debug_text() for child in self.children)


class FractionBox(Box):
    def __init__(self, numerator: Box, denominator: Box, pad: int):
        self.numerator = numerator
        self.denominator = denominator
        self.pad = pad
        self.width = max(numerator.width, denominator.width) + pad * 2
        self.baseline = numerator.height + pad + 2
        self.height = numerator.height + denominator.height + pad * 3 + 2

    def draw(self, ctx: DrawContext, x: int, y: int) -> None:
        nx = x + (self.width - self.numerator.width) // 2
        dx = x + (self.width - self.denominator.width) // 2
        self.numerator.draw(ctx, nx, y)
        line_y = y + self.numerator.height + self.pad
        ctx.draw.line((x + self.pad // 2, line_y, x + self.width - self.pad // 2, line_y), fill=ctx.color(), width=max(2, self.pad // 3))
        self.denominator.draw(ctx, dx, line_y + self.pad)

    def debug_text(self) -> str:
        return f"({self.numerator.debug_text()})/({self.denominator.debug_text()})"


class SqrtBox(Box):
    def __init__(self, child: Box, size: int, fonts: FontCache):
        self.child = child
        self.root = TextBox("√", fonts, int(size * 1.15))
        self.pad = max(4, size // 16)
        self.width = self.root.width + child.width + self.pad * 2
        self.baseline = max(self.root.baseline, child.baseline + self.pad)
        self.height = max(self.root.height, child.height + self.pad * 2)

    def draw(self, ctx: DrawContext, x: int, y: int) -> None:
        self.root.draw(ctx, x, y + self.baseline - self.root.baseline)
        child_x = x + self.root.width + self.pad
        child_y = y + self.baseline - self.child.baseline + self.pad // 2
        self.child.draw(ctx, child_x, child_y)
        over_y = child_y + max(2, self.pad // 2)
        ctx.draw.line((child_x, over_y, child_x + self.child.width, over_y), fill=ctx.color(), width=max(1, self.pad // 3))

    def debug_text(self) -> str:
        return f"√({self.child.debug_text()})"


class ScriptBox(Box):
    def __init__(self, base: Box, sup: Box | None, sub: Box | None, limits: bool = False):
        self.base = base
        self.sup = sup
        self.sub = sub
        self.limits = limits
        if limits:
            side = max(sup.width if sup else 0, sub.width if sub else 0)
            self.width = max(base.width, side)
            top = (sup.height + 2) if sup else 0
            bottom = (sub.height + 2) if sub else 0
            self.height = top + base.height + bottom
            self.baseline = top + base.baseline
        else:
            side_width = max(sup.width if sup else 0, sub.width if sub else 0)
            self.width = base.width + side_width
            sup_top = sup.height if sup else 0
            sub_bottom = sub.height if sub else 0
            self.baseline = base.baseline + max(0, sup_top - base.baseline // 2)
            self.height = self.baseline + max(base.height - base.baseline, sub_bottom)

    def draw(self, ctx: DrawContext, x: int, y: int) -> None:
        if self.limits:
            if self.sup:
                self.sup.draw(ctx, x + (self.width - self.sup.width) // 2, y)
            base_y = y + (self.sup.height + 2 if self.sup else 0)
            self.base.draw(ctx, x + (self.width - self.base.width) // 2, base_y)
            if self.sub:
                self.sub.draw(ctx, x + (self.width - self.sub.width) // 2, base_y + self.base.height + 2)
            return
        self.base.draw(ctx, x, y + self.baseline - self.base.baseline)
        if self.sup:
            self.sup.draw(ctx, x + self.base.width, y)
        if self.sub:
            self.sub.draw(ctx, x + self.base.width, y + self.baseline + 2)

    def debug_text(self) -> str:
        text = self.base.debug_text()
        if self.sub:
            text += "_" + self.sub.debug_text()
        if self.sup:
            text += "^" + self.sup.debug_text()
        return text


class MatrixBox(Box):
    def __init__(self, rows: list[list[Box]], env: str, size: int, fonts: FontCache):
        self.rows = rows
        self.env = env
        self.pad = max(8, size // 8)
        columns = max((len(row) for row in rows), default=0)
        self.col_widths = [
            max((row[i].width for row in rows if i < len(row)), default=size // 2)
            for i in range(columns)
        ]
        self.row_heights = [max((cell.height for cell in row), default=size) for row in rows]
        body_width = sum(self.col_widths) + self.pad * max(0, columns - 1)
        body_height = sum(self.row_heights) + self.pad * max(0, len(rows) - 1)
        self.left_delim = TextBox(self._left_delim(), fonts, max(size, int(body_height * 0.75))) if self._left_delim() else None
        self.right_delim = TextBox(self._right_delim(), fonts, max(size, int(body_height * 0.75))) if self._right_delim() else None
        self.width = body_width + (self.left_delim.width if self.left_delim else 0) + (self.right_delim.width if self.right_delim else 0) + self.pad * 2
        self.height = max(body_height, self.left_delim.height if self.left_delim else 0, self.right_delim.height if self.right_delim else 0)
        self.baseline = self.height // 2

    def _left_delim(self) -> str:
        return {"pmatrix": "(", "bmatrix": "[", "vmatrix": "|", "cases": "{"}.get(self.env, "")

    def _right_delim(self) -> str:
        return {"pmatrix": ")", "bmatrix": "]", "vmatrix": "|"}.get(self.env, "")

    def draw(self, ctx: DrawContext, x: int, y: int) -> None:
        cursor_x = x
        if self.left_delim:
            self.left_delim.draw(ctx, cursor_x, y + (self.height - self.left_delim.height) // 2)
            cursor_x += self.left_delim.width + self.pad // 2
        body_y = y + (self.height - (sum(self.row_heights) + self.pad * max(0, len(self.rows) - 1))) // 2
        row_y = body_y
        for row_index, row in enumerate(self.rows):
            cell_x = cursor_x
            for col_index, cell in enumerate(row):
                cell.draw(ctx, cell_x + (self.col_widths[col_index] - cell.width) // 2, row_y + (self.row_heights[row_index] - cell.height) // 2)
                cell_x += self.col_widths[col_index] + self.pad
            row_y += self.row_heights[row_index] + self.pad
        if self.right_delim:
            self.right_delim.draw(ctx, x + self.width - self.right_delim.width, y + (self.height - self.right_delim.height) // 2)

    def debug_text(self) -> str:
        rows = ["[" + ",".join(cell.debug_text() for cell in row) + "]" for row in self.rows]
        return "[" + ";".join(rows) + "]"


class LatexParser:
    def __init__(self, text: str, fonts: FontCache, size: int):
        self.text = text.strip()
        self.fonts = fonts
        self.size = size
        self.pos = 0

    def parse(self) -> Box:
        box = self._parse_until("")
        return box if box.width else TextBox(" ", self.fonts, self.size)

    def _parse_until(self, terminator: str) -> Box:
        children: list[Box] = []
        while self.pos < len(self.text):
            if terminator and self.text.startswith(terminator, self.pos):
                self.pos += len(terminator)
                break
            if terminator == "}" and self.text[self.pos] == "}":
                self.pos += 1
                break
            atom = self._parse_atom()
            atom = self._parse_scripts(atom)
            if atom:
                children.append(atom)
        return HBox(children, gap=max(0, self.size // 18))

    def _parse_atom(self) -> Box:
        ch = self.text[self.pos]
        if ch.isspace():
            self.pos += 1
            return TextBox(" ", self.fonts, self.size // 2)
        if ch == "{":
            self.pos += 1
            return self._parse_until("}")
        if ch == "\\":
            return self._parse_command()
        self.pos += 1
        return TextBox(ch, self.fonts, self.size)

    def _read_command_name(self) -> str:
        self.pos += 1
        start = self.pos
        while self.pos < len(self.text) and self.text[self.pos].isalpha():
            self.pos += 1
        if self.pos == start and self.pos < len(self.text):
            self.pos += 1
        return self.text[start:self.pos]

    def _read_group_text(self) -> str:
        self._skip_space()
        if self.pos >= len(self.text) or self.text[self.pos] != "{":
            if self.pos < len(self.text):
                ch = self.text[self.pos]
                self.pos += 1
                return ch
            return ""
        depth = 0
        start = self.pos + 1
        self.pos += 1
        while self.pos < len(self.text):
            ch = self.text[self.pos]
            if ch == "{":
                depth += 1
            elif ch == "}":
                if depth == 0:
                    content = self.text[start:self.pos]
                    self.pos += 1
                    return content
                depth -= 1
            self.pos += 1
        return self.text[start:]

    def _parse_group(self, scale: float = 1.0) -> Box:
        content = self._read_group_text()
        return LatexParser(content, self.fonts, max(8, int(self.size * scale))).parse()

    def _parse_command(self) -> Box:
        name = self._read_command_name()
        if name in {"left", "right"}:
            self._skip_space()
            if self.pos < len(self.text):
                delimiter = self.text[self.pos]
                self.pos += 1
                if delimiter == ".":
                    return TextBox("", self.fonts, self.size)
                return TextBox(delimiter, self.fonts, int(self.size * 1.05))
            return TextBox("", self.fonts, self.size)
        if name in STYLE_COMMANDS:
            return TextBox("", self.fonts, self.size)
        if name in SIZE_DELIMITERS:
            return TextBox("", self.fonts, self.size)
        if name in GROUP_WRAPPERS:
            return self._parse_group()
        if name == "frac":
            numerator = self._parse_group(0.8)
            denominator = self._parse_group(0.8)
            return FractionBox(numerator, denominator, max(5, self.size // 10))
        if name == "sqrt":
            self._skip_optional()
            return SqrtBox(self._parse_group(0.9), self.size, self.fonts)
        if name == "begin":
            env = self._read_group_text().strip()
            return self._parse_environment(env)
        if name in BIG_OPERATORS:
            return TextBox(BIG_OPERATORS[name], self.fonts, int(self.size * (1.35 if name != "lim" else 1.0)))
        if name in GREEK:
            return TextBox(GREEK[name], self.fonts, self.size)
        if name in SYMBOLS:
            return TextBox(SYMBOLS[name], self.fonts, self.size)
        if name in {"\\", ",", ";", ":", "!", " "}:
            return TextBox(" ", self.fonts, self.size // 2)
        if name in {"{", "}", "_", "%", "#", "&"}:
            return TextBox(name, self.fonts, self.size)
        return TextBox(name, self.fonts, self.size)

    def _parse_environment(self, env: str) -> Box:
        end_marker = f"\\end{{{env}}}"
        end = self.text.find(end_marker, self.pos)
        if end < 0:
            return TextBox(env, self.fonts, self.size)
        content = self.text[self.pos:end]
        self.pos = end + len(end_marker)
        if env not in MATRIX_ENVS:
            return LatexParser(content, self.fonts, self.size).parse()
        rows = []
        for raw_row in re.split(r"\\\\", content):
            cells = [LatexParser(cell.strip(), self.fonts, max(8, int(self.size * 0.86))).parse() for cell in raw_row.split("&")]
            if cells:
                rows.append(cells)
        return MatrixBox(rows, env, self.size, self.fonts)

    def _parse_scripts(self, base: Box) -> Box:
        sup = None
        sub = None
        while True:
            self._skip_space()
            if self.pos >= len(self.text) or self.text[self.pos] not in {"^", "_"}:
                break
            marker = self.text[self.pos]
            self.pos += 1
            script = self._parse_group(0.62)
            if marker == "^":
                sup = script
            else:
                sub = script
        limits = base.debug_text() in {"∑", "∫", "∏", "lim"}
        return ScriptBox(base, sup, sub, limits=limits) if sup or sub else base

    def _skip_optional(self) -> None:
        self._skip_space()
        if self.pos >= len(self.text) or self.text[self.pos] != "[":
            return
        depth = 0
        self.pos += 1
        while self.pos < len(self.text):
            if self.text[self.pos] == "[":
                depth += 1
            elif self.text[self.pos] == "]":
                if depth == 0:
                    self.pos += 1
                    return
                depth -= 1
            self.pos += 1

    def _skip_space(self) -> None:
        while self.pos < len(self.text) and self.text[self.pos].isspace():
            self.pos += 1


def latex_to_box(expr: str, fonts: FontCache, size: int) -> Box:
    return LatexParser(normalize_latex_math(expr), fonts, size).parse()


def latex_to_debug_text(expr: str, font_path: Path | None = None, size: int = 48) -> str:
    font_path = font_path or Path(__file__).resolve().parent / "font_assets" / "神韵英子楷书.ttf"
    font = ImageFont.truetype(str(font_path), size=size)
    return latex_to_box(expr, FontCache(font), size).debug_text()


def _strip_markdown_markup(line: str) -> str:
    line = re.sub(r"^\s{0,3}#{1,6}\s*", "", line)
    line = re.sub(r"^\s*[-*+]\s+", "", line)
    line = re.sub(r"^\s*\d+[.)]\s+", lambda m: m.group(0).strip() + " ", line)
    line = line.replace("**", "").replace("__", "").replace("`", "")
    return line


def _blocks(markdown: str) -> list[tuple[str, str]]:
    lines = normalize_math_markdown(markdown).splitlines()
    blocks: list[tuple[str, str]] = []
    paragraph: list[str] = []
    in_math = False
    math_lines: list[str] = []
    for line in lines:
        stripped = line.strip()
        if stripped == "$$":
            if in_math:
                blocks.append(("math", "\n".join(math_lines)))
                math_lines = []
                in_math = False
            else:
                if paragraph:
                    blocks.append(("text", " ".join(paragraph)))
                    paragraph = []
                in_math = True
            continue
        if in_math:
            math_lines.append(line)
            continue
        if not stripped:
            if paragraph:
                blocks.append(("text", " ".join(paragraph)))
                paragraph = []
            continue
        paragraph.append(_strip_markdown_markup(line))
    if math_lines:
        blocks.append(("math", "\n".join(math_lines)))
    if paragraph:
        blocks.append(("text", " ".join(paragraph)))
    return blocks


INLINE_MATH_RE = re.compile(r"\$([^$\n]+)\$")


def _split_top_level_math(expr: str, separators: set[str]) -> list[str]:
    parts: list[str] = []
    start = 0
    brace_depth = 0
    paren_depth = 0
    bracket_depth = 0
    i = 0
    while i < len(expr):
        if expr.startswith("\\left", i) or expr.startswith("\\right", i):
            i += 5 if expr.startswith("\\left", i) else 6
            continue
        ch = expr[i]
        if ch == "\\":
            i += 1
            while i < len(expr) and expr[i].isalpha():
                i += 1
            continue
        if ch == "{":
            brace_depth += 1
        elif ch == "}":
            brace_depth = max(0, brace_depth - 1)
        elif ch == "(":
            paren_depth += 1
        elif ch == ")":
            paren_depth = max(0, paren_depth - 1)
        elif ch == "[":
            bracket_depth += 1
        elif ch == "]":
            bracket_depth = max(0, bracket_depth - 1)
        elif ch in separators and not brace_depth and not paren_depth and not bracket_depth:
            end = i + 1 if ch in {"=", "+", "-"} else i
            part = expr[start:end].strip()
            if part:
                parts.append(part)
            start = i + 1
        i += 1
    tail = expr[start:].strip()
    if tail:
        parts.append(tail)
    return parts if len(parts) > 1 else [expr.strip()]


def _display_math_lines(expr: str, available_width: int, fonts: FontCache, size: int) -> list[str]:
    expr = normalize_latex_math(expr)
    if not expr:
        return []
    if "\\begin{" in expr:
        return [expr]
    raw_lines = [line.strip() for line in re.split(r"\\\\", expr) if line.strip()]
    lines: list[str] = []
    for raw in raw_lines:
        if latex_to_box(raw, fonts, size).width <= available_width:
            lines.append(raw)
            continue
        split = _split_top_level_math(raw, {"="})
        if len(split) == 1:
            split = _split_top_level_math(raw, {"+", "-"})
        if len(split) == 1:
            split = _split_top_level_math(raw, {","})
        lines.extend(split)
    return lines


def _text_to_boxes(text: str, fonts: FontCache, size: int) -> list[Box]:
    boxes: list[Box] = []
    pos = 0
    for match in INLINE_MATH_RE.finditer(text):
        if match.start() > pos:
            boxes.extend(TextBox(ch, fonts, size) for ch in text[pos:match.start()])
        boxes.append(latex_to_box(match.group(1), fonts, size))
        pos = match.end()
    if pos < len(text):
        boxes.extend(TextBox(ch, fonts, size) for ch in text[pos:])
    return boxes


def _layout_inline(boxes: list[Box], available_width: int, word_spacing: int) -> list[HBox]:
    lines: list[list[Box]] = []
    current: list[Box] = []
    width = 0
    for box in boxes:
        gap = word_spacing if current else 0
        if current and width + gap + box.width > available_width:
            lines.append(current)
            current = [box]
            width = box.width
        else:
            current.append(box)
            width += gap + box.width
    if current:
        lines.append(current)
    return [HBox(line, gap=word_spacing) for line in lines]


def render_markdown_handwriting(
    markdown: str,
    background: Image.Image,
    font,
    config: HandwritingRenderConfig,
    progress_hook: Callable[[str, str, int], None] | None = None,
) -> list[Image.Image]:
    markdown = normalize_math_markdown(markdown)
    fonts = FontCache(font)
    rand = random.Random(config.seed)
    available_width = background.width - config.left_margin - config.right_margin
    max_y = background.height - config.bottom_margin
    pages: list[Image.Image] = []
    page = background.copy()
    draw = ImageDraw.Draw(page)
    ctx = DrawContext(draw, fonts, config, rand)
    y = config.top_margin

    def new_page() -> None:
        nonlocal page, draw, ctx, y
        pages.append(page)
        page = background.copy()
        draw = ImageDraw.Draw(page)
        ctx = DrawContext(draw, fonts, config, rand)
        y = config.top_margin

    def draw_line(line: Box, extra_gap: int = 0, center: bool = False) -> None:
        nonlocal y
        line_height = max(config.line_spacing, line.height + extra_gap)
        if y + line_height > max_y and y > config.top_margin:
            new_page()
        x = config.left_margin + ((available_width - line.width) // 2 if center and line.width < available_width else 0)
        line.draw(ctx, x, y)
        y += line_height

    for index, (kind, content) in enumerate(_blocks(markdown), start=1):
        if progress_hook:
            progress_hook("rendering", f"正在处理第 {index} 段", min(90, 45 + index))
        if kind == "math":
            for math_line in _display_math_lines(content, available_width, fonts, config.font_size):
                size = config.font_size
                box = latex_to_box(math_line, fonts, size)
                while box.width > available_width and size > 24:
                    size = max(24, int(size * 0.9))
                    box = latex_to_box(math_line, fonts, size)
                draw_line(box, extra_gap=config.line_spacing // 2, center=True)
            continue
        for line in _layout_inline(_text_to_boxes(content, fonts, config.font_size), available_width, config.word_spacing):
            draw_line(line)
        y += max(4, config.line_spacing // 5)

    pages.append(page)
    return pages


def images_to_docx(image_paths: Iterable[Path], output_path: Path, dpi: int = 300) -> Path:
    paths = list(image_paths)
    if not paths:
        raise ValueError("No images to export")
    first = Image.open(paths[0])
    width_in = first.width / dpi
    height_in = first.height / dpi
    first.close()
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(width_in)
    section.page_height = Inches(height_in)
    section.top_margin = Inches(0)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0)
    section.right_margin = Inches(0)
    for index, path in enumerate(paths):
        if index:
            doc.add_page_break()
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.space_after = 0
        paragraph.paragraph_format.space_before = 0
        run = paragraph.add_run()
        run.add_picture(str(path), width=Inches(width_in))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return output_path


def write_images_to_docx_bytes(image_paths: Iterable[Path]) -> bytes:
    with tempfile.TemporaryDirectory(prefix="handwriting_docx_") as tmp:
        output = Path(tmp) / "handwriting.docx"
        images_to_docx(image_paths, output)
        return output.read_bytes()
