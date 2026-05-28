from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from zipfile import ZipFile

import pypandoc


FRONT_MATTER_RE = re.compile(r"\A---\s*\n.*?\n---\s*\n", re.S)
IMAGE_MD_RE = re.compile(r"!\[[^\]]*\]\([^)]+\)")
HTML_IMAGE_RE = re.compile(r"<img\b[^>]*>", re.I)
DISPLAY_DOLLAR_RE = re.compile(r"(?<!\\)\$\$(.+?)(?<!\\)\$\$", re.S)
INLINE_DOLLAR_RE = re.compile(r"(?<!\\)\$(?!\$)(.+?)(?<!\\)\$", re.S)
INLINE_PAREN_RE = re.compile(r"\\\((.*?)\\\)", re.S)
DISPLAY_BRACKET_RE = re.compile(r"\\\[(.*?)\\\]", re.S)
LATEX_COMMAND_NAMES = (
    "scriptscriptstyle",
    "Leftrightarrow",
    "leftrightarrow",
    "operatorname",
    "displaystyle",
    "overrightarrow",
    "Rightarrow",
    "Leftarrow",
    "rightarrow",
    "leftarrow",
    "varnothing",
    "leqslant",
    "geqslant",
    "boldsymbol",
    "overline",
    "underline",
    "textstyle",
    "scriptstyle",
    "therefore",
    "parallel",
    "because",
    "partial",
    "setminus",
    "subseteq",
    "supseteq",
    "emptyset",
    "mathfrak",
    "mathcal",
    "mathbb",
    "mathrm",
    "mathbf",
    "mathsf",
    "mathtt",
    "mathit",
    "textrm",
    "textbf",
    "mapsto",
    "notin",
    "pmod",
    "cfrac",
    "dfrac",
    "tfrac",
    "begin",
    "right",
    "left",
    "sqrt",
    "frac",
    "iiint",
    "prod",
    "oint",
    "iint",
    "lim",
    "sum",
    "int",
    "forall",
    "exists",
    "nabla",
    "approx",
    "equiv",
    "subset",
    "supset",
    "propto",
    "triangle",
    "times",
    "cdot",
    "infty",
    "cdots",
    "ldots",
    "vdots",
    "ddots",
    "qquad",
    "alpha",
    "gamma",
    "delta",
    "epsilon",
    "varepsilon",
    "theta",
    "vartheta",
    "lambda",
    "Lambda",
    "sigma",
    "omega",
    "Gamma",
    "Omega",
    "Theta",
    "Delta",
    "Sigma",
    "Phi",
    "Psi",
    "simeq",
    "cong",
    "perp",
    "angle",
    "colon",
    "dots",
    "quad",
    "beta",
    "zeta",
    "eta",
    "rho",
    "tau",
    "varphi",
    "phi",
    "chi",
    "psi",
    "gets",
    "cup",
    "cap",
    "mid",
    "vee",
    "lor",
    "land",
    "wedge",
    "oplus",
    "otimes",
    "neq",
    "not",
    "leq",
    "geq",
    "sin",
    "cos",
    "tan",
    "log",
    "exp",
    "xi",
    "pi",
    "mu",
    "nu",
    "mp",
    "le",
    "ge",
    "to",
    "ni",
    "in",
    "ne",
    "pm",
    "Pr",
    "ln",
)
LATEX_NAMED_COMMAND_RE = re.compile(r"\\(?:" + "|".join(LATEX_COMMAND_NAMES) + r")(?![A-Za-z])")
LATEX_RESIDUAL_RE = re.compile(r"\\(?:[A-Za-z]+|[,;:! ]|[{}_^~'\"`|])|(?<!\\)\$")
INLINE_COMMAND_RE = LATEX_NAMED_COMMAND_RE
DISPLAY_COMMAND_RE = LATEX_NAMED_COMMAND_RE
MATH_RELATION_RE = re.compile(
    r"[=<>≤≥≠≈≡∈∉⊥∥→←⇒⇐↔⇔⊂⊆⊃⊇∝]|"
    r"\\(?:leqslant|geqslant|leq|geq|le|ge|neq|ne|approx|sim|cong|simeq|equiv|"
    r"pmod|perp|parallel|in|notin|ni|subset|subseteq|supset|supseteq|propto|"
    r"to|rightarrow|leftarrow|Rightarrow|Leftarrow|leftrightarrow|Leftrightarrow|mapsto)(?![A-Za-z])"
)
STYLE_COMMAND_RE = re.compile(r"\\(?:displaystyle|textstyle|scriptstyle|scriptscriptstyle)\b")


def _looks_like_display_math(text: str) -> bool:
    compact = text.strip()
    if not compact:
        return False
    if re.search(r"[\u4e00-\u9fff]", compact):
        return False
    if "\\begin{" in compact:
        return True
    command_count = len(DISPLAY_COMMAND_RE.findall(compact))
    relation_count = len(MATH_RELATION_RE.findall(compact))
    script_count = len(re.findall(r"[_^]\s*(?:\{|[A-Za-z0-9])", compact))
    return command_count >= 1 and (relation_count >= 1 or script_count >= 2)


def normalize_latex_math(expr: str) -> str:
    expr = expr.replace("\r\n", "\n").replace("\r", "\n")
    expr = STYLE_COMMAND_RE.sub("", expr)
    expr = expr.replace("\\dots", "\\ldots")
    expr = expr.replace("\\dotsc", "\\ldots").replace("\\dotsb", "\\ldots").replace("\\dotso", "\\ldots")
    expr = re.sub(r"\s*([_^])\s*", r"\1", expr)
    expr = re.sub(r"\{\s*([^{}\n]+?)\s*\}", r"{\1}", expr)
    expr = re.sub(r"\\(?:,|;|:|!)", " ", expr)
    expr = re.sub(r"[ \t]+", " ", expr)
    expr = re.sub(r"\n{3,}", "\n\n", expr)
    return expr.strip()


def _normalize_inline_math(text: str) -> str:
    text = INLINE_PAREN_RE.sub(lambda m: f"${normalize_latex_math(m.group(1))}$", text)
    return INLINE_DOLLAR_RE.sub(lambda m: f"${normalize_latex_math(m.group(1))}$", text)


def _flush_paragraph(result: list[str], paragraph: list[str]) -> None:
    if not paragraph:
        return
    text = "\n".join(paragraph).strip()
    if _looks_like_display_math(text):
        result.extend(["$$", normalize_latex_math(text), "$$", ""])
    else:
        normalized = _normalize_inline_math(text)
        result.extend(normalized.splitlines())
        result.append("")
    paragraph.clear()


def normalize_math_markdown(markdown: str) -> str:
    markdown = FRONT_MATTER_RE.sub("", markdown or "")
    markdown = markdown.replace("\r\n", "\n").replace("\r", "\n")
    markdown = DISPLAY_BRACKET_RE.sub(lambda m: f"\n\n$$\n{normalize_latex_math(m.group(1))}\n$$\n\n", markdown)
    markdown = DISPLAY_DOLLAR_RE.sub(lambda m: f"\n\n$$\n{normalize_latex_math(m.group(1))}\n$$\n\n", markdown)

    result: list[str] = []
    paragraph: list[str] = []
    in_math = False
    math_lines: list[str] = []

    for line in markdown.splitlines():
        stripped = line.strip()
        if stripped == "$$":
            if in_math:
                result.extend(["$$", normalize_latex_math("\n".join(math_lines)), "$$", ""])
                math_lines = []
                in_math = False
                continue
            if paragraph and _looks_like_display_math("\n".join(paragraph)):
                result.extend(["$$", normalize_latex_math("\n".join(paragraph)), "$$", ""])
                paragraph = []
                continue
            _flush_paragraph(result, paragraph)
            in_math = True
            continue

        if in_math:
            math_lines.append(line)
            continue

        if not stripped:
            _flush_paragraph(result, paragraph)
            continue
        paragraph.append(line)

    if in_math:
        content = "\n".join(math_lines).strip()
        if content:
            result.extend(["$$", normalize_latex_math(content), "$$", ""])
    _flush_paragraph(result, paragraph)

    normalized = "\n".join(result)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized).strip()
    return normalized + "\n" if normalized else ""


def editable_docx_bytes(markdown: str) -> bytes:
    normalized = normalize_math_markdown(markdown)
    with tempfile.TemporaryDirectory(prefix="handwriting_standard_docx_") as tmp:
        tmp_dir = Path(tmp)
        md_path = tmp_dir / "source.md"
        docx_path = tmp_dir / "standard.docx"
        md_path.write_text(normalized, encoding="utf-8")

        command = shutil.which("codex-md-to-docx")
        local_command = Path("/Users/hwaigc/.local/bin/codex-md-to-docx")
        if command or local_command.exists():
            subprocess.run([command or str(local_command), str(md_path), str(docx_path)], check=True)
        else:
            pypandoc.convert_file(
                str(md_path),
                to="docx",
                outputfile=str(docx_path),
                extra_args=["--from=markdown+tex_math_dollars", "--standalone"],
            )
        if not docx_path.exists():
            raise FileNotFoundError("标准 Word 校对稿生成失败")
        return docx_path.read_bytes()


def inspect_docx_math(docx_bytes: bytes) -> dict[str, int | bool]:
    with tempfile.TemporaryDirectory(prefix="handwriting_docx_inspect_") as tmp:
        path = Path(tmp) / "inspect.docx"
        path.write_bytes(docx_bytes)
        with ZipFile(path) as zf:
            document_xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
    visible_text = re.sub(r"<[^>]+>", "", document_xml)
    return {
        "office_math_objects": document_xml.count("<m:oMath") + document_xml.count("<m:oMathPara"),
        "has_latex_residuals": bool(LATEX_RESIDUAL_RE.search(visible_text)),
    }
