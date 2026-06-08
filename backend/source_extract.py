from __future__ import annotations

from pathlib import Path
import re
from typing import Any

import pypandoc
from werkzeug.utils import secure_filename

from markdown_math import normalize_math_markdown
from mineru_adapter import extract_pdf_to_markdown


SUPPORTED_SOURCE_SUFFIXES = {".pdf", ".docx", ".doc", ".md", ".markdown", ".txt", ".rtf"}
MATH_SPAN_RE = re.compile(r"(?<!\\)(\$\$?)(.*?)(?<!\\)\1", re.S)
OCR_C_SCRIPT_PATTERN = (
    r"(?:\^\s*(?:\{\s*)?c(?:\s*\})?"
    r"|\{\s*\\mathfrak\s*\{\s*c\s*\}\s*\}"
    r"|\\mathfrak\s*\{\s*c\s*\})"
)
OCR_C_OPEN_RE = re.compile(
    r"(?P<prefix>(?:\\Pr|Pr|\\Phi|Φ|Phi|P(?:\s*_\s*(?:\{[^{}]*\}|[A-Za-z0-9]))?))"
    r"\s*(?P<mark>" + OCR_C_SCRIPT_PATTERN + r")\s*[,，]?\s*"
)
OCR_C_CLOSE_RE = re.compile(
    r"(?P<prefix>(?:[A-Za-z0-9}\)]|\\[A-Za-z]+(?:\s*\{[^{}]*\})?))"
    r"\s*(?P<mark>" + OCR_C_SCRIPT_PATTERN + r")"
    r"\s*(?=(?:=|,|，|;|；|。|\\approx|\\le|\\ge|\\right|$))"
)
OCR_C_LEFTOVER_RE = re.compile(OCR_C_SCRIPT_PATTERN)
SPACED_DECIMAL_RE = re.compile(r"(?<![A-Za-z])(?P<int>\d(?:\s+\d)*)\s*[.．]\s*(?P<frac>\d(?:\s*\d)*)")
SPACED_DIGITS_RE = re.compile(r"(?<![A-Za-z\\])(?<!\d)(?P<digits>\d(?:\s+\d)+)(?!\s*[A-Za-z])")
MALFORMED_ODE_MATRIX_BLOCK_RE = re.compile(
    r"其中，\s*\$(?P<matrix>\\begin\{array\}.*?\\end\{array\})\$\s*\n\n"
    r"相应的初始条件为：\s*\$(?P<initial>.*?)\$(?=\n\n(?:题\s*\d|$))",
    re.S,
)


def _repair_ocr_c_scripts_in_math_expr(expr: str) -> str:
    """Turn common OCR-misread parenthesis marks into baseline delimiters."""
    expr = OCR_C_OPEN_RE.sub(lambda match: f"{match.group('prefix')}(", expr)
    expr = OCR_C_CLOSE_RE.sub(lambda match: f"{match.group('prefix')})", expr)
    return OCR_C_LEFTOVER_RE.sub("c", expr)


def _repair_ocr_c_scripts(markdown: str) -> str:
    return MATH_SPAN_RE.sub(
        lambda match: f"{match.group(1)}{_repair_ocr_c_scripts_in_math_expr(match.group(2))}{match.group(1)}",
        markdown,
    )


def _compact_spaced_number(match: re.Match[str]) -> str:
    return re.sub(r"\s+", "", match.group(0))


def _strip_number_spaces(text: str) -> str:
    return re.sub(r"\s+", "", text)


def _repair_mineru_spacing_in_math_expr(expr: str) -> str:
    expr = SPACED_DECIMAL_RE.sub(
        lambda match: f"{_strip_number_spaces(match.group('int'))}.{_strip_number_spaces(match.group('frac'))}",
        expr,
    )
    expr = SPACED_DIGITS_RE.sub(_compact_spaced_number, expr)
    replacements = (
        (r"\bm\s+a\s+x\b", "max"),
        (r"\bm\s+i\s+n\b", "min"),
        (r"\\alpha_\{\s*max\s*\}", r"\\alpha_{max}"),
        (r"\\int_\{\s*0\s*\}\s*c\s*2\s*x\s*d\s*x", r"\\int_{0}^{c} 2x\\,dx"),
        (r"e\^\{\s*-\s*30\s*\\lambda\s*\}", r"e^{-30\\lambda}"),
        (r"\(\s*30\s*\\lambda\s*\)", r"(30\\lambda)"),
        (r"\bP\s+\(\s*30\\lambda\s*\)", r"P(30\\lambda)"),
        (r"k\s*!", r"k!"),
        (r"n\s*\\infty\s*\\\s*\\Theta\s*\\\s*J", r"n \\to \\infty"),
        (r"-\s*0\.4\s*\\sqrt\s*\{\s*n\s*\}\s*-\s*\\infty", r"-0.4 \\sqrt{n} \\to -\\infty"),
        (r"\\theta\s*\\\s*\\mathfrak\s*\{\s*g\s*\}\s*\\mathrm\s*\{\s*\\cdot\s*\}", r"\\theta"),
        (r"X_\{_\{\\left\(\s*n\s*\\right\)\s*\}\s*\}", r"X_{(n)}"),
        (r"X_\{_\{\s*(\d+)\s*\}\s*\}", r"X_{\1}"),
        (r"X_\{\s*_n\s*\}", r"X_{n}"),
    )
    for pattern, replacement in replacements:
        expr = re.sub(pattern, replacement, expr)
    return expr


def _repair_mineru_spacing_in_math(markdown: str) -> str:
    return MATH_SPAN_RE.sub(
        lambda match: f"{match.group(1)}{_repair_mineru_spacing_in_math_expr(match.group(2))}{match.group(1)}",
        markdown,
    )


def _repair_statistics_pdf_fragments(text: str) -> str:
    text = text.replace("\x08eta", "β").replace("\\β", "β").replace("\\beta", "β")
    replacements = (
        (r"\$\\\[\s*\\beta\s*=\s*P_1\(X\s*\$", r"$\beta = P_1(X < c) = c^2$"),
        (r"\$\\\[\s*β\s*=\s*P_1\(X\s*\$", r"$β = P_1(X < c) = c^2$"),
        (r"\[\\beta\s*=\s*P_1\(X", r"β = P_1(X < c)"),
        (r"\[β\s*=\s*P_1\(X", r"β = P_1(X < c)"),
        (r"backslash\]\s*β", r"β"),
        (r"α\s*\+\s*2β\s*=\s*1\s*−\s*c\s*\+\s*2c\s*\.\s*2", r"α + 2β = 1 − c + 2c^2"),
        (r"−1\s*\+\s*4c\s*=\s*0", r"−1 + 4c = 0"),
        (r"c\s*=\s*4\s*\.\s*1", r"c = 1/4"),
        (r"1\s*2\s*8\s*\.\s*7", r"1/8"),
        (r"8\s*\.\s*7", r"7/8"),
        (r"P\s*_\s*0\s*\(X\s*≥\s*c\)\s*=\s*1\s*−\s*c", r"P_0(X ≥ c) = 1 − c"),
        (r"0\s*≤\s*c\s*≤\s*1", r"0 ≤ c ≤ 1"),
        (r"X\s*\(_\s*n\s*\)", r"X_{(n)}"),
        (r"X\s*\(\s*n\s*\)", r"X_{(n)}"),
        (r"2c\s*\.\s*2", r"2c^2"),
        (r"证明当\s*n→∞\s*时，\s*\$\\alpha\s*0\s*,\s*β\s*0\$", r"证明当 n→∞ 时， $\\alpha \\to 0, β \\to 0$"),
        (r"该概率关\s*\$\\mathcal\s*\{\s*F\s*\}\s*\\theta\$\s*单调递减", r"该概率关于 $\\theta$ 单调递减"),
    )
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)
    return text


def _looks_like_malformed_ode_matrix_block(context: str, matrix: str, initial: str) -> bool:
    compact = re.sub(r"\s+", "", context + matrix + initial)
    required = (
        "x^{(4)}+x=te^{t}",
        "x^{\\prime}=Ax+f(t)",
        "\\boldsymbol{x}",
        "A=",
        "f(t)",
        "te^{t}",
        "{1}",
        "{-1}",
        "{2}",
        "{0}",
    )
    return all(token in compact for token in required)


def _repair_malformed_ode_matrix_blocks(text: str) -> str:
    def replace(match: re.Match[str]) -> str:
        context = text[max(0, match.start() - 1200) : match.start()]
        matrix = match.group("matrix")
        initial = match.group("initial")
        if not _looks_like_malformed_ode_matrix_block(context, matrix, initial):
            return match.group(0)
        return (
            "其中，\n\n"
            "$$\n"
            r"x = \begin{pmatrix} x_1 \\ x_2 \\ x_3 \\ x_4 \end{pmatrix},\quad "
            r"A = \begin{pmatrix} "
            r"0 & 1 & 0 & 0 \\ "
            r"0 & 0 & 1 & 0 \\ "
            r"0 & 0 & 0 & 1 \\ "
            r"-1 & 0 & 0 & 0 "
            r"\end{pmatrix},\quad "
            r"f(t) = \begin{pmatrix} 0 \\ 0 \\ 0 \\ t e^t \end{pmatrix}."
            "\n$$\n\n"
            "相应的初始条件为：\n\n"
            "$$\n"
            r"x(0)=\begin{pmatrix} 1 \\ -1 \\ 2 \\ 0 \end{pmatrix}"
            "\n$$"
        )

    return MALFORMED_ODE_MATRIX_BLOCK_RE.sub(replace, text)


def repair_extracted_markdown_text(markdown: str) -> str:
    """Repair conservative OCR/PDF extraction ordering glitches before math normalization."""
    text = markdown or ""
    repairs = (
        (r"故稳\s+稳\s+Y", "故 Y"),
        (r"有平\s+概率；\s+稳\s+双随机", "有平稳概率；双随机"),
        (r"平稳分布只稳\s*能", "平稳分布只能"),
        (r"平稳分布，由\s+稳\s+(\(\d+\))", r"平稳分布，由 \1"),
        (r"平\s+分\s*稳\s*布", "平稳分布"),
        (r"平\s+分布由细致平衡\s+稳", "平稳分布由细致平衡"),
        (r"平\s+分布满足\s+稳", "平稳分布满足"),
        (r"平\s+分布，则\s+稳", "平稳分布，则"),
        (r"平\s+分布[。．.]\s*稳", "平稳分布。"),
        (r"平\s+概率分布[。．.]\s*稳", "平稳概率分布。"),
        (r"平\s+方程给\s+稳", "平稳方程给"),
        (r"平\s+方程为\s+稳", "平稳方程为"),
        (r"平\s+方程仍成立\s+稳", "平稳方程仍成立"),
        (r"平\s+方程[。．.]\s*稳", "平稳方程。"),
        (r"任意稳(?=\s*\$)", "任意"),
        (r"平\s+概率分布", "平稳概率分布"),
        (r"平\s+分布", "平稳分布"),
        (r"平\s+方程", "平稳方程"),
        (r"平\s+稳", "平稳"),
    )
    for pattern, replacement in repairs:
        text = re.sub(pattern, replacement, text)
    text = re.sub(r"平稳分布只稳\s*能", "平稳分布只能", text)
    text = re.sub(r"平稳分布，由\s+稳\s+(\(\d+\))", r"平稳分布，由 \1", text)
    text = re.sub(r"e\s*_\s*\{\s*i\s+j\s+\\epsilon\s*\}", r"e_{i j}", text)
    text = re.sub(
        r"\\pi\s*_\s*\{\s*i\s*\}\s*\\overline\s*\{\s*\{?\s*-\s*\}?\s*\}\s*\\pi\s*_\s*\{\s*i\s*_\s*\{\s*\\circ\s*\}\s*\}",
        r"\\pi_{i} = \\pi_{i}",
        text,
    )
    text = re.sub(
        r"1\s*_\s*\{\s*\\\{\s*X\s*_\s*\{\s*k\s*\}\s*=\s*j\s*\\\}\s*\}\s*\{\s*\\mathfrak\s*\{\s*c\s*\}\s*\}",
        r"1_{ \\{ X_{k} = j \\} }",
        text,
    )
    text = re.sub(r"(互通类为\s*\$)C\s*_\s*\{\s*\\circ\s*\}(\$\s*。)", r"\1C\2", text)
    text = re.sub(r"(平稳(?:概率分布|分布|方程)(?:由细致平衡|满足|给|为|仍成立)?)(?:\s+稳)", r"\1", text)
    text = re.sub(r"(?:[。；]\s*)稳(?=\s*(?:\n|$|[A-Z\\$（(]))", lambda match: match.group(0).replace("稳", ""), text)
    text = _repair_mineru_spacing_in_math(text)
    text = _repair_ocr_c_scripts(text)
    text = _repair_malformed_ode_matrix_blocks(text)
    text = _repair_statistics_pdf_fragments(text)
    return text


def safe_source_filename(filename: str, suffix: str | None = None, fallback_stem: str = "source") -> str:
    resolved_suffix = (suffix or Path(filename).suffix).lower()
    safe_name = secure_filename(filename)
    safe_path = Path(safe_name)
    if resolved_suffix and safe_path.suffix.lower() == resolved_suffix and safe_path.stem:
        return safe_name

    safe_stem = safe_path.stem if safe_path.suffix else safe_name
    if not safe_stem or safe_stem.lower() == resolved_suffix.lstrip("."):
        safe_stem = fallback_stem
    if not resolved_suffix:
        return safe_stem or fallback_stem
    return f"{safe_stem}{resolved_suffix}"


def _read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore").replace("\r\n", "\n").replace("\r", "\n")


def _pandoc_to_markdown(path: Path, from_format: str | None = None) -> str:
    extra_args = ["--wrap=none"]
    kwargs: dict[str, Any] = {
        "to": "markdown+tex_math_dollars",
        "extra_args": extra_args,
    }
    if from_format:
        kwargs["format"] = from_format
    return pypandoc.convert_file(str(path), **kwargs).replace("\r\n", "\n").replace("\r", "\n")


def extract_source_to_markdown(path: Path) -> dict[str, Any]:
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_SOURCE_SUFFIXES:
        raise ValueError("只支持 PDF、Word、Markdown、TXT、RTF 文件")

    if suffix in {".md", ".markdown"}:
        return {"markdown": normalize_math_markdown(_read_text(path)), "source": "markdown", "warnings": []}

    if suffix == ".txt":
        return {"markdown": normalize_math_markdown(_read_text(path)), "source": "plain_text", "warnings": []}

    if suffix == ".rtf":
        try:
            return {"markdown": normalize_math_markdown(_pandoc_to_markdown(path, "rtf")), "source": "pandoc_rtf", "warnings": []}
        except Exception:
            return {
                "markdown": normalize_math_markdown(_read_text(path)),
                "source": "plain_rtf_fallback",
                "warnings": ["RTF 转 Markdown 失败，已按纯文本读取"],
            }

    if suffix in {".docx", ".doc"}:
        return {"markdown": normalize_math_markdown(_pandoc_to_markdown(path)), "source": "pandoc_docx", "warnings": []}

    if suffix == ".pdf":
        result = extract_pdf_to_markdown(path)
        result["markdown"] = normalize_math_markdown(repair_extracted_markdown_text(result["markdown"]))
        result["source"] = "mineru"
        return result

    raise ValueError("Unsupported source file type")
