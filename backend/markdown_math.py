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
    "Longleftrightarrow",
    "Longrightarrow",
    "Longleftarrow",
    "longleftrightarrow",
    "longrightarrow",
    "longleftarrow",
    "Leftrightarrow",
    "leftrightarrow",
    "xrightarrow",
    "xleftarrow",
    "twoheadrightarrow",
    "twoheadleftarrow",
    "rightsquigarrow",
    "rightleftharpoons",
    "leftrightharpoons",
    "rightharpoondown",
    "rightharpoonup",
    "leftharpoondown",
    "leftharpoonup",
    "hookrightarrow",
    "hookleftarrow",
    "nearrow",
    "searrow",
    "nwarrow",
    "swarrow",
    "operatorname",
    "operatornamewithlimits",
    "displaystyle",
    "overrightarrow",
    "underrightarrow",
    "underleftarrow",
    "overleftarrow",
    "multicolumn",
    "smallmatrix",
    "textcolor",
    "underparen",
    "overparen",
    "Rightarrow",
    "Leftarrow",
    "rightarrow",
    "leftarrow",
    "varnothing",
    "leqslant",
    "geqslant",
    "lessapprox",
    "gtrapprox",
    "lesssim",
    "gtrsim",
    "leqsim",
    "geqsim",
    "boldsymbol",
    "overline",
    "underline",
    "textstyle",
    "scriptstyle",
    "therefore",
    "smallsetminus",
    "rightthreetimes",
    "leftthreetimes",
    "blacksquare",
    "nonumber",
    "substack",
    "subarray",
    "parallel",
    "because",
    "partial",
    "setminus",
    "nsubseteq",
    "nsupseteq",
    "subseteq",
    "supseteq",
    "subsetneq",
    "supsetneq",
    "emptyset",
    "lVert",
    "rVert",
    "Vert",
    "mathfrak",
    "mathscr",
    "mathcal",
    "mathds",
    "mathbb",
    "boldmath",
    "pmb",
    "Bbb",
    "cal",
    "mathrel",
    "mathbin",
    "mathord",
    "mathopen",
    "mathclose",
    "mathpunct",
    "mathinner",
    "mathclap",
    "raisebox",
    "stackrel",
    "buildrel",
    "genfrac",
    "hdotsfor",
    "mathrm",
    "mathbf",
    "mathop",
    "mathsf",
    "mathtt",
    "mathit",
    "textrm",
    "textnormal",
    "textit",
    "textup",
    "textsl",
    "texttt",
    "textsf",
    "textbf",
    "aleph",
    "backprime",
    "daleth",
    "gimel",
    "beth",
    "Game",
    "Finv",
    "mho",
    "eth",
    "atop",
    "overwithdelims",
    "atopwithdelims",
    "abovewithdelims",
    "eqref",
    "hbar",
    "label",
    "ell",
    "notag",
    "middle",
    "hphantom",
    "vphantom",
    "phantom",
    "smash",
    "choose",
    "multirow",
    "mapsto",
    "implies",
    "notin",
    "pmod",
    "brack",
    "brace",
    "above",
    "bmod",
    "preceq",
    "succeq",
    "nless",
    "ngtr",
    "nleq",
    "ngeq",
    "lneq",
    "gneq",
    "lnsim",
    "gnsim",
    "leqq",
    "geqq",
    "lessgtr",
    "gtrless",
    "models",
    "vdash",
    "dashv",
    "vDash",
    "VDash",
    "Vdash",
    "Vvdash",
    "nvdash",
    "nvDash",
    "nVdash",
    "nVDash",
    "smile",
    "frown",
    "bowtie",
    "bigcup",
    "bigcap",
    "bigsqcup",
    "bigvee",
    "bigwedge",
    "bigoplus",
    "bigotimes",
    "boxtimes",
    "boxminus",
    "boxplus",
    "boxdot",
    "Subset",
    "Supset",
    "sqsubseteq",
    "sqsupseteq",
    "sqsubset",
    "sqsupset",
    "limsup",
    "liminf",
    "injlim",
    "projlim",
    "cfrac",
    "dfrac",
    "tfrac",
    "begin",
    "right",
    "left",
    "sqrt",
    "boxed",
    "color",
    "grave",
    "acute",
    "breve",
    "check",
    "fbox",
    "hline",
    "cline",
    "mbox",
    "hbox",
    "rlap",
    "llap",
    "sout",
    "bcancel",
    "xcancel",
    "cancel",
    "frac",
    "mathring",
    "iiint",
    "prod",
    "oint",
    "iint",
    "lim",
    "ref",
    "tag",
    "pod",
    "sum",
    "int",
    "forall",
    "exists",
    "nabla",
    "approx",
    "equiv",
    "asymp",
    "doteq",
    "subset",
    "supset",
    "propto",
    "triangle",
    "triangleleft",
    "triangleright",
    "times",
    "cdot",
    "cdotp",
    "ldotp",
    "bullet",
    "diamond",
    "Box",
    "square",
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
    "Xi",
    "Pi",
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
    "varrho",
    "tau",
    "iota",
    "kappa",
    "upsilon",
    "varpi",
    "varsigma",
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
    "sqcup",
    "sqcap",
    "oplus",
    "otimes",
    "uplus",
    "wr",
    "amalg",
    "ltimes",
    "rtimes",
    "neq",
    "neg",
    "lnot",
    "not",
    "leq",
    "geq",
    "sin",
    "cos",
    "tan",
    "arcsin",
    "arccos",
    "arctan",
    "sinh",
    "cosh",
    "tanh",
    "cot",
    "sec",
    "csc",
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
    "prec",
    "succ",
    "ll",
    "gg",
    "ln",
    "bm",
    "thinspace",
    "medspace",
    "thickspace",
    "negthinspace",
    "negmedspace",
    "negthickspace",
    "Re",
    "Im",
    "wp",
)
LATEX_NAMED_COMMAND_RE = re.compile(r"\\(?:" + "|".join(LATEX_COMMAND_NAMES) + r")(?![A-Za-z])")
LATEX_RESIDUAL_RE = re.compile(r"\\(?:[A-Za-z]+|[,;:! ]|[{}_^~'\"`|])|(?<!\\)\$")
INLINE_COMMAND_RE = LATEX_NAMED_COMMAND_RE
DISPLAY_COMMAND_RE = LATEX_NAMED_COMMAND_RE
MATH_RELATION_RE = re.compile(
    r"[=<>≤≥≠≈≡∈∉⊥∥→←⇒⇐↔⇔⊂⊆⊃⊇∝]|"
    r"\\(?:leqslant|geqslant|leq|geq|le|ge|neq|ne|approx|sim|cong|simeq|equiv|"
    r"pmod|bmod|pod|perp|parallel|in|notin|ni|subset|subseteq|nsubseteq|subsetneq|"
    r"supset|supseteq|nsupseteq|supsetneq|Subset|Supset|sqsubset|sqsupset|sqsubseteq|sqsupseteq|propto|"
    r"prec|preceq|succ|succeq|nleq|ngeq|nless|ngtr|lneq|gneq|lnsim|gnsim|"
    r"leqq|geqq|leqsim|geqsim|lesssim|gtrsim|lessapprox|gtrapprox|"
    r"lessgtr|gtrless|ll|gg|asymp|doteq|models|vdash|dashv|vDash|Vdash|VDash|Vvdash|"
    r"nvdash|nvDash|nVdash|nVDash|"
    r"smallsetminus|sqcup|sqcap|uplus|wr|amalg|boxtimes|boxplus|boxminus|boxdot|ltimes|rtimes|leftthreetimes|rightthreetimes|"
    r"bigcup|bigcap|bigsqcup|bigvee|bigwedge|bigoplus|bigotimes|"
    r"to|rightarrow|leftarrow|xrightarrow|xleftarrow|hookrightarrow|hookleftarrow|twoheadrightarrow|twoheadleftarrow|rightsquigarrow|"
    r"leftharpoonup|leftharpoondown|rightharpoonup|rightharpoondown|leftrightharpoons|rightleftharpoons|nearrow|searrow|nwarrow|swarrow|"
    r"Rightarrow|Leftarrow|Longrightarrow|Longleftarrow|longrightarrow|longleftarrow|longleftrightarrow|"
    r"leftrightarrow|Leftrightarrow|Longleftrightarrow|mapsto|implies|middle)(?![A-Za-z])"
)
INLINE_STRUCTURAL_COMMAND_RE = re.compile(
    r"\\(?:frac|dfrac|tfrac|cfrac|sqrt|binom|pmod|partial|xrightarrow|xleftarrow|substack|"
    r"boxed|fbox|cancel|bcancel|xcancel|sout|color|textcolor|multicolumn|multirow|"
    r"hline|cline|hdotsfor|acute|grave|breve|check|mathring|mathscr|mathds|bm|Re|Im|ell|hbar|aleph|wp|"
    r"beth|gimel|daleth|mho|Game|Finv|backprime|eth|"
    r"stackrel|buildrel|genfrac|overwithdelims|atopwithdelims|abovewithdelims|"
    r"mathrel|mathbin|mathord|mathopen|mathclose|mathpunct|mathinner|"
    r"pmb|boldmath|cal|Bbb|operatornamewithlimits|textnormal|textit|textup|textsl|texttt|textsf|"
    r"smash|rlap|llap|mathclap|raisebox|"
    r"varpi|varsigma|varrho|bullet|diamond|Box|square|blacksquare|triangleleft|triangleright|"
    r"thinspace|medspace|thickspace|negthinspace|negmedspace|negthickspace|"
    r"bmod|pod|choose|atop|brack|brace|above|phantom|hphantom|vphantom|mbox|hbox|limsup|liminf|injlim|projlim|eqref|ref)(?![A-Za-z])"
)
TEXT_MATH_BOUNDARY_RE = re.compile(r"[\u4e00-\u9fff，。；：！？、]")
STYLE_COMMAND_RE = re.compile(r"\\(?:displaystyle|textstyle|scriptstyle|scriptscriptstyle)\b")
TEXT_GROUP_RE = re.compile(r"\\(?:text|textrm|textbf)\s*\{[^{}]*\}")
BARE_BUILDREL_RE = re.compile(r"\\buildrel\s+.+?\s+\\over\s+(?:\\[A-Za-z]+|[^\s\u4e00-\u9fff，。；：！？、]+)")


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
    protected_text_groups: list[str] = []

    def protect_text_group(match: re.Match[str]) -> str:
        protected_text_groups.append(match.group(0))
        return f"@@TEXTGROUP{len(protected_text_groups) - 1}@@"

    expr = TEXT_GROUP_RE.sub(protect_text_group, expr)
    expr = expr.replace("\\dots", "\\ldots")
    expr = expr.replace("\\dotsc", "\\ldots").replace("\\dotsb", "\\ldots").replace("\\dotso", "\\ldots")
    expr = re.sub(r"\s*([_^])\s*", r"\1", expr)
    expr = re.sub(r"\{\s*([^{}\n]+?)\s*\}", r"{\1}", expr)
    for index, text_group in enumerate(protected_text_groups):
        expr = expr.replace(f"@@TEXTGROUP{index}@@", text_group)
    expr = re.sub(r"\\(?:,|;|:|!)", " ", expr)
    expr = re.sub(r"[ \t]+", " ", expr)
    expr = re.sub(r"\n{3,}", "\n\n", expr)
    return expr.strip()


def _is_bare_math_char(ch: str) -> bool:
    if TEXT_MATH_BOUNDARY_RE.match(ch) or ch in "$":
        return False
    return ch.isascii() and (ch.isalnum() or ch.isspace() or ch in "\\{}[]()_^+-=*/<>.,;:|~'\"!&")


def _bare_math_span_end(text: str, end: int) -> int:
    i = end
    brace_depth = 0
    while i < len(text):
        ch = text[i]
        if not _is_bare_math_char(ch):
            break
        if ch == "{":
            brace_depth += 1
        elif ch == "}":
            brace_depth = max(0, brace_depth - 1)
        elif ch.isspace() and not brace_depth:
            next_nonspace = i
            while next_nonspace < len(text) and text[next_nonspace].isspace():
                next_nonspace += 1
            tail = text[next_nonspace:]
            if re.match(r"[A-Za-z]{2,}(?:\s|$|[.。])", tail):
                break
        i += 1
    return i


def _trim_bare_math_expr(expr: str) -> str:
    expr = expr.strip()
    while True:
        trimmed = re.sub(r"\s+[A-Za-z]{2,}\.?$", "", expr).rstrip()
        if trimmed == expr:
            return expr
        expr = trimmed


def _looks_like_inline_bare_math(expr: str) -> bool:
    if not expr or not LATEX_NAMED_COMMAND_RE.search(expr):
        return False
    command_count = len(LATEX_NAMED_COMMAND_RE.findall(expr))
    operator_count = len(re.findall(r"[+\-*/=<>]", expr))
    return bool(
        MATH_RELATION_RE.search(expr)
        or INLINE_STRUCTURAL_COMMAND_RE.search(expr)
        or len(re.findall(r"[_^]\s*(?:\{|[A-Za-z0-9])", expr)) >= 1
        or (command_count >= 1 and operator_count >= 1)
    )


def _wrap_bare_latex_spans(text: str) -> str:
    result: list[str] = []
    pos = 0
    while pos < len(text):
        match = LATEX_NAMED_COMMAND_RE.search(text, pos)
        if not match:
            result.append(text[pos:])
            break
        if text.startswith("\\buildrel", match.start()):
            buildrel_match = BARE_BUILDREL_RE.match(text, match.start())
            if buildrel_match:
                start = match.start()
                expr = buildrel_match.group(0).strip()
                result.append(text[pos:start])
                result.append(f"${normalize_latex_math(expr)}$")
                pos = buildrel_match.end()
                continue
        start = match.start()
        matched_command = match.group(0)[1:]
        if matched_command in {"above", "overwithdelims", "atopwithdelims", "abovewithdelims"}:
            while start > pos and text[start - 1].isspace():
                start -= 1
            while start > pos and not text[start - 1].isspace() and _is_bare_math_char(text[start - 1]):
                start -= 1
        while start > pos and not text[start - 1].isspace() and _is_bare_math_char(text[start - 1]):
            start -= 1
        end = _bare_math_span_end(text, match.end())
        expr = _trim_bare_math_expr(text[start:end])
        if not _looks_like_inline_bare_math(expr):
            result.append(text[pos:match.end()])
            pos = match.end()
            continue
        expr_start = text.find(expr, start, end)
        expr_end = expr_start + len(expr)
        result.append(text[pos:expr_start])
        result.append(f"${normalize_latex_math(expr)}$")
        pos = expr_end
    return "".join(result)


def _normalize_bare_inline_math(text: str) -> str:
    parts: list[str] = []
    pos = 0
    for match in INLINE_DOLLAR_RE.finditer(text):
        parts.append(_wrap_bare_latex_spans(text[pos:match.start()]))
        parts.append(match.group(0))
        pos = match.end()
    parts.append(_wrap_bare_latex_spans(text[pos:]))
    return "".join(parts)


def _normalize_inline_math(text: str) -> str:
    text = INLINE_PAREN_RE.sub(lambda m: f"${normalize_latex_math(m.group(1))}$", text)
    text = INLINE_DOLLAR_RE.sub(lambda m: f"${normalize_latex_math(m.group(1))}$", text)
    return _normalize_bare_inline_math(text)



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
