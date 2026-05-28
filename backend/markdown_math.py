from __future__ import annotations

import html
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
    "xleftrightarrow",
    "xRightarrow",
    "xLeftarrow",
    "xLeftrightarrow",
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
    "downarrow",
    "uparrow",
    "nolimits",
    "overrightarrow",
    "overleftrightarrow",
    "underrightarrow",
    "underleftarrow",
    "overleftarrow",
    "multicolumn",
    "underbrace",
    "overbrace",
    "smallmatrix",
    "textcolor",
    "underparen",
    "overparen",
    "widetilde",
    "widehat",
    "smallmatrix",
    "pmatrix",
    "bmatrix",
    "Bmatrix",
    "vmatrix",
    "Vmatrix",
    "matrix",
    "cases",
    "Biggl",
    "Biggr",
    "biggl",
    "biggr",
    "Rightarrow",
    "Leftarrow",
    "rightarrow",
    "leftarrow",
    "varnothing",
    "leqslant",
    "geqslant",
    "nparallel",
    "smallfrown",
    "smallsmile",
    "blacklozenge",
    "blacktriangle",
    "lessapprox",
    "gtrapprox",
    "lesssim",
    "gtrsim",
    "precsim",
    "succsim",
    "leqsim",
    "geqsim",
    "coloneqq",
    "eqqcolon",
    "triangleq",
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
    "napprox",
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
    "underset",
    "overset",
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
    "text",
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
    "limits",
    "hspace",
    "vspace",
    "enspace",
    "lbrace",
    "rbrace",
    "lparen",
    "rparen",
    "lbrack",
    "rbrack",
    "langle",
    "rangle",
    "lfloor",
    "rfloor",
    "lceil",
    "rceil",
    "lvert",
    "rvert",
    "hphantom",
    "vphantom",
    "phantom",
    "smash",
    "choose",
    "Bigg",
    "bigg",
    "Bigl",
    "Bigr",
    "bigl",
    "bigr",
    "Big",
    "big",
    "multirow",
    "mapsto",
    "longmapsto",
    "leadsto",
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
    "ncong",
    "nsim",
    "sim",
    "nmid",
    "nsmile",
    "nfrown",
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
    "degree",
    "prime",
    "bigcup",
    "bigcap",
    "bigsqcup",
    "bigvee",
    "bigwedge",
    "bigoplus",
    "bigotimes",
    "triangledown",
    "diamondsuit",
    "heartsuit",
    "spadesuit",
    "clubsuit",
    "lozenge",
    "bigstar",
    "bigcirc",
    "natural",
    "diagdown",
    "diagup",
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
    "dbinom",
    "tbinom",
    "binom",
    "begin",
    "end",
    "right",
    "left",
    "sqrt",
    "boxed",
    "color",
    "grave",
    "acute",
    "breve",
    "check",
    "tilde",
    "hat",
    "bar",
    "ddot",
    "dot",
    "vec",
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
    "coprod",
    "oint",
    "iint",
    "lim",
    "ref",
    "tag",
    "pod",
    "over",
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
    "circ",
    "star",
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
    "vert",
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
    "varDelta",
    "varGamma",
    "varTheta",
    "varLambda",
    "varOmega",
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
    "flat",
    "sharp",
    "top",
    "bot",
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
    "iff",
    "to",
    "ni",
    "in",
    "ne",
    "pm",
    "Pr",
    "arg",
    "det",
    "dim",
    "ker",
    "gcd",
    "min",
    "max",
    "sup",
    "inf",
    "mod",
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
    "lessdot",
    "gtrdot",
    "lll",
    "ggg",
    "nleqslant",
    "ngeqslant",
    "lneqq",
    "gneqq",
    "subsetneqq",
    "supsetneqq",
    "varsubsetneq",
    "varsupsetneq",
    "nexists",
    "complement",
    "Bbbk",
    "ulcorner",
    "urcorner",
    "llcorner",
    "lrcorner",
    "curlywedge",
    "curlyvee",
    "Cap",
    "Cup",
    "circledast",
    "circledcirc",
    "circleddash",
    "blacktriangleright",
    "blacktriangleleft",
    "trianglerighteq",
    "trianglelefteq",
    "maltese",
)
LATEX_NAMED_COMMAND_RE = re.compile(r"\\(?:" + "|".join(LATEX_COMMAND_NAMES) + r")(?![A-Za-z])")
LATEX_RESIDUAL_RE = re.compile(r"\\(?:[A-Za-z]+|[,;:! ]|[{}_^~'\"`|])|(?<!\\)\$")
INLINE_COMMAND_RE = LATEX_NAMED_COMMAND_RE
DISPLAY_COMMAND_RE = LATEX_NAMED_COMMAND_RE
MATH_RELATION_RE = re.compile(
    r"[=<>≤≥≠≈≡∈∉⊥∥→←⇒⇐↔⇔⊂⊆⊃⊇∝]|"
    r"\\(?:leqslant|geqslant|leq|geq|le|ge|neq|ne|approx|sim|cong|simeq|equiv|"
    r"pmod|bmod|pod|perp|parallel|nparallel|in|notin|ni|subset|subseteq|nsubseteq|subsetneq|"
    r"supset|supseteq|nsupseteq|supsetneq|Subset|Supset|sqsubset|sqsupset|sqsubseteq|sqsupseteq|propto|"
    r"prec|preceq|succ|succeq|nleq|ngeq|nless|ngtr|lneq|gneq|lnsim|gnsim|ncong|napprox|nsim|nmid|nsmile|nfrown|smallsmile|smallfrown|"
    r"leqq|geqq|leqsim|geqsim|lesssim|gtrsim|lessapprox|gtrapprox|precsim|succsim|"
    r"lessdot|gtrdot|lll|ggg|nleqslant|ngeqslant|lneqq|gneqq|subsetneqq|supsetneqq|varsubsetneq|varsupsetneq|"
    r"lessgtr|gtrless|ll|gg|asymp|doteq|models|vdash|dashv|vDash|Vdash|VDash|Vvdash|"
    r"nvdash|nvDash|nVdash|nVDash|"
    r"smallsetminus|diagup|diagdown|sqcup|sqcap|uplus|wr|amalg|boxtimes|boxplus|boxminus|boxdot|ltimes|rtimes|leftthreetimes|rightthreetimes|"
    r"bigcup|bigcap|bigsqcup|bigvee|bigwedge|bigoplus|bigotimes|"
    r"to|rightarrow|leftarrow|xrightarrow|xleftarrow|xleftrightarrow|xRightarrow|xLeftarrow|xLeftrightarrow|"
    r"hookrightarrow|hookleftarrow|twoheadrightarrow|twoheadleftarrow|rightsquigarrow|leadsto|"
    r"leftharpoonup|leftharpoondown|rightharpoonup|rightharpoondown|leftrightharpoons|rightleftharpoons|nearrow|searrow|nwarrow|swarrow|"
    r"Rightarrow|Leftarrow|Longrightarrow|Longleftarrow|longrightarrow|longleftarrow|longleftrightarrow|"
    r"leftrightarrow|Leftrightarrow|Longleftrightarrow|mapsto|longmapsto|implies|"
    r"coloneqq|eqqcolon|triangleq|nexists|complement|curlywedge|curlyvee|Cap|Cup|"
    r"circledast|circledcirc|circleddash|blacktriangleright|blacktriangleleft|trianglerighteq|trianglelefteq|maltese|"
    r"middle)(?![A-Za-z])"
)
INLINE_STRUCTURAL_COMMAND_RE = re.compile(
    r"\\(?:frac|dfrac|tfrac|cfrac|sqrt|binom|dbinom|tbinom|pmod|partial|xrightarrow|xleftarrow|"
    r"xleftrightarrow|xRightarrow|xLeftarrow|xLeftrightarrow|substack|"
    r"boxed|fbox|cancel|bcancel|xcancel|sout|color|textcolor|multicolumn|multirow|"
    r"hline|cline|hdotsfor|acute|grave|breve|check|mathring|hat|widehat|bar|tilde|widetilde|vec|dot|ddot|overleftrightarrow|"
    r"overline|underline|boldsymbol|boldmath|textrm|textnormal|textup|textsl|hbox|"
    r"mathscr|mathds|bm|Re|Im|ell|hbar|aleph|wp|"
    r"lbrace|rbrace|lparen|rparen|lbrack|rbrack|langle|rangle|lfloor|rfloor|lceil|rceil|lvert|rvert|vert|"
    r"big|Big|bigg|Bigg|bigl|bigr|Bigl|Bigr|biggl|biggr|Biggl|Biggr|limits|nolimits|"
    r"hspace|vspace|quad|qquad|enspace|thinspace|medspace|thickspace|negthinspace|negmedspace|negthickspace|"
    r"beth|gimel|daleth|mho|Game|Finv|backprime|eth|"
    r"coprod|bigstar|bigcirc|lozenge|blacklozenge|blacktriangle|triangledown|"
    r"clubsuit|diamondsuit|heartsuit|spadesuit|natural|flat|sharp|top|bot|"
    r"iff|circ|star|degree|prime|uparrow|downarrow|min|max|sup|inf|arg|det|dim|ker|gcd|mod|"
    r"overset|underset|stackrel|buildrel|overbrace|underbrace|genfrac|overwithdelims|atopwithdelims|abovewithdelims|"
    r"mathrel|mathbin|mathord|mathopen|mathclose|mathpunct|mathinner|"
    r"pmb|boldmath|cal|Bbb|operatornamewithlimits|textnormal|textit|textup|textsl|texttt|textsf|"
    r"smash|rlap|llap|mathclap|raisebox|"
    r"varpi|varsigma|varrho|varDelta|varGamma|varTheta|varLambda|varOmega|"
    r"lessdot|gtrdot|lll|ggg|nleqslant|ngeqslant|lneqq|gneqq|subsetneqq|supsetneqq|varsubsetneq|varsupsetneq|"
    r"nexists|complement|Bbbk|ulcorner|urcorner|llcorner|lrcorner|curlywedge|curlyvee|Cap|Cup|"
    r"circledast|circledcirc|circleddash|blacktriangleright|blacktriangleleft|trianglerighteq|trianglelefteq|maltese|"
    r"bullet|diamond|Box|square|blacksquare|triangleleft|triangleright|"
    r"thinspace|medspace|thickspace|negthinspace|negmedspace|negthickspace|"
    r"bmod|pod|over|choose|atop|brack|brace|above|phantom|hphantom|vphantom|mbox|hbox|limsup|liminf|injlim|projlim|"
    r"eqref|ref|tag|label|notag|nonumber|text)(?![A-Za-z])"
)
TEXT_MATH_BOUNDARY_RE = re.compile(r"[\u4e00-\u9fff，。；：！？、]")
STYLE_COMMAND_RE = re.compile(r"\\(?:displaystyle|textstyle|scriptstyle|scriptscriptstyle)\b")
TEXT_GROUP_RE = re.compile(r"\\(?:text|textrm|textbf)\s*\{[^{}]*\}")
BARE_BUILDREL_RE = re.compile(r"\\buildrel\s+.+?\s+\\over\s+(?:\\[A-Za-z]+|[^\s\u4e00-\u9fff，。；：！？、]+)")
BUILDREL_RE = re.compile(r"\\buildrel\s+(.+?)\s+\\over\s+(\{[^{}]*\}|\\[A-Za-z]+|[A-Za-z0-9]+|[^\s])")
LEGACY_INFIX_COMMANDS = {
    "over",
    "choose",
    "above",
    "atop",
    "brack",
    "brace",
    "overwithdelims",
    "atopwithdelims",
    "abovewithdelims",
}
PANDOC_SAFE_SYMBOL_REPLACEMENTS = {
    "degree": "°",
    "setminus": "∖",
    "diagup": "⟋",
    "diagdown": "⟍",
    "ldotp": ".",
    "cdotp": "·",
    "leqsim": "⪅",
    "geqsim": "⪆",
    "lessapprox": "⪅",
    "gtrapprox": "⪆",
    "precsim": "≾",
    "succsim": "≿",
    "coloneqq": "≔",
    "eqqcolon": "≕",
    "triangleq": "≜",
    "leadsto": "↝",
    "longmapsto": "⟼",
    "varDelta": "Δ",
    "varGamma": "Γ",
    "varTheta": "Θ",
    "varLambda": "Λ",
    "varOmega": "Ω",
    "lessdot": "⋖",
    "gtrdot": "⋗",
    "lll": "⋘",
    "ggg": "⋙",
    "nleqslant": "≰",
    "ngeqslant": "≱",
    "lneqq": "≨",
    "gneqq": "≩",
    "subsetneqq": "⫋",
    "supsetneqq": "⫌",
    "varsubsetneq": "⊊",
    "varsupsetneq": "⊋",
    "nexists": "∄",
    "complement": "∁",
    "Bbbk": "𝕜",
    "ulcorner": "⌜",
    "urcorner": "⌝",
    "llcorner": "⌞",
    "lrcorner": "⌟",
    "curlywedge": "⋏",
    "curlyvee": "⋎",
    "Cap": "⋒",
    "Cup": "⋓",
    "circledast": "⊛",
    "circledcirc": "⊚",
    "circleddash": "⊝",
    "blacktriangleright": "▸",
    "blacktriangleleft": "◂",
    "trianglerighteq": "⊵",
    "trianglelefteq": "⊴",
    "maltese": "✠",
    "nsmile": "¬⌣",
    "nfrown": "¬⌢",
}
PANDOC_SAFE_WRAPPER_COMMANDS = (
    "llap",
    "rlap",
    "mathclap",
    "mathinner",
    "smash",
)
PLAIN_TEX_MATRIX_COMMANDS = {
    "smallmatrix": "smallmatrix",
    "matrix": "matrix",
    "pmatrix": "pmatrix",
    "bmatrix": "bmatrix",
    "Bmatrix": "Bmatrix",
    "vmatrix": "vmatrix",
    "Vmatrix": "Vmatrix",
    "cases": "cases",
}
PLAIN_TEX_MATRIX_COMMAND_RE = re.compile(r"\\(" + "|".join(PLAIN_TEX_MATRIX_COMMANDS) + r")(?![A-Za-z])")
UNKNOWN_LATEX_COMMAND_RE = re.compile(r"(?<!\\)\\([A-Za-z]+)(?![A-Za-z])")


def _read_balanced_brace_group(text: str, pos: int) -> tuple[str, int] | None:
    pos = _skip_ws(text, pos)
    if pos >= len(text) or text[pos] != "{":
        return None
    depth = 0
    start = pos + 1
    pos += 1
    while pos < len(text):
        ch = text[pos]
        if ch == "\\" and pos + 1 < len(text) and text[pos + 1] in "{}":
            pos += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            if depth == 0:
                return text[start:pos], pos + 1
            depth -= 1
        pos += 1
    return None


def _rewrite_unknown_latex_commands(expr: str, *, math_mode: bool) -> str:
    result: list[str] = []
    pos = 0
    while pos < len(expr):
        match = UNKNOWN_LATEX_COMMAND_RE.search(expr, pos)
        if not match:
            result.append(expr[pos:])
            break
        name = match.group(1)
        result.append(expr[pos:match.start()])
        if name in LATEX_COMMAND_NAMES:
            result.append(match.group(0))
            pos = match.end()
            continue
        group = _read_balanced_brace_group(expr, match.end())
        if group:
            content, end = group
            content = _rewrite_unknown_latex_commands(content, math_mode=math_mode)
            if math_mode:
                result.append(rf"\operatorname{{{name}}}({content})")
            else:
                result.append(f"{name}({content})")
            pos = end
            continue
        if math_mode:
            result.append(rf"\operatorname{{{name}}}")
        else:
            result.append(name)
        pos = match.end()
    return "".join(result)


def _unknown_latex_command_matches(expr: str) -> list[re.Match[str]]:
    return [match for match in UNKNOWN_LATEX_COMMAND_RE.finditer(expr) if match.group(1) not in LATEX_COMMAND_NAMES]


def _unknown_latex_command_has_group(expr: str) -> bool:
    return any(_read_balanced_brace_group(expr, match.end()) for match in _unknown_latex_command_matches(expr))


def _is_likely_path_escape(text: str, start: int) -> bool:
    if start > 0 and text[start - 1] in {":", "/", "\\"}:
        return True
    segment_start = max(text.rfind(ch, 0, start) for ch in " \t\r\n\"'()[]{}<>，。；：！？、") + 1
    segment_prefix = text[segment_start:start]
    return bool(re.search(r"[A-Za-z]:[\\/]", segment_prefix) or "/" in segment_prefix or "\\" in segment_prefix)


def _find_top_level_legacy_infix(expr: str) -> tuple[str, int, int] | None:
    depth = 0
    i = 0
    while i < len(expr):
        ch = expr[i]
        if ch == "\\":
            start = i
            i += 1
            name_start = i
            while i < len(expr) and expr[i].isalpha():
                i += 1
            if i == name_start and i < len(expr):
                i += 1
            name = expr[name_start:i]
            if depth == 0 and name in LEGACY_INFIX_COMMANDS and (i >= len(expr) or not expr[i].isalpha()):
                return name, start, i
            continue
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth = max(0, depth - 1)
        i += 1
    return None


def _skip_ws(text: str, pos: int) -> int:
    while pos < len(text) and text[pos].isspace():
        pos += 1
    return pos


def _read_latex_token(text: str, pos: int) -> tuple[str, int]:
    pos = _skip_ws(text, pos)
    if pos >= len(text):
        return "", pos
    if text[pos] == "\\":
        start = pos
        pos += 1
        while pos < len(text) and text[pos].isalpha():
            pos += 1
        if pos == start + 1 and pos < len(text):
            pos += 1
        return text[start:pos], pos
    return text[pos], pos + 1


def _read_dimension_token(text: str, pos: int) -> tuple[str, int]:
    pos = _skip_ws(text, pos)
    start = pos
    while pos < len(text) and not text[pos].isspace():
        pos += 1
    return text[start:pos], pos


def _latex_delimiter(token: str) -> str:
    if not token:
        return ""
    if token == ".":
        return "."
    return {
        "\\lbrace": "\\{",
        "\\rbrace": "\\}",
        "\\lparen": "(",
        "\\rparen": ")",
        "\\lbrack": "[",
        "\\rbrack": "]",
    }.get(token, token)


def _with_latex_delimiters(core: str, left: str, right: str) -> str:
    left = _latex_delimiter(left)
    right = _latex_delimiter(right)
    if not left and not right:
        return core
    return f"\\left{left}{core}\\right{right}"


def _rewrite_legacy_infix_math(expr: str) -> str:
    found = _find_top_level_legacy_infix(expr)
    if not found:
        return expr
    name, start, after = found
    left = expr[:start].strip()
    rest = expr[after:]
    if not left:
        return expr

    left_delimiter = ""
    right_delimiter = ""
    thickness = ""
    pos = 0
    if name in {"overwithdelims", "atopwithdelims", "abovewithdelims"}:
        left_delimiter, pos = _read_latex_token(rest, pos)
        right_delimiter, pos = _read_latex_token(rest, pos)
    if name in {"above", "abovewithdelims"}:
        thickness, pos = _read_dimension_token(rest, pos)

    right = rest[pos:].strip()
    if not right:
        return expr

    left = _rewrite_legacy_infix_math(left)
    right = _rewrite_legacy_infix_math(right)
    if name in {"over", "overwithdelims"}:
        return _with_latex_delimiters(f"\\frac{{{left}}}{{{right}}}", left_delimiter, right_delimiter)
    if name == "choose":
        return f"\\binom{{{left}}}{{{right}}}"
    if name in {"atop", "atopwithdelims", "brack", "brace"} or (
        name in {"above", "abovewithdelims"} and thickness.strip().startswith("0")
    ):
        core = f"\\substack{{{left}\\\\{right}}}"
        if name == "brack":
            return _with_latex_delimiters(core, "[", "]")
        if name == "brace":
            return _with_latex_delimiters(core, "\\{", "\\}")
        return _with_latex_delimiters(core, left_delimiter, right_delimiter)
    return _with_latex_delimiters(f"\\frac{{{left}}}{{{right}}}", left_delimiter, right_delimiter)


def _rewrite_buildrel(match: re.Match[str]) -> str:
    over = match.group(1).strip()
    base = match.group(2).strip()
    if base.startswith("{") and base.endswith("}"):
        base = base[1:-1]
    return f"\\overset{{{over}}}{{{base}}}"


def _rewrite_plain_tex_matrix_commands(expr: str) -> str:
    result: list[str] = []
    pos = 0
    while pos < len(expr):
        match = PLAIN_TEX_MATRIX_COMMAND_RE.search(expr, pos)
        if not match:
            result.append(expr[pos:])
            break
        group = _read_balanced_brace_group(expr, match.end())
        if group is None:
            result.append(expr[pos:match.end()])
            pos = match.end()
            continue
        content, end = group
        env = PLAIN_TEX_MATRIX_COMMANDS[match.group(1)]
        result.append(expr[pos:match.start()])
        result.append(rf"\begin{{{env}}}{_rewrite_plain_tex_matrix_commands(content)}\end{{{env}}}")
        pos = end
    return "".join(result)


def _rewrite_unsupported_presentation_helpers(expr: str) -> str:
    expr = _rewrite_plain_tex_matrix_commands(expr)
    expr = re.sub(r"\\cfrac(?![A-Za-z])", r"\\frac", expr)
    expr = re.sub(r"\\(?:dbinom|tbinom)(?![A-Za-z])", r"\\binom", expr)
    expr = re.sub(
        r"\\begin\s*\{subarray\}\s*\{[^{}]*\}(.*?)\\end\s*\{subarray\}",
        lambda match: r"\substack{" + match.group(1).strip() + "}",
        expr,
        flags=re.S,
    )
    expr = re.sub(
        r"\\begin\s*\{(alignedat\*?|alignat\*)\}\s*(?:\{[^{}]*\})?",
        lambda match: r"\begin{" + ("aligned" if match.group(1).startswith("aligned") else "align") + "}",
        expr,
    )
    expr = re.sub(
        r"\\end\s*\{(alignedat\*?|alignat\*)\}",
        lambda match: r"\end{" + ("aligned" if match.group(1).startswith("aligned") else "align") + "}",
        expr,
    )
    expr = re.sub(
        r"\\begin\s*\{array\}(?!\s*\{)(.*?)\\end\s*\{array\}",
        lambda match: r"\begin{matrix}" + match.group(1) + r"\end{matrix}",
        expr,
        flags=re.S,
    )
    for command, replacement in PANDOC_SAFE_SYMBOL_REPLACEMENTS.items():
        expr = re.sub(rf"\\{command}(?![A-Za-z])", replacement, expr)
    expr = re.sub(r"\\injlim(?![A-Za-z])", r"\\operatorname{inj lim}", expr)
    expr = re.sub(r"\\projlim(?![A-Za-z])", r"\\operatorname{proj lim}", expr)
    expr = re.sub(r"\\boldmath\s*\{([^{}]*)\}", r"\1", expr)
    expr = re.sub(r"\\cal\s*\{([^{}]*)\}", r"\\mathcal{\1}", expr)
    expr = re.sub(r"\\Bbb\s*\{([^{}]*)\}", r"\\mathbb{\1}", expr)
    expr = re.sub(
        r"\\(?:" + "|".join(PANDOC_SAFE_WRAPPER_COMMANDS) + r")\s*\{([^{}]*)\}",
        r"\1",
        expr,
    )
    expr = re.sub(
        r"\\(?:textnormal|textup|textsl|hbox|mbox)\s*\{([^{}]*)\}",
        r"\\text{\1}",
        expr,
    )
    expr = re.sub(r"\\operatornamewithlimits\s*\{([^{}]*)\}", r"\\operatorname{\1}", expr)
    expr = re.sub(r"\\textcolor\s*\{[^{}]*\}\s*\{([^{}]*)\}", r"\1", expr)
    expr = re.sub(r"\\(?:cancel|bcancel|xcancel|sout)\s*\{([^{}]*)\}", r"\1", expr)
    expr = re.sub(r"\\fbox\s*\{([^{}]*)\}", r"\\boxed{\1}", expr)
    expr = re.sub(r"\\raisebox\s*\{[^{}]*\}(?:\s*\[[^\]]*\]){0,2}\s*\{([^{}]*)\}", r"\1", expr)
    expr = re.sub(
        r"\\hdotsfor\s*\{(\d+)\}",
        lambda match: r"\cdots" * max(1, min(int(match.group(1)), 12)),
        expr,
    )
    expr = re.sub(r"\\(?:hspace|vspace)\s*\{[^{}]*\}", " ", expr)
    expr = re.sub(r"\\(?:phantom|hphantom|vphantom)\s*\{[^{}]*\}", "", expr)
    expr = re.sub(r"\\(?:limits|nolimits)(?![A-Za-z])", "", expr)
    expr = re.sub(r"\\(?:thinspace|medspace|thickspace|negthinspace|negmedspace|negthickspace)(?![A-Za-z])", " ", expr)
    expr = re.sub(r"\\notag(?![A-Za-z])|\\nonumber(?![A-Za-z])", "", expr)
    expr = re.sub(r"\\eqref\s*\{([^{}]*)\}", r"(\\text{\1})", expr)
    expr = re.sub(r"\\ref\s*\{([^{}]*)\}", r"\\text{\1}", expr)
    expr = re.sub(r"\\(?:hline|cline)\s*(?:\{[^{}]*\})?", " ", expr)
    expr = re.sub(r"\\(?:multicolumn|multirow)\s*\{[^{}]*\}\s*\{[^{}]*\}\s*\{([^{}]*)\}", r"\1", expr)
    return expr


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
    expr = re.sub(r"\\(?:dotsc|dotsi|dotsb|dotsm|dotso|dots)(?![A-Za-z])", r"\\ldots", expr)
    expr = re.sub(r"\s*([_^])\s*", r"\1", expr)
    expr = re.sub(r"\{\s*([^{}\n]+?)\s*\}", r"{\1}", expr)
    for index, text_group in enumerate(protected_text_groups):
        expr = expr.replace(f"@@TEXTGROUP{index}@@", text_group)
    expr = re.sub(r"\\(?:,|;|:|!)", " ", expr)
    expr = re.sub(r"[ \t]+", " ", expr)
    expr = re.sub(r"\n{3,}", "\n\n", expr)
    expr = _rewrite_unsupported_presentation_helpers(expr)
    expr = _rewrite_unknown_latex_commands(expr, math_mode=True)
    expr = BUILDREL_RE.sub(_rewrite_buildrel, expr)
    expr = _rewrite_legacy_infix_math(expr.strip())
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
    if not expr:
        return False
    unknown_commands = _unknown_latex_command_matches(expr)
    if not LATEX_NAMED_COMMAND_RE.search(expr) and not unknown_commands:
        return False
    command_count = len(LATEX_NAMED_COMMAND_RE.findall(expr))
    operator_count = len(re.findall(r"[+\-*/=<>]", expr))
    if command_count >= 1:
        return True
    if unknown_commands and (
        _unknown_latex_command_has_group(expr)
        or operator_count >= 1
        or len(re.findall(r"[_^]\s*(?:\{|[A-Za-z0-9])", expr)) >= 1
    ):
        return True
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
        match = UNKNOWN_LATEX_COMMAND_RE.search(text, pos)
        if not match:
            result.append(text[pos:])
            break
        if text.startswith("\\begin", match.start()):
            env_match = re.match(r"\\begin\s*\{([^{}]+)\}", text[match.start() :])
            if env_match:
                env = env_match.group(1)
                end_marker = f"\\end{{{env}}}"
                end = text.find(end_marker, match.start() + env_match.end())
                if end >= 0:
                    expr_start = match.start()
                    expr_end = end + len(end_marker)
                    result.append(text[pos:expr_start])
                    result.append(f"\n\n$$\n{normalize_latex_math(text[expr_start:expr_end])}\n$$\n\n")
                    pos = expr_end
                    continue
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
        if matched_command not in LATEX_COMMAND_NAMES and _is_likely_path_escape(text, start):
            result.append(text[pos:match.end()])
            pos = match.end()
            continue
        if matched_command not in LATEX_COMMAND_NAMES:
            group = _read_balanced_brace_group(text, match.end())
            next_char = text[match.end()] if match.end() < len(text) else ""
            if group is None and len(matched_command) == 1 and (not next_char or next_char.isspace() or TEXT_MATH_BOUNDARY_RE.match(next_char)):
                result.append(text[pos:match.end()])
                pos = match.end()
                continue
            rewrite_end = group[1] if group else match.end()
            result.append(text[pos:match.start()])
            result.append(_rewrite_unknown_latex_commands(text[match.start():rewrite_end], math_mode=False))
            pos = rewrite_end
            continue
        if matched_command in {
            "over",
            "choose",
            "above",
            "atop",
            "brack",
            "brace",
            "overwithdelims",
            "atopwithdelims",
            "abovewithdelims",
        }:
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
    has_inline_math_delimiter = bool(INLINE_DOLLAR_RE.search(text) or INLINE_PAREN_RE.search(text))
    if not has_inline_math_delimiter and _looks_like_display_math(text):
        result.extend(["$$", normalize_latex_math(text), "$$", ""])
    else:
        normalized = _normalize_inline_math(text)
        result.extend(normalized.splitlines())
        result.append("")
    paragraph.clear()


def normalize_math_markdown(markdown: str) -> str:
    markdown = FRONT_MATTER_RE.sub("", html.unescape(markdown or ""))
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
