from __future__ import annotations

import argparse
import copy
import json
import os
import random
import re
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree
from markdown_math import LATEX_RESIDUAL_RE


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_FONT_PATH = ROOT_DIR / "ttf_files" / "神韵英子楷书.ttf"
DEFAULT_FONT_NAME = "神韵英子楷书"
FONT_ALIASES = {
    "神韵英子楷书": {"神韵英子楷书", "SYyingzikaishu"},
    "SYyingzikaishu": {"神韵英子楷书", "SYyingzikaishu"},
    "李国夫手写体": {"李国夫手写体", "liguofu"},
    "liguofu": {"李国夫手写体", "liguofu"},
    "青叶手写体": {"青叶手写体", "QYSXT"},
    "QYSXT": {"青叶手写体", "QYSXT"},
    "云烟体": {"云烟体", "YYT"},
    "YYT": {"云烟体", "YYT"},
}
BODY_INK = "151515"
HEADING_INK = "0B4F71"
MIN_TEXT_SIZE = 22

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

IMAGE_MD_RE = re.compile(r"!\[[^\]]*\]\([^)]+\)")
HTML_IMAGE_RE = re.compile(r"<img\b[^>]*>", re.I)
FRONT_MATTER_RE = re.compile(r"\A---\s*\n.*?\n---\s*\n", re.S)
GREEK_REPLACEMENTS = {
    "alpha": "α",
    "beta": "β",
    "gamma": "γ",
    "delta": "δ",
    "mu": "μ",
    "sigma": "σ",
    "theta": "θ",
    "lambda": "λ",
    "pi": "π",
    "chi": "χ",
    "varLambda": "Λ",
    "Lambda": "Λ",
}


def w(tag: str) -> str:
    return f"{{{NS['w']}}}{tag}"


def m(tag: str) -> str:
    return f"{{{NS['m']}}}{tag}"


def ensure_child(parent: etree._Element, tag: str) -> etree._Element:
    child = parent.find(tag)
    if child is None:
        child = etree.SubElement(parent, tag)
    return child


def set_w_attr(element: etree._Element, name: str, value: str) -> None:
    element.set(w(name), value)


def font_aliases(font_name: str) -> set[str]:
    return {alias.lower() for alias in FONT_ALIASES.get(font_name, {font_name})}


def normalize_markdown(markdown: str) -> tuple[str, dict[str, Any]]:
    """Prepare Markdown for no-image handwritten DOCX generation.

    The final document must not contain copied/cropped images. Image references
    are therefore removed and counted for the conversion report.
    """

    image_refs = IMAGE_MD_RE.findall(markdown) + HTML_IMAGE_RE.findall(markdown)
    markdown = FRONT_MATTER_RE.sub("", markdown)
    markdown = HTML_IMAGE_RE.sub("", markdown)
    markdown = IMAGE_MD_RE.sub("", markdown)
    markdown = markdown.replace("\r\n", "\n").replace("\r", "\n")

    # Pandoc accepts \(..\), but normalizing to dollar math makes later residual
    # checks and manual debugging simpler.
    markdown = re.sub(r"\\\[(.*?)\\\]", lambda m: "\n\n$$\n" + m.group(1).strip() + "\n$$\n\n", markdown, flags=re.S)
    markdown = re.sub(r"\\\((.*?)\\\)", lambda m: "$" + m.group(1).strip() + "$", markdown, flags=re.S)
    markdown = re.sub(r"\n{3,}", "\n\n", markdown).strip() + "\n"
    return markdown, {"removed_images": len(image_refs), "image_refs": image_refs}


def _replace_simple_frac(text: str) -> str:
    frac_re = re.compile(r"\\(?:cfrac|dfrac|tfrac|frac)\s*\{([^{}]+)\}\s*\{([^{}]+)\}")
    previous = None
    while previous != text:
        previous = text
        text = frac_re.sub(r"(\1)/(\2)", text)
    return text


def _replace_latex_command(text: str, command: str, replacement: str) -> str:
    return re.sub(rf"\\{command}(?![A-Za-z])", replacement, text)


def plainify_latex_text(text: str) -> str:
    """Convert unparsed TeX fallback into readable non-LaTeX text.

    This is a last-resort cleanup for OCR-damaged math that Pandoc cannot turn
    into Office Math. It intentionally avoids images and removes visible TeX
    commands from the final document.
    """

    if not LATEX_RESIDUAL_RE.search(text):
        return text
    text = text.replace("$", "")
    text = text.replace("\\cfrac", "\\frac").replace("\\dfrac", "\\frac").replace("\\tfrac", "\\frac")
    text = _replace_simple_frac(text)
    text = re.sub(r"\\sqrt\s*\[([^\[\]]+)\]\s*\{([^{}]+)\}", r"√[\1](\2)", text)
    text = re.sub(r"\\sqrt\s*\{([^{}]+)\}", r"√(\1)", text)
    text = re.sub(
        r"\\xrightarrow(?:\[([^\[\]]*)\])?\s*\{([^{}]*)\}",
        lambda match: "→"
        + (f"_{{{match.group(1)}}}" if match.group(1) else "")
        + (f"^{{{match.group(2)}}}" if match.group(2) else ""),
        text,
    )
    text = re.sub(
        r"\\xleftarrow(?:\[([^\[\]]*)\])?\s*\{([^{}]*)\}",
        lambda match: "←"
        + (f"_{{{match.group(1)}}}" if match.group(1) else "")
        + (f"^{{{match.group(2)}}}" if match.group(2) else ""),
        text,
    )
    text = re.sub(r"\\tag\*?\s*\{([^{}]*)\}", r"(\1)", text)
    text = re.sub(r"\\eqref\s*\{([^{}]*)\}", r"(\1)", text)
    text = re.sub(r"\\ref\s*\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\(?:label|notag|nonumber)\s*(?:\{[^{}]*\})?", "", text)
    text = re.sub(r"\\substack\s*\{([^{}]*)\}", lambda match: match.group(1).replace("\\\\", " "), text)
    text = re.sub(r"\\begin\s*\{array\}\s*\{[^{}]*\}", " ", text)
    text = re.sub(r"\\begin\s*\{subarray\}\s*\{[^{}]*\}", " ", text)
    text = re.sub(r"\\(?:hline|cline)\s*(?:\{[^{}]*\})?", " ", text)
    text = re.sub(r"\\(?:multicolumn|multirow)\s*\{[^{}]*\}\s*\{[^{}]*\}\s*\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\(?:color|textcolor)\s*\{[^{}]*\}\s*\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\color\s*\{[^{}]*\}", "", text)
    text = re.sub(r"\\(?:phantom|hphantom|vphantom)\s*\{[^{}]*\}", "", text)
    text = re.sub(r"\\(?:mbox|hbox)\s*\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\(?:boxed|fbox)\s*\{([^{}]*)\}", r"[\1]", text)
    text = re.sub(
        r"\\(?:cancel|bcancel|xcancel|sout|overparen|underparen|overleftarrow|underleftarrow|underrightarrow)\s*\{([^{}]*)\}",
        r"\1",
        text,
    )
    text = re.sub(
        r"\\(?:boldsymbol|boldmath|mathbf|mathrm|mathbb|mathscr|mathds|mathcal|mathfrak|mathsf|mathtt|mathit|textbf|bm|operatorname\*?|mathop\*?|text)\s*\{([^{}]*)\}",
        r"\1",
        text,
    )
    text = re.sub(r"\\(?:textrm|cal|small)\s*\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\bf\s*([A-Za-z0-9]+)", r"\1", text)
    text = re.sub(r"\\overline\s*\{\{?([^{}]+)\}?\}", r"\1̄", text)
    text = re.sub(r"\\(?:acute|')\s*\{([^{}]+)\}", r"´\1", text)
    text = re.sub(r"\\(?:grave|`)\s*\{([^{}]+)\}", r"`\1", text)
    text = re.sub(r"\\breve\s*\{([^{}]+)\}", r"˘\1", text)
    text = re.sub(r"\\check\s*\{([^{}]+)\}", r"ˇ\1", text)
    text = re.sub(r"\\mathring\s*\{([^{}]+)\}", r"˚\1", text)
    text = re.sub(r"\\stackrel\s*\{[^{}]*\}\s*\{([^{}]+)\}", r"\1", text)
    text = re.sub(r"\\begin\s*\{[^{}]+\}|\\end\s*\{[^{}]+\}", " ", text)
    text = text.replace("&", " ")
    text = text.replace("\\\\", " ")
    text = re.sub(r"\\(?:left|right|middle|big|Big|bigl|bigr|Bigl|Bigr|bigg|biggl|biggr|Bigg|Biggl|Biggr)", "", text)
    text = _replace_latex_command(text, "lbrace", "@@LBRACE@@")
    text = _replace_latex_command(text, "rbrace", "@@RBRACE@@")
    text = re.sub(r"\\not\s*\\in(?![A-Za-z])", "∉", text)
    text = re.sub(r"\\pod\s*\{([^{}]*)\}", r"(\1)", text)
    text = re.sub(r"\\pmod\s*\{([^{}]*)\}", r"mod \1", text)
    symbol_replacements = {
        "leqslant": "≤",
        "leq": "≤",
        "geqslant": "≥",
        "geq": "≥",
        "neq": "≠",
        "ne": "≠",
        "approx": "≈",
        "sim": "∼",
        "equiv": "≡",
        "notin": "∉",
        "in": "∈",
        "lVert": "‖",
        "rVert": "‖",
        "Vert": "‖",
        "perp": "⊥",
        "parallel": "∥",
        "angle": "∠",
        "therefore": "∴",
        "because": "∵",
        "colon": ":",
        "rightarrow": "→",
        "leftarrow": "←",
        "Longrightarrow": "⟹",
        "Longleftarrow": "⟸",
        "Longleftrightarrow": "⟺",
        "implies": "⇒",
        "hookrightarrow": "↪",
        "hookleftarrow": "↩",
        "twoheadrightarrow": "↠",
        "twoheadleftarrow": "↞",
        "rightsquigarrow": "↝",
        "to": "→",
        "lparen": "(",
        "rparen": ")",
        "lbrack": "[",
        "rbrack": "]",
        "langle": "〈",
        "rangle": "〉",
        "Re": "ℜ",
        "Im": "ℑ",
        "ell": "ℓ",
        "hbar": "ℏ",
        "aleph": "ℵ",
        "wp": "℘",
        "neg": "¬",
        "lnot": "¬",
        "prec": "≺",
        "preceq": "≼",
        "succ": "≻",
        "succeq": "≽",
        "ll": "≪",
        "gg": "≫",
        "asymp": "≍",
        "doteq": "≐",
        "smallsetminus": "∖",
        "sqcup": "⊔",
        "sqcap": "⊓",
        "bmod": "mod",
        "limsup": "lim sup",
        "liminf": "lim inf",
        "injlim": "inj lim",
        "projlim": "proj lim",
        "infty": "∞",
        "times": "×",
        "cdot": "·",
        "pm": "±",
        "cdots": "⋯",
        "ldots": "…",
        "dots": "…",
        "quad": " ",
        "qquad": " ",
        "ln": "ln",
        "exp": "exp",
    }
    for source, target in symbol_replacements.items():
        text = _replace_latex_command(text, source, target)
    for name, replacement in GREEK_REPLACEMENTS.items():
        text = _replace_latex_command(text, name, replacement)
    text = re.sub(r"‖\s+", "‖", text)
    text = re.sub(r"\s+‖", "‖", text)
    text = re.sub(
        r"\{?([A-Za-z0-9_^\-]+)\}?\s*\\over\s*\{?([A-Za-z0-9_^\-]+)\}?",
        r"(\1)/(\2)",
        text,
    )
    text = re.sub(r"\\[,;:! ]", " ", text)
    text = text.replace("\\{", "{").replace("\\}", "}").replace("\\_", "_")
    text = re.sub(r"\\([A-Za-z]+)", r"\1", text)
    text = text.replace("{", "").replace("}", "")
    text = text.replace("@@LBRACE@@", "{").replace("@@RBRACE@@", "}")
    text = text.replace("~", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def markdown_from_pdf(source: Path, out: Path) -> Path:
    """Extract the text layer from a PDF as Markdown, intentionally ignoring images."""

    import fitz

    parts = [f"# {source.stem}", ""]
    text_chars = 0
    with fitz.open(source) as doc:
        for page_index, page in enumerate(doc, start=1):
            if doc.page_count > 1:
                parts.extend([f"## 第{page_index}页", ""])
            blocks = page.get_text("blocks", sort=True)
            lines: list[str] = []
            for block in blocks:
                if len(block) >= 7 and int(block[6]) != 0:
                    continue
                block_text = str(block[4] if len(block) > 4 else "")
                for raw_line in block_text.splitlines():
                    line = re.sub(r"\s+", " ", raw_line).strip()
                    if line:
                        lines.append(line)
            page_text = "\n\n".join(lines).strip()
            if page_text:
                text_chars += len(page_text)
                parts.append(page_text)
                parts.append("")

    if text_chars < 20:
        raise RuntimeError(
            "PDF text layer is too sparse for no-image conversion. "
            "Use OCR/MinerU to extract Markdown first, then run the handwritten document pipeline."
        )
    out.write_text("\n".join(parts).strip() + "\n", encoding="utf-8")
    return out


def markdown_from_source(source: Path, work_dir: Path) -> Path:
    source = source.resolve()
    if source.suffix.lower() in {".md", ".markdown"}:
        return source
    out = work_dir / f"{source.stem}.md"
    if source.suffix.lower() == ".docx":
        media_dir = work_dir / f"{source.stem}_media"
        subprocess.run(
            [
                "pandoc",
                "--from=docx",
                "--to=markdown+tex_math_dollars",
                f"--extract-media={media_dir}",
                str(source),
                "-o",
                str(out),
            ],
            check=True,
        )
        return out
    if source.suffix.lower() == ".pdf":
        return markdown_from_pdf(source, out)
    raise ValueError(f"Unsupported input type for direct conversion: {source}")


def run_md_to_docx(markdown: Path, docx: Path) -> None:
    subprocess.run(["codex-md-to-docx", str(markdown), str(docx)], check=True)
    if not docx.exists():
        raise FileNotFoundError(f"Expected DOCX was not created: {docx}")


def install_font_for_local_user(font_path: Path) -> Path | None:
    """Install the handwriting font in the macOS user font folder if needed.

    LibreOffice PDF export can only use the TTF if the font is discoverable.
    The operation is intentionally local to the current user and does not need sudo.
    """

    if not font_path.exists():
        return None
    target_dir = Path.home() / "Library" / "Fonts"
    if not target_dir.exists():
        return None
    target = target_dir / font_path.name
    if target.exists() and target.stat().st_size == font_path.stat().st_size:
        return target
    shutil.copy2(font_path, target)
    return target


def rewrite_zip_dir(src_dir: Path, dst_docx: Path) -> None:
    with ZipFile(dst_docx, "w", ZIP_DEFLATED) as zout:
        for path in src_dir.rglob("*"):
            if path.is_file():
                zout.write(path, path.relative_to(src_dir).as_posix())


def remove_media_files(unpacked: Path) -> int:
    media_dir = unpacked / "word" / "media"
    if not media_dir.exists():
        return 0
    files = [p for p in media_dir.rglob("*") if p.is_file()]
    shutil.rmtree(media_dir)
    return len(files)


def set_run_font(
    run_pr: etree._Element,
    font_name: str,
    size_half_points: int,
    *,
    bold: bool = False,
    color: str | None = None,
) -> None:
    fonts = ensure_child(run_pr, w("rFonts"))
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        set_w_attr(fonts, key, font_name)
    size = ensure_child(run_pr, w("sz"))
    set_w_attr(size, "val", str(size_half_points))
    size_cs = ensure_child(run_pr, w("szCs"))
    set_w_attr(size_cs, "val", str(size_half_points))
    if bold:
        ensure_child(run_pr, w("b"))
        ensure_child(run_pr, w("bCs"))
    else:
        for child in run_pr.findall(w("b")) + run_pr.findall(w("bCs")):
            run_pr.remove(child)
    if color:
        color_node = ensure_child(run_pr, w("color"))
        set_w_attr(color_node, "val", color)


def jitter_color(color: str, rng: random.Random, amplitude: int) -> str:
    color = color.strip().lstrip("#")
    if len(color) != 6:
        return color
    channels = [int(color[i : i + 2], 16) for i in (0, 2, 4)]
    jittered = [max(0, min(255, channel + rng.randint(-amplitude, amplitude))) for channel in channels]
    return "".join(f"{channel:02X}" for channel in jittered)


def set_run_jitter(
    run_pr: etree._Element,
    *,
    size_delta: int,
    position_delta: int,
    spacing_delta: int = 0,
    scale_delta: int = 0,
    color: str | None = None,
) -> None:
    size = run_pr.find(w("sz"))
    if size is not None:
        base = int(size.get(w("val"), "30") or "30")
        set_w_attr(size, "val", str(max(MIN_TEXT_SIZE, base + size_delta)))
    size_cs = run_pr.find(w("szCs"))
    if size_cs is not None:
        base = int(size_cs.get(w("val"), "30") or "30")
        set_w_attr(size_cs, "val", str(max(MIN_TEXT_SIZE, base + size_delta)))
    if position_delta:
        position = ensure_child(run_pr, w("position"))
        set_w_attr(position, "val", str(position_delta))
    else:
        for child in run_pr.findall(w("position")):
            run_pr.remove(child)
    if spacing_delta:
        spacing = ensure_child(run_pr, w("spacing"))
        set_w_attr(spacing, "val", str(spacing_delta))
    else:
        for child in run_pr.findall(w("spacing")):
            run_pr.remove(child)
    if scale_delta:
        scale = ensure_child(run_pr, w("w"))
        set_w_attr(scale, "val", str(max(92, min(108, 100 + scale_delta))))
    else:
        for child in run_pr.findall(w("w")):
            run_pr.remove(child)
    if color:
        color_node = ensure_child(run_pr, w("color"))
        set_w_attr(color_node, "val", color)


def run_text(run: etree._Element) -> str:
    return "".join(t.text or "" for t in run.xpath(".//w:t", namespaces=NS))


def clear_run_text(run: etree._Element, text: str) -> None:
    for child in list(run):
        if etree.QName(child).localname not in {"rPr"}:
            run.remove(child)
    t = etree.SubElement(run, w("t"))
    if text.isspace() or text.startswith(" ") or text.endswith(" "):
        t.set(XML_SPACE, "preserve")
    t.text = text


def omath_text(omath: etree._Element) -> str:
    text = "".join(omath.xpath(".//m:t/text()", namespaces=NS))
    text = text.replace("\u2009", " ").replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def is_simple_math(omath: etree._Element) -> bool:
    text = omath_text(omath)
    if not text or len(text) > 90:
        return False
    complex_nodes = omath.xpath(
        ".//m:f | .//m:rad | .//m:nary | .//m:m | .//m:eqArr | .//m:groupChr | .//m:limLow | .//m:limUpp",
        namespaces=NS,
    )
    if complex_nodes:
        return False
    return True


def make_text_run(
    text: str,
    font_name: str,
    size_half_points: int,
    color: str,
    *,
    vert_align: str | None = None,
) -> etree._Element:
    run = etree.Element(w("r"))
    run_pr = etree.SubElement(run, w("rPr"))
    set_run_font(run_pr, font_name, size_half_points, bold=False, color=color)
    if vert_align:
        align = ensure_child(run_pr, w("vertAlign"))
        set_w_attr(align, "val", vert_align)
    clear_run_text(run, text)
    return run


MathToken = tuple[str, str | None]


def _direct_children(element: etree._Element, local_name: str) -> list[etree._Element]:
    return [child for child in element if etree.QName(child).localname == local_name]


def math_tokens(element: etree._Element, vert_align: str | None = None) -> list[MathToken]:
    """Flatten simple Office Math into text tokens while preserving sub/superscript."""

    local = etree.QName(element).localname
    if local == "r":
        text = "".join(element.xpath("./m:t/text()", namespaces=NS))
        return [(text, vert_align)] if text else []
    if local == "sSup":
        tokens: list[MathToken] = []
        for child in _direct_children(element, "e"):
            tokens.extend(math_tokens(child, vert_align))
        for child in _direct_children(element, "sup"):
            tokens.extend(math_tokens(child, "superscript"))
        return tokens
    if local == "sSub":
        tokens = []
        for child in _direct_children(element, "e"):
            tokens.extend(math_tokens(child, vert_align))
        for child in _direct_children(element, "sub"):
            tokens.extend(math_tokens(child, "subscript"))
        return tokens
    if local == "sSubSup":
        tokens = []
        for child in _direct_children(element, "e"):
            tokens.extend(math_tokens(child, vert_align))
        for child in _direct_children(element, "sub"):
            tokens.extend(math_tokens(child, "subscript"))
        for child in _direct_children(element, "sup"):
            tokens.extend(math_tokens(child, "superscript"))
        return tokens

    tokens = []
    for child in element:
        tokens.extend(math_tokens(child, vert_align))
    return tokens


def text_runs_from_math(omath: etree._Element, font_name: str) -> list[etree._Element]:
    tokens = math_tokens(omath)
    if not tokens:
        text = omath_text(omath)
        return [make_text_run(text, font_name, 30, BODY_INK)] if text else []

    merged: list[MathToken] = []
    for text, vert_align in tokens:
        if not text:
            continue
        text = text.replace("\u2009", " ").replace("\u00a0", " ")
        if merged and merged[-1][1] == vert_align:
            merged[-1] = (merged[-1][0] + text, vert_align)
        else:
            merged.append((text, vert_align))
    return [make_text_run(text, font_name, 30, BODY_INK, vert_align=vert_align) for text, vert_align in merged if text]


def demote_simple_math(root: etree._Element, font_name: str) -> dict[str, int]:
    demoted_inline = 0
    demoted_display = 0

    for omath_para in list(root.xpath("//m:oMathPara", namespaces=NS)):
        omaths = omath_para.xpath("./m:oMath", namespaces=NS)
        if len(omaths) != 1 or not is_simple_math(omaths[0]):
            continue
        parent = omath_para.getparent()
        if parent is None or etree.QName(parent).localname != "p":
            continue
        runs = text_runs_from_math(omaths[0], font_name)
        if not runs:
            continue
        index = parent.index(omath_para)
        parent.remove(omath_para)
        for offset, run in enumerate(runs):
            parent.insert(index + offset, run)
        demoted_display += 1

    for omath in list(root.xpath("//m:oMath[not(ancestor::m:oMathPara)]", namespaces=NS)):
        if not is_simple_math(omath):
            continue
        parent = omath.getparent()
        if parent is None or etree.QName(parent).localname != "p":
            continue
        runs = text_runs_from_math(omath, font_name)
        if not runs:
            continue
        index = parent.index(omath)
        parent.remove(omath)
        for offset, run in enumerate(runs):
            parent.insert(index + offset, run)
        demoted_inline += 1

    return {"inline": demoted_inline, "display": demoted_display, "total": demoted_inline + demoted_display}


def jitter_profile_for_char(ch: str, rng: random.Random) -> tuple[int, int, int, int]:
    if ch.isspace():
        return 0, 0, 0, 0
    if ch.isascii() and (ch.isalpha() or ch.isdigit()):
        return (
            rng.choice([-2, -1, 0, 0, 1, 2]),
            rng.choice([-3, -2, -1, 0, 0, 1, 2, 3]),
            rng.choice([-6, -3, 0, 0, 2, 4, 7, 10]),
            rng.choice([-5, -3, 0, 0, 2, 4, 6]),
        )
    if re.match(r"[\u4e00-\u9fff]", ch):
        return (
            rng.choice([-2, -1, 0, 0, 1, 2]),
            rng.choice([-2, -1, 0, 0, 1, 2]),
            rng.choice([-4, -2, 0, 0, 2, 5, 8]),
            rng.choice([-4, -2, 0, 0, 2, 4]),
        )
    return (
        rng.choice([-1, 0, 0, 1]),
        rng.choice([-2, -1, 0, 0, 1, 2]),
        rng.choice([-4, -2, 0, 0, 2, 5]),
        rng.choice([-3, 0, 0, 3]),
    )


def split_run_with_jitter(
    run: etree._Element,
    rng: random.Random,
    base_size: int,
    font_name: str,
    ink_color: str,
) -> list[etree._Element] | None:
    original_text = run_text(run)
    text = plainify_latex_text(original_text)
    if len(text) <= 1 or len(text) > 240:
        if text != original_text:
            clear_run_text(run, text)
        run_pr = ensure_child(run, w("rPr"))
        set_run_font(run_pr, font_name, base_size, bold=False, color=ink_color)
        set_run_jitter(
            run_pr,
            size_delta=rng.choice([-2, -1, 0, 0, 1, 2]),
            position_delta=rng.choice([-2, -1, 0, 0, 1, 2]),
            spacing_delta=rng.choice([-3, 0, 0, 3, 6]),
            scale_delta=rng.choice([-4, -2, 0, 0, 2, 4]),
            color=jitter_color(ink_color, rng, 14),
        )
        return None

    new_runs: list[etree._Element] = []
    for ch in text:
        new_run = copy.deepcopy(run)
        run_pr = ensure_child(new_run, w("rPr"))
        set_run_font(run_pr, font_name, base_size, bold=False, color=ink_color)
        size_delta, position_delta, spacing_delta, scale_delta = jitter_profile_for_char(ch, rng)
        set_run_jitter(
            run_pr,
            size_delta=size_delta,
            position_delta=position_delta,
            spacing_delta=spacing_delta,
            scale_delta=scale_delta,
            color=jitter_color(ink_color, rng, 12) if not ch.isspace() else None,
        )
        clear_run_text(new_run, ch)
        new_runs.append(new_run)
    return new_runs


def paragraph_text(para: etree._Element) -> str:
    return "".join(para.xpath(".//w:t/text()", namespaces=NS))


def paragraph_style(para: etree._Element) -> str:
    styles = para.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
    return styles[0] if styles else ""


def paragraph_has_math(para: etree._Element) -> bool:
    return bool(para.xpath(".//m:oMath | .//m:oMathPara", namespaces=NS))


def looks_like_math_line(text: str) -> bool:
    stripped = text.strip()
    if not stripped or len(stripped) > 120:
        return False
    if re.search(r"[\u4e00-\u9fff]", stripped):
        return False
    return bool(re.search(r"[A-Za-z0-9α-ωΑ-Ω]", stripped) and re.search(r"[=+\-*/^≤≥<>≈]", stripped))


def paragraph_kind(para: etree._Element) -> str:
    style = paragraph_style(para)
    text = paragraph_text(para).strip()
    numbered_heading = bool(re.match(r"^第?[一二三四五六七八九十\d]+[、.．)]\s*$", text)) or bool(
        re.match(r"^第?[一二三四五六七八九十\d]+题$", text)
    )
    if style.startswith("Heading") or numbered_heading:
        return "heading"
    if paragraph_has_math(para) and not text:
        return "formula"
    if looks_like_math_line(text):
        return "hand_formula"
    return "body"


def apply_paragraph_style(para: etree._Element, rng: random.Random) -> tuple[int, str]:
    kind = paragraph_kind(para)
    if kind == "heading":
        base_size = 34
        before_base = rng.choice([0, 20, 40])
        after_base = rng.choice([80, 100, 120])
        line = rng.choice([400, 415, 430])
        first_line = 0
        left = rng.choice([0, 30, 60, 90])
        align = "left"
        ink_color = HEADING_INK
    elif kind == "formula":
        base_size = 30
        before_base = rng.choice([10, 20, 30])
        after_base = rng.choice([40, 55, 70])
        line = rng.choice([360, 375, 390])
        first_line = 0
        left = 0
        align = "center"
        ink_color = BODY_INK
    elif kind == "hand_formula":
        base_size = rng.choice([29, 30, 31])
        before_base = rng.choice([15, 25, 40, 55])
        after_base = rng.choice([45, 60, 80, 100])
        line = rng.choice([360, 380, 400, 420])
        first_line = 0
        align = rng.choice(["left", "left", "center"])
        left = 0 if align == "center" else rng.choice([260, 360, 480, 620, 760])
        ink_color = BODY_INK
    else:
        base_size = 30
        before_base = rng.choice([0, 0, 10, 25, 40])
        after_base = rng.choice([55, 70, 90, 115, 135])
        line = rng.choice([395, 415, 435, 455, 475, 495])
        first_line = rng.choice([0, 100, 160, 220, 300])
        left = rng.choice([0, 0, 45, 90, 140, 190])
        align = "left"
        ink_color = BODY_INK

    para_pr = ensure_child(para, w("pPr"))
    spacing = ensure_child(para_pr, w("spacing"))
    set_w_attr(spacing, "before", str(max(0, before_base)))
    set_w_attr(spacing, "after", str(max(0, after_base)))
    set_w_attr(spacing, "line", str(line))
    set_w_attr(spacing, "lineRule", "auto")
    ind = ensure_child(para_pr, w("ind"))
    set_w_attr(ind, "left", str(left))
    set_w_attr(ind, "firstLine", str(first_line))
    jc = ensure_child(para_pr, w("jc"))
    set_w_attr(jc, "val", align)
    return base_size, ink_color


def strip_drawings_and_relationships(unpacked: Path, root: etree._Element) -> int:
    drawings = root.xpath("//w:drawing | //w:pict | //w:object", namespaces=NS)
    removed = len(drawings)
    for drawing in drawings:
        parent = drawing.getparent()
        if parent is not None:
            parent.remove(drawing)

    rels_path = unpacked / "word" / "_rels" / "document.xml.rels"
    if rels_path.exists():
        rel_root = etree.fromstring(rels_path.read_bytes())
        rel_ns = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
        for rel in list(rel_root.xpath("//rel:Relationship", namespaces=rel_ns)):
            target = rel.get("Target", "")
            rel_type = rel.get("Type", "")
            if "image" in rel_type or target.startswith("media/"):
                rel_root.remove(rel)
        rels_path.write_bytes(etree.tostring(rel_root, xml_declaration=True, encoding="UTF-8", standalone=True))
    return removed


def apply_page_layout(root: etree._Element, rng: random.Random) -> None:
    for sect_pr in root.xpath("//w:sectPr", namespaces=NS):
        pg_sz = ensure_child(sect_pr, w("pgSz"))
        set_w_attr(pg_sz, "w", "11906")
        set_w_attr(pg_sz, "h", "16838")
        pg_mar = ensure_child(sect_pr, w("pgMar"))
        margins = {
            "top": 960 + rng.choice([0, 40, 80]),
            "right": 1080 + rng.choice([-40, 0, 60]),
            "bottom": 900 + rng.choice([0, 40, 80]),
            "left": 1180 + rng.choice([0, 80, 120]),
            "header": 300,
            "footer": 300,
            "gutter": 0,
        }
        for key, value in margins.items():
            set_w_attr(pg_mar, key, str(value))


def apply_styles(styles_xml: Path, font_name: str) -> None:
    if not styles_xml.exists():
        return
    root = etree.fromstring(styles_xml.read_bytes())
    for style in root.xpath("//w:style", namespaces=NS):
        style_id = style.get(w("styleId"), "")
        run_pr = ensure_child(style, w("rPr"))
        if style_id == "Heading1":
            set_run_font(run_pr, font_name, 36, bold=False, color=HEADING_INK)
        elif style_id.startswith("Heading"):
            set_run_font(run_pr, font_name, 34, bold=False, color=HEADING_INK)
        else:
            set_run_font(run_pr, font_name, 30, bold=False, color=BODY_INK)
    styles_xml.write_bytes(etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True))


def postprocess_docx(
    docx_path: Path,
    *,
    output_docx: Path | None = None,
    font_name: str = DEFAULT_FONT_NAME,
    seed: int = 14790897,
) -> dict[str, Any]:
    output_docx = output_docx or docx_path
    with tempfile.TemporaryDirectory(prefix="handwriting_docx_") as tmp:
        tmp_dir = Path(tmp)
        with ZipFile(docx_path) as zin:
            zin.extractall(tmp_dir)
        document_path = tmp_dir / "word" / "document.xml"
        root = etree.fromstring(document_path.read_bytes())

        rng = random.Random(seed)
        removed_drawings = strip_drawings_and_relationships(tmp_dir, root)
        removed_media = remove_media_files(tmp_dir)
        apply_page_layout(root, rng)
        apply_styles(tmp_dir / "word" / "styles.xml", font_name)
        demoted_math = demote_simple_math(root, font_name)

        changed_runs = 0
        split_runs = 0
        for para in root.xpath("//w:p[not(ancestor::m:oMath) and not(ancestor::m:oMathPara)]", namespaces=NS):
            base_size, ink_color = apply_paragraph_style(para, rng)
            runs = list(para.xpath("./w:r[not(.//w:drawing) and not(ancestor::m:oMath) and not(ancestor::m:oMathPara)]", namespaces=NS))
            for run in runs:
                if not run_text(run):
                    continue
                new_runs = split_run_with_jitter(run, rng, base_size, font_name, ink_color)
                changed_runs += 1
                if new_runs is not None:
                    parent = run.getparent()
                    if parent is None:
                        continue
                    index = parent.index(run)
                    parent.remove(run)
                    for offset, new_run in enumerate(new_runs):
                        parent.insert(index + offset, new_run)
                    split_runs += 1

        document_path.write_bytes(etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True))
        rewrite_zip_dir(tmp_dir, output_docx)
    return {
        "docx": str(output_docx),
        "font_name": font_name,
        "changed_runs": changed_runs,
        "split_runs": split_runs,
        "demoted_inline_math": demoted_math["inline"],
        "demoted_display_math": demoted_math["display"],
        "demoted_simple_math": demoted_math["total"],
        "removed_drawings": removed_drawings,
        "removed_media": removed_media,
    }


def convert_docx_to_pdf(docx: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    expected_pdf = out_dir / f"{docx.stem}.pdf"
    if expected_pdf.exists():
        expected_pdf.unlink()
    pdf = expected_pdf
    for attempt in range(2):
        subprocess.run(["codex-docx-to-pdf", str(docx), str(out_dir)], check=True)
        for _ in range(80):
            if pdf.exists():
                break
            time.sleep(0.25)
        if pdf.exists():
            break
        if attempt == 0:
            time.sleep(1.0)
    if not pdf.exists():
        candidates = sorted(out_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not candidates:
            raise FileNotFoundError(f"PDF was not created for {docx}")
        pdf = candidates[0]
    return pdf


def run_font_values(run: etree._Element) -> list[str]:
    fonts = run.xpath("./w:rPr/w:rFonts", namespaces=NS)
    if not fonts:
        return []
    values: list[str] = []
    for font in fonts:
        for key in ("ascii", "hAnsi", "eastAsia", "cs"):
            value = font.get(w(key))
            if value:
                values.append(value)
    return values


def run_prop_values(run: etree._Element, prop: str) -> list[str]:
    values: list[str] = []
    for element in run.xpath(f"./w:rPr/w:{prop}", namespaces=NS):
        value = element.get(w("val"))
        if value:
            values.append(value)
    return values


def inspect_docx(docx: Path, *, font_name: str = DEFAULT_FONT_NAME) -> dict[str, Any]:
    with ZipFile(docx) as zf:
        names = zf.namelist()
        document_xml = zf.read("word/document.xml")
    root = etree.fromstring(document_xml)
    visible_text = "".join(root.xpath("//w:t/text()", namespaces=NS))
    non_math_text = "".join(root.xpath("//w:t[not(ancestor::m:oMath) and not(ancestor::m:oMathPara)]/text()", namespaces=NS))
    text_runs = [
        run
        for run in root.xpath("//w:r[.//w:t and not(ancestor::m:oMath) and not(ancestor::m:oMathPara)]", namespaces=NS)
        if run_text(run).strip()
    ]
    non_handwritten_runs: list[str] = []
    accepted_fonts = font_aliases(font_name)
    for run in text_runs:
        values = run_font_values(run)
        if not values or any(value.lower() not in accepted_fonts for value in values):
            non_handwritten_runs.append(run_text(run)[:40])
    positions = [value for run in text_runs for value in run_prop_values(run, "position")]
    spacings = [value for run in text_runs for value in run_prop_values(run, "spacing")]
    scales = [value for run in text_runs for value in run_prop_values(run, "w")]
    sizes = [value for run in text_runs for value in run_prop_values(run, "sz")]
    colors = [value for run in text_runs for value in run_prop_values(run, "color")]
    return {
        "docx": str(docx),
        "visible_chars": len(visible_text),
        "math_objects": len(root.xpath("//m:oMath | //m:oMathPara", namespaces=NS)),
        "media_files": [name for name in names if name.startswith("word/media/")],
        "drawing_objects": len(root.xpath("//w:drawing | //w:pict | //w:object", namespaces=NS)),
        "latex_residuals": sorted(set(LATEX_RESIDUAL_RE.findall(non_math_text))),
        "non_math_text_runs": len(text_runs),
        "handwritten_text_runs": len(text_runs) - len(non_handwritten_runs),
        "non_handwritten_text_runs": len(non_handwritten_runs),
        "non_handwritten_samples": non_handwritten_runs[:12],
        "positioned_text_runs": len(positions),
        "spaced_text_runs": len(spacings),
        "scaled_text_runs": len(scales),
        "distinct_text_sizes": sorted(set(sizes), key=lambda value: int(value))[:24],
        "distinct_text_colors": sorted(set(colors))[:24],
    }


def inspect_pdf(pdf: Path) -> dict[str, Any]:
    import fitz

    doc = fitz.open(pdf)
    try:
        image_count = sum(len(page.get_images(full=True)) for page in doc)
        text = "\n".join(page.get_text() for page in doc)
        fonts: set[str] = set()
        for page in doc:
            for item in page.get_fonts(full=True):
                if len(item) >= 4:
                    fonts.add(str(item[3]))
        return {
            "pdf": str(pdf),
            "pages": doc.page_count,
            "text_chars": len(text),
            "image_count": image_count,
            "fonts": sorted(fonts),
            "latex_residuals": sorted(set(LATEX_RESIDUAL_RE.findall(text))),
        }
    finally:
        doc.close()


def convert_to_handwritten(
    source: Path,
    out_dir: Path,
    *,
    output_stem: str | None = None,
    font_path: Path = DEFAULT_FONT_PATH,
    font_name: str = DEFAULT_FONT_NAME,
    seed: int = 14790897,
    keep_markdown: bool = True,
) -> dict[str, Any]:
    source = source.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    install_font_for_local_user(font_path)

    with tempfile.TemporaryDirectory(prefix="handwriting_md_") as tmp:
        tmp_dir = Path(tmp)
        md_source = markdown_from_source(source, tmp_dir)
        raw_markdown = md_source.read_text(encoding="utf-8")
        normalized, normalize_info = normalize_markdown(raw_markdown)

        stem = output_stem or f"{source.stem}_手写体"
        normalized_md = out_dir / f"{stem}.md"
        normalized_md.write_text(normalized, encoding="utf-8")

        docx = out_dir / f"{stem}.docx"
        run_md_to_docx(normalized_md, docx)
        post_info = postprocess_docx(docx, font_name=font_name, seed=seed)
        pdf = convert_docx_to_pdf(docx, out_dir)

    docx_info = inspect_docx(docx, font_name=font_name)
    pdf_info = inspect_pdf(pdf)
    accepted_fonts = font_aliases(font_name)
    handwritten_font_found = any(any(alias in font.lower() for alias in accepted_fonts) for font in pdf_info["fonts"])
    report = {
        "source": str(source),
        "markdown": str(normalized_md) if keep_markdown else None,
        "docx": str(docx),
        "pdf": str(pdf),
        "normalize": normalize_info,
        "postprocess": post_info,
        "docx_inspect": docx_info,
        "pdf_inspect": pdf_info,
        "handwritten_font_found": handwritten_font_found,
        "handwritten_font_aliases": sorted(accepted_fonts),
        "passed": (
            not docx_info["media_files"]
            and docx_info["drawing_objects"] == 0
            and not docx_info["latex_residuals"]
            and docx_info["non_handwritten_text_runs"] == 0
            and pdf_info["image_count"] == 0
            and not pdf_info["latex_residuals"]
            and pdf_info["text_chars"] > 0
            and handwritten_font_found
        ),
    }
    report_path = out_dir / f"{Path(report['pdf']).stem}_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    report["report"] = str(report_path)
    if not report["passed"]:
        raise RuntimeError(f"Handwritten conversion did not pass checks: {report_path}")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert Markdown/DOCX into no-image handwritten DOCX and PDF.")
    parser.add_argument("source", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--output-stem")
    parser.add_argument("--font-path", type=Path, default=DEFAULT_FONT_PATH)
    parser.add_argument("--font-name", default=DEFAULT_FONT_NAME)
    parser.add_argument("--seed", type=int, default=14790897)
    args = parser.parse_args()

    report = convert_to_handwritten(
        args.source,
        args.out_dir,
        output_stem=args.output_stem,
        font_path=args.font_path,
        font_name=args.font_name,
        seed=args.seed,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
