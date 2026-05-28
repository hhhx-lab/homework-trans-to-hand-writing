from __future__ import annotations

from pathlib import Path
from typing import Any

import pypandoc
from werkzeug.utils import secure_filename

from markdown_math import normalize_math_markdown
from mineru_adapter import MinerUConfigError, MinerUExtractionError, extract_pdf_to_markdown, user_facing_mineru_error


SUPPORTED_SOURCE_SUFFIXES = {".pdf", ".docx", ".doc", ".md", ".markdown", ".txt", ".rtf"}


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


def _read_pdf_text_layer_with_pymupdf(path: Path) -> str:
    import fitz

    parts: list[str] = []
    with fitz.open(path) as doc:
        for page_index, page in enumerate(doc, start=1):
            text = page.get_text(sort=True).strip()
            if not text:
                continue
            if doc.page_count > 1:
                parts.extend([f"## 第{page_index}页", "", text, ""])
            else:
                parts.append(text)
    return "\n".join(parts).strip()


def _read_pdf_text_layer_with_pypdf2(path: Path) -> str:
    import PyPDF2

    parts: list[str] = []
    with path.open("rb") as pdf_file_obj:
        pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
        for page_index, page in enumerate(pdf_reader.pages, start=1):
            text = (page.extract_text() or "").strip()
            if not text:
                continue
            if len(pdf_reader.pages) > 1:
                parts.extend([f"## 第{page_index}页", "", text, ""])
            else:
                parts.append(text)
    return "\n".join(parts).strip()


def _read_pdf_text_layer(path: Path) -> str:
    try:
        text = _read_pdf_text_layer_with_pymupdf(path)
        if text.strip():
            return text
    except Exception:
        pass
    try:
        return _read_pdf_text_layer_with_pypdf2(path)
    except Exception:
        return ""


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
        try:
            result = extract_pdf_to_markdown(path)
            result["markdown"] = normalize_math_markdown(result["markdown"])
            return result
        except (MinerUConfigError, MinerUExtractionError) as e:
            text_layer = _read_pdf_text_layer(path)
            if not text_layer.strip():
                raise
            return {
                "markdown": normalize_math_markdown(text_layer),
                "source": "pymupdf_pdf_fallback",
                "warnings": [
                    user_facing_mineru_error(e),
                    "已改用 PDF 文本层提取；扫描件或图片公式仍建议配置 MinerU/OCR",
                ],
                "metadata": {"fallback": "pdf_text_layer"},
            }

    raise ValueError("Unsupported source file type")
