from __future__ import annotations

from pathlib import Path
from typing import Any

import pypandoc

from markdown_math import normalize_math_markdown
from mineru_adapter import extract_pdf_to_markdown


SUPPORTED_SOURCE_SUFFIXES = {".pdf", ".docx", ".doc", ".md", ".markdown", ".txt", ".rtf"}


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
        result["markdown"] = normalize_math_markdown(result["markdown"])
        return result

    raise ValueError("Unsupported source file type")
