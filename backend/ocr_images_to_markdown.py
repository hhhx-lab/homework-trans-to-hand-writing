from __future__ import annotations

import argparse
import re
from pathlib import Path


GREEK_COMMANDS = {
    "alpha",
    "beta",
    "gamma",
    "delta",
    "mu",
    "sigma",
    "theta",
    "lambda",
    "pi",
    "chi",
    "Lambda",
    "varLambda",
}


def image_sort_key(path: Path) -> tuple[int, int, str]:
    match = re.search(r"题(\d+)(?:_(\d+))?", path.stem)
    if match:
        return int(match.group(1)), int(match.group(2) or 0), path.name
    match = re.search(r"page[_-]?(\d+)", path.stem, re.I)
    if match:
        return int(match.group(1)), 0, path.name
    return 999999, 0, path.name


def cleanup_ocr_markdown(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    replacements = {
        "\\operatorname{l n}": "\\ln",
        "\\operatorname { l n }": "\\ln",
        "\\operatorname{e x p}": "\\exp",
        "\\operatorname { e x p }": "\\exp",
        "\\boldsymbol{X}": "X",
        "\\boldsymbol{Y}": "Y",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    text = text.replace("日于", "由于")
    text = text.replace("暨信", "置信")
    text = text.replace("直表", "查表")
    text = text.replace("956能", "95%的")
    text = text.replace("l}-\\boldsymbol{\\alpha}", "1-\\alpha")
    text = text.replace("\\cfrac", "\\frac")
    text = re.sub(r"\\(?:boldsymbol|mathbf|mathbb|textbf)\s*\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\{\\bf\s*\\,\s*([^{}]+)\}", r"\1", text)
    text = re.sub(r"\{\\bf\s*([^{}]+)\}", r"\1", text)
    text = re.sub(r"\\bf\s*([A-Za-z0-9]+)", r"\1", text)
    text = re.sub(r"\{\\cal\s+([A-Za-z])\}", r"\1", text)

    def strip_boldmath(match: re.Match[str]) -> str:
        name = match.group(1).strip()
        if name in GREEK_COMMANDS:
            return f"\\{name}"
        return name

    text = re.sub(r"\\mathrm\s*\{\\boldmath~\s*\\?([A-Za-z]+)\s*~\}", strip_boldmath, text)
    text = re.sub(r"\\mathrm\s*\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\small\s*\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"\\[,;:!]", " ", text)
    text = re.sub(r"(?<=\d)\s+\.\s+(?=\d)", ".", text)
    previous = None
    while previous != text:
        previous = text
        text = re.sub(r"(\d+\.\d+)\s+(\d)", r"\1\2", text)
    text = re.sub(r"\s+([,.;:，。；：])", r"\1", text)
    cleaned_lines: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped in {"中会", "321", "323", "325", "328", "2nx 2nx", "0.13"}:
            continue
        if stripped.startswith("·CN") or stripped.startswith("CN"):
            cleaned_lines.append("某种材料抗压强度数据如下:")
            continue
        if stripped == "482493457471510446435418394469":
            cleaned_lines.append("482, 493, 457, 471, 510, 446, 435, 418, 394, 469")
            continue
        if stripped.rstrip("，,") == "154550536065708390":
            cleaned_lines.append("15, 45, 50, 53, 60, 65, 70, 83, 90,")
            continue
        if stripped.startswith("13.汉总体"):
            cleaned_lines.append("13.设总体 $X$ 的密度函数为")
            continue
        if "w h e r s i s h o l" in stripped:
            cleaned_lines.append(
                "这里 $s_w^2=\\frac{(n_1-1)s_x^2+(n_2-1)s_y^2}{n_1+n_2-2}=54.0043$。"
            )
            continue
        cleaned_lines.append(line)
    text = "\n".join(cleaned_lines)
    text = text.replace("（2）由手", "（2）由于")
    text = text.replace("(3求", "（3）求")
    text = text.replace("四南", "因而")
    text = text.replace("名度图效力", "密度函数为")
    text = text.replace("E x p", "Exp")
    text = text.replace("\\vdash", "-")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def ocr_images_to_markdown(images: list[Path], *, title: str, resized_shape: int) -> str:
    try:
        from pix2text import Pix2Text
    except ImportError as exc:
        raise RuntimeError("pix2text is required for image OCR. Use the h2_math_ocr_work venv or install pix2text.") from exc

    p2t = Pix2Text.from_config(enable_formula=True, enable_table=False, device="cpu")
    parts = [f"# {title}", ""]
    current_problem: int | None = None
    for image in sorted(images, key=image_sort_key):
        problem_match = re.search(r"题(\d+)", image.stem)
        problem_no = int(problem_match.group(1)) if problem_match else None
        if problem_no is not None and problem_no != current_problem:
            current_problem = problem_no
            parts.extend(["", f"## 第{problem_no}题", ""])
        text = p2t.recognize_text_formula(str(image), resized_shape=resized_shape, return_text=True)
        parts.append(cleanup_ocr_markdown(str(text)))
        parts.append("")
    return "\n".join(parts).strip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="OCR cropped math-homework images into Markdown.")
    parser.add_argument("image_dir", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--title", default="数理统计作业答案")
    parser.add_argument("--resized-shape", type=int, default=960)
    args = parser.parse_args()

    images = [
        path
        for path in args.image_dir.iterdir()
        if path.suffix.lower() in {".png", ".jpg", ".jpeg"} and "预览" not in path.name and "contact" not in path.name
    ]
    if not images:
        raise RuntimeError(f"No OCR images found in {args.image_dir}")
    markdown = ocr_images_to_markdown(images, title=args.title, resized_shape=args.resized_shape)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(markdown, encoding="utf-8")
    print(args.out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
