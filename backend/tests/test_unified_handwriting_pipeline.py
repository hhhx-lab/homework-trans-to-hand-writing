from __future__ import annotations

import io
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from PIL import Image, ImageChops, ImageDraw, ImageFont
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import mineru_adapter
from handwriting_markdown_renderer import (
    FontCache,
    HandwritingRenderConfig,
    ScriptBox,
    TextBox,
    latex_to_debug_text,
    latex_to_box,
    markdown_render_debug_text,
    render_markdown_handwriting,
    write_images_to_docx_bytes,
)
from markdown_math import editable_docx_bytes, inspect_docx_math, normalize_math_markdown
from mineru_adapter import (
    MinerUConfigError,
    MinerUExtractionError,
    extract_pdf_to_markdown,
    sanitize_mineru_markdown,
    user_facing_mineru_error,
)
from source_extract import extract_source_to_markdown, safe_source_filename


FONT_PATH = Path(__file__).resolve().parents[1] / "font_assets" / "神韵英子楷书.ttf"


class UnifiedHandwritingPipelineTests(unittest.TestCase):
    def test_formula_debug_text_does_not_keep_latex_macros(self):
        expr = (
            r"\textstyle \frac{a_1}{b^2}+\sqrt{x}+\sum_{i=1}^{n}x_i+\ldots+"
            r"\begin{pmatrix}1&2\\3&4\end{pmatrix}"
        )
        text = latex_to_debug_text(expr, FONT_PATH)
        for macro in ("\\frac", "\\sqrt", "\\sum", "\\begin", "textstyle", "ldots", "dots"):
            self.assertNotIn(macro, text)
        self.assertIn("√", text)
        self.assertIn("∑", text)
        self.assertIn("…", text)
        self.assertIn("[[1,2];[3,4]]", text)

    def test_markdown_math_normalizer_repairs_lone_display_delimiter_and_docx_math(self):
        markdown = (
            r"P \left(Y_{n+k}=j_{n+k}, 1 \leq k \leq m / X_n=i\right)"
            "\n$$\n\n"
            r"其中 $x _ { n } = i.$ 右端只含当前状态 $\textstyle X _ { n } = i$ 和未来参数。"
        )
        normalized = normalize_math_markdown(markdown)
        self.assertIn("$$\nP \\left", normalized)
        self.assertIn("$X_{n} = i$", normalized)
        self.assertNotIn("\\textstyle", normalized)
        docx_info = inspect_docx_math(editable_docx_bytes(normalized))
        self.assertGreater(docx_info["office_math_objects"], 0)
        self.assertFalse(docx_info["has_latex_residuals"])

    def test_markdown_normalizer_preserves_inline_display_math_and_image_markers(self):
        markdown = "前文 ![图1](assets/a.png) 中间 $$x^2+1$$ 后文 <img src=\"b.png\" alt=\"图2\">"
        normalized = normalize_math_markdown(markdown)
        self.assertIn("![图1](assets/a.png)", normalized)
        self.assertIn("<img src=\"b.png\" alt=\"图2\">", normalized)
        self.assertIn("$$\nx^2+1\n$$", normalized)
        debug_text = markdown_render_debug_text(normalized, FONT_PATH)
        self.assertIn("[图片:![图1](assets/a.png)]", debug_text)
        self.assertIn("[图片:<img src=\"b.png\" alt=\"图2\">]", debug_text)
        self.assertIn("x^2+1", debug_text)

    def test_unknown_latex_commands_remain_visible_without_raw_latex(self):
        debug_text = latex_to_debug_text(r"\unknowncmd{x}+\overset{a}{b}", FONT_PATH)
        self.assertNotIn("\\", debug_text)
        self.assertIn("unknowncmd", debug_text)
        self.assertIn("x", debug_text)
        self.assertIn("a", debug_text)
        self.assertIn("b", debug_text)

    def test_common_math_decorations_have_visible_marks(self):
        debug_text = latex_to_debug_text(r"\overline{x}+\hat{y}+\vec{z}", FONT_PATH)
        self.assertIn("¯x", debug_text)
        self.assertIn("^y", debug_text)
        self.assertIn("→z", debug_text)

    def test_escaped_accent_commands_render_as_decorations(self):
        debug_text = latex_to_debug_text(r"\~{\pi}+\~\pi+\'{e}+\`{a}+\"{u}+x^\pi", FONT_PATH)
        self.assertNotIn("\\", debug_text)
        self.assertNotIn("~(", debug_text)
        self.assertIn("~π", debug_text)
        self.assertIn("´e", debug_text)
        self.assertIn("`a", debug_text)
        self.assertIn("¨u", debug_text)
        self.assertIn("x^π", debug_text)

    def test_common_latex_commands_render_readably_without_raw_latex(self):
        expr = (
            r"\dfrac{a_1}{b^2}+\binom{n}{k}+\overset{a}{b}+"
            r"\underset{0}{\lim}+\operatorname{Var}(X)+\mathbb{R}+"
            r"\iff+\mapsto+\left\{x\mid x\geq0\right\}"
        )
        debug_text = latex_to_debug_text(expr, FONT_PATH)
        self.assertNotIn("\\", debug_text)
        for macro in ("dfrac", "binom", "overset", "underset", "iff", "mapsto", "left", "right"):
            self.assertNotIn(macro, debug_text)
        for token in ("a", "1", "b", "2", "n", "k", "0", "lim", "Var", "X", "R", "x"):
            self.assertIn(token, debug_text)
        self.assertIn("↔", debug_text)
        self.assertIn("↦", debug_text)

    def test_markdown_debug_text_preserves_body_and_formula_tokens(self):
        markdown = (
            r"第12题：已知 A_n=3，求 $P(X_n=i)=\dfrac{a_1}{b^2}$ 的值。"
            "\n\n$$\n"
            r"f(x)=\begin{cases}x^2+1,&x\geq0\\-x,&x<0\end{cases}"
            "\n$$\n"
            "最后一行 z_9 不可丢。"
        )
        debug_text = markdown_render_debug_text(markdown, FONT_PATH)
        self.assertNotIn("\\", debug_text)
        for token in ("第12题", "已知", "A", "n", "3", "P", "X", "i", "a", "1", "b", "2", "f", "x", "0", "最后一行", "z", "9"):
            self.assertIn(token, debug_text)

    def test_plain_requests_with_math_use_markdown_renderer(self):
        from app import should_render_with_markdown_renderer

        self.assertTrue(should_render_with_markdown_renderer("plain", r"题目 \dfrac{a}{b}"))
        self.assertTrue(should_render_with_markdown_renderer("plain", r"题目 $x+1$"))
        self.assertTrue(should_render_with_markdown_renderer("plain", r"a\equiv b\pmod{n}"))
        self.assertTrue(should_render_with_markdown_renderer("plain", r"x\perp y,\ \angle ABC"))
        self.assertTrue(should_render_with_markdown_renderer("plain", r"\therefore x\ne 0"))
        self.assertFalse(should_render_with_markdown_renderer("plain", "纯文本内容"))

    def test_renderer_places_text_on_first_ruled_line_band(self):
        background = Image.new("RGB", (900, 1100), "white")
        top_margin = 70
        line_spacing = 96
        left_margin = 70
        right_margin = 70
        draw = ImageDraw.Draw(background)
        first_rule_y = top_margin + line_spacing
        draw.line((left_margin, first_rule_y, background.width - right_margin, first_rule_y), fill="black")
        font = ImageFont.truetype(str(FONT_PATH), 52)
        config = HandwritingRenderConfig(
            line_spacing=line_spacing,
            font_size=52,
            left_margin=left_margin,
            top_margin=top_margin,
            right_margin=right_margin,
            bottom_margin=70,
            word_spacing=1,
            perturb_x_sigma=0,
            perturb_y_sigma=0,
            ink_depth_sigma=0,
        )
        page = render_markdown_handwriting("横线内文字 $\\frac{1}{2}$", background, font, config)[0]
        ink_bbox = ImageChops.difference(background, page).getbbox()
        self.assertIsNotNone(ink_bbox)
        self.assertGreaterEqual(ink_bbox[1], top_margin)
        self.assertGreaterEqual(ink_bbox[3], first_rule_y - 12)

    def test_missing_math_symbol_glyphs_use_fallback_font(self):
        font = ImageFont.truetype(str(FONT_PATH), 52)
        self.assertIsNone(font.getmask("↔").getbbox())
        box = TextBox("↔", FontCache(font), 52)
        self.assertIsNotNone(box.font.getmask("↔").getbbox())

    def test_more_common_latex_commands_render_as_math_symbols(self):
        expr = (
            r"\mathcal{F}+\mathfrak{c}+a\equiv b\pmod{n},\quad "
            r"x\perp y,\quad l\parallel m,\quad \angle ABC,\quad "
            r"\therefore x\ne 0,\because y\leq z,\colon,\quad y\not\in B,\quad \|v\|"
        )
        debug_text = latex_to_debug_text(expr, FONT_PATH)
        self.assertNotIn("\\", debug_text)
        self.assertNotRegex(
            debug_text,
            r"mathfrak|equiv|pmod|quad|perp|parallel|angle|therefore|because|colon|not∈",
        )
        for token in ("F", "c", "≡", "mod", "n", "⊥", "∥", "∠", "ABC", "∴", "≠", "∵", "≤", ":", "∉", "B", "‖v‖"):
            self.assertIn(token, debug_text)

    def test_limits_commands_attach_scripts_to_big_operator(self):
        font = ImageFont.truetype(str(FONT_PATH), 52)
        box = latex_to_box(r"\sum\limits_{i=1}^{n} x_i", FontCache(font), 52)
        first = box.children[0]
        self.assertIsInstance(first, ScriptBox)
        self.assertEqual(first.base.debug_text(), "∑")
        self.assertEqual(first.sub.debug_text(), "i=1")
        self.assertEqual(first.sup.debug_text(), "n")

    def test_mineru_sanitize_preserves_image_placeholders(self):
        sanitized = sanitize_mineru_markdown("题面\n\n![scan](images/p1.png)\n\n答案")
        self.assertIn("题面", sanitized)
        self.assertIn("[图片:![scan](images/p1.png)]", sanitized)
        self.assertIn("答案", sanitized)

    def test_markdown_renderer_outputs_nonblank_image_and_docx(self):
        background = Image.new("RGB", (900, 1100), "white")
        font = ImageFont.truetype(str(FONT_PATH), 52)
        config = HandwritingRenderConfig(
            line_spacing=96,
            font_size=52,
            left_margin=70,
            top_margin=70,
            right_margin=70,
            bottom_margin=70,
            word_spacing=1,
            perturb_x_sigma=2,
            perturb_y_sigma=2,
            ink_depth_sigma=8,
        )
        pages = render_markdown_handwriting(
            "测试 $\\frac{1}{2}$\n\n$$\n\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}\n$$\n",
            background,
            font,
            config,
        )
        self.assertGreaterEqual(len(pages), 1)
        self.assertIsNotNone(ImageChops.difference(background, pages[0]).getbbox())

        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "page.png"
            pages[0].save(image_path)
            docx_bytes = write_images_to_docx_bytes([image_path])
        self.assertGreater(len(docx_bytes), 1000)

    def test_source_extract_markdown_txt_and_docx(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            md = tmp_dir / "source.md"
            md.write_text("# 标题\n\n$x+1$\n", encoding="utf-8")
            self.assertIn("$x+1$", extract_source_to_markdown(md)["markdown"])

            txt = tmp_dir / "source.txt"
            txt.write_text("普通文本", encoding="utf-8")
            self.assertEqual(extract_source_to_markdown(txt)["markdown"].strip(), "普通文本")

            docx = tmp_dir / "source.docx"
            document = Document()
            document.add_paragraph("Word 文本 x + 1")
            document.save(docx)
            extracted = extract_source_to_markdown(docx)["markdown"]
            self.assertIn("Word 文本", extracted)

    def test_safe_source_filename_keeps_suffix_for_chinese_names(self):
        self.assertEqual(safe_source_filename("随机过程三次作业答案.pdf", ".pdf"), "source.pdf")
        self.assertEqual(safe_source_filename("作业答案.docx", ".docx", "draft"), "draft.docx")
        self.assertEqual(safe_source_filename("homework.pdf", ".pdf"), "homework.pdf")

    def test_mineru_timeout_error_is_user_facing(self):
        message = user_facing_mineru_error(MinerUExtractionError("MinerU request failed: <urlopen error timed out>"))
        self.assertIn("PDF 识别服务 MinerU 连接超时", message)
        self.assertIn("MINERU_BASE_URL", message)

    def test_pdf_without_mineru_config_has_readable_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "source.pdf"
            pdf.write_bytes(b"%PDF-1.4\n%%EOF\n")
            with mock.patch.dict(os.environ, {}, clear=True):
                with self.assertRaises(MinerUConfigError) as ctx:
                    extract_source_to_markdown(pdf)
            self.assertIn("MINERU_BASE_URL", str(ctx.exception))

    def test_mineru_adapter_with_mocked_service(self):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            zf.writestr("result/full.md", "识别文本 $\\frac{1}{2}$\n")

        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "source.pdf"
            pdf.write_bytes(b"%PDF-1.4\n%%EOF\n")
            with mock.patch.dict(
                os.environ,
                {
                    "MINERU_BASE_URL": "https://mineru.example",
                    "MINERU_API_TOKEN": "token",
                    "MINERU_PUBLIC_BASE_URL": "https://handwrite.example",
                },
                clear=True,
            ), mock.patch.object(mineru_adapter, "_submit_task", return_value="task-1"), mock.patch.object(
                mineru_adapter, "_poll_task", return_value="https://download.example/result.zip"
            ), mock.patch.object(mineru_adapter, "_download", return_value=zip_buffer.getvalue()):
                result = extract_pdf_to_markdown(pdf)

        self.assertEqual(result["source"], "mineru")
        self.assertIn("识别文本", result["markdown"])


if __name__ == "__main__":
    unittest.main()
