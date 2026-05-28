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
    _layout_inline,
    _text_to_boxes,
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

    def test_docx_math_inspector_detects_broad_latex_residuals(self):
        document = Document()
        document.add_paragraph(r"泄漏公式 a\equiv b, \therefore x\ne0")
        with tempfile.TemporaryDirectory() as tmp:
            docx = Path(tmp) / "raw_latex.docx"
            document.save(docx)
            docx_info = inspect_docx_math(docx.read_bytes())

        self.assertTrue(docx_info["has_latex_residuals"])

    def test_docx_math_inspector_detects_unknown_latex_residuals(self):
        document = Document()
        document.add_paragraph(r"泄漏公式 \partial f/\partial x 与 \unknowncmd{x}")
        with tempfile.TemporaryDirectory() as tmp:
            docx = Path(tmp) / "unknown_latex.docx"
            document.save(docx)
            docx_info = inspect_docx_math(docx.read_bytes())

        self.assertTrue(docx_info["has_latex_residuals"])

    def test_markdown_normalizer_wraps_common_bare_calculus_line(self):
        normalized = normalize_math_markdown(r"\partial f/\partial x=0")
        self.assertIn("$$\n\\partial f/\\partial x=0\n$$", normalized)
        docx_info = inspect_docx_math(editable_docx_bytes(normalized))
        self.assertGreater(docx_info["office_math_objects"], 0)
        self.assertFalse(docx_info["has_latex_residuals"])

    def test_markdown_normalizer_wraps_bare_math_inside_chinese_text(self):
        normalized = normalize_math_markdown(r"解：\partial f/\partial x=0，所以 a\equiv b\pmod{n}。")
        self.assertIn(r"解：$\partial f/\partial x=0$，所以 $a\equiv b\pmod{n}$。", normalized)
        docx_info = inspect_docx_math(editable_docx_bytes(normalized))
        self.assertGreaterEqual(docx_info["office_math_objects"], 2)
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

    def test_infix_over_renders_as_fraction(self):
        self.assertEqual("(a)/(b)", latex_to_debug_text(r"a \over b", FONT_PATH))

    def test_text_command_preserves_inner_spaces(self):
        self.assertEqual("if x>0", latex_to_debug_text(r"\text{if }x>0", FONT_PATH))

    def test_mathop_and_starred_operatorname_do_not_render_command_names(self):
        mathop_text = latex_to_debug_text(r"\mathop{\lim}\limits_{x\to0} f(x)", FONT_PATH)
        operator_text = latex_to_debug_text(r"\operatorname*{arg\,max}_{x} f(x)", FONT_PATH)
        self.assertNotRegex(mathop_text + operator_text, r"mathop|operatorname|\\")
        self.assertIn("lim_x→0", mathop_text)
        self.assertIn("arg max_x", operator_text)

    def test_optional_root_index_and_extensible_arrow_labels_are_preserved(self):
        debug_text = latex_to_debug_text(
            r"\sqrt[3]{x^2+y}+\xrightarrow[n\to0]{m\to\infty} y+\xleftarrow{k} z",
            FONT_PATH,
        )
        self.assertNotRegex(debug_text, r"\\|sqrt|xrightarrow|xleftarrow")
        for token in ("√[3]", "x^2", "y", "→", "n→0", "m→∞", "←", "k", "z"):
            self.assertIn(token, debug_text)

    def test_alignment_environment_and_equation_metadata_do_not_render_control_words(self):
        debug_text = latex_to_debug_text(
            r"\begin{align}a&=b+c\\d&=e+f\tag{1}\label{eq:one}\end{align}",
            FONT_PATH,
        )
        self.assertNotRegex(debug_text, r"\\|begin|end|align|tag|label|&|eq:one")
        for token in ("a", "=b+c", "d", "=e+f", "(1)"):
            self.assertIn(token, debug_text)

    def test_stacked_limit_helpers_do_not_render_control_words_or_alignment_specs(self):
        debug_text = latex_to_debug_text(
            r"\sum_{\substack{i=1\\j=2}}^n a_{ij}+"
            r"\lim_{\begin{subarray}{c}x\to0\\y\to1\end{subarray}} f(x,y)",
            FONT_PATH,
        )
        self.assertNotRegex(debug_text, r"\\|substack|subarray|begin|end|_c")
        for token in ("∑", "i=1", "j=2", "^n", "a", "ij", "lim", "x→0", "y→1", "f"):
            self.assertIn(token, debug_text)

    def test_matrix_variants_render_as_matrix_rows_without_control_words(self):
        debug_text = latex_to_debug_text(
            r"\begin{smallmatrix}1&2\\3&4\end{smallmatrix}+"
            r"\begin{Bmatrix}a&b\\c&d\end{Bmatrix}+"
            r"\begin{Vmatrix}p&q\\r&s\end{Vmatrix}",
            FONT_PATH,
        )
        self.assertNotRegex(debug_text, r"\\|begin|end|smallmatrix|Bmatrix|Vmatrix|&")
        for token in ("[[1,2];[3,4]]", "[[a,b];[c,d]]", "[[p,q];[r,s]]"):
            self.assertIn(token, debug_text)

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

    def test_legacy_textfileprocess_uses_markdown_formula_extraction(self):
        from app import extract_textfileprocess_content

        with tempfile.TemporaryDirectory() as tmp:
            docx = Path(tmp) / "formula.docx"
            docx.write_bytes(editable_docx_bytes(r"旧入口公式 $\frac{a_1}{b^2}+\sum_{i=1}^{n}x_i$ 完成"))
            result = extract_textfileprocess_content(docx)
        self.assertIn("旧入口公式", result["text"])
        self.assertIn(r"\frac", result["text"])
        self.assertIn(r"\sum", result["text"])
        self.assertIn("完成", result["text"])

    def test_raw_latex_commands_in_text_are_rendered_as_math(self):
        debug_text = markdown_render_debug_text(
            r"题目 a\equiv b\pmod{n} 结束；因此 \therefore x\ne0，且 y\not\in B。",
            FONT_PATH,
        )
        self.assertNotIn("\\", debug_text)
        for token in ("题目", "a", "≡", "b", "mod", "n", "结束", "∴", "x", "≠", "0", "y", "∉", "B"):
            self.assertIn(token, debug_text)

    def test_raw_latex_commands_do_not_swallow_adjacent_plain_text(self):
        debug_text = markdown_render_debug_text(r"This is a \frac{1}{2} test.", FONT_PATH)
        self.assertEqual("This is a (1)/(2) test.", debug_text)
        self.assertNotIn("\\frac", debug_text)

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

    def test_oversized_inline_formula_wraps_within_available_width(self):
        font = ImageFont.truetype(str(FONT_PATH), 52)
        fonts = FontCache(font)
        formula = "$" + "+".join(f"a_{i}" for i in range(1, 18)) + "$"
        lines = _layout_inline(_text_to_boxes("长公式 " + formula + " 结束", fonts, 52), 360, 1)
        self.assertGreater(len(lines), 1)
        self.assertTrue(all(line.width <= 360 for line in lines))

    def test_oversized_display_formula_stays_within_right_margin(self):
        background = Image.new("RGB", (600, 900), "white")
        font = ImageFont.truetype(str(FONT_PATH), 52)
        config = HandwritingRenderConfig(
            line_spacing=90,
            font_size=52,
            left_margin=70,
            top_margin=70,
            right_margin=70,
            bottom_margin=70,
            word_spacing=1,
            perturb_x_sigma=0,
            perturb_y_sigma=0,
            ink_depth_sigma=0,
        )
        page = render_markdown_handwriting(
            "$$\nabcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890\n$$",
            background,
            font,
            config,
        )[0]
        ink_bbox = ImageChops.difference(background, page).getbbox()
        self.assertIsNotNone(ink_bbox)
        self.assertLessEqual(ink_bbox[2], background.width - config.right_margin)

    def test_tall_display_formula_stays_above_bottom_margin(self):
        background = Image.new("RGB", (500, 320), "white")
        font = ImageFont.truetype(str(FONT_PATH), 72)
        config = HandwritingRenderConfig(
            line_spacing=80,
            font_size=72,
            left_margin=50,
            top_margin=40,
            right_margin=50,
            bottom_margin=40,
            word_spacing=1,
            perturb_x_sigma=0,
            perturb_y_sigma=0,
            perturb_theta_sigma=0,
            ink_depth_sigma=0,
        )
        page = render_markdown_handwriting(
            r"$$\frac{\frac{a}{b}}{\frac{c}{d}}$$",
            background,
            font,
            config,
        )[0]
        ink_bbox = ImageChops.difference(background, page).getbbox()
        self.assertIsNotNone(ink_bbox)
        self.assertLessEqual(ink_bbox[3], background.height - config.bottom_margin)

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

    def test_norm_delimiter_commands_render_as_double_bars(self):
        debug_text = latex_to_debug_text(r"\lVert v\rVert+\Vert x\Vert", FONT_PATH)
        self.assertNotRegex(debug_text, r"lVert|rVert|\\")
        self.assertEqual("‖v‖+‖x‖", debug_text)

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

            formula_docx = tmp_dir / "formula.docx"
            formula_docx.write_bytes(editable_docx_bytes(r"Word 公式 $\frac{a_1}{b^2}+\sum_{i=1}^{n}x_i$ 完成"))
            formula_extracted = extract_source_to_markdown(formula_docx)["markdown"]
            self.assertIn("Word 公式", formula_extracted)
            self.assertIn(r"\frac", formula_extracted)
            self.assertIn(r"\sum", formula_extracted)
            self.assertIn("完成", formula_extracted)

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
