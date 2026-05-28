from __future__ import annotations

import shutil
import sys
import tempfile
import unittest
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx import Document

from handwriting_document import (
    convert_to_handwritten,
    inspect_docx,
    markdown_from_source,
    normalize_markdown,
    paragraph_kind,
    plainify_latex_text,
    w,
)


class HandwritingDocumentTests(unittest.TestCase):
    def test_long_numbered_problem_prompt_is_body_not_heading(self):
        def para(text: str):
            p = etree.Element(w("p"))
            r = etree.SubElement(p, w("r"))
            t = etree.SubElement(r, w("t"))
            t.text = text
            return p

        self.assertEqual(paragraph_kind(para("3. 0.50,1.25,0.80,2.00 是取自总体 X 的样本")), "body")
        self.assertEqual(paragraph_kind(para("3.")), "heading")
        self.assertEqual(paragraph_kind(para("第3题")), "heading")

    def test_normalize_markdown_removes_images_and_backslash_math(self):
        markdown, info = normalize_markdown(
            "示例 \\(x^2+1\\)\n\n![](assets/a.png)\n\n\\[\\frac{1}{2}\\]\n"
        )
        self.assertEqual(info["removed_images"], 1)
        self.assertNotIn("![]", markdown)
        self.assertIn("$x^2+1$", markdown)
        self.assertIn("$$", markdown)

    def test_plainify_latex_text_handles_broad_relation_commands(self):
        text = plainify_latex_text(
            r"a\equiv b\pmod{n}, x\perp y, l\parallel m, "
            r"\therefore x\ne0, \because y\leq z, y\not\in B"
        )
        self.assertNotIn("\\", text)
        for token in ("a", "≡", "b", "mod", "n", "⊥", "∥", "∴", "≠", "∵", "≤", "∉", "B"):
            self.assertIn(token, text)

    def test_plainify_latex_text_handles_norm_and_infix_over(self):
        text = plainify_latex_text(r"\lVert v\rVert+\Vert x\Vert+a \over b+\text{if }x>0")
        self.assertNotRegex(text, r"\\|lVert|rVert|Vert|over")
        for token in ("‖v‖", "‖x‖", "(a)/(b)", "if x>0"):
            self.assertIn(token, text)

    def test_plainify_latex_text_preserves_optional_root_arrow_and_tag_content(self):
        text = plainify_latex_text(
            r"\sqrt[3]{x}+\xrightarrow[n\to0]{m\to\infty}y+"
            r"\begin{align}a&=b\tag{1}\label{eq:a}\end{align}"
        )
        self.assertNotRegex(text, r"\\|sqrt|xrightarrow|begin|align|tag|label|eq:a|&")
        for token in ("√[3](x)", "→_n→0^m→∞", "a", "b", "(1)"):
            self.assertIn(token, text)

    def test_plainify_latex_text_preserves_substack_rows_without_control_words(self):
        text = plainify_latex_text(
            r"\sum_{\substack{i=1\\j=2}}+\begin{subarray}{c}x\to0\\y\to1\end{subarray}"
        )
        self.assertNotRegex(text, r"\\|substack|subarray|begin|end|&")
        for token in ("sum", "i=1", "j=2", "x→0", "y→1"):
            self.assertIn(token, text)

    def test_plainify_latex_text_preserves_decorated_color_and_cancel_content(self):
        text = plainify_latex_text(
            r"\color{red}{x+y}+\boxed{a+b}+\cancel{z}+\overleftarrow{AB}+\left\lbrace x\middle|x>0\right\rbrace"
        )
        self.assertNotRegex(text, r"\\|color|boxed|cancel|overleftarrow|middle")
        for token in ("x+y", "[a+b]", "z", "AB", "{", "|", "x>0", "}"):
            self.assertIn(token, text)

    def test_plainify_latex_text_preserves_array_span_and_named_accent_content(self):
        text = plainify_latex_text(
            r"\begin{array}{c|c}\hline a&b\\\cline{1-2}\multicolumn{2}{c}{c+d}\end{array}"
            r"+\acute{x}+\breve{z}"
        )
        self.assertNotRegex(text, r"\\|array|hline|cline|multicolumn|acute|breve|c\|c")
        for token in ("a", "b", "c+d", "´x", "˘z"):
            self.assertIn(token, text)

    def test_plainify_latex_text_preserves_font_wrappers_and_named_symbol_content(self):
        text = plainify_latex_text(r"\mathscr{F}+\mathds{1}+\bm{x}+\Re z+\Im z+\ell+\hbar")
        self.assertNotRegex(text, r"\\|mathscr|mathds|bm|Re|Im|ell|hbar")
        for token in ("F", "1", "x", "ℜ", "z", "ℑ", "ℓ", "ℏ"):
            self.assertIn(token, text)

    def test_docx_inspection_detects_broad_latex_residuals(self):
        with tempfile.TemporaryDirectory(prefix="handwriting_docx_residual_") as tmp:
            docx = Path(tmp) / "raw_latex.docx"
            document = Document()
            document.add_paragraph(r"泄漏公式 a\equiv b, \therefore x\ne0")
            document.save(docx)
            report = inspect_docx(docx)

        self.assertTrue(report["latex_residuals"])

    def test_extract_text_layer_pdf_to_markdown(self):
        try:
            import fitz
        except ImportError:
            self.skipTest("PyMuPDF is not available")

        with tempfile.TemporaryDirectory(prefix="handwriting_pdf_extract_") as tmp:
            tmp_dir = Path(tmp)
            source = tmp_dir / "homework.pdf"
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((72, 72), "Homework 1")
            page.insert_text((72, 96), "Solve x^2 + y^2 = 1.")
            doc.save(source)
            doc.close()

            markdown = markdown_from_source(source, tmp_dir)
            text = markdown.read_text(encoding="utf-8")

        self.assertIn("# homework", text)
        self.assertIn("Homework 1", text)
        self.assertIn("x^2 + y^2", text)

    def test_convert_markdown_to_no_image_handwritten_pdf(self):
        for command in ("codex-md-to-docx", "codex-docx-to-pdf"):
            if shutil.which(command) is None:
                self.skipTest(f"{command} is not available")

        with tempfile.TemporaryDirectory(prefix="handwriting_doc_test_") as tmp:
            tmp_dir = Path(tmp)
            source = tmp_dir / "sample.md"
            source.write_text(
                "# 测试作业\n\n"
                "这是第 1 小题, 令 $X_n$ 表示状态.\n\n"
                "$$\n"
                "x+a\\ln|x|=t+C_1\n"
                "$$\n\n"
                "$$\n"
                "\\frac{d x}{d t}=x^2+1\n"
                "$$\n\n"
                "因此答案为 $x=\\tan(t+C)$.\n",
                encoding="utf-8",
            )
            report = convert_to_handwritten(source, tmp_dir / "out", output_stem="sample_handwritten")

        self.assertTrue(report["passed"])
        self.assertEqual(report["docx_inspect"]["media_files"], [])
        self.assertEqual(report["docx_inspect"]["drawing_objects"], 0)
        self.assertGreater(report["docx_inspect"]["math_objects"], 0)
        self.assertGreaterEqual(report["postprocess"]["demoted_inline_math"], 1)
        self.assertGreaterEqual(report["postprocess"]["demoted_display_math"], 1)
        self.assertGreater(report["docx_inspect"]["non_math_text_runs"], 0)
        self.assertEqual(report["docx_inspect"]["non_handwritten_text_runs"], 0)
        self.assertGreater(report["docx_inspect"]["positioned_text_runs"], 0)
        self.assertGreater(report["docx_inspect"]["spaced_text_runs"], 0)
        self.assertGreater(report["docx_inspect"]["scaled_text_runs"], 0)
        self.assertGreater(len(report["docx_inspect"]["distinct_text_sizes"]), 1)
        self.assertGreater(len(report["docx_inspect"]["distinct_text_colors"]), 1)
        self.assertEqual(report["pdf_inspect"]["image_count"], 0)
        self.assertTrue(report["handwritten_font_found"])


if __name__ == "__main__":
    unittest.main()
