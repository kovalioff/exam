#!/usr/bin/env python3

from __future__ import annotations

import argparse
import html
import re
from pathlib import Path

from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, Preformatted, SimpleDocTemplate, Spacer


def register_fonts() -> None:
    font_dir = Path("/usr/share/fonts/truetype/dejavu")
    pdfmetrics.registerFont(TTFont("DejaVuSerif", str(font_dir / "DejaVuSerif.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVuSerif-Bold", str(font_dir / "DejaVuSerif-Bold.ttf")))
    pdfmetrics.registerFont(TTFont("DejaVuSansMono", str(font_dir / "DejaVuSansMono.ttf")))


def build_styles() -> dict[str, ParagraphStyle]:
    return {
        "title": ParagraphStyle(
            "title",
            fontName="DejaVuSerif-Bold",
            fontSize=22,
            leading=26,
            spaceAfter=10,
        ),
        "heading": ParagraphStyle(
            "heading",
            fontName="DejaVuSerif-Bold",
            fontSize=16,
            leading=20,
            spaceBefore=14,
            spaceAfter=8,
            textColor=HexColor("#1a1a1a"),
        ),
        "body": ParagraphStyle(
            "body",
            fontName="DejaVuSerif",
            fontSize=11,
            leading=15,
            alignment=TA_JUSTIFY,
            spaceAfter=6,
        ),
        "bullet": ParagraphStyle(
            "bullet",
            fontName="DejaVuSerif",
            fontSize=11,
            leading=15,
            leftIndent=14,
            firstLineIndent=-10,
            spaceAfter=2,
        ),
        "math": ParagraphStyle(
            "math",
            fontName="DejaVuSansMono",
            fontSize=10,
            leading=13,
            leftIndent=10,
            backColor=HexColor("#f7f7f7"),
            borderColor=HexColor("#dddddd"),
            borderWidth=0.5,
            borderPadding=6,
            borderRadius=3,
            spaceBefore=4,
            spaceAfter=6,
        ),
    }


def apply_inline_markup(raw: str) -> str:
    tokens: dict[str, str] = {}

    def protect(pattern: str, formatter, text: str) -> str:
        def repl(match: re.Match[str]) -> str:
            token = f"@@TOKEN{len(tokens)}@@"
            tokens[token] = formatter(match.group(1))
            return token

        return re.sub(pattern, repl, text)

    raw = protect(
        r"`([^`]+)`",
        lambda s: f'<font name="DejaVuSansMono">{html.escape(s)}</font>',
        raw,
    )
    raw = protect(
        r"\$([^$]+)\$",
        lambda s: f'<font name="DejaVuSansMono">{html.escape(s)}</font>',
        raw,
    )

    rendered = html.escape(raw)
    rendered = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", rendered)
    rendered = re.sub(r"\*([^*]+)\*", r"\1", rendered)

    for token, replacement in tokens.items():
        rendered = rendered.replace(html.escape(token), replacement)

    return rendered


def render_markdown_to_story(text: str, styles: dict[str, ParagraphStyle]):
    story = []
    lines = text.splitlines()
    paragraph: list[str] = []
    math_block: list[str] = []
    in_math = False
    seen_question = False

    def flush_paragraph() -> None:
        nonlocal paragraph
        if paragraph:
            content = " ".join(paragraph).strip()
            story.append(Paragraph(apply_inline_markup(content), styles["body"]))
            paragraph = []

    for line in lines:
        stripped = line.strip()

        if in_math:
            if stripped == "$$":
                story.append(
                    Preformatted(
                        "\n".join(math_block),
                        styles["math"],
                    )
                )
                story.append(Spacer(1, 2))
                math_block = []
                in_math = False
            else:
                math_block.append(line)
            continue

        if stripped == "$$":
            flush_paragraph()
            in_math = True
            math_block = []
            continue

        if not stripped:
            flush_paragraph()
            continue

        if line.startswith("# "):
            flush_paragraph()
            story.append(Paragraph(apply_inline_markup(line[2:].strip()), styles["title"]))
            continue

        if line.startswith("## "):
            flush_paragraph()
            title = line[3:].strip()
            if title.startswith("Pytanie "):
                if seen_question:
                    story.append(PageBreak())
                seen_question = True
            story.append(Paragraph(apply_inline_markup(title), styles["heading"]))
            continue

        if line.startswith("- "):
            flush_paragraph()
            story.append(Paragraph(f"• {apply_inline_markup(line[2:].strip())}", styles["bullet"]))
            continue

        paragraph.append(stripped)

    flush_paragraph()
    return story


def main() -> None:
    parser = argparse.ArgumentParser(description="Render markdown notes to PDF.")
    parser.add_argument("input_markdown", type=Path)
    parser.add_argument("output_pdf", type=Path)
    args = parser.parse_args()

    register_fonts()
    styles = build_styles()

    text = args.input_markdown.read_text(encoding="utf-8")
    story = render_markdown_to_story(text, styles)

    doc = SimpleDocTemplate(
        str(args.output_pdf),
        pagesize=A4,
        leftMargin=45,
        rightMargin=45,
        topMargin=45,
        bottomMargin=45,
        title="Opracowanie wybranych pytań z prezentacji",
        author="Cursor",
    )
    doc.build(story)


if __name__ == "__main__":
    main()
