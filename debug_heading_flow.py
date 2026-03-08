from __future__ import annotations

import base64
import re
import time
from pathlib import Path

import pythoncom

from services.applier import FormatApplier
from services.extractor import FormatExtractor
from services.word_com import WordComClient

ROOT = Path(__file__).resolve().parent
TEMPLATE_PATH = ROOT / "调试模板.docx"
TARGET_PATH = ROOT / "调试目标.docx"
OUTPUT_PATH = ROOT / "调试输出.docx"
IMAGE_PATH = ROOT / "调试图片.png"

WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_LINE_SPACE_SINGLE = 0
WD_LINE_SPACE_ONE_POINT_FIVE = 1
WD_LINE_SPACE_DOUBLE = 2
WD_LINE_SPACE_MULTIPLE = 5
RPC_E_CALL_REJECTED = -2147418111

TEST_IMAGE_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wn0K3sAAAAASUVORK5CYII="
)


def ensure_test_image() -> None:
    IMAGE_PATH.write_bytes(base64.b64decode(TEST_IMAGE_BASE64))


def _retry_word_call(callback, *, attempts: int = 20, delay: float = 0.1):
    for attempt in range(attempts):
        try:
            return callback()
        except Exception as exc:
            code = exc.args[0] if getattr(exc, "args", None) else None
            if code != RPC_E_CALL_REJECTED or attempt >= attempts - 1:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(delay)
    raise RuntimeError("Word 调用重试失败")


def _document_end_range(doc):
    end = max(int(doc.Content.End) - 1, 0)
    return doc.Range(end, end)


def _set_line_spacing(fmt, mode: str, value: float) -> None:
    if mode == "单倍":
        fmt.LineSpacingRule = WD_LINE_SPACE_SINGLE
        fmt.LineSpacing = 12
        return
    if mode == "1.5 倍":
        fmt.LineSpacingRule = WD_LINE_SPACE_ONE_POINT_FIVE
        fmt.LineSpacing = 18
        return
    if mode == "2 倍":
        fmt.LineSpacingRule = WD_LINE_SPACE_DOUBLE
        fmt.LineSpacing = 24
        return
    fmt.LineSpacingRule = WD_LINE_SPACE_MULTIPLE
    fmt.LineSpacing = max(value * 12, 12)


def _set_paragraph_indents(fmt, first_line_indent: float) -> None:
    try:
        fmt.CharacterUnitLeftIndent = 0.0
        fmt.CharacterUnitRightIndent = 0.0
        fmt.CharacterUnitFirstLineIndent = first_line_indent
    except Exception:
        fmt.LeftIndent = 0
        fmt.RightIndent = 0
        fmt.FirstLineIndent = first_line_indent * 12


def _apply_direct_paragraph_format(
    paragraph,
    *,
    font_name: str,
    font_size: float,
    bold: bool,
    alignment: int,
    first_line_indent: float,
    line_spacing_mode: str,
    line_spacing: float,
) -> None:
    para_range = paragraph.Range
    fmt = paragraph.Format
    para_range.Font.NameFarEast = font_name
    para_range.Font.Name = font_name
    para_range.Font.Size = font_size
    para_range.Font.Bold = -1 if bold else 0
    para_range.Font.Italic = 0
    fmt.Alignment = alignment
    _set_paragraph_indents(fmt, first_line_indent)
    fmt.SpaceBefore = 0
    fmt.SpaceAfter = 0
    _set_line_spacing(fmt, line_spacing_mode, line_spacing)


def _append_paragraph(
    doc,
    *,
    style_name: str,
    text: str,
    font_name: str,
    font_size: float,
    bold: bool,
    alignment: int,
    first_line_indent: float,
    line_spacing_mode: str,
    line_spacing: float,
):
    insert_range = _retry_word_call(lambda: _document_end_range(doc))
    _retry_word_call(lambda: setattr(insert_range, "Text", text))
    paragraph = _retry_word_call(lambda: insert_range.Paragraphs(1))
    _retry_word_call(lambda: setattr(paragraph, "Style", style_name))
    _apply_direct_paragraph_format(
        paragraph,
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        alignment=alignment,
        first_line_indent=first_line_indent,
        line_spacing_mode=line_spacing_mode,
        line_spacing=line_spacing,
    )
    _retry_word_call(lambda: paragraph.Range.InsertParagraphAfter())
    return paragraph


def _append_inline_image(doc, *, first_line_indent: float, alignment: int) -> None:
    insert_range = _retry_word_call(lambda: _document_end_range(doc))
    shape = _retry_word_call(lambda: insert_range.InlineShapes.AddPicture(str(IMAGE_PATH)))
    paragraph = _retry_word_call(lambda: shape.Range.Paragraphs(1))
    fmt = paragraph.Format
    fmt.Alignment = alignment
    _set_paragraph_indents(fmt, first_line_indent)
    fmt.SpaceBefore = 0
    fmt.SpaceAfter = 0
    _set_line_spacing(fmt, "单倍", 1.0)
    _retry_word_call(lambda: paragraph.Range.InsertParagraphAfter())


def create_template_document() -> None:
    with WordComClient(visible=False) as client:
        doc = client.app.Documents.Add()
        try:
            _append_paragraph(
                doc,
                style_name="标题 1",
                text="一、 模板一级标题",
                font_name="黑体",
                font_size=18.0,
                bold=True,
                alignment=WD_ALIGN_PARAGRAPH_CENTER,
                first_line_indent=0.0,
                line_spacing_mode="单倍",
                line_spacing=1.0,
            )
            _append_paragraph(
                doc,
                style_name="标题 2",
                text="（一） 模板二级标题",
                font_name="黑体",
                font_size=16.0,
                bold=True,
                alignment=WD_ALIGN_PARAGRAPH_LEFT,
                first_line_indent=0.0,
                line_spacing_mode="1.5 倍",
                line_spacing=1.5,
            )
            _append_paragraph(
                doc,
                style_name="标题 3",
                text="1、 模板三级标题",
                font_name="黑体",
                font_size=14.0,
                bold=True,
                alignment=WD_ALIGN_PARAGRAPH_LEFT,
                first_line_indent=0.0,
                line_spacing_mode="1.5 倍",
                line_spacing=1.5,
            )
            _append_paragraph(
                doc,
                style_name="正文",
                text="这是模板正文第一段。",
                font_name="宋体",
                font_size=12.0,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH_LEFT,
                first_line_indent=2.0,
                line_spacing_mode="1.5 倍",
                line_spacing=1.5,
            )
            _append_inline_image(doc, first_line_indent=0.0, alignment=WD_ALIGN_PARAGRAPH_CENTER)
            _retry_word_call(lambda: doc.SaveAs2(str(TEMPLATE_PATH)))
        finally:
            client.close_document(doc, save_changes=False)


def create_target_document() -> None:
    with WordComClient(visible=False) as client:
        doc = client.app.Documents.Add()
        try:
            _append_paragraph(
                doc,
                style_name="标题 1",
                text="第一章 项目概述",
                font_name="宋体",
                font_size=12.0,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH_LEFT,
                first_line_indent=2.0,
                line_spacing_mode="多倍",
                line_spacing=2.0,
            )
            _append_paragraph(
                doc,
                style_name="标题 2",
                text="1.1 建设目标",
                font_name="宋体",
                font_size=12.0,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH_CENTER,
                first_line_indent=2.0,
                line_spacing_mode="多倍",
                line_spacing=2.0,
            )
            _append_paragraph(
                doc,
                style_name="标题 3",
                text="1.1.1 总体原则",
                font_name="宋体",
                font_size=12.0,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH_CENTER,
                first_line_indent=2.0,
                line_spacing_mode="多倍",
                line_spacing=2.0,
            )
            _append_paragraph(
                doc,
                style_name="正文",
                text="目标文档正文第一段。",
                font_name="宋体",
                font_size=12.0,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH_LEFT,
                first_line_indent=2.0,
                line_spacing_mode="1.5 倍",
                line_spacing=1.5,
            )
            _append_paragraph(
                doc,
                style_name="标题 1",
                text="第二章 实施路径",
                font_name="宋体",
                font_size=12.0,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH_LEFT,
                first_line_indent=2.0,
                line_spacing_mode="多倍",
                line_spacing=2.0,
            )
            _append_paragraph(
                doc,
                style_name="标题 2",
                text="1.2 推进步骤",
                font_name="宋体",
                font_size=12.0,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH_CENTER,
                first_line_indent=2.0,
                line_spacing_mode="多倍",
                line_spacing=2.0,
            )
            _append_paragraph(
                doc,
                style_name="标题 3",
                text="1.2.1 阶段安排",
                font_name="宋体",
                font_size=12.0,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH_CENTER,
                first_line_indent=2.0,
                line_spacing_mode="多倍",
                line_spacing=2.0,
            )
            _append_inline_image(doc, first_line_indent=2.0, alignment=WD_ALIGN_PARAGRAPH_LEFT)
            _append_paragraph(
                doc,
                style_name="正文",
                text="1.2 这是普通正文，不应该被识别成标题。",
                font_name="宋体",
                font_size=12.0,
                bold=False,
                alignment=WD_ALIGN_PARAGRAPH_LEFT,
                first_line_indent=2.0,
                line_spacing_mode="1.5 倍",
                line_spacing=1.5,
            )
            _retry_word_call(lambda: doc.SaveAs2(str(TARGET_PATH)))
        finally:
            client.close_document(doc, save_changes=False)


def dump_outline(title: str, outline: list[dict]) -> None:
    print(title)
    for item in outline:
        print(item)
    print()


def _heading_level_from_style(style_name: str) -> int:
    match = re.search(r"(?:标题|Heading)\s*(\d+)", style_name or "", re.IGNORECASE)
    if not match:
        return 0
    return int(match.group(1))


def _same_float(actual: float, expected: float, tolerance: float = 0.15) -> bool:
    return abs(float(actual) - float(expected)) <= tolerance


def inspect_output_document(template_result) -> None:
    expected_heading_texts = [
        "一、 项目概述",
        "（一） 建设目标",
        "1、 总体原则",
        "二、 实施路径",
        "（二） 推进步骤",
        "1、 阶段安排",
    ]
    actual_heading_texts: list[str] = []
    errors: list[str] = []

    with WordComClient(visible=False) as client:
        doc = client.open_document(str(OUTPUT_PATH), read_only=True)
        try:
            page_count = client.get_page_count(doc)
            target_range = client.get_page_span_range(doc, 1, page_count)
            print("输出文档段落：")
            for paragraph in client.iter_paragraphs(target_range):
                info = client.get_paragraph_info(paragraph)
                text = info["text"].strip()
                has_inline_shape = False
                try:
                    has_inline_shape = paragraph.Range.InlineShapes.Count > 0
                except Exception:
                    has_inline_shape = False
                if not text and not has_inline_shape:
                    continue
                print(
                    {
                        "text": text,
                        "style_name": info.get("style_name"),
                        "font_name": info.get("font_name"),
                        "font_size": info.get("font_size"),
                        "bold": info.get("bold"),
                        "alignment": info.get("alignment"),
                        "first_line_indent": info.get("first_line_indent_chars"),
                        "line_spacing_mode": info.get("line_spacing_mode"),
                        "line_spacing": info.get("line_spacing"),
                    }
                )
                level = _heading_level_from_style(str(info.get("style_name", "")))
                if level:
                    actual_heading_texts.append(text)
                    profile = template_result.catalog.headings[f"标题{level}"]
                    if info.get("font_name") != profile.font_name:
                        errors.append(f"标题{level} 字体未按模板应用：{info.get('font_name')} != {profile.font_name}")
                    if not _same_float(info.get("font_size", 0), profile.font_size):
                        errors.append(f"标题{level} 字号未按模板应用：{info.get('font_size')} != {profile.font_size}")
                    if bool(info.get("bold")) != bool(profile.bold):
                        errors.append(f"标题{level} 加粗未按模板应用：{info.get('bold')} != {profile.bold}")
                    if info.get("alignment") != profile.alignment:
                        errors.append(f"标题{level} 对齐未按模板应用：{info.get('alignment')} != {profile.alignment}")
                    if not _same_float(info.get("first_line_indent_chars", 0), profile.first_line_indent):
                        errors.append(
                            f"标题{level} 首行缩进未按模板应用：{info.get('first_line_indent_chars')} != {profile.first_line_indent}"
                        )

            if actual_heading_texts != expected_heading_texts:
                errors.append(f"标题编号结果不符合预期：{actual_heading_texts} != {expected_heading_texts}")

            inline_shapes = client.iter_inline_shapes(target_range)
            print("输出文档图片段落：")
            for index, shape in enumerate(inline_shapes, start=1):
                paragraph = shape.Range.Paragraphs(1)
                info = client.get_paragraph_info(paragraph)
                print(
                    index,
                    {
                        "alignment": info.get("alignment"),
                        "left_indent": info.get("left_indent_chars"),
                        "right_indent": info.get("right_indent_chars"),
                        "first_line_indent": info.get("first_line_indent_chars"),
                        "line_spacing_mode": info.get("line_spacing_mode"),
                        "line_spacing": info.get("line_spacing"),
                    },
                )
                if not _same_float(info.get("left_indent_chars", 0), 0.0):
                    errors.append(f"图片左缩进未清零：{info.get('left_indent_chars')}")
                if not _same_float(info.get("right_indent_chars", 0), 0.0):
                    errors.append(f"图片右缩进未清零：{info.get('right_indent_chars')}")
                if not _same_float(info.get("first_line_indent_chars", 0), 0.0):
                    errors.append(f"图片首行缩进未清零：{info.get('first_line_indent_chars')}")

            if errors:
                raise AssertionError("\n".join(errors))
        finally:
            client.close_document(doc, save_changes=False)


def main() -> None:
    for path in [TEMPLATE_PATH, TARGET_PATH, OUTPUT_PATH, IMAGE_PATH]:
        if path.exists():
            path.unlink()

    ensure_test_image()
    create_template_document()
    create_target_document()

    extractor = FormatExtractor(WordComClient(visible=False))
    template_result = extractor.extract_template_document(str(TEMPLATE_PATH))
    target_outline, outline_logs = extractor.extract_outline_from_page_range(str(TARGET_PATH), 1, 1)

    dump_outline("模板 outline", template_result.outline)
    dump_outline("目标 outline", target_outline)
    print("目标 outline 日志", outline_logs)
    print("模板标题参数")
    for key in ["标题1", "标题2", "标题3"]:
        profile = template_result.catalog.headings.get(key)
        if profile is None:
            continue
        print(
            key,
            {
                "font_name": profile.font_name,
                "font_size": profile.font_size,
                "bold": profile.bold,
                "alignment": profile.alignment,
                "first_line_indent": profile.first_line_indent,
                "numbering_style": profile.numbering_style,
                "numbering_pattern": profile.numbering_pattern,
                "renumber": profile.renumber,
                "clear_existing_numbering": profile.clear_existing_numbering,
            },
        )
    print("模板图片参数", template_result.catalog.image)
    print()

    applier = FormatApplier(WordComClient(visible=False))
    result = applier.apply_to_document(
        source_path=str(TARGET_PATH),
        catalog=template_result.catalog,
        start_page=1,
        end_page=1,
        outline=target_outline,
        output_path=str(OUTPUT_PATH),
    )
    print("应用日志")
    for log in result.logs:
        print(log)
    print()

    inspect_output_document(template_result)
    print("调试验证通过。")


if __name__ == "__main__":
    main()
