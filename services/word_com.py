from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List

import pythoncom
import win32com.client as win32

from models.format_profile import round_float


WD_GO_TO_PAGE = 1
WD_GO_TO_ABSOLUTE = 1
WD_ACTIVE_END_PAGE_NUMBER = 3
WD_ALIGN_PARAGRAPH_LEFT = 0
WD_ALIGN_PARAGRAPH_CENTER = 1
WD_ALIGN_PARAGRAPH_RIGHT = 2
WD_ALIGN_PARAGRAPH_JUSTIFY = 3
WD_LINE_SPACE_SINGLE = 0
WD_LINE_SPACE_ONE_POINT_FIVE = 1
WD_LINE_SPACE_DOUBLE = 2
WD_LINE_SPACE_AT_LEAST = 3
WD_LINE_SPACE_EXACTLY = 4
WD_LINE_SPACE_MULTIPLE = 5
WD_TABLE_ALIGNMENT_LEFT = 0
WD_TABLE_ALIGNMENT_CENTER = 1
WD_TABLE_ALIGNMENT_RIGHT = 2
WD_CELL_ALIGN_VERTICAL_TOP = 0
WD_CELL_ALIGN_VERTICAL_CENTER = 1
WD_CELL_ALIGN_VERTICAL_BOTTOM = 3
WD_WRAP_INLINE = 7
WD_WRAP_SQUARE = 0
WD_WRAP_TIGHT = 1
WD_WRAP_TOP_BOTTOM = 4
WD_WRAP_FRONT = 3
WD_WRAP_BEHIND = 5
WD_BORDER_TOP = 1
WD_BORDER_BOTTOM = 3
WD_BORDER_HORIZONTAL = 5
WD_BORDER_VERTICAL = 6
WD_BORDER_NONE = 0
WD_OUTLINE_LEVEL_BODY_TEXT = 10

STYLE_NAME_TO_HEADING_LEVEL = {
    "标题 1": 1,
    "标题1": 1,
    "heading 1": 1,
    "heading1": 1,
    "标题 2": 2,
    "标题2": 2,
    "heading 2": 2,
    "heading2": 2,
    "标题 3": 3,
    "标题3": 3,
    "heading 3": 3,
    "heading3": 3,
    "标题 4": 4,
    "标题4": 4,
    "heading 4": 4,
    "heading4": 4,
    "标题 5": 5,
    "标题5": 5,
    "heading 5": 5,
    "heading5": 5,
    "标题 6": 6,
    "标题6": 6,
    "heading 6": 6,
    "heading6": 6,
    "标题 7": 7,
    "标题7": 7,
    "heading 7": 7,
    "heading7": 7,
    "标题 8": 8,
    "标题8": 8,
    "heading 8": 8,
    "heading8": 8,
    "标题 9": 9,
    "标题9": 9,
    "heading 9": 9,
    "heading9": 9,
}

ALIGNMENT_MAP = {
    WD_ALIGN_PARAGRAPH_LEFT: "左对齐",
    WD_ALIGN_PARAGRAPH_CENTER: "居中",
    WD_ALIGN_PARAGRAPH_RIGHT: "右对齐",
    WD_ALIGN_PARAGRAPH_JUSTIFY: "两端对齐",
}
ALIGNMENT_REVERSE_MAP = {value: key for key, value in ALIGNMENT_MAP.items()}
TABLE_ALIGNMENT_MAP = {
    WD_TABLE_ALIGNMENT_LEFT: "左对齐",
    WD_TABLE_ALIGNMENT_CENTER: "居中",
    WD_TABLE_ALIGNMENT_RIGHT: "右对齐",
}
TABLE_ALIGNMENT_REVERSE_MAP = {value: key for key, value in TABLE_ALIGNMENT_MAP.items()}
VERTICAL_ALIGNMENT_MAP = {
    WD_CELL_ALIGN_VERTICAL_TOP: "顶端",
    WD_CELL_ALIGN_VERTICAL_CENTER: "居中",
    WD_CELL_ALIGN_VERTICAL_BOTTOM: "底端",
}
VERTICAL_ALIGNMENT_REVERSE_MAP = {value: key for key, value in VERTICAL_ALIGNMENT_MAP.items()}
WRAP_TYPE_MAP = {
    WD_WRAP_INLINE: "嵌入型",
    WD_WRAP_SQUARE: "四周型",
    WD_WRAP_TIGHT: "紧密型",
    WD_WRAP_TOP_BOTTOM: "上下型",
    WD_WRAP_FRONT: "浮于文字上方",
    WD_WRAP_BEHIND: "衬于文字下方",
}
WRAP_TYPE_REVERSE_MAP = {value: key for key, value in WRAP_TYPE_MAP.items()}
LINE_SPACING_MAP = {
    WD_LINE_SPACE_SINGLE: "单倍",
    WD_LINE_SPACE_ONE_POINT_FIVE: "1.5 倍",
    WD_LINE_SPACE_DOUBLE: "2 倍",
    WD_LINE_SPACE_AT_LEAST: "最小值",
    WD_LINE_SPACE_EXACTLY: "固定值",
    WD_LINE_SPACE_MULTIPLE: "多倍",
}


class WordAutomationError(RuntimeError):
    pass


class WordComClient:
    def __init__(self, visible: bool = False) -> None:
        self.visible = visible
        self.app = None
        self._initialized = False

    def __enter__(self) -> "WordComClient":
        self.start()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def start(self) -> None:
        if self.app is not None:
            return
        pythoncom.CoInitialize()
        self._initialized = True
        try:
            self.app = win32.DispatchEx("Word.Application")
            self.app.Visible = self.visible
            self.app.DisplayAlerts = 0
        except Exception as exc:
            self._release_com()
            raise WordAutomationError("无法启动 Microsoft Word，请确认已正确安装桌面版 Word。") from exc

    def close(self) -> None:
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
            self.app = None
        self._release_com()

    def _release_com(self) -> None:
        if self._initialized:
            pythoncom.CoUninitialize()
            self._initialized = False

    def open_document(self, path: str, read_only: bool = True):
        self.start()
        normalized_path = str(Path(path).expanduser().resolve(strict=False))
        try:
            return self.app.Documents.Open(normalized_path, ReadOnly=read_only, AddToRecentFiles=False)
        except Exception as exc:
            raise WordAutomationError(f"无法打开文档：{normalized_path}") from exc

    def close_document(self, document, save_changes: bool = False) -> None:
        if document is None:
            return
        try:
            document.Close(SaveChanges=save_changes)
        except Exception:
            pass

    def get_page_count(self, document) -> int:
        candidates: List[int] = []
        try:
            document.Repaginate()
        except Exception:
            pass

        def collect(value) -> None:
            try:
                numeric = int(value)
            except Exception:
                return
            if numeric > 0:
                candidates.append(numeric)

        try:
            collect(document.ComputeStatistics(2))
        except Exception:
            pass

        try:
            collect(document.Content.Information(4))
        except Exception:
            pass

        try:
            content = document.Content
            last_pos = max(int(content.Start), int(content.End) - 1)
            collect(document.Range(Start=last_pos, End=last_pos).Information(WD_ACTIVE_END_PAGE_NUMBER))
        except Exception:
            pass

        try:
            collect(document.BuiltInDocumentProperties("Number of Pages").Value)
        except Exception:
            pass

        if candidates:
            return max(candidates)
        raise WordAutomationError("无法计算文档页数。")

    def get_page_range(self, document, page_number: int):
        page_count = self.get_page_count(document)
        if page_number < 1 or page_number > page_count:
            raise WordAutomationError(f"页码超出范围：{page_number}，总页数为 {page_count}。")
        start_range = document.GoTo(What=WD_GO_TO_PAGE, Which=WD_GO_TO_ABSOLUTE, Count=page_number)
        start = start_range.Start
        if page_number < page_count:
            next_range = document.GoTo(What=WD_GO_TO_PAGE, Which=WD_GO_TO_ABSOLUTE, Count=page_number + 1)
            end = next_range.Start - 1
        else:
            end = document.Content.End
        return document.Range(Start=start, End=end)

    def get_page_span_range(self, document, start_page: int, end_page: int):
        page_count = self.get_page_count(document)
        if start_page < 1 or end_page < start_page or end_page > page_count:
            raise WordAutomationError(f"页码范围无效：{start_page}-{end_page}，总页数为 {page_count}。")
        start_range = document.GoTo(What=WD_GO_TO_PAGE, Which=WD_GO_TO_ABSOLUTE, Count=start_page)
        start = start_range.Start
        if end_page < page_count:
            next_range = document.GoTo(What=WD_GO_TO_PAGE, Which=WD_GO_TO_ABSOLUTE, Count=end_page + 1)
            end = next_range.Start - 1
        else:
            end = document.Content.End
        return document.Range(Start=start, End=end)

    def iter_paragraphs(self, word_range) -> List[Any]:
        return [paragraph for paragraph in word_range.Paragraphs]

    def iter_tables(self, word_range) -> List[Any]:
        return [table for table in word_range.Tables]

    def iter_inline_shapes(self, word_range) -> List[Any]:
        return [shape for shape in word_range.InlineShapes]

    def iter_shapes_on_page(self, document, page_number: int) -> List[Any]:
        shapes = []
        for shape in document.Shapes:
            try:
                if int(shape.Anchor.Information(WD_ACTIVE_END_PAGE_NUMBER)) == page_number:
                    shapes.append(shape)
            except Exception:
                continue
        return shapes

    def get_paragraph_info(self, paragraph) -> Dict[str, Any]:
        para_range = paragraph.Range
        style_name = self._get_style_name(paragraph)
        font = para_range.Font
        fmt = paragraph.Format
        line_spacing_mode = LINE_SPACING_MAP.get(int(fmt.LineSpacingRule), "多倍")
        line_spacing_value = round_float(fmt.LineSpacing, 2)
        if int(fmt.LineSpacingRule) == WD_LINE_SPACE_ONE_POINT_FIVE:
            line_spacing_value = 1.5
        elif int(fmt.LineSpacingRule) == WD_LINE_SPACE_DOUBLE:
            line_spacing_value = 2.0
        elif int(fmt.LineSpacingRule) == WD_LINE_SPACE_SINGLE:
            line_spacing_value = 1.0
        elif int(fmt.LineSpacingRule) == WD_LINE_SPACE_MULTIPLE:
            line_spacing_value = round_float(float(fmt.LineSpacing) / 12.0, 2)
        page_number = 0
        try:
            page_number = int(para_range.Information(WD_ACTIVE_END_PAGE_NUMBER))
        except Exception:
            page_number = 0
        left_indent_chars = self._safe_character_indent(fmt, "CharacterUnitLeftIndent")
        right_indent_chars = self._safe_character_indent(fmt, "CharacterUnitRightIndent")
        first_line_indent_chars = self._safe_character_indent(fmt, "CharacterUnitFirstLineIndent")
        return {
            "text": para_range.Text.strip("\r\x07"),
            "style_name": style_name,
            "font_name": getattr(font, "NameFarEast", "") or getattr(font, "Name", "") or "宋体",
            "font_size": round_float(getattr(font, "Size", 12.0), 2),
            "bold": bool(getattr(font, "Bold", False)),
            "italic": bool(getattr(font, "Italic", False)),
            "alignment": ALIGNMENT_MAP.get(int(fmt.Alignment), "左对齐"),
            "left_indent": left_indent_chars,
            "right_indent": right_indent_chars,
            "first_line_indent": first_line_indent_chars,
            "left_indent_chars": left_indent_chars,
            "right_indent_chars": right_indent_chars,
            "first_line_indent_chars": first_line_indent_chars,
            "space_before": round_float(fmt.SpaceBefore),
            "space_after": round_float(fmt.SpaceAfter),
            "line_spacing_mode": line_spacing_mode,
            "line_spacing": line_spacing_value,
            "list_string": self._safe_list_string(para_range),
            "range_start": int(para_range.Start),
            "range_end": int(para_range.End),
            "page_number": page_number,
        }

    def get_inline_shape_info(self, shape) -> Dict[str, Any]:
        paragraph = shape.Range.Paragraphs(1)
        info = self.get_paragraph_info(paragraph)
        return {
            "alignment": info["alignment"],
            "space_before": info["space_before"],
            "space_after": info["space_after"],
            "line_spacing_mode": info["line_spacing_mode"],
            "line_spacing": info["line_spacing"],
            "wrap_type": "嵌入型",
            "anchor_start": int(shape.Range.Start),
        }

    def get_floating_shape_info(self, shape) -> Dict[str, Any]:
        paragraph = shape.Anchor.Paragraphs(1)
        info = self.get_paragraph_info(paragraph)
        wrap_type = WRAP_TYPE_MAP.get(int(shape.WrapFormat.Type), "四周型")
        return {
            "alignment": info["alignment"],
            "space_before": info["space_before"],
            "space_after": info["space_after"],
            "line_spacing_mode": info["line_spacing_mode"],
            "line_spacing": info["line_spacing"],
            "wrap_type": wrap_type,
            "anchor_start": int(shape.Anchor.Start),
        }

    def get_table_info(self, table) -> Dict[str, Any]:
        cell_vertical_alignment = "居中"
        try:
            cell_vertical_alignment = VERTICAL_ALIGNMENT_MAP.get(int(table.Rows(1).Cells(1).VerticalAlignment), "居中")
        except Exception:
            pass

        return {
            "alignment": TABLE_ALIGNMENT_MAP.get(int(table.Rows.Alignment), "居中"),
            "cell_vertical_alignment": cell_vertical_alignment,
            "top_border": self._border_enabled(table.Borders(WD_BORDER_TOP)),
            "header_bottom_border": self._header_bottom_border(table),
            "bottom_border": self._border_enabled(table.Borders(WD_BORDER_BOTTOM)),
            "inside_horizontal_border": self._border_enabled(table.Borders(WD_BORDER_HORIZONTAL)),
            "inside_vertical_border": self._border_enabled(table.Borders(WD_BORDER_VERTICAL)),
            "range_start": int(table.Range.Start),
            "range_end": int(table.Range.End),
            "table_obj": table,
        }

    def duplicate_document_editable(self, input_path: str, output_path: str):
        document = self.open_document(input_path, read_only=False)
        try:
            document.SaveAs2(output_path)
            document.Close(SaveChanges=False)
        except Exception as exc:
            self.close_document(document, save_changes=False)
            raise WordAutomationError(f"无法另存为文档：{output_path}") from exc
        return self.open_document(output_path, read_only=False)

    def create_temp_document(self):
        self.start()
        try:
            return self.app.Documents.Add()
        except Exception as exc:
            raise WordAutomationError("无法创建临时文档。") from exc

    def reorder_ranges(self, document, range_specs: List[Dict[str, Any]], ordered_range_ids: List[str], insert_at: int) -> None:
        specs = self._normalize_range_specs(range_specs)
        if len(specs) <= 1:
            return
        current_order = [str(spec["range_id"]) for spec in specs]
        desired_order = [str(range_id) for range_id in ordered_range_ids]
        if current_order == desired_order:
            return
        spec_by_id = {str(spec["range_id"]): spec for spec in specs}
        if len(spec_by_id) != len(specs):
            raise WordAutomationError("文档内容块标识重复，无法重排。")
        ordered_specs: List[Dict[str, Any]] = []
        seen: set[str] = set()
        for range_id in desired_order:
            spec = spec_by_id.get(range_id)
            if spec is None or range_id in seen:
                raise WordAutomationError("检查项顺序与文档内容块不一致，无法重排。")
            ordered_specs.append(spec)
            seen.add(range_id)
        if len(ordered_specs) != len(specs):
            raise WordAutomationError("检查项数量与文档内容块不一致，无法重排。")

        temp_document = self.create_temp_document()
        try:
            captured_ranges = self._capture_ranges_to_temp_document(document, temp_document, specs)
            self._delete_ranges(document, specs)
            insert_position = self._clamp_insert_position(document, insert_at)
            for spec in ordered_specs:
                stored = captured_ranges[str(spec["range_id"])]
                insert_position = self._insert_range_from_document(document, temp_document, stored["start"], stored["end"], insert_position)
        except WordAutomationError:
            raise
        except Exception as exc:
            raise WordAutomationError("重排文档内容失败。") from exc
        finally:
            self.close_document(temp_document, save_changes=False)

    def save_document(self, document) -> None:
        try:
            document.Save()
        except Exception as exc:
            raise WordAutomationError("保存文档失败。") from exc

    def apply_paragraph_format(self, paragraph, profile) -> None:
        para_range = paragraph.Range
        fmt = paragraph.Format
        heading_level = self._safe_outline_level(paragraph)
        if not heading_level:
            heading_level = self._heading_level_from_style_name(self._get_style_name(paragraph))
        para_range.Font.NameFarEast = profile.font_name
        para_range.Font.Name = profile.font_name
        para_range.Font.Size = profile.font_size
        para_range.Font.Bold = -1 if profile.bold else 0
        para_range.Font.Italic = -1 if profile.italic else 0
        fmt.Alignment = ALIGNMENT_REVERSE_MAP.get(profile.alignment, WD_ALIGN_PARAGRAPH_LEFT)
        self._set_character_indent(fmt, "CharacterUnitLeftIndent", profile.left_indent)
        self._set_character_indent(fmt, "CharacterUnitRightIndent", profile.right_indent)
        self._set_character_indent(fmt, "CharacterUnitFirstLineIndent", profile.first_line_indent)
        fmt.SpaceBefore = profile.space_before
        fmt.SpaceAfter = profile.space_after
        rule, value = self._line_spacing_to_word(profile.line_spacing_mode, profile.line_spacing)
        fmt.LineSpacingRule = rule
        fmt.LineSpacing = value
        if heading_level:
            try:
                paragraph.OutlineLevel = heading_level
            except Exception:
                pass
            try:
                paragraph.Range.ParagraphFormat.OutlineLevel = heading_level
            except Exception:
                pass

    def apply_image_format(self, paragraph, profile) -> None:
        fmt = paragraph.Format
        fmt.Alignment = ALIGNMENT_REVERSE_MAP.get(profile.alignment, WD_ALIGN_PARAGRAPH_CENTER)
        fmt.SpaceBefore = profile.space_before
        fmt.SpaceAfter = profile.space_after
        rule, value = self._line_spacing_to_word(profile.line_spacing_mode, profile.line_spacing)
        fmt.LineSpacingRule = rule
        fmt.LineSpacing = value
        self._set_character_indent(fmt, "CharacterUnitLeftIndent", 0.0)
        self._set_character_indent(fmt, "CharacterUnitRightIndent", 0.0)
        self._set_character_indent(fmt, "CharacterUnitFirstLineIndent", 0.0)
        try:
            fmt.LeftIndent = 0
            fmt.RightIndent = 0
            fmt.FirstLineIndent = 0
        except Exception:
            pass

        target_shape = None
        try:
            anchor_range = paragraph.Range
            inline_shapes = [shape for shape in anchor_range.InlineShapes]
            if inline_shapes:
                target_shape = inline_shapes[0]
        except Exception:
            target_shape = None

        if target_shape is not None:
            return

        try:
            document = paragraph.Range.Document
            anchor_start = int(paragraph.Range.Start)
            for shape in document.Shapes:
                try:
                    if int(shape.Anchor.Start) == anchor_start:
                        target_shape = shape
                        break
                except Exception:
                    continue
        except Exception:
            target_shape = None

        if target_shape is None:
            return

        try:
            target_shape.WrapFormat.Type = WRAP_TYPE_REVERSE_MAP.get(profile.wrap_type, WD_WRAP_SQUARE)
        except Exception:
            pass

    def apply_table_format(self, table, profile) -> None:
        table.Rows.Alignment = TABLE_ALIGNMENT_REVERSE_MAP.get(profile.alignment, WD_TABLE_ALIGNMENT_CENTER)
        self._set_border(table.Borders(WD_BORDER_TOP), profile.top_border)
        self._set_border(table.Borders(WD_BORDER_BOTTOM), profile.bottom_border)
        self._set_border(table.Borders(WD_BORDER_HORIZONTAL), profile.inside_horizontal_border)
        self._set_border(table.Borders(WD_BORDER_VERTICAL), profile.inside_vertical_border)
        if profile.header_bottom_border and table.Rows.Count >= 1:
            try:
                self._set_border(table.Rows(1).Borders(WD_BORDER_BOTTOM), True)
            except Exception:
                pass
        try:
            for paragraph in table.Range.Paragraphs:
                self._set_character_indent(paragraph.Format, "CharacterUnitFirstLineIndent", 0.0)
        except Exception:
            pass
        try:
            for row in table.Rows:
                for cell in row.Cells:
                    cell.VerticalAlignment = VERTICAL_ALIGNMENT_REVERSE_MAP.get(profile.cell_vertical_alignment, WD_CELL_ALIGN_VERTICAL_CENTER)
        except Exception:
            pass

    @staticmethod
    def _normalize_style_name(style_name: str) -> str:
        return re.sub(r"\s+", " ", (style_name or "").strip()).lower()

    @classmethod
    def _heading_level_from_style_name(cls, style_name: str) -> int:
        return STYLE_NAME_TO_HEADING_LEVEL.get(cls._normalize_style_name(style_name), 0)

    @classmethod
    def _capture_heading_style(cls, paragraph):
        style_obj = None
        try:
            style_obj = paragraph.Style
        except Exception:
            try:
                style_obj = paragraph.Range.Style
            except Exception:
                style_obj = None
        style_name = cls._style_name(style_obj) or cls._get_style_name(paragraph)
        heading_level = cls._safe_outline_level(paragraph)
        if not heading_level:
            heading_level = cls._heading_level_from_style_name(style_name)
        if not heading_level:
            return None
        return style_obj, style_name, heading_level

    @staticmethod
    def _style_name(style) -> str:
        if style is None:
            return ""
        try:
            return style.NameLocal
        except Exception:
            try:
                return str(style)
            except Exception:
                return ""

    @classmethod
    def _preserve_heading_style(cls, paragraph) -> None:
        heading_level = cls._safe_outline_level(paragraph)
        if not heading_level:
            style_name = cls._get_style_name(paragraph)
            heading_level = cls._heading_level_from_style_name(style_name)
        if not heading_level:
            return
        try:
            paragraph.OutlineLevel = heading_level
        except Exception:
            pass

    @classmethod
    def _restore_heading_style(cls, paragraph, style_obj, style_name: str, heading_level: int) -> None:
        if style_obj is not None:
            try:
                paragraph.Style = style_obj
            except Exception:
                pass
            try:
                paragraph.Range.Style = style_obj
            except Exception:
                pass
        if style_name:
            try:
                paragraph.Style = style_name
            except Exception:
                pass
            try:
                paragraph.Range.Style = style_name
            except Exception:
                pass
        try:
            paragraph.OutlineLevel = heading_level
        except Exception:
            pass

    @staticmethod
    def _safe_outline_level(paragraph) -> int:
        try:
            level = int(paragraph.OutlineLevel)
            if 1 <= level < WD_OUTLINE_LEVEL_BODY_TEXT:
                return level
        except Exception:
            pass
        try:
            level = int(paragraph.Range.ParagraphFormat.OutlineLevel)
            if 1 <= level < WD_OUTLINE_LEVEL_BODY_TEXT:
                return level
        except Exception:
            pass
        return 0

    @staticmethod
    def _safe_character_indent(fmt, attr_name: str) -> float:
        try:
            return round_float(getattr(fmt, attr_name), 2)
        except Exception:
            return 0.0

    @staticmethod
    def _set_character_indent(fmt, attr_name: str, value: float) -> None:
        try:
            setattr(fmt, attr_name, round_float(value, 2))
        except Exception:
            fallback_map = {
                "CharacterUnitLeftIndent": "LeftIndent",
                "CharacterUnitRightIndent": "RightIndent",
                "CharacterUnitFirstLineIndent": "FirstLineIndent",
            }
            fallback_attr = fallback_map.get(attr_name)
            if not fallback_attr:
                return
            try:
                setattr(fmt, fallback_attr, round_float(value, 2) * 12)
            except Exception:
                pass

    @staticmethod
    def _get_style_name(paragraph) -> str:
        try:
            style = paragraph.Style
        except Exception:
            try:
                style = paragraph.Range.Style
            except Exception:
                return ""
        try:
            return style.NameLocal
        except Exception:
            try:
                return str(style)
            except Exception:
                return ""

    @staticmethod
    def _header_bottom_border(table) -> bool:
        try:
            if table.Rows.Count >= 1:
                return WordComClient._border_enabled(table.Rows(1).Borders(WD_BORDER_BOTTOM))
        except Exception:
            return False
        return False

    @staticmethod
    def _safe_list_string(word_range) -> str:
        try:
            return word_range.ListFormat.ListString or ""
        except Exception:
            return ""

    @staticmethod
    def _border_enabled(border) -> bool:
        try:
            return int(border.LineStyle) != WD_BORDER_NONE
        except Exception:
            return False

    @staticmethod
    def _set_border(border, enabled: bool) -> None:
        try:
            border.LineStyle = 1 if enabled else WD_BORDER_NONE
        except Exception:
            pass

    @staticmethod
    def _line_spacing_to_word(mode: str, value: float):
        if mode == "固定值":
            return WD_LINE_SPACE_EXACTLY, max(value, 1)
        if mode == "最小值":
            return WD_LINE_SPACE_AT_LEAST, max(value, 1)
        if mode == "单倍":
            return WD_LINE_SPACE_SINGLE, 12
        if mode == "1.5 倍":
            return WD_LINE_SPACE_ONE_POINT_FIVE, 18
        if mode == "2 倍":
            return WD_LINE_SPACE_DOUBLE, 24
        return WD_LINE_SPACE_MULTIPLE, max(value * 12, 12)

    @staticmethod
    def _clamp_insert_position(document, insert_at: int) -> int:
        content = document.Content
        start = int(content.Start)
        end = int(content.End)
        return max(start, min(int(insert_at), end))

    @staticmethod
    def _normalize_range_specs(range_specs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        normalized: List[Dict[str, Any]] = []
        for spec in range_specs:
            try:
                range_id = str(spec["range_id"])
                start = int(spec["start"])
                end = int(spec["end"])
            except Exception as exc:
                raise WordAutomationError("文档内容块范围无效。") from exc
            if end <= start:
                continue
            normalized.append({"range_id": range_id, "start": start, "end": end})
        normalized.sort(key=lambda item: (item["start"], item["end"]))
        for index in range(1, len(normalized)):
            previous = normalized[index - 1]
            current = normalized[index]
            if current["start"] < previous["end"]:
                raise WordAutomationError("文档内容块存在重叠，无法重排。")
        return normalized

    def _capture_ranges_to_temp_document(self, source_document, temp_document, specs: List[Dict[str, Any]]) -> Dict[str, Dict[str, int]]:
        captured: Dict[str, Dict[str, int]] = {}
        for spec in specs:
            target_range = temp_document.Range(Start=int(temp_document.Content.End) - 1, End=int(temp_document.Content.End) - 1)
            source_range = source_document.Range(Start=spec["start"], End=spec["end"])
            target_range.FormattedText = source_range.FormattedText
            inserted = temp_document.Range(Start=target_range.Start, End=target_range.End)
            captured[str(spec["range_id"])] = {"start": int(inserted.Start), "end": int(inserted.End)}
        return captured

    @staticmethod
    def _delete_ranges(document, specs: List[Dict[str, Any]]) -> None:
        for spec in reversed(specs):
            document.Range(Start=spec["start"], End=spec["end"]).Delete()

    def _insert_range_from_document(self, target_document, source_document, source_start: int, source_end: int, insert_at: int) -> int:
        insert_position = self._clamp_insert_position(target_document, insert_at)
        target_range = target_document.Range(Start=insert_position, End=insert_position)
        source_range = source_document.Range(Start=source_start, End=source_end)
        target_range.FormattedText = source_range.FormattedText
        inserted = target_document.Range(Start=target_range.Start, End=target_range.End)
        return int(inserted.End)
