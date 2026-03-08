from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

from models.format_profile import FormatCatalog, ImageFormat, ParagraphFormat, TableFormat
from services.word_com import WordComClient
from utils.numbering import (
    detect_heading_level,
    detect_heading_numbering_style,
    looks_like_figure_caption,
    looks_like_table_caption,
    resolve_numbering_pattern,
    split_leading_numbering,
    strip_existing_numbering,
)


@dataclass
class ExtractResult:
    catalog: FormatCatalog
    outline: List[dict]
    logs: List[str]


class FormatExtractor:
    def __init__(self, client: WordComClient) -> None:
        self.client = client

    def extract_template_document(self, document_path: str) -> ExtractResult:
        with self.client as client:
            document = client.open_document(document_path, read_only=True)
            try:
                page_count = client.get_page_count(document)
                result = self._build_catalog(document, document.Content, floating_pages=list(range(1, page_count + 1)))
                result.logs.insert(0, f"已按整份模板识别，共 {page_count} 页。")
                return result
            finally:
                client.close_document(document, save_changes=False)

    def extract_from_page_range(self, document_path: str, start_page: int, end_page: int) -> ExtractResult:
        with self.client as client:
            document = client.open_document(document_path, read_only=True)
            try:
                page_range = client.get_page_span_range(document, start_page, end_page)
                result = self._build_catalog(document, page_range, floating_pages=list(range(start_page, end_page + 1)))
                result.logs.insert(0, f"已按第 {start_page}-{end_page} 页识别模板范围。")
                return result
            finally:
                client.close_document(document, save_changes=False)

    def extract_from_template_page(self, document_path: str, page_number: int) -> Tuple[FormatCatalog, List[str]]:
        result = self.extract_from_page_range(document_path, page_number, page_number)
        return result.catalog, result.logs

    def extract_from_whole_document(self, document_path: str) -> Tuple[FormatCatalog, List[str]]:
        result = self.extract_template_document(document_path)
        return result.catalog, result.logs

    def extract_outline_from_page_range(self, document_path: str, start_page: int, end_page: int) -> Tuple[List[dict], List[str]]:
        with self.client as client:
            document = client.open_document(document_path, read_only=True)
            try:
                page_range = client.get_page_span_range(document, start_page, end_page)
                _buckets, _tables, _images, outline = self._collect_structure(
                    document,
                    page_range,
                    floating_pages=list(range(start_page, end_page + 1)),
                )
                logs = [f"已加载目标文档第 {start_page}-{end_page} 页检查项。"]
                if outline:
                    logs.append(f"识别目标文档检查项 {len(outline)} 条。")
                else:
                    logs.append("目标文档未识别到检查项。")
                return outline, logs
            finally:
                client.close_document(document, save_changes=False)

    def _build_catalog(self, document, target_range, floating_pages: Optional[List[int]]) -> ExtractResult:
        catalog = FormatCatalog()
        logs: List[str] = []
        buckets, tables, images, outline = self._collect_structure(document, target_range, floating_pages)
        visible_categories: List[str] = []

        if buckets["正文"]:
            catalog.body = self._pick_paragraph_format(buckets["正文"], category="正文")
            logs.append(f"识别正文样本 {len(buckets['正文'])} 段。")
            visible_categories.append("正文")
        for level in range(1, 10):
            key = f"标题{level}"
            if buckets[key]:
                catalog.headings[key] = self._pick_paragraph_format(buckets[key], heading_level=level, category=key)
                logs.append(f"识别 {key} 样本 {len(buckets[key])} 段。")
                visible_categories.append(key)
        if buckets["图注"]:
            catalog.figure_caption = self._pick_paragraph_format(buckets["图注"], category="图注")
            logs.append(f"识别图注样本 {len(buckets['图注'])} 段。")
            visible_categories.append("图注")
        if buckets["表注"]:
            catalog.table_caption = self._pick_paragraph_format(buckets["表注"], category="表注")
            logs.append(f"识别表注样本 {len(buckets['表注'])} 段。")
            visible_categories.append("表注")
        if images:
            catalog.image = self._pick_image_format(images)
            logs.append(f"识别图片样本 {len(images)} 个。")
            visible_categories.append("图片")
        if tables:
            catalog.table = self._pick_table_format(tables)
            logs.append(f"识别表格样本 {len(tables)} 个。")
            visible_categories.append("表格")
        if outline:
            logs.append(f"识别检查项目 {len(outline)} 条。")
        if not logs:
            logs.append("未识别到可用格式样本，已保留默认参数。")
        catalog.set_visible_categories(visible_categories)
        return ExtractResult(catalog=catalog, outline=outline, logs=logs)

    def _collect_structure(
        self,
        document,
        target_range,
        floating_pages: Optional[List[int]],
    ) -> Tuple[Dict[str, List[dict]], List[dict], List[dict], List[dict]]:
        paragraphs = [self.client.get_paragraph_info(p) for p in self.client.iter_paragraphs(target_range)]
        tables = [self.client.get_table_info(t) for t in self.client.iter_tables(target_range)]
        images = [self.client.get_inline_shape_info(s) for s in self.client.iter_inline_shapes(target_range)]
        outline: List[dict] = []

        for page_number in sorted(set(floating_pages or [])):
            images.extend(self.client.get_floating_shape_info(s) for s in self.client.iter_shapes_on_page(document, page_number))

        image_anchor_starts = {
            int(info["anchor_start"])
            for info in images
            if info.get("anchor_start") is not None
        }

        buckets: Dict[str, List[dict]] = defaultdict(list)
        previous_info: Optional[dict] = None
        for info in paragraphs:
            category = self._classify_paragraph(info, tables, image_anchor_starts, previous_info)
            if category:
                buckets[category].append(info)
                if category.startswith("标题") or category in {"图注", "表注"}:
                    outline.append(self._build_outline_item(info, category, len(outline) + 1))
            previous_info = info
        return buckets, tables, images, outline

    def _classify_paragraph(
        self,
        info: dict,
        tables: List[dict],
        image_anchor_starts: Optional[set[int]] = None,
        previous_info: Optional[dict] = None,
    ) -> Optional[str]:
        text = info["text"].strip()
        if not text:
            return None
        if self._inside_table(info, tables):
            return None

        style = info.get("style_name", "")
        list_string = info.get("list_string", "")
        heading_level = detect_heading_level(style, text, list_string)
        if heading_level:
            return f"标题{heading_level}"

        image_anchor_starts = image_anchor_starts or set()
        current_is_image_anchor = info.get("range_start") in image_anchor_starts
        previous_is_image_anchor = bool(previous_info and previous_info.get("range_start") in image_anchor_starts)
        previous_is_table = bool(previous_info and self._inside_table(previous_info, tables))

        if looks_like_figure_caption(text, style):
            if current_is_image_anchor or previous_is_image_anchor or len(text) <= 80:
                return "图注"
        if looks_like_table_caption(text, style):
            if previous_is_table or len(text) <= 80:
                return "表注"
        return "正文"

    @staticmethod
    def _build_outline_item(info: dict, category: str, order: int) -> dict:
        level = int(category.replace("标题", "")) if category.startswith("标题") else 0
        prefix, clean_text = split_leading_numbering(info.get("text", ""))
        range_start = info.get("range_start", 0)
        return {
            "order": order,
            "source_order": order,
            "outline_id": f"outline-{order}-{range_start}",
            "page_number": info.get("page_number", 0),
            "level": level,
            "category": category,
            "text": clean_text,
            "original_text": info.get("text", ""),
            "original_prefix": prefix,
            "range_start": range_start,
            "range_end": info.get("range_end", range_start),
            "space_after_numbering": bool(prefix and info.get("text", "").strip().startswith(f"{prefix} ")),
        }

    @staticmethod
    def _inside_table(info: dict, tables: List[dict]) -> bool:
        start = info["range_start"]
        end = info["range_end"]
        for table in tables:
            if start >= table["range_start"] and end <= table["range_end"]:
                return True
        return False

    @staticmethod
    def _pick_paragraph_format(samples: List[dict], heading_level: Optional[int] = None, category: Optional[str] = None) -> ParagraphFormat:
        base = samples[0]
        numbering_style = "none"
        pattern = ""
        renumber = False
        clear_existing_numbering = False
        if heading_level is not None:
            numbering_style = FormatExtractor._infer_heading_numbering_style(samples, heading_level)
            pattern = resolve_numbering_pattern(numbering_style)
            renumber = numbering_style != "none"
            clear_existing_numbering = numbering_style != "none"
        elif category == "图注":
            numbering_style = "figure_decimal"
            pattern = resolve_numbering_pattern(numbering_style)
        elif category == "表注":
            numbering_style = "table_decimal"
            pattern = resolve_numbering_pattern(numbering_style)
        return ParagraphFormat(
            font_name=FormatExtractor._most_common(samples, "font_name", fallback=base["font_name"]),
            font_size=FormatExtractor._most_common(samples, "font_size", fallback=base["font_size"]),
            bold=FormatExtractor._most_common(samples, "bold", fallback=base["bold"]),
            italic=FormatExtractor._most_common(samples, "italic", fallback=base["italic"]),
            alignment=FormatExtractor._most_common(samples, "alignment", fallback=base["alignment"]),
            left_indent=0.0 if category in {"图注", "表注"} else FormatExtractor._most_common(samples, "left_indent_chars", fallback=base["left_indent_chars"]),
            right_indent=0.0 if category in {"图注", "表注"} else FormatExtractor._most_common(samples, "right_indent_chars", fallback=base["right_indent_chars"]),
            first_line_indent=0.0 if category in {"图注", "表注"} else FormatExtractor._most_common(samples, "first_line_indent_chars", fallback=base["first_line_indent_chars"]),
            space_before=FormatExtractor._most_common(samples, "space_before", fallback=base["space_before"]),
            space_after=FormatExtractor._most_common(samples, "space_after", fallback=base["space_after"]),
            line_spacing_mode=FormatExtractor._most_common(samples, "line_spacing_mode", fallback=base["line_spacing_mode"]),
            line_spacing=FormatExtractor._most_common(samples, "line_spacing", fallback=base["line_spacing"]),
            numbering_style=numbering_style,
            custom_numbering_pattern="",
            numbering_pattern=pattern,
            renumber=renumber,
            clear_existing_numbering=clear_existing_numbering,
            continue_previous_numbering=False,
            space_after_numbering=FormatExtractor._infer_space_after_numbering(samples, renumber),
        )

    @staticmethod
    def _pick_image_format(samples: List[dict]) -> ImageFormat:
        base = samples[0]
        return ImageFormat(
            alignment=FormatExtractor._most_common(samples, "alignment", fallback=base["alignment"]),
            space_before=FormatExtractor._most_common(samples, "space_before", fallback=base["space_before"]),
            space_after=FormatExtractor._most_common(samples, "space_after", fallback=base["space_after"]),
            line_spacing_mode=FormatExtractor._most_common(samples, "line_spacing_mode", fallback=base.get("line_spacing_mode", "单倍")),
            line_spacing=FormatExtractor._most_common(samples, "line_spacing", fallback=base.get("line_spacing", 1.0)),
            wrap_type=FormatExtractor._most_common(samples, "wrap_type", fallback=base["wrap_type"]),
        )

    @staticmethod
    def _pick_table_format(samples: List[dict]) -> TableFormat:
        base = samples[0]
        top = FormatExtractor._most_common(samples, "top_border", fallback=base["top_border"])
        header_bottom = FormatExtractor._most_common(samples, "header_bottom_border", fallback=base["header_bottom_border"])
        bottom = FormatExtractor._most_common(samples, "bottom_border", fallback=base["bottom_border"])
        inside_horizontal = FormatExtractor._most_common(samples, "inside_horizontal_border", fallback=base["inside_horizontal_border"])
        inside_vertical = FormatExtractor._most_common(samples, "inside_vertical_border", fallback=base["inside_vertical_border"])
        return TableFormat(
            alignment=FormatExtractor._most_common(samples, "alignment", fallback=base["alignment"]),
            cell_vertical_alignment=FormatExtractor._most_common(samples, "cell_vertical_alignment", fallback=base["cell_vertical_alignment"]),
            top_border=top,
            header_bottom_border=header_bottom,
            bottom_border=bottom,
            inside_horizontal_border=inside_horizontal,
            inside_vertical_border=inside_vertical,
            three_line_status=FormatExtractor._infer_three_line(top, header_bottom, bottom, inside_vertical),
        )

    @staticmethod
    def _infer_space_after_numbering(samples: List[dict], renumber: bool) -> bool:
        if not renumber:
            return True
        values = [bool(sample.get("space_after_numbering", True)) for sample in samples]
        if not values:
            return True
        return Counter(values).most_common(1)[0][0]

    @staticmethod
    def _infer_three_line(top: bool, header_bottom: bool, bottom: bool, inside_vertical: bool) -> str:
        if top and header_bottom and bottom and not inside_vertical:
            return "是"
        if not top and not bottom:
            return "否"
        return "不确定"

    @staticmethod
    def _infer_heading_numbering_style(samples: List[dict], heading_level: int) -> str:
        detected: List[str] = []
        for sample in samples:
            style = detect_heading_numbering_style(sample.get("text", ""), sample.get("list_string", ""), heading_level)
            if style != "none":
                detected.append(style)
        if detected:
            return Counter(detected).most_common(1)[0][0]
        if heading_level == 1:
            return "chinese_comma"
        if heading_level == 2:
            return "paren_chinese"
        if heading_level == 3:
            return "decimal_comma"
        return "none"

    @staticmethod
    def _most_common(samples: Iterable[dict], key: str, fallback):
        values = [sample.get(key) for sample in samples]
        counter = Counter(values)
        if not counter:
            return fallback
        return counter.most_common(1)[0][0]
