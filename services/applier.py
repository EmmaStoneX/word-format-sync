from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

from models.format_profile import FormatCatalog
from services.word_com import WordComClient
from utils.numbering import (
    NumberingFormatter,
    detect_heading_level,
    detect_heading_numbering_style,
    detect_numbering_values,
    looks_like_figure_caption,
    looks_like_table_caption,
    resolve_numbering_pattern,
    strip_existing_numbering,
)


@dataclass
class ApplyResult:
    output_path: str
    logs: List[str]


class FormatApplier:
    def __init__(self, client: WordComClient) -> None:
        self.client = client

    def apply_to_document(
        self,
        source_path: str,
        catalog: FormatCatalog,
        start_page: int,
        end_page: int,
        outline: Optional[List[dict]] = None,
        output_path: Optional[str] = None,
    ) -> ApplyResult:
        output_path = output_path or self._build_output_path(source_path)
        logs: List[str] = []
        with self.client as client:
            document = client.duplicate_document_editable(source_path, output_path)
            try:
                logs.extend(self._reorder_document_by_outline(document, start_page, end_page, outline))
                target_range = client.get_page_span_range(document, start_page, end_page)
                initial_tables = client.iter_tables(target_range)
                initial_inline_shapes = client.iter_inline_shapes(target_range)
                initial_floating_shapes_by_page = {
                    page: client.iter_shapes_on_page(document, page)
                    for page in range(start_page, end_page + 1)
                }
                initial_floating_shapes = [
                    shape
                    for shapes in initial_floating_shapes_by_page.values()
                    for shape in shapes
                ]
                table_infos = [client.get_table_info(table) for table in initial_tables]
                image_anchor_starts = self._collect_image_anchor_starts(initial_inline_shapes, initial_floating_shapes_by_page)
                paragraphs = [
                    (paragraph, client.get_paragraph_info(paragraph))
                    for paragraph in client.iter_paragraphs(target_range)
                ]
                formatter = NumberingFormatter()
                figure_formatter = NumberingFormatter()
                table_formatter = NumberingFormatter()
                self._initialize_heading_counters(document, formatter, catalog, start_page)

                body_count = 0
                heading_count = 0
                figure_caption_count = 0
                table_caption_count = 0
                table_count = 0
                image_count = 0
                heading_index = 0
                started_levels: set[int] = set()
                previous_info = None

                for paragraph, info in paragraphs:
                    category = self._classify_target_paragraph(info, table_infos, image_anchor_starts, previous_info)
                    outline_item = outline[heading_index] if outline and heading_index < len(outline) else None
                    if category == "正文":
                        client.apply_paragraph_format(paragraph, catalog.body)
                        body_count += 1
                    elif category and category.startswith("标题"):
                        effective_category = self._resolve_outline_category(category, outline_item)
                        profile = catalog.headings.get(effective_category)
                        if profile is not None:
                            self._apply_heading(paragraph, effective_category, profile, formatter, started_levels, outline_item)
                            client.apply_paragraph_format(paragraph, profile)
                            heading_count += 1
                        heading_index += 1
                    elif category == "图注":
                        self._apply_caption(paragraph, catalog.figure_caption, figure_formatter, outline_item)
                        client.apply_paragraph_format(paragraph, catalog.figure_caption)
                        figure_caption_count += 1
                        heading_index += 1
                    elif category == "表注":
                        self._apply_caption(paragraph, catalog.table_caption, table_formatter, outline_item)
                        client.apply_paragraph_format(paragraph, catalog.table_caption)
                        table_caption_count += 1
                        heading_index += 1
                    previous_info = info

                for table in initial_tables:
                    client.apply_table_format(table, catalog.table)
                    table_count += 1

                for shape in initial_inline_shapes:
                    paragraph = shape.Range.Paragraphs(1)
                    client.apply_image_format(paragraph, catalog.image)
                    image_count += 1

                for shape in initial_floating_shapes:
                    paragraph = shape.Anchor.Paragraphs(1)
                    client.apply_image_format(paragraph, catalog.image)
                    image_count += 1

                client.save_document(document)
                logs.extend(
                    [
                        f"已将格式应用到第 {start_page}-{end_page} 页，并输出到：{output_path}",
                        f"正文：{body_count} 段",
                        f"标题：{heading_count} 段",
                        f"图注：{figure_caption_count} 段",
                        f"表注：{table_caption_count} 段",
                        f"表格：{table_count} 个",
                        f"图片：{image_count} 个",
                    ]
                )
                return ApplyResult(output_path=output_path, logs=logs)
            finally:
                client.close_document(document, save_changes=False)

    def _initialize_heading_counters(self, document, formatter: NumberingFormatter, catalog: FormatCatalog, start_page: int) -> None:
        formatter.reset()
        if start_page <= 1:
            return
        previous_range = self.client.get_page_span_range(document, 1, start_page - 1)
        table_infos = [self.client.get_table_info(table) for table in self.client.iter_tables(previous_range)]
        for paragraph in self.client.iter_paragraphs(previous_range):
            info = self.client.get_paragraph_info(paragraph)
            category = self._classify_target_paragraph(info, table_infos)
            if not category or not category.startswith("标题"):
                continue
            profile = catalog.headings.get(category)
            if profile is None or not profile.continue_previous_numbering:
                continue
            level = int(category.replace("标题", ""))
            pattern = self._resolve_profile_pattern(profile)
            if not pattern:
                continue
            formatter.next_for_level(level, pattern)

    def _reorder_document_by_outline(self, document, start_page: int, end_page: int, outline: Optional[List[dict]]) -> List[str]:
        if not outline:
            return []
        target_range = self.client.get_page_span_range(document, start_page, end_page)
        range_specs = self._build_outline_range_specs(document, target_range, outline)
        if not range_specs:
            return ["检查项与文档内容块未能一一对应，已跳过内容重排。"]
        if len(range_specs) <= 1:
            return []
        current_ids = [str(spec["range_id"]) for spec in range_specs]
        current_id_set = set(current_ids)
        ordered_ids = [str(item.get("outline_id", "")) for item in outline if item.get("outline_id") in current_id_set]
        if len(ordered_ids) != len(current_ids):
            return ["检查项与文档内容块数量不一致，已跳过内容重排。"]
        if current_ids == ordered_ids:
            return []
        insert_at = min(int(spec["start"]) for spec in range_specs)
        self.client.reorder_ranges(document, range_specs, ordered_ids, insert_at)
        return [f"已按检查项顺序重排 {len(range_specs)} 个内容块。"]

    def _build_outline_range_specs(self, document, target_range, outline: List[dict]) -> List[dict]:
        paragraphs = [
            (paragraph, self.client.get_paragraph_info(paragraph))
            for paragraph in self.client.iter_paragraphs(target_range)
        ]
        if not paragraphs:
            return []
        source_outline = self._sort_outline_by_source_order(outline)
        if len(source_outline) != len(outline):
            return []
        target_start = int(target_range.Start)
        target_end = int(target_range.End)
        paragraph_infos = [info for _, info in paragraphs]
        table_infos = [self.client.get_table_info(table) for table in self.client.iter_tables(target_range)]
        page_numbers = [int(info.get("page_number", 0) or 0) for info in paragraph_infos if int(info.get("page_number", 0) or 0) > 0]
        inline_shapes = self.client.iter_inline_shapes(target_range)
        floating_shapes_by_page = {
            page: self.client.iter_shapes_on_page(document, page)
            for page in range((min(page_numbers) if page_numbers else 1), (max(page_numbers) if page_numbers else 0) + 1)
        }
        image_anchor_starts = self._collect_image_anchor_starts(inline_shapes, floating_shapes_by_page)

        current_anchors: List[dict] = []
        previous_info = None
        for _index, (_paragraph, info) in enumerate(paragraphs):
            category = self._classify_target_paragraph(info, table_infos, image_anchor_starts, previous_info)
            if category and (category.startswith("标题") or category in {"图注", "表注"}):
                adjusted_start = int(info["range_start"])
                if category == "图注":
                    adjusted_start = self._figure_block_start(info, previous_info, image_anchor_starts)
                elif category == "表注":
                    adjusted_start = self._table_caption_block_start(info, previous_info, table_infos)
                current_anchors.append(
                    {
                        "category": category,
                        "start": adjusted_start,
                    }
                )
            previous_info = info
        if len(current_anchors) != len(source_outline):
            return []
        specs: List[dict] = []
        for index, anchor in enumerate(current_anchors):
            outline_item = source_outline[index]
            range_id = str(outline_item.get("outline_id", "") or "")
            if not range_id:
                return []
            start = max(target_start, int(anchor["start"]))
            end = target_end if index + 1 >= len(current_anchors) else min(target_end, int(current_anchors[index + 1]["start"]))
            if end > start:
                specs.append({"range_id": range_id, "start": start, "end": end})
        return specs

    @staticmethod
    def _figure_block_start(info: dict, previous_info: Optional[dict], image_anchor_starts: set[int]) -> int:
        current_start = int(info.get("range_start", 0) or 0)
        if current_start in image_anchor_starts:
            return current_start
        if previous_info and int(previous_info.get("range_start", 0) or 0) in image_anchor_starts:
            return int(previous_info.get("range_start", 0) or current_start)
        return current_start

    @staticmethod
    def _table_caption_block_start(info: dict, previous_info: Optional[dict], table_infos: List[dict]) -> int:
        current_start = int(info.get("range_start", 0) or 0)
        if previous_info and FormatApplier._inside_table(previous_info, table_infos):
            for table_info in table_infos:
                if current_start >= int(table_info["range_start"]) and current_start <= int(table_info["range_end"]):
                    return int(table_info["range_start"])
            for table_info in table_infos:
                start = int(table_info["range_start"])
                end = int(table_info["range_end"])
                previous_start = int(previous_info.get("range_start", 0) or 0)
                previous_end = int(previous_info.get("range_end", 0) or 0)
                if previous_start >= start and previous_end <= end:
                    return start
        return current_start

    def _apply_heading(
        self,
        paragraph,
        category: str,
        profile,
        formatter: NumberingFormatter,
        started_levels: set[int],
        outline_item: Optional[dict],
    ) -> None:
        level = int(category.replace("标题", ""))
        pattern = self._resolve_profile_pattern(profile)
        text = self._heading_text(paragraph, outline_item)
        existing_prefix = self._outline_prefix(outline_item)
        if profile.renumber or profile.clear_existing_numbering:
            self._clear_paragraph_list_numbering(paragraph)
        if level not in started_levels and not profile.continue_previous_numbering:
            formatter.clear_from_level(level)
        if profile.renumber or profile.clear_existing_numbering:
            text = strip_existing_numbering(text)
        prefix = ""
        if profile.renumber and pattern:
            prefix = formatter.next_for_level(level, pattern)
        elif not profile.clear_existing_numbering:
            prefix = existing_prefix
            formatter.clear_from_level(level)
        else:
            formatter.clear_from_level(level)
        started_levels.add(level)
        self._rewrite_prefixed_paragraph(paragraph, prefix, text, profile.space_after_numbering)

    def _apply_caption(self, paragraph, profile, formatter: NumberingFormatter, outline_item: Optional[dict] = None) -> None:
        pattern = self._resolve_profile_pattern(profile)
        text = self._heading_text(paragraph, outline_item)
        existing_prefix = self._outline_prefix(outline_item)
        if profile.renumber or profile.clear_existing_numbering:
            self._clear_paragraph_list_numbering(paragraph)
            text = strip_existing_numbering(text)
        if profile.renumber and pattern:
            prefix = formatter.next_for_level(1, pattern)
        elif not profile.clear_existing_numbering:
            prefix = existing_prefix
        else:
            prefix = ""
        self._rewrite_prefixed_paragraph(paragraph, prefix, text, profile.space_after_numbering)

    @staticmethod
    def _pattern_uses_explicit_levels(pattern: str, level: int) -> bool:
        return any(f"{{{index}" in (pattern or "") for index in range(1, min(level, 9) + 1))

    @classmethod
    def _select_numbering_values(cls, pattern: str, level: int, values: List[int]) -> List[int]:
        cleaned: List[int] = []
        for value in values[:level]:
            try:
                numeric = max(int(value), 0)
            except (TypeError, ValueError):
                numeric = 0
            if numeric > 0:
                cleaned.append(numeric)
        if not cleaned:
            return []
        if cls._pattern_uses_explicit_levels(pattern, level):
            return cleaned[:level]
        return [cleaned[-1]]

    @classmethod
    def _detect_heading_numbering_values(cls, outline_item: Optional[dict], level: int, pattern: str) -> List[int]:
        for candidate in cls._outline_numbering_candidates(outline_item):
            style = detect_heading_numbering_style(candidate, "", level)
            if style == "none":
                continue
            values = cls._select_numbering_values(pattern, level, detect_numbering_values(candidate, style))
            if values:
                return values
        return []

    @classmethod
    def _detect_caption_numbering_values(cls, outline_item: Optional[dict], numbering_style: str, pattern: str) -> List[int]:
        if not numbering_style or numbering_style == "custom":
            return []
        for candidate in cls._outline_numbering_candidates(outline_item):
            values = cls._select_numbering_values(pattern, 1, detect_numbering_values(candidate, numbering_style))
            if values:
                return values
        return []

    @staticmethod
    def _outline_numbering_candidates(outline_item: Optional[dict]) -> List[str]:
        if not outline_item:
            return []
        candidates: List[str] = []
        for raw_value in [outline_item.get("original_prefix", ""), outline_item.get("original_text", "")]:
            cleaned = str(raw_value or "").strip()
            if cleaned and cleaned not in candidates:
                candidates.append(cleaned)
        return candidates

    @staticmethod
    def _sort_outline_by_source_order(outline: List[dict]) -> List[dict]:
        def sort_key(item_with_index):
            index, item = item_with_index
            try:
                source_order = int(item.get("source_order", 0) or 0)
            except Exception:
                source_order = 0
            try:
                range_start = int(item.get("range_start", 0) or 0)
            except Exception:
                range_start = 0
            order_key = source_order if source_order > 0 else (range_start if range_start > 0 else index + 1)
            return (order_key, range_start if range_start > 0 else index + 1, index)

        return [dict(item) for _, item in sorted(enumerate(outline), key=sort_key)]

    @staticmethod
    def _build_output_path(source_path: str) -> str:
        source = Path(source_path)
        return str(source.with_name(f"{source.stem}_格式化.docx"))

    @staticmethod
    def _resolve_profile_pattern(profile) -> str:
        if profile.numbering_style == "custom":
            pattern = (profile.custom_numbering_pattern or "").strip()
            profile.numbering_pattern = pattern
            return pattern
        if profile.numbering_style:
            pattern = resolve_numbering_pattern(profile.numbering_style, profile.custom_numbering_pattern)
            profile.numbering_pattern = pattern
            return pattern
        return (profile.numbering_pattern or "").strip()

    @staticmethod
    def _resolve_outline_category(default_category: str, outline_item: Optional[dict]) -> str:
        level = outline_item.get("level") if outline_item else None
        if isinstance(level, int) and 1 <= level <= 9:
            return f"标题{level}"
        return default_category

    @staticmethod
    def _outline_prefix(outline_item: Optional[dict]) -> str:
        if not outline_item:
            return ""
        return str(outline_item.get("original_prefix", "") or "").strip()

    @staticmethod
    def _heading_text(paragraph, outline_item: Optional[dict]) -> str:
        if outline_item and outline_item.get("text"):
            return str(outline_item["text"])
        return paragraph.Range.Text.rstrip("\r")

    @staticmethod
    def _rewrite_prefixed_paragraph(paragraph, prefix: str, text: str, space_after_numbering: bool = True) -> None:
        style_name = FormatApplier._paragraph_style_name(paragraph)
        original_text = paragraph.Range.Text.rstrip("\r")
        heading_level = FormatApplier._safe_outline_level(paragraph)
        if not heading_level and style_name:
            heading_level = detect_heading_level(style_name, original_text, "")
        body = text.strip()
        separator = " " if space_after_numbering else ""
        if prefix and body:
            new_text = f"{prefix}{separator}{body}"
        elif prefix:
            new_text = prefix
        else:
            new_text = body
        text_range = paragraph.Range.Duplicate
        if text_range.End > text_range.Start:
            text_range.End -= 1
        text_range.Text = new_text
        if heading_level:
            try:
                paragraph.OutlineLevel = heading_level
            except Exception:
                pass
            try:
                paragraph.Range.ParagraphFormat.OutlineLevel = heading_level
            except Exception:
                pass

    @staticmethod
    def _paragraph_style(paragraph):
        try:
            return paragraph.Style
        except Exception:
            try:
                return paragraph.Range.Style
            except Exception:
                return None

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

    @staticmethod
    def _paragraph_style_name(paragraph) -> str:
        return FormatApplier._style_name(FormatApplier._paragraph_style(paragraph))

    @staticmethod
    def _safe_outline_level(paragraph) -> int:
        try:
            level = int(paragraph.OutlineLevel)
            if 1 <= level <= 9:
                return level
        except Exception:
            pass
        try:
            level = int(paragraph.Range.ParagraphFormat.OutlineLevel)
            if 1 <= level <= 9:
                return level
        except Exception:
            pass
        return 0

    @staticmethod
    def _restore_heading_style(paragraph, style_obj, style_name: str, heading_level: int) -> None:
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
    def _clear_paragraph_list_numbering(paragraph) -> None:
        try:
            paragraph.Range.ListFormat.RemoveNumbers()
        except Exception:
            pass

    @staticmethod
    def _collect_image_anchor_starts(inline_shapes: List, floating_shapes_by_page: dict) -> set[int]:
        anchors: set[int] = set()
        for shape in inline_shapes:
            try:
                anchors.add(int(shape.Range.Paragraphs(1).Range.Start))
            except Exception:
                continue
        for shapes in floating_shapes_by_page.values():
            for shape in shapes:
                try:
                    anchors.add(int(shape.Anchor.Paragraphs(1).Range.Start))
                except Exception:
                    continue
        return anchors

    @staticmethod
    def _classify_target_paragraph(
        info: dict,
        table_infos: List[dict],
        image_anchor_starts: Optional[set[int]] = None,
        previous_info: Optional[dict] = None,
    ) -> Optional[str]:
        text = info["text"].strip()
        if not text:
            return None
        if FormatApplier._inside_table(info, table_infos):
            return None

        style_name = info.get("style_name", "")
        list_string = info.get("list_string", "")
        heading_level = detect_heading_level(style_name, text, list_string)
        if heading_level:
            return f"标题{heading_level}"

        image_anchor_starts = image_anchor_starts or set()
        current_is_image_anchor = info.get("range_start") in image_anchor_starts
        previous_is_image_anchor = bool(previous_info and previous_info.get("range_start") in image_anchor_starts)
        previous_is_table = bool(previous_info and FormatApplier._inside_table(previous_info, table_infos))

        if looks_like_figure_caption(text, style_name):
            if current_is_image_anchor or previous_is_image_anchor or len(text) <= 80:
                return "图注"
        if looks_like_table_caption(text, style_name):
            if previous_is_table or len(text) <= 80:
                return "表注"
        return "正文"

    @staticmethod
    def _inside_table(info: dict, table_infos: List[dict]) -> bool:
        start = info["range_start"]
        end = info["range_end"]
        for table_info in table_infos:
            if start >= table_info["range_start"] and end <= table_info["range_end"]:
                return True
        return False
