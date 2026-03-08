from __future__ import annotations

import re
from typing import List, Optional, Tuple

TOKEN_PATTERN = re.compile(r"\{(\d+|n)(?::([a-z_]+))?\}")
HEADING_STYLE_PATTERN = re.compile(r"(?:标题|Heading)\s*(\d+)", re.IGNORECASE)
BODY_STYLE_PATTERN = re.compile(r"(?:正文|normal|body\s*text)", re.IGNORECASE)
TOC_STYLE_PATTERN = re.compile(r"(?:^toc\b|目录)", re.IGNORECASE)
FIGURE_CAPTION_STYLE_PATTERN = re.compile(r"(?:图注|figure\s*caption)", re.IGNORECASE)
TABLE_CAPTION_STYLE_PATTERN = re.compile(r"(?:表注|table\s*caption)", re.IGNORECASE)
CHINESE_DIGITS = "零一二三四五六七八九"
CHINESE_UNITS = ["", "十", "百", "千"]
SECTION_UNITS = ["", "万", "亿"]
CHINESE_NUMBER_CHARS = "零〇一二两三四五六七八九十百千万"
CHAPTER_VALUE_PATTERN = re.compile(rf"^第\s*([{CHINESE_NUMBER_CHARS}\d]+)[章节篇部]\s*")
CHINESE_COMMA_VALUE_PATTERN = re.compile(rf"^\s*([{CHINESE_NUMBER_CHARS}]+)、\s*")
PAREN_CHINESE_VALUE_PATTERN = re.compile(rf"^[（(]\s*([{CHINESE_NUMBER_CHARS}]+)\s*[)）]\s*")
DECIMAL_COMMA_VALUE_PATTERN = re.compile(r"^\s*(\d+)、\s*")
DECIMAL_DOT_VALUE_PATTERN = re.compile(r"^\s*(\d+(?:[.．]\d+)+)")
PAREN_DECIMAL_VALUE_PATTERN = re.compile(r"^[（(]\s*(\d+)\s*[)）]\s*")
SIMPLE_DECIMAL_VALUE_PATTERN = re.compile(r"^\s*(\d+)[.)．]\s*")
FIGURE_CAPTION_VALUE_PATTERN = re.compile(r"^图\s*(\d+(?:[-－—.．]\d+)*)")
TABLE_CAPTION_VALUE_PATTERN = re.compile(r"^表\s*(\d+(?:[-－—.．]\d+)*)")
NUMBERING_STYLE_TO_PATTERN = {
    "none": "",
    "chapter": "第{n:zh}章",
    "chinese_comma": "{n:zh}、",
    "paren_chinese": "（{n:zh}）",
    "decimal": "{n}",
    "decimal_comma": "{n}、",
    "decimal_2": "{1}.{n}",
    "decimal_3": "{1}.{2}.{n}",
    "paren_decimal": "（{n}）",
    "figure_decimal": "图{n}.",
    "table_decimal": "表{n}.",
    "custom": "",
}
CHAPTER_PATTERN = re.compile(rf"^第[{CHINESE_NUMBER_CHARS}\d]+[章节篇部]\s*")
CHINESE_COMMA_PATTERN = re.compile(rf"^[{CHINESE_NUMBER_CHARS}]+、\s*")
PAREN_CHINESE_PATTERN = re.compile(rf"^[（(]\s*[{CHINESE_NUMBER_CHARS}]+\s*[)）]\s*")
DECIMAL_COMMA_PATTERN = re.compile(r"^\d+、\s*")
DECIMAL_DOT_PATTERN = re.compile(r"^\d+(?:\.\d+)+(?:\s*[、.．:：])?\s*")
PAREN_DECIMAL_PATTERN = re.compile(r"^[（(]\s*\d+\s*[)）]\s*")
SIMPLE_DECIMAL_PATTERN = re.compile(r"^\d+[.)．]\s*")
FIGURE_CAPTION_PATTERN = re.compile(r"^图\s*\d+(?:[-－—.．]\d+)*(?:\s*[.．、:：])?\s*")
TABLE_CAPTION_PATTERN = re.compile(r"^表\s*\d+(?:[-－—.．]\d+)*(?:\s*[.．、:：])?\s*")
LEADING_NUMBERING_PATTERNS: List[re.Pattern[str]] = [
    CHAPTER_PATTERN,
    FIGURE_CAPTION_PATTERN,
    TABLE_CAPTION_PATTERN,
    DECIMAL_DOT_PATTERN,
    CHINESE_COMMA_PATTERN,
    PAREN_CHINESE_PATTERN,
    DECIMAL_COMMA_PATTERN,
    PAREN_DECIMAL_PATTERN,
    SIMPLE_DECIMAL_PATTERN,
]


def resolve_numbering_pattern(style: str, custom_pattern: str = "") -> str:
    if style == "custom":
        return (custom_pattern or "").strip()
    return NUMBERING_STYLE_TO_PATTERN.get(style, "")


def detect_heading_level(style_name: str, text: str, list_string: str = "") -> Optional[int]:
    match = HEADING_STYLE_PATTERN.search(style_name or "")
    if match:
        level = int(match.group(1))
        return level if 1 <= level <= 9 else None
    if BODY_STYLE_PATTERN.search(style_name or ""):
        return None
    if TOC_STYLE_PATTERN.search(style_name or ""):
        return None
    if not _is_probable_structured_paragraph(text, list_string):
        return None
    for candidate in _heading_candidates(text, list_string):
        if not candidate:
            continue
        if CHAPTER_PATTERN.match(candidate) or CHINESE_COMMA_PATTERN.match(candidate):
            return 1
        if PAREN_CHINESE_PATTERN.match(candidate):
            return 2
        if DECIMAL_COMMA_PATTERN.match(candidate):
            return 3
        dot_match = DECIMAL_DOT_PATTERN.match(candidate)
        if dot_match:
            depth = dot_match.group(0).count(".") + 1
            return min(max(depth, 1), 9)
    return None


def detect_heading_numbering_style(text: str, list_string: str = "", level: Optional[int] = None) -> str:
    for candidate in _heading_candidates(text, list_string):
        if not candidate:
            continue
        if CHAPTER_PATTERN.match(candidate):
            return "chapter"
        if CHINESE_COMMA_PATTERN.match(candidate):
            return "chinese_comma"
        if PAREN_CHINESE_PATTERN.match(candidate):
            return "paren_chinese"
        if DECIMAL_COMMA_PATTERN.match(candidate):
            return "decimal_comma"
        dot_match = DECIMAL_DOT_PATTERN.match(candidate)
        if dot_match:
            depth = dot_match.group(0).count(".") + 1
            if depth <= 1:
                return "decimal"
            if depth == 2:
                return "decimal_2"
            return "decimal_3"
        if PAREN_DECIMAL_PATTERN.match(candidate):
            return "paren_decimal"
    if level == 1:
        return "chinese_comma"
    if level == 2:
        return "paren_chinese"
    if level == 3:
        return "decimal_comma"
    return "none"


def looks_like_figure_caption(text: str, style_name: str = "") -> bool:
    cleaned = (text or "").strip()
    if not cleaned:
        return False
    if FIGURE_CAPTION_STYLE_PATTERN.search(style_name or ""):
        return True
    return bool(FIGURE_CAPTION_PATTERN.match(cleaned))


def looks_like_table_caption(text: str, style_name: str = "") -> bool:
    cleaned = (text or "").strip()
    if not cleaned:
        return False
    if TABLE_CAPTION_STYLE_PATTERN.search(style_name or ""):
        return True
    return bool(TABLE_CAPTION_PATTERN.match(cleaned))


def split_leading_numbering(text: str) -> Tuple[str, str]:
    result = (text or "").strip()
    for pattern in LEADING_NUMBERING_PATTERNS:
        match = pattern.match(result)
        if match:
            prefix = match.group(0).strip()
            body = result[match.end():].strip()
            return prefix, body
    return "", result


def number_to_chinese_lower(value: int) -> str:
    try:
        value = int(value)
    except (TypeError, ValueError):
        return "零"
    if value <= 0:
        return "零"
    if value < 10:
        return CHINESE_DIGITS[value]

    sections: List[str] = []
    unit_index = 0
    remaining = value
    need_zero = False

    while remaining > 0:
        section = remaining % 10000
        remaining //= 10000
        if section == 0:
            need_zero = True
            unit_index += 1
            continue
        section_text = _section_to_chinese(section)
        if need_zero and sections and not section_text.endswith("零"):
            section_text += "零"
        section_text += SECTION_UNITS[unit_index]
        sections.append(section_text)
        need_zero = section < 1000
        unit_index += 1

    result = "".join(reversed(sections)).replace("零零", "零")
    result = result.rstrip("零")
    if result.startswith("一十"):
        result = result[1:]
    return result or "零"


def _section_to_chinese(value: int) -> str:
    result = ""
    zero_pending = False
    digits = [int(char) for char in str(value)]
    total = len(digits)
    for index, digit in enumerate(digits):
        unit_pos = total - index - 1
        if digit == 0:
            zero_pending = result != ""
            continue
        if zero_pending:
            result += "零"
            zero_pending = False
        result += CHINESE_DIGITS[digit] + CHINESE_UNITS[unit_pos]
    return result


def chinese_to_number(value: str) -> int:
    text = (value or "").strip().replace("〇", "零")
    if not text:
        return 0
    if text.isdigit():
        return int(text)
    total = 0
    section = 0
    number = 0
    digit_map = {char: index for index, char in enumerate(CHINESE_DIGITS)}
    digit_map["两"] = 2
    unit_map = {"十": 10, "百": 100, "千": 1000}
    section_unit_map = {"万": 10000, "亿": 100000000}
    for char in text:
        if char in digit_map:
            number = digit_map[char]
            continue
        if char in unit_map:
            if number == 0:
                number = 1
            section += number * unit_map[char]
            number = 0
            continue
        if char in section_unit_map:
            section = (section + number) * section_unit_map[char]
            total += section
            section = 0
            number = 0
            continue
        return 0
    return total + section + number


class NumberingFormatter:
    def __init__(self) -> None:
        self.counters = [0] * 9

    def reset(self) -> None:
        self.counters = [0] * 9

    def clear_from_level(self, level: int) -> None:
        if level < 1 or level > 9:
            return
        for index in range(level - 1, 9):
            self.counters[index] = 0

    def set_counters(self, counters: List[int]) -> None:
        normalized = [0] * 9
        for index, value in enumerate(counters[:9]):
            try:
                normalized[index] = max(int(value), 0)
            except (TypeError, ValueError):
                normalized[index] = 0
        self.counters = normalized

    def ensure_counters(self, values: List[int]) -> None:
        for index, value in enumerate(values[:9]):
            try:
                numeric = max(int(value), 0)
            except (TypeError, ValueError):
                numeric = 0
            if numeric > self.counters[index]:
                self.counters[index] = numeric

    def preview_for_level(self, level: int, pattern: str) -> str:
        snapshot = list(self.counters)
        try:
            return self.next_for_level(level, pattern)
        finally:
            self.counters = snapshot

    def format_values(self, pattern: str, values: List[int]) -> str:
        if not pattern.strip():
            return ""
        normalized: List[int] = []
        for value in values:
            try:
                normalized.append(max(int(value), 0))
            except (TypeError, ValueError):
                normalized.append(0)

        def replace(match: re.Match[str]) -> str:
            token = match.group(1)
            formatter = (match.group(2) or "").strip().lower()
            if token == "n":
                raw_value = normalized[-1] if normalized else 0
            else:
                number = int(token)
                if 1 <= number <= len(normalized):
                    raw_value = normalized[number - 1]
                else:
                    return ""
            if formatter == "zh":
                return number_to_chinese_lower(raw_value)
            return str(raw_value)

        return TOKEN_PATTERN.sub(replace, pattern).strip()

    def apply_values_for_level(self, level: int, pattern: str, values: List[int]) -> str:
        if level < 1 or level > 9:
            return ""
        cleaned: List[int] = []
        for value in values[:level]:
            try:
                cleaned.append(max(int(value), 0))
            except (TypeError, ValueError):
                cleaned.append(0)
        cleaned = [value for value in cleaned if value > 0]
        if not cleaned:
            return self.next_for_level(level, pattern)
        if len(cleaned) >= level:
            normalized = cleaned[:level]
        else:
            normalized = list(self.counters[:level])
            start = level - len(cleaned)
            for offset, value in enumerate(cleaned):
                normalized[start + offset] = value
        for index in range(level):
            self.counters[index] = normalized[index]
        for index in range(level, 9):
            self.counters[index] = 0
        return self.format_values(pattern, normalized)

    def next_for_level(self, level: int, pattern: str) -> str:
        if level < 1 or level > 9:
            return ""
        self.counters[level - 1] += 1
        for index in range(level, 9):
            self.counters[index] = 0
        return self.format_values(pattern, self.counters[:level])


def strip_existing_numbering(text: str) -> str:
    return split_leading_numbering(text)[1]


def detect_numbering_values(text: str, style: str) -> List[int]:
    candidate = (text or "").strip()
    if not candidate:
        return []
    if style == "chapter":
        match = CHAPTER_VALUE_PATTERN.match(candidate)
        return [_parse_number_token(match.group(1))] if match else []
    if style == "chinese_comma":
        match = CHINESE_COMMA_VALUE_PATTERN.match(candidate)
        return [_parse_number_token(match.group(1))] if match else []
    if style == "paren_chinese":
        match = PAREN_CHINESE_VALUE_PATTERN.match(candidate)
        return [_parse_number_token(match.group(1))] if match else []
    if style == "decimal":
        match = SIMPLE_DECIMAL_VALUE_PATTERN.match(candidate)
        return [_safe_int(match.group(1))] if match else []
    if style == "decimal_comma":
        match = DECIMAL_COMMA_VALUE_PATTERN.match(candidate)
        return [_safe_int(match.group(1))] if match else []
    if style == "decimal_2":
        match = DECIMAL_DOT_VALUE_PATTERN.match(candidate)
        return _parse_decimal_parts(match.group(1), expected_parts=2) if match else []
    if style == "decimal_3":
        match = DECIMAL_DOT_VALUE_PATTERN.match(candidate)
        return _parse_decimal_parts(match.group(1), expected_parts=3) if match else []
    if style == "paren_decimal":
        match = PAREN_DECIMAL_VALUE_PATTERN.match(candidate)
        return [_safe_int(match.group(1))] if match else []
    if style == "figure_decimal":
        match = FIGURE_CAPTION_VALUE_PATTERN.match(candidate)
        return _parse_decimal_parts(match.group(1)) if match else []
    if style == "table_decimal":
        match = TABLE_CAPTION_VALUE_PATTERN.match(candidate)
        return _parse_decimal_parts(match.group(1)) if match else []
    return []


def _heading_candidates(text: str, list_string: str = "") -> List[str]:
    candidates: List[str] = []
    list_part = (list_string or "").strip()
    text_part = (text or "").strip()
    for candidate in [list_part, text_part, f"{list_part}{text_part}" if list_part and text_part else ""]:
        cleaned = candidate.strip()
        if cleaned and cleaned not in candidates:
            candidates.append(cleaned)
    return candidates


def _safe_int(value: str) -> int:
    try:
        return max(int(value), 0)
    except (TypeError, ValueError):
        return 0


def _parse_number_token(token: str) -> int:
    cleaned = (token or "").strip()
    if not cleaned:
        return 0
    if cleaned.isdigit():
        return int(cleaned)
    return chinese_to_number(cleaned)


def _parse_decimal_parts(value: str, expected_parts: int = 0) -> List[int]:
    cleaned = (value or "").strip()
    if not cleaned:
        return []
    normalized = re.split(r"[-－—.．]", cleaned)
    parts = [_safe_int(part) for part in normalized if str(part).strip()]
    if not parts or any(part <= 0 for part in parts):
        return []
    if expected_parts and len(parts) < expected_parts:
        return []
    if expected_parts:
        return parts[:expected_parts]
    return parts


def _is_probable_structured_paragraph(text: str, list_string: str = "") -> bool:
    if (list_string or "").strip():
        return True
    cleaned = re.sub(r"\s+", "", (text or "").strip())
    if not cleaned:
        return False
    if len(cleaned) > 80:
        return False
    return not cleaned.endswith(("。", "；", "，", "：", ":", ";", "!", "！", "?", "？"))
