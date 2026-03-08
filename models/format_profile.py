from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict, dataclass, field, fields, is_dataclass
from typing import Any, Dict, List

TEXT_ALIGNMENTS = ["左对齐", "居中", "右对齐", "两端对齐"]
LINE_SPACING_MODES = ["单倍", "1.5 倍", "2 倍", "多倍", "最小值", "固定值"]
TABLE_ALIGNMENTS = ["左对齐", "居中", "右对齐"]
VERTICAL_ALIGNMENTS = ["顶端", "居中", "底端"]
THREE_LINE_OPTIONS = ["是", "否", "不确定"]
WRAP_TYPES = ["嵌入型", "四周型", "紧密型", "上下型", "浮于文字上方", "衬于文字下方"]
NUMBERING_STYLE_OPTIONS = [
    ("无编号", "none"),
    ("第一章", "chapter"),
    ("一、", "chinese_comma"),
    ("（一）", "paren_chinese"),
    ("1", "decimal"),
    ("1、", "decimal_comma"),
    ("1.1", "decimal_2"),
    ("1.1.1", "decimal_3"),
    ("（1）", "paren_decimal"),
    ("图n.", "figure_decimal"),
    ("表n.", "table_decimal"),
    ("自定义", "custom"),
]
FONT_SIZE_PRESETS = [
    ("初号", 42.0),
    ("小初", 36.0),
    ("一号", 26.0),
    ("小一", 24.0),
    ("二号", 22.0),
    ("小二", 18.0),
    ("三号", 16.0),
    ("小三", 15.0),
    ("四号", 14.0),
    ("小四", 12.0),
    ("五号", 10.5),
    ("小五", 9.0),
    ("六号", 7.5),
    ("小六", 6.5),
    ("七号", 5.5),
    ("八号", 5.0),
]
FONT_SIZE_NAME_TO_VALUE = {name: value for name, value in FONT_SIZE_PRESETS}
DEFAULT_CATEGORIES = ["正文", *[f"标题{i}" for i in range(1, 10)], "图片", "图注", "表格", "表注"]


def default_numbering_style(level: int) -> str:
    if level == 1:
        return "chinese_comma"
    if level == 2:
        return "paren_chinese"
    if level == 3:
        return "decimal_comma"
    return "none"


@dataclass
class ParagraphFormat:
    font_name: str = "宋体"
    font_size: float = 12.0
    bold: bool = False
    italic: bool = False
    alignment: str = "左对齐"
    left_indent: float = 0.0
    right_indent: float = 0.0
    first_line_indent: float = 0.0
    space_before: float = 0.0
    space_after: float = 0.0
    line_spacing_mode: str = "1.5 倍"
    line_spacing: float = 1.5
    numbering_style: str = "none"
    custom_numbering_pattern: str = ""
    numbering_pattern: str = ""
    renumber: bool = False
    clear_existing_numbering: bool = False
    continue_previous_numbering: bool = False
    space_after_numbering: bool = True


@dataclass
class ImageFormat:
    alignment: str = "居中"
    space_before: float = 0.0
    space_after: float = 0.0
    line_spacing_mode: str = "单倍"
    line_spacing: float = 1.0
    wrap_type: str = "嵌入型"


@dataclass
class TableFormat:
    alignment: str = "居中"
    cell_vertical_alignment: str = "居中"
    top_border: bool = True
    header_bottom_border: bool = True
    bottom_border: bool = True
    inside_horizontal_border: bool = False
    inside_vertical_border: bool = False
    three_line_status: str = "不确定"


@dataclass
class FormatCatalog:
    body: ParagraphFormat = field(default_factory=ParagraphFormat)
    headings: Dict[str, ParagraphFormat] = field(
        default_factory=lambda: {
            f"标题{i}": ParagraphFormat(
                font_name="黑体",
                font_size=max(18.0 - i, 10.5),
                bold=True,
                line_spacing_mode="1.5 倍",
                line_spacing=1.5,
                numbering_style=default_numbering_style(i),
                numbering_pattern="" if i > 3 else ("{n:zh}、" if i == 1 else "（{n:zh}）" if i == 2 else "{n}、"),
                renumber=(i <= 3),
                clear_existing_numbering=(i <= 3),
                continue_previous_numbering=False,
            )
            for i in range(1, 10)
        }
    )
    image: ImageFormat = field(default_factory=ImageFormat)
    figure_caption: ParagraphFormat = field(
        default_factory=lambda: ParagraphFormat(
            font_name="宋体",
            font_size=10.5,
            alignment="居中",
            left_indent=0.0,
            right_indent=0.0,
            first_line_indent=0.0,
            line_spacing_mode="单倍",
            line_spacing=1.0,
            numbering_style="figure_decimal",
            numbering_pattern="图{n}.",
            renumber=False,
            clear_existing_numbering=False,
            continue_previous_numbering=False,
        )
    )
    table: TableFormat = field(default_factory=TableFormat)
    table_caption: ParagraphFormat = field(
        default_factory=lambda: ParagraphFormat(
            font_name="宋体",
            font_size=10.5,
            alignment="居中",
            left_indent=0.0,
            right_indent=0.0,
            first_line_indent=0.0,
            line_spacing_mode="单倍",
            line_spacing=1.0,
            numbering_style="table_decimal",
            numbering_pattern="表{n}.",
            renumber=False,
            clear_existing_numbering=False,
            continue_previous_numbering=False,
        )
    )
    category_order: List[str] = field(default_factory=lambda: list(DEFAULT_CATEGORIES))
    visible_categories: List[str] = field(default_factory=lambda: list(DEFAULT_CATEGORIES))

    @classmethod
    def category_names(cls) -> List[str]:
        return list(DEFAULT_CATEGORIES)

    @classmethod
    def default_visible_category_names(cls) -> List[str]:
        return list(DEFAULT_CATEGORIES)

    def ordered_categories(self) -> List[str]:
        names: List[str] = []
        seen = set()
        for name in self.category_order:
            if name not in seen and self._supports_category(name):
                names.append(name)
                seen.add(name)
        for name in self.category_names():
            if name not in seen and self._supports_category(name):
                names.append(name)
        return names

    def ordered_visible_categories(self) -> List[str]:
        visible = set(self.visible_categories or [])
        return [name for name in self.ordered_categories() if name in visible]

    def set_visible_categories(self, names: List[str]) -> None:
        allowed = [name for name in names if self._supports_category(name)]
        self.visible_categories = allowed or ["正文"]

    def set_category_order(self, names: List[str]) -> None:
        ordered = [name for name in names if self._supports_category(name)]
        self.category_order = ordered or ["正文"]

    def reset_category_lists(self) -> None:
        names = self.ordered_categories()
        self.category_order = list(names)
        self.visible_categories = list(names)

    def get_category(self, name: str) -> Any:
        if name == "正文":
            return self.body
        if name.startswith("标题"):
            return self.headings[name]
        if name == "图片":
            return self.image
        if name == "图注":
            return self.figure_caption
        if name == "表格":
            return self.table
        if name == "表注":
            return self.table_caption
        raise KeyError(f"未知类别: {name}")

    def set_category(self, name: str, value: Any) -> None:
        if name == "正文":
            self.body = value
            return
        if name.startswith("标题"):
            self.headings[name] = value
            return
        if name == "图片":
            self.image = value
            return
        if name == "图注":
            self.figure_caption = value
            return
        if name == "表格":
            self.table = value
            return
        if name == "表注":
            self.table_caption = value
            return
        raise KeyError(f"未知类别: {name}")

    def clone(self) -> "FormatCatalog":
        return deepcopy(self)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "正文": asdict(self.body),
            **{name: asdict(profile) for name, profile in self.headings.items()},
            "图片": asdict(self.image),
            "图注": asdict(self.figure_caption),
            "表格": asdict(self.table),
            "表注": asdict(self.table_caption),
            "__meta__": {
                "category_order": list(self.category_order),
                "visible_categories": list(self.visible_categories),
            },
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "FormatCatalog":
        catalog = cls()
        meta = data.get("__meta__") if isinstance(data.get("__meta__"), dict) else {}
        for name, payload in data.items():
            if name == "__meta__":
                continue
            current = catalog.get_category(name)
            if is_dataclass(current):
                allowed_fields = {item.name for item in fields(type(current))}
                normalized_payload = {key: value for key, value in payload.items() if key in allowed_fields}
                catalog.set_category(name, type(current)(**normalized_payload))
        catalog.set_category_order(meta.get("category_order") or catalog.category_names())
        catalog.set_visible_categories(meta.get("visible_categories") or catalog.category_names())
        return catalog

    def merge_from(self, other: "FormatCatalog") -> None:
        for name in self.category_names():
            self.set_category(name, deepcopy(other.get_category(name)))
        self.category_order = list(other.category_order)
        self.visible_categories = list(other.visible_categories)

    @staticmethod
    def _supports_category(name: str) -> bool:
        if name == "正文":
            return True
        if name == "图片":
            return True
        if name == "图注":
            return True
        if name == "表格":
            return True
        if name == "表注":
            return True
        if name.startswith("标题"):
            suffix = name[2:]
            return suffix.isdigit() and 1 <= int(suffix) <= 9
        return False


def font_size_to_text(value: Any) -> str:
    numeric = round_float(value)
    for name, points in FONT_SIZE_PRESETS:
        if abs(points - numeric) < 0.01:
            return name
    return f"{numeric:g}"


def font_size_from_text(text: str, fallback: float = 12.0) -> float:
    cleaned = (text or "").strip().replace("磅", "").replace("pt", "").replace("PT", "")
    if cleaned in FONT_SIZE_NAME_TO_VALUE:
        return FONT_SIZE_NAME_TO_VALUE[cleaned]
    try:
        return round(float(cleaned), 2)
    except (TypeError, ValueError):
        return fallback


def round_float(value: Any, digits: int = 2) -> float:
    try:
        return round(float(value), digits)
    except (TypeError, ValueError):
        return 0.0
