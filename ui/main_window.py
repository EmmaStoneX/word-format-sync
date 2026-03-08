from __future__ import annotations

import json
from pathlib import Path
from typing import Callable, List, Optional

from PySide6.QtCore import QItemSelectionModel, QMimeData, QPoint, QSignalBlocker, QTimer, Qt, Signal
from PySide6.QtGui import QColor, QDrag, QDragEnterEvent, QDropEvent, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFontComboBox,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHeaderView,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from models.format_profile import (
    FONT_SIZE_PRESETS,
    FormatCatalog,
    ImageFormat,
    LINE_SPACING_MODES,
    NUMBERING_STYLE_OPTIONS,
    ParagraphFormat,
    TABLE_ALIGNMENTS,
    TEXT_ALIGNMENTS,
    THREE_LINE_OPTIONS,
    TableFormat,
    VERTICAL_ALIGNMENTS,
    WRAP_TYPES,
    font_size_from_text,
    font_size_to_text,
)
from services.applier import FormatApplier
from services.extractor import ExtractResult, FormatExtractor
from services.word_com import WordAutomationError, WordComClient
from utils.numbering import (
    NumberingFormatter,
    resolve_numbering_pattern,
)

WORD_SUFFIXES = {".doc", ".docx"}


class FileDropLineEdit(QLineEdit):
    fileDropped = Signal(str)

    def __init__(self, placeholder: str) -> None:
        super().__init__()
        self.setAcceptDrops(True)
        self.setPlaceholderText(placeholder)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        urls = event.mimeData().urls()
        if len(urls) == 1 and urls[0].isLocalFile():
            suffix = Path(urls[0].toLocalFile()).suffix.lower()
            if suffix in WORD_SUFFIXES:
                event.acceptProposedAction()
                return
        event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        if len(urls) != 1 or not urls[0].isLocalFile():
            event.ignore()
            return
        path = urls[0].toLocalFile()
        if Path(path).suffix.lower() not in WORD_SUFFIXES:
            event.ignore()
            return
        self.setText(path)
        self.fileDropped.emit(path)
        event.acceptProposedAction()


class OutlineTableWidget(QTableWidget):
    rowsMoved = Signal(object, int)
    dragPreviewChanged = Signal(object, int)
    dragPreviewCleared = Signal()

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self._drag_origin_rows: List[int] = []
        self._preview_row: Optional[int] = None
        self._drop_highlight_rows: List[int] = []
        self._auto_scroll_margin = 28
        self._auto_scroll_step = 3
        self._last_drag_pos = QPoint()
        self._auto_scroll_timer = QTimer(self)
        self._auto_scroll_timer.setInterval(30)
        self._auto_scroll_timer.timeout.connect(self._perform_auto_scroll)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropOverwriteMode(False)
        self.setDragDropMode(QAbstractItemView.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.verticalScrollBar().setSingleStep(20)

    def startDrag(self, supportedActions) -> None:
        selected_rows = self._selected_rows()
        if not selected_rows:
            current_row = self.currentRow()
            if current_row < 0:
                return
            selected_rows = [current_row]
        self._drag_origin_rows = selected_rows
        self._preview_row = None
        self._set_drop_highlight_rows(selected_rows)
        drag = QDrag(self)
        mime = QMimeData()
        mime.setData("application/x-outline-rows", json.dumps(selected_rows).encode("utf-8"))
        drag.setMimeData(mime)
        drag.exec(Qt.MoveAction)
        self._finish_drag_state()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.source() is self:
            event.setDropAction(Qt.MoveAction)
            event.accept()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:
        if event.source() is self and self._drag_origin_rows:
            self._last_drag_pos = event.position().toPoint()
            self._update_auto_scroll_state()
            insert_row = self._resolve_insert_row(self._last_drag_pos.y())
            if insert_row != self._preview_row:
                self._preview_row = insert_row
                self.dragPreviewChanged.emit(list(self._drag_origin_rows), insert_row)
            event.setDropAction(Qt.MoveAction)
            event.accept()
            return
        super().dragMoveEvent(event)

    def dragLeaveEvent(self, event) -> None:
        if self._drag_origin_rows:
            self._preview_row = None
            self.dragPreviewCleared.emit()
        self._finish_drag_visuals()
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        if event.source() is not self or not self._drag_origin_rows:
            super().dropEvent(event)
            return
        origin_rows = list(self._drag_origin_rows)
        target_row = self._preview_row
        if target_row is not None:
            self.rowsMoved.emit(origin_rows, target_row)
        else:
            self.dragPreviewCleared.emit()
        self._finish_drag_state()
        event.setDropAction(Qt.MoveAction)
        event.accept()

    def set_drag_target_rows(self, rows: List[int]) -> None:
        highlight_rows = self._selected_rows() if rows is None else sorted({row for row in rows if 0 <= row < self.rowCount()})
        self._set_drop_highlight_rows(highlight_rows)

    def clear_drag_target_rows(self) -> None:
        self._set_drop_highlight_rows(self._selected_rows())

    def _finish_drag_state(self) -> None:
        self._drag_origin_rows = []
        self._preview_row = None
        self.dragPreviewCleared.emit()
        self._finish_drag_visuals()

    def _finish_drag_visuals(self) -> None:
        self._auto_scroll_timer.stop()
        self._set_drop_highlight_rows(self._selected_rows())

    def _selected_rows(self) -> List[int]:
        selection_model = self.selectionModel()
        if selection_model is None:
            return []
        return sorted({index.row() for index in selection_model.selectedRows()})

    def _resolve_insert_row(self, pos_y: int) -> int:
        if self.rowCount() <= 0:
            return 0
        row = self.rowAt(pos_y)
        if row < 0:
            return self.rowCount()
        rect = self.visualRect(self.model().index(row, 0))
        return row + 1 if pos_y > rect.center().y() else row

    def _update_auto_scroll_state(self) -> None:
        viewport_height = self.viewport().height()
        if viewport_height <= 0:
            self._auto_scroll_timer.stop()
            return
        pos_y = self._last_drag_pos.y()
        if pos_y <= self._auto_scroll_margin or pos_y >= viewport_height - self._auto_scroll_margin:
            if not self._auto_scroll_timer.isActive():
                self._auto_scroll_timer.start()
        else:
            self._auto_scroll_timer.stop()

    def _perform_auto_scroll(self) -> None:
        if not self._drag_origin_rows:
            self._auto_scroll_timer.stop()
            return
        viewport_height = self.viewport().height()
        pos_y = self._last_drag_pos.y()
        scrollbar = self.verticalScrollBar()
        if pos_y <= self._auto_scroll_margin:
            scrollbar.setValue(scrollbar.value() - self._auto_scroll_step)
        elif pos_y >= viewport_height - self._auto_scroll_margin:
            scrollbar.setValue(scrollbar.value() + self._auto_scroll_step)
        else:
            self._auto_scroll_timer.stop()
            return
        insert_row = self._resolve_insert_row(self._last_drag_pos.y())
        if insert_row != self._preview_row:
            self._preview_row = insert_row
            self.dragPreviewChanged.emit(list(self._drag_origin_rows), insert_row)

    def _set_drop_highlight_rows(self, rows: List[int]) -> None:
        normalized = sorted({row for row in rows if 0 <= row < self.rowCount()})
        if normalized == self._drop_highlight_rows:
            return
        all_rows = set(self._drop_highlight_rows) | set(normalized)
        self._drop_highlight_rows = normalized
        for row in sorted(all_rows):
            self._apply_row_highlight(row, row in normalized)

    def refresh_row_styles(self) -> None:
        highlighted_rows = set(self._drop_highlight_rows)
        for row in range(self.rowCount()):
            self._apply_row_highlight(row, row in highlighted_rows)

    def _row_background_color(self, row: int, highlighted: bool) -> str:
        if highlighted:
            return "#cfe3ff"
        category_item = self.item(row, 2)
        if category_item is not None and category_item.text() == "标题1":
            return "#d9e6ff"
        return "transparent"

    def _apply_row_highlight(self, row: int, highlighted: bool) -> None:
        color_name = self._row_background_color(row, highlighted)
        color = QColor(color_name)
        for column in range(self.columnCount()):
            item = self.item(row, column)
            if item is not None:
                item.setBackground(color)
        widget = self.cellWidget(row, 3)
        if widget is not None:
            widget.setStyleSheet("" if color_name == "transparent" else f"background-color: {color_name};")


class ParagraphEditor(QWidget):
    def __init__(self, on_change: Optional[Callable[[], None]] = None) -> None:
        super().__init__()
        self._profile: Optional[ParagraphFormat] = None
        self._loading = False
        self._heading_mode = False
        self._caption_mode = False
        self._on_change = on_change

        layout = QFormLayout(self)
        layout.setLabelAlignment(Qt.AlignRight)

        self.font_name = QFontComboBox()
        self.font_size = QComboBox()
        self.font_size.setEditable(True)
        self.font_size.addItems([name for name, _ in FONT_SIZE_PRESETS])
        self.bold = QCheckBox("加粗")
        self.italic = QCheckBox("斜体")
        self.alignment = self._combo(TEXT_ALIGNMENTS)
        self.left_indent = self._double_spin(-1000, 1000, 0.1)
        self.left_indent.setSuffix(" 字符")
        self.right_indent = self._double_spin(-1000, 1000, 0.1)
        self.right_indent.setSuffix(" 字符")
        self.first_line_indent = self._double_spin(-1000, 1000, 0.1)
        self.first_line_indent.setSuffix(" 字符")
        self.space_before = self._double_spin(0, 1000, 0.1)
        self.space_after = self._double_spin(0, 1000, 0.1)
        self.line_spacing_mode = self._combo(LINE_SPACING_MODES)
        self.line_spacing = self._double_spin(0.5, 200, 0.1)
        self.numbering_style = QComboBox()
        for label, value in NUMBERING_STYLE_OPTIONS:
            self.numbering_style.addItem(label, value)
        self.custom_numbering_pattern = QLineEdit()
        self.custom_numbering_pattern.setPlaceholderText("自定义示例：第{n:zh}章、{n:zh}、图{n}.、表{n}.")
        self.renumber = QCheckBox("重新编号")
        self.clear_existing_numbering = QCheckBox("清除已有序号")
        self.continue_previous_numbering = QCheckBox("延续前文序号")
        self.space_after_numbering = QCheckBox("编号后加空格")

        self.left_indent_label = QLabel("左缩进（字符）")
        self.right_indent_label = QLabel("右缩进（字符）")
        self.first_line_indent_label = QLabel("首行缩进（字符）")
        self.numbering_style_label = QLabel("编号样式")
        self.custom_numbering_pattern_label = QLabel("自定义编号")
        self.numbering_options_label = QLabel("编号处理")

        layout.addRow("字体", self.font_name)
        layout.addRow("字号", self.font_size)
        layout.addRow("字形", self._row_widget(self.bold, self.italic))
        layout.addRow("对齐", self.alignment)
        layout.addRow(self.left_indent_label, self.left_indent)
        layout.addRow(self.right_indent_label, self.right_indent)
        layout.addRow(self.first_line_indent_label, self.first_line_indent)
        layout.addRow("段前", self.space_before)
        layout.addRow("段后", self.space_after)
        layout.addRow("行距方式", self.line_spacing_mode)
        layout.addRow("行距值", self.line_spacing)
        layout.addRow(self.numbering_style_label, self.numbering_style)
        layout.addRow(self.custom_numbering_pattern_label, self.custom_numbering_pattern)
        layout.addRow(self.numbering_options_label, self._row_widget(self.renumber, self.clear_existing_numbering, self.continue_previous_numbering, self.space_after_numbering))

        for widget in [
            self.font_name,
            self.font_size,
            self.bold,
            self.italic,
            self.alignment,
            self.left_indent,
            self.right_indent,
            self.first_line_indent,
            self.space_before,
            self.space_after,
            self.line_spacing,
            self.custom_numbering_pattern,
            self.clear_existing_numbering,
            self.continue_previous_numbering,
            self.space_after_numbering,
        ]:
            self._connect_change(widget)

        self.line_spacing_mode.currentTextChanged.connect(self._handle_line_spacing_mode_changed)
        self.numbering_style.currentIndexChanged.connect(self._handle_numbering_widgets_changed)
        self.renumber.toggled.connect(self._handle_numbering_widgets_changed)
        self._update_line_spacing_editor()
        self._update_numbering_widgets()

    def load_profile(self, profile: ParagraphFormat, heading_mode: bool, caption_mode: bool = False) -> None:
        self._profile = profile
        self._heading_mode = heading_mode
        self._caption_mode = caption_mode
        self._loading = True
        self.font_name.setCurrentFont(QFont(profile.font_name))
        self.font_size.setEditText(font_size_to_text(profile.font_size))
        self.bold.setChecked(profile.bold)
        self.italic.setChecked(profile.italic)
        self.alignment.setCurrentText(profile.alignment)
        self.left_indent.setValue(profile.left_indent)
        self.right_indent.setValue(profile.right_indent)
        self.first_line_indent.setValue(profile.first_line_indent)
        self.space_before.setValue(profile.space_before)
        self.space_after.setValue(profile.space_after)
        self.line_spacing_mode.setCurrentText(profile.line_spacing_mode)
        self._update_line_spacing_editor()
        self.line_spacing.setValue(profile.line_spacing)
        index = self.numbering_style.findData(profile.numbering_style)
        self.numbering_style.setCurrentIndex(index if index >= 0 else 0)
        self.custom_numbering_pattern.setText(profile.custom_numbering_pattern or "")
        self.renumber.setChecked(profile.renumber)
        self.clear_existing_numbering.setChecked(profile.clear_existing_numbering)
        self.continue_previous_numbering.setChecked(profile.continue_previous_numbering if heading_mode else False)
        self.space_after_numbering.setChecked(profile.space_after_numbering)
        self._update_numbering_widgets()
        self._loading = False

    def sync_to_profile(self) -> None:
        if not self._profile or self._loading:
            return
        self._profile.font_name = self.font_name.currentFont().family()
        self._profile.font_size = font_size_from_text(self.font_size.currentText(), fallback=self._profile.font_size)
        self._profile.bold = self.bold.isChecked()
        self._profile.italic = self.italic.isChecked()
        self._profile.alignment = self.alignment.currentText()
        if self._caption_mode and not self._heading_mode:
            self._profile.left_indent = 0.0
            self._profile.right_indent = 0.0
            self._profile.first_line_indent = 0.0
        else:
            self._profile.left_indent = self.left_indent.value()
            self._profile.right_indent = self.right_indent.value()
            self._profile.first_line_indent = self.first_line_indent.value()
        self._profile.space_before = self.space_before.value()
        self._profile.space_after = self.space_after.value()
        self._profile.line_spacing_mode = self.line_spacing_mode.currentText()
        self._profile.line_spacing = self.line_spacing.value()
        self._profile.numbering_style = self.numbering_style.currentData() or "none"
        self._profile.custom_numbering_pattern = self.custom_numbering_pattern.text().strip()
        self._profile.numbering_pattern = resolve_numbering_pattern(
            self._profile.numbering_style,
            self._profile.custom_numbering_pattern,
        )
        self._profile.renumber = self.renumber.isChecked()
        self._profile.clear_existing_numbering = self.clear_existing_numbering.isChecked()
        self._profile.continue_previous_numbering = self.continue_previous_numbering.isChecked() if self._heading_mode else False
        enabled = self._heading_mode or self._caption_mode
        self._profile.space_after_numbering = self.space_after_numbering.isChecked() if enabled else False
        if self._on_change:
            self._on_change()

    def _handle_line_spacing_mode_changed(self) -> None:
        self._update_line_spacing_editor()
        self.sync_to_profile()

    def _handle_numbering_widgets_changed(self) -> None:
        self._update_numbering_widgets()
        self.sync_to_profile()

    def _update_line_spacing_editor(self) -> None:
        mode = self.line_spacing_mode.currentText()
        current_value = self.line_spacing.value()
        with QSignalBlocker(self.line_spacing):
            self.line_spacing.setDecimals(2)
            if mode == "单倍":
                self.line_spacing.setRange(1.0, 1.0)
                self.line_spacing.setSingleStep(0.1)
                self.line_spacing.setSuffix(" 倍")
                self.line_spacing.setValue(1.0)
                self.line_spacing.setEnabled(False)
            elif mode == "1.5 倍":
                self.line_spacing.setRange(1.5, 1.5)
                self.line_spacing.setSingleStep(0.1)
                self.line_spacing.setSuffix(" 倍")
                self.line_spacing.setValue(1.5)
                self.line_spacing.setEnabled(False)
            elif mode == "2 倍":
                self.line_spacing.setRange(2.0, 2.0)
                self.line_spacing.setSingleStep(0.1)
                self.line_spacing.setSuffix(" 倍")
                self.line_spacing.setValue(2.0)
                self.line_spacing.setEnabled(False)
            elif mode == "多倍":
                self.line_spacing.setRange(0.5, 10.0)
                self.line_spacing.setSingleStep(0.1)
                self.line_spacing.setSuffix(" 倍")
                self.line_spacing.setEnabled(True)
                self.line_spacing.setValue(current_value if current_value > 0 else 1.5)
            else:
                self.line_spacing.setRange(1.0, 200.0)
                self.line_spacing.setSingleStep(0.5)
                self.line_spacing.setSuffix(" 磅")
                self.line_spacing.setEnabled(True)
                self.line_spacing.setValue(current_value if current_value > 0 else 12.0)

    def _update_numbering_widgets(self) -> None:
        enabled = self._heading_mode or self._caption_mode
        self.numbering_style.setEnabled(enabled)
        self.renumber.setEnabled(enabled)
        self.clear_existing_numbering.setEnabled(enabled)
        self.continue_previous_numbering.setEnabled(self._heading_mode)
        self.space_after_numbering.setEnabled(enabled)
        style = self.numbering_style.currentData() or "none"
        self.custom_numbering_pattern.setEnabled(enabled and style == "custom")

        is_caption = self._caption_mode and not self._heading_mode
        for label, widget in [
            (self.left_indent_label, self.left_indent),
            (self.right_indent_label, self.right_indent),
            (self.first_line_indent_label, self.first_line_indent),
        ]:
            label.setVisible(not is_caption)
            widget.setVisible(not is_caption)
        for label, widget in [
            (self.numbering_style_label, self.numbering_style),
            (self.custom_numbering_pattern_label, self.custom_numbering_pattern),
            (self.numbering_options_label, self.renumber.parentWidget()),
        ]:
            label.setVisible(enabled)
            widget.setVisible(enabled)

    def _connect_change(self, widget) -> None:
        if hasattr(widget, "currentFontChanged"):
            widget.currentFontChanged.connect(lambda *_: self.sync_to_profile())
        elif hasattr(widget, "currentTextChanged"):
            widget.currentTextChanged.connect(lambda *_: self.sync_to_profile())
        elif hasattr(widget, "valueChanged"):
            widget.valueChanged.connect(lambda *_: self.sync_to_profile())
        elif hasattr(widget, "toggled"):
            widget.toggled.connect(lambda *_: self.sync_to_profile())
        elif hasattr(widget, "textChanged"):
            widget.textChanged.connect(lambda *_: self.sync_to_profile())

    @staticmethod
    def _double_spin(minimum: float, maximum: float, step: float) -> QDoubleSpinBox:
        spin = QDoubleSpinBox()
        spin.setRange(minimum, maximum)
        spin.setDecimals(2)
        spin.setSingleStep(step)
        return spin

    @staticmethod
    def _combo(items) -> QComboBox:
        combo = QComboBox()
        combo.addItems(items)
        return combo

    @staticmethod
    def _row_widget(*widgets) -> QWidget:
        wrapper = QWidget()
        layout = QHBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        for widget in widgets:
            layout.addWidget(widget)
        layout.addStretch(1)
        return wrapper


class ImageEditor(QWidget):
    def __init__(self, on_change: Optional[Callable[[], None]] = None) -> None:
        super().__init__()
        self._profile: Optional[ImageFormat] = None
        self._loading = False
        self._on_change = on_change

        layout = QFormLayout(self)
        layout.setLabelAlignment(Qt.AlignRight)
        self.alignment = QComboBox()
        self.alignment.addItems(TEXT_ALIGNMENTS)
        self.space_before = ParagraphEditor._double_spin(0, 1000, 0.1)
        self.space_after = ParagraphEditor._double_spin(0, 1000, 0.1)
        self.line_spacing_mode = ParagraphEditor._combo(LINE_SPACING_MODES)
        self.line_spacing = ParagraphEditor._double_spin(0.5, 200, 0.1)
        self.wrap_type = QComboBox()
        self.wrap_type.addItems(WRAP_TYPES)

        layout.addRow("对齐", self.alignment)
        layout.addRow("段前", self.space_before)
        layout.addRow("段后", self.space_after)
        layout.addRow("行距方式", self.line_spacing_mode)
        layout.addRow("行距值", self.line_spacing)
        layout.addRow("环绕方式", self.wrap_type)

        for widget in [self.alignment, self.space_before, self.space_after, self.line_spacing, self.wrap_type]:
            if hasattr(widget, "currentTextChanged"):
                widget.currentTextChanged.connect(lambda *_: self.sync_to_profile())
            else:
                widget.valueChanged.connect(lambda *_: self.sync_to_profile())
        self.line_spacing_mode.currentTextChanged.connect(self._handle_line_spacing_mode_changed)
        self._update_line_spacing_editor()

    def load_profile(self, profile: ImageFormat) -> None:
        self._profile = profile
        self._loading = True
        self.alignment.setCurrentText(profile.alignment)
        self.space_before.setValue(profile.space_before)
        self.space_after.setValue(profile.space_after)
        self.line_spacing_mode.setCurrentText(profile.line_spacing_mode)
        self._update_line_spacing_editor()
        self.line_spacing.setValue(profile.line_spacing)
        self.wrap_type.setCurrentText(profile.wrap_type)
        self._loading = False

    def sync_to_profile(self) -> None:
        if not self._profile or self._loading:
            return
        self._profile.alignment = self.alignment.currentText()
        self._profile.space_before = self.space_before.value()
        self._profile.space_after = self.space_after.value()
        self._profile.line_spacing_mode = self.line_spacing_mode.currentText()
        self._profile.line_spacing = self.line_spacing.value()
        self._profile.wrap_type = self.wrap_type.currentText()
        if self._on_change:
            self._on_change()

    def _handle_line_spacing_mode_changed(self) -> None:
        self._update_line_spacing_editor()
        self.sync_to_profile()

    def _update_line_spacing_editor(self) -> None:
        mode = self.line_spacing_mode.currentText()
        current_value = self.line_spacing.value()
        with QSignalBlocker(self.line_spacing):
            self.line_spacing.setDecimals(2)
            if mode == "单倍":
                self.line_spacing.setRange(1.0, 1.0)
                self.line_spacing.setSingleStep(0.1)
                self.line_spacing.setSuffix(" 倍")
                self.line_spacing.setValue(1.0)
                self.line_spacing.setEnabled(False)
            elif mode == "1.5 倍":
                self.line_spacing.setRange(1.5, 1.5)
                self.line_spacing.setSingleStep(0.1)
                self.line_spacing.setSuffix(" 倍")
                self.line_spacing.setValue(1.5)
                self.line_spacing.setEnabled(False)
            elif mode == "2 倍":
                self.line_spacing.setRange(2.0, 2.0)
                self.line_spacing.setSingleStep(0.1)
                self.line_spacing.setSuffix(" 倍")
                self.line_spacing.setValue(2.0)
                self.line_spacing.setEnabled(False)
            elif mode == "多倍":
                self.line_spacing.setRange(0.5, 10.0)
                self.line_spacing.setSingleStep(0.1)
                self.line_spacing.setSuffix(" 倍")
                self.line_spacing.setEnabled(True)
                self.line_spacing.setValue(current_value if current_value > 0 else 1.0)
            else:
                self.line_spacing.setRange(1.0, 200.0)
                self.line_spacing.setSingleStep(0.5)
                self.line_spacing.setSuffix(" 磅")
                self.line_spacing.setEnabled(True)
                self.line_spacing.setValue(current_value if current_value > 0 else 12.0)


class TableEditor(QWidget):
    def __init__(self, on_change: Optional[Callable[[], None]] = None) -> None:
        super().__init__()
        self._profile: Optional[TableFormat] = None
        self._loading = False
        self._on_change = on_change

        layout = QFormLayout(self)
        layout.setLabelAlignment(Qt.AlignRight)
        self.alignment = QComboBox()
        self.alignment.addItems(TABLE_ALIGNMENTS)
        self.cell_vertical_alignment = QComboBox()
        self.cell_vertical_alignment.addItems(VERTICAL_ALIGNMENTS)
        self.top_border = QCheckBox("顶线")
        self.header_bottom_border = QCheckBox("表头下横线")
        self.bottom_border = QCheckBox("底线")
        self.inside_horizontal_border = QCheckBox("内部横线")
        self.inside_vertical_border = QCheckBox("内部竖线")
        self.three_line_status = QComboBox()
        self.three_line_status.addItems(THREE_LINE_OPTIONS)

        layout.addRow("表格对齐", self.alignment)
        layout.addRow("垂直对齐", self.cell_vertical_alignment)
        layout.addRow("边框", ParagraphEditor._row_widget(self.top_border, self.header_bottom_border, self.bottom_border))
        layout.addRow("内部线", ParagraphEditor._row_widget(self.inside_horizontal_border, self.inside_vertical_border))
        layout.addRow("三线表判断", self.three_line_status)

        for widget in [
            self.alignment,
            self.cell_vertical_alignment,
            self.top_border,
            self.header_bottom_border,
            self.bottom_border,
            self.inside_horizontal_border,
            self.inside_vertical_border,
            self.three_line_status,
        ]:
            if hasattr(widget, "currentTextChanged"):
                widget.currentTextChanged.connect(lambda *_: self.sync_to_profile())
            elif hasattr(widget, "valueChanged"):
                widget.valueChanged.connect(lambda *_: self.sync_to_profile())
            else:
                widget.toggled.connect(lambda *_: self.sync_to_profile())

    def load_profile(self, profile: TableFormat) -> None:
        self._profile = profile
        self._loading = True
        self.alignment.setCurrentText(profile.alignment)
        self.cell_vertical_alignment.setCurrentText(profile.cell_vertical_alignment)
        self.top_border.setChecked(profile.top_border)
        self.header_bottom_border.setChecked(profile.header_bottom_border)
        self.bottom_border.setChecked(profile.bottom_border)
        self.inside_horizontal_border.setChecked(profile.inside_horizontal_border)
        self.inside_vertical_border.setChecked(profile.inside_vertical_border)
        self.three_line_status.setCurrentText(profile.three_line_status)
        self._loading = False

    def sync_to_profile(self) -> None:
        if not self._profile or self._loading:
            return
        self._profile.alignment = self.alignment.currentText()
        self._profile.cell_vertical_alignment = self.cell_vertical_alignment.currentText()
        self._profile.top_border = self.top_border.isChecked()
        self._profile.header_bottom_border = self.header_bottom_border.isChecked()
        self._profile.bottom_border = self.bottom_border.isChecked()
        self._profile.inside_horizontal_border = self.inside_horizontal_border.isChecked()
        self._profile.inside_vertical_border = self.inside_vertical_border.isChecked()
        self._profile.three_line_status = self.three_line_status.currentText()
        if self._on_change:
            self._on_change()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Word 格式识别与套用工具")
        self.catalog = FormatCatalog()
        self.outline: List[dict] = []
        self.current_category = "正文"
        self._loading_outline = False
        self._outline_target_signature: Optional[tuple[str, int, int]] = None
        self._last_predicted_output_path = ""
        self._page_count_cache: dict[tuple[str, int, int], int] = {}
        self._outline_preview_override: Optional[List[dict]] = None
        self._outline_search_matches: List[int] = []
        self._outline_search_index = -1
        self._wublub_click_count = 0

        central = QWidget()
        self.setCentralWidget(central)
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        central_layout.addWidget(self.scroll_area)

        container = QWidget()
        self.scroll_area.setWidget(container)
        root_layout = QVBoxLayout(container)

        root_layout.addWidget(self._build_file_group())
        root_layout.addWidget(self._build_mode_group())
        root_layout.addWidget(self._build_center_area(), stretch=1)
        root_layout.addWidget(self._build_bottom_area(), stretch=1)

        self._apply_initial_window_size()
        self._refresh_category_list()
        self._load_category_into_editor(self.current_category)
        self._refresh_output_hint()
        self._reload_outline_table()
        self._refresh_preview()

    def _build_file_group(self) -> QGroupBox:
        group = QGroupBox("文档设置")
        layout = QGridLayout(group)

        self.template_path = FileDropLineEdit("可拖入模板 Word 文档")
        self.target_path = FileDropLineEdit("可拖入目标 Word 文档")
        self.output_path = QLineEdit()
        self.output_hint = QLabel()
        self.output_hint.setWordWrap(True)
        self.output_hint.setStyleSheet("color: #666666;")
        self.credit_label = QLabel('感谢产品经理 <a href="wublub://thanks">wublub</a>，感谢大模型 gpt-5.4')
        self.credit_label.setOpenExternalLinks(False)
        self.credit_label.setTextFormat(Qt.RichText)
        self.credit_label.setTextInteractionFlags(Qt.TextBrowserInteraction)
        self.credit_label.setStyleSheet("color: #666666;")
        self.credit_label.linkActivated.connect(self._handle_wublub_link_clicked)

        self.template_path.fileDropped.connect(lambda path: self._append_logs([f"已拖入模板文档：{path}"]))
        self.template_path.fileDropped.connect(self._handle_template_file_changed)
        self.target_path.fileDropped.connect(lambda path: self._append_logs([f"已拖入目标文档：{path}"]))
        self.target_path.fileDropped.connect(self._handle_target_file_changed)
        self.template_path.textChanged.connect(self._handle_template_path_text_changed)
        self.target_path.textChanged.connect(self._handle_target_path_text_changed)
        self.target_path.editingFinished.connect(lambda: self._sync_outline_from_target_inputs(silent=True))
        self.output_path.setPlaceholderText("可在原文件名基础上修改输出文件名")
        self.output_path.textChanged.connect(self._refresh_output_hint)

        template_btn = QPushButton("选择模板文档")
        target_btn = QPushButton("选择目标文档")
        output_btn = QPushButton("选择输出文件")
        export_profile_btn = QPushButton("导出格式参数")
        import_profile_btn = QPushButton("导入格式参数")

        template_btn.clicked.connect(lambda: self._choose_word_file(self.template_path))
        target_btn.clicked.connect(lambda: self._choose_word_file(self.target_path))
        output_btn.clicked.connect(self._choose_output_path)
        export_profile_btn.clicked.connect(self.export_catalog)
        import_profile_btn.clicked.connect(self.import_catalog)

        layout.addWidget(QLabel("模板文档"), 0, 0)
        layout.addWidget(self.template_path, 0, 1)
        layout.addWidget(template_btn, 0, 2)
        layout.addWidget(QLabel("目标文档"), 1, 0)
        layout.addWidget(self.target_path, 1, 1)
        layout.addWidget(target_btn, 1, 2)
        layout.addWidget(QLabel("输出文件"), 2, 0)
        layout.addWidget(self.output_path, 2, 1)
        layout.addWidget(output_btn, 2, 2)
        layout.addWidget(QLabel("输出位置"), 3, 0)
        layout.addWidget(self.output_hint, 3, 1, 1, 2)
        layout.addWidget(QLabel("致谢"), 4, 0)
        layout.addWidget(self.credit_label, 4, 1, 1, 2)
        layout.addWidget(QLabel("参数配置"), 5, 0)
        layout.addWidget(ParagraphEditor._row_widget(export_profile_btn, import_profile_btn), 5, 1, 1, 2)
        return group

    def _build_mode_group(self) -> QGroupBox:
        group = QGroupBox("识别与应用")
        layout = QGridLayout(group)

        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["模板页抽取", "范围识别"])
        self.extract_start_label = QLabel("识别开始页")
        self.extract_end_label = QLabel("识别结束页")
        self.extract_start_page = QSpinBox()
        self.extract_end_page = QSpinBox()
        self.target_start_page = QSpinBox()
        self.target_end_page = QSpinBox()
        for spin in [self.extract_start_page, self.extract_end_page, self.target_start_page, self.target_end_page]:
            spin.setRange(1, 99999)
            spin.setValue(1)
        self.extract_btn = QPushButton("识别格式")
        self.apply_btn = QPushButton("应用到目标文档")
        self.mode_hint = QLabel()
        self.mode_hint.setWordWrap(True)
        self.mode_hint.setStyleSheet("color: #666666;")

        self.extract_btn.clicked.connect(self.extract_formats)
        self.apply_btn.clicked.connect(self.apply_formats)
        self.mode_combo.currentTextChanged.connect(self._toggle_extract_range)
        self.target_start_page.editingFinished.connect(lambda: self._sync_outline_from_target_inputs(silent=True))
        self.target_end_page.editingFinished.connect(lambda: self._sync_outline_from_target_inputs(silent=True))

        layout.addWidget(QLabel("识别模式"), 0, 0)
        layout.addWidget(self.mode_combo, 0, 1)
        layout.addWidget(self.extract_btn, 0, 4)
        layout.addWidget(QLabel("应用开始页"), 0, 2)
        layout.addWidget(self.target_start_page, 0, 3)
        layout.addWidget(QLabel("应用结束页"), 1, 2)
        layout.addWidget(self.target_end_page, 1, 3)
        layout.addWidget(self.apply_btn, 1, 4)
        layout.addWidget(self.extract_start_label, 1, 0)
        layout.addWidget(self.extract_start_page, 1, 1)
        layout.addWidget(self.extract_end_label, 2, 0)
        layout.addWidget(self.extract_end_page, 2, 1)
        layout.addWidget(self.mode_hint, 3, 0, 1, 5)
        self._toggle_extract_range(self.mode_combo.currentText())
        return group

    def _build_center_area(self) -> QWidget:
        splitter = QSplitter(Qt.Horizontal)

        category_panel = QWidget()
        category_layout = QVBoxLayout(category_panel)
        category_layout.setContentsMargins(0, 0, 0, 0)
        category_layout.addWidget(QLabel("当前分类"))

        self.category_list = QListWidget()
        self.category_list.setSelectionMode(QAbstractItemView.SingleSelection)
        self.category_list.currentTextChanged.connect(self._load_category_into_editor)
        category_layout.addWidget(self.category_list, stretch=1)

        self.category_choice = QComboBox()
        self.add_category_btn = QPushButton("添加分类")
        self.remove_category_btn = QPushButton("移除当前")
        self.move_category_up_btn = QPushButton("上移")
        self.move_category_down_btn = QPushButton("下移")
        self.show_all_categories_btn = QPushButton("全部显示")

        self.add_category_btn.clicked.connect(self._add_selected_category)
        self.remove_category_btn.clicked.connect(self._remove_current_category)
        self.move_category_up_btn.clicked.connect(lambda: self._move_current_category(-1))
        self.move_category_down_btn.clicked.connect(lambda: self._move_current_category(1))
        self.show_all_categories_btn.clicked.connect(self._show_all_categories)

        category_layout.addWidget(QLabel("可添加分类"))
        category_layout.addWidget(self.category_choice)
        category_layout.addWidget(self.add_category_btn)
        category_layout.addWidget(self.remove_category_btn)
        category_layout.addWidget(self.move_category_up_btn)
        category_layout.addWidget(self.move_category_down_btn)
        category_layout.addWidget(self.show_all_categories_btn)

        self.editor_stack = QStackedWidget()
        self.paragraph_editor = ParagraphEditor(on_change=self._refresh_preview)
        self.image_editor = ImageEditor(on_change=self._refresh_preview)
        self.table_editor = TableEditor(on_change=self._refresh_preview)
        self.editor_stack.addWidget(self.paragraph_editor)
        self.editor_stack.addWidget(self.image_editor)
        self.editor_stack.addWidget(self.table_editor)

        self.outline_table = OutlineTableWidget()
        self.outline_search_input = QLineEdit()
        self.outline_search_input.setPlaceholderText("搜索标题文本")
        self.outline_search_btn = QPushButton("定位下一个")
        self.outline_move_start = QSpinBox()
        self.outline_move_end = QSpinBox()
        self.outline_move_target = QSpinBox()
        for spin in [self.outline_move_start, self.outline_move_end, self.outline_move_target]:
            spin.setRange(1, 1)
            spin.setValue(1)
        self.outline_move_btn = QPushButton("移动到下方")
        self.outline_search_input.returnPressed.connect(self._find_next_outline_heading)
        self.outline_search_input.textChanged.connect(self._reset_outline_search_state)
        self.outline_search_btn.clicked.connect(self._find_next_outline_heading)
        self.outline_move_btn.clicked.connect(self._move_outline_rows_by_order)
        self.outline_table.setColumnCount(6)
        self.outline_table.setHorizontalHeaderLabels(["顺序", "页码", "分类", "级别", "文本", "编号预览"])
        self.outline_table.verticalHeader().setVisible(False)
        self.outline_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.outline_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.outline_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.outline_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.outline_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.outline_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.outline_table.cellChanged.connect(self._handle_outline_cell_changed)
        self.outline_table.rowsMoved.connect(self._handle_outline_rows_moved)
        self.outline_table.dragPreviewChanged.connect(self._handle_outline_drag_preview_changed)
        self.outline_table.dragPreviewCleared.connect(self._handle_outline_drag_preview_cleared)

        outline_panel = QWidget()
        outline_layout = QVBoxLayout(outline_panel)
        outline_layout.setContentsMargins(0, 0, 0, 0)
        outline_layout.addWidget(QLabel("标题搜索"))
        outline_layout.addWidget(ParagraphEditor._row_widget(self.outline_search_input, self.outline_search_btn))
        outline_layout.addWidget(QLabel("标题快速移动"))
        outline_layout.addWidget(
            ParagraphEditor._row_widget(
                QLabel("开始顺序"),
                self.outline_move_start,
                QLabel("结束顺序"),
                self.outline_move_end,
                QLabel("放到顺序"),
                self.outline_move_target,
                self.outline_move_btn,
            )
        )
        outline_layout.addWidget(self.outline_table, stretch=1)

        splitter.addWidget(self._wrap_group("分类", category_panel))
        splitter.addWidget(self._wrap_group("格式参数", self.editor_stack))
        splitter.addWidget(self._wrap_group("检查项", outline_panel))
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 1)
        return splitter

    def _build_bottom_area(self) -> QWidget:
        splitter = QSplitter(Qt.Horizontal)

        self.preview_edit = QPlainTextEdit()
        self.preview_edit.setReadOnly(True)

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)

        splitter.addWidget(self._wrap_group("识别/配置预览", self.preview_edit))
        splitter.addWidget(self._wrap_group("执行日志", self.log_edit))
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        return splitter

    def _load_category_into_editor(self, category: str) -> None:
        if not category:
            return
        self.current_category = category
        profile = self.catalog.get_category(category)
        if isinstance(profile, ParagraphFormat):
            self.editor_stack.setCurrentWidget(self.paragraph_editor)
            self.paragraph_editor.load_profile(
                profile,
                heading_mode=category.startswith("标题"),
                caption_mode=category in {"图注", "表注"},
            )
        elif isinstance(profile, ImageFormat):
            self.editor_stack.setCurrentWidget(self.image_editor)
            self.image_editor.load_profile(profile)
        elif isinstance(profile, TableFormat):
            self.editor_stack.setCurrentWidget(self.table_editor)
            self.table_editor.load_profile(profile)
        self._refresh_category_controls()
        self._refresh_preview()

    def _toggle_extract_range(self, mode: str) -> None:
        is_template_mode = mode == "模板页抽取"
        for widget in [self.extract_start_label, self.extract_start_page, self.extract_end_label, self.extract_end_page]:
            widget.setVisible(not is_template_mode)
            widget.setEnabled(not is_template_mode)
        self.mode_hint.setText(
            "模板页抽取：默认识别整份模板文档。"
            if is_template_mode
            else "范围识别：按模板文档指定开始页和结束页统计格式。"
        )

    def _choose_word_file(self, line_edit: QLineEdit) -> None:
        start = str(Path.home())
        path, _ = QFileDialog.getOpenFileName(self, "选择 Word 文档", start, "Word 文档 (*.doc;*.docx)")
        if path:
            line_edit.setText(path)
            if line_edit is self.target_path:
                self._handle_target_file_changed(path)

    def _choose_output_path(self) -> None:
        source = self.target_path.text().strip()
        default_path = self.output_path.text().strip() or (self._predict_output_path(source) if source else str(Path.home() / "输出文档.docx"))
        path, _ = QFileDialog.getSaveFileName(self, "选择输出文件", default_path, "Word 文档 (*.docx)")
        if not path:
            return
        target = Path(path)
        if target.suffix.lower() != ".docx":
            target = target.with_suffix(".docx")
        self.output_path.setText(str(target))

    def export_catalog(self) -> None:
        self._sync_current_editor()
        default_path = str(Path.home() / "word格式参数.json")
        path, _ = QFileDialog.getSaveFileName(self, "导出格式参数", default_path, "JSON 文件 (*.json)")
        if not path:
            return
        target = Path(path)
        if target.suffix.lower() != ".json":
            target = target.with_suffix(".json")
        try:
            target.write_text(json.dumps(self.catalog.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
            self._append_logs([f"已导出格式参数：{target}"])
            QMessageBox.information(self, "完成", f"格式参数已导出：\n{target}")
        except Exception as exc:
            self._error(f"导出格式参数失败：{exc}")

    def import_catalog(self) -> None:
        start = str(Path.home())
        path, _ = QFileDialog.getOpenFileName(self, "导入格式参数", start, "JSON 文件 (*.json)")
        if not path:
            return
        try:
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
            imported_catalog = FormatCatalog.from_dict(payload)
        except Exception as exc:
            self._error(f"导入格式参数失败：{exc}")
            return

        self.catalog = imported_catalog
        visible_names = self.catalog.ordered_visible_categories()
        self.current_category = self.current_category if self.current_category in visible_names else (visible_names[0] if visible_names else "正文")
        self._refresh_category_list()
        self._load_category_into_editor(self.current_category)
        self._refresh_preview()
        self._append_logs([f"已导入格式参数：{path}"])
        QMessageBox.information(self, "完成", f"格式参数已导入：\n{path}")

    def _apply_initial_window_size(self) -> None:
        screen = QApplication.primaryScreen()
        if screen is None:
            self.resize(1200, 800)
            return
        available = screen.availableGeometry()
        margin = 80
        width = min(1500, max(960, available.width() - margin), available.width())
        height = min(900, max(720, available.height() - margin), available.height())
        self.resize(width, height)

    def _handle_template_path_text_changed(self, path: str) -> None:
        self._invalidate_page_count_cache(path)

    def _handle_target_path_text_changed(self, path: str) -> None:
        self._invalidate_page_count_cache(path)
        self._sync_output_path_from_target(path)
        self._refresh_output_hint()

    def _handle_template_file_changed(self, path: str) -> None:
        self._invalidate_page_count_cache(path)

    def _handle_target_file_changed(self, path: str) -> None:
        self._invalidate_page_count_cache(path)
        self._sync_output_path_from_target(path)
        self._sync_target_page_range_to_document(path)
        self._sync_outline_from_target_inputs(silent=True, force=True)

    def _sync_target_page_range_to_document(self, path: str) -> None:
        total_pages = self._get_document_page_count(path)
        if total_pages is None or total_pages < 1:
            return
        with QSignalBlocker(self.target_start_page), QSignalBlocker(self.target_end_page):
            self.target_start_page.setValue(1)
            self.target_end_page.setValue(total_pages)

    def _sync_output_path_from_target(self, path: str) -> None:
        source = path.strip()
        if not source:
            self._last_predicted_output_path = ""
            if not self.output_path.text().strip():
                self.output_path.clear()
            return
        predicted = self._predict_output_path(source)
        current_output = self.output_path.text().strip()
        if not current_output or current_output == self._last_predicted_output_path:
            self.output_path.setText(predicted)
        self._last_predicted_output_path = predicted

    def _refresh_output_hint(self, *_args) -> None:
        source = self.target_path.text().strip()
        custom_output = self.output_path.text().strip()
        if custom_output:
            self.output_hint.setText(f"将输出到：{custom_output}")
        elif source:
            self.output_hint.setText(f"将自动输出到：{self._predict_output_path(source)}")
        else:
            self.output_hint.setText("将自动输出到目标文档同文件夹，文件名为 原文件名_格式化.docx")

    def _refresh_category_list(self) -> None:
        current = self.current_category or "正文"
        names = self.catalog.ordered_visible_categories()
        with QSignalBlocker(self.category_list):
            self.category_list.clear()
            for name in names:
                self.category_list.addItem(QListWidgetItem(name))
            target = current if current in names else (names[0] if names else "正文")
            for index in range(self.category_list.count()):
                if self.category_list.item(index).text() == target:
                    self.category_list.setCurrentRow(index)
                    break
        self._refresh_category_controls()
        self._load_category_into_editor(self.category_list.currentItem().text() if self.category_list.currentItem() else "正文")

    def _refresh_category_controls(self) -> None:
        visible_names = self.catalog.ordered_visible_categories()
        available_names = [name for name in self.catalog.ordered_categories() if name not in visible_names]
        current = self.current_category or "正文"
        with QSignalBlocker(self.category_choice):
            self.category_choice.clear()
            for name in available_names:
                self.category_choice.addItem(name)
        can_remove = current in visible_names and len(visible_names) > 1
        self.remove_category_btn.setEnabled(can_remove)
        current_index = visible_names.index(current) if current in visible_names else -1
        self.move_category_up_btn.setEnabled(current_index > 0)
        self.move_category_down_btn.setEnabled(0 <= current_index < len(visible_names) - 1)
        self.add_category_btn.setEnabled(self.category_choice.count() > 0)

    def _add_selected_category(self) -> None:
        name = self.category_choice.currentText().strip()
        if not name:
            return
        visible = self.catalog.ordered_visible_categories()
        if name in visible:
            return
        visible.append(name)
        self.catalog.set_visible_categories(visible)
        self.current_category = name
        self._refresh_category_list()

    def _remove_current_category(self) -> None:
        current = self.current_category or "正文"
        visible = self.catalog.ordered_visible_categories()
        if current not in visible or len(visible) <= 1:
            return
        visible.remove(current)
        self.catalog.set_visible_categories(visible)
        self.current_category = visible[0]
        self._refresh_category_list()

    def _move_current_category(self, offset: int) -> None:
        current = self.current_category or "正文"
        visible = self.catalog.ordered_visible_categories()
        if current not in visible:
            return
        index = visible.index(current)
        target = index + offset
        if target < 0 or target >= len(visible):
            return
        visible[index], visible[target] = visible[target], visible[index]
        self.catalog.set_category_order(visible)
        self.catalog.set_visible_categories(visible)
        self.current_category = current
        self._refresh_category_list()

    def _show_all_categories(self) -> None:
        names = self.catalog.ordered_categories()
        self.catalog.set_category_order(names)
        self.catalog.set_visible_categories(names)
        self.current_category = self.current_category if self.current_category in names else names[0]
        self._refresh_category_list()

    def _reload_outline_table(self, selected_rows: Optional[List[int]] = None) -> None:
        self._loading_outline = True
        rows_to_restore = self._normalize_outline_drag_rows(selected_rows if selected_rows is not None else self._selected_outline_rows())
        with QSignalBlocker(self.outline_table):
            self.outline_table.setRowCount(len(self.outline))
            for row, item in enumerate(self.outline):
                order_item = QTableWidgetItem(str(row + 1))
                order_item.setFlags(order_item.flags() & ~Qt.ItemIsEditable)
                page_item = QTableWidgetItem(str(item.get("page_number", 0)))
                page_item.setFlags(page_item.flags() & ~Qt.ItemIsEditable)
                category_item = QTableWidgetItem(str(item.get("category", "")))
                category_item.setFlags(category_item.flags() & ~Qt.ItemIsEditable)
                text_item = QTableWidgetItem(str(item.get("text", "")))
                preview_item = QTableWidgetItem("")
                preview_item.setFlags(preview_item.flags() & ~Qt.ItemIsEditable)

                level_combo = QComboBox()
                level_combo.addItem("-", 0)
                for value in range(1, 10):
                    level_combo.addItem(str(value), value)
                category = str(item.get("category", "") or "")
                level = int(item.get("level", 0) or 0) if category.startswith("标题") else 0
                level_combo.setCurrentIndex(level if 0 <= level <= 9 else 0)
                level_combo.setEnabled(category.startswith("标题"))
                level_combo.currentIndexChanged.connect(lambda _index, row=row: self._handle_outline_level_changed(row))

                self.outline_table.setItem(row, 0, order_item)
                self.outline_table.setItem(row, 1, page_item)
                self.outline_table.setItem(row, 2, category_item)
                self.outline_table.setCellWidget(row, 3, level_combo)
                self.outline_table.setItem(row, 4, text_item)
                self.outline_table.setItem(row, 5, preview_item)
        self._restore_outline_selection(rows_to_restore)
        self.outline_table.refresh_row_styles()
        self._sync_outline_move_controls()
        self._outline_preview_override = None
        self._loading_outline = False
        self._refresh_outline_preview()

    def _selected_outline_rows(self) -> List[int]:
        selection_model = self.outline_table.selectionModel()
        rows = sorted({index.row() for index in selection_model.selectedRows()}) if selection_model is not None else []
        if rows:
            return self._normalize_outline_drag_rows(rows)
        current_row = self.outline_table.currentRow()
        if 0 <= current_row < len(self.outline):
            return self._normalize_outline_drag_rows([current_row])
        return []

    def _sync_outline_move_controls(self) -> None:
        maximum = max(1, len(self.outline))
        has_outline = bool(self.outline)
        for spin in [self.outline_move_start, self.outline_move_end, self.outline_move_target]:
            with QSignalBlocker(spin):
                spin.setRange(1, maximum)
                if spin.value() > maximum:
                    spin.setValue(maximum)
        self.outline_move_btn.setEnabled(has_outline)

    def _reset_outline_search_state(self) -> None:
        self._outline_search_matches = []
        self._outline_search_index = -1

    def _restore_outline_selection(self, rows: List[int]) -> None:
        selection_model = self.outline_table.selectionModel()
        if selection_model is None:
            return
        normalized = self._normalize_outline_drag_rows(rows)
        with QSignalBlocker(self.outline_table):
            self.outline_table.clearSelection()
            if not normalized:
                return
            flags = QItemSelectionModel.Select | QItemSelectionModel.Rows
            for row in normalized:
                selection_model.select(self.outline_table.model().index(row, 0), flags)
            self.outline_table.setCurrentCell(normalized[0], 0, QItemSelectionModel.NoUpdate)

    def _find_next_outline_heading(self) -> None:
        keyword = self.outline_search_input.text().strip()
        if not keyword:
            self._outline_search_matches = []
            self._outline_search_index = -1
            return
        matches = [
            row
            for row, item in enumerate(self.outline)
            if str(item.get("category", "") or "").startswith("标题") and keyword in str(item.get("text", "") or "")
        ]
        if not matches:
            self._outline_search_matches = []
            self._outline_search_index = -1
            self._warn(f"未找到标题：{keyword}")
            return
        if matches != self._outline_search_matches:
            self._outline_search_matches = matches
            self._outline_search_index = 0
        else:
            self._outline_search_index = (self._outline_search_index + 1) % len(matches)
        row = self._outline_search_matches[self._outline_search_index]
        self._restore_outline_selection([row])
        self.outline_table.scrollToItem(self.outline_table.item(row, 4), QAbstractItemView.PositionAtCenter)

    def _move_outline_rows_by_order(self) -> None:
        if not self.outline:
            self._warn("当前没有可移动的检查项。")
            return
        start_order = self.outline_move_start.value()
        end_order = self.outline_move_end.value()
        target_order = self.outline_move_target.value()
        if end_order < start_order:
            self._warn("结束顺序不能小于开始顺序。")
            return
        max_order = len(self.outline)
        if start_order < 1 or end_order > max_order or target_order < 1 or target_order > max_order:
            self._warn("快速移动的顺序超出检查项范围。")
            return
        source_rows = list(range(start_order - 1, end_order))
        rows = self._normalize_outline_drag_rows(source_rows)
        if not rows:
            self._warn("未找到可移动的标题范围。")
            return
        if rows[0] <= target_order - 1 <= rows[-1]:
            self._warn("目标顺序不能落在待移动范围内。")
            return
        normalized_target = self._normalize_outline_drop_target(rows, target_order)
        if normalized_target is None:
            self._warn("当前目标位置无需移动。")
            return
        moved_rows = self._outline_rows_after_move(rows, normalized_target)
        self._apply_outline_reorder(rows, normalized_target)
        self._outline_preview_override = None
        self._reload_outline_table(selected_rows=moved_rows)
        self._restore_outline_selection(moved_rows)
        self.outline_table.scrollToItem(self.outline_table.item(moved_rows[0], 4), QAbstractItemView.PositionAtCenter)

    def _normalize_outline_drag_rows(self, rows: List[int]) -> List[int]:
        normalized = sorted({row for row in rows if 0 <= row < len(self.outline)})
        if not normalized:
            return []
        intervals: List[tuple[int, int]] = []
        for row in normalized:
            block_end = self._outline_block_end(row)
            intervals.append((row, block_end))
        merged: List[tuple[int, int]] = []
        for start, end in sorted(intervals):
            if not merged or start > merged[-1][1]:
                merged.append((start, end))
            else:
                merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        expanded: List[int] = []
        for start, end in merged:
            expanded.extend(range(start, end))
        return expanded

    def _outline_block_end(self, row: int) -> int:
        if row < 0 or row >= len(self.outline):
            return row
        item = self.outline[row]
        category = str(item.get("category", "") or "")
        if not category.startswith("标题"):
            return min(row + 1, len(self.outline))
        level = int(item.get("level", 0) or 0)
        if level <= 0:
            try:
                level = int(category.replace("标题", ""))
            except Exception:
                level = 1
        for index in range(row + 1, len(self.outline)):
            next_item = self.outline[index]
            next_category = str(next_item.get("category", "") or "")
            if not next_category.startswith("标题"):
                continue
            next_level = int(next_item.get("level", 0) or 0)
            if next_level <= 0:
                try:
                    next_level = int(next_category.replace("标题", ""))
                except Exception:
                    next_level = 1
            if next_level <= level:
                return index
        return len(self.outline)

    def _normalize_outline_drop_target(self, source_rows: List[int], insert_row: int) -> Optional[int]:
        rows = self._normalize_outline_drag_rows(source_rows)
        if not rows:
            return None
        normalized_insert = max(0, min(insert_row, len(self.outline)))
        if rows[0] <= normalized_insert <= rows[-1] + 1:
            return None
        removed_before_insert = sum(1 for row in rows if row < normalized_insert)
        target_row = normalized_insert - removed_before_insert
        max_target = len(self.outline) - len(rows)
        return max(0, min(target_row, max_target))

    @staticmethod
    def _outline_rows_after_move(rows: List[int], target_row: int) -> List[int]:
        return [target_row + offset for offset in range(len(rows))]

    def _handle_outline_rows_moved(self, source_rows: List[int], target_row: int) -> None:
        rows = self._normalize_outline_drag_rows(source_rows)
        normalized_target = self._normalize_outline_drop_target(rows, target_row)
        if self._loading_outline or not rows or normalized_target is None:
            return
        self._apply_outline_reorder(rows, normalized_target)
        self._outline_preview_override = None
        self._reload_outline_table(selected_rows=self._outline_rows_after_move(rows, normalized_target))
        self.outline_table.clear_drag_target_rows()

    def _handle_outline_drag_preview_changed(self, source_rows: List[int], target_row: int) -> None:
        rows = self._normalize_outline_drag_rows(source_rows)
        normalized_target = self._normalize_outline_drop_target(rows, target_row)
        if self._loading_outline or not rows:
            return
        if normalized_target is None:
            self._outline_preview_override = None
            self.outline_table.clear_drag_target_rows()
            self._refresh_outline_preview()
            return
        preview_outline = self._reordered_outline(rows, normalized_target)
        self._outline_preview_override = preview_outline
        moved_rows = self._outline_rows_after_move(rows, normalized_target)
        self.outline_table.set_drag_target_rows(moved_rows)
        self._refresh_outline_preview()

    def _handle_outline_drag_preview_cleared(self) -> None:
        if self._outline_preview_override is None:
            self.outline_table.clear_drag_target_rows()
            return
        self._outline_preview_override = None
        self.outline_table.clear_drag_target_rows()
        self._refresh_outline_preview()

    def _apply_outline_reorder(self, source_rows: List[int], target_row: int) -> None:
        self.outline = self._reordered_outline(source_rows, target_row)
        for index, item in enumerate(self.outline, start=1):
            item["order"] = index

    def _reordered_outline(self, source_rows: List[int], target_row: int) -> List[dict]:
        rows = self._normalize_outline_drag_rows(source_rows)
        reordered = [dict(item) for item in self.outline]
        if not rows:
            return reordered
        moved_set = set(rows)
        moved_items = [reordered[row] for row in rows]
        remaining = [item for index, item in enumerate(reordered) if index not in moved_set]
        insert_row = max(0, min(target_row, len(remaining)))
        reordered = remaining[:insert_row] + moved_items + remaining[insert_row:]
        for index, item in enumerate(reordered, start=1):
            item["order"] = index
        return reordered

    def _handle_outline_level_changed(self, row: int) -> None:
        if self._loading_outline or row >= len(self.outline):
            return
        item = self.outline[row]
        category = str(item.get("category", "") or "")
        if not category.startswith("标题"):
            return
        combo = self.outline_table.cellWidget(row, 3)
        if not isinstance(combo, QComboBox):
            return
        level = int(combo.currentData() or 0)
        level = max(1, min(level, 9))
        item["level"] = level
        item["category"] = f"标题{level}"
        category_item = self.outline_table.item(row, 2)
        if category_item is not None:
            category_item.setText(item["category"])
        self._refresh_outline_preview()

    def _handle_outline_cell_changed(self, row: int, column: int) -> None:
        if self._loading_outline or column != 4 or row >= len(self.outline):
            return
        item = self.outline_table.item(row, column)
        self.outline[row]["text"] = item.text().strip() if item else ""
        self._refresh_outline_preview()

    def _refresh_outline_preview(self) -> None:
        previews = self._build_outline_previews(self._outline_preview_override or self.outline)
        self._loading_outline = True
        with QSignalBlocker(self.outline_table):
            for row, preview in enumerate(previews):
                preview_item = self.outline_table.item(row, 5)
                if preview_item is None:
                    preview_item = QTableWidgetItem()
                    preview_item.setFlags(preview_item.flags() & ~Qt.ItemIsEditable)
                    self.outline_table.setItem(row, 5, preview_item)
                preview_item.setText(preview)
        self._loading_outline = False

    def _build_outline_previews(self, outline_items: Optional[List[dict]] = None) -> List[str]:
        outline = outline_items if outline_items is not None else self.outline
        heading_formatter = NumberingFormatter()
        figure_formatter = NumberingFormatter()
        table_formatter = NumberingFormatter()
        started_levels: set[int] = set()
        previews: List[str] = []
        for item in outline:
            category = str(item.get("category", "") or "")
            existing_prefix = str(item.get("original_prefix", "") or "").strip()
            if category.startswith("标题"):
                level = max(1, min(int(item.get("level", 1) or 1), 9))
                profile = self.catalog.headings.get(f"标题{level}")
                if profile is None:
                    previews.append("")
                    continue
                pattern = self._resolve_heading_pattern(profile)
                if level not in started_levels and not profile.continue_previous_numbering:
                    heading_formatter.clear_from_level(level)
                prefix = ""
                if profile.renumber and pattern:
                    prefix = heading_formatter.next_for_level(level, pattern)
                elif not profile.clear_existing_numbering:
                    prefix = existing_prefix
                    heading_formatter.clear_from_level(level)
                else:
                    heading_formatter.clear_from_level(level)
                started_levels.add(level)
                previews.append(prefix)
                continue

            if category == "图注":
                profile = self.catalog.figure_caption
                pattern = self._resolve_heading_pattern(profile)
                if profile.renumber and pattern:
                    previews.append(figure_formatter.next_for_level(1, pattern))
                elif not profile.clear_existing_numbering:
                    previews.append(existing_prefix)
                else:
                    previews.append("")
                continue

            if category == "表注":
                profile = self.catalog.table_caption
                pattern = self._resolve_heading_pattern(profile)
                if profile.renumber and pattern:
                    previews.append(table_formatter.next_for_level(1, pattern))
                elif not profile.clear_existing_numbering:
                    previews.append(existing_prefix)
                else:
                    previews.append("")
                continue

            previews.append("")
        return previews

    def _sync_outline_from_target_inputs(self, silent: bool = False, force: bool = False) -> bool:
        source = self.target_path.text().strip()
        start_page = self.target_start_page.value()
        end_page = self.target_end_page.value()
        if not source:
            self._outline_target_signature = None
            self.outline = []
            self._reload_outline_table()
            return False
        target = Path(source)
        if not target.exists() or not target.is_file() or target.suffix.lower() not in WORD_SUFFIXES:
            self._outline_target_signature = None
            self.outline = []
            self._reload_outline_table()
            if not silent:
                self._warn(f"目标文档不存在：{source}" if target.suffix.lower() in WORD_SUFFIXES else "目标文档必须是 .doc 或 .docx 文件。")
            return False
        if end_page < start_page:
            self._outline_target_signature = None
            self.outline = []
            self._reload_outline_table()
            if not silent:
                self._warn("应用页码范围的结束页不能小于开始页。")
            return False
        signature = (str(target), start_page, end_page)
        if not force and self._outline_target_signature == signature:
            return True
        extractor = FormatExtractor(WordComClient(visible=False))
        try:
            outline, logs = extractor.extract_outline_from_page_range(str(target), start_page, end_page)
            self.outline = outline
            self._outline_target_signature = signature
            self._reload_outline_table()
            if not silent:
                self._append_logs(logs)
            return True
        except WordAutomationError as exc:
            self._outline_target_signature = None
            self.outline = []
            self._reload_outline_table()
            if not silent:
                self._error(str(exc))
            return False
        except Exception as exc:
            self._outline_target_signature = None
            self.outline = []
            self._reload_outline_table()
            if not silent:
                self._error(f"加载目标文档检查项失败：{exc}")
            return False

    def extract_formats(self) -> None:
        path = self.template_path.text().strip()
        if not self._validate_source_path(path, "模板文档"):
            return
        if self.mode_combo.currentText() == "范围识别" and not self._validate_page_range(
            self.extract_start_page.value(),
            self.extract_end_page.value(),
            "识别页码范围",
            path,
        ):
            return

        self._sync_current_editor()
        extractor = FormatExtractor(WordComClient(visible=False))
        self._append_logs([f"开始识别模板文档：{path}"])
        try:
            result: ExtractResult
            if self.mode_combo.currentText() == "模板页抽取":
                result = extractor.extract_template_document(path)
            else:
                result = extractor.extract_from_page_range(path, self.extract_start_page.value(), self.extract_end_page.value())
            self.catalog = result.catalog
            self.outline = result.outline
            self._outline_target_signature = None
            self._refresh_category_list()
            self._load_category_into_editor(self.current_category)
            self._reload_outline_table()
            self._append_logs(result.logs)
            self._append_logs(["格式识别完成。"])
        except WordAutomationError as exc:
            self._error(str(exc))
        except Exception as exc:
            self._error(f"识别失败：{exc}")

    def apply_formats(self) -> None:
        source = self.target_path.text().strip()
        if not self._validate_source_path(source, "目标文档"):
            return
        if not self._validate_page_range(self.target_start_page.value(), self.target_end_page.value(), "应用页码范围", source):
            return

        self._sync_current_editor()
        if not self._sync_outline_from_target_inputs(silent=False):
            return
        output_path = self.output_path.text().strip() or self._predict_output_path(source)
        output_target = Path(output_path)
        if output_target.suffix.lower() != ".docx":
            output_target = output_target.with_suffix(".docx")
        try:
            output_target.parent.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            self._error(f"输出目录无效：{exc}")
            return
        applier = FormatApplier(WordComClient(visible=False))
        self._append_logs([f"开始应用到目标文档：{source}", f"输出文件：{output_target}"])
        try:
            result = applier.apply_to_document(
                source_path=source,
                catalog=self.catalog,
                start_page=self.target_start_page.value(),
                end_page=self.target_end_page.value(),
                outline=[dict(item) for item in self.outline] if self.outline else None,
                output_path=str(output_target),
            )
            self.output_path.setText(result.output_path)
            self._append_logs(result.logs)
            QMessageBox.information(self, "完成", f"已输出：\n{result.output_path}")
        except WordAutomationError as exc:
            self._error(str(exc))
        except Exception as exc:
            self._error(f"应用失败：{exc}")

    def _validate_source_path(self, path: str, label: str) -> bool:
        if not path:
            self._warn(f"请先选择{label}。")
            return False
        target = Path(path)
        if not target.exists() or not target.is_file():
            self._warn(f"{label}不存在：{path}")
            return False
        if target.suffix.lower() not in WORD_SUFFIXES:
            self._warn(f"{label}必须是 .doc 或 .docx 文件。")
            return False
        return True

    def _validate_page_range(self, start: int, end: int, label: str, source_path: str = "") -> bool:
        if end < start:
            self._warn(f"{label}的结束页不能小于开始页。")
            return False
        total_pages = self._get_document_page_count(source_path) if source_path else None
        if total_pages is not None:
            if start > total_pages:
                self._warn(f"{label}的开始页超出范围：{start}，文档总页数为 {total_pages}。")
                return False
            if end > total_pages:
                self._warn(f"{label}的结束页超出范围：{end}，文档总页数为 {total_pages}。")
                return False
        return True

    def _get_document_page_count(self, source_path: str) -> Optional[int]:
        path = source_path.strip()
        if not path:
            return None
        target = Path(path)
        if not target.exists() or not target.is_file() or target.suffix.lower() not in WORD_SUFFIXES:
            return None
        try:
            resolved = target.resolve()
            stat = resolved.stat()
        except Exception:
            return None
        cache_key = (str(resolved), int(stat.st_mtime_ns), int(stat.st_size))
        if cache_key in self._page_count_cache:
            return self._page_count_cache[cache_key]
        client = WordComClient(visible=False)
        try:
            with client as word_client:
                document = word_client.open_document(str(resolved), read_only=True)
                try:
                    page_count = word_client.get_page_count(document)
                finally:
                    word_client.close_document(document, save_changes=False)
        except Exception:
            return None
        self._page_count_cache[cache_key] = page_count
        stale_keys = [key for key in self._page_count_cache if key[0] == str(resolved) and key != cache_key]
        for key in stale_keys:
            self._page_count_cache.pop(key, None)
        return page_count

    def _invalidate_page_count_cache(self, source_path: str) -> None:
        path = source_path.strip()
        if not path:
            return
        try:
            resolved_path = str(Path(path).resolve())
        except Exception:
            return
        stale_keys = [key for key in self._page_count_cache if key[0] == resolved_path]
        for key in stale_keys:
            self._page_count_cache.pop(key, None)

    def _sync_current_editor(self) -> None:
        widget = self.editor_stack.currentWidget()
        if hasattr(widget, "sync_to_profile"):
            widget.sync_to_profile()
        self._refresh_preview()

    def _refresh_preview(self) -> None:
        self.preview_edit.setPlainText(json.dumps(self.catalog.to_dict(), ensure_ascii=False, indent=2))
        self._refresh_outline_preview()

    def _append_logs(self, messages) -> None:
        existing = self.log_edit.toPlainText().strip()
        text = "\n".join(messages)
        self.log_edit.setPlainText(f"{existing}\n{text}".strip())
        self._refresh_preview()

    def _warn(self, message: str) -> None:
        QMessageBox.warning(self, "提示", message)

    def _error(self, message: str) -> None:
        self._append_logs([message])
        QMessageBox.critical(self, "错误", message)

    def _handle_wublub_link_clicked(self, link: str) -> None:
        if link != "wublub://thanks":
            return
        self._wublub_click_count += 1
        if self._wublub_click_count >= 3:
            self._wublub_click_count = 0
            QMessageBox.information(self, "彩蛋", "J")

    @staticmethod
    def _predict_output_path(source_path: str) -> str:
        source = Path(source_path)
        return str(source.with_name(f"{source.stem}_格式化.docx"))

    @staticmethod
    def _resolve_heading_pattern(profile: ParagraphFormat) -> str:
        if profile.numbering_style == "custom":
            return (profile.custom_numbering_pattern or "").strip()
        if profile.numbering_style:
            return resolve_numbering_pattern(profile.numbering_style, profile.custom_numbering_pattern)
        return (profile.numbering_pattern or "").strip()

    @staticmethod
    def _wrap_group(title: str, widget: QWidget) -> QGroupBox:
        group = QGroupBox(title)
        layout = QVBoxLayout(group)
        layout.addWidget(widget)
        return group
