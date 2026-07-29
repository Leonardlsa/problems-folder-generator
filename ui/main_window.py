"""Main window for 错题集生成器 v3.1."""

from __future__ import annotations

import configparser
import os
import re
from datetime import datetime

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from version import APP_VERSION

from .styles import get_dark_theme, get_light_theme
from .widgets import DigitInput, DigitInputGroup

_ICON_DIR = os.path.join(os.path.dirname(__file__), "icons")
_LOG_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log")


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"错题集生成器 v{APP_VERSION}")
        self.setFixedSize(600, 640)

        self._is_dark = True
        self.student_images_dir = ""
        self.teacher_images_dir = ""
        self.feedback_file = ""
        self._log_buffer: list[str] = []

        self.config = configparser.ConfigParser()
        self.config.read("config.ini", encoding="utf-8")

        self._setup_ui()
        self._setup_navigation()
        self._apply_theme()
        self._load_saved_config()

    def _setup_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(30, 20, 30, 30)
        main_layout.setSpacing(0)

        header_layout = QHBoxLayout()
        title_label = QLabel(f"错题集生成器 v{APP_VERSION}")
        title_label.setObjectName("titleLabel")
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        self.restore_btn = QPushButton()
        self.restore_btn.setObjectName("themeButton")
        self.restore_btn.setToolTip("恢复备份")
        self.restore_btn.clicked.connect(self._restore_backup)
        header_layout.addWidget(self.restore_btn)
        header_layout.addSpacing(8)

        self.theme_btn = QPushButton()
        self.theme_btn.setObjectName("themeButton")
        self.theme_btn.setToolTip("切换主题")
        self.theme_btn.clicked.connect(self._toggle_theme)
        header_layout.addWidget(self.theme_btn)
        header_layout.addSpacing(8)

        self.info_btn = QPushButton("i")
        self.info_btn.setObjectName("infoButton")
        self.info_btn.setToolTip("使用帮助")
        self.info_btn.clicked.connect(self._show_help)
        header_layout.addWidget(self.info_btn)
        main_layout.addLayout(header_layout)

        author_label = QLabel("作者: Leonard")
        author_label.setObjectName("authorLabel")
        main_layout.addSpacing(2)
        main_layout.addWidget(author_label)
        main_layout.addSpacing(20)

        form_layout = QGridLayout()
        form_layout.setHorizontalSpacing(12)
        form_layout.setVerticalSpacing(20)
        form_layout.setColumnStretch(1, 1)
        form_layout.setColumnMinimumWidth(2, 50)
        row = 0

        subject_label = QLabel("科目:")
        form_layout.addWidget(
            subject_label,
            row,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )
        self.subject_combo = QComboBox()
        self.subject_combo.addItems(["数学", "物理", "化学", "生物"])
        form_layout.addWidget(self.subject_combo, row, 1)
        row += 1

        class_label = QLabel("班级:")
        form_layout.addWidget(
            class_label,
            row,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )
        self.class_input = QLineEdit()
        self.class_input.setPlaceholderText("请输入班级")
        form_layout.addWidget(self.class_input, row, 1)
        row += 1

        date_label = QLabel("日期:")
        form_layout.addWidget(
            date_label,
            row,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )
        date_week_layout = QHBoxLayout()
        date_week_layout.setSpacing(8)
        self.date_inputs = [DigitInput() for _ in range(4)]
        self.date_group = DigitInputGroup(self.date_inputs)
        for item in self.date_inputs:
            date_week_layout.addWidget(item)
        date_week_layout.addSpacing(30)
        date_week_layout.addWidget(QLabel("周数:"))
        self.week_inputs = [DigitInput() for _ in range(2)]
        self.week_group = DigitInputGroup(self.week_inputs)
        for item in self.week_inputs:
            date_week_layout.addWidget(item)
        date_week_layout.addStretch()
        form_layout.addLayout(date_week_layout, row, 1, 1, 2)
        row += 1

        self.student_folder_input, self.student_folder_btn = self._add_folder_row(
            form_layout,
            row,
            "学生版图片:",
            self._select_student_folder,
        )
        row += 1
        self.teacher_folder_input, self.teacher_folder_btn = self._add_folder_row(
            form_layout,
            row,
            "教师版图片:",
            self._select_teacher_folder,
        )
        row += 1

        file_label = QLabel("错题反馈表:")
        form_layout.addWidget(
            file_label,
            row,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )
        self.file_input = QLineEdit()
        self.file_input.setReadOnly(True)
        self.file_input.setPlaceholderText("点击右侧按钮选择 Excel 文件")
        form_layout.addWidget(self.file_input, row, 1)
        self.file_btn = QPushButton()
        self.file_btn.setObjectName("browseButton")
        self.file_btn.setIconSize(QSize(24, 24))
        self.file_btn.clicked.connect(self._select_file)
        form_layout.addWidget(self.file_btn, row, 2)
        main_layout.addLayout(form_layout)

        main_layout.addSpacing(25)
        main_layout.addStretch()
        button_layout = QHBoxLayout()
        button_layout.setSpacing(40)
        button_layout.addStretch()
        self.reset_btn = QPushButton("重置")
        self.reset_btn.setObjectName("resetButton")
        self.reset_btn.clicked.connect(self._reset)
        button_layout.addWidget(self.reset_btn)
        self.go_btn = QPushButton("GO!")
        self.go_btn.setObjectName("goButton")
        self.go_btn.clicked.connect(self._start_generate)
        button_layout.addWidget(self.go_btn)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)

    def _add_folder_row(
        self,
        layout: QGridLayout,
        row: int,
        label_text: str,
        callback,
    ) -> tuple[QLineEdit, QPushButton]:
        label = QLabel(label_text)
        layout.addWidget(
            label,
            row,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )
        line_edit = QLineEdit()
        line_edit.setReadOnly(True)
        line_edit.setPlaceholderText("点击右侧按钮选择文件夹")
        layout.addWidget(line_edit, row, 1)
        button = QPushButton()
        button.setObjectName("browseButton")
        button.setIconSize(QSize(24, 24))
        button.clicked.connect(callback)
        layout.addWidget(button, row, 2)
        return line_edit, button

    def _setup_navigation(self) -> None:
        self.class_input.returnPressed.connect(
            lambda: self.date_inputs[0].setFocus()
        )
        self.date_inputs[-1].digitEntered.connect(
            lambda: self.week_inputs[0].setFocus()
        )

    def _apply_theme(self) -> None:
        if self._is_dark:
            self.setStyleSheet(get_dark_theme())
            self.theme_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "moon.svg")))
            self.restore_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "undo.svg")))
            folder_icon = QIcon(os.path.join(_ICON_DIR, "folder.svg"))
            self.file_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "file.svg")))
        else:
            self.setStyleSheet(get_light_theme())
            self.theme_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "sun.svg")))
            self.restore_btn.setIcon(
                QIcon(os.path.join(_ICON_DIR, "undo_light.svg"))
            )
            folder_icon = QIcon(os.path.join(_ICON_DIR, "folder_light.svg"))
            self.file_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "file_light.svg")))

        self.student_folder_btn.setIcon(folder_icon)
        self.teacher_folder_btn.setIcon(folder_icon)
        self.theme_btn.setIconSize(QSize(20, 20))
        self.restore_btn.setIconSize(QSize(20, 20))
        self.date_group.set_dark_mode(self._is_dark)
        self.week_group.set_dark_mode(self._is_dark)

    def _toggle_theme(self) -> None:
        self._is_dark = not self._is_dark
        self._apply_theme()
        if not self.config.has_section("ui"):
            self.config.add_section("ui")
        self.config.set("ui", "dark_mode", str(self._is_dark))
        self._write_config()

    def _load_saved_config(self) -> None:
        if self.config.has_section("info"):
            section = self.config["info"]
            self.class_input.setText(section.get("class", ""))
            self.date_group.set_value(section.get("date", ""))
            self.week_group.set_value(section.get("week", ""))
            subject = section.get("subject", "")
            index = self.subject_combo.findText(subject)
            if index >= 0:
                self.subject_combo.setCurrentIndex(index)

            # v3.0 used one `folder` key. Migrate it to the student edition.
            student_folder = section.get("student_folder", section.get("folder", ""))
            teacher_folder = section.get("teacher_folder", "")
            if os.path.isdir(student_folder):
                self.student_images_dir = student_folder
                self.student_folder_input.setText(student_folder)
            if os.path.isdir(teacher_folder):
                self.teacher_images_dir = teacher_folder
                self.teacher_folder_input.setText(teacher_folder)

            file_path = section.get("file", "")
            if os.path.isfile(file_path):
                self.feedback_file = file_path
                self.file_input.setText(file_path)

        if self.config.has_section("ui"):
            self._is_dark = self.config.getboolean(
                "ui", "dark_mode", fallback=True
            )
            self._apply_theme()

    def _save_config(self) -> None:
        if not self.config.has_section("info"):
            self.config.add_section("info")
        self.config.set("info", "class", self.class_input.text())
        self.config.set("info", "date", self.date_group.get_value())
        self.config.set("info", "week", self.week_group.get_value())
        self.config.set("info", "subject", self.subject_combo.currentText())
        self.config.set("info", "student_folder", self.student_images_dir)
        self.config.set("info", "teacher_folder", self.teacher_images_dir)
        self.config.set("info", "file", self.feedback_file)
        self._write_config()

    def _write_config(self) -> None:
        with open("config.ini", "w", encoding="utf-8") as file:
            self.config.write(file)

    def _select_student_folder(self) -> None:
        folder = self._select_folder("选择学生版图片文件夹", self.student_images_dir)
        if folder:
            self.student_images_dir = folder
            self.student_folder_input.setText(folder)

    def _select_teacher_folder(self) -> None:
        folder = self._select_folder("选择教师版图片文件夹", self.teacher_images_dir)
        if folder:
            self.teacher_images_dir = folder
            self.teacher_folder_input.setText(folder)

    def _select_folder(self, title: str, start_dir: str) -> str:
        return QFileDialog.getExistingDirectory(
            self,
            title,
            start_dir,
            options=QFileDialog.Option.ShowDirsOnly,
        )

    def _select_file(self) -> None:
        start_dir = os.path.dirname(self.feedback_file) if self.feedback_file else ""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择错题反馈表",
            start_dir,
            "Excel Files (*.xlsx)",
        )
        if file_path:
            self.feedback_file = file_path
            self.file_input.setText(file_path)
            self._auto_fill_from_filename(file_path)

    def _auto_fill_from_filename(self, file_path: str) -> None:
        filename = os.path.basename(file_path)
        date_match = re.search(r"(\d{4})", filename)
        if date_match:
            self.date_group.set_value(date_match.group(1))
        for subject in ("数学", "物理", "化学", "生物"):
            if subject in filename:
                self.subject_combo.setCurrentText(subject)
                break

    def _reset(self) -> None:
        self.class_input.clear()
        self.subject_combo.setCurrentIndex(0)
        self.date_group.clear()
        self.week_group.clear()
        self.student_folder_input.clear()
        self.teacher_folder_input.clear()
        self.file_input.clear()
        self.student_images_dir = ""
        self.teacher_images_dir = ""
        self.feedback_file = ""
        self._load_saved_config()

    def _restore_backup(self) -> None:
        if not os.path.exists("backup"):
            QMessageBox.warning(
                self,
                "警告",
                "备份不存在！\n这可能是因为您本学期第一次使用错题集生成器。",
            )
            return

        reply = QMessageBox.question(
            self,
            "确认恢复",
            "确定要恢复到上次生成之前的状态吗？\n"
            "当前学生目录中的学生版和教师版都会被还原。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        try:
            from core import restore_outputs

            restored = restore_outputs()
            QMessageBox.information(
                self, "成功", f"备份已恢复：{'、'.join(restored)}"
            )
        except Exception as exc:
            QMessageBox.critical(self, "错误", f"恢复失败: {exc}")

    def _validate_inputs(self) -> bool:
        checks = (
            (
                bool(self.class_input.text().strip()),
                "请输入班级!",
                self.class_input,
            ),
            (
                self.date_group.is_complete(),
                "请输入完整的日期 (MMDD)!",
                self.date_inputs[0],
            ),
            (
                self.week_group.is_complete(),
                "请输入周数!",
                self.week_inputs[0],
            ),
        )
        for valid, message, widget in checks:
            if not valid:
                QMessageBox.warning(self, "警告", message)
                widget.setFocus()
                return False

        folders = (
            (self.student_images_dir, "请选择学生版图片文件夹!"),
            (self.teacher_images_dir, "请选择教师版图片文件夹!"),
        )
        for folder, missing_message in folders:
            if not folder:
                QMessageBox.warning(self, "警告", missing_message)
                return False
            if not os.path.isdir(folder):
                QMessageBox.warning(self, "警告", f"图片文件夹不存在: {folder}")
                return False

        if not self.feedback_file:
            QMessageBox.warning(self, "警告", "请选择错题反馈表!")
            return False
        if not os.path.isfile(self.feedback_file):
            QMessageBox.warning(self, "警告", "错题反馈表文件不存在!")
            return False
        return True

    def _log(self, text: str) -> None:
        clean_text = re.sub(r"<[^>]+>", "", text)
        self._log_buffer.append(clean_text)
        print(clean_text)

    def _save_log(self) -> None:
        if not self._log_buffer:
            return
        os.makedirs(_LOG_DIR, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(_LOG_DIR, f"generate_{timestamp}.log")
        with open(log_file, "w", encoding="utf-8") as file:
            file.write(f"错题集生成器 v{APP_VERSION} - 生成日志\n")
            file.write(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            file.write(f"班级: {self.class_input.text()}\n")
            file.write(f"科目: {self.subject_combo.currentText()}\n")
            file.write(f"日期: {self.date_group.get_value()}\n")
            file.write(f"周数: {self.week_group.get_value()}\n")
            file.write(f"学生版图片: {self.student_images_dir}\n")
            file.write(f"教师版图片: {self.teacher_images_dir}\n")
            file.write("-" * 50 + "\n\n")
            file.write("\n".join(self._log_buffer))
        self._log_buffer.clear()

    def _start_generate(self) -> None:
        if not self._validate_inputs():
            return

        self._save_config()
        self._log_buffer.clear()
        try:
            from core import generate

            adapter = TextBrowserAdapter(self)
            generate(
                self.student_images_dir,
                self.teacher_images_dir,
                self.feedback_file,
                self.class_input.text().strip(),
                self.subject_combo.currentText(),
                self.week_group.get_value(),
                self.date_group.get_value(),
                adapter,
            )
            self._save_log()
            QMessageBox.information(
                self,
                "完成",
                "每名学生的学生版和教师版已同时生成！",
            )
        except Exception as exc:
            self._log(f"错误: {exc}")
            self._save_log()
            QMessageBox.critical(self, "错误", f"生成失败: {exc}")

    def _show_help(self) -> None:
        help_text = f"""<h3>使用方法：</h3>
<ol>
    <li>准备学生版和教师版两套图片，均按题号命名（如 01、08、15）</li>
    <li>两套目录必须包含相同题号；图片内容可以不同</li>
    <li>准备 .xlsx 反馈表：第一列姓名，第二列错题号</li>
    <li>错题号可用中英文逗号或空格分隔，也可填写“全对”/“未交”/“未带”</li>
    <li>填写班级、科目、日期、周数后点击 GO!</li>
</ol>
<p><b>一次生成两套文档：</b></p>
<pre>学生/张三/物理学生版/第3周错题集.docx
学生/张三/物理教师版/第3周错题集.docx</pre>
<p>两版的学生、错题号、日期和排版相同，仅插入图片来源不同；可直接打包学生姓名目录发给家长。</p>
<p><b>恢复备份：</b>点击左上角↩按钮，恢复上次生成前的完整学生目录。</p>
<hr>
<p><b>关于：</b>初版由 Leonard 于 2022 年 @武汉天任教育 开发，Lrj 参与命令行初版开发。<br>
v{APP_VERSION} 将旧版分别运行“科目学生版/科目教师版”的流程整合为一次生成。</p>
<p><small>项目地址: github.com/Leonardlsa/problems-folder-generator</small></p>
"""
        QMessageBox.information(self, "使用帮助", help_text)


class TextBrowserAdapter:
    """Adapter that forwards core logs and confirmations to the main window."""

    def __init__(self, window: MainWindow) -> None:
        self.window = window
        self.textBrowser = self

    def append(self, text: str) -> None:
        self.window._log(text)

    def confirm_missing_images(
        self, missing_by_student: dict[str, list[int]]
    ) -> bool:
        from core import _format_missing_message

        result = QMessageBox.warning(
            self.window,
            "缺少图片",
            _format_missing_message(missing_by_student),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return result == QMessageBox.StandardButton.Yes
