"""Main window for 错题集生成器 v3.0 - Dark minimalist UI"""

import os
import re
import shutil
import configparser
from datetime import datetime

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QFileDialog, QMessageBox
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon

from .styles import get_dark_theme, get_light_theme
from .widgets import DigitInput, DigitInputGroup

# Resolve icon paths relative to this file
_ICON_DIR = os.path.join(os.path.dirname(__file__), "icons")
_LOG_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "log")


class MainWindow(QMainWindow):
    """Main application window with dark minimalist theme."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("错题集生成器 v3.0")
        self.setFixedSize(520, 540)

        # Theme state
        self._is_dark = True

        # Initialize data
        self.problem_shots_dir = ""
        self.feedback_file = ""

        # Log buffer for current operation
        self._log_buffer = []

        # Load config
        self.config = configparser.ConfigParser()
        self.config.read("config.ini")

        self._setup_ui()
        self._setup_navigation()
        self._apply_theme()
        self._load_saved_config()

    def _setup_ui(self):
        """Setup the main UI layout."""
        central = QWidget()
        self.setCentralWidget(central)

        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(30, 20, 30, 30)
        main_layout.setSpacing(0)

        # Header
        header_layout = QHBoxLayout()

        title_label = QLabel("错题集生成器v3.0")
        title_label.setObjectName("titleLabel")
        header_layout.addWidget(title_label)

        header_layout.addStretch()

        # Restore backup button
        self.restore_btn = QPushButton()
        self.restore_btn.setObjectName("themeButton")
        self.restore_btn.setToolTip("恢复备份")
        self.restore_btn.clicked.connect(self._restore_backup)
        header_layout.addWidget(self.restore_btn)

        header_layout.addSpacing(8)

        # Theme toggle button
        self.theme_btn = QPushButton()
        self.theme_btn.setObjectName("themeButton")
        self.theme_btn.setToolTip("切换主题")
        self.theme_btn.clicked.connect(self._toggle_theme)
        header_layout.addWidget(self.theme_btn)

        header_layout.addSpacing(8)

        # Info button
        self.info_btn = QPushButton("i")
        self.info_btn.setObjectName("infoButton")
        self.info_btn.setToolTip("使用帮助")
        self.info_btn.clicked.connect(self._show_help)
        header_layout.addWidget(self.info_btn)

        main_layout.addLayout(header_layout)

        # Author
        author_label = QLabel("作者: Leonard")
        author_label.setObjectName("authorLabel")
        main_layout.addSpacing(2)
        main_layout.addWidget(author_label)

        main_layout.addSpacing(20)

        # Form layout using grid for alignment
        form_layout = QGridLayout()
        form_layout.setHorizontalSpacing(12)
        form_layout.setVerticalSpacing(24)
        form_layout.setColumnStretch(1, 1)
        form_layout.setColumnMinimumWidth(2, 50)

        row = 0

        # 科目
        subject_label = QLabel("科目:")
        form_layout.addWidget(subject_label, row, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.subject_combo = QComboBox()
        self.subject_combo.addItems(["数学", "物理", "化学", "生物"])
        form_layout.addWidget(self.subject_combo, row, 1)

        row += 1

        # 班级
        class_label = QLabel("班级:")
        form_layout.addWidget(class_label, row, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.class_input = QLineEdit()
        self.class_input.setPlaceholderText("请输入班级")
        form_layout.addWidget(self.class_input, row, 1)

        row += 1

        # 日期 + 周数
        date_label = QLabel("日期:")
        form_layout.addWidget(date_label, row, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        date_week_layout = QHBoxLayout()
        date_week_layout.setSpacing(8)

        self.date_inputs = [DigitInput() for _ in range(4)]
        self.date_group = DigitInputGroup(self.date_inputs)

        for inp in self.date_inputs:
            date_week_layout.addWidget(inp)

        date_week_layout.addSpacing(30)

        week_label = QLabel("周数:")
        date_week_layout.addWidget(week_label)

        self.week_inputs = [DigitInput() for _ in range(2)]
        self.week_group = DigitInputGroup(self.week_inputs)

        for inp in self.week_inputs:
            date_week_layout.addWidget(inp)

        date_week_layout.addStretch()

        form_layout.addLayout(date_week_layout, row, 1, 1, 2)

        row += 1

        # 截图文件夹
        folder_label = QLabel("截图文件夹:")
        form_layout.addWidget(folder_label, row, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.folder_input = QLineEdit()
        self.folder_input.setReadOnly(True)
        self.folder_input.setPlaceholderText("点击右侧按钮选择文件夹")
        form_layout.addWidget(self.folder_input, row, 1)

        self.folder_btn = QPushButton()
        self.folder_btn.setObjectName("browseButton")
        self.folder_btn.setIconSize(QSize(24, 24))
        self.folder_btn.clicked.connect(self._select_folder)
        form_layout.addWidget(self.folder_btn, row, 2)

        row += 1

        # 错题反馈表
        file_label = QLabel("错题反馈表:")
        form_layout.addWidget(file_label, row, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.file_input = QLineEdit()
        self.file_input.setReadOnly(True)
        self.file_input.setPlaceholderText("点击右侧按钮选择Excel文件")
        form_layout.addWidget(self.file_input, row, 1)

        self.file_btn = QPushButton()
        self.file_btn.setObjectName("browseButton")
        self.file_btn.setIconSize(QSize(24, 24))
        self.file_btn.clicked.connect(self._select_file)
        form_layout.addWidget(self.file_btn, row, 2)

        main_layout.addLayout(form_layout)

        main_layout.addSpacing(25)
        main_layout.addStretch()

        # Buttons
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

    def _setup_navigation(self):
        """Setup keyboard navigation."""
        self.class_input.returnPressed.connect(
            lambda: self.date_inputs[0].setFocus()
        )
        self.date_inputs[-1].digitEntered.connect(
            lambda: self.week_inputs[0].setFocus()
        )

    def _apply_theme(self):
        """Apply current theme to all widgets."""
        if self._is_dark:
            self.setStyleSheet(get_dark_theme())
            self.theme_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "moon.svg")))
            self.restore_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "undo.svg")))
            self.folder_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "folder.svg")))
            self.file_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "file.svg")))
        else:
            self.setStyleSheet(get_light_theme())
            self.theme_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "sun.svg")))
            self.restore_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "undo_light.svg")))
            self.folder_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "folder_light.svg")))
            self.file_btn.setIcon(QIcon(os.path.join(_ICON_DIR, "file_light.svg")))

        self.theme_btn.setIconSize(QSize(20, 20))
        self.restore_btn.setIconSize(QSize(20, 20))
        self.date_group.set_dark_mode(self._is_dark)
        self.week_group.set_dark_mode(self._is_dark)

    def _toggle_theme(self):
        """Toggle between dark and light theme."""
        self._is_dark = not self._is_dark
        self._apply_theme()
        if not self.config.has_section("ui"):
            self.config.add_section("ui")
        self.config.set("ui", "dark_mode", str(self._is_dark))
        self._write_config()

    def _load_saved_config(self):
        """Load saved configuration values."""
        if self.config.has_section("info"):
            if self.config.has_option("info", "class"):
                self.class_input.setText(self.config.get("info", "class"))
            if self.config.has_option("info", "date"):
                self.date_group.set_value(self.config.get("info", "date"))
            if self.config.has_option("info", "week"):
                self.week_group.set_value(self.config.get("info", "week"))
            if self.config.has_option("info", "subject"):
                subject = self.config.get("info", "subject")
                idx = self.subject_combo.findText(subject)
                if idx >= 0:
                    self.subject_combo.setCurrentIndex(idx)
            if self.config.has_option("info", "folder"):
                folder = self.config.get("info", "folder")
                if os.path.isdir(folder):
                    self.problem_shots_dir = folder
                    self.folder_input.setText(folder)
            if self.config.has_option("info", "file"):
                file_path = self.config.get("info", "file")
                if os.path.isfile(file_path):
                    self.feedback_file = file_path
                    self.file_input.setText(file_path)

        if self.config.has_section("ui"):
            if self.config.has_option("ui", "dark_mode"):
                self._is_dark = self.config.get("ui", "dark_mode") == "True"
                self._apply_theme()

    def _save_config(self):
        """Save current configuration (called on GO)."""
        if not self.config.has_section("info"):
            self.config.add_section("info")
        self.config.set("info", "class", self.class_input.text())
        self.config.set("info", "date", self.date_group.get_value())
        self.config.set("info", "week", self.week_group.get_value())
        self.config.set("info", "subject", self.subject_combo.currentText())
        self.config.set("info", "folder", self.problem_shots_dir)
        self.config.set("info", "file", self.feedback_file)
        self._write_config()

    def _write_config(self):
        """Write config to file."""
        with open("config.ini", "w") as f:
            self.config.write(f)

    def _select_folder(self):
        """Open folder selection dialog."""
        start_dir = self.problem_shots_dir if self.problem_shots_dir else ""
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择截图文件夹",
            start_dir,
            options=QFileDialog.Option.ShowDirsOnly
        )
        if folder:
            self.problem_shots_dir = folder
            self.folder_input.setText(folder)
            self._rename_single_digit_files(folder)

    def _rename_single_digit_files(self, folder: str):
        """Rename single-digit image files to two-digit format."""
        for f in os.listdir(folder):
            if f.endswith((".png", ".jpg")):
                name = f[:-4]
                if name.isdigit() and len(name) == 1:
                    new_name = f"0{name}{f[-4:]}"
                    os.rename(
                        os.path.join(folder, f),
                        os.path.join(folder, new_name)
                    )

    def _select_file(self):
        """Open file selection dialog."""
        start_dir = os.path.dirname(self.feedback_file) if self.feedback_file else ""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择错题反馈表",
            start_dir,
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.feedback_file = file_path
            self.file_input.setText(file_path)
            self._auto_fill_from_filename(file_path)

    def _auto_fill_from_filename(self, file_path: str):
        """Auto-fill date and subject from filename."""
        filename = os.path.basename(file_path)

        date_match = re.search(r'(\d{4})', filename)
        if date_match:
            self.date_group.set_value(date_match.group(1))

        subjects = ["数学", "物理", "化学", "生物"]
        for subject in subjects:
            if subject in filename:
                self.subject_combo.setCurrentText(subject)
                break

    def _reset(self):
        """Reset all input fields."""
        self.class_input.clear()
        self.subject_combo.setCurrentIndex(0)
        self.date_group.clear()
        self.week_group.clear()
        self.folder_input.clear()
        self.file_input.clear()
        self.problem_shots_dir = ""
        self.feedback_file = ""
        self._load_saved_config()

    def _restore_backup(self):
        """Restore from backup."""
        backup_dir = os.path.join(".", "backup")
        student_dir = os.path.join(".", "学生")

        if not os.path.exists(backup_dir):
            QMessageBox.warning(self, "警告", "备份不存在！\n这可能是因为您本学期第一次使用错题集生成器。")
            return

        reply = QMessageBox.question(
            self,
            "确认恢复",
            "确定要恢复到上次操作之前的状态吗？\n当前的学生文件夹将被覆盖。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                if os.path.exists(student_dir):
                    shutil.rmtree(student_dir)
                shutil.copytree(backup_dir, student_dir)
                shutil.rmtree(backup_dir)
                QMessageBox.information(self, "成功", "备份已恢复！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"恢复失败: {str(e)}")

    def _validate_inputs(self) -> bool:
        """Validate all inputs before generation."""
        if not self.class_input.text().strip():
            QMessageBox.warning(self, "警告", "请输入班级!")
            self.class_input.setFocus()
            return False

        if not self.date_group.is_complete():
            QMessageBox.warning(self, "警告", "请输入完整的日期 (MMDD)!")
            self.date_inputs[0].setFocus()
            return False

        if not self.week_group.is_complete():
            QMessageBox.warning(self, "警告", "请输入周数!")
            self.week_inputs[0].setFocus()
            return False

        if not self.problem_shots_dir:
            QMessageBox.warning(self, "警告", "请选择截图文件夹!")
            return False

        if not os.path.isdir(self.problem_shots_dir):
            QMessageBox.warning(self, "警告", "截图文件夹不存在!")
            return False

        if not self.feedback_file:
            QMessageBox.warning(self, "警告", "请选择错题反馈表!")
            return False

        if not os.path.isfile(self.feedback_file):
            QMessageBox.warning(self, "警告", "错题反馈表文件不存在!")
            return False

        return True

    def _log(self, text: str):
        """Add text to log buffer and print to console."""
        clean_text = re.sub(r'<[^>]+>', '', text)
        self._log_buffer.append(clean_text)
        print(clean_text)

    def _save_log(self):
        """Save log buffer to file."""
        if not self._log_buffer:
            return

        if not os.path.exists(_LOG_DIR):
            os.makedirs(_LOG_DIR)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(_LOG_DIR, f"generate_{timestamp}.log")

        with open(log_file, "w", encoding="utf-8") as f:
            f.write(f"错题集生成器 v3.0 - 生成日志\n")
            f.write(f"时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"班级: {self.class_input.text()}\n")
            f.write(f"科目: {self.subject_combo.currentText()}\n")
            f.write(f"日期: {self.date_group.get_value()}\n")
            f.write(f"周数: {self.week_group.get_value()}\n")
            f.write("-" * 50 + "\n\n")
            f.write("\n".join(self._log_buffer))

        self._log_buffer.clear()

    def _start_generate(self):
        """Start the generation process."""
        if not self._validate_inputs():
            return

        self._save_config()
        self._log_buffer.clear()

        class_name = self.class_input.text().strip()
        subject = self.subject_combo.currentText()
        date = self.date_group.get_value()
        week = self.week_group.get_value()

        try:
            from core import generate
            adapter = TextBrowserAdapter(self)
            generate(
                self.problem_shots_dir,
                self.feedback_file,
                class_name,
                subject,
                week,
                date,
                adapter
            )
            self._save_log()
            QMessageBox.information(self, "完成", "错题集已生成!")
        except Exception as e:
            self._log(f"错误: {str(e)}")
            self._save_log()
            QMessageBox.critical(self, "错误", f"生成失败: {str(e)}")

    def _show_help(self):
        """Show help dialog."""
        help_text = """<h3>使用方法：</h3>
<ol>
    <li>将讲义题目截图(PNG格式)按题号命名(如01, 08, 15)，放入同一文件夹</li>
    <li>准备错题反馈表Excel：第一列姓名，第二列错题号</li>
    <li>错题号格式：数字用逗号分隔，或填"全对"/"未交"/"未带"</li>
    <li>填写班级、科目、日期、周数</li>
    <li>选择截图文件夹和反馈表</li>
    <li>点击 GO! 开始生成</li>
</ol>
<p><b>恢复备份：</b>点击左上角↩按钮可恢复到上次操作前的状态</p>
<hr>
<p><b>关于：</b>初版由 Leonard 于 2022 年 @武汉天任教育 开发，Lrj 参与命令行初版开发。<br>
v3.0 是 2026 年 1 月 Leonard 翻新的第三版（第二个图形化版本）。<br>
感谢孙总、硕姐、万幸哥对我的肯定和支持！祝大家越过越开心！</p>
<p><small>项目地址: github.com/Leonardlsa/problems-folder-generator</small></p>
"""
        QMessageBox.information(self, "使用帮助", help_text)


class TextBrowserAdapter:
    """Adapter to make MainWindow compatible with legacy core.py."""

    def __init__(self, window: MainWindow):
        self.window = window
        self.textBrowser = self

    def append(self, text: str):
        """Append text to log."""
        self.window._log(text)


# Monkey-patch PyQt5 imports for core.py compatibility
import sys

class _QMessageBoxCompat:
    Yes = QMessageBox.StandardButton.Yes
    No = QMessageBox.StandardButton.No

    @staticmethod
    def warning(parent, title, text, buttons=None, default=None):
        if buttons is None:
            return QMessageBox.warning(parent, title, text)
        return QMessageBox.warning(parent, title, text, buttons, default)

class _FakeQtWidgets:
    QMessageBox = _QMessageBoxCompat

class _FakePyQt5:
    QtWidgets = _FakeQtWidgets()

if 'PyQt5' not in sys.modules:
    sys.modules['PyQt5'] = _FakePyQt5()
    sys.modules['PyQt5.QtWidgets'] = _FakeQtWidgets()
