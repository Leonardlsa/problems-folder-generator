"""Custom widgets for 错题集生成器 v3.1"""

from PySide6.QtWidgets import QLineEdit
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QKeyEvent

from .styles import DIGIT_INPUT_STYLE_DARK, DIGIT_INPUT_STYLE_LIGHT


class DigitInput(QLineEdit):
    """Single digit input box with auto-advance functionality."""

    # Signal emitted when user presses backspace on empty input
    backspaceOnEmpty = Signal()
    # Signal emitted when a digit is entered (for auto-advance)
    digitEntered = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMaxLength(1)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._is_dark = True
        self.setStyleSheet(DIGIT_INPUT_STYLE_DARK)

        # Connect text changed to handle auto-advance
        self.textChanged.connect(self._on_text_changed)

    def set_dark_mode(self, is_dark: bool):
        """Set the theme mode."""
        self._is_dark = is_dark
        if is_dark:
            self.setStyleSheet(DIGIT_INPUT_STYLE_DARK)
        else:
            self.setStyleSheet(DIGIT_INPUT_STYLE_LIGHT)

    def _on_text_changed(self, text: str):
        """Handle text change - emit signal when digit entered."""
        if text and text.isdigit():
            self.digitEntered.emit()

    def keyPressEvent(self, event: QKeyEvent):
        """Handle key press events."""
        key = event.key()

        # Handle backspace on empty input
        if key == Qt.Key.Key_Backspace and not self.text():
            self.backspaceOnEmpty.emit()
            return

        # Only allow digits
        if event.text() and not event.text().isdigit():
            if key not in (Qt.Key.Key_Backspace, Qt.Key.Key_Delete,
                          Qt.Key.Key_Left, Qt.Key.Key_Right,
                          Qt.Key.Key_Tab, Qt.Key.Key_Backtab):
                return

        super().keyPressEvent(event)


class DigitInputGroup:
    """Manages a group of DigitInput widgets with auto-navigation."""

    def __init__(self, inputs: list[DigitInput]):
        self.inputs = inputs
        self._setup_navigation()

    def _setup_navigation(self):
        """Connect signals for auto-navigation between inputs."""
        for i, inp in enumerate(self.inputs):
            # Auto-advance to next input
            if i < len(self.inputs) - 1:
                inp.digitEntered.connect(
                    lambda idx=i: self.inputs[idx + 1].setFocus()
                )

            # Backspace to previous input
            if i > 0:
                inp.backspaceOnEmpty.connect(
                    lambda idx=i: self._focus_and_select(idx - 1)
                )

    def _focus_and_select(self, index: int):
        """Focus input at index and select its content."""
        self.inputs[index].setFocus()
        self.inputs[index].selectAll()

    def get_value(self) -> str:
        """Get concatenated value from all inputs."""
        return ''.join(inp.text() for inp in self.inputs)

    def set_value(self, value: str):
        """Set value across all inputs."""
        for i, inp in enumerate(self.inputs):
            if i < len(value):
                inp.setText(value[i])
            else:
                inp.clear()

    def clear(self):
        """Clear all inputs."""
        for inp in self.inputs:
            inp.clear()

    def is_complete(self) -> bool:
        """Check if all inputs have values."""
        return all(inp.text() for inp in self.inputs)

    def set_dark_mode(self, is_dark: bool):
        """Set theme mode for all inputs."""
        for inp in self.inputs:
            inp.set_dark_mode(is_dark)
