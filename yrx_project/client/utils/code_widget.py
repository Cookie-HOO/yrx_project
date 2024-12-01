import sys
import typing

from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QDialog, QVBoxLayout, QPlainTextEdit
from PyQt5.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QFont
from PyQt5.QtCore import Qt, QRegExp


class PythonHighlighter(QSyntaxHighlighter):
    def __init__(self, document):
        super().__init__(document)

        # Define formats for different Python syntax elements
        self.keywordFormat = QTextCharFormat()
        self.keywordFormat.setForeground(QColor("blue"))
        self.keywordFormat.setFontWeight(QFont.Bold)

        self.stringFormat = QTextCharFormat()
        self.stringFormat.setForeground(QColor("magenta"))

        self.commentFormat = QTextCharFormat()
        self.commentFormat.setForeground(QColor("green"))
        self.commentFormat.setFontItalic(True)

        # Define Python keywords
        keywords = [
            'and', 'as', 'assert', 'break', 'class', 'continue', 'def', 'del', 'elif', 'else', 'except',
            'False', 'finally', 'for', 'from', 'global', 'if', 'import', 'in', 'is', 'lambda', 'None',
            'nonlocal', 'not', 'or', 'pass', 'raise', 'return', 'True', 'try', 'while', 'with', 'yield'
        ]

        # Create regex patterns for syntax highlighting
        self.highlightingRules = [(QRegExp(r'\b' + keyword + r'\b'), self.keywordFormat) for keyword in keywords]
        self.highlightingRules.append((QRegExp(r'".*?"'), self.stringFormat))
        self.highlightingRules.append((QRegExp(r"'.*?'"), self.stringFormat))
        self.highlightingRules.append((QRegExp(r'#.*'), self.commentFormat))

    def highlightBlock(self, text):
        for pattern, format in self.highlightingRules:
            expression = QRegExp(pattern)
            index = expression.indexIn(text)
            while index >= 0:
                length = expression.matchedLength()
                self.setFormat(index, length, format)
                index = expression.indexIn(text, index + length)


class CodeDialog(QDialog):
    def __init__(self, apply_func: typing.Callable, language="python", size=(400, 300), init_code=None):
        super().__init__()
        self.setWindowTitle(f"{language.capitalize()}代码")
        self.apply_func = apply_func
        self.resize(*size)

        layout = QVBoxLayout()
        self.setLayout(layout)

        self.editor = QPlainTextEdit()
        self.editor.setPlainText(init_code or "")
        self.button = QPushButton("应用代码", self)
        self.button.clicked.connect(self.apply_code)

        layout.addWidget(self.editor)
        layout.addWidget(self.button)

        if language == "python":
            # Set tab width (e.g., 4 spaces)
            self.editor.setTabStopWidth(4 * self.editor.fontMetrics().width(' '))
            # Apply syntax highlighter
            self.highlighter = PythonHighlighter(self.editor.document())

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Tab:
            cursor = self.editor.textCursor()
            cursor.insertText(" " * 4)  # Insert 4 spaces instead of a tab character
            event.accept()
        else:
            super().keyPressEvent(event)

    def apply_code(self):
        self.apply_func(self.editor.document().toPlainText())
        self.accept()  # Close the dialog


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Main Window")
        self.resize(300, 200)

        self.button = QPushButton("Open Code Editor", self)
        self.button.clicked.connect(self.open_code_editor)
        self.setCentralWidget(self.button)

    def open_code_editor(self):
        dialog = CodeDialog()
        dialog.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())