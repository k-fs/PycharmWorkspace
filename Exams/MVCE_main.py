import sys
import time
from code import InteractiveConsole

from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot, QThread, QObject, pyqtSignal, QTimer
from PyQt5.QtGui import QTextOption, QTextCursor
from PyQt5.QtWidgets import QApplication

__author__ = "daegontaven"
__copyright__ = "daegontaven"
__license__ = "gpl3"


class ConsoleBuffer(QObject):
    excrete = pyqtSignal(str)

    def __init__(self, parent=None, minimum=0.050):
        super(ConsoleBuffer, self).__init__(parent)
        self.minimum = minimum
        self.last_time = time.monotonic() - minimum
        self.buffer = []
        self.timer = QTimer()
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self._excrete)

    @pyqtSlot(str)
    def consume(self, s):
        self.buffer.append(s)

        delta = time.monotonic() - self.last_time
        remaining = self.minimum - delta
        if remaining <= 0:
            self._excrete()
        elif not self.timer.isActive():
            self.timer.start(int(1000 * remaining))

    def _excrete(self):
        self.timer.stop()
        s = ''.join(self.buffer)
        if len(s):
            self.last_time = time.monotonic()
            self.excrete.emit(s)
        self.buffer = []


class ConsoleStream(QObject):
    """
    Custom StreamIO class that handles when send data
    to console_log
    """
    written = pyqtSignal(str)

    def __init__(self, parent=None):
        super(ConsoleStream, self).__init__(parent)

    def write(self, string):
        self.written.emit(string)


class PythonInterpreter(QObject, InteractiveConsole):
    """
    A reimplementation of the builtin InteractiveConsole to
    work with threads.
    """
    push_command = pyqtSignal(str)
    multi_line = pyqtSignal(bool)
    output = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, parent=None):
        super(PythonInterpreter, self).__init__(parent)
        self.locals = {}
        InteractiveConsole.__init__(self, self.locals)
        self.stream = ConsoleStream(self)
        self.push_command.connect(self.command)

    def write(self, string):
        self.error.emit(string)

    def runcode(self, code):
        """
        Overrides and captures stdout and stdin from
        InteractiveConsole.
        """
        sys.stdout = self.stream
        sys.stderr = self.stream
        sys.excepthook = sys.__excepthook__
        result = InteractiveConsole.runcode(self, code)
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        return result

    @pyqtSlot(str)
    def command(self, command):
        """
        :param command: line retrieved from console_input on
                        returnPressed
        """
        result = self.push(command)
        self.multi_line.emit(result)


class MainWindow:
    """
    The main GUI window. Opens maximized.
    """
    def __init__(self):

        self.ui = uic.loadUi("main.ui")
        self.ui.showMaximized()

        # Console Properties
        self.ui.console_log.document().setMaximumBlockCount(500)
        self.ui.console_log.setWordWrapMode(QTextOption.WrapAnywhere)

        self.ps1 = '>>>'
        self.ps2 = '...'
        self.ui.console_prompt.setText(self.ps1)

        # Controls
        self.cursor = self.ui.console_log.textCursor()
        self.scrollbar = self.ui.console_log.verticalScrollBar()

        # Spawn Interpreter
        self.thread = QThread()
        self.thread.start()

        self.buffer = ConsoleBuffer()

        self.interpreter = PythonInterpreter()
        self.interpreter.moveToThread(self.thread)

        # Interpreter Signals
        self.ui.console_input.returnPressed.connect(self.send_console_input)
        self.interpreter.stream.written.connect(self.buffer.consume)
        self.buffer.excrete.connect(self.send_console_log)
        self.interpreter.error.connect(self.send_console_log)
        self.interpreter.multi_line.connect(self.prompt)

    def prompt(self, multi_line):
        """
        Sets what prompt to use.
        """
        if multi_line:
            self.ui.console_prompt.setText(self.ps2)
        else:
            self.ui.console_prompt.setText(self.ps1)

    def send_console_input(self):
        """
        Send input grabbed from the QLineEdit prompt to the console.
        """
        command = self.ui.console_input.text()
        self.ui.console_input.clear()
        self.interpreter.push_command.emit(str(command))

    def send_console_log(self, output):
        """
        Set the output from InteractiveConsole in the QTextEdit.
        Auto scroll scrollbar.
        """
        # Move cursor
        self.cursor.movePosition(QTextCursor.End)
        self.ui.console_log.setTextCursor(self.cursor)

        # Insert text
        self.ui.console_log.insertPlainText(output)

        # Move scrollbar
        self.scrollbar.setValue(self.scrollbar.maximum())


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()