import sys
import zipfile
import os
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QProgressBar, QMessageBox
import win32com.client

def extract_zip(zip_file, extract_dir):
    zip_path = os.path.join(os.path.expanduser('~'), 'Downloads', zip_file)

    if not os.path.exists(zip_path) or not zipfile.is_zipfile(zip_path):
        return False

    os.makedirs(extract_dir, exist_ok=True)

    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        return True
    except Exception as e:
        return False


class ZipExtractorApp(QWidget):
    def __init__(self):
        super().__init__()

        self.successful_extractions = 0
        self.failed_extractions = 0

        self.setWindowTitle("ZIP Extractor")
        self.setGeometry(100, 100, 500, 250)

        icon_path = os.path.join(os.path.dirname(__file__), 'img', 'zip.png')
        self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout(self)

        zip_layout = QHBoxLayout()
        self.zip_label = QLabel("ZIP файлы:")
        self.zip_entry = QLineEdit()
        self.zip_entry.setPlaceholderText("Выберите ZIP файлы")
        self.zip_button = QPushButton("Обзор")
        self.zip_button.clicked.connect(self.browse_zip)

        zip_layout.addWidget(self.zip_label)
        zip_layout.addWidget(self.zip_entry)
        zip_layout.addWidget(self.zip_button)

        extract_layout = QHBoxLayout()
        self.extract_label = QLabel("Папка для извлечения:")
        self.extract_entry = QLineEdit()
        self.extract_entry.setPlaceholderText("Выберите папку")
        self.extract_button = QPushButton("Обзор")
        self.extract_button.clicked.connect(self.browse_extract_dir)

        extract_layout.addWidget(self.extract_label)
        extract_layout.addWidget(self.extract_entry)
        extract_layout.addWidget(self.extract_button)

        self.progress = QProgressBar(self)
        self.progress.setRange(0, 1)

        self.extract_button_main = QPushButton("Извлечь")
        self.extract_button_main.clicked.connect(self.extract_zip_gui)

        layout.addLayout(zip_layout)
        layout.addLayout(extract_layout)
        layout.addWidget(self.progress)
        layout.addWidget(self.extract_button_main)

    def browse_zip(self):
        filenames, _ = QFileDialog.getOpenFileNames(self, "Выберите ZIP файлы", os.path.expanduser('~') + '/Downloads', "ZIP files (*.zip)")
        if filenames:
            self.zip_entry.setText(', '.join([os.path.basename(f) for f in filenames]))

    def browse_extract_dir(self):
        dirname = QFileDialog.getExistingDirectory(self, "Выберите папку для извлечения", os.path.expanduser('~'))
        if dirname:
            self.extract_entry.setText(dirname)

    def extract_zip_gui(self):
        zip_files = self.zip_entry.text().split(', ')
        extract_dir = self.extract_entry.text()

        if not zip_files or not extract_dir:
            self.show_error("Пожалуйста, выберите ZIP-файлы и папку для извлечения.")
            return

        self.progress.setMaximum(len(zip_files))
        self.progress.setValue(0)

        for zip_file in zip_files:
            if extract_zip(zip_file, extract_dir):
                self.successful_extractions += 1
            else:
                self.failed_extractions += 1
            self.progress.setValue(self.progress.value() + 1)

        self.show_result()

    def show_result(self):
        result_text = f"Извлечено: {self.successful_extractions} архива(ов)\nНеудачных попыток: {self.failed_extractions}"
        self.show_info(result_text)

    def show_info(self, text):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setText(text)
        msg.setWindowTitle("Результаты")
        msg.exec()

    def show_error(self, text):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Critical)
        msg.setText(text)
        msg.setWindowTitle("Ошибка")
        msg.exec()

    def create_shortcut(self, target, shortcut_name, icon_path):
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_name)
        shortcut.TargetPath = target
        shortcut.WorkingDirectory = os.path.dirname(target)
        shortcut.IconLocation = icon_path
        shortcut.save()

    def create_app_shortcut(self):
        app_path = os.path.abspath(sys.argv[0])
        shortcut_path = os.path.join(os.path.expanduser("~"), "Desktop", "ZIP Extractor.lnk")
        icon_path = os.path.join(os.path.dirname(__file__), 'img', 'zip.ico')
        self.create_shortcut(app_path, shortcut_path, icon_path)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ZipExtractorApp()
    window.show()

    window.create_app_shortcut()

    sys.exit(app.exec())
