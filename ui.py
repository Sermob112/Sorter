import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QLineEdit, QVBoxLayout, QHBoxLayout, QFileDialog
)
from PySide6.QtCore import Qt
from sorter import Sorter

class DuplicateChecker(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Duplicate Checker")
        self.setGeometry(100, 100, 400, 300)
        
        # Основной layout
        self.layout = QVBoxLayout()

        # Поле для первой папки
        self.label_folder1 = QLabel("Выберите первую папку:")
        self.button_folder1 = QPushButton("Выбрать папку")
        self.path_folder1 = QLineEdit()
        self.button_folder1.clicked.connect(self.select_folder1)

        # Поле для второй папки
        self.label_folder2 = QLabel("Выберите вторую папку:")
        self.button_folder2 = QPushButton("Выбрать папку")
        self.path_folder2 = QLineEdit()
        self.button_folder2.clicked.connect(self.select_folder2)

        # Поле для отображения количества дубликатов
        self.label_duplicates = QLabel("Количество дубликатов:")
        self.duplicate_count = QLineEdit()
        self.duplicate_count.setReadOnly(True)

        # Добавляем элементы в интерфейс
        self.layout.addWidget(self.label_folder1)
        self.layout.addWidget(self.path_folder1)
        self.layout.addWidget(self.button_folder1)

        self.layout.addWidget(self.label_folder2)
        self.layout.addWidget(self.path_folder2)
        self.layout.addWidget(self.button_folder2)

        self.layout.addWidget(self.label_duplicates)
        self.layout.addWidget(self.duplicate_count)

        # Нижний горизонтальный layout для кнопок
        self.button_layout = QHBoxLayout()

        # Кнопка "Сортировать"
        self.button_sort = QPushButton("Сортировать")
        self.button_sort.clicked.connect(self.sort_folders)

        # Кнопка "Сгенерировать XLSX"
        self.button_generate_xlsx = QPushButton("Сгенерировать XLSX")
        self.button_generate_xlsx.clicked.connect(self.generate_xlsx)

        # Добавляем кнопки в горизонтальный layout
        self.button_layout.addStretch()  # Добавляем растяжение для кнопок в конце
        self.button_layout.addWidget(self.button_sort)
        self.button_layout.addWidget(self.button_generate_xlsx)

        # Добавляем основной layout и layout кнопок
        self.layout.addStretch()  # Добавляем растяжение между верхней частью и кнопками
        self.layout.addLayout(self.button_layout)

        self.setLayout(self.layout)

    def select_folder1(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите первую папку")
        if folder:
            self.path_folder1.setText(folder)
            self.directory_path = folder

    def select_folder2(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите вторую папку")
        if folder:
            self.path_folder2.setText(folder)

    def sort_folders(self):
        self.sorter = Sorter(self.directory_path)
        self.duplicate_count.setText(str(self.sorter.count_files()))


    def generate_xlsx(self):
        # Логика генерации XLSX
        print("Генерация XLSX...")


