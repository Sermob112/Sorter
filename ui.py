import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QLineEdit, QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox
)
from PySide6.QtCore import Qt
from sorter import Sorter

class DuplicateChecker(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Сортировщик")
        self.setGeometry(100, 100, 400, 300)
        
        # Основной layout
        self.layout = QVBoxLayout()

        # Поле для первой папки
        self.label_folder1 = QLabel("Выберите папку с файлами, которую нужно сортировать:")
        self.button_folder1 = QPushButton("Выбрать папку")
        self.path_folder1 = QLineEdit()
        self.button_folder1.clicked.connect(self.select_folder1)

        # Поле для второй папки
        self.label_folder2 = QLabel("Выберите папку, в которую будут скопированы сортированные файлы:")
        self.button_folder2 = QPushButton("Выбрать папку")
        self.path_folder2 = QLineEdit()
        self.button_folder2.clicked.connect(self.select_folder2)

        # Поле для отображения количества дубликатов
        self.label_duplicates = QLabel("Количество дубликатов:")
        self.label_files_count = QLabel("Количество файлов:")
       
        # Добавляем элементы в интерфейс
        self.layout.addWidget(self.label_folder1)
        self.layout.addWidget(self.path_folder1)
        self.layout.addWidget(self.button_folder1)

        self.layout.addWidget(self.label_folder2)
        self.layout.addWidget(self.path_folder2)
        self.layout.addWidget(self.button_folder2)

        self.layout.addWidget(self.label_files_count)
        self.layout.addWidget(self.label_duplicates)


        # Нижний горизонтальный layout для кнопок
        self.button_layout = QHBoxLayout()

        # Кнопка "Сортировать"
        self.button_check = QPushButton("Проверить")
        self.button_check.clicked.connect(self.check_folders)

        # Кнопка "Сгенерировать XLSX"
        self.button_sort = QPushButton("Сортировать")
        self.button_sort.clicked.connect(self.generate_sort)

        # Добавляем кнопки в горизонтальный layout
        self.button_layout.addStretch()  # Добавляем растяжение для кнопок в конце
        self.button_layout.addWidget(self.button_check)
        # self.button_layout.addWidget(self.button_sort)

        # Добавляем основной layout и layout кнопок
        self.layout.addStretch()  # Добавляем растяжение между верхней частью и кнопками
        self.layout.addLayout(self.button_layout)

        self.setLayout(self.layout)

    def export_files_with_notification(self,text):

        # Отображаем сообщение об успешном сохранении
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("Внимание!")
        msg_box.setText(text)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()
    def select_folder1(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите первую папку")
        if folder:
            self.path_folder1.setText(folder)
            self.directory_path = folder

    def select_folder2(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите вторую папку")
        if folder:
            self.path_folder2.setText(folder)
            self.directory_path_for_sort = folder

    def check_folders(self):
        if not hasattr(self, 'directory_path') or not self.directory_path:
            self.export_files_with_notification("Выберите папку с файлами, которую нужно сортировать")
            return 
        self.sorter = Sorter(self.directory_path)
        self.label_files_count.setText(f"Количество файлов: {str(self.sorter.count_files())}")
        self.sorter.export_to_xlsx("Контрольная сумма файлов.xlsx")
        self.export_files_with_notification("Сгенерирован файл: 'Контрольная сумма файлов.xlsx' в корневой папке программы")
        self.label_duplicates.setText(f"Количество дубликатов: {str(self.sorter.duplicate_counter)}")
        self.button_layout.addWidget(self.button_sort)

        # self.duplicate_count.setText(str(self.sorter.count_files()))


    def generate_sort(self):
        if not hasattr(self, 'directory_path_for_sort') or not self.directory_path_for_sort:
            self.export_files_with_notification("Выберите папку, в которую будут перемещены отсортированные файлы")
            return 
        self.sorter.move_files_to_folders(self.directory_path_for_sort)
        self.export_files_with_notification("Готово")


