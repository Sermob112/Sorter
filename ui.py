import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QLineEdit, QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox,QCheckBox
)
from PySide6.QtCore import Qt,Signal,Slot
from sorter import Sorter
from excel import ExcelGenerator
from PySide6.QtCore import QThread
from PySide6.QtWidgets import QApplication
class DuplicateChecker(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Сортировщик")
        self.setGeometry(100, 100, 400, 300)
        
        self.files_moved_count = 0
        self.layout = QVBoxLayout()

   
        self.label_folder1 = QLabel("Выберите папку с файлами, которую нужно сортировать:")
        self.button_folder1 = QPushButton("Выбрать папку")
        self.path_folder1 = QLineEdit()
        self.button_folder1.clicked.connect(self.select_folder1)

        self.label_folder2 = QLabel("Выберите папку, в которую будут скопированы сортированные файлы:")
        self.button_folder2 = QPushButton("Выбрать папку")
        self.path_folder2 = QLineEdit()
        self.button_folder2.clicked.connect(self.select_folder2)

        self.label_duplicates = QLabel("Количество дубликатов:")
        self.label_files_count = QLabel("Количество файлов:")
        self.checkboxMove = QCheckBox("Переместить файлы.")
        self.checkboxStay = QCheckBox("Переместить\Копировать файлы если соответсвуют формату.")
        self.label_files_moved_count = QLabel("Файлов перемещено: 0")
 
        self.layout.addWidget(self.label_folder1)
        self.layout.addWidget(self.path_folder1)
        self.layout.addWidget(self.button_folder1)

        self.layout.addWidget(self.label_folder2)
        self.layout.addWidget(self.path_folder2)
        self.layout.addWidget(self.button_folder2)

        self.layout.addWidget(self.label_files_count)
        self.layout.addWidget(self.label_duplicates)
        self.layout.addWidget(self.label_files_moved_count)
        self.layout.addWidget(self.checkboxMove)
        self.layout.addWidget(self.checkboxStay)

        self.button_layout = QHBoxLayout()


        self.button_check = QPushButton("Проверить")
        self.button_check.clicked.connect(self.check_folders)


        self.button_sort = QPushButton("Сортировать")
        self.button_sort.clicked.connect(self.generate_sort)
        self.button_report = QPushButton("Сгенерировать отчет")
        self.button_report.clicked.connect(self.generate_report)

        self.button_layout.addStretch()  
        self.button_layout.addWidget(self.button_check)

        self.layout.addStretch() 
        self.layout.addLayout(self.button_layout)


        self.setLayout(self.layout)

    def export_files_with_notification(self,text):


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
        self.excel_generator = ExcelGenerator(self.directory_path)
        self.label_files_count.setText(f"Количество файлов: {str(self.excel_generator.count_files())}")
        self.excel_generator.export_to_xlsx("Контрольная сумма файлов.xlsx")
        self.export_files_with_notification("Сгенерирован файл: 'Контрольная сумма файлов.xlsx' в корневой папке программы")
        self.label_duplicates.setText(f"Количество дубликатов: {str(self.excel_generator.duplicate_counter)}")
        self.button_layout.addWidget(self.button_sort)

        # self.duplicate_count.setText(str(self.sorter.count_files()))


    def generate_sort(self):
        self.sorter = Sorter(self.directory_path)
        
        if not hasattr(self, 'directory_path_for_sort') or not self.directory_path_for_sort:
            self.export_files_with_notification("Выберите папку, в которую будут перемещены отсортированные файлы")
            return 
         # Создаем поток
        self.thread = QThread()
        self.sorter.moveToThread(self.thread)
        # status =  self.checkboxMove.isChecked()
        # statusS =  self.checkboxStay.isChecked()
        # self.sorter.move_files_to_folders(self.directory_path_for_sort,status,statusS,self.update_files_moved_count)
        # self.button_layout.addWidget(self.button_report)
        # self.export_files_with_notification("Готово")
        self.sorter.file_moved.connect(self.update_files_moved_count)
        self.thread.started.connect(lambda: self.sorter.move_files_to_folders(
            self.directory_path_for_sort, 
            self.checkboxMove.isChecked(), 
            self.checkboxStay.isChecked()
        ))
        
        self.thread.start()
        self.thread.finished.connect(self.thread.deleteLater)
        # self.export_files_with_notification("Готово")
    def generate_report(self):
        try:
            self.excel_generator.generate_hierarchy_report(self.directory_path_for_sort)
            self.export_files_with_notification("Отчет сгенерирован!")
        except :
            self.export_files_with_notification("Ошибка генерации отчета!")
    @Slot(int)
    def update_files_moved_count(self,_):
        
        self.files_moved_count += 1
        print(f"Файлов перемещено: {self.files_moved_count}") 
        self.label_files_moved_count.setText(f"Количество файлов:{ self.files_moved_count}")
        QApplication.processEvents()  

# if __name__ == "__main__":
#     app = QApplication(sys.argv)
#     window = DuplicateChecker()
#     window.show()
#     sys.exit(app.exec())
