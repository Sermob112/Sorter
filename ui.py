
import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QLineEdit, QVBoxLayout, 
    QHBoxLayout, QFileDialog, QMessageBox, QCheckBox, QTabWidget, QTextEdit
)
from PySide6.QtCore import *
from sorter import Sorter
from excel import ExcelGenerator
from PySide6.QtCore import QThread
from PySide6.QtWidgets import QApplication
from ExmWritter import *
import traceback
class DuplicateChecker(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Сортировщик и XML генератор")
        self.setGeometry(100, 100, 800, 850)
        
        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_sorter_tab(), "Сортировщик")
        self.tabs.addTab(self.create_xml_generator_tab(), "XML генератор")
        
        main_layout = QVBoxLayout()
        main_layout.addWidget(self.tabs)

        layout = QVBoxLayout()

        #  поле для логов
        self.log_label = QLabel("Лог ошибок:")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(350)
        
        main_layout.addWidget(self.log_label)
        main_layout.addWidget(self.log_text)


        self.control_buttons_layout = QHBoxLayout()
    

        # кнопки управления потоком
        # self.button_pause = QPushButton("Пауза")
        # self.button_pause.setEnabled(False)
        # self.button_pause.clicked.connect(self.pause_sorting)
        
        self.button_stop = QPushButton("Стоп") 
        self.button_stop.setEnabled(False)
        self.button_stop.clicked.connect(self.stop_sorting)
        
        # self.control_buttons_layout.addWidget(self.button_pause)
        self.control_buttons_layout.addWidget(self.button_stop)
        
        layout.addLayout(self.control_buttons_layout)
        main_layout.addLayout(layout)
        self.setLayout(main_layout)


    @Slot(str)
    def log_error(self, message):
        """Слот для получения сообщений из Sorter"""
        self.log_text.append(message)
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        QApplication.processEvents()

    def create_sorter_tab(self):
        sorter_tab = QWidget()
        layout = QVBoxLayout()

        self.files_moved_count = 0

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
        self.checkboxStay = QCheckBox("Переместить\Копировать файлы, если соответсвуют формату.")
        self.checkboxJoiner = QCheckBox("Объединить PDF листы в один PDF файл")
        self.label_files_moved_count = QLabel("Файлов перемещено: 0")

        layout.addWidget(self.label_folder1)
        layout.addWidget(self.path_folder1)
        layout.addWidget(self.button_folder1)

        layout.addWidget(self.label_folder2)
        layout.addWidget(self.path_folder2)
        layout.addWidget(self.button_folder2)

        layout.addWidget(self.label_files_count)
        layout.addWidget(self.label_duplicates)
        layout.addWidget(self.label_files_moved_count)
        # layout.addWidget(self.checkboxMove)
        # layout.addWidget(self.checkboxStay)
        layout.addWidget(self.checkboxJoiner)

        self.button_layout = QHBoxLayout()

        self.button_check = QPushButton("Проверить")
        self.button_check.clicked.connect(self.check_folders)

        self.button_sort = QPushButton("Сортировать")
        self.button_sort.clicked.connect(self.generate_sort)
        self.button_report = QPushButton("Сгенерировать отчет")
        self.button_report.clicked.connect(self.generate_report)

        self.button_layout.addStretch()
        self.button_layout.addWidget(self.button_check)

        layout.addStretch()
        layout.addLayout(self.button_layout)

        sorter_tab.setLayout(layout)
        return sorter_tab

    def create_xml_generator_tab(self):
        xml_tab = QWidget()
        layout = QVBoxLayout()

        self.label_xml_folder1 = QLabel("Выберите папку с файлами для генерации XML:")
        self.button_xml_folder1 = QPushButton("Выбрать папку")
        self.path_xml_folder1 = QLineEdit()
        self.button_xml_folder1.clicked.connect(self.select_xml_folder1)

        self.label_xml_folder2 = QLabel("Выберите папку для сохранения XML файлов:")
        self.button_xml_folder2 = QPushButton("Выбрать папку")
        self.path_xml_folder2 = QLineEdit()
        self.button_xml_folder2.clicked.connect(self.select_xml_folder2)

        self.button_generate_xml = QPushButton("Генерировать")
        self.button_generate_xml.clicked.connect(self.generate_xml)

        layout.addWidget(self.label_xml_folder1)
        layout.addWidget(self.path_xml_folder1)
        layout.addWidget(self.button_xml_folder1)

        layout.addWidget(self.label_xml_folder2)
        layout.addWidget(self.path_xml_folder2)
        layout.addWidget(self.button_xml_folder2)

        layout.addWidget(self.button_generate_xml)

        xml_tab.setLayout(layout)
        return xml_tab

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

    def select_xml_folder1(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с файлами для генерации XML")
        if folder:
            self.path_xml_folder1.setText(folder)
            self.xml_directory_path = folder

    def select_xml_folder2(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения XML файла")
        if folder:
            self.path_xml_folder2.setText(folder)
            self.xml_directory_path_for_save = folder

    def check_folders(self):
        try:
            if not hasattr(self, 'directory_path') or not self.directory_path:
                self.export_files_with_notification("Выберите папку с файлами, которую нужно сортировать")
                return

            self.excel_generator = ExcelGenerator(self.directory_path)
            count = self.excel_generator.count_files()
            self.label_files_count.setText(f"Количество файлов: {count}")
            
            self.excel_generator.export_to_xlsx("Контрольная сумма файлов.xlsx")
            self.label_duplicates.setText(f"Количество дубликатов: {self.excel_generator.duplicate_counter}")
            self.export_files_with_notification("Сгенерирован файл: 'Контрольная сумма файлов.xlsx'")
            
            self.button_layout.addWidget(self.button_sort)
        except Exception as e:
            error_msg = f"[Ошибка] Ошибка при проверке папок: {e}\n{traceback.format_exc()}"
            self.log_error(error_msg)
            self.export_files_with_notification("Произошла ошибка при проверке папок.")

        # self.duplicate_count.setText(str(self.sorter.count_files()))

    def generate_xml(self):
        try:
            if not hasattr(self, 'xml_directory_path') or not self.xml_directory_path:
                self.export_files_with_notification("Выберите папку с файлами для генерации XML")
                return 
            if not hasattr(self, 'xml_directory_path_for_save') or not self.xml_directory_path_for_save:
                self.export_files_with_notification("Выберите папку для сохранения XML файлов")
                return 

            walker = LocalFileSystemWalker()
            files = walker.collect_files(self.xml_directory_path)

            processor = FileListProcessor(files)
            files = processor.save_only_filenames()
            processor.create_assembly_xml(files, f"{self.xml_directory_path_for_save}\\XML.xml")

            self.export_files_with_notification("XML файл успешно сгенерирован!")
        except Exception as e:
            error_msg = f"[Ошибка] Ошибка генерации XML: {e}\n{traceback.format_exc()}"
            self.log_error(error_msg)
            self.export_files_with_notification("Ошибка генерации XML.")

    def generate_sort(self):
        self.log_text.clear()
        self.files_moved_count = 0  
        try:
            if not hasattr(self, 'directory_path_for_sort') or not self.directory_path_for_sort:
                self.export_files_with_notification("Выберите папку, в которую будут перемещены отсортированные файлы")
                return 

            self.sorter = Sorter(self.directory_path)
            self.thread = QThread()
            # Подключаем сигналы
            self.sorter.log_message.connect(self.log_error)
            self.sorter.file_moved.connect(self.update_files_moved_count)
            self.sorter.finished.connect(self.thread.quit)
            self.sorter.finished.connect(self.sorting_finished)
            
        
            self.sorter.moveToThread(self.thread)
            self.thread.finished.connect(self.thread.deleteLater)

            self.thread.started.connect(lambda: self.sorter.move_files_to_folders(
                self.directory_path_for_sort,
                self.checkboxMove.isChecked(),
                self.checkboxStay.isChecked(),
                self.checkboxJoiner.isChecked()
            ))

            self.fc = self.sorter.count_files(self.checkboxStay.isChecked())

            #  кнопки управления
            # self.button_pause.setEnabled(True)
            self.button_stop.setEnabled(True)
            self.button_sort.setEnabled(False)

            self.thread.start()

            self.button_layout.addWidget(self.button_report)
        except Exception as e:
            error_msg = f"[Ошибка] Ошибка при запуске сортировки: {e}\n{traceback.format_exc()}"
            self.log_error(error_msg)
            self.export_files_with_notification("Ошибка запуска сортировки.")

    def generate_report(self):
        try:
            if not hasattr(self, 'excel_generator') or not hasattr(self, 'directory_path_for_sort'):
                self.export_files_with_notification("Сначала выполните проверку папок и сортировку.")
                return

            self.excel_generator.generate_hierarchy_report(self.directory_path_for_sort)
            self.export_files_with_notification("Отчет сгенерирован!")
        except Exception as e:
            error_msg = f"[Ошибка] Ошибка при генерации отчета: {e}\n{traceback.format_exc()}"
            self.log_error(error_msg)
            self.export_files_with_notification("Ошибка генерации отчета.")
    def pause_sorting(self):
        """Постановка на паузу с защитой от зависания"""
        try:
            if not hasattr(self, 'sorter') or not hasattr(self, 'thread'):
                return

            if self.button_pause.text() == "Пауза":
                # Асинхронный вызов через очередь событий
                QTimer.singleShot(0, lambda: (
                    self.sorter.pause(),
                    QApplication.processEvents()  # Принудительная обработка событий
                ))
                # self.button_pause.setText("Продолжить")
            else:
                QTimer.singleShot(0, lambda: (
                    self.sorter.resume(),
                    QApplication.processEvents()
                ))
                self.button_pause.setText("Пауза")
        except Exception as e:
            self.log_error(f"Ошибка паузы: {str(e)}")
            # self.button_pause.setText("Пауза")  # Сброс состояния кнопки при ошибке
    def stop_sorting(self):
        """Полная остановка потока"""
        if not hasattr(self, 'sorter') or not hasattr(self, 'thread'):
            return
            
        # Отправляем команду остановки
        self.sorter.stop()
        
        # Даем потоку время на корректное завершение
        if self.thread.isRunning():
            self.thread.quit()
            if not self.thread.wait(3000):  # Увеличиваем время ожидания
                self.thread.terminate()
                self.log_error("Поток был принудительно остановлен")
        
        # Обновляем UI
        # self.button_pause.setEnabled(False)
        self.button_stop.setEnabled(False)
        self.button_sort.setEnabled(True)
        self.log_error("Сортировка остановлена пользователем")

    def sorting_finished(self):
        """Обработчик завершения сортировки"""
        # self.button_pause.setEnabled(False)
        self.button_stop.setEnabled(False)
        self.button_sort.setEnabled(True)
        self.export_files_with_notification("Сортировка закончена")
    @Slot(int)
    def update_files_moved_count(self,_):
        self.files_moved_count += 1
        self.label_files_moved_count.setText(f"Количество файлов:{ self.files_moved_count} из {self.fc}")
        QApplication.processEvents()  
   
# if __name__ == "__main__":
#     app = QApplication(sys.argv)
#     window = DuplicateChecker()
#     window.show()
#     sys.exit(app.exec())
