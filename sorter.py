import os
import re
import shutil
import traceback
from model import File
from PySide6.QtCore import *


class Sorter(QObject):
    file_moved = Signal(int)
    finished = Signal()
    log_message = Signal(str) 
    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path
        self._mutex = QMutex()
        self._condition = QWaitCondition()
        self._paused = False
        self._pause_condition = QWaitCondition()
        self._stopped = False
        self._current_operation: dict[str, str] | None = None
        self._lock = QMutex() 


    def get_file_names(self):
        try:
            if not os.path.isdir(self.folder_path):
                self.log("[Ошибка] Папка не существует: " + self.folder_path)
                return []
            file_paths = []
            for root, _, files in os.walk(self.folder_path):
                for f in files:
                    file_paths.append(os.path.join(root, f))  # теперь это путь, а не просто имя
            return file_paths
        except Exception as e:
            self.log(f"[Ошибка] Ошибка в get_file_names: {e}")
            traceback.print_exc()
            return []
    @Slot()
    def pause(self):
        with QMutexLocker(self._lock):
            self._paused = True
        self.log_message.emit("Операция приостановлена")

    @Slot()
    def resume(self):
        with QMutexLocker(self._lock):
            self._paused = False
            self._cv.wakeAll()
        self.log_message.emit("Пауза: операция возобновлена")

    @Slot()
    def stop(self):
        with QMutexLocker(self._lock):
            self._stopped = True
            self._cv.wakeAll()
        self.log_message.emit("Операция остановлена")

    def _wait_if_paused(self) -> bool:
        with QMutexLocker(self._lock):
            while self._paused and not self._stopped:
                self._cv.wait(self._lock, 100)  # таймаут для регулярной проверки
            return self._stopped

    def safe_copy(self, src: str, dst: str, move: bool = False) -> bool:
        self._current_operation = {"src": src, "dst": dst}
        tmp = dst + ".part"
        try:
            if self._wait_if_paused():
                return False
            buf = 1024 * 1024
            with open(src, "rb") as f_src, open(tmp, "wb") as f_dst:
                while True:
                    if self._wait_if_paused():
                        return False
                    chunk = f_src.read(buf)
                    if not chunk:
                        break
                    f_dst.write(chunk)
            os.replace(tmp, dst)  # атомарно
            if move and not self._stopped:
                try:
                    os.remove(src)
                except FileNotFoundError:
                    pass
            return True
        except Exception as e:
            self.log_message.emit(f"Ошибка копирования: {e}")
            return False
        finally:
            self._current_operation = None
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except Exception:
                pass
    def check_pause_stop(self):
        with QMutexLocker(self._mutex):
            while self._paused and not self._stopped:
                self._condition.wait(self._mutex)
            return self._stopped
        

    def check_pause(self):
        """Неблокирующая проверка паузы"""
        with QMutexLocker(self._lock):
            if self._paused and not self._stopped:
                self.log_message.emit("Пауза: ожидание...")
                self._pause_condition.wait(self._lock, 100)  # Таймаут 100 мс
            return self._stopped
        

    def safe_copy(self, src, dst, move=False):
        """Безопасное копирование с поддержкой паузы"""
        self._current_operation = {'src': src, 'dst': dst}
        
        try:
            # Проверяем паузу/остановку перед началом
            if self.check_pause() or self._stopped:
                return False

            buffer_size = 1024 * 1024  # 1MB
            with open(src, 'rb') as f_src:
                with open(dst, 'wb') as f_dst:
                    while True:
                        # Частая проверка паузы во время копирования
                        if self.check_pause() or self._stopped:
                            f_dst.close()
                            if os.path.exists(dst):
                                os.remove(dst)
                            return False
                        
                        data = f_src.read(buffer_size)
                        if not data:
                            break
                        f_dst.write(data)
                        QThread.msleep(1)  

            if move and not (self._paused or self._stopped):
                os.remove(src)
                
            return True
        except Exception as e:
            self.log_message.emit(f"Ошибка копирования: {str(e)}")
            return False
        finally:
            self._current_operation = None
    def log(self, message, error=None):
        """Универсальный метод для логирования сообщений"""
        full_message = message
        if error:
            full_message += f"\n{str(error)}\n{traceback.format_exc()}"
        self.log_message.emit(full_message)  # Отправляем сообщение в UI
    def replace_cyrillic_to_latin(self, text):
        try:
            translation_table = str.maketrans({'Н': 'H', 'В': 'B'})
            return text.translate(translation_table)
        except Exception as e:
            self.log(f"[Ошибка] Ошибка при замене кириллицы: {e}")
            traceback.print_exc()
            return text

    def extract_prefix(self, file_name):
        try:
            file_name_latin = self.replace_cyrillic_to_latin(file_name)

            # Проверяем сначала паттерн "Письма проектировщика"
            if re.search(r'120-\d{3}-\d{2,3}-пр[._\s-]*(HB\d{3}|НВ\d{3})[-_ ]?\d{4,6}', file_name_latin, re.IGNORECASE):
                # Это "ОССЗ - Письма проектировщика", возвращаем None, чтобы приоритезировать этот вариант
                return None

            # Ищем стандартный префикс HB600.360060
            pattern = r'(HB\d{3})[\s._,-]?(\d{6})'
            match = re.search(pattern, file_name_latin)

            if match:
                return f"{match.group(1)}.{match.group(2)}"
        except Exception as e:
            self.log(f"[Ошибка] Ошибка при извлечении префикса: {e}")
            traceback.print_exc()
        return None

    def match_alternative_patterns(self, file_name):
        try:
            lat_file_name = self.replace_cyrillic_to_latin(file_name)

            # 120-007-15-пр.HB600-106223 (любой набор разделителей)
            pattern1 = (
                r'120-\d{3}-\d{2,3}[\s._-]*пр[\s._-]*'
                r'HB[\s\-_]*\d{3}[\s\-_]*\d{4,6}'
            )
            if re.search(pattern1, lat_file_name, re.IGNORECASE):
                return "ОССЗ - Письма проектировщика"

            # ОССЗ - Письма
            if re.search(r'HB\d{3}-(ОССЗ|ЭДС)-\d{2,3}', lat_file_name, re.IGNORECASE):
                return "ОССЗ - Письма"

            # ОССЗ - Письма проектировщика (формат HB600-93-07)
            if re.search(r'HB\d{3}-\d{2,3}-\d{2,4}', lat_file_name):
                return "ОССЗ - Письма проектировщика"

            # ТР-HB600 изм.1
            pattern2 = (
                r'ТР-\d{3}-\d{2}[\s_\-]*HB\d{3}'
                r'(?:[\s._\-]*изм(?:\.\d+)?)?'
            )
            if re.search(pattern2, lat_file_name, re.IGNORECASE):
                return "ОССЗ - Письма проектировщика"

        except Exception as e:
            self.log(f"[Ошибка] Ошибка в match_alternative_patterns: {e}")
            traceback.print_exc()

        return None




    def create_escd_dict(self):
        print("Запуск метода create_escd_dict")
        escd_dict = {}
        try:
            for file in File.select():
                escd_dict[file.num] = file.name
        except Exception as e:
            print(f"[Ошибка] Ошибка при чтении из базы данных: {e}")
            traceback.print_exc()
        return escd_dict

    def count_files(self, statusStay):
        count = 0
        try:
            file_paths = self.get_file_names()
            for file_path in file_paths:
                file_name = os.path.basename(file_path)
                try:
                    lat_file_name = self.replace_cyrillic_to_latin(file_name)
                    prefix = self.extract_prefix(file_name)
                    if prefix or re.match(r'^120-\d{3}', lat_file_name) and re.search(r'(HB\d{3})', lat_file_name) or re.match(r'^HB\d{3}-\d{3}', lat_file_name) or not statusStay:
                        count += 1
                except Exception as e:
                    print(f"[Ошибка] Ошибка при анализе файла {file_name}: {e}")
                    traceback.print_exc()
        except Exception as e:
            self.log(f"[Ошибка] Ошибка при подсчете файлов: {e}")
            traceback.print_exc()
        return count

    def move_files_to_folders(self, destination_folder, statusMove, statusStay):
        try:
            file_paths  = self.get_file_names()
            escd_dict = self.create_escd_dict()
            seen = {}

            for file_path in file_paths:
                file_name = os.path.basename(file_path)
                try:
                    lat_file_name = self.replace_cyrillic_to_latin(file_name)

             
                    alt_folder = self.match_alternative_patterns(file_name)
                    if alt_folder:
                        project_code_match = re.search(r'(HB|НВ)[\s._-]?(\d{3})', lat_file_name, re.IGNORECASE)
                        if project_code_match:
                            project_code = f"{project_code_match.group(1).upper()}{project_code_match.group(2)}"
                        else:
                            project_code = "ОШИБКА"
                        target_folder = os.path.join(destination_folder, project_code, f"{project_code}.{alt_folder}")
                        self.moveable(target_folder, file_path, seen, statusMove)
                        continue 

                 
                    prefix = self.extract_prefix(file_name)
                    if prefix:
                        target_folder = self.handle_prefix_case(destination_folder, prefix, escd_dict)
                        self.moveable(target_folder, file_path, seen, statusMove)
                    elif not statusStay:
                        target_folder = os.path.join(destination_folder, "Прочие документы")
                        self.moveable(target_folder, file_path, seen, statusMove)

                except Exception as e:
                    self.log(f"[Ошибка] Ошибка при обработке файла {file_name}: {e}")
                    traceback.print_exc()


        except Exception as e:
            self.log(f"[Ошибка] Ошибка в move_files_to_folders: {e}")
            traceback.print_exc()
        finally:
            self.finished.emit()


    def extract_description(self, file_name):
        try:
            parts = file_name.split("_")
            if len(parts) > 1:
                return parts[1].replace(".dwg", "").replace(".xlsx", "")
        except Exception as e:
            self.log(f"[Ошибка] Ошибка при извлечении описания из имени файла {file_name}: {e}")
            traceback.print_exc()
        return ""

    def moveable(self, target_folder, file_path, seen, statusMove):
        try:
            self.copy_file_to_folder(file_path, target_folder, seen, statusMove)
        except Exception as e:
            self.log(f"[Ошибка] Ошибка в moveable для файла {file_path}: {e}")
            traceback.print_exc()

    def handle_prefix_case(self, destination_folder, prefix, escd_dict):
        try:
            project_code = prefix[:5]
            project_number = prefix[6:]
            target_folder = os.path.join(destination_folder, project_code)

            for i in range(2, len(project_number) + 1):
                sub_folder = project_number[:i]
                folder_description = escd_dict.get(sub_folder, "")
                folder_name = f"{project_code}.{sub_folder} - {folder_description}" if folder_description else f"{project_code}.{sub_folder}"
                target_folder = os.path.join(target_folder, folder_name)
            return target_folder
        except Exception as e:
            self.log(f"[Ошибка] Ошибка при формировании пути по префиксу: {e}")
            traceback.print_exc()
            return destination_folder

    def handle_specific_format_case_120(self, destination_folder, file_name, specific_folder):
        try:
            project_code = re.search(r'(HB\d{3})', file_name).group(1)
            return os.path.join(destination_folder, project_code, f"{project_code}.{specific_folder}")
        except Exception as e:
            self.log(f"[Ошибка] Ошибка при разборе 120-файла: {e}")
            traceback.print_exc()
            return destination_folder

    def handle_specific_format_case(self, destination_folder, file_name, specific_folder):
        try:
            project_code = file_name[:5]
            return os.path.join(destination_folder, project_code, f"{project_code}.{specific_folder}")
        except Exception as e:
            self.log(f"[Ошибка] Ошибка при разборе HB-файла: {e}")
            traceback.print_exc()
            return destination_folder

    def append_extension_folder(self, target_folder, file_name):
        try:
            extension = os.path.splitext(file_name)[1][1:].lower()
            return os.path.join(target_folder, extension) if extension else target_folder
        except Exception as e:
            self.log(f"[Ошибка] Ошибка при добавлении папки по расширению: {e}")
            traceback.print_exc()
            return target_folder

    def copy_file_to_folder(self, file_path, target_folder, seen, status):
        if self.check_pause_stop():
            return

        try:
            os.makedirs(target_folder, exist_ok=True)

            file_name = os.path.basename(file_path)  # ← имя файла без пути
            new_file_name = self.replace_cyrillic_to_latin(file_name)
            dest_path = os.path.join(target_folder, new_file_name)

            counter = 1
            while os.path.exists(dest_path):
                if self.check_pause_stop():
                    return
                name, ext = os.path.splitext(new_file_name)
                dest_path = os.path.join(target_folder, f"{name}_{counter}{ext}")
                counter += 1

            if self.safe_copy(file_path, dest_path, move=status):
                self.file_moved.emit(1)
                self.log_message.emit(f"Успешно: {'перемещен' if status else 'скопирован'} {file_path} -> {dest_path}")
            else:
                self.log_message.emit(f"Операция отменена для файла: {file_path}")

        except Exception as e:
            self.log_message.emit(f"[Ошибка] Ошибка обработки файла {file_path}: {str(e)}")
            traceback.print_exc()



    def find_file_recursive(self, file_name):
        try:
            file_name_encoded = file_name.encode('utf-8', errors='ignore').decode('utf-8')
            for root, _, files in os.walk(self.folder_path):
                for f in files:
                    f_encoded = f.encode('utf-8', errors='ignore').decode('utf-8')
                    if f_encoded == file_name_encoded:
                        return os.path.join(root, f)
            return None
        except UnicodeError:
            return None


