import os
import re
import shutil
import traceback
from model import File,ProjectFile  
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
        
    def _extract_project_and_drawing(self, file_name: str) -> tuple[str, str] | None:
        # пример: "26.00026-901-002", допускаем разные разделители и пробелы
        s = self.replace_cyrillic_to_latin(file_name)
        m = re.search(r'\b(\d{1,5})\s*[._-]\s*(\d{5}-\d{3}-\d{3})\b', s)
        if m:
            project = str(int(m.group(1)))  # нормализуем без ведущих нулей: "026" -> "26"
            drawing = m.group(2)
            return project, drawing
        # fallback: если нашли только чертежный номер — попробу ем найти проект по БД
        m2 = re.search(r'\b(\d{5}-\d{3}-\d{3})\b', s)
        if m2:
            drawing = m2.group(1)
            rec = (ProjectFile
                .select(ProjectFile.project)
                .where(ProjectFile.drawing_number == drawing)
                .first())
            if rec and rec.project:
                return str(rec.project), drawing
        return None

    def check_pause(self):
        """Неблокирующая проверка паузы"""
        with QMutexLocker(self._lock):
            if self._paused and not self._stopped:
                self.log_message.emit("Пауза: ожидание...")
                self._pause_condition.wait(self._lock, 100)  # Таймаут 100 мс
            return self._stopped
        


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

            # 1. Приоритетное исключение для писем проектировщика
            if re.search(
                r'120-\d{3}-\d{2,3}-пр[._\s-]*(HB\d{3}|НВ\d{3})[-_ ]?\d{4,6}',
                file_name_latin,
                re.IGNORECASE
            ):
                return None  # Пусть их сначала обрабатывает match_alternative_patterns

            # 2. Универсальный поиск шаблона <любые буквы/цифры>.<6 цифр>
            # Например: "ТПР2201.362671", "02020.362671"
            match = re.search(r'([A-Za-zА-Яа-я0-9]{2,10})[\s._-]?(\d{6})', file_name_latin)
            if match:
                prefix = f"{match.group(1)}.{match.group(2)}"
                return prefix

            # 3. Поддержка древних форматов 00036-010-012
            match_alt = re.search(r'(\d{3,5}[-_]\d{2,3}[-_]\d{2,3})', file_name_latin)
            if match_alt:
                # Можно заменить "-" на "." для согласованности
                prefix = match_alt.group(1).replace("-", ".").replace("_", ".")
                return prefix

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


    def moveable(self, target_folder, file_path, seen, statusMove):
        try:
            self.copy_file_to_folder(file_path, target_folder, seen, statusMove)
        except Exception as e:
            self.log(f"[Ошибка] Ошибка в moveable для файла {file_path}: {e}")
            traceback.print_exc()

   

    def copy_file_to_folder(self, file_path, target_folder, seen, status):
        if self.check_pause_stop():
            return
        try:
            os.makedirs(target_folder, exist_ok=True)
            file_name = os.path.basename(file_path)
            # было: new_file_name = self.replace_cyrillic_to_latin(file_name)
            new_file_name = self.sanitize_filename(file_name)  # ← очищаем имя файла

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

    
    def _parse_project_and_code(self, file_name: str) -> tuple[str, str | None]:
        base = os.path.basename(file_name)
        m = re.match(r'^\s*([^\.\-\s]+)\s*[.\-]\s*(.*)$', base)
        if not m:
            return base.split()[0], None
        project = m.group(1).strip()
        rest = m.group(2)
        m6 = re.search(r'(\d{6})', rest.replace(' ', ''))
        code6 = m6.group(1) if m6 else None
        return project, code6

    def _build_hierarchy_path(self, destination_folder: str, project: str,
                          code6: str | None, escd_dict: dict[str, str]) -> str:
        base = os.path.join(destination_folder, self.sanitize_component(project))  # ← очистка проекта
        path = base
        if code6:
            for L in range(2, min(len(code6), 6) + 1):
                key = code6[:L]
                if key in escd_dict:
                    desc = escd_dict.get(key, "")
                    dirname = f"{project}.{key} - {desc}" if desc else f"{project}.{key}"
                    dirname = self.sanitize_component(dirname)  # ← очистка имени папки
                    path = os.path.join(path, dirname)
        return path

    def sanitize_component(self, name: str) -> str:
        """
        Очищает имя папки/компонента пути:
        - применяет replace_cyrillic_to_latin;
        - удаляет вхождения вида "(123)";
        - убирает запрещённые для Windows и нежелательные символы;
        - схлопывает повторные пробелы/разделители и обрезает края.
        """
        try:
            name = self.replace_cyrillic_to_latin(name)
            # убрать (68), ( 12 ), любые скобки с числами
            name = re.sub(r'\(\s*\d+\s*\)', '', name)
            # запреты Windows и нежелательные символы (добавлен '!')
            name = re.sub(r'[<>:"/\\|?*\!]+', '', name)
            # заменить множественные пробелы на один
            name = re.sub(r'\s+', ' ', name)
            # схлопнуть подряд идущие точки/дефисы/подчёркивания
            name = re.sub(r'([.\-_])\1+', r'\1', name)
            # обрезать пробелы/точки/дефисы по краям
            name = name.strip(' .-_')
            # подстраховка на случай полного очищения
            return name if name else 'unnamed'
        except Exception as e:
            self.log(f"[Ошибка] sanitize_component: {e}")
            traceback.print_exc()
            return name

    def sanitize_filename(self, filename: str) -> str:
        """
        Возвращает очищённое имя файла:
        - сначала replace_cyrillic_to_latin;
        - очистка базовой части (без расширения) по тем же правилам;
        - расширение сохраняется.
        """
        try:
            filename = self.replace_cyrillic_to_latin(filename)
            base, ext = os.path.splitext(filename)
            base = re.sub(r'\(\s*\d+\s*\)', '', base)                 # убрать (68)
            base = re.sub(r'[<>:"/\\|?*\!]+', '', base)               # убрать ! и запрещённые
            base = re.sub(r'\s+', ' ', base)                          # схлопнуть пробелы
            base = re.sub(r'([.\-_])\1+', r'\1', base)                # схлопнуть разделители
            base = base.strip(' .-_')                                  # обрезать края
            if not base:
                base = 'unnamed'
            # нормализуем расширение (оставляем точку и исходный регистр/можно .lower())
            return f"{base}{ext}"
        except Exception as e:
            self.log(f"[Ошибка] sanitize_filename: {e}")
            traceback.print_exc()
            return filename
        


    def _build_hierarchy_from_prefix(self, destination_folder: str, prefix: str,
                                 escd_dict: dict[str, str]) -> str:
        """
        prefix вида "HB900.360061" или "00036.010012" -> проект = часть до точки,
        number = часть после точки; строим проект/проект.K2 - .../ ... /проект.Kn - ...
        по всем ключам из escd_dict (2..6), если присутствуют.
        """
        try:
            if "." not in prefix:
                return destination_folder
            project_code, number = prefix.split(".", 1)
            base = os.path.join(destination_folder, project_code)
            path = base
            number = number.strip()
            for L in range(2, min(len(number), 6) + 1):
                key = number[:L]
                if key in escd_dict:
                    desc = escd_dict.get(key, "")
                    dirname = f"{project_code}.{key} - {desc}" if desc else f"{project_code}.{key}"
                    path = os.path.join(path, dirname)
            return path
        except Exception as e:
            self.log(f"[Ошибка] Ошибка при сборке пути по старому префиксу: {e}")
            traceback.print_exc()
            return destination_folder
        

    def move_files_to_folders(self, destination_folder, statusMove, statusStay):
        try:
            file_paths = self.get_file_names()
            escd_dict = self.create_escd_dict()
            seen = {}

            for file_path in file_paths:
                file_name = os.path.basename(file_path)
                try:
                    # 1) Новый приоритет: проект + 6-значный код -> многоуровневая иерархия
                    project, code6 = self._parse_project_and_code(file_name)
                    target_folder = self._build_hierarchy_path(destination_folder, project, code6, escd_dict)
                    if target_folder and target_folder != os.path.join(destination_folder, project):
                        self.moveable(target_folder, file_path, seen, statusMove)
                        continue

                    # 2) Старый префикс (низкий приоритет): extract_prefix
                    prefix = self.extract_prefix(file_name)
                    if prefix:
                        target_folder = self._build_hierarchy_from_prefix(destination_folder, prefix, escd_dict)
                        self.moveable(target_folder, file_path, seen, statusMove)
                        continue

                    # 3) Самый последний шаг: альтернативные паттерны
                    alt_folder = self.match_alternative_patterns(file_name)
                    if alt_folder:
                        # Складываем внутрь папки проекта (без дублирования названия проекта в подпапке)
                        target_folder = os.path.join(destination_folder, project, alt_folder)
                        self.moveable(target_folder, file_path, seen, statusMove)
                        continue

                    # 4) Фоллбек: просто в проект
                    self.moveable(os.path.join(destination_folder, project), file_path, seen, statusMove)

                except Exception as e:
                    self.log(f"[Ошибка] Ошибка при обработке файла {file_name}: {e}")
                    traceback.print_exc()

        except Exception as e:
            self.log(f"[Ошибка] Ошибка в move_files_to_folders: {e}")
            traceback.print_exc()
        finally:
            self.finished.emit()