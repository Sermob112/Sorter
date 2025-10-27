import os
import re
import shutil
import traceback
from model import File,ProjectFile  
from PySide6.QtCore import *
try:
    from pypdf import PdfMerger  # для новых pypdf
except Exception:
    try:
        from pypdf.merger import PdfMerger  # альтернативный путь
    except Exception:
        try:
            from PyPDF2 import PdfMerger  # для старых PyPDF2
        except Exception:
            PdfMerger = None
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
        self._moved_pdfs: list[str] = []

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
        

    def move_files_to_folders(self, destination_folder, statusMove, statusStay, merge_enabled):
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
                finally:
        # Последняя стадия: слияние листов
                    try:
                        self.merge_pdf_siblings(merge_enabled)
                    except Exception as e:
                        self.log_message.emit(f"[PDF] Ошибка финального слияния: {e}")
        except Exception as e:
            self.log(f"[Ошибка] Ошибка в move_files_to_folders: {e}")
            traceback.print_exc()
        finally:
            self.finished.emit()

    def _lookup_document_name(self, file_name: str) -> str | None:
        """
        Возвращает очищенное (для имени файла) наименование документа из БД
        по паре (проект, 6-значный код), либо None, если не найдено/пусто.
        """
        try:
            project, code6 = self._parse_project_and_code(file_name)
            if not code6:
                return None

            # 1) Проект + точное/частичное совпадение кода в чертежном номере
            rec = (
                ProjectFile
                .select(ProjectFile.document_name)
                .where(
                    (ProjectFile.project == project) &
                    ((ProjectFile.drawing_number == code6) |
                    (ProjectFile.drawing_number.contains(code6)))
                )
                .first()
            )

            # 2) Fallback: поиск без проекта (если формат чертежного номера в БД иной)
            if not rec:
                rec = (
                    ProjectFile
                    .select(ProjectFile.document_name)
                    .where(
                        (ProjectFile.drawing_number == code6) |
                        (ProjectFile.drawing_number.contains(code6))
                    )
                    .first()
                )

            if not rec or not rec.document_name:
                return None

            title = str(rec.document_name).strip()
            # Игнорируем технические значения
            if title.lower() == "нет данных" or title == "":
                return None

            # Очистка для безопасного включения в имя файла
            return self.sanitize_component(title)
        except Exception as e:
            self.log(f"[Ошибка] _lookup_document_name: {e}")
            traceback.print_exc()
            return None
        



    def _infer_drawing_candidates(self, file_name: str) -> tuple[str, list[str]]:
        """
        Возвращает (project, candidates) — проект и список возможных вариантов чертежного номера
        для поиска в БД: полная базовая строка, ее дефисная версия, 6-значный код и комбинации.
        """
        base = os.path.splitext(os.path.basename(file_name))[0]
        project, code6 = self._parse_project_and_code(file_name)
        # Варианты базового: как есть и с заменой разделителей на дефис
        hyph = re.sub(r'[.\s_]+', '-', base)
        candidates = [base, hyph]

        # Трёхгруппный формат типа 14701-212-011(ВО) если присутствует в базовой строке
        m3 = re.search(r'\b\d{5}-\d{3}-\d{3}[A-Za-zА-Яа-я]*\b', base)
        if m3:
            candidates.append(m3.group(0))

        # Если есть 6-значный код — добавить комбинации
        if code6:
            candidates.extend([code6, f"{project}-{code6}", f"{project}.{code6}"])

        # Уникализировать, сохранить порядок
        seen = set()
        uniq = []
        for c in candidates:
            if c not in seen:
                uniq.append(c)
                seen.add(c)
        return project, uniq
    


    def _lookup_document_name_any(self, file_name: str) -> str | None:
        """
        Ищет ProjectFile.document_name по набору кандидатов drawing_number и проекту.
        Приоритет: точное совпадение с проектом → contains с проектом → точное без проекта → contains без проекта.
        """
        try:
            project, candidates = self._infer_drawing_candidates(file_name)

            # 1) Точное совпадение + проект
            for cand in candidates:
                rec = (ProjectFile
                    .select(ProjectFile.document_name)
                    .where((ProjectFile.project == project) &
                            (ProjectFile.drawing_number == cand))
                    .first())
                if rec and rec.document_name:
                    title = str(rec.document_name).strip()
                    return self.sanitize_component(title) if title and title.lower() != "нет данных" else None

            # 2) Содержит + проект
            for cand in candidates:
                rec = (ProjectFile
                    .select(ProjectFile.document_name)
                    .where((ProjectFile.project == project) &
                            (ProjectFile.drawing_number.contains(cand)))
                    .first())
                if rec and rec.document_name:
                    title = str(rec.document_name).strip()
                    return self.sanitize_component(title) if title and title.lower() != "нет данных" else None

            # 3) Точное без проекта
            for cand in candidates:
                rec = (ProjectFile
                    .select(ProjectFile.document_name)
                    .where(ProjectFile.drawing_number == cand)
                    .first())
                if rec and rec.document_name:
                    title = str(rec.document_name).strip()
                    return self.sanitize_component(title) if title and title.lower() != "нет данных" else None

            # 4) Содержит без проекта
            for cand in candidates:
                rec = (ProjectFile
                    .select(ProjectFile.document_name)
                    .where(ProjectFile.drawing_number.contains(cand))
                    .first())
                if rec and rec.document_name:
                    title = str(rec.document_name).strip()
                    return self.sanitize_component(title) if title and title.lower() != "нет данных" else None

            return None
        except Exception as e:
            self.log(f"[Ошибка] _lookup_document_name_any: {e}")
            traceback.print_exc()
            return None
        

    def copy_file_to_folder(self, file_path, target_folder, seen, status):
        if self.check_pause_stop():
            return
        try:
            os.makedirs(target_folder, exist_ok=True)
            orig_name = os.path.basename(file_path)

            # Базовая очистка имени файла
            clean_file_name = self.sanitize_filename(orig_name)
            name, ext = os.path.splitext(clean_file_name)

            # Всегда дописываем через "_" при наличии названия из БД
            doc_title = self._lookup_document_name_any(orig_name)  # использует sanitize_component внутри
            if doc_title:
                if not name.endswith(f"_{doc_title}"):
                    name = f"{name}_{doc_title}"

            new_file_name = f"{name}{ext}"
            dest_path = os.path.join(target_folder, new_file_name)

            counter = 1
            while os.path.exists(dest_path):
                if self.check_pause_stop():
                    return
                base, ext2 = os.path.splitext(new_file_name)
                dest_path = os.path.join(target_folder, f"{base}_{counter}{ext2}")
                counter += 1
            
            if self.safe_copy(file_path, dest_path, move=status):
                self.file_moved.emit(1)
                self.log_message.emit(
                    f"Успешно: {'перемещен' if status else 'скопирован'} {file_path} -> {dest_path}"
                )
            else:
                self.log_message.emit(f"Операция отменена для файла: {orig_name}")

            if os.path.splitext(dest_path)[1].lower() == ".pdf":
                # Поворот выполняется сразу (как у вас уже сделано)
                try:
                    self.rotate_pdf_to_portrait(dest_path)
                except Exception as e:
                    self.log_message.emit(f"[PDF] Ошибка поворота {os.path.basename(dest_path)}: {e}")
                # Буферизуем для последующего слияния
                self._moved_pdfs.append(dest_path)

        except Exception as e:
            self.log_message.emit(f"[Ошибка] Ошибка обработки файла {file_path}: {str(e)}")
            traceback.print_exc()
    

    def _merge_key_from_name(self, file_path: str) -> tuple[str, str, str | None] | None:
        try:
            stem = os.path.splitext(os.path.basename(file_path))[0]

            # 1) Убрать конечные '(n)'
            stem = re.sub(r'\(\s*\d+\s*\)\s*$', '', stem).strip()
            # 2) Убрать ВСЕ вхождения 'Лист N' (глобально)
            stem_no_list = re.sub(r'\s*лист\s*\d+\s*', ' ', stem, flags=re.IGNORECASE)
            stem_no_list = re.sub(r'\s+', ' ', stem_no_list).strip()

            # 3) Разделить на '<чертёж>[_]<титул>' по последнему '_'
            idx = stem_no_list.rfind('_')
            if idx != -1:
                drawing_part = stem_no_list[:idx].strip()
                title_part = stem_no_list[idx+1:].strip()
                title_part = re.sub(r'\s*лист\s*\d+\s*', ' ', title_part, flags=re.IGNORECASE)
                title_part = re.sub(r'\s+', ' ', title_part).strip()
                title = title_part if title_part else None
            else:
                drawing_part = stem_no_list
                title = None

            # 4) Проект = всё до первой '.' или '-'
            m_proj = re.match(r'^\s*([^\.\-\s]+)\s*[.\-]', drawing_part)
            project = m_proj.group(1) if m_proj else drawing_part.split()[0]

            drawing_key = drawing_part
            return project, drawing_key, title
        except Exception:
            return None


        
    def merge_pdf_siblings(self, merge_enabled: bool):
        if not merge_enabled:
            return

        groups: dict[tuple[str, str], list[tuple[int, str, str | None]]] = {}
        titles_by_key: dict[tuple[str, str], list[str]] = {}

        for p in self._moved_pdfs:
            if os.path.splitext(p)[1].lower() != ".pdf":
                continue
            key = self._merge_key_from_name(p)
            if not key:
                continue
            project, drawing_key, title = key
            stem = os.path.splitext(os.path.basename(p))[0]
            m = re.search(r'лист\s*(\d+)', stem, flags=re.IGNORECASE)
            num = int(m.group(1)) if m else 10**9
            gk = (project, drawing_key)
            groups.setdefault(gk, []).append((num, p, title))
            if title:
                titles_by_key.setdefault(gk, []).append(title)

        for (project, drawing_key), items in groups.items():
            if len(items) < 2:
                continue
            items.sort(key=lambda x: x[0])

            # Выбор общего title
            title_candidates = titles_by_key.get((project, drawing_key), [])
            title_unique = sorted(set([t for t in title_candidates if t]), key=len, reverse=True)
            common_title = None
            if len(title_unique) == 1:
                common_title = title_unique[0]
            elif len(title_unique) > 1:
                common_title = title_unique[0]
                self.log_message.emit(
                    f"[PDF] Несовпадающие названия у листов {drawing_key}: {title_unique}"
                )

            out_dir = os.path.dirname(items[0][1])
            out_name = f"{drawing_key}.pdf" if not common_title else f"{drawing_key}_{common_title}.pdf"
            out_path = os.path.join(out_dir, out_name)
            tmp_path = out_path + ".tmp"

            try:
                if PdfMerger is not None:
                    merger = PdfMerger()
                    for _, f, _ in items:
                        merger.append(f)
                    with open(tmp_path, "wb") as fo:
                        merger.write(fo)
                    try:
                        merger.close()
                    except Exception:
                        pass
                else:
                    writer = PdfWriter()
                    for _, f, _ in items:
                        reader = PdfReader(f)
                        for page in reader.pages:
                            writer.add_page(page)
                    with open(tmp_path, "wb") as fo:
                        writer.write(fo)

                os.replace(tmp_path, out_path)

                outs = {os.path.normcase(out_path)}
                for _, f, _ in items:
                    try:
                        if os.path.normcase(f) not in outs and os.path.exists(f):
                            os.remove(f)
                    except Exception as e:
                        self.log_message.emit(f"[PDF] Не удалось удалить лист {os.path.basename(f)}: {e}")
                self.log_message.emit(f"[PDF] Объединено: {drawing_key} -> {out_name}")
            except Exception as e:
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except Exception:
                    pass
                self.log_message.emit(f"[PDF] Ошибка слияния {drawing_key}: {e}")
                continue


    def rotate_pdf_to_portrait(self, pdf_path: str) -> None:
        try:
            from pypdf import PdfReader, PdfWriter
        except Exception as e:
            self.log_message.emit(f"[PDF] Нет pypdf: {e} — пропуск {os.path.basename(pdf_path)}")
            return
        try:
            reader = PdfReader(pdf_path)
            writer = PdfWriter()
            changed = False

            for page in reader.pages:
                # Текущие размеры и поворот
                w = float(page.mediabox.width)
                h = float(page.mediabox.height)
                rotation = getattr(page, "rotation", 0) or 0

                # Эффективная ориентация с учётом поворота
                # landscape, если ширина "на экране" больше высоты
                effective_landscape = ((w > h and rotation % 180 == 0) or
                                    (h > w and rotation % 180 != 0))

                if effective_landscape:
                    # Совместимость с разными версиями pypdf/PyPDF2
                    try:
                        page.rotate(90)
                    except Exception:
                        try:
                            page.rotate_clockwise(90)
                        except Exception:
                            # запасной вариант через set rotation
                            page.rotate(90)
                    changed = True

                writer.add_page(page)

            if changed:
                tmp_path = pdf_path + ".tmp"
                with open(tmp_path, "wb") as f:
                    writer.write(f)
                os.replace(tmp_path, pdf_path)
                self.log_message.emit(f"[PDF] Поворот в портрет: {os.path.basename(pdf_path)}")
        except Exception as e:
            self.log_message.emit(f"[PDF] Ошибка поворота {os.path.basename(pdf_path)}: {e}")
