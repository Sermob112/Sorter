from openpyxl import Workbook,load_workbook
from openpyxl.styles import PatternFill
from collections import defaultdict
import re
import os
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from collections import defaultdict
import re
import os
import pandas as pd
import traceback  # для вывода полной информации об ошибках


class ExcelGenerator:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.duplicate_counter = 0

    def export_to_xlsx(self, output_file):
        def human_readable_size(size):
            for unit in ['Б', 'КБ', 'МБ', 'ГБ', 'ТБ']:
                if size < 1024.0:
                    return f"{size:.1f} {unit}"
                size /= 1024.0
            return f"{size:.1f} ТБ"

        try:
            file_names = self.get_file_names()

            if not file_names:
                print("Нет файлов для экспорта.")
                return

            duplicates = self.find_duplicates(file_names)
            self.duplicate_counter = len(duplicates)

            wb = Workbook()
            ws = wb.active
            ws.title = "File List"

            red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            ws.append(["Название файла", "Расширение", "Размер файла"])

            for full_path in file_names:
                try:
                    file_name = os.path.relpath(full_path, self.folder_path)
                    extension = os.path.splitext(file_name)[1].lower()
                    file_size = os.path.getsize(full_path)
                    readable_size = human_readable_size(file_size)

                    ws.append([file_name, extension, readable_size])

                    # Проверка на дубликат по имени файла (без пути)
                    if os.path.basename(file_name) in duplicates:
                        for col in range(1, 4):
                            ws.cell(row=ws.max_row, column=col).fill = red_fill
                except Exception as e:
                    print(f"[Ошибка] Не удалось обработать файл {full_path}: {e}")
                    traceback.print_exc()

            wb.save(output_file)

        except Exception as e:
            print(f"[Ошибка] Ошибка при экспорте в Excel: {e}")
            traceback.print_exc()

    def find_duplicates(self, file_paths):
        duplicates = []
        seen = {}

        for full_path in file_paths:
            file_name = os.path.basename(full_path)
            clean_name = self.clean_file_name(file_name)

            if clean_name in seen:
                duplicates.append(file_name)
            else:
                seen[clean_name] = file_name

        return duplicates


    def clean_file_name(self, file_name):
        try:
            name_without_extension, extension = os.path.splitext(file_name)
            clean_name = re.sub(r'\s*\(\d+\)$', '', name_without_extension)
            return clean_name + extension
        except Exception as e:
            print(f"[Ошибка] Ошибка при очистке имени файла {file_name}: {e}")
            traceback.print_exc()
            return file_name

    def count_files(self):
        try:
            if not os.path.isdir(self.folder_path):
                print(f"Папка {self.folder_path} не существует")
                return 0
            file_count = sum([len(files) for _, _, files in os.walk(self.folder_path)])
            return file_count
        except Exception as e:
            print(f"[Ошибка] Ошибка при подсчете файлов: {e}")
            traceback.print_exc()
            return 0

    def get_file_names(self):
        """
        Возвращает список файлов с полными путями во всех подкаталогах.
        """
        if not os.path.isdir(self.folder_path):
            print(f"[Ошибка] Папка {self.folder_path} не существует")
            return []

        file_paths = []
        for root, _, files in os.walk(self.folder_path):
            for file in files:
                full_path = os.path.join(root, file)
                file_paths.append(full_path)
        return file_paths


    def generate_hierarchy_report(self, root_folder, excel_path="Отчет по файлам.xlsx"):
        """
        Формирует Excel-отчёт по текущей иерархии:
        Проект / <Проект>.<код2 - имя2> / <код3 - имя3> / <код4 - имя4> / <код5 - имя5> / <код6 - имя6>
        + поддержка «особых» папок (ОССЗ - Письма, Прочие документы и т.п.) как уровня без кода.
        """
        import os, re, traceback
        import pandas as pd
        from collections import defaultdict

        # Регистр колонок-кодов/имен по уровням
        code_cols = ["Код 2", "Код 3", "Код 4", "Код 5", "Код 6"]
        name_cols = ["Наименование 2", "Наименование 3", "Наименование 4", "Наименование 5", "Наименование 6"]

        def parse_class_segment(project: str, segment: str):
            """
            Возвращает (level, code, name) для сегмента "<Проект>.<код> - <Описание>",
            где level ∈ {2,3,4,5,6} по длине кода, либо (None, None, special_name) для "особых" папок,
            либо (None, None, None) если это не классификаторный сегмент.
            """
            seg = segment.strip()
            # Особые папки
            if any(key in seg for key in ["ОССЗ - Письма проектировщика", "ОССЗ - Письма", "Прочие документы"]):
                return (None, None, seg)
            # Совпадение "<Проект>.<код> - <Описание>"
            # Пример: "02020.36 - Суда, судовое оборудование"
            m = re.match(rf"^{re.escape(project)}\.(\d{{2,6}})\s*-\s*(.+)$", seg)
            if m:
                code = m.group(1)
                name = m.group(2).strip()
                L = len(code)
                if 2 <= L <= 6:
                    return (L, code, name)
            return (None, None, None)

        # Хранилище: ключ строки -> счётчики по расширениям
        rows = {}
        ext_counts = defaultdict(lambda: defaultdict(int))
        all_exts = set()

        # Обход дерева
        for root, _, files in os.walk(root_folder):
            try:
                rel = os.path.relpath(root, root_folder)
                # Пропуск корня
                if rel == ".":
                    project = None
                    segments = []
                else:
                    segments = rel.split(os.sep)
                    project = segments[0] if segments else None

                if not project:
                    continue

                # Инициализация структуры уровней
                codes = {"2": "", "3": "", "4": "", "5": "", "6": ""}
                names = {"2": "", "3": "", "4": "", "5": "", "6": ""}
                special = None

                # Разобрать сегменты после проекта
                for seg in segments[1:]:
                    level, code, name = parse_class_segment(project, seg)
                    if level is None and code is None and name:
                        # особая папка — положим в "Наименование 2" если пусто; иначе в первый пустой "Наименование N"
                        placed = False
                        for k in ["2","3","4","5","6"]:
                            if not names[k]:
                                names[k] = name
                                placed = True
                                break
                        if not placed:
                            # если все заняты, добавим в конец последнего уровня
                            names["6"] = names["6"] + (" | " if names["6"] else "") + name
                    elif level:
                        key = str(level)
                        codes[key] = code
                        names[key] = name

                # Ключ строки (без расширений)
                row_key = (
                    project,
                    codes["2"], names["2"],
                    codes["3"], names["3"],
                    codes["4"], names["4"],
                    codes["5"], names["5"],
                    codes["6"], names["6"],
                )
                if row_key not in rows:
                    rows[row_key] = True  # маркер наличия строки

                # Учёт файлов
                for fn in files:
                    ext = os.path.splitext(fn)[1][1:].lower()
                    if not ext:
                        ext = "_noext"
                    ext_counts[row_key][ext] += 1
                    all_exts.add(ext)
            except Exception as e:
                self.log_message.emit(f"[Отчет] Ошибка при обходе папки {root}: {e}")
                traceback.print_exc()
                continue

        # Сборка датафрейма
        data = []
        for row_key in rows.keys():
            rec = {
                "Проект": row_key[0],
                "Классификатор код 2": row_key[1],
                "Классификатор имя 2": row_key[2],
                "Классификатор код 3": row_key[3],
                "Классификатор имя 3": row_key[4],
                "Классификатор код 4": row_key[5],
                "Классификатор имя 4": row_key[6],
                "Классификатор код 5": row_key[7],
                "Классификатор имя 5": row_key[8],
                "Классификатор код 6": row_key[9],
                "Классификатор имя 6": row_key[10],
            }
            for ext in all_exts:
                rec[ext] = ext_counts[row_key].get(ext, 0)
            rec["Всего файлов (строка)"] = sum(ext_counts[row_key].values())
            data.append(rec)

        try:
            import pandas as pd
            df = pd.DataFrame(data)
            # Упорядочим колонки: метаданные уровней + расширения + итог
            meta_cols = [
                "Проект",
                "Классификатор код 2","Классификатор имя 2",
                "Классификатор код 3","Классификатор имя 3",
                "Классификатор код 4","Классификатор имя 4",
                "Классификатор код 5","Классификатор имя 5",
                "Классификатор код 6","Классификатор имя 6",
            ]
            ext_cols = sorted([c for c in df.columns if c not in meta_cols + ["Всего файлов (строка)"]])
            df = df[meta_cols + ext_cols + ["Всего файлов (строка)"]]

            with pd.ExcelWriter(excel_path) as writer:
                df.to_excel(writer, index=False, sheet_name="Отчет")
        except Exception as e:
            self.log_message.emit(f"[Отчет] Ошибка записи Excel {excel_path}: {e}")
            traceback.print_exc()
