# project_loader.py
import re
from datetime import date
import pandas as pd
from model import ProjectFile, db  # Предполагается, что ProjectFile уже добавлен в model.py
from peewee import IntegrityError
DEFAULT_TEXT = "нет данных"
DEFAULT_DATE = date(1900, 1, 1)
class ProjectLoader:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path

    @staticmethod
    def _normalize_name(s: str) -> str:
        # Приводим заголовок к каноничному виду: трим, нижний регистр, один пробел, замена ё->е
        s = str(s).strip()
        s = re.sub(r"\s+", " ", s)
        s = s.replace("Ё", "Е").replace("ё", "е")
        # Исправляем возможные смешения кириллицы и латиницы в слове "примечание"
        s = s.replace("примечанеe", "примечанее").replace("примечание", "примечанее")
        return s.lower()

    def _read_excel(self) -> pd.DataFrame:
        df = pd.read_excel(self.excel_path, header=0)
        # Удаляем полностью пустые столбцы (часто "Unnamed")
        df = df.dropna(axis=1, how="all")
        # Нормализуем заголовки
        df.columns = [self._normalize_name(c) for c in df.columns]
        return df

    def _select_and_rename(self, df: pd.DataFrame) -> pd.DataFrame:
        expected_ru = [
            "№пп",
            "проект",
            "чертежный номер",
            "наименование документа",
            "листов, ед.",
            "масштаб",
            "формат",
            "утвердил",
            "дата",
            "компания",
            "ссылка файла",
            "примечанее",
            "размер файла",
            "дата файла",
        ]

        # Нормализация как в вашем коде
        expected_ru_norm = [self._normalize_name(c) for c in expected_ru]
        aliases = {
            "№пп": "ignore_num",
            "проект": "project",
            "чертежный номер": "drawing_number",
            "наименование документа": "document_name",
            "листов, ед.": "sheets",
            "масштаб": "scale",
            "формат": "format",
            "утвердил": "approved_by",
            "дата": "date",
            "компания": "company",
            "ссылка файла": "file_link",
            "примечанее": "note",
            "размер файла": "file_size",
            "дата файла": "file_date",
        }
        aliases_norm = {self._normalize_name(k): v for k, v in aliases.items()}

        present = [c for c in df.columns if c in aliases_norm]

        # Fallback: берём первые 14 или меньше, если их меньше
        if len(present) < 10:
            n = min(14, df.shape[1])
            df14 = df.iloc[:, :n].copy()
            # присвоить только n имён
            df14.columns = expected_ru_norm[:n]
            df = df14

        # Переименуем найденные
        rename_map = {c: aliases_norm[c] for c in df.columns if c in aliases_norm}
        df = df.rename(columns=rename_map)

        keep = [
            "project",
            "drawing_number",
            "document_name",
            "sheets",
            "scale",
            "format",
            "approved_by",
            "date",
            "company",
            "file_link",
            "note",
            "file_size",
            "file_date",
        ]
        # Добавить отсутствующие поля как пустые
        for k in keep:
            if k not in df.columns:
                df[k] = None

        df = df[keep]
        return df


   # Было (ошибочно комбинирует @staticmethod + self):
# @staticmethod
# def _coerce_types(self, df: pd.DataFrame) -> pd.DataFrame:

# Стало (вариант A — метод экземпляра, без декоратора):
    def _coerce_types(self, df: pd.DataFrame) -> pd.DataFrame:
            # Даты -> python date или None
            for col in ["date", "file_date"]:
                s = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
                df[col] = s.apply(lambda x: x.date() if pd.notna(x) else None)

            # Строковые поля -> str без лишних пробелов, None для пропусков
            str_cols = [
                "project", "drawing_number", "document_name", "sheets", "scale",
                "format", "approved_by", "company", "file_link", "note", "file_size"
            ]
            for col in str_cols:
                df[col] = df[col].astype(object)
                df[col] = df[col].where(pd.notna(df[col]), None)
                df[col] = df[col].apply(lambda v: v.strip() if isinstance(v, str) else v)
            return df

    def _apply_defaults(self, df: pd.DataFrame) -> pd.DataFrame:
        # Текстовые поля: пустые -> "нет данных"
        str_cols = [
            "project", "drawing_number", "document_name", "sheets", "scale",
            "format", "approved_by", "company", "file_link", "note", "file_size"
        ]
        for col in str_cols:
            df[col] = df[col].apply(
                lambda v: DEFAULT_TEXT if v in (None, "", "nan", "None") else v
            )

        # Даты: пустые -> DEFAULT_DATE
        for col in ["date", "file_date"]:
            df[col] = df[col].apply(lambda d: d if isinstance(d, date) else DEFAULT_DATE)
        return df

    def load_to_db(self) -> int:
        df = self._read_excel()
        df = self._select_and_rename(df)
        df = self._coerce_types(df)          # приведение типов
        df = df.astype(object).where(pd.notna(df), None)  # глобальная зачистка пропусков
        df = self._apply_defaults(df)        # применение дефолтов для NOT NULL

        records = df.to_dict(orient="records")
        inserted = 0
        with db.atomic():
            for batch_start in range(0, len(records), 100):
                batch = records[batch_start: batch_start + 100]
                try:
                    ProjectFile.insert_many(batch).execute()
                    inserted += len(batch)
                except IntegrityError:
                    # Даже при построчной вставке данные уже с дефолтами
                    for row in batch:
                        ProjectFile.create(**row)
                        inserted += 1
        return inserted


    @staticmethod
    def print_first_100():
        print("Первые 100 записей из ProjectFile:")
        for r in ProjectFile.select().limit(100):
            print(
                f"ID={r.id} | Проект: {r.project} | Чертеж: {r.drawing_number} | "
                f"Документ: {r.document_name} | Компания: {r.company}"
            )

if __name__ == "__main__":
    db.connect()
    db.create_tables([ProjectFile])
    loader = ProjectLoader("redan.xlsx")  # замените на путь к вашему Excel
    count = loader.load_to_db()
    print(f"Импортировано записей: {count}")
    loader.print_first_100()
    db.close()
