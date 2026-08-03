# [file name]: xlsx_processor.py
import os
from pathlib import Path
from typing import List, Dict, Any

from document_handlers.base_processor import BaseDocumentProcessor

# Импорты для XLSX
try:
    import pandas as pd
    import openpyxl

    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False


class XLSXProcessor(BaseDocumentProcessor):
    """Процессор для XLSX файлов"""

    def can_process(self, file_path: str) -> bool:
        return file_path.lower().endswith('.xlsx') and HAS_XLSX

    def extract_content(self, file_path: str) -> str:
        """
        Извлечение содержимого из XLSX (реализация абстрактного метода)
        """
        return self.extract_sheets_content(file_path)

    def extract_sheets_content(self, file_path: str) -> str:
        """Извлечение содержимого всех листов XLSX в формате Markdown"""
        try:
            content_lines = []
            table_counter = 0

            # Открываем файл для получения информации о листах
            workbook = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            sheet_names = workbook.sheetnames

            for sheet_name in sheet_names:
                sheet_content = self._extract_sheet_content(file_path, sheet_name, table_counter)
                if sheet_content:
                    content_lines.append(f"### Лист: {sheet_name}")
                    content_lines.append("")
                    content_lines.append(sheet_content)
                    content_lines.append("")  # Пустая строка между листами

            workbook.close()
            return '\n'.join(content_lines)

        except Exception as e:
            self.logger.error(f"Ошибка извлечения содержимого XLSX {file_path}: {e}")
            raise

    def _extract_sheet_content(self, file_path: str, sheet_name: str, table_counter: int) -> str:
        """Извлечение содержимого конкретного листа"""
        try:
            # Используем pandas для чтения данных
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)

            if df.empty:
                return ""

            content_lines = []
            table_data = []

            # Преобразуем DataFrame в список списков
            for _, row in df.iterrows():
                row_data = [str(cell) if pd.notna(cell) else "" for cell in row]
                table_data.append(row_data)

            if not table_data:
                return ""

            # Форматируем как таблицу в Markdown (используем TableProcessor)
            formatted_table = self.table_processor.format_table(table_data)
            if formatted_table:
                content_lines.append(formatted_table)

            return '\n'.join(content_lines)

        except Exception as e:
            self.logger.warning(f"Ошибка извлечения листа {sheet_name}: {e}")
            return ""