# [file name]: table_processor.py
"""
Модуль для унифицированной обработки таблиц из разных форматов документов
"""

import re
from typing import List, Dict, Any, Optional, Union


class TableProcessor:
    """
    Универсальный процессор для таблиц из разных источников
    """

    def __init__(self, config: Dict[str, Any] = None):
        self.config = config or {}
        self.max_header_rows = self.config.get('processors', {}).get('max_header_rows', 3)
        self.header_keywords = [
            'характеристика', 'показатель', 'параметр', 'наименование',
            'описание', 'тип', 'вид', 'категория', 'название', 'позиция',
            'обозначение', 'марка', 'свойство', 'значение', 'единица'
        ]

    def format_table(self, table_data: List[List[str]], table_number: int = None) -> str:
        """
        Универсальное форматирование таблицы в Markdown

        Args:
            table_data: список строк таблицы, каждая строка - список ячеек
            table_number: номер таблицы (если известен)

        Returns:
            отформатированная таблица в Markdown
        """
        if not table_data or not any(any(cell for cell in row if cell) for row in table_data):
            return ""

        # Очистка данных
        cleaned_data = self._clean_table_data(table_data)

        table_lines = []

        # Добавляем заголовок таблицы в Markdown
        if table_number:
            table_lines.append(f"### Таблица {table_number}")
        else:
            table_lines.append("### Таблица")

        # Автоматическое определение количества строк шапки
        header_row_count = self._detect_header_rows(cleaned_data)

        # Разделение на шапку и данные
        header_rows = cleaned_data[:header_row_count] if header_row_count > 0 else []
        data_rows = cleaned_data[header_row_count:]

        # Формируем Markdown-таблицу
        if header_rows:
            # Объединяем шапку
            merged_header = self._merge_header_columns(header_rows)
            if merged_header:
                # Заголовок таблицы
                table_lines.append("| " + " | ".join(merged_header) + " |")
                # Разделитель
                table_lines.append("|" + "|".join([" --- " for _ in merged_header]) + "|")

        # Добавление данных
        for row in data_rows:
            if row and any(cell for cell in row if cell):
                # Экранируем спецсимволы Markdown
                cleaned_cells = [self._escape_markdown(cell) for cell in row]
                table_lines.append("| " + " | ".join(cleaned_cells) + " |")

        table_lines.append("")  # Пустая строка после таблицы

        return "\n".join(table_lines)

    def _escape_markdown(self, text: str) -> str:
        """Экранирование спецсимволов Markdown"""
        if not text:
            return ""
        # Экранируем символы |, *, _, `, [, ], (, ), #, +, -, ., !
        text = str(text)
        special_chars = ['|', '*', '_', '`', '[', ']', '(', ')', '#', '+', '-', '.', '!']
        for char in special_chars:
            text = text.replace(char, '\\' + char)
        return text

    def _clean_table_data(self, table_data: List[List[str]]) -> List[List[str]]:
        """Очистка данных таблицы"""
        cleaned = []
        max_cols = max(len(row) for row in table_data) if table_data else 0

        for row in table_data:
            # Очищаем каждую ячейку
            cleaned_row = []
            for cell in row:
                if cell is None:
                    cleaned_cell = ""
                else:
                    cell_text = str(cell).strip()
                    cell_text = re.sub(r'\s+', ' ', cell_text)
                    cleaned_cell = cell_text
                cleaned_row.append(cleaned_cell)

            # Выравниваем длину строки
            while len(cleaned_row) < max_cols:
                cleaned_row.append("")

            # Добавляем только если есть непустые ячейки
            if any(cleaned_row):
                cleaned.append(cleaned_row)

        return cleaned

    def _detect_header_rows(self, table_data: List[List[str]]) -> int:
        """Автоматически определяет количество строк шапки (0-max_header_rows)"""
        if not table_data:
            return 0

        max_header_rows = min(self.max_header_rows, len(table_data))

        for i in range(max_header_rows):
            row = table_data[i]

            # Критерии для определения конца шапки:

            # 1. Следующая строка содержит числовые данные
            if i + 1 < len(table_data) and self._has_numeric_data(table_data[i + 1]):
                return i + 1

            # 2. Текущая строка выглядит как данные (длинный текст, числа)
            if self._looks_like_data_row(row):
                return i

            # 3. Текущая строка - явная шапка (короткий текст, заголовочные слова)
            if self._is_explicit_header_row(row):
                continue

            # 4. По умолчанию считаем первую строку шапкой
            if i == 0:
                continue

        return max_header_rows

    def _has_numeric_data(self, row: List[str]) -> bool:
        """Проверяет, содержит ли строка числовые данные"""
        for cell in row:
            if cell and re.search(r'\d+[,\.]\d+|\b\d+\b', str(cell)):
                return True
        return False

    def _looks_like_data_row(self, row: List[str]) -> bool:
        """Проверяет, выглядит ли строка как строка данных"""
        if not row:
            return False

        # Данные обычно содержат более длинный текст или числа
        total_chars = sum(len(str(cell)) for cell in row if cell)
        avg_chars_per_cell = total_chars / len(row) if row else 0

        return avg_chars_per_cell > 25 or self._has_numeric_data(row)

    def _is_explicit_header_row(self, row: List[str]) -> bool:
        """Проверяет, является ли строка явной шапкой"""
        if not row:
            return False

        row_text = ' '.join(str(cell).lower() for cell in row if cell)
        has_header_keywords = any(keyword in row_text for keyword in self.header_keywords)

        # Шапка обычно короткая
        total_chars = sum(len(str(cell)) for cell in row if cell)
        avg_chars_per_cell = total_chars / len(row) if row else 0

        return has_header_keywords and avg_chars_per_cell < 30

    def _merge_header_columns(self, header_rows: List[List[str]]) -> List[str]:
        """Объединяет шапку по колонкам (вертикальное объединение)"""
        if not header_rows:
            return []

        if len(header_rows) == 1:
            return header_rows[0]

        # Определяем максимальное количество колонок
        max_cols = max(len(row) for row in header_rows)

        # Создаем объединенную шапку по колонкам
        merged_header = []

        for col_idx in range(max_cols):
            column_cells = []
            for row in header_rows:
                if col_idx < len(row) and row[col_idx]:
                    cell_text = str(row[col_idx]).strip()
                    if cell_text:
                        column_cells.append(cell_text)

            if column_cells:
                # Объединяем ячейки вертикально через пробел
                merged_cell = " ".join(column_cells)
                merged_cell = re.sub(r'\s+', ' ', merged_cell).strip()
                merged_header.append(merged_cell)
            else:
                merged_header.append("")

        return merged_header