# [file name]: pdf_processor.py
import os
import re
from pathlib import Path
from typing import List, Dict, Any

from document_handlers.base_processor import BaseDocumentProcessor

# Импорты для PDF
try:
    import fitz  # PyMuPDF
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False


class PDFProcessor(BaseDocumentProcessor):
    """Процессор для PDF файлов"""

    def can_process(self, file_path: str) -> bool:
        return file_path.lower().endswith('.pdf') and HAS_PDF and HAS_PDFPLUMBER

    def extract_content(self, file_path: str) -> str:
        """
        Извлечение содержимого из PDF (реализация абстрактного метода)
        """
        return self.extract_text_with_tables_ordered(file_path)

    def extract_text_with_tables_ordered(self, file_path: str) -> str:
        """Извлечение текста из PDF с таблицами в формате Markdown"""
        try:
            full_content = []
            table_count = 0

            with pdfplumber.open(file_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    page_content = []

                    # Извлекаем таблицы
                    tables = page.find_tables()
                    table_bboxes = []

                    if tables:
                        for table in tables:
                            bbox = table.bbox
                            table_data = table.extract()

                            if table_data and any(any(cell for cell in row if cell) for row in table_data):
                                table_bboxes.append({
                                    'bbox': bbox,
                                    'data': table_data,
                                    'page_num': page_num
                                })

                    # Извлекаем текст
                    text = page.extract_text()
                    if text and text.strip():
                        lines = text.strip().split('\n')
                        for line in lines:
                            line = line.strip()
                            if line:
                                page_content.append(line)

                    # Добавляем таблицы (используем TableProcessor)
                    if table_bboxes:
                        # Сортируем таблицы по y-координате
                        table_bboxes.sort(key=lambda x: x['bbox'][1])

                        for table_info in table_bboxes:
                            table_count += 1
                            # Используем TableProcessor для форматирования в Markdown
                            table_content = self.table_processor.format_table(
                                table_info['data'],
                                table_number=table_count
                            )
                            if table_content:
                                page_content.append(table_content)

                    if page_content:
                        # Добавляем разделитель страницы в Markdown
                        full_content.append(f"--- **Страница {page_num}** ---")
                        full_content.append("")
                        full_content.extend(page_content)
                        full_content.append("")

            result = '\n'.join(full_content)
            self.logger.info(f"PDF извлечен: {table_count} таблиц")
            return result

        except Exception as e:
            self.logger.warning(f"Ошибка PDFPlumber: {e}, используется резервный метод")
            return self._extract_text_fallback(file_path)

    def _extract_text_fallback(self, file_path: str) -> str:
        """Резервный метод извлечения текста"""
        try:
            doc = fitz.open(file_path)
            full_content = []

            for page_num in range(len(doc)):
                page = doc[page_num]
                text = page.get_text("text", sort=True)

                if text.strip():
                    full_content.append(f"--- **Страница {page_num + 1}** ---")
                    full_content.append("")
                    lines = text.strip().split('\n')
                    for line in lines:
                        line = line.strip()
                        if line:
                            full_content.append(line)
                    full_content.append("")

            doc.close()
            return '\n'.join(full_content)

        except Exception as e:
            self.logger.error(f"Ошибка резервного метода извлечения текста: {e}")
            raise