# -*- coding: utf-8 -*-
"""Адаптер для вызова существующих процессоров документов"""

import os
import sys
from pathlib import Path
from typing import Dict, Any

# Добавляем путь к родительской директории для импорта процессоров
sys.path.insert(0, str(Path(__file__).parent.parent))

from document_handlers.docx_processor import DOCXProcessor
from document_handlers.pdf_processor import PDFProcessor
from document_handlers.xlsx_processor import XLSXProcessor
from document_handlers.txt_processor import TXTProcessor
from document_handlers.doc_processor import DOCProcessor
from document_handlers.xls_processor import XLSProcessor

# Глобальный кэш процессоров (создаём один раз)
_processors_cache = None


def _get_processors(config: Dict[str, Any] = None):
    """Инициализирует и возвращает список процессоров"""
    global _processors_cache
    if _processors_cache is None:
        config = config or {}
        _processors_cache = [
            PDFProcessor(config),
            DOCXProcessor(config),
            XLSXProcessor(config),
            TXTProcessor(config),
            XLSProcessor(config),
        ]
    return _processors_cache


def get_document_content(file_path: str, config: Dict[str, Any] = None) -> Dict[str, Any]:
    """
    Извлекает содержимое документа в формате Markdown/JSON.

    Args:
        file_path: путь к файлу
        config: конфигурация для процессоров

    Returns:
        dict с полями:
            - format: "markdown" или "json"
            - content: извлечённое содержимое
            - metadata: метаданные (опционально)
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Файл не найден: {file_path}")

    processors = _get_processors(config)

    for processor in processors:
        if processor.can_process(file_path):
            # Извлекаем содержимое
            content = processor.extract_content(file_path)

            # Определяем формат (по расширению или по умолчанию)
            ext = Path(file_path).suffix.lower()
            if ext in ('.xlsx', '.xls'):
                output_format = "json"
            else:
                output_format = "markdown"

            # Получаем метаданные
            metadata = processor.get_file_metadata(file_path)

            return {
                "format": output_format,
                "content": content,
                "metadata": metadata
            }

    raise ValueError(f"No processor for {file_path}")


def get_document_content_simple(file_path: str) -> str:
    """
    Упрощённая версия, возвращает только содержимое как строку.
    """
    result = get_document_content(file_path)
    return result["content"]