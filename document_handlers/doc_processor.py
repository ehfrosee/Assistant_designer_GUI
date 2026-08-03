# -*- coding: utf-8 -*-
"""Процессор для DOC файлов (старый формат Word)"""

import os
import re
from pathlib import Path
from typing import List, Dict, Any

from document_handlers.base_processor import BaseDocumentProcessor

try:
    import textract
    HAS_TEXTRACT = True
except ImportError:
    HAS_TEXTRACT = False


class DOCProcessor(BaseDocumentProcessor):
    """Процессор для DOC файлов с использованием textract"""

    def __init__(self, config: Dict[str, Any]):
        super().__init__(config)
        if not HAS_TEXTRACT:
            self.logger.warning("textract не установлен. DOCProcessor недоступен.")

    def can_process(self, file_path: str) -> bool:
        return file_path.lower().endswith('.doc') and HAS_TEXTRACT

    def extract_content(self, file_path: str) -> str:
        """Извлечение текста из DOC с помощью textract"""
        try:
            text = textract.process(file_path).decode('utf-8', errors='ignore')
            # Простая пост-обработка: разбивка на абзацы
            return text
        except Exception as e:
            self.logger.error(f"Ошибка извлечения текста из DOC: {e}")
            raise