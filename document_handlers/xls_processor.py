# -*- coding: utf-8 -*-
"""Процессор для XLS файлов (старый формат Excel)"""

import os
from pathlib import Path
from typing import List, Dict, Any

from document_handlers.base_processor import BaseDocumentProcessor

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


class XLSProcessor(BaseDocumentProcessor):
    """Процессор для XLS файлов с использованием pandas (xlrd)"""

    def __init__(self, config: Dict[str, Any]):
        super().__init__(config)
        if not HAS_PANDAS:
            self.logger.warning("pandas не установлен. XLSProcessor недоступен.")

    def can_process(self, file_path: str) -> bool:
        return file_path.lower().endswith('.xls') and HAS_PANDAS

    def extract_content(self, file_path: str) -> str:
        """Извлечение содержимого из XLS в формате Markdown-таблицы"""
        try:
            # Читаем все листы с помощью pandas + xlrd
            xls = pd.ExcelFile(file_path, engine='xlrd')
            all_content = []
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, dtype=str)
                table_data = df.values.tolist()
                if table_data:
                    all_content.append(f"### Лист: {sheet_name}")
                    all_content.append(self.table_processor.format_table(table_data))
            return '\n'.join(all_content)
        except Exception as e:
            self.logger.error(f"Ошибка извлечения содержимого из XLS: {e}")
            raise