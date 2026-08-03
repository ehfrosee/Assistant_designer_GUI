# -*- coding: utf-8 -*-
"""Загрузка файлов через диалог (поддержка всех форматов)"""

from pathlib import Path
from typing import Optional, Tuple, List
from tkinter import filedialog
import sys
import os

# Добавляем путь для импорта адаптера документов
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from document_handlers.adapter import get_document_content


class FileLoader:
    """Утилита для загрузки файлов всех поддерживаемых форматов"""

    # Расширения для документов (обрабатываются через адаптер)
    DOCUMENT_EXTENSIONS = {".docx", ".xlsx", ".pdf", ".xls"}
    # Текстовые расширения (читаются напрямую)
    TEXT_EXTENSIONS = {".txt", ".md", ".py", ".json", ".csv"}

    @staticmethod
    def load_from_dialog(parent=None, config_manager=None) -> Optional[Tuple[str, str, str]]:
        """
        Открывает диалог выбора одного файла.
        Возвращает (имя_файла, содержимое, формат) или None.
        Формат: "markdown", "json" или "text"
        """
        file_path = filedialog.askopenfilename(
            parent=parent,
            title="Выберите файл",
            filetypes=[
                ("Все поддерживаемые", "*.docx *.xlsx *.pdf *.xls *.txt *.md *.py *.json *.csv"),
                ("Документы Word", "*.docx"),
                ("Таблицы Excel", "*.xlsx *.xls"),
                ("PDF", "*.pdf"),
                ("Текстовые", "*.txt *.md *.py *.json *.csv"),
                ("Все файлы", "*.*")
            ]
        )
        if not file_path:
            return None
        return FileLoader._read_file(file_path, config_manager)

    @staticmethod
    def load_multiple_from_dialog(parent=None, config_manager=None) -> List[Tuple[str, str, str]]:
        """Загружает несколько файлов через диалог."""
        file_paths = filedialog.askopenfilenames(
            parent=parent,
            title="Выберите файлы",
            filetypes=[
                ("Все поддерживаемые", "*.docx *.xlsx *.pdf *.xls *.txt *.md *.py *.json *.csv"),
                ("Документы Word", "*.docx"),
                ("Таблицы Excel", "*.xlsx *.xls"),
                ("PDF", "*.pdf"),
                ("Текстовые", "*.txt *.md *.py *.json *.csv"),
                ("Все файлы", "*.*")
            ]
        )
        result = []
        for file_path in file_paths:
            r = FileLoader._read_file(file_path, config_manager)
            if r:
                result.append(r)
        return result

    @staticmethod
    def load_from_drop(mime_data, config_manager=None) -> Optional[Tuple[str, str, str]]:
        """Заглушка для drag & drop (можно реализовать позже)."""
        # Здесь можно обработать mime_data, но пока возвращаем None
        return None

    @staticmethod
    def _read_file(file_path: str, config_manager=None) -> Optional[Tuple[str, str, str]]:
        """
        Читает файл и возвращает (имя, содержимое, формат) или None.
        Формат: "markdown", "json" или "text".
        """
        path = Path(file_path)
        ext = path.suffix.lower()
        name = path.name

        # --- Текстовые файлы ---
        if ext in FileLoader.TEXT_EXTENSIONS:
            encodings = ["utf-8", "cp1251", "latin-1"]
            for enc in encodings:
                try:
                    content = path.read_text(encoding=enc)
                    fmt = "json" if ext == ".json" else "text"
                    return name, content, fmt
                except (UnicodeDecodeError, PermissionError, OSError):
                    continue
            return None  # Не удалось прочитать

        # --- Документы (DOCX, XLSX, PDF) ---
        if ext in FileLoader.DOCUMENT_EXTENSIONS:
            config = config_manager.config if config_manager else {}
            try:
                result = get_document_content(str(path), config)
                content = result["content"]
                fmt = result.get("format", "markdown")  # "markdown" или "json"
                return name, content, fmt
            except Exception as e:
                print(f"Ошибка извлечения содержимого {name}: {e}")
                return None

        # Неподдерживаемый формат
        return None
