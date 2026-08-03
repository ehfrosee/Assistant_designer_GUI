# [file name]: base_processor.py
import json
import logging
import os
import re
from pathlib import Path
from typing import Dict, List, Any, Optional, Union, Tuple
from abc import ABC, abstractmethod
import datetime

from document_handlers.table_processor import TableProcessor


class BaseDocumentProcessor(ABC):
    """Базовый класс для обработки документов"""

    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.logger = logging.getLogger(self.__class__.__name__)
        self.table_processor = TableProcessor(config)

    @abstractmethod
    def can_process(self, file_path: str) -> bool:
        """Может ли процессор обработать данный файл"""
        pass

    @abstractmethod
    def extract_content(self, file_path: str) -> str:
        """Извлечение содержимого из файла"""
        pass

    def convert_to_txt(self, file_path: str, output_dir: str) -> str:
        """Конвертация файла в TXT с Markdown-разметкой"""
        try:
            self.logger.info(f"Конвертация {file_path}")

            # Извлекаем содержимое
            content = self.extract_content(file_path)

            # Получаем метаданные
            metadata = self.get_file_metadata(file_path)

            # Формируем итоговый контент с метаданными в markdown
            final_content = self._format_as_markdown(content, metadata)

            # Сохраняем результат
            output_path = os.path.join(output_dir, f"{Path(file_path).stem}.md")
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(final_content)

            self.logger.info(f"Сохранено в: {output_path}")
            return output_path

        except Exception as e:
            self.logger.error(f"Ошибка конвертации {file_path}: {e}")
            raise

    def _format_as_markdown(self, content: str, metadata: Dict) -> str:
        """Форматирует контент с метаданными в markdown"""
        lines = []

        # Заголовок документа (уровень 1)
        title = metadata.get('file_name', 'Документ')
        lines.append(f"# {title}")
        lines.append("")

        # Секция метаданных
        lines.append("## Метаданные документа")
        lines.append("")

        # Метаданные в виде списка
        lines.append(f"- **Файл**: `{metadata.get('file_name', '')}`")
        lines.append(f"- **Тип**: {metadata.get('processor', '').replace('Processor', '')}")
        lines.append(f"- **Размер**: {self._format_size(metadata.get('file_size', 0))}")
        lines.append(f"- **Создан**: {metadata.get('created', '')}")
        lines.append(f"- **Изменен**: {metadata.get('modified', '')}")
        lines.append("")

        # Дополнительные метаданные из процессора (штамп, колонтитулы и т.д.)
        if 'stamp' in metadata:
            lines.append("### Штамп документа")
            lines.append("")
            stamp = metadata['stamp']
            for key, value in stamp.items():
                if key != 'raw_footer_lines':
                    lines.append(f"- **{key}**: {value}")
            lines.append("")

        if 'header_text' in metadata and metadata['header_text']:
            lines.append("### Верхние колонтитулы")
            lines.append("")
            lines.append(metadata['header_text'])
            lines.append("")

        # Разделитель
        lines.append("---")
        lines.append("")

        # Содержимое документа
        lines.append("## Содержимое документа")
        lines.append("")
        lines.append(content)

        return "\n".join(lines)

    def _format_size(self, size_bytes: int) -> str:
        """Форматирует размер файла"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"

    def get_file_metadata(self, file_path: str) -> dict:
        """Получает метаданные файла"""
        stat = os.stat(file_path)

        return {
            'file_path': file_path,
            'file_name': os.path.basename(file_path),
            'file_size': stat.st_size,
            'created': datetime.datetime.fromtimestamp(stat.st_ctime).isoformat(),
            'modified': datetime.datetime.fromtimestamp(stat.st_mtime).isoformat(),
            'processor': self.__class__.__name__
        }


class DocumentPostProcessor:
    """Класс для пост-обработки и улучшения структуры документов в Markdown"""

    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.extract_metadata = config.get('document_structure', {}).get('extract_metadata', True)
        self.enhance_structure = config.get('document_structure', {}).get('enhance_structure', True)

    def process_document(self, content: str, file_path: str) -> Tuple[str, Dict[str, str]]:
        """
        Полная обработка документа

        Returns:
            кортеж (обработанный_контент, метаданные)
        """
        metadata = {}

        if self.extract_metadata:
            title, metadata = self.extract_metadata_and_title(content, file_path)

        if self.enhance_structure:
            content = self.enhance_content_structure(content)

        return content, metadata

    def extract_metadata_and_title(self, content: str, file_path: str) -> Tuple[str, Dict[str, str]]:
        """Извлекает заголовок и метаданные документа"""
        lines = content.split('\n')
        metadata = {}

        # Поиск метаданных в первых 20 строках
        for line in lines[:20]:
            self._extract_metadata_from_line(line.strip(), metadata)

        # Формируем заголовок из метаданных
        title = self._create_title_from_metadata(metadata, file_path)

        return title, metadata

    def _extract_metadata_from_line(self, line: str, metadata: Dict[str, str]):
        """Извлекает метаданные из строки"""
        # Код документа (СП, ГОСТ и т.д.)
        code_match = re.search(r'(СП|ГОСТ|СНиП|ТР|СанПиН|ПБ)\s+[\d\.\-]+', line)
        if code_match and 'code' not in metadata:
            metadata['code'] = code_match.group(0)

        # ОКС
        oks_match = re.search(r'ОКС\s+[\d\.]+', line)
        if oks_match and 'oks' not in metadata:
            metadata['oks'] = oks_match.group(0).replace('ОКС ', '')

        # ОКВЭД
        okved_match = re.search(r'ОК\s*ВЭД\s+[A-Z]\s+\d+', line)
        if okved_match and 'okved' not in metadata:
            metadata['okved'] = okved_match.group(0)

        # Дата введения
        date_match = re.search(r'Дата введения\s+[\d\-\.]+', line)
        if date_match and 'date_intro' not in metadata:
            metadata['date_intro'] = date_match.group(0).replace('Дата введения ', '')

        # Название документа (ищем многострочное название)
        if 'code' in metadata and 'title' not in metadata:
            code_pos = line.find(metadata['code'])
            if code_pos != -1:
                remaining_text = line[code_pos + len(metadata['code']):].strip()
                if remaining_text and len(remaining_text) > 5:
                    metadata['title'] = remaining_text

    def _create_title_from_metadata(self, metadata: Dict[str, str], file_path: str) -> str:
        """Создает заголовок документа из метаданных"""
        title_parts = []

        if 'code' in metadata:
            title_parts.append(metadata['code'])

        if 'title' in metadata:
            title_parts.append(metadata['title'])

        if title_parts:
            return ' '.join(title_parts)
        else:
            return Path(file_path).stem

    def enhance_content_structure(self, content: str) -> str:
        """Улучшает структуру содержимого документа в формате Markdown"""
        lines = content.split('\n')
        enhanced_lines = []

        i = 0
        table_counter = 0
        page_counter = 0

        while i < len(lines):
            line = lines[i].rstrip()

            if not line:
                enhanced_lines.append("")
                i += 1
                continue

            # Обработка маркеров страниц
            if line.startswith('--- Страница'):
                page_counter += 1
                # Преобразуем в Markdown-разделитель с номером страницы
                page_num = re.search(r'Страница (\d+)', line)
                if page_num:
                    enhanced_lines.append(f"--- **Страница {page_num.group(1)}** ---")
                else:
                    enhanced_lines.append("--- **Новая страница** ---")
                enhanced_lines.append("")
                i += 1
                continue

            # Обработка маркеров таблиц (из старых версий)
            if line == "--- Таблица начало ---":
                table_counter += 1
                # Пропускаем этот маркер, таблица будет обработана отдельно
                i += 1
                continue

            if line == "--- Таблица конец ---":
                # Пропускаем этот маркер
                i += 1
                continue

            # Обработка названий листов Excel
            if line.startswith('--- Лист:'):
                sheet_name = line.replace('--- Лист:', '').replace('---', '').strip()
                enhanced_lines.append(f"### Лист: {sheet_name}")
                enhanced_lines.append("")
                i += 1
                continue

            # Определение заголовков по их содержимому
            header_level = self._detect_header_level(line)
            if header_level > 0:
                # Убираем возможные старые маркеры и добавляем правильный уровень
                clean_text = re.sub(r'^#+\s+', '', line)
                enhanced_lines.append(f"{'#' * header_level} {clean_text}")
                i += 1
                continue

            # Обработка таблиц в Markdown-формате (они уже отформатированы TableProcessor)
            if line.startswith('### Таблица'):
                # Это уже обработанная таблица, оставляем как есть
                enhanced_lines.append(line)
                i += 1

                # Добавляем все строки таблицы до пустой строки
                while i < len(lines) and lines[i].strip() != "":
                    enhanced_lines.append(lines[i])
                    i += 1
                continue

            # Обычный текст
            enhanced_lines.append(line)
            i += 1

        return '\n'.join(enhanced_lines)

    def _detect_header_level(self, line: str) -> int:
        """
        Определяет уровень заголовка на основе содержимого строки
        Возвращает 0 если это не заголовок, иначе уровень (1-4)
        """
        # Проверяем, не является ли строка уже размеченным заголовком
        if line.startswith('# '):
            return 1
        elif line.startswith('## '):
            return 2
        elif line.startswith('### '):
            return 3
        elif line.startswith('#### '):
            return 4

        # Анализируем структуру нумерации
        clean_line = line.strip()

        # Проверяем на наличие нумерации вида "1.", "1.1.", "1.1.1."
        if re.match(r'^\d+\.\d+\.\d+\.', clean_line):
            return 4  # Четвертый уровень
        elif re.match(r'^\d+\.\d+\.', clean_line):
            return 3  # Третий уровень
        elif re.match(r'^\d+\.', clean_line):
            return 2  # Второй уровень

        # Проверяем на ключевые слова заголовков
        header_keywords = ['введение', 'предисловие', 'приложение', 'глава',
                           'раздел', 'содержание', 'заключение', 'список']

        if any(keyword in clean_line.lower() for keyword in header_keywords):
            # Определяем важность по длине
            if len(clean_line) < 50:
                return 2
            else:
                return 3

        # Проверяем на заглавные буквы (возможный заголовок)
        words = clean_line.split()
        if len(words) > 2 and all(word[0].isupper() for word in words if word):
            if len(clean_line) < 100:
                return 3

        return 0  # Не заголовок