# [file name]: txt_processor.py
import os
from pathlib import Path
from typing import Dict, Any

from document_handlers.base_processor import BaseDocumentProcessor


class TXTProcessor(BaseDocumentProcessor):
    """Процессор для TXT файлов"""

    def can_process(self, file_path: str) -> bool:
        return file_path.lower().endswith('.txt')

    def extract_content(self, file_path: str) -> str:
        """
        Извлечение содержимого из TXT (реализация абстрактного метода)
        """
        return self._read_file_with_encoding(file_path)

    def _read_file_with_encoding(self, file_path: str) -> str:
        """Чтение файла с определением кодировки"""
        try:
            # Пробуем UTF-8
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except UnicodeDecodeError:
            # Пробуем другие кодировки
            encodings = ['cp1251', 'cp866', 'iso-8859-1', 'macroman']
            
            for encoding in encodings:
                try:
                    self.logger.info(f"Попытка чтения с кодировкой {encoding}: {file_path}")
                    with open(file_path, 'r', encoding=encoding) as f:
                        return f.read()
                except UnicodeDecodeError:
                    continue
                except Exception as e:
                    self.logger.warning(f"Ошибка с кодировкой {encoding}: {e}")
                    continue
            
            raise ValueError(f"Не удалось прочитать файл {file_path} с доступными кодировками")

    def convert_to_txt(self, file_path: str, output_dir: str) -> str:
        """Копирование TXT файла в выходную директорию"""
        try:
            self.logger.info(f"Копирование TXT файла: {file_path}")

            # Читаем содержимое файла через extract_content
            content = self.extract_content(file_path)

            # Создаем путь для выходного файла
            output_path = os.path.join(output_dir, f"{Path(file_path).stem}.txt")
            
            # Записываем содержимое в выходной файл
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(content)

            self.logger.info(f"TXT файл скопирован в: {output_path}")
            return output_path

        except Exception as e:
            self.logger.error(f"Ошибка обработки TXT файла {file_path}: {e}")
            raise