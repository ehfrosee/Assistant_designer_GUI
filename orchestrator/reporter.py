# orchestrator/reporter.py

import json
import csv
import io
from typing import Any, List, Dict


class Reporter:
    @staticmethod
    def format_result(content: str, source_format: str, target_format: str) -> str:
        """Преобразует результат из source_format в target_format."""
        if source_format == target_format:
            return content

        if source_format == "json":
            try:
                data = json.loads(content)
                if target_format == "markdown":
                    return Reporter._json_to_markdown(data)
                elif target_format == "csv":
                    return Reporter._json_to_csv(data)
                else:
                    # fallback: возвращаем как JSON
                    return json.dumps(data, indent=2, ensure_ascii=False)
            except json.JSONDecodeError:
                # если невалидный JSON, возвращаем как есть
                return content

        # другие преобразования пока не поддерживаются
        return content

    @staticmethod
    def _json_to_markdown(data: Any) -> str:
        """Преобразует JSON (массив объектов) в Markdown-таблицу."""
        if isinstance(data, list) and data and all(isinstance(item, dict) for item in data):
            # Массив объектов → таблица
            keys = list(data[0].keys())
            rows = ["| " + " | ".join(keys) + " |"]
            rows.append("|" + "|".join([" --- " for _ in keys]) + "|")
            for item in data:
                row = []
                for k in keys:
                    val = item.get(k, "")
                    if isinstance(val, (list, dict)):
                        val = json.dumps(val, ensure_ascii=False)
                    row.append(str(val).replace("|", "\\|"))
                rows.append("| " + " | ".join(row) + " |")
            return "\n".join(rows)
        else:
            # Если не массив объектов, выводим как JSON-блок
            return f"```json\n{json.dumps(data, indent=2, ensure_ascii=False)}\n```"

    @staticmethod
    def _json_to_csv(data: Any) -> str:
        """Преобразует JSON (массив объектов) в CSV-строку."""
        if not isinstance(data, list) or not data:
            # Пустой список или не массив
            return ""

        # Определяем заголовки из первого объекта
        if isinstance(data[0], dict):
            headers = list(data[0].keys())
            output = io.StringIO()
            writer = csv.DictWriter(output, fieldnames=headers, quoting=csv.QUOTE_MINIMAL)
            writer.writeheader()
            for row in data:
                # Приводим значения к строке, обрабатываем вложенные структуры
                row_clean = {}
                for k, v in row.items():
                    if isinstance(v, (list, dict)):
                        row_clean[k] = json.dumps(v, ensure_ascii=False)
                    else:
                        row_clean[k] = str(v) if v is not None else ""
                writer.writerow(row_clean)
            return output.getvalue()
        else:
            # Если не объекты, попробуем записать как простой список
            output = io.StringIO()
            writer = csv.writer(output, quoting=csv.QUOTE_MINIMAL)
            for item in data:
                if isinstance(item, (list, tuple)):
                    writer.writerow([str(v) for v in item])
                else:
                    writer.writerow([str(item)])
            return output.getvalue()