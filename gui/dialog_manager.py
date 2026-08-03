import json
import threading
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional


class DialogManager:
    def __init__(self, config_manager, chat_client=None):
        self.config_manager = config_manager
        self.chat_client = chat_client
        self.dialogs_dir = Path(self.config_manager.get_paths().get("dialogs_dir", "./dialogs"))
        self.dialogs_dir.mkdir(exist_ok=True)
        self.current_dialog = self._new_dialog()
        self._auto_save = True

    def _new_dialog(self) -> Dict[str, Any]:
        return {
            "dialog_id": datetime.now().strftime("%Y%m%d_%H%M%S"),
            "created_at": datetime.now().isoformat(),
            "messages": [],
            "summary": "",
            "attached_files": [],
            "loaded_documents": [],  # список {filename, content, format, path}
            "loaded_scenarios": [],  # список {path, data, name}
            "last_analysis_result": None,  # {result, format}
            "is_active": True
        }

    def new_dialog(self):
        self.current_dialog = self._new_dialog()
        if self._auto_save:
            self.save_dialog()

    def save_dialog(self, file_path: str = None):
        if file_path is None:
            file_path = self.dialogs_dir / f"{self.current_dialog['dialog_id']}.json"
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(self.current_dialog, f, ensure_ascii=False, indent=2)

    def load_dialog(self, file_path: str):
        with open(file_path, "r", encoding="utf-8") as f:
            self.current_dialog = json.load(f)
            # Обратная совместимость: если нет новых полей, создаём их
            if "loaded_documents" not in self.current_dialog:
                self.current_dialog["loaded_documents"] = []
            if "loaded_scenarios" not in self.current_dialog:
                self.current_dialog["loaded_scenarios"] = []
            if "last_analysis_result" not in self.current_dialog:
                self.current_dialog["last_analysis_result"] = None

    def add_message(self, role: str, content: str, tokens_used: int = 0):
        self.current_dialog["messages"].append({
            "role": role,
            "content": content,
            "timestamp": datetime.now().isoformat(),
            "tokens_used": tokens_used
        })
        if self._auto_save:
            self.save_dialog()
        if role == "assistant":
            self._auto_update_summary()

    def add_attached_file(self, filename: str, content: str):
        if not any(f["filename"] == filename for f in self.current_dialog["attached_files"]):
            self.current_dialog["attached_files"].append({"filename": filename, "content": content})
            if self._auto_save:
                self.save_dialog()

    def clear_attached_files(self):
        self.current_dialog["attached_files"] = []
        if self._auto_save:
            self.save_dialog()

    # ---------- Новые методы для сценариев ----------
    def add_loaded_document(self, filename: str, content: str, file_format: str, file_path: str = None):
        """Добавляет загруженный документ в диалог."""
        if file_path is None:
            file_path = filename  # используем имя файла как путь
        doc = {
            "filename": filename,
            "content": content,
            "format": file_format,
            "path": file_path or filename  # сохраняем полный путь
        }
        # Если документ с таким именем уже есть, заменяем
        for i, d in enumerate(self.current_dialog["loaded_documents"]):
            if d["filename"] == filename:
                self.current_dialog["loaded_documents"][i] = doc
                if self._auto_save:
                    self.save_dialog()
                return
        self.current_dialog["loaded_documents"].append(doc)
        if self._auto_save:
            self.save_dialog()

    def get_loaded_documents(self) -> List[Dict]:
        """Возвращает список загруженных документов."""
        return self.current_dialog.get("loaded_documents", [])

    def clear_loaded_documents(self):
        """Очищает список загруженных документов."""
        self.current_dialog["loaded_documents"] = []
        if self._auto_save:
            self.save_dialog()

    def add_loaded_scenario(self, path: str, data: dict, name: str = None):
        """Добавляет загруженный сценарий в диалог."""
        if name is None:
            name = Path(path).stem
        scenario = {"path": path, "data": data, "name": name}
        # Если сценарий с таким путём уже есть, заменяем
        for i, s in enumerate(self.current_dialog["loaded_scenarios"]):
            if s["path"] == path:
                self.current_dialog["loaded_scenarios"][i] = scenario
                if self._auto_save:
                    self.save_dialog()
                return
        self.current_dialog["loaded_scenarios"].append(scenario)
        if self._auto_save:
            self.save_dialog()

    def get_loaded_scenarios(self) -> List[Dict]:
        """Возвращает список загруженных сценариев."""
        return self.current_dialog.get("loaded_scenarios", [])

    def get_last_scenario(self) -> Optional[Dict]:
        """Возвращает последний загруженный сценарий."""
        scenarios = self.get_loaded_scenarios()
        return scenarios[-1] if scenarios else None

    def get_last_document(self) -> Optional[Dict]:
        """Возвращает последний загруженный документ."""
        docs = self.get_loaded_documents()
        return docs[-1] if docs else None

    def clear_loaded_scenarios(self):
        """Очищает список загруженных сценариев."""
        self.current_dialog["loaded_scenarios"] = []
        if self._auto_save:
            self.save_dialog()

    def set_last_result(self, result: str, fmt: str):
        """Сохраняет последний результат анализа."""
        self.current_dialog["last_analysis_result"] = {
            "result": result,
            "format": fmt
        }
        if self._auto_save:
            self.save_dialog()

    def get_last_result(self) -> Optional[Dict]:
        """Возвращает последний результат анализа."""
        return self.current_dialog.get("last_analysis_result")

    # ---------- Существующие методы ----------
    def get_conversation_pairs(self, n: int) -> List[tuple]:
        pairs = []
        msgs = self.current_dialog["messages"]
        i = 0
        while i < len(msgs) - 1:
            if msgs[i]["role"] == "user" and msgs[i+1]["role"] == "assistant":
                pairs.append((msgs[i]["content"], msgs[i+1]["content"]))
                i += 2
            else:
                i += 1
        return pairs[-n:]

    def get_full_conversation_text(self) -> str:
        return "\n".join(f"{m['role']}: {m['content']}" for m in self.current_dialog["messages"])

    def _auto_update_summary(self):
        if not self.chat_client:
            return
        dialog_params = self.config_manager.get_dialog_params()
        summary_pairs = dialog_params.get("summary_pairs", 3)
        pairs = self.get_conversation_pairs(summary_pairs)
        if not pairs:
            return
        conv_text = "Предыдущее обобщение:\n" + self.current_dialog.get("summary", "") + "\n\n"
        for i, (q, a) in enumerate(pairs):
            conv_text += f"Вопрос {i+1}: {q}\nОтвет {i+1}: {a}\n"
        if self.current_dialog.get("attached_files"):
            conv_text += "\nСодержимое загруженных файлов:\n"
            for f in self.current_dialog["attached_files"]:
                conv_text += f"--- {f['filename']} ---\n{f['content']}\n"
        def generate():
            summary_model = dialog_params.get("summary_model", "gpt-4o-mini")
            summary = self.chat_client.generate_summary(conv_text, model=summary_model)
            if summary:
                self.current_dialog["summary"] = summary
                if self._auto_save:
                    self.save_dialog()
        threading.Thread(target=generate, daemon=True).start()