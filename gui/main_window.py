# -*- coding: utf-8 -*-
"""Главное окно приложения AI Document Analyst (Chat)"""

import logging
import threading
import os
import glob
import json
import re
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, Frame, ttk, Menu
from tkinter import END, DISABLED, NORMAL, LEFT, RIGHT, X, Y, BOTH
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# ---- Импорты из других модулей проекта ----
from config.config_manager import ConfigManager
from core.chat_client import ChatGPTClient
from gui.file_loader import FileLoader
from gui.dialog_manager import DialogManager
from orchestrator.dispatcher import run_analysis

# Импорты виджетов и хендлеров
from gui.chat_widget import ChatWidget
from gui.settings_dialog import SettingsDialog
from gui.export_handlers import ExportHandlers
from gui.selection_handlers import SelectionHandlers


class MainWindow(tb.Window):
    """Главное окно приложения, объединяющее чат, команды и настройки"""

    # ---------- Инициализация и настройка ----------
    def __init__(self, config_manager: ConfigManager):
        theme = config_manager.config.get("ui", {}).get("theme", "light")
        theme_name = "darkly" if theme == "dark" else "cosmo"
        super().__init__(themename=theme_name)

        self.config_manager = config_manager
        self.chat_client = ChatGPTClient(config_manager)
        self.dialog_manager = DialogManager(config_manager, self.chat_client)
        self._stream_buffer = ""
        self.logger = logging.getLogger("main_window")
        self.analysis_thread = None
        self.cancel_event = None
        self.result_format = "text"

        self.setup_logging()

        # === Создаём хендлеры ДО вызова init_ui ===
        self.export_handlers = ExportHandlers(self)
        self.selection_handlers = SelectionHandlers(self)

        # Перенаправляем методы экспорта
        self._parse_markdown_table = self.export_handlers._parse_markdown_table
        self._convert_csv_to_table = self.export_handlers._convert_csv_to_table
        self._is_likely_csv = self.export_handlers._is_likely_csv
        self._convert_csv_block_to_table = self.export_handlers._convert_csv_block_to_table
        self._apply_inline_formatting = self.export_handlers._apply_inline_formatting
        self._json_to_word_table = self.export_handlers._json_to_word_table
        self._save_as_docx = self.export_handlers.save_as_docx
        self._save_as_xlsx = self.export_handlers.save_as_xlsx
        self._save_chat_as_docx = self.export_handlers.save_chat_as_docx

        # Перенаправляем методы сохранения выделенного
        self._safe_filename = self.selection_handlers._safe_filename
        self._detect_list_type = self.selection_handlers._detect_list_type
        self._save_selected_as_file = self.selection_handlers.save_selected_as_file
        self._save_selected_as_sections = self.selection_handlers.save_selected_as_sections
        self._save_selected_as_docx = self.selection_handlers.save_selected_as_docx

        # Теперь можно вызывать init_ui
        self.init_ui()
        self.bind_signals()
        self.check_api_key()
        self.after(500, self.test_api_connection)
        self.after(1000, self._auto_load_prompt)

        self.available_models = []
        self.load_models_at_startup()
        self.attached_files = []  # список кортежей (filename, content)

        # Инициализация хендлеров
        self.export_handlers = ExportHandlers(self)
        self.selection_handlers = SelectionHandlers(self)

        # Перенаправление методов экспорта
        self._parse_markdown_table = self.export_handlers._parse_markdown_table
        self._convert_csv_to_table = self.export_handlers._convert_csv_to_table
        self._is_likely_csv = self.export_handlers._is_likely_csv
        self._convert_csv_block_to_table = self.export_handlers._convert_csv_block_to_table
        self._apply_inline_formatting = self.export_handlers._apply_inline_formatting
        self._json_to_word_table = self.export_handlers._json_to_word_table
        self._save_as_docx = self.export_handlers.save_as_docx
        self._save_as_xlsx = self.export_handlers.save_as_xlsx
        self._save_chat_as_docx = self.export_handlers.save_chat_as_docx

        # Перенаправление методов сохранения выделенного
        self._safe_filename = self.selection_handlers._safe_filename
        self._detect_list_type = self.selection_handlers._detect_list_type
        self._save_selected_as_file = self.selection_handlers.save_selected_as_file
        self._save_selected_as_sections = self.selection_handlers.save_selected_as_sections
        self._save_selected_as_docx = self.selection_handlers.save_selected_as_docx

    def setup_logging(self):
        """Настройка логирования в файл и консоль"""
        logs_dir = self.config_manager.config.get("paths", {}).get("logs_dir", "./logs")
        Path(logs_dir).mkdir(exist_ok=True)
        log_file = Path(logs_dir) / "app.log"
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            handlers=[
                logging.FileHandler(log_file, encoding="utf-8"),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger("main_window")
        logging.getLogger("chat_client").setLevel(logging.DEBUG)

    # ---------- Построение пользовательского интерфейса ----------
    def init_ui(self):
        self.title("AI Document Analyst (Chat)")
        self.geometry("900x600")
        self.minsize(600, 400)

        # ---- Меню ----
        menubar = tb.Menu(self)
        self.config(menu=menubar)

        file_menu = tb.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Выход", command=self.destroy)

        dialog_menu = tb.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Диалог", menu=dialog_menu)
        dialog_menu.add_command(label="Новый диалог", command=self.new_dialog)
        dialog_menu.add_command(label="Сохранить диалог", command=self.save_dialog)
        dialog_menu.add_command(label="Загрузить диалог", command=self.load_dialog)
        dialog_menu.add_command(label="Обобщить диалог (summary)", command=self.summarize_dialog)
        dialog_menu.add_separator()
        dialog_menu.add_command(label="Очистить загруженные документы", command=self._cmd_clear_documents)
        dialog_menu.add_command(label="Очистить загруженные сценарии", command=self._cmd_clear_scenarios)
        dialog_menu.add_separator()
        dialog_menu.add_command(label="Загрузить промпт", command=self._load_prompt_global)
        dialog_menu.add_command(label="Сохранить промпт", command=self._save_prompt_global)

        tools_menu = tb.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Инструменты", menu=tools_menu)
        tools_menu.add_command(label="Копировать чат", command=self.copy_chat)
        tools_menu.add_command(label="Сохранить чат в DOCX", command=self._save_chat_as_docx_wrapper)
        tools_menu.add_separator()
        tools_menu.add_command(label="Сохранить выделенное как файл", command=self._save_selected_as_file)
        tools_menu.add_command(label="Сохранить выделенное по разделам", command=self._save_selected_as_sections)
        tools_menu.add_command(label="Сохранить выделенное как DOCX", command=self._save_selected_as_docx)
        tools_menu.add_separator()
        tools_menu.add_command(label="Очистить чат", command=self.clear_chat)
        tools_menu.add_command(label="Настройки", command=self.open_settings)

        commands_menu = tb.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Команды", menu=commands_menu)
        commands_menu.add_command(label="📄 Загрузить документ", command=self._cmd_load_document_dialog)
        commands_menu.add_command(label="📋 Загрузить сценарий", command=self._cmd_load_scenario_dialog)
        commands_menu.add_separator()
        commands_menu.add_command(label="▶️ Запустить сценарий", command=self._cmd_run_scenario)
        commands_menu.add_command(label="⏹️ Остановить сценарий", command=self._cmd_stop_scenario)
        commands_menu.add_separator()
        commands_menu.add_command(label="📁 Импортировать папку", command=self._cmd_import_folder_dialog)

        export_menu = tb.Menu(commands_menu, tearoff=0)
        commands_menu.add_cascade(label="💾 Экспортировать результат", menu=export_menu)
        for fmt, label in [("json", "JSON"), ("csv", "CSV"), ("md", "Markdown (MD)"),
                           ("docx", "Word (DOCX)"), ("xlsx", "Excel (XLSX)")]:
            export_menu.add_command(label=label, command=lambda f=fmt: self._cmd_export_result(f))

        commands_menu.add_separator()
        commands_menu.add_command(label="🗑️ Очистить документы", command=self._cmd_clear_documents)
        commands_menu.add_command(label="🗑️ Очистить сценарии", command=self._cmd_clear_scenarios)
        commands_menu.add_separator()
        commands_menu.add_command(label="❓ Помощь", command=self._cmd_help)

        help_menu = tb.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Помощь", menu=help_menu)
        help_menu.add_command(label="О программе", command=self.show_about)

        # ---- Статусная строка (самая нижняя) ----
        self.statusbar = tb.Label(self, text="Готов", bootstyle="info", anchor=W)
        self.statusbar.pack(side=BOTTOM, fill=X)

        # ---- Панель загрузки файлов (над статусной) ----
        file_frame = Frame(self)
        file_frame.pack(fill=X, side=BOTTOM, padx=5, pady=2)
        file_frame.config(height=35)
        file_frame.pack_propagate(False)

        self.attach_btn = tb.Button(
            file_frame, text="📎 Прикрепить файл", bootstyle="secondary",
            command=self.load_file
        )
        self.attach_btn.pack(side=LEFT, padx=2)

        self.file_label = tb.Label(
            file_frame,
            text="Файлы не загружены",
            bootstyle="info",
            anchor=W
        )
        self.file_label.pack(side=LEFT, fill=X, expand=True, padx=10)

        self.clear_file_btn = tb.Button(
            file_frame, text="✕ Убрать", bootstyle="danger",
            command=self.clear_file, state=DISABLED,
            width=10
        )
        self.clear_file_btn.pack(side=RIGHT, padx=2)

        # ---- Виджет чата (занимает всё оставшееся пространство) ----
        self.chat_frame = ttk.Frame(self)
        self.chat_frame.pack(fill=BOTH, expand=True, padx=0, pady=0)
        self.chat_widget = ChatWidget(self.chat_frame, self.on_user_message, self)
        self.chat_widget.pack(fill=BOTH, expand=True)

        # ---- Приветственное сообщение ----
        self.chat_widget.add_message("Приложение запущено. Введите вопрос.", "assistant")

    # ---------- Сигналы и колбэки ----------
    def bind_signals(self):
        self.chat_client.on_response = self._on_ai_response
        self.chat_client.on_chunk = self._on_ai_chunk
        self.chat_client.on_error = self._on_api_error
        self.chat_client.on_start = self._on_request_start
        self.chat_client.on_finish = self._on_request_finish

    # ---------- Обработчики событий чата ----------
    def on_user_message(self, text: str):
        if text.startswith('/'):
            self._handle_command(text)
            return

        self.chat_widget.add_message(text, "user")
        self.dialog_manager.add_message("user", text)

        dialog_params = self.config_manager.get_dialog_params()
        history_pairs = dialog_params.get("history_pairs", 3)
        pairs = self.dialog_manager.get_conversation_pairs(history_pairs)
        history_text = ""
        for i, (q, a) in enumerate(pairs):
            history_text += f"Вопрос {i + 1}: {q}\nОтвет {i + 1}: {a}\n"
        summary = self.dialog_manager.current_dialog.get("summary", "")

        files_text = ""
        loaded_docs = self.dialog_manager.get_loaded_documents()
        for doc in loaded_docs:
            files_text += f"--- {doc['filename']} ---\n{doc['content']}\n"
        for fname, content in self.attached_files:
            if not any(doc["filename"] == fname for doc in loaded_docs):
                files_text += f"--- {fname} ---\n{content}\n"

        prompts = self.chat_client._load_prompts()
        system_prompt = prompts.get("dialog_system_prompt", "Ты полезный ассистент.")
        user_prompt_template = prompts.get(
            "dialog_user_prompt",
            "Краткое содержание предыдущих сообщений (summary):\n{summary}\n\n"
            "Последние вопросы и ответы:\n{history}\n\n"
            "Содержимое загруженных файлов:\n{files}\n\n"
            "Текущий вопрос пользователя: {question}"
        )
        user_prompt = user_prompt_template.format(
            summary=summary,
            history=history_text,
            files=files_text,
            question=text
        )

        # Проверка размера файла
        if files_text:
            import tiktoken
            try:
                model = self.config_manager.get_llm_params().get("model", "gpt-4o-mini")
                encoding = tiktoken.encoding_for_model(model)
                total_tokens = len(encoding.encode(files_text))
                self.logger.info(f"📊 Размер файла в токенах: {total_tokens}")
                max_safe_tokens = 100000 if ("gpt-4" in model or "gpt-5" in model) else 50000
                if total_tokens > max_safe_tokens:
                    recommendation = self._recommend_file_split(files_text, model)
                    if recommendation:
                        self.chat_widget.add_message(recommendation, "assistant")
                        self.dialog_manager.add_message("assistant", recommendation)
                    else:
                        self.chat_widget.add_message("⚠️ Файл слишком большой. Попробуйте разбить его на части и загружать по отдельности.", "assistant")
                    return
            except Exception as e:
                self.logger.warning(f"Ошибка при оценке размера файла: {e}")

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]
        self._stream_buffer = ""
        self.chat_client.send_message(messages, stream=True)

    # ---------- Обработка команд ----------
    def _handle_command(self, text: str):
        parts = text.split(maxsplit=1)
        cmd = parts[0].lower()
        args = parts[1] if len(parts) > 1 else ""

        commands = {
            "/help": self._cmd_help,
            "/load_document": lambda: self._cmd_load_document(args),
            "/load_scenario": lambda: self._cmd_load_scenario(args),
            "/run_scenario": self._cmd_run_scenario,
            "/stop_scenario": self._cmd_stop_scenario,
            "/export_result": lambda: self._cmd_export_result(args),
            "/import_folder": lambda: self._cmd_import_folder(args),
            "/clear_documents": self._cmd_clear_documents,
            "/clear_scenarios": self._cmd_clear_scenarios,
        }
        if cmd in commands:
            commands[cmd]()
        else:
            self.chat_widget.add_message(f"❌ Неизвестная команда: {cmd}. Введите /help.", "assistant")
            self.dialog_manager.add_message("assistant", f"Неизвестная команда: {cmd}")

    # ---------- Реализация команд (меню и чат) ----------
    def _cmd_help(self):
        help_text = (
            "**Доступные команды:**\n\n"
            "`/load_document <путь>` — загрузить документ в контекст чата\n"
            "`/load_scenario <путь>` — загрузить JSON-сценарий\n"
            "`/run_scenario` — выполнить загруженный сценарий на загруженном документе\n"
            "`/export_result <формат>` — экспортировать результат (json, csv, md, docx, xlsx)\n"
            "`/import_folder <путь>` — загрузить все документы из папки\n"
            "`/clear_documents` — очистить загруженные документы\n"
            "`/clear_scenarios` — очистить загруженные сценарии\n"
            "`/help` — показать эту справку"
        )
        self.chat_widget.add_message(help_text, "assistant")
        self.dialog_manager.add_message("assistant", help_text)

    def _cmd_load_document(self, file_path: str):
        if not file_path:
            self.chat_widget.add_message("❌ Укажите путь к файлу: /load_document <путь>", "assistant")
            return
        file_path = file_path.strip()
        if os.path.isdir(file_path):
            self._cmd_import_folder(file_path)
            return
        if ',' in file_path or ';' in file_path:
            separator = ',' if ',' in file_path else ';'
            paths = [p.strip().strip('"').strip("'") for p in file_path.split(separator) if p.strip()]
            count = 0
            for p in paths:
                p = p.strip('"').strip("'")
                if os.path.exists(p):
                    self._load_single_document(p)
                    count += 1
                else:
                    self.chat_widget.add_message(f"⚠️ Файл не найден: {p}", "assistant")
            if count > 0:
                self.chat_widget.add_message(f"✅ Загружено {count} документов", "assistant")
            return
        file_path = file_path.strip('"').strip("'")
        self._load_single_document(file_path)

    def _load_single_document(self, file_path: str):
        if not os.path.exists(file_path):
            self.chat_widget.add_message(f"⚠️ Файл не найден: {file_path}", "assistant")
            return
        try:
            result = FileLoader._read_file(file_path, self.config_manager)
            if result:
                filename, content, file_format = result
                self.dialog_manager.add_loaded_document(filename, content, file_format, file_path)
                self.attached_files.append((filename, content))
                self._update_file_label()
                self.chat_widget.add_message(f"✅ Документ '{filename}' загружен в контекст", "assistant")
                self.dialog_manager.add_message("assistant", f"Документ '{filename}' загружен в контекст")
                # Сохраняем конвертированный файл в папку сценария, если сценарий загружен
                scenario_path = None
                last_scenario = self.dialog_manager.get_last_scenario()
                if last_scenario:
                    scenario_path = last_scenario.get("path")
                if scenario_path:
                    scenario_dir = Path(scenario_path).parent
                    input_dir = scenario_dir / "input"
                    input_dir.mkdir(parents=True, exist_ok=True)
                    md_filename = Path(filename).stem + ".md"
                    md_path = input_dir / md_filename
                    with open(md_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                    self.chat_widget.add_message(f"*Система*: Конвертированный файл сохранён в {md_path}", "assistant")
            else:
                self.chat_widget.add_message(f"❌ Не удалось загрузить документ: {file_path}", "assistant")
        except Exception as e:
            self.chat_widget.add_message(f"❌ Ошибка загрузки: {e}", "assistant")

    def _cmd_load_scenario(self, file_path: str):
        if not file_path:
            self.chat_widget.add_message("❌ Укажите путь к сценарию: /load_scenario <путь>", "assistant")
            return
        file_path = file_path.strip().strip('"').strip("'")
        if not os.path.exists(file_path):
            self.chat_widget.add_message(f"❌ Файл сценария не найден: {file_path}", "assistant")
            return
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if 'stages' not in data:
                    raise ValueError("Нет ключа 'stages'")
            self.dialog_manager.add_loaded_scenario(file_path, data)
            self.chat_widget.add_message(f"✅ Сценарий '{Path(file_path).stem}' загружен", "assistant")
            self.dialog_manager.add_message("assistant", f"Сценарий '{Path(file_path).stem}' загружен")
        except Exception as e:
            self.chat_widget.add_message(f"❌ Ошибка загрузки сценария: {e}", "assistant")

    def _cmd_run_scenario(self):
        scenario = self.dialog_manager.get_last_scenario()
        document = self.dialog_manager.get_last_document()
        if not scenario:
            self.chat_widget.add_message("❌ Сценарий не загружен. Используйте /load_scenario", "assistant")
            return
        if not document:
            self.chat_widget.add_message("❌ Документ не загружен. Используйте /load_document", "assistant")
            return
        doc_path = document.get("path")
        doc_paths = [doc_path]
        scenario_path = scenario.get("path")
        if not doc_path or not os.path.exists(doc_path):
            self.chat_widget.add_message("⚠️ Путь к документу не найден. Используйте /load_document с полным путём", "assistant")
            return
        if not scenario_path or not os.path.exists(scenario_path):
            self.chat_widget.add_message("⚠️ Путь к сценарию не найден. Используйте /load_scenario с полным путём", "assistant")
            return
        api_key = self.config_manager.get_api_key()
        if not api_key:
            self.chat_widget.add_message("❌ API ключ не найден. Задайте его в настройках.", "assistant")
            return
        llm_params = self.config_manager.get_llm_params()
        api_settings = {
            "api_key": api_key,
            "model": llm_params.get("model", "gpt-4o-mini"),
            "temperature": llm_params.get("temperature", 0.2),
            "max_tokens": llm_params.get("max_tokens", 2000),
            "timeout": llm_params.get("timeout", 60)
        }
        self.chat_widget.add_message("🔄 Запуск анализа по сценарию...", "assistant")
        self.dialog_manager.add_message("assistant", "Запуск анализа по сценарию...")
        self.cancel_event = threading.Event()
        self.analysis_thread = threading.Thread(
            target=self._run_scenario_thread,
            args=(doc_paths, scenario_path, api_settings),
            daemon=True
        )
        self.analysis_thread.start()

    def _run_scenario_thread(self, doc_paths, scenario_path, api_settings):
        def chat_callback(msg_type, data):
            if msg_type == "log":
                self.after(0, lambda: self._append_chat_message(f"*Система*: {data}"))
            elif msg_type == "stage":
                current = data.get("current", 0)
                total = data.get("total", 1)
                name = data.get("name", "unknown")
                self.after(0, lambda: self._append_chat_message(f"*Система*: Этап {current}/{total}: {name}"))
            elif msg_type == "stage_result":
                stage_id = data.get("stage_id", "unknown")
                formatted = data.get("formatted_content")
                if formatted:
                    self.after(0, lambda: self._append_chat_message(f"*Система*: Результат этапа '{stage_id}':\n{formatted}"))
                else:
                    full_result = data.get("full_result", "")
                    self.after(0, lambda: self._append_chat_message(f"*Система*: Результат этапа '{stage_id}':\n{full_result}"))

        try:
            result = run_analysis(
                document_paths=doc_paths,
                scenario_path=scenario_path,
                api_settings=api_settings,
                progress_callback=chat_callback,
                cancel_event=self.cancel_event
            )
            self.after(0, lambda: self._on_scenario_finished(result))
        except Exception as e:
            self.after(0, lambda: self._append_chat_message(f"❌ Критическая ошибка: {e}"))

    def _append_chat_message(self, text: str):
        self.chat_widget.add_message(text, "assistant")
        self.dialog_manager.add_message("assistant", text)

    def _on_scenario_finished(self, result):
        if result is None:
            self._append_chat_message("❌ Ошибка: результат анализа пустой")
            self.analysis_thread = None
            return
        if not isinstance(result, dict):
            self._append_chat_message(f"❌ Ошибка: неверный формат результата: {result}")
            self.analysis_thread = None
            return
        status = result.get("status")
        if status == "success":
            self.dialog_manager.set_last_result(result["result"], result.get("format", "text"))
            self._append_chat_message(f"✅ Анализ завершён успешно! Результат в формате {result.get('format', 'text')}")
            self._append_chat_message(f"**Результат:**\n{result.get('result', '')}")
        elif status == "cancelled":
            self._append_chat_message("⚠️ Анализ отменён пользователем")
        else:
            error_msg = result.get("message", "Неизвестная ошибка")
            self._append_chat_message(f"❌ Ошибка: {error_msg}")
        self.analysis_thread = None

    def add_system_message(self, text: str):
        self.chat_widget.add_message(f"*Система*: {text}", "assistant")
        self.dialog_manager.add_message("assistant", f"*Система*: {text}")

    def _cmd_export_result(self, format_type: str):
        if not format_type:
            self.chat_widget.add_message("❌ Укажите формат: /export_result <json|csv|md|docx|xlsx>", "assistant")
            return
        format_type = format_type.strip().lower()
        valid = {"json", "csv", "md", "docx", "xlsx"}
        if format_type not in valid:
            self.chat_widget.add_message(f"❌ Неподдерживаемый формат: {format_type}. Доступные: {', '.join(valid)}", "assistant")
            return
        result = self.dialog_manager.get_last_result()
        if not result:
            self.chat_widget.add_message("❌ Нет сохранённых результатов анализа", "assistant")
            return

        ext_map = {"json": ".json", "csv": ".csv", "md": ".md", "docx": ".docx", "xlsx": ".xlsx"}
        file_path = filedialog.asksaveasfilename(
            defaultextension=ext_map[format_type],
            filetypes=[(f"{format_type.upper()} files", f"*{ext_map[format_type]}"), ("All files", "*.*")],
            title=f"Сохранить результат как {format_type.upper()}"
        )
        if not file_path:
            return

        try:
            ext = Path(file_path).suffix.lower()
            content = result["result"]
            if ext == '.docx':
                doc = self._save_as_docx(content, result.get("format", "text"))
                doc.save(file_path)
            elif ext == '.xlsx':
                wb = self._save_as_xlsx(content, result.get("format", "text"))
                wb.save(file_path)
            elif ext == '.json':
                try:
                    data = json.loads(content)
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False, indent=2)
                except json.JSONDecodeError:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(content)
                self.chat_widget.add_message(f"✅ Результат сохранён: {file_path}", "assistant")
                self.dialog_manager.add_message("assistant", f"Результат сохранён: {file_path}")
                self.statusbar.config(text=f"Результат сохранён: {file_path}")
            else:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                self.chat_widget.add_message(f"✅ Результат сохранён: {file_path}", "assistant")
                self.dialog_manager.add_message("assistant", f"Результат сохранён: {file_path}")
                self.statusbar.config(text=f"Результат сохранён: {file_path}")
        except Exception as e:
            self.chat_widget.add_message(f"❌ Ошибка сохранения: {e}", "assistant")
            self.logger.error(f"Ошибка экспорта: {e}")

    def _cmd_import_folder(self, folder_path: str):
        if not folder_path:
            self.chat_widget.add_message("❌ Укажите путь к папке: /import_folder <путь>", "assistant")
            return
        folder_path = folder_path.strip().strip('"').strip("'")
        if not os.path.isdir(folder_path):
            self.chat_widget.add_message(f"❌ Папка не найдена: {folder_path}", "assistant")
            return
        supported_ext = {'.docx', '.xlsx', '.pdf', '.xls', '.txt', '.md', '.json', '.csv'}
        count = 0
        errors = []
        for ext in supported_ext:
            for file_path in glob.glob(os.path.join(folder_path, f"*{ext}")):
                try:
                    result = FileLoader._read_file(file_path, self.config_manager)
                    if result:
                        filename, content, file_format = result
                        self.dialog_manager.add_loaded_document(filename, content, file_format, file_path)
                        self.attached_files.append((filename, content))
                        count += 1
                except Exception as e:
                    errors.append(f"{Path(file_path).name}: {e}")
        self._update_file_label()
        if count > 0:
            self.chat_widget.add_message(f"✅ Загружено {count} документов из папки", "assistant")
            self.dialog_manager.add_message("assistant", f"Загружено {count} документов из папки")
        if errors:
            self.chat_widget.add_message(f"⚠️ Ошибки при загрузке:\n" + "\n".join(errors), "assistant")

    def _cmd_clear_documents(self):
        self.dialog_manager.clear_loaded_documents()
        self.attached_files = []
        self._update_file_label()
        self.chat_widget.add_message("✅ Загруженные документы очищены", "assistant")
        self.dialog_manager.add_message("assistant", "Загруженные документы очищены")

    def _cmd_clear_scenarios(self):
        self.dialog_manager.clear_loaded_scenarios()
        self.chat_widget.add_message("✅ Загруженные сценарии очищены", "assistant")
        self.dialog_manager.add_message("assistant", "Загруженные сценарии очищены")

    def _cmd_stop_scenario(self):
        if self.analysis_thread and self.analysis_thread.is_alive():
            if self.cancel_event:
                self.cancel_event.set()
                self.chat_widget.add_message("⏹️ Запрос на остановку сценария отправлен...", "assistant")
                self.dialog_manager.add_message("assistant", "Запрос на остановку сценария отправлен...")
                self.statusbar.config(text="Остановка сценария...")
            else:
                self.chat_widget.add_message("⚠️ Сценарий выполняется, но событие отмены не найдено.", "assistant")
        else:
            self.chat_widget.add_message("ℹ️ Нет выполняющегося сценария для остановки.", "assistant")

    # ---------- Диалоги выбора файлов (меню) ----------
    def _cmd_load_document_dialog(self):
        file_path = filedialog.askopenfilename(
            title="Выберите документ для загрузки",
            filetypes=[
                ("Поддерживаемые", "*.docx *.xlsx *.pdf *.xls *.txt *.md *.json *.csv"),
                ("Документы Word", "*.docx"),
                ("Таблицы Excel", "*.xlsx *.xls"),
                ("PDF", "*.pdf"),
                ("Текстовые", "*.txt *.md *.json *.csv"),
                ("Все файлы", "*.*")
            ]
        )
        if file_path:
            self._load_single_document(file_path)

    def _cmd_load_scenario_dialog(self):
        file_path = filedialog.askopenfilename(
            title="Выберите JSON-сценарий",
            filetypes=[("JSON files", "*.json")]
        )
        if file_path:
            self._cmd_load_scenario(file_path)

    def _cmd_import_folder_dialog(self):
        folder_path = filedialog.askdirectory(title="Выберите папку с документами")
        if folder_path:
            self._cmd_import_folder(folder_path)

    # ---------- Загрузка/сохранение промптов ----------
    def _load_prompt_global(self):
        self.chat_widget._load_prompt_from_dialog()

    def _save_prompt_global(self):
        self.chat_widget._save_prompt_to_dialog()

    def _auto_load_prompt(self):
        prompt_defaults = self.config_manager.get_prompt_defaults()
        auto_file = prompt_defaults.get("auto_load_file", "")
        if auto_file and os.path.exists(auto_file):
            try:
                with open(auto_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.chat_widget.input_field.delete("1.0", END)
                self.chat_widget.input_field.insert("1.0", content)
                self.add_system_message(f"📂 Автозагрузка промпта из {auto_file}")
            except Exception as e:
                self.logger.error(f"Ошибка автозагрузки промпта: {e}")

    # ---------- Работа с файлами ----------
    def load_file(self):
        result = FileLoader.load_from_dialog(self, self.config_manager)
        if result:
            filename, content, file_format = result
            self.attached_files.append((filename, content))
            self.dialog_manager.add_attached_file(filename, content)
            file_path = self._find_file_by_name(filename)
            self.dialog_manager.add_loaded_document(filename, content, file_format, file_path)
            self._update_file_label()
            self.statusbar.config(text=f"Файл '{filename}' добавлен")
            self.chat_widget.add_message(f"*Система*: Загружен файл \"{filename}\".", "assistant")
            saved_path = self._save_converted_file(filename, content, file_format)
            if saved_path:
                self.chat_widget.add_message(f"*Система*: Извлечённый текст сохранён в {saved_path}", "assistant")
        else:
            messagebox.showerror("Ошибка", "Не удалось загрузить файл.\nПоддерживаются: docx, xlsx, pdf, txt, md, py, json, csv")

    def _find_file_by_name(self, filename: str) -> str:
        if os.path.exists(filename):
            return filename
        docs_dir = self.config_manager.config.get("input_settings", {}).get("default_input_directory", "./documents")
        full_path = os.path.join(docs_dir, filename)
        if os.path.exists(full_path):
            return full_path
        converted_dir = self.config_manager.get_paths().get("converted_files_dir", "./converted_files")
        full_path = os.path.join(converted_dir, filename)
        if os.path.exists(full_path):
            return full_path
        return filename

    def _update_file_label(self):
        if self.attached_files:
            names = ", ".join(f[0] for f in self.attached_files)
            if len(names) > 60:
                names = names[:57] + "..."
            self.file_label.config(text=f"📎 Файлы: {names}")
            self.clear_file_btn.config(state=NORMAL)
        else:
            self.file_label.config(text="📎 Файлы не загружены")
            self.clear_file_btn.config(state=DISABLED)

    def clear_file(self):
        self.attached_files = []
        self.dialog_manager.clear_attached_files()
        self._update_file_label()
        self.statusbar.config(text="Все файлы удалены")
        self.chat_widget.add_message("*Система*: Все загруженные файлы удалены.", "assistant")

    def _save_converted_file(self, original_filename: str, content: str, file_format: str) -> Optional[str]:
        try:
            save_dir = self.config_manager.get_paths().get("converted_files_dir", "./converted_files")
            Path(save_dir).mkdir(parents=True, exist_ok=True)
            base = Path(original_filename).stem
            ext = ".json" if file_format == "json" else ".md"
            save_path = Path(save_dir) / f"{base}{ext}"
            counter = 1
            while save_path.exists():
                save_path = Path(save_dir) / f"{base}_{counter}{ext}"
                counter += 1
            with open(save_path, 'w', encoding='utf-8') as f:
                f.write(content)
            self.logger.info(f"Сохранён конвертированный файл: {save_path}")
            return str(save_path)
        except Exception as e:
            self.logger.error(f"Ошибка сохранения конвертированного файла: {e}")
            return None

    # ---------- Работа с диалогами ----------
    def new_dialog(self):
        if self.dialog_manager.current_dialog["messages"]:
            reply = messagebox.askyesno("Новый диалог", "Текущий диалог будет потерян. Продолжить?")
            if not reply:
                return
        self.dialog_manager.new_dialog()
        self.chat_widget.clear()
        self.attached_files = []
        self._update_file_label()
        self.statusbar.config(text="Новый диалог создан")

    def save_dialog(self):
        if not self.dialog_manager.current_dialog["messages"]:
            messagebox.showwarning("Сохранение", "Нет сообщений для сохранения.")
            return
        for fname, content in self.attached_files:
            self.dialog_manager.add_attached_file(fname, content)
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=self.dialog_manager.dialogs_dir,
            initialfile=f"{self.dialog_manager.current_dialog['dialog_id']}.json"
        )
        if file_path:
            self.dialog_manager.save_dialog(file_path)
            self.statusbar.config(text=f"Диалог сохранён в {file_path}")

    def load_dialog(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            initialdir=self.dialog_manager.dialogs_dir
        )
        if file_path:
            self.dialog_manager.load_dialog(file_path)
            self.chat_widget.clear()
            for msg in self.dialog_manager.current_dialog["messages"]:
                self.chat_widget.add_message(msg["content"], msg["role"])
            self.attached_files = [(f["filename"], f["content"]) for f in
                                   self.dialog_manager.current_dialog.get("attached_files", [])]
            self._update_file_label()
            self.statusbar.config(text=f"Диалог загружен: {file_path}")

    def summarize_dialog(self):
        if not self.dialog_manager.current_dialog["messages"]:
            messagebox.showwarning("Обобщение", "Нет сообщений для обобщения.")
            return
        self.statusbar.config(text="Генерация summary...")
        dialog_params = self.config_manager.get_dialog_params()
        summary_pairs = dialog_params.get("summary_pairs", 3)
        pairs = self.dialog_manager.get_conversation_pairs(summary_pairs)
        conv_text = "Предыдущее обобщение:\n" + self.dialog_manager.current_dialog.get("summary", "") + "\n\n"
        for i, (q, a) in enumerate(pairs):
            conv_text += f"Вопрос {i+1}: {q}\nОтвет {i+1}: {a}\n"
        if self.attached_files:
            conv_text += "\nСодержимое загруженных файлов:\n"
            for fname, content in self.attached_files:
                conv_text += f"--- {fname} ---\n{content}\n"
        def generate():
            summary = self.chat_client.generate_summary(conv_text)
            self.after(0, lambda: self._on_summary_ready(summary))
        threading.Thread(target=generate, daemon=True).start()

    def _on_summary_ready(self, summary):
        self.dialog_manager.current_dialog["summary"] = summary
        self.dialog_manager.save_dialog()
        self.chat_widget.add_message(f"*Система*: Обобщение диалога:\n{summary}", "assistant")
        self.statusbar.config(text="Summary готов")

    # ---------- Сохранение чата в DOCX ----------
    def _save_chat_as_docx_wrapper(self):
        chat_text = self.chat_widget.get_chat_text()
        if not chat_text:
            messagebox.showwarning("Сохранение", "Чат пуст, нечего сохранять")
            return
        file_path = self._save_chat_as_docx(chat_text)
        if file_path:
            self.statusbar.config(text=f"Чат сохранён в DOCX: {file_path}")
            self.add_system_message(f"📄 Чат сохранён в DOCX: {file_path}")

    # ---------- Копирование чата ----------
    def copy_chat(self):
        chat_text = self.chat_widget.get_chat_text()
        if chat_text:
            self.clipboard_clear()
            self.clipboard_append(chat_text)
            self.statusbar.config(text="Чат скопирован в буфер обмена")
            self.chat_widget.add_message("*Система*: Содержимое чата скопировано в буфер обмена.", "assistant")
        else:
            self.statusbar.config(text="Чат пуст, нечего копировать")

    def clear_chat(self):
        self.chat_widget.clear()
        self.statusbar.config(text="Чат очищен")
        self.chat_widget.add_message("*Система*: История чата очищена.", "assistant")

    # ---------- Обработчики API ----------
    def _on_ai_chunk(self, chunk):
        self.after(0, lambda: self._append_chunk(chunk))

    def _append_chunk(self, chunk):
        self._stream_buffer += chunk

    def _on_ai_response(self, text, metadata):
        self.after(0, lambda: self._display_response(text))

    def _display_response(self, text):
        full = text if text else self._stream_buffer
        if full:
            self.chat_widget.add_message(full, "assistant")
            self.dialog_manager.add_message("assistant", full, tokens_used=0)
        else:
            self.chat_widget.add_message("*Ассистент не дал ответа*", "assistant")
        self._stream_buffer = ""

    def _on_api_error(self, msg, recoverable):
        self.after(0, lambda: self._show_error(msg, recoverable))

    def _show_error(self, msg, recoverable):
        self.logger.error(f"API error: {msg}")
        self.chat_widget.add_message(f"*Ошибка*: {msg}", "assistant")
        if not recoverable:
            self.statusbar.config(text="Критическая ошибка API. Проверьте ключ.")
        else:
            self.statusbar.config(text="Ошибка API, повторите попытку.")

    def _on_request_start(self):
        self.after(0, lambda: self._set_sending_state(True))

    def _on_request_finish(self):
        self.after(0, lambda: self._set_sending_state(False))

    def _set_sending_state(self, is_sending):
        if is_sending:
            self.chat_widget.set_send_enabled(False)
            self.statusbar.config(text="Ассистент печатает...")
        else:
            self.chat_widget.set_send_enabled(True)
            self.statusbar.config(text="Готов")

    # ---------- Настройки, проверка API, загрузка моделей ----------
    def open_settings(self):
        dialog = SettingsDialog(self, self.config_manager, self.available_models)
        self.wait_window(dialog)
        theme = self.config_manager.config.get("ui", {}).get("theme", "light")
        new_theme = "darkly" if theme == "dark" else "cosmo"
        self.style.theme_use(new_theme)

    def check_api_key(self):
        if not self.config_manager.get_api_key():
            reply = messagebox.askyesno("API ключ", "OPENAI_API_KEY не найден. Ввести сейчас?")
            if reply:
                from tkinter import simpledialog
                key = simpledialog.askstring("API ключ", "Введите OpenAI API ключ:", show='*')
                if key:
                    self.config_manager.set_api_key(key)
                    self.statusbar.config(text="API ключ сохранён")
                    self.chat_widget.add_message("*Система*: API ключ успешно сохранён.", "assistant")
                else:
                    self.statusbar.config(text="API ключ не задан")
            else:
                self.statusbar.config(text="API ключ не задан")

    def test_api_connection(self):
        def test():
            success = self.chat_client.test_connection()
            self.after(0, lambda: self._on_test_result(success))
        threading.Thread(target=test, daemon=True).start()

    def _on_test_result(self, success):
        if success:
            self.chat_widget.add_message("*Система*: ✅ Соединение с OpenAI API установлено.", "assistant")
            self.statusbar.config(text="API готов")
        else:
            self.chat_widget.add_message("*Система*: ❌ Не удалось подключиться к API. Проверьте ключ и интернет.", "assistant")
            self.statusbar.config(text="Ошибка подключения к API")

    def load_models_at_startup(self):
        def load():
            try:
                api_key = self.config_manager.get_api_key()
                if not api_key:
                    self.logger.warning("Не удалось загрузить модели: нет API ключа")
                    return
                from openai import OpenAI
                client = OpenAI(api_key=api_key)
                models = client.models.list()
                model_names = [m.id for m in models if 'gpt' in m.id and not m.id.startswith('whisper')]
                priority = ["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-4"]
                sorted_models = sorted(model_names, key=lambda x: (x not in priority, priority.index(x) if x in priority else len(priority)))
                self.available_models = sorted_models
                self.logger.info(f"Загружено {len(sorted_models)} моделей")
            except Exception as e:
                self.logger.error(f"Ошибка загрузки моделей: {e}")
                self.available_models = ["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-4"]
        threading.Thread(target=load, daemon=True).start()

    def show_about(self):
        messagebox.showinfo(
            "О программе",
            "AI Document Analyst (Chat)\nВерсия 1.0\n\n"
            "Агентная система анализа документов с ИИ.\n"
            "Использует OpenAI API для ответов на вопросы.\n\n© 2024"
        )

    # ---------- Рекомендация по разбиению файла ----------
    def _recommend_file_split(self, content: str, model: str = "gpt-4o-mini") -> str:
        try:
            import tiktoken
            encoding = tiktoken.encoding_for_model(model)
            total_tokens = len(encoding.encode(content))
            max_safe_tokens = 100000 if ("gpt-4" in model or "gpt-5" in model) else 50000
            if total_tokens <= max_safe_tokens:
                return None
            parts = (total_tokens // max_safe_tokens) + 1
            tokens_per_part = total_tokens // parts + 1
            return (
                f"⚠️ **Файл слишком большой для модели {model}.**\n"
                f"Размер: ~{total_tokens:,} токенов.\n"
                f"Рекомендация: разбейте файл на **{parts}** частей (примерно по {tokens_per_part:,} токенов каждая)\n"
                f"и загружайте их по очереди, задавая вопросы к каждой части отдельно.\n"
                f"Или используйте команду /load_document для загрузки отдельных файлов."
            )
        except Exception as e:
            self.logger.error(f"Ошибка подсчёта токенов: {e}")
            return None