# -*- coding: utf-8 -*-
"""
Визуальный редактор сценариев анализа документов (JSON)
Поддерживает создание, редактирование, импорт/экспорт сценариев.
"""

import json
import re
from pathlib import Path
from tkinter import filedialog, messagebox, Toplevel, END, INSERT, NORMAL, DISABLED
from tkinter import ttk

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import QueryDialog
from ttkbootstrap.scrolled import ScrolledFrame

from orchestrator.scenario_manager import ScenarioManager


class ScenarioEditor:
    """Редактор сценариев"""

    def __init__(self, parent, config_manager=None):
        self.parent = parent
        self.config_manager = config_manager
        self.window = tb.Toplevel(parent)
        self.window.title("Редактор сценариев")
        self.window.geometry("1000x700")
        self.window.minsize(800, 500)

        # Текущие данные сценария
        self.scenario_data = {
            "scenario_name": "Новый сценарий",
            "stages": [],
            "final_output": {}
        }
        self.current_file_path = None
        self.stage_edit_mode = False  # режим редактирования этапа (True) или добавления (False)
        self.editing_stage_id = None   # id редактируемого этапа

        self._build_ui()
        self._update_stage_list()

    def _build_ui(self):
        """Построение интерфейса"""
        # Основной контейнер с панелями
        main_frame = tb.Frame(self.window)
        main_frame.pack(fill=BOTH, expand=True, padx=5, pady=5)

        # Верхняя панель (название сценария и кнопки)
        top_frame = tb.Frame(main_frame)
        top_frame.pack(fill=X, pady=(0, 5))

        tb.Label(top_frame, text="Название сценария:").pack(side=LEFT, padx=(0,5))
        self.name_entry = tb.Entry(top_frame, width=40)
        self.name_entry.pack(side=LEFT, fill=X, expand=True, padx=(0,10))
        self.name_entry.bind("<KeyRelease>", self._on_name_changed)

        btn_frame = tb.Frame(top_frame)
        btn_frame.pack(side=RIGHT)
        tb.Button(btn_frame, text="Загрузить JSON", bootstyle="secondary",
                  command=self._load_scenario).pack(side=LEFT, padx=2)
        tb.Button(btn_frame, text="Сохранить", bootstyle="primary",
                  command=self._save_scenario).pack(side=LEFT, padx=2)
        tb.Button(btn_frame, text="Сохранить как...", bootstyle="secondary",
                  command=self._save_scenario_as).pack(side=LEFT, padx=2)

        # Средняя область: список этапов (слева) и форма редактирования (справа)
        middle = tb.Frame(main_frame)
        middle.pack(fill=BOTH, expand=True, pady=5)

        # Левая панель: список этапов
        left_panel = tb.Frame(middle, width=350)
        left_panel.pack(side=LEFT, fill=Y, padx=(0,5))
        left_panel.pack_propagate(False)

        tb.Label(left_panel, text="Этапы (порядок важен)", font=("Segoe UI", 10, "bold")).pack(anchor=W)

        # Treeview для списка этапов
        columns = ("id", "input", "output_format")
        self.tree = ttk.Treeview(left_panel, columns=columns, show="headings", height=15)
        self.tree.heading("id", text="ID")
        self.tree.heading("input", text="Источник")
        self.tree.heading("output_format", text="Выходной формат")
        self.tree.column("id", width=100)
        self.tree.column("input", width=150)
        self.tree.column("output_format", width=100)
        self.tree.pack(fill=BOTH, expand=True, pady=(5,5))

        # Кнопки управления этапами
        stage_btn_frame = tb.Frame(left_panel)
        stage_btn_frame.pack(fill=X, pady=5)
        tb.Button(stage_btn_frame, text="Добавить этап", bootstyle="success",
                  command=self._add_stage).pack(side=LEFT, padx=2)
        tb.Button(stage_btn_frame, text="Редактировать", bootstyle="info",
                  command=self._edit_selected_stage).pack(side=LEFT, padx=2)
        tb.Button(stage_btn_frame, text="Удалить", bootstyle="danger",
                  command=self._delete_selected_stage).pack(side=LEFT, padx=2)
        tb.Button(stage_btn_frame, text="Вверх", bootstyle="secondary",
                  command=lambda: self._move_stage(-1)).pack(side=LEFT, padx=2)
        tb.Button(stage_btn_frame, text="Вниз", bootstyle="secondary",
                  command=lambda: self._move_stage(1)).pack(side=LEFT, padx=2)

        # Правая панель: форма редактирования этапа
        right_panel = tb.Frame(middle)
        right_panel.pack(side=RIGHT, fill=BOTH, expand=True)

        tb.Label(right_panel, text="Редактор этапа", font=("Segoe UI", 10, "bold")).pack(anchor=W)

        # Контейнер с прокруткой для формы
        form_scroll = ScrolledFrame(right_panel, autohide=True)
        form_scroll.pack(fill=BOTH, expand=True)
        form_frame = form_scroll

        # ID этапа
        row1 = tb.Frame(form_frame)
        row1.pack(fill=X, pady=5)
        tb.Label(row1, text="ID этапа:", width=15).pack(side=LEFT)
        self.stage_id_entry = tb.Entry(row1, width=30)
        self.stage_id_entry.pack(side=LEFT, fill=X, expand=True)

        # Источник входных данных
        row2 = tb.Frame(form_frame)
        row2.pack(fill=X, pady=5)
        tb.Label(row2, text="Источник:", width=15).pack(side=LEFT)
        self.source_combo = tb.Combobox(row2, values=["document", "stage."], state="readonly", width=30)
        self.source_combo.pack(side=LEFT, fill=X, expand=True)
        self.source_combo.bind("<<ComboboxSelected>>", self._on_source_changed)
        self.source_ref_entry = tb.Entry(row2, width=20)
        self.source_ref_entry.pack(side=LEFT, padx=(5,0))
        self.source_ref_entry.config(state=DISABLED)

        # Формат входных данных (не редактируется, выводится автоматически)
        row2b = tb.Frame(form_frame)
        row2b.pack(fill=X, pady=2)
        tb.Label(row2b, text="Формат входа:", width=15).pack(side=LEFT)
        self.input_format_label = tb.Label(row2b, text="(определяется автоматически)", font=("Segoe UI", 9, "italic"))
        self.input_format_label.pack(side=LEFT)

        # Промпт (многострочный)
        row3 = tb.Frame(form_frame)
        row3.pack(fill=BOTH, expand=True, pady=5)
        tb.Label(row3, text="Промпт (используйте {content}):", anchor=W).pack(fill=X)
        self.prompt_text = tb.Text(row3, height=8, wrap=WORD)
        self.prompt_text.pack(fill=BOTH, expand=True)

        # Ожидаемый формат выхода
        row4 = tb.Frame(form_frame)
        row4.pack(fill=X, pady=5)
        tb.Label(row4, text="Выходной формат:", width=15).pack(side=LEFT)
        self.output_combo = tb.Combobox(row4, values=["markdown", "json", "text"], state="readonly", width=20)
        self.output_combo.pack(side=LEFT)

        # Параметры модели (опционально)
        row5 = tb.Frame(form_frame)
        row5.pack(fill=X, pady=5)
        tb.Label(row5, text="Переопределить параметры модели (опционально):", anchor=W).pack(fill=X)
        param_frame = tb.Frame(row5)
        param_frame.pack(fill=X, pady=2)
        tb.Label(param_frame, text="max_tokens:").pack(side=LEFT, padx=(0,5))
        self.max_tokens_spin = tb.Spinbox(param_frame, from_=100, to=8000, width=8)
        self.max_tokens_spin.pack(side=LEFT, padx=(0,10))
        tb.Label(param_frame, text="temperature:").pack(side=LEFT, padx=(0,5))
        self.temp_spin = tb.Spinbox(param_frame, from_=0.0, to=2.0, increment=0.1, width=6)
        self.temp_spin.pack(side=LEFT)

        # Кнопки управления этапом
        btn_frame2 = tb.Frame(form_frame)
        btn_frame2.pack(fill=X, pady=10)
        self.save_stage_btn = tb.Button(btn_frame2, text="Сохранить этап", bootstyle="primary",
                                        command=self._save_current_stage)
        self.save_stage_btn.pack(side=LEFT, padx=2)
        self.cancel_stage_btn = tb.Button(btn_frame2, text="Отменить", bootstyle="secondary",
                                          command=self._clear_form)
        self.cancel_stage_btn.pack(side=LEFT, padx=2)

        # Нижняя панель: настройки финального вывода
        bottom_frame = tb.Frame(main_frame)
        bottom_frame.pack(fill=X, pady=5)

        tb.Label(bottom_frame, text="Финальный вывод:").pack(side=LEFT, padx=(0,5))
        self.final_source_combo = tb.Combobox(bottom_frame, values=["last_stage"] + [], state="readonly", width=30)
        self.final_source_combo.pack(side=LEFT, fill=X, expand=True, padx=(0,10))
        self.final_source_combo.bind("<<ComboboxSelected>>", self._on_final_source_changed)
        self.final_format_combo = tb.Combobox(bottom_frame, values=["markdown", "json", "text"], state="readonly", width=10)
        self.final_format_combo.pack(side=LEFT)

        # Статусная строка
        self.status_var = tb.StringVar(value="Готов")
        status_bar = tb.Label(self.window, textvariable=self.status_var, bootstyle="info", anchor=W)
        status_bar.pack(side=BOTTOM, fill=X)

        # Инициализация
        self._update_final_source_list()
        self._clear_form()
        self.tree.bind("<<TreeviewSelect>>", self._on_stage_selected)

    def _on_name_changed(self, event=None):
        self.scenario_data["scenario_name"] = self.name_entry.get()

    def _update_stage_list(self):
        """Обновляет список этапов в Treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, stage in enumerate(self.scenario_data["stages"]):
            stage_id = stage.get("id", f"stage_{idx}")
            input_source = stage.get("input", {}).get("source", "")
            output_fmt = stage.get("output_format", "text")
            self.tree.insert("", END, values=(stage_id, input_source, output_fmt), tags=(idx,))
        self._update_final_source_list()

    def _update_final_source_list(self):
        """Обновляет выпадающий список для выбора финального вывода"""
        stage_ids = [s["id"] for s in self.scenario_data["stages"]]
        values = ["last_stage"] + [f"stage.{sid}" for sid in stage_ids]
        self.final_source_combo["values"] = values
        # Установка текущего значения
        final = self.scenario_data.get("final_output", {})
        if final.get("source"):
            self.final_source_combo.set(final["source"])
        else:
            self.final_source_combo.set("last_stage")
        self.final_format_combo.set(final.get("format", "markdown"))

    def _on_final_source_changed(self, event=None):
        source = self.final_source_combo.get()
        if source == "last_stage":
            self.scenario_data["final_output"] = {"source": "last_stage", "format": self.final_format_combo.get()}
        else:
            self.scenario_data["final_output"] = {"source": source, "format": self.final_format_combo.get()}

    def _clear_form(self):
        """Очищает форму редактирования этапа"""
        self.stage_id_entry.delete(0, END)
        self.source_combo.set("")
        self.source_ref_entry.delete(0, END)
        self.source_ref_entry.config(state=DISABLED)
        self.prompt_text.delete("1.0", END)
        self.output_combo.set("")
        self.max_tokens_spin.delete(0, END)
        self.temp_spin.delete(0, END)
        self.stage_edit_mode = False
        self.editing_stage_id = None
        self.save_stage_btn.config(text="Сохранить этап")
        self.input_format_label.config(text="(определяется автоматически)")

    def _add_stage(self):
        """Подготовка формы для добавления нового этапа"""
        self._clear_form()
        self.stage_edit_mode = False
        self.editing_stage_id = None
        self.save_stage_btn.config(text="Добавить этап")
        # Предлагаем ID по умолчанию
        base_id = "stage"
        counter = 1
        existing_ids = [s["id"] for s in self.scenario_data["stages"]]
        new_id = base_id
        while new_id in existing_ids:
            new_id = f"{base_id}_{counter}"
            counter += 1
        self.stage_id_entry.insert(0, new_id)

    def _edit_selected_stage(self):
        """Загружает выбранный этап в форму для редактирования"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Редактирование", "Выберите этап для редактирования")
            return
        idx = int(self.tree.item(selection[0], "tags")[0])
        stage = self.scenario_data["stages"][idx]
        self._clear_form()
        self.stage_edit_mode = True
        self.editing_stage_id = stage["id"]
        self.save_stage_btn.config(text="Обновить этап")
        # Заполняем поля
        self.stage_id_entry.insert(0, stage["id"])
        self.stage_id_entry.config(state="readonly")  # ID нельзя менять при редактировании
        source = stage["input"].get("source", "")
        if source.startswith("stage."):
            self.source_combo.set("stage.")
            ref = source.split(".", 1)[1]
            self.source_ref_entry.config(state=NORMAL)
            self.source_ref_entry.delete(0, END)
            self.source_ref_entry.insert(0, ref)
            self.source_ref_entry.config(state=DISABLED)
        else:
            self.source_combo.set("document")
            self.source_ref_entry.delete(0, END)
            self.source_ref_entry.config(state=DISABLED)
        self.prompt_text.insert("1.0", stage.get("prompt_template", ""))
        self.output_combo.set(stage.get("output_format", "text"))
        # Параметры модели
        params = stage.get("params", {})
        self.max_tokens_spin.delete(0, END)
        self.max_tokens_spin.insert(0, str(params.get("max_tokens", "")))
        self.temp_spin.delete(0, END)
        self.temp_spin.insert(0, str(params.get("temperature", "")))
        # Формат входа выводится (если определён)
        in_fmt = stage["input"].get("format", "не указан")
        self.input_format_label.config(text=in_fmt)

    def _delete_selected_stage(self):
        selection = self.tree.selection()
        if not selection:
            return
        idx = int(self.tree.item(selection[0], "tags")[0])
        if messagebox.askyesno("Удаление", f"Удалить этап '{self.scenario_data['stages'][idx]['id']}'?"):
            del self.scenario_data["stages"][idx]
            self._update_stage_list()
            self._clear_form()
            self.status_var.set("Этап удалён")

    def _move_stage(self, direction):
        """Перемещение этапа вверх/вниз"""
        selection = self.tree.selection()
        if not selection:
            return
        idx = int(self.tree.item(selection[0], "tags")[0])
        new_idx = idx + direction
        if new_idx < 0 or new_idx >= len(self.scenario_data["stages"]):
            return
        stages = self.scenario_data["stages"]
        stages[idx], stages[new_idx] = stages[new_idx], stages[idx]
        self._update_stage_list()
        # Выделяем перемещённый элемент
        for item in self.tree.get_children():
            if int(self.tree.item(item, "tags")[0]) == new_idx:
                self.tree.selection_set(item)
                break

    def _on_stage_selected(self, event):
        """При выборе этапа в дереве показываем его (необязательно)"""
        pass

    def _on_source_changed(self, event=None):
        """При изменении источника показываем/скрываем поле ссылки"""
        if self.source_combo.get() == "stage.":
            self.source_ref_entry.config(state=NORMAL)
        else:
            self.source_ref_entry.config(state=DISABLED)

    def _save_current_stage(self):
        """Сохраняет этап из формы в данные сценария"""
        stage_id = self.stage_id_entry.get().strip()
        if not stage_id:
            messagebox.showerror("Ошибка", "ID этапа не может быть пустым")
            return
        # Проверка уникальности ID
        existing_ids = [s["id"] for s in self.scenario_data["stages"]]
        if self.stage_edit_mode:
            if stage_id != self.editing_stage_id and stage_id in existing_ids:
                messagebox.showerror("Ошибка", f"ID '{stage_id}' уже существует")
                return
        else:
            if stage_id in existing_ids:
                messagebox.showerror("Ошибка", f"ID '{stage_id}' уже существует")
                return

        source_type = self.source_combo.get()
        if source_type == "document":
            source = "document"
        elif source_type == "stage.":
            ref = self.source_ref_entry.get().strip()
            if not ref:
                messagebox.showerror("Ошибка", "Укажите ID этапа после 'stage.'")
                return
            source = f"stage.{ref}"
        else:
            messagebox.showerror("Ошибка", "Выберите источник")
            return

        prompt = self.prompt_text.get("1.0", END).strip()
        if not prompt:
            messagebox.showerror("Ошибка", "Промпт не может быть пустым")
            return
        output_fmt = self.output_combo.get()
        if not output_fmt:
            messagebox.showerror("Ошибка", "Выберите выходной формат")
            return

        # Параметры модели
        params = {}
        if self.max_tokens_spin.get():
            try:
                params["max_tokens"] = int(self.max_tokens_spin.get())
            except:
                pass
        if self.temp_spin.get():
            try:
                params["temperature"] = float(self.temp_spin.get())
            except:
                pass

        stage_data = {
            "id": stage_id,
            "input": {"source": source},
            "prompt_template": prompt,
            "output_format": output_fmt
        }
        if params:
            stage_data["params"] = params

        if self.stage_edit_mode:
            # Находим индекс и заменяем
            for idx, st in enumerate(self.scenario_data["stages"]):
                if st["id"] == self.editing_stage_id:
                    self.scenario_data["stages"][idx] = stage_data
                    break
        else:
            self.scenario_data["stages"].append(stage_data)

        self._update_stage_list()
        self._clear_form()
        self.status_var.set(f"Этап '{stage_id}' сохранён")
        # После сохранения пересчитаем список для финального вывода
        self._update_final_source_list()

    def _load_scenario(self):
        """Загружает сценарий из JSON файла"""
        file_path = filedialog.askopenfilename(
            title="Загрузить сценарий",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not file_path:
            return
        try:
            data = ScenarioManager.load_scenario(file_path)  # используем валидатор
            self.scenario_data = data
            self.current_file_path = file_path
            self.name_entry.delete(0, END)
            self.name_entry.insert(0, data.get("scenario_name", Path(file_path).stem))
            self._update_stage_list()
            self._clear_form()
            self._update_final_source_list()
            self.status_var.set(f"Загружен: {file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка загрузки", str(e))

    def _save_scenario(self):
        """Сохраняет текущий сценарий в файл (если путь известен, иначе вызывает Save as)"""
        if self.current_file_path:
            self._write_scenario(self.current_file_path)
        else:
            self._save_scenario_as()

    def _save_scenario_as(self):
        file_path = filedialog.asksaveasfilename(
            title="Сохранить сценарий",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not file_path:
            return
        self.current_file_path = file_path
        self._write_scenario(file_path)

    def _write_scenario(self, file_path):
        """Записывает сценарий в файл"""
        # Обновляем название
        self.scenario_data["scenario_name"] = self.name_entry.get()
        # Обновляем финальный вывод
        self._on_final_source_changed()
        # Валидация перед сохранением (проверка ссылок)
        try:
            ScenarioManager._validate(self.scenario_data)  # используем внутренний метод валидации
        except Exception as e:
            messagebox.showerror("Ошибка валидации", str(e))
            return
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(self.scenario_data, f, ensure_ascii=False, indent=2)
        self.status_var.set(f"Сценарий сохранён: {file_path}")
        messagebox.showinfo("Сохранение", f"Сценарий сохранён в {file_path}")


def launch_scenario_editor(parent, config_manager=None):
    """Функция для запуска редактора из главного окна"""
    editor = ScenarioEditor(parent, config_manager)
    editor.window.grab_set()
    return editor