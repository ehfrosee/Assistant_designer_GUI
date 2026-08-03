# -*- coding: utf-8 -*-
"""Диалог настроек приложения (модель, температура, лимиты токенов, диалог, тема)"""

import tkinter as tk
from tkinter import Frame, Toplevel, ttk
import ttkbootstrap as tb
from ttkbootstrap.constants import *

from config.config_manager import ConfigManager


class SettingsDialog(Toplevel):
    """Окно настроек с прокруткой и группировкой параметров"""

    def __init__(self, parent, config_manager: ConfigManager, available_models: list):
        """
        Args:
            parent: родительское окно
            config_manager: менеджер конфигурации
            available_models: список доступных моделей (для выпадающих списков)
        """
        super().__init__(parent)
        self.config_manager = config_manager
        self.available_models = available_models
        self.title("Настройки")
        self.geometry("500x550")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # Основной фрейм с прокруткой
        main_frame = Frame(self)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Контейнер с прокруткой
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=canvas.yview)
        self.scrollable_frame = Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # --- Секция: Модель ---
        self._add_model_section()

        # --- Секция: Температура ---
        self._add_temperature_section()

        # --- Секция: Max tokens ---
        self._add_max_tokens_section()

        # --- Секция: Настройки диалога ---
        self._add_dialog_section()

        # --- Секция: Тема интерфейса ---
        self._add_theme_section()

        # --- Кнопки ---
        btn_frame = Frame(self.scrollable_frame)
        btn_frame.pack(fill=X, pady=(20, 10))
        tb.Button(btn_frame, text="Сохранить", bootstyle="primary", command=self._save).pack(side=RIGHT, padx=5)
        tb.Button(btn_frame, text="Отмена", bootstyle="secondary", command=self.destroy).pack(side=RIGHT, padx=5)

    # ---------- Вспомогательные методы для построения секций ----------
    def _add_model_section(self):
        tb.Label(self.scrollable_frame, text="Модель:").pack(anchor=W, pady=(0, 5))
        self.model_combo = tb.Combobox(self.scrollable_frame, state="readonly", width=50)
        self.model_combo.pack(fill=X, pady=(0, 10))

        models_to_show = self.available_models if self.available_models else ["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-4"]
        self.model_combo['values'] = models_to_show
        current_model = self.config_manager.config.get("llm_defaults", {}).get("model", "gpt-4o-mini")
        if current_model in models_to_show:
            self.model_combo.set(current_model)
        else:
            self.model_combo.set(models_to_show[0] if models_to_show else "gpt-4o-mini")

    def _add_temperature_section(self):
        tb.Label(self.scrollable_frame, text="Температура (0.0-2.0):").pack(anchor=W, pady=(0, 5))
        temp_frame = Frame(self.scrollable_frame)
        temp_frame.pack(fill=X, pady=(0, 10))
        self.temp_scale = tb.Scale(
            temp_frame,
            from_=0.0,
            to=2.0,
            value=self.config_manager.config.get("llm_defaults", {}).get("temperature", 0.3),
            orient=HORIZONTAL
        )
        self.temp_scale.pack(side=LEFT, fill=X, expand=True)
        self.temp_label = tb.Label(temp_frame, text=f"{self.temp_scale.get():.1f}", width=5)
        self.temp_label.pack(side=RIGHT, padx=(5, 0))
        self.temp_scale.configure(command=self._update_temp_label)

    def _add_max_tokens_section(self):
        tb.Label(self.scrollable_frame, text="Max tokens:").pack(anchor=W, pady=(0, 5))
        self.tokens_spin = tb.Spinbox(self.scrollable_frame, from_=100, to=8000, width=15)
        self.tokens_spin.set(str(self.config_manager.config.get("llm_defaults", {}).get("max_tokens", 2000)))
        self.tokens_spin.pack(anchor=W, pady=(0, 10))

    def _add_dialog_section(self):
        dialog_params = self.config_manager.get_dialog_params()
        tb.Label(self.scrollable_frame, text="Настройки диалога:", font=("Segoe UI", 10, "bold")).pack(anchor=W, pady=(10, 5))

        frame_dialog = Frame(self.scrollable_frame)
        frame_dialog.pack(fill=X, pady=5)

        tb.Label(frame_dialog, text="Пар для контекста:").grid(row=0, column=0, sticky=W, padx=5, pady=2)
        self.history_pairs_spin = tb.Spinbox(frame_dialog, from_=1, to=20, width=5)
        self.history_pairs_spin.set(dialog_params.get("history_pairs", 3))
        self.history_pairs_spin.grid(row=0, column=1, padx=5, pady=2)

        tb.Label(frame_dialog, text="Пар для summary:").grid(row=1, column=0, sticky=W, padx=5, pady=2)
        self.summary_pairs_spin = tb.Spinbox(frame_dialog, from_=1, to=20, width=5)
        self.summary_pairs_spin.set(dialog_params.get("summary_pairs", 3))
        self.summary_pairs_spin.grid(row=1, column=1, padx=5, pady=2)

        tb.Label(frame_dialog, text="Модель для summary:").grid(row=2, column=0, sticky=W, padx=5, pady=2)
        self.summary_model_combo = tb.Combobox(frame_dialog, state="readonly", width=30)
        self.summary_model_combo['values'] = self.available_models if self.available_models else ["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-4"]
        current_summary_model = dialog_params.get("summary_model", "gpt-4o-mini")
        if current_summary_model in self.summary_model_combo['values']:
            self.summary_model_combo.set(current_summary_model)
        else:
            self.summary_model_combo.set(self.summary_model_combo['values'][0])
        self.summary_model_combo.grid(row=2, column=1, padx=5, pady=2)

    def _add_theme_section(self):
        tb.Label(self.scrollable_frame, text="Тема:", font=("Segoe UI", 10, "bold")).pack(anchor=W, pady=(10, 5))
        self.theme_combo = tb.Combobox(self.scrollable_frame, values=["light", "dark"], state="readonly", width=20)
        self.theme_combo.set(self.config_manager.config.get("ui", {}).get("theme", "light"))
        self.theme_combo.pack(anchor=W, pady=(0, 10))

    # ---------- Обработчики ----------
    def _update_temp_label(self, *args):
        self.temp_label.config(text=f"{self.temp_scale.get():.1f}")

    def _save(self):
        """Сохраняет изменения в конфигурацию"""
        changes = {
            "llm_defaults": {
                "model": self.model_combo.get(),
                "temperature": self.temp_scale.get(),
                "max_tokens": int(self.tokens_spin.get())
            },
            "dialog": {
                "history_pairs": int(self.history_pairs_spin.get()),
                "summary_pairs": int(self.summary_pairs_spin.get()),
                "summary_model": self.summary_model_combo.get().strip()
            },
            "ui": {
                "theme": self.theme_combo.get()
            }
        }
        self.config_manager.update_config(changes)
        self.destroy()