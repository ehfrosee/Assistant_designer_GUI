# -*- coding: utf-8 -*-
"""Виджет чата с поддержкой форматирования, Ctrl+C/X/V, загрузкой/сохранением промптов и контекстным меню"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Frame, Text, Menu, ttk, END, DISABLED, NORMAL
from datetime import datetime

import ttkbootstrap as tb
from ttkbootstrap.constants import *


class ChatWidget(Frame):
    """Виджет чата с областью сообщений, полем ввода и кнопками управления"""

    def __init__(self, parent, on_send_callback, main_window):
        super().__init__(parent)
        self.parent = parent
        self.main_window = main_window
        self.on_send_callback = on_send_callback

        self.paned = ttk.PanedWindow(self, orient=VERTICAL)
        self.paned.pack(fill=BOTH, expand=True)

        # Область чата
        self.chat_area = scrolledtext.ScrolledText(
            self.paned, wrap=WORD, state=DISABLED,
            font=("Segoe UI", 10)
        )
        self.paned.add(self.chat_area, weight=3)

        self._setup_chat_context_menu()
        self.chat_area.bind("<Control-KeyPress>", self._on_ctrl_key_chat)

        # Панель ввода и кнопки
        input_frame = Frame(self.paned)
        self.paned.add(input_frame, weight=1)

        self.input_field = Text(input_frame, height=4, wrap=WORD, font=("Segoe UI", 10))
        self.input_field.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 5))

        self.input_field.bind("<Control-KeyPress>", self._on_ctrl_key_input)

        self.send_btn = tb.Button(
            input_frame, text="Отправить", bootstyle="primary",
            command=self._send_message
        )
        self.send_btn.pack(side=RIGHT)
        self.input_field.bind("<Control-Return>", lambda e: self._send_message())

        # Кнопки загрузки/сохранения промпта
        self.load_prompt_btn = tb.Button(
            input_frame, text="📂", bootstyle="secondary",
            command=self._load_prompt_from_dialog, width=3
        )
        self.load_prompt_btn.pack(side=RIGHT, padx=1)

        self.save_prompt_btn = tb.Button(
            input_frame, text="💾", bootstyle="secondary",
            command=self._save_prompt_to_dialog, width=3
        )
        self.save_prompt_btn.pack(side=RIGHT, padx=1)

        # Теги форматирования
        self.chat_area.tag_config("user_header", font=("Segoe UI", 9, "bold"), foreground="blue")
        self.chat_area.tag_config("assistant_header", font=("Segoe UI", 9, "bold"), foreground="green")
        self.chat_area.tag_config("user_text", font=("Segoe UI", 10), lmargin1=10, lmargin2=10, rmargin=10)
        self.chat_area.tag_config("assistant_text", font=("Segoe UI", 10), lmargin1=10, lmargin2=10, rmargin=10)

    # ---------- Контекстное меню чата ----------
    def _setup_chat_context_menu(self):
        self.chat_context_menu = Menu(self.chat_area, tearoff=0)
        self.chat_context_menu.add_command(label="Копировать выделенное", command=self._copy_selected_from_chat)
        self.chat_context_menu.add_command(label="Копировать всё", command=self._copy_all_from_chat)
        self.chat_context_menu.add_separator()
        self.chat_context_menu.add_command(label="Сохранить выделенное как файл", command=self._save_selected_as_file)
        self.chat_context_menu.add_command(label="Сохранить выделенное по разделам", command=self._save_selected_as_sections)
        self.chat_area.bind("<Button-3>", self._show_chat_context_menu)

    def _show_chat_context_menu(self, event):
        try:
            self.chat_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.chat_context_menu.grab_release()

    # ---------- Команды контекстного меню ----------
    def _copy_selected_from_chat(self):
        try:
            self.chat_area.configure(state=NORMAL)
            selected = self.chat_area.get("sel.first", "sel.last")
            self.chat_area.configure(state=DISABLED)
            if selected:
                self.clipboard_clear()
                self.clipboard_append(selected)
        except:
            pass

    def _copy_all_from_chat(self):
        self.chat_area.configure(state=NORMAL)
        all_text = self.chat_area.get("1.0", END)
        self.chat_area.configure(state=DISABLED)
        if all_text.strip():
            self.clipboard_clear()
            self.clipboard_append(all_text)

    def _save_selected_as_file(self):
        if hasattr(self.main_window, '_save_selected_as_file'):
            self.main_window._save_selected_as_file()
        else:
            messagebox.showerror("Ошибка", "Метод сохранения не найден в главном окне.")

    def _save_selected_as_sections(self):
        if hasattr(self.main_window, '_save_selected_as_sections'):
            self.main_window._save_selected_as_sections()
        else:
            messagebox.showerror("Ошибка", "Метод сохранения по разделам не найден в главном окне.")

    # ---------- Обработчики Ctrl+Key ----------
    def _on_ctrl_key_chat(self, event):
        if event.keycode == 67:  # C
            self._copy_from_chat()
            return "break"
        return None

    def _on_ctrl_key_input(self, event):
        if event.keycode == 67:  # C
            self._copy_from_input()
            return "break"
        elif event.keycode == 88:  # X
            self._cut_from_input()
            return "break"
        elif event.keycode == 86:  # V
            self._paste_to_input()
            return "break"
        return None

    # ---------- Вспомогательные функции для копирования/вставки ----------
    def _copy_from_chat(self):
        try:
            self.chat_area.configure(state=NORMAL)
            selected = self.chat_area.get("sel.first", "sel.last")
            self.chat_area.configure(state=DISABLED)
            if selected:
                self.clipboard_clear()
                self.clipboard_append(selected)
        except:
            pass

    def _copy_from_input(self):
        try:
            self.input_field.event_generate("<<Copy>>")
        except:
            try:
                selected = self.input_field.get("sel.first", "sel.last")
                if selected:
                    self.clipboard_clear()
                    self.clipboard_append(selected)
            except:
                pass

    def _cut_from_input(self):
        try:
            self._copy_from_input()
            self.input_field.delete("sel.first", "sel.last")
        except:
            pass

    def _paste_to_input(self):
        try:
            self.input_field.event_generate("<<Paste>>")
        except:
            try:
                text = self.clipboard_get()
                self.input_field.insert("insert", text)
            except:
                pass

    # ---------- Отправка сообщения ----------
    def _send_message(self):
        text = self.input_field.get("1.0", END).strip()
        if text:
            self.on_send_callback(text)
            self.input_field.delete("1.0", END)

    def add_message(self, text: str, sender: str):
        self.chat_area.configure(state=NORMAL)
        timestamp = datetime.now().strftime("%H:%M")
        header = "Вы" if sender == "user" else "Ассистент"
        header_tag = "user_header" if sender == "user" else "assistant_header"
        text_tag = "user_text" if sender == "user" else "assistant_text"
        self.chat_area.insert(END, f"{header}  {timestamp}\n", header_tag)
        self.chat_area.insert(END, f"{text}\n", text_tag)
        self.chat_area.insert(END, "-" * 60 + "\n", "separator")
        self.chat_area.tag_config("separator", foreground="gray", font=("Segoe UI", 8))
        self.chat_area.see(END)
        self.chat_area.configure(state=DISABLED)

    def get_chat_text(self) -> str:
        self.chat_area.configure(state=NORMAL)
        text = self.chat_area.get("1.0", END)
        self.chat_area.configure(state=DISABLED)
        return text.strip()

    def get_selected_text(self) -> str:
        try:
            if self.chat_area.tag_ranges("sel"):
                return self.chat_area.get("sel.first", "sel.last")
            else:
                return ""
        except:
            return ""

    def clear(self):
        self.chat_area.configure(state=NORMAL)
        self.chat_area.delete("1.0", END)
        self.chat_area.configure(state=DISABLED)

    def set_send_enabled(self, enabled: bool):
        self.send_btn.configure(state=NORMAL if enabled else DISABLED)

    # ---------- Загрузка/сохранение промптов ----------
    def _load_prompt_from_dialog(self):
        file_path = filedialog.askopenfilename(
            title="Загрузить промпт",
            filetypes=[("Текстовые файлы", "*.txt *.md *.prompt"), ("Все файлы", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                self.input_field.delete("1.0", END)
                self.input_field.insert("1.0", content)
                if hasattr(self.main_window, 'add_system_message'):
                    self.main_window.add_system_message(f"📂 Промпт загружен из {file_path}")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить промпт:\n{e}")

    def _save_prompt_to_dialog(self):
        text = self.input_field.get("1.0", END).strip()
        if not text:
            messagebox.showwarning("Сохранение", "Поле ввода пустое")
            return
        file_path = filedialog.asksaveasfilename(
            defaultextension=".md",
            filetypes=[("Markdown файлы", "*.md"), ("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
            title="Сохранить промпт как Markdown"
        )
        if not file_path:
            return
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(text)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить промпт:\n{e}")
            return
        if hasattr(self.main_window, 'add_system_message'):
            self.main_window.add_system_message(f"💾 Промпт сохранён в {file_path}")