# -*- coding: utf-8 -*-
"""Менеджер конфигурации: загрузка/сохранение config.json и чтение .env"""

import json
import os
from pathlib import Path
from typing import Any, Dict

from dotenv import load_dotenv, set_key

from config.default_config import default_config


class ConfigManager:
    """Управляет глобальными настройками и API-ключом (без сигналов)"""

    def __init__(self, config_path: str = "config.json"):
        self.config_path = Path(config_path)
        self.config: Dict[str, Any] = {}
        self._load_config()
        self._ensure_env_file()

    def _load_config(self) -> None:
        """Загружает config.json, при отсутствии создаёт из default_config"""
        if self.config_path.exists():
            with open(self.config_path, "r", encoding="utf-8") as f:
                self.config = json.load(f)
        else:
            self.config = default_config()
            self.save_config()

    def save_config(self) -> None:
        """Сохраняет текущую конфигурацию в config.json"""
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)

    def update_config(self, changes: Dict[str, Any]) -> None:
        """Обновляет конфигурацию (без сигнала)"""
        def update_recursive(target, source):
            for key, value in source.items():
                if isinstance(value, dict) and key in target and isinstance(target[key], dict):
                    update_recursive(target[key], value)
                else:
                    target[key] = value

        update_recursive(self.config, changes)
        self.save_config()

    def _ensure_env_file(self) -> None:
        """Проверяет наличие .env, при отсутствии создаёт пустой"""
        env_path = self.config.get("paths", {}).get("env_file", ".env")
        if not Path(env_path).exists():
            Path(env_path).touch()

    def get_api_key(self) -> str:
        """Возвращает API-ключ из .env"""
        env_path = self.config.get("paths", {}).get("env_file", ".env")
        load_dotenv(env_path)
        return os.getenv("OPENAI_API_KEY", "")

    def set_api_key(self, key: str) -> None:
        """Сохраняет API-ключ в .env"""
        env_path = self.config.get("paths", {}).get("env_file", ".env")
        set_key(str(env_path), "OPENAI_API_KEY", key)

    def get_llm_params(self) -> Dict[str, Any]:
        """Возвращает параметры LLM по умолчанию"""
        return self.config.get("llm_defaults", {}).copy()

    def get_rag_config(self) -> Dict[str, Any]:
        """Возвращает настройки RAG"""
        return self.config.get("rag", {}).copy()

    def get_paths(self) -> Dict[str, str]:
        """Возвращает пути из конфига"""
        return self.config.get("paths", {}).copy()

    def get_dialog_params(self) -> Dict[str, Any]:
        """Возвращает параметры диалога"""
        return self.config.get("dialog", {}).copy()

    def get_processing_params(self) -> Dict[str, Any]:
        """Возвращает параметры обработки больших файлов."""
        return self.config.get("processing", {}).copy()

    def get_prompt_defaults(self) -> Dict[str, Any]:
        """Возвращает настройки промптов."""
        return self.config.get("prompt_defaults", {}).copy()
