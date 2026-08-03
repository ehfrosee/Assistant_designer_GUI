# -*- coding: utf-8 -*-
"""Асинхронный клиент для работы с OpenAI ChatGPT API (threading, callbacks)"""

import logging
import threading
from typing import Dict, List, Optional, Any

from openai import OpenAI, APIError, APIConnectionError, RateLimitError, AuthenticationError, BadRequestError


class ChatGPTClient:
    """Асинхронный клиент ChatGPT с колбэками и адаптацией параметров под модель"""

    # Список префиксов моделей, требующих адаптации (без учёта регистра)
    NEW_MODEL_PREFIXES = ('o1', 'o3', 'gpt-4o', 'gpt-5')

    def __init__(self, config_manager):
        self.config_manager = config_manager
        self.logger = logging.getLogger("chat_client")
        self._thread: Optional[threading.Thread] = None
        self._stop_requested = False

        # Колбэки (устанавливаются из GUI)
        self.on_response = None  # func(text, metadata)
        self.on_chunk = None  # func(chunk)
        self.on_error = None  # func(error_msg, recoverable)
        self.on_start = None  # func()
        self.on_finish = None  # func()

    def _get_api_key_and_timeout(self):
        api_key = self.config_manager.get_api_key()
        timeout = self.config_manager.get_llm_params().get("timeout", 60)
        return api_key, timeout

    def _prepare_params_for_model(self, params: dict, model: str) -> dict:
        """
        Адаптирует параметры запроса под конкретную модель:
        - для новых моделей (o1, o3, gpt-4o, gpt-5): заменяет max_tokens на max_completion_tokens,
          удаляет temperature, top_p, presence_penalty, frequency_penalty.
        - для остальных моделей: оставляет без изменений.
        """
        prepared = params.copy()
        model_lower = model.lower()
        is_new = any(model_lower.startswith(prefix) for prefix in self.NEW_MODEL_PREFIXES)

        if is_new:
            # 1. max_tokens -> max_completion_tokens
            if 'max_tokens' in prepared:
                prepared['max_completion_tokens'] = prepared.pop('max_tokens')
                self.logger.debug(f"Для модели {model}: max_tokens заменён на max_completion_tokens")
            # 2. Удаляем неподдерживаемые параметры
            for param in ('temperature', 'top_p', 'presence_penalty', 'frequency_penalty'):
                if param in prepared:
                    del prepared[param]
                    self.logger.debug(f"Для модели {model}: удалён параметр {param}")
        else:
            # Обратная конвертация (если вдруг пришло max_completion_tokens)
            if 'max_completion_tokens' in prepared:
                prepared['max_tokens'] = prepared.pop('max_completion_tokens')
        return prepared

    def send_message(self, messages: List[dict], override_params: Optional[dict] = None,
                     stream: bool = False) -> None:
        """Отправляет сообщение асинхронно"""
        if self._thread and self._thread.is_alive():
            self.logger.warning("Previous request still running")
            if self.on_error:
                self.on_error("Предыдущий запрос ещё выполняется", True)
            return

        api_key, timeout = self._get_api_key_and_timeout()
        if not api_key:
            if self.on_error:
                self.on_error("OPENAI_API_KEY не найден в .env", False)
            return

        params = self._get_merged_params(override_params)
        self._stop_requested = False

        if self.on_start:
            self.on_start()

        self._thread = threading.Thread(
            target=self._worker,
            args=(api_key, timeout, messages, params, stream),
            daemon=True
        )
        self._thread.start()

    def _worker(self, api_key: str, timeout: int, messages: List[dict],
                params: dict, stream: bool):
        """Выполняется в отдельном потоке"""
        client = None
        try:
            client = OpenAI(api_key=api_key, timeout=timeout)
            model = params.get('model', 'gpt-4o-mini')
            # Адаптируем параметры под модель
            adapted_params = self._prepare_params_for_model(params, model)

            # --- Логирование запроса (ДО вызова API) ---
            self.logger.debug(f"API Request: model={model}, params={adapted_params}, messages_count={len(messages)}")

            if stream:
                full_response = ""
                self.logger.debug(
                    f"API Request: model={model}, params={adapted_params}, messages_count={len(messages)}")
                response = client.chat.completions.create(
                    messages=messages,
                    stream=True,
                    **adapted_params
                )
                for chunk in response:
                    if self._stop_requested:
                        return
                    if chunk.choices and chunk.choices[0].delta.content:
                        content = chunk.choices[0].delta.content
                        full_response += content
                        if self.on_chunk:
                            self._safe_callback(self.on_chunk, content)
                if self.on_response:
                    self._safe_callback(self.on_response, full_response, {"model": model})
            else:
                response = client.chat.completions.create(
                    messages=messages,
                    stream=False,
                    **adapted_params
                )
                if self._stop_requested:
                    return
                text = response.choices[0].message.content
                # --- Логирование ответа (после получения) ---
                self.logger.debug(f"API Response: {text[:200]}...")
                metadata = {
                    "model": response.model,
                    "usage": response.usage.model_dump() if response.usage else {}
                }
                if self.on_response:
                    self._safe_callback(self.on_response, text, metadata)

        except BadRequestError as e:
            self.logger.error(f"BadRequestError: {e}")
            if self.on_error:
                self.on_error(f"Ошибка запроса: {e}", False)
        except (APIError, APIConnectionError, RateLimitError, AuthenticationError) as e:
            recoverable = isinstance(e, (RateLimitError, APIConnectionError))
            self.logger.error(f"API error: {e}")
            if self.on_error:
                self.on_error(str(e), recoverable)
        except Exception as e:
            self.logger.exception(f"Unexpected error: {e}")
            if self.on_error:
                self.on_error(str(e), False)
        finally:
            if client:
                del client
            if self.on_finish:
                self._safe_callback(self.on_finish)
            self._thread = None

    def _safe_callback(self, callback, *args):
        try:
            callback(*args)
        except Exception as e:
            self.logger.error(f"Callback error: {e}")

    def cancel_current_request(self) -> None:
        self._stop_requested = True
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)

    def _get_merged_params(self, override: Optional[dict]) -> dict:
        defaults = self.config_manager.get_llm_params()
        defaults.pop("timeout", None)
        if override:
            merged = {**defaults, **override}
        else:
            merged = defaults.copy()
        return merged

    def test_connection(self) -> bool:
        api_key, timeout = self._get_api_key_and_timeout()
        if not api_key:
            return False
        try:
            client = OpenAI(api_key=api_key, timeout=timeout)
            client.chat.completions.create(
                messages=[{"role": "user", "content": "test"}],
                model="gpt-4o-mini",
                max_tokens=5
            )
            return True
        except Exception as e:
            self.logger.error(f"Test connection failed: {e}")
            return False

    def generate_summary(self, conversation_text: str, model: str = None) -> str:
        prompts = self._load_prompts()
        system_prompt = prompts.get("summarize_system_prompt",
                                    "Ты - нейро-саммаризатор. Составь краткое обобщение диалога.")
        user_prompt_template = prompts.get("summarize_user_prompt", "{conversation_text}")
        user_prompt = user_prompt_template.format(conversation_text=conversation_text)

        if model is None:
            model = self.config_manager.get_dialog_params().get("summary_model", "gpt-4o-mini")

        params = {"model": model, "max_tokens": 500, "temperature": 0.2}
        adapted_params = self._prepare_params_for_model(params, model)

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ]

        try:
            api_key = self.config_manager.get_api_key()
            if not api_key:
                return ""
            client = OpenAI(api_key=api_key, timeout=60)
            response = client.chat.completions.create(
                messages=messages,
                **adapted_params
            )
            summary = response.choices[0].message.content.strip()
            self.logger.info(f"Generated summary: {summary[:100]}...")
            return summary
        except Exception as e:
            self.logger.error(f"Summary generation error: {e}")
            return ""

    def _load_prompt(self, key: str) -> str:
        import json
        from pathlib import Path
        prompts_path = Path("core/prompts.json")
        if prompts_path.exists():
            with open(prompts_path, "r", encoding="utf-8") as f:
                prompts = json.load(f)
                return prompts.get(key, "")
        return ""

    def _load_prompts(self) -> dict:
        import json
        from pathlib import Path
        prompts_path = Path("core/prompts.json")
        if prompts_path.exists():
            with open(prompts_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}
