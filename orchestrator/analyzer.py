import threading
import time
import logging
import json
import tiktoken
import re
from typing import Optional, List, Dict, Any

from openai import OpenAI, APIError, APIConnectionError, RateLimitError

logger = logging.getLogger("analyzer")

class Analyzer:
    def __init__(self, api_key: str, model: str = "gpt-4o-mini",
                 temperature: float = 0.2, max_tokens: int = 2000,
                 timeout: int = 60):
        self.client = OpenAI(api_key=api_key, timeout=timeout)
        self.model = model
        self.temperature = temperature
        self.max_tokens = max_tokens
        self.timeout = timeout

    def analyze(self,
                system_prompt: str,
                user_prompt: str,
                examples: Optional[List[tuple]] = None,
                output_format: str = "text",
                cancel_event: Optional[threading.Event] = None) -> str:
        """
        Отправляет запрос к LLM с разделением на system, user и примеры (assistant).
        """
        # Формируем список сообщений
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})

        # Добавляем примеры (чередование user/assistant)
        if examples:
            for user, assistant in examples:
                messages.append({"role": "user", "content": user})
                messages.append({"role": "assistant", "content": assistant})

        # Добавляем финальный user-запрос с данными
        messages.append({"role": "user", "content": user_prompt})

        # Оценка токенов (только для user_prompt, чтобы решить, нужен ли чанкинг)
        try:
            encoding = tiktoken.encoding_for_model(self.model)
        except KeyError:
            encoding = tiktoken.get_encoding("cl100k_base")
        total_tokens = len(encoding.encode(user_prompt))
        max_context = 20000  # безопасное значение

        if total_tokens > max_context:
            # Чанкинг: разбиваем только user_prompt, system и примеры оставляем неизменными
            chunks = self._chunk_text(user_prompt, max_context // 2)
            results = []
            for chunk in chunks:
                if cancel_event and cancel_event.is_set():
                    raise InterruptedError("Cancelled")
                # Создаём копию messages, но заменяем последнее user-сообщение на чанк
                chunk_messages = []
                for msg in messages[:-1]:  # все кроме последнего user
                    chunk_messages.append(msg.copy())
                chunk_messages.append({"role": "user", "content": chunk})
                resp = self._send_request(chunk_messages, output_format)
                results.append(resp)
            return self._aggregate_results(results, output_format)
        else:
            return self._send_request(messages, output_format)

    def _send_request(self, messages: List[Dict], output_format: str, retries: int = 3) -> str:
        for attempt in range(retries):
            try:
                params = {
                    "model": self.model,
                    "temperature": self.temperature,
                    "max_tokens": self.max_tokens,
                }
                adapted_params = self._prepare_params_for_model(params)
                response = self.client.chat.completions.create(
                    messages=messages,
                    **adapted_params
                )
                content = response.choices[0].message.content
                logger.debug(f"RAW RESPONSE (attempt {attempt+1}): {content[:200]}...")

                if output_format == "json":
                    cleaned = self._clean_json_response(content)
                    json.loads(cleaned)
                    return cleaned
                return content

            except (APIError, APIConnectionError, RateLimitError) as e:
                if attempt == retries - 1:
                    raise
                wait = 2 ** attempt
                logger.warning(f"API error, retry in {wait}s: {e}")
                time.sleep(wait)
            except json.JSONDecodeError as e:
                if attempt == retries - 1:
                    raise ValueError(f"Invalid JSON after {retries} attempts: {content}")
                logger.warning(f"JSON decode error, attempt {attempt+1}: {e}")
                try:
                    extracted = self._extract_json_from_text(content)
                    json.loads(extracted)
                    return extracted
                except:
                    pass
                # Отправляем исправляющий запрос
                fix_prompt = f"Your previous response was not valid JSON. Please respond with valid JSON only, without any markdown formatting or extra text.\nPrevious response:\n{content}\n\nCorrect JSON:"
                # Заменяем последнее user-сообщение на исправляющий промпт
                if messages and messages[-1]["role"] == "user":
                    messages[-1]["content"] = fix_prompt
                else:
                    messages.append({"role": "user", "content": fix_prompt})
                continue
        raise RuntimeError("Request failed after retries")

    def _chunk_text(self, text: str, max_tokens_per_chunk: int) -> list:
        try:
            encoding = tiktoken.encoding_for_model(self.model)
        except KeyError:
            encoding = tiktoken.get_encoding("cl100k_base")
            logger.warning(f"Model {self.model} not found in tiktoken, using cl100k_base fallback")
        paragraphs = text.split('\n\n')
        chunks = []
        current = []
        current_len = 0
        for para in paragraphs:
            para_tokens = len(encoding.encode(para))
            if current_len + para_tokens > max_tokens_per_chunk and current:
                chunks.append('\n\n'.join(current))
                current = [para]
                current_len = para_tokens
            else:
                current.append(para)
                current_len += para_tokens
        if current:
            chunks.append('\n\n'.join(current))
        return chunks

    def _aggregate_results(self, results: list, output_format: str) -> str:
        if output_format == "json":
            merged = []
            for r in results:
                try:
                    data = json.loads(r)
                    merged.append(data)
                except:
                    merged.append(r)
            return json.dumps(merged, ensure_ascii=False)
        else:
            return "\n\n".join(results)

    def _clean_json_response(self, text: str) -> str:
        import re
        pattern = r'```(?:json)?\s*\n?(.*?)\n?```'
        match = re.search(pattern, text, re.DOTALL)
        if match:
            return match.group(1).strip()
        start = text.find('{')
        if start == -1:
            start = text.find('[')
        if start != -1:
            end = text.rfind('}')
            if end == -1:
                end = text.rfind(']')
            if end != -1 and end > start:
                return text[start:end+1].strip()
        return text.strip()

    def _extract_json_from_text(self, text: str) -> str:
        import re
        match = re.search(r'(\[.*\]|\{.*\})', text, re.DOTALL)
        if match:
            return match.group(1)
        return text

    def _prepare_params_for_model(self, params: dict) -> dict:
        prepared = params.copy()
        model_lower = self.model.lower()
        new_model_prefixes = ('o1', 'o3', 'gpt-4o', 'gpt-5')
        is_new = any(model_lower.startswith(prefix) for prefix in new_model_prefixes)

        if is_new:
            if 'max_tokens' in prepared:
                prepared['max_completion_tokens'] = prepared.pop('max_tokens')
            for param in ('temperature', 'top_p', 'presence_penalty', 'frequency_penalty'):
                prepared.pop(param, None)
        else:
            if 'max_completion_tokens' in prepared:
                prepared['max_tokens'] = prepared.pop('max_completion_tokens')
        return prepared
