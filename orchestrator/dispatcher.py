# orchestrator/dispatcher.py
import threading
import logging
import sys
import json
from pathlib import Path
from typing import Dict, Any, Callable, Optional, List, Union

sys.path.insert(0, str(Path(__file__).parent.parent))
from document_handlers.adapter import get_document_content
from orchestrator.scenario_manager import ScenarioManager
from orchestrator.analyzer import Analyzer
from orchestrator.reporter import Reporter

logger = logging.getLogger("dispatcher")


def run_analysis(document_paths: Union[str, List[str]],
                 scenario_path: str,
                 api_settings: Dict[str, Any],
                 progress_callback: Optional[Callable] = None,
                 cancel_event: Optional[threading.Event] = None) -> Dict[str, Any]:
    """
    Основная функция оркестратора.

    Args:
        document_paths: путь к файлу или список путей
        scenario_path: путь к файлу сценария
        api_settings: настройки API (api_key, model, temperature, max_tokens, timeout)
        progress_callback: функция обратного вызова для прогресса
        cancel_event: событие для отмены

    Returns:
        Dict с ключами: status, result, format
    """
    if isinstance(document_paths, str):
        document_paths = [document_paths]
    if not document_paths:
        return {"status": "error", "message": "No documents provided"}

    try:
        if progress_callback:
            progress_callback("log", "Загрузка сценария...")
        scenario = ScenarioManager.load_scenario(scenario_path)
        stages = scenario["stages"]

        if not stages:
            return {"status": "error", "message": "Scenario has no stages"}

        files_mode = scenario.get("files_processing", {}).get("mode", "combined")
        aggregate_stage_id = scenario.get("files_processing", {}).get("aggregate_stage")

        # Получаем содержимое всех документов
        docs_content = []
        for path in document_paths:
            if progress_callback:
                progress_callback("log", f"Извлечение содержимого {path}...")
            doc_content = get_document_content(path, {})
            docs_content.append({
                "path": path,
                "content": doc_content["content"],
                "format": doc_content["format"]
            })

        if files_mode == "combined":
            combined_text = ""
            for idx, doc in enumerate(docs_content):
                combined_text += f"--- Документ {idx + 1}: {doc['path']} ---\n{doc['content']}\n\n"

            # Выполняем пайплайн и оборачиваем результат в словарь со статусом
            pipeline_result = _run_pipeline(combined_text, stages, api_settings, progress_callback, cancel_event)
            return {
                "status": "success",
                "result": pipeline_result["content"],
                "format": pipeline_result["format"]
            }

        elif files_mode == "separate":
            per_file_results = []
            for idx, doc in enumerate(docs_content):
                if cancel_event and cancel_event.is_set():
                    return {"status": "cancelled"}
                if progress_callback:
                    progress_callback("log", f"Обработка файла {idx + 1}/{len(docs_content)}: {doc['path']}")

                result = _run_pipeline(
                    doc["content"],
                    stages,
                    api_settings,
                    progress_callback,
                    cancel_event,
                    stop_before=aggregate_stage_id
                )
                per_file_results.append({
                    "file": doc["path"],
                    "result": result
                })
                if cancel_event and cancel_event.is_set():
                    return {"status": "cancelled"}

            if aggregate_stage_id:
                # Находим этап слияния
                merge_stage = next((s for s in stages if s["id"] == aggregate_stage_id), None)
                if not merge_stage:
                    raise ValueError(f"Aggregate stage '{aggregate_stage_id}' not found in stages")

                # Собираем результаты всех файлов
                aggregated_text = ""
                for item in per_file_results:
                    aggregated_text += f"--- Файл: {item['file']} ---\n{item['result']['content']}\n\n"

                # Выполняем этап слияния (передаём aggregated_text как input_text)
                merge_result = _execute_stage(
                    merge_stage,
                    aggregated_text,
                    api_settings,
                    progress_callback,
                    cancel_event,
                    {}  # stage_results для ссылок (здесь они не нужны)
                )
                final_result = merge_result
            else:
                # Нет этапа слияния — возвращаем JSON со всеми результатами
                final_result = {
                    "content": json.dumps(per_file_results, ensure_ascii=False, indent=2),
                    "format": "json"
                }

            return {
                "status": "success",
                "result": final_result["content"],
                "format": final_result["format"]
            }

        else:
            raise ValueError(f"Unknown files_processing.mode: {files_mode}")

    except Exception as e:
        logger.exception("Ошибка в оркестраторе")
        return {"status": "error", "message": str(e)}


def _run_pipeline(input_text: str,
                  stages: list,
                  api_settings: dict,
                  progress_callback: Optional[Callable],
                  cancel_event: Optional[threading.Event],
                  stop_before: Optional[str] = None) -> Dict[str, Any]:
    """
    Выполняет последовательность этапов до указанного (или все).

    Returns:
        Dict с ключами: content, format
    """
    if not stages:
        return {"content": input_text, "format": "text"}

    stage_results = {}
    total = len(stages)
    execute_stages = []

    # Определяем, какие этапы выполнять
    for stage in stages:
        if stop_before and stage["id"] == stop_before:
            break
        execute_stages.append(stage)

    # Если нет этапов для выполнения
    if not execute_stages:
        return {"content": input_text, "format": "text"}

    # Выполняем этапы
    for idx, stage in enumerate(execute_stages, 1):
        if cancel_event and cancel_event.is_set():
            raise InterruptedError("Cancelled")

        if progress_callback:
            progress_callback("stage", {"name": stage["id"], "current": idx, "total": total})
            progress_callback("log", f"Этап {idx}/{total}: {stage['id']}")

        result = _execute_stage(
            stage,
            input_text,
            api_settings,
            progress_callback,
            cancel_event,
            stage_results
        )
        stage_results[stage["id"]] = result

    # Определяем финальный результат
    if stop_before:
        # Если мы остановились перед каким-то этапом, берём результат последнего выполненного
        last_stage = execute_stages[-1] if execute_stages else None
        if last_stage and last_stage["id"] in stage_results:
            return stage_results[last_stage["id"]]
        else:
            return {"content": input_text, "format": "text"}
    else:
        # Берём результат последнего этапа
        last_stage = stages[-1]
        if last_stage["id"] in stage_results:
            return stage_results[last_stage["id"]]
        else:
            return {"content": input_text, "format": "text"}


def _apply_template(template: str, data: Any) -> str:
    if not template:
        return str(data)
    if isinstance(data, dict):
        def replace(match):
            key = match.group(1).strip()
            parts = key.split('.')
            value = data
            for part in parts:
                if isinstance(value, dict) and part in value:
                    value = value[part]
                else:
                    return match.group(0)
            return str(value) if value is not None else ""
        import re
        return re.sub(r'\{\{\s*([^}]+)\s*\}\}', replace, template)
    else:
        return template.replace("{{content}}", str(data))

def _execute_stage(stage: Dict,
                   input_text: str,
                   api_settings: Dict,
                   progress_callback: Optional[Callable],
                   cancel_event: Optional[threading.Event],
                   stage_results: Optional[Dict] = None) -> Dict[str, Any]:
    if stage_results is None:
        stage_results = {}

    source = stage["input"]["source"]
    sources = source if isinstance(source, list) else [source]
    processing_mode = stage.get("processing_mode", "llm")

    # ---- Режим "concat" ----
    if processing_mode == "concat":
        combined_text = ""
        for s in sources:
            if s.startswith("stage."):
                ref_id = s.split(".", 1)[1]
                if ref_id in stage_results:
                    combined_text += stage_results[ref_id]["content"] + "\n\n"
            elif s == "document":
                combined_text += input_text + "\n\n"
        result = {
            "content": combined_text.strip(),
            "format": "text",
            "formatted_content": combined_text.strip()
        }
        if progress_callback:
            preview = combined_text[:500] + ("..." if len(combined_text) > 500 else "")
            progress_callback("stage_result", {
                "stage_id": stage["id"],
                "result_preview": preview,
                "full_result": combined_text,
                "formatted_content": combined_text,
                "format": "text"
            })
        return result

    # ---- Режим "template" ----
    if processing_mode == "template":
        template_content = stage.get("template_file_content", "")
        if not template_content:
            raise ValueError(f"Stage {stage['id']}: template_file_content is empty for template mode")

        import re
        def replace_stage_ref(match):
            ref = match.group(1).strip()
            if ref.startswith("stage."):
                ref_id = ref.split(".", 1)[1]
                if ref_id in stage_results:
                    return stage_results[ref_id]["content"]
            return match.group(0)

        result_text = re.sub(r'\{\{\s*(stage\.[^}]+)\s*\}\}', replace_stage_ref, template_content)

        result = {
            "content": result_text,
            "format": "text",
            "formatted_content": result_text
        }
        if progress_callback:
            preview = result_text[:500] + ("..." if len(result_text) > 500 else "")
            progress_callback("stage_result", {
                "stage_id": stage["id"],
                "result_preview": preview,
                "full_result": result_text,
                "formatted_content": result_text,
                "format": "text"
            })
        return result

    # ---- Режим "llm" (по умолчанию) ----
    # Определяем источники входных данных
    input_texts = []
    for s in sources:
        if s == "document":
            input_texts.append(input_text)
        elif s == "aggregated_results":
            input_texts.append(input_text)
        elif s.startswith("stage."):
            ref_id = s.split(".", 1)[1]
            if ref_id not in stage_results:
                raise ValueError(f"Stage {stage['id']}: reference to unknown stage '{ref_id}'")
            input_texts.append(stage_results[ref_id]["content"])
        else:
            raise ValueError(f"Unknown source: {s}")

    if len(input_texts) == 1:
        combined_input = input_texts[0]
    else:
        combined_input = "\n\n---\n\n".join(input_texts)

    # Получаем системный промпт
    system_prompt = stage.get("system_prompt", "")
    # Получаем пользовательский промпт (с подстановкой {content})
    prompt_template = stage.get("prompt_template", "")
    if not prompt_template:
        raise ValueError(f"Stage {stage['id']}: missing prompt_template")
    user_prompt = prompt_template.replace("{content}", combined_input)

    # Примеры (format_examples) уже загружены как список пар
    examples = stage.get("format_examples")

    # Параметры модели
    stage_params = stage.get("params", {})
    model = stage_params.get("model", api_settings.get("model", "gpt-4o-mini"))
    temperature = stage_params.get("temperature", api_settings.get("temperature", 0.2))
    max_tokens = stage_params.get("max_tokens", api_settings.get("max_tokens", 2000))
    timeout = stage_params.get("timeout", api_settings.get("timeout", 60))

    analyzer = Analyzer(
        api_key=api_settings["api_key"],
        model=model,
        temperature=temperature,
        max_tokens=max_tokens,
        timeout=timeout
    )

    output_format = stage["output_format"]
    try:
        response = analyzer.analyze(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            examples=examples,
            output_format=output_format,
            cancel_event=cancel_event
        )
    except InterruptedError:
        raise

    result = {
        "content": response,
        "format": output_format
    }

    # Применяем шаблон вывода
    if "output_template" in stage and stage["output_template"]:
        try:
            data = json.loads(response)
        except:
            data = response
        formatted = _apply_template(stage["output_template"], data)
        result["formatted_content"] = formatted
    else:
        result["formatted_content"] = response

    # Отправляем промежуточный результат
    if progress_callback:
        preview = response[:500] + ("..." if len(response) > 500 else "")
        progress_callback("stage_result", {
            "stage_id": stage["id"],
            "result_preview": preview,
            "full_result": response,
            "formatted_content": result.get("formatted_content", response),
            "format": output_format
        })

    return result

def _resolve_stage_reference(ref: str, stage_results: Dict) -> str:
    """Извлекает содержимое этапа по ссылке stage.<id> или возвращает None."""
    if ref.startswith("stage."):
        ref_id = ref.split(".", 1)[1]
        if ref_id in stage_results:
            return stage_results[ref_id]["content"]
    return None
