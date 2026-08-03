# orchestrator/scenario_manager.py
import json
import os
from pathlib import Path
from typing import Dict, Any

class ScenarioManager:
    @staticmethod
    def load_scenario(file_path: str) -> Dict[str, Any]:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Scenario not found: {file_path}")
        with open(file_path, 'r', encoding='utf-8') as f:
            scenario = json.load(f)
        ScenarioManager._validate(scenario)
        scenario_dir = Path(file_path).parent

        for stage in scenario.get("stages", []):
            # ---- Загрузка системного промпта (system) ----
            system_ref = stage.get("system_prompt")
            if system_ref:
                system_file = system_ref if system_ref.endswith('.md') else f"{system_ref}.md"
                system_path = scenario_dir / system_file
                if system_path.exists():
                    with open(system_path, 'r', encoding='utf-8') as f:
                        stage["system_prompt"] = f.read()
                else:
                    raise FileNotFoundError(f"System prompt file not found: {system_path}")
            else:
                stage["system_prompt"] = ""

            # ---- Загрузка пользовательского промпта (user) ----
            prompt_ref = stage.get("prompt_template")
            if prompt_ref and not prompt_ref.startswith('{') and not prompt_ref.startswith('<'):
                prompt_file = prompt_ref if prompt_ref.endswith('.md') else f"{prompt_ref}.md"
                prompt_path = scenario_dir / prompt_file
                if prompt_path.exists():
                    with open(prompt_path, 'r', encoding='utf-8') as f:
                        stage["prompt_template"] = f.read()
                else:
                    raise FileNotFoundError(f"User prompt file not found: {prompt_path}")

            # ---- Загрузка примеров (assistant) ----
            examples_ref = stage.get("format_examples")
            if examples_ref:
                examples_file = examples_ref if examples_ref.endswith('.md') else f"{examples_ref}.md"
                examples_path = scenario_dir / examples_file
                if examples_path.exists():
                    with open(examples_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    stage["format_examples"] = ScenarioManager._parse_examples(content)
                else:
                    raise FileNotFoundError(f"Format examples file not found: {examples_path}")

            # ---- Загрузка шаблона вывода (output_template) ----
            template_ref = stage.get("output_template")
            if template_ref:
                template_file = template_ref if template_ref.endswith('.md') else f"{template_ref}.md"
                template_path = scenario_dir / template_file
                if template_path.exists():
                    with open(template_path, 'r', encoding='utf-8') as f:
                        stage["output_template"] = f.read()
                else:
                    raise FileNotFoundError(f"Output template file not found: {template_path}")

            # ---- Загрузка файла шаблона для режима "template" ----
            if stage.get("processing_mode") == "template":
                template_ref = stage.get("template_file")
                if template_ref:
                    template_file = template_ref if template_ref.endswith('.md') else f"{template_ref}.md"
                    template_path = scenario_dir / template_file
                    if template_path.exists():
                        with open(template_path, 'r', encoding='utf-8') as f:
                            stage["template_file_content"] = f.read()
                    else:
                        raise FileNotFoundError(f"Template file not found: {template_path}")

        return scenario

    @staticmethod
    def _validate(scenario: Dict):
        if 'stages' not in scenario or not isinstance(scenario['stages'], list):
            raise ValueError("Missing or invalid 'stages' array")

        stage_ids = set()
        valid_output_formats = {'markdown', 'json', 'text'}

        for i, stage in enumerate(scenario['stages']):
            if 'id' not in stage or not isinstance(stage['id'], str):
                raise ValueError(f"Stage {i}: missing 'id'")
            if stage['id'] in stage_ids:
                raise ValueError(f"Duplicate stage id: {stage['id']}")
            stage_ids.add(stage['id'])

            if 'input' not in stage or not isinstance(stage['input'], dict):
                raise ValueError(f"Stage {stage['id']}: missing 'input'")
            source = stage['input'].get('source')
            if not source:
                raise ValueError(f"Stage {stage['id']}: input.source missing")
            sources = source if isinstance(source, list) else [source]
            for s in sources:
                if not isinstance(s, str):
                    raise ValueError(
                        f"Stage {stage['id']}: input.source elements must be strings (got {type(s).__name__})")
                if s.startswith('stage.'):
                    ref_id = s.split('.', 1)[1]
                    if ref_id not in stage_ids and ref_id != stage['id']:
                        raise ValueError(
                            f"Stage {stage['id']}: references unknown stage '{ref_id}'. Available: {sorted(stage_ids)}")
                elif s not in ['document', 'aggregated_results']:
                    raise ValueError(
                        f"Stage {stage['id']}: invalid input.source value '{s}'. Must be 'document', 'aggregated_results' or 'stage.<id>'")

            processing_mode = stage.get("processing_mode", "llm")
            if processing_mode not in ["llm", "concat", "template"]:
                raise ValueError(
                    f"Stage {stage['id']}: invalid processing_mode '{processing_mode}'. "
                    "Must be one of 'llm', 'concat', 'template'"
                )

            if processing_mode == "llm":
                if 'prompt_template' not in stage:
                    raise ValueError(
                        f"Stage {stage['id']}: missing prompt_template (required for processing_mode='llm')")
                if 'output_format' not in stage:
                    raise ValueError(f"Stage {stage['id']}: missing output_format (required for processing_mode='llm')")
            else:
                # Для concat и template эти поля не обязательны
                pass

            if processing_mode == "template":
                if 'template_file' not in stage:
                    raise ValueError(
                        f"Stage {stage['id']}: missing 'template_file' (required for processing_mode='template')")
                if not isinstance(stage['template_file'], str):
                    raise ValueError(f"Stage {stage['id']}: 'template_file' must be a string")

            # Проверка output_format только если оно есть и режим llm
            if 'output_format' in stage and stage['output_format'] not in valid_output_formats:
                raise ValueError(
                    f"Stage {stage['id']}: invalid output_format '{stage['output_format']}'. "
                    f"Must be one of {valid_output_formats}"
                )

            if 'params' in stage and not isinstance(stage['params'], dict):
                raise ValueError(f"Stage {stage['id']}: 'params' must be a dictionary")
            if 'system_prompt' in stage and not isinstance(stage['system_prompt'], str):
                raise ValueError(f"Stage {stage['id']}: 'system_prompt' must be a string")

            # Проверка необязательных полей (типы)
            if 'output_template' in stage and not isinstance(stage['output_template'], str):
                raise ValueError(f"Stage {stage['id']}: 'output_template' must be a string")
            if 'format_examples' in stage and not isinstance(stage['format_examples'], str):
                raise ValueError(f"Stage {stage['id']}: 'format_examples' must be a string")

        # Проверка final_output
        final = scenario.get('final_output')
        if final:
            if not isinstance(final, dict):
                raise ValueError("final_output must be a dictionary")
            if 'format' in final:
                valid_final = {'json', 'csv', 'markdown', 'text'}
                if final['format'] not in valid_final:
                    raise ValueError(f"final_output.format '{final['format']}' invalid. Must be one of {valid_final}")
            if 'source' in final:
                source = final['source']
                if source.startswith('stage.'):
                    ref_id = source.split('.', 1)[1]
                    if ref_id not in stage_ids:
                        raise ValueError(f"final_output.source references unknown stage {ref_id}")
                else:
                    raise ValueError("final_output.source must start with 'stage.'")
            else:
                raise ValueError("final_output.source is required")

        # Валидация files_processing
        fp = scenario.get('files_processing')
        if fp:
            if not isinstance(fp, dict):
                raise ValueError("files_processing must be a dictionary")
            mode = fp.get('mode')
            if mode not in ['separate', 'combined']:
                raise ValueError("files_processing.mode must be 'separate' or 'combined'")
            if mode == 'separate':
                agg_id = fp.get('aggregate_stage')
                if agg_id:
                    if agg_id not in stage_ids:
                        raise ValueError(f"Aggregate stage '{agg_id}' not found in stages")
                    agg_stage = next((s for s in scenario['stages'] if s['id'] == agg_id), None)
                    if agg_stage:
                        src = agg_stage['input'].get('source')
                        if src != 'aggregated_results' and (isinstance(src, list) and 'aggregated_results' not in src):
                            raise ValueError(
                                f"Aggregate stage '{agg_id}' must have input.source='aggregated_results' or include it in list")

    @staticmethod
    def _parse_examples(content: str) -> list[tuple[str, str]]:
        """
        Парсит файл с примерами вида:
           --- user ---
           текст
           --- assistant ---
           текст
        Возвращает список кортежей (user_text, assistant_text).
        """
        import re
        examples = []
        parts = re.split(r'\n?---\s*(user|assistant)\s*---\s*\n?', content, flags=re.IGNORECASE)
        # parts: [maybe_text, role1, text1, role2, text2, ...]
        if len(parts) < 3:
            return examples
        i = 1
        while i + 2 <= len(parts):
            role = parts[i].strip().lower()
            text = parts[i + 1].strip()
            if role == 'user':
                if i + 2 < len(parts) and parts[i + 2].strip().lower() == 'assistant':
                    assistant_text = parts[i + 3].strip()
                    examples.append((text, assistant_text))
                    i += 4
                else:
                    i += 2
            else:
                i += 2
        return examples