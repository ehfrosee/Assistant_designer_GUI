# -*- coding: utf-8 -*-
"""Настройки по умолчанию для config.json"""

def default_config() -> dict:
    return {
        "llm_defaults": {
            "model": "gpt-4o-mini",
            "temperature": 0.3,
            "max_tokens": 2000,
            "timeout": 60
        },
        "dialog": {
            "history_pairs": 3,
            "summary_pairs": 3,
            "summary_model": "gpt-4o-mini"
        },
        "paths": {
            "scenarios_dir": "./scenarios",
            "rag_db_dir": "./rag_db",
            "logs_dir": "./logs",
            "env_file": ".env",
            "dialogs_dir": "./dialogs",
            "converted_files_dir": "./converted_files"   # <-- новая строка
        },
        "rag": {
            "embedding_model": "text-embedding-ada-002",
            "top_k": 5,
            "vector_db_type": "faiss"
        },
        "ui": {
            "theme": "light",
            "language": "ru"
        },
        "processing": {
            "enable_chunking": True,
            "chunk_size_tokens": 60000,  # максимальный размер чанка в токенах
            "chunk_overlap_tokens": 200,  # перекрытие между чанками (в токенах)
            "chunk_overlap_percent": 10  # альтернативно в процентах от размера чанка
        },
        "prompt_defaults": {
            "auto_load_file": ""  # путь к файлу для автозагрузки промпта
        }
    }