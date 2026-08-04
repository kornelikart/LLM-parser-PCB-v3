# Eval-набор для пайплайна распознавания ПП

Набор реальных спецификаций + эталонные ожидания для проверки качества
извлечения текста и LLM-промптов. Сами файлы лежат в `eval/data/` и
**не коммитятся** (реальные данные заказчиков) — см. `.gitignore`.

## Структура

```
eval/
├── run_extraction_eval.py   # оффлайн-проверка извлечения (без API-ключа)
├── run_llm_eval.py          # полный прогон LLM-пайплайна (нужен MISTRAL_API_KEY)
├── data/                    # (gitignored) спецификации + <имя>.expected.json
└── results/                 # (gitignored) извлечённые тексты и ответы LLM
```

## Запуск

```bash
# 1. Извлечение (быстро, без сети): ключевые значения попали в текст, шум - нет
python eval/run_extraction_eval.py -v

# 2. Полный пайплайн (Промпт 1 + Промпт 2 + маппинг), ~2 LLM-вызова на файл
python eval/run_llm_eval.py
python eval/run_llm_eval.py --only NICEVT      # только один файл
```

`run_extraction_eval.py` возвращает код 0, если все проверки пройдены —
удобно для CI. Извлечённые тексты сохраняются в `eval/results/extraction/`,
их полезно просматривать глазами после правок экстрактора.

## Формат `<имя файла>.expected.json`

```json
{
  "description": "что проверяет кейс",
  "extract": {
    "must_contain":     ["строки, обязанные попасть в текст для LLM"],
    "must_not_contain": ["шум: скрытые колонки, листы-примеры, списки меню"]
  },
  "fields": {                        // проверка Промпта 1 (structured output)
    "board_name":  {"contains": "САНТ.758725.191"},
    "layer_count": {"equals": 2},
    "coverage_type": {"contains_any": ["ПОС", "HASL"]},
    "ipc_class":   {"contains_all": ["IPC", "3"]},
    "quantity":    {"not_empty": true, "soft": true}   // soft = информационная
  },
  "normalized": {                    // проверка Промпта 2 (канонические значения)
    "finish_type": "HASL (PbSn)",
    "base_material": null,           // null = ожидаем «не найдено»
    "copper_thickness": {"value": "0.5 OZ (17 um)", "soft": true}
  },
  "bitrix_required_ok": true         // map_pcb_to_bitrix24_fields не должен упасть
}
```

## Как добавить новый кейс

1. Положите файл спецификации в `eval/data/` (форматы: xlsx, xls, docx, doc, txt).
2. Создайте рядом `<имя>.expected.json` — начните с `extract.must_contain`
   по 3–5 ключевым значениям из документа.
3. Прогоните `run_extraction_eval.py`; посмотрите извлечённый текст в
   `eval/results/extraction/` и добавьте `must_not_contain` для замеченного шума.
4. Прогоните `run_llm_eval.py --only <имя>` и зафиксируйте `fields`/`normalized`.
   Неоднозначные ожидания помечайте `"soft": true`, чтобы они не валили прогон.
