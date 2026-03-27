# LLM-Parser: Характеристики печатных плат

Интеллектуальный парсер для извлечения технических характеристик печатных плат из **Excel** и **Word** файлов с автоматической нормализацией значений и интеграцией с Битрикс24 (смарт-процесс, entityTypeId=182).

---

## Описание

Система использует **двухшаговый LLM-пайплайн** на базе **Mistral AI** для обработки листов технических требований ПП (на русском и английском) из форматов `.xlsx`, `.xls`, `.docx`, `.doc`:

**Шаг 1 — Извлечение (Промпт 1)**
`mistral-medium-latest` + `with_structured_output(PCBCharacteristics)` извлекает характеристики платы в структурированную Pydantic-модель.

**Шаг 2 — Нормализация (Промпт 2)**
`mistral-small-latest` приводит извлечённые текстовые значения к каноническим из справочников Битрикс24 (`"FR-4 High TG"` → `"FR4 TG-180"`, `"ENIG"` → `"Imm. gold (chem.Ni/Au)"` и т.д.), затем маппит на `item_id` для передачи в API.

---

## Как это работает

```
Файл (Excel/Word)
      │
      ▼
extract_document_data()        ← utils.py: читает все листы / параграфы / таблицы
      │
      ▼
LLM Промпт 1                   ← mistral-medium + structured output → PCBCharacteristics
      │  Извлечённые характеристики (raw dict)
      ▼
LLM Промпт 2                   ← mistral-small + JSON mode → нормализованные значения
      │  {finish_type, copper_thickness, base_material, pcb_type, ipc_class}
      ▼
map_to_bitrix24_ids()          ← pcb_normalizer.py → {ufCrm24_*: item_id}
      │
      ▼
Интерфейс Gradio               ← Таблица характеристик + таблица полей Битрикс24
      │
      ▼
crm.item.add (REST API)        ← bitrix24.py → создание элемента в смарт-процессе 182
```

---

## Извлекаемые характеристики (PCBCharacteristics)

| Поле | Описание |
|------|----------|
| `company_name` | Название компании |
| `board_name` | Обозначение платы (OEM PN) |
| `base_material` | Материал основания (FR4, Rogers, Polyimide…) |
| `board_thickness` | Толщина платы с допуском, мм |
| `foil_thickness` | Толщина меди (CU foil) |
| `layer_count` | Количество слоёв |
| `coverage_type` | Финишное покрытие (Surface Finish) |
| `board_size` | Габариты платы (Длина × Ширина) |
| `panelization` | Параметры панели |
| `solder_mask_colour` | Цвет паяльной маски |
| `solder_mark_colour` | Цвет маркировки |
| `soldering_surface` | Тип монтажа |
| `electrical_testing` | Электроконтроль |
| `edge_plating` | Металлизация торца |
| `contour_treatment` | Обработка контура |
| `pcb_type` | Тип платы (Rigid / Flex / Semi-Flex…) |
| `peelable_mask` | Пилинг-маска (Yes/No) |
| `gold_fingers` | Золотые пальцы (Yes/No) |
| `ipc_class` | Класс качества IPC (Class 2 / Class 3) |
| `back_drill` | Обратное сверление (Yes/No) |
| `flex_type` | Тип гибкой платы |
| `cover_layer` | Coverlay (Yes/No) |
| `flex_layer_location` | Расположение гибких слоёв |
| `coin` | Copper coin (Yes/No) |
| `embedded_components` | Встроенные компоненты (Yes/No) |

---

## Маппинг в поля Битрикс24 (UF_CRM_24_*)

### Обязательные поля

| Код поля | Название | Источник |
|----------|----------|----------|
| `ufCrm24_1709799376061` | OEM PN | `board_name` |
| `ufCrm24_1709799393816` | OEM Description | Генерируется |
| `ufCrm24_1709799420584` | Rev. | Пустая строка |
| `ufCrm24_1708374728464` | Board Thickness (mm) | `board_thickness` |
| `ufCrm24_1707768819` | Finish Type (iblock 74) | `coverage_type` → Промпт 2 |
| `ufCrm24_1707838248` | Materials (iblock 56) | `base_material` → Промпт 2 |
| `ufCrm24_1707838441` | Max Copper base OZ (iblock 62) | `foil_thickness` → Промпт 2 |
| `ufCrm24_1707838074` | PCB type (iblock 52) | `pcb_type` → Промпт 2 |
| `ufCrm24_1709815185` | No of Layers (iblock 54) | `layer_count` |
| `ufCrm24_1707839629` | Peelable SM (iblock 86) | `peelable_mask` |
| `ufCrm24_1707838030` | Order unit (iblock 50) | Дефолт: `ea` |
| `ufCrm24_1707849863` | Production Unit (iblock 160) | Дефолт: `ea` |

### Дополнительные поля

| Код поля | Название | Источник |
|----------|----------|----------|
| `ufCrm24_1707849930` | Outer layers Copper | `foil_thickness` → Промпт 2 |
| `ufCrm24_1707841162` | IPC Class | `ipc_class` → Промпт 2 |
| `ufCrm24_1707838528` | Gold Fingers (iblock 70) | `gold_fingers` |
| `ufCrm24_1707839110` | Edge plating (iblock 72) | `edge_plating` |
| `ufCrm24_1707851410` | Back Drill Holes (iblock 230) | `back_drill` |
| `ufCrm24_1707851442` | Coin (iblock 232) | `coin` |
| `ufCrm24_1707851467` | Embedded Components (iblock 234) | `embedded_components` |
| `ufCrm24_1707840205` | Cover layer (iblock 110) | `cover_layer` |
| `ufCrm24_1707840178` | Flex Type (iblock 108) | `flex_type` |
| `ufCrm24_1707851507` | Flex Layer Location (iblock 236) | `flex_layer_location` |
| `ufCrm24_1708353384301` | Board Length (mm) | Из `board_size` |
| `ufCrm24_1708353402068` | Board Width (mm) | Из `board_size` |
| `ufCrm24_1708374692747` | Board Size (sqr dec) | Length × Width / 100 |
| `ufCrm24_1708375852081` | Panel Length (mm) | Из `panelization` |
| `ufCrm24_1708375871512` | Panel Width (mm) | Из `panelization` |
| `ufCrm24_1708375895460` | Panel Size (sqr dec) | Panel L × W / 100 |

---

## Структура проекта

```
LLM-parser-PCB/
├── main.py                       # Точка входа FastAPI + Gradio mount
├── requirements.txt              # Зависимости Python
├── .env                          # API-ключи (не коммитится)
│
└── app/
    ├── config.py                 # Загрузка env, mistral_params, bitrix24_config
    ├── model.py                  # Pydantic PCBCharacteristics (структурированный вывод LLM)
    ├── utils.py                  # Извлечение текста из Excel/Word, вызов LLM Промпт 1
    ├── pcb_normalizer.py         # Словари Битрикс24, LLM Промпт 2, маппинг item_id
    ├── bitrix24.py               # REST API Битрикс24, MistralChatAdapter, map_pcb_to_bitrix24_fields
    ├── bitrix24_dictionaries.py  # Fallback-словари (без БД) + fuzzy-маппинг
    ├── db_dictionaries.py        # SQLAlchemy-клиент для справочников из БД (продакшен)
    ├── interface.py              # Gradio веб-интерфейс (загрузка, результаты, отправка)
    └── logger.py                 # Настройка логирования
```

---

## Требования

- Python 3.9+
- API ключ Mistral AI
- Webhook URL Битрикс24 (опционально, для отправки в CRM)
- Для `.doc` (Word 97–2003): [LibreOffice](https://www.libreoffice.org/) или Microsoft Word

---

## Установка

### 1. Клонирование

```bash
git clone https://github.com/kornelikart/LLM-parser-PCB-work-v.327.git
cd LLM-parser-PCB-work-v.327
```

### 2. Виртуальное окружение

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux/Mac
python -m venv venv
source venv/bin/activate
```

### 3. Зависимости

```bash
pip install -r requirements.txt
```

### 4. Файл `.env`

Создайте `.env` в корне проекта (рядом с `main.py`):

```env
# Обязательно
MISTRAL_API_KEY=ваш_ключ_mistral_ai

# Опционально: интеграция с Битрикс24
BITRIX24_WEBHOOK_URL=https://fineline.bitrix24.ru/rest/6/<token>/crm.item.add

# Опционально: БД справочников (для продакшена)
DICTIONARIES_DB_URL=postgresql://user:password@host:port/dbname
```

### 5. Запуск

```bash
fastapi run main.py
```

Откройте браузер: [http://localhost:8000](http://localhost:8000)

---

## Пример запроса к Битрикс24

```bash
curl --location 'https://fineline.bitrix24.ru/rest/6/<token>/crm.item.add' \
--header 'Content-Type: application/json' \
--data '{
  "entityTypeId": 182,
  "fields": {
    "ufCrm24_1707768819": 5934,
    "ufCrm24_1707838030": 5256,
    "ufCrm24_1707838074": 6610,
    "ufCrm24_1707838248": 5774,
    "ufCrm24_1707838441": 5824,
    "ufCrm24_1707839629": 6014,
    "ufCrm24_1707849863": 6270,
    "ufCrm24_1708374728464": 1.6,
    "ufCrm24_1709799376061": "OEM-12345",
    "ufCrm24_1709799393816": "4-layer PCB, FR4 TG-180, ENIG",
    "ufCrm24_1709799420584": ""
  }
}'
```

---

## Справочники (iblock_id → поле Битрикс24)

| iblock_id | Название | Поле |
|-----------|----------|------|
| 50 | Order unit | `ufCrm24_1707838030` |
| 52 | PCB type | `ufCrm24_1707838074` |
| 54 | No of Layers | `ufCrm24_1709815185` |
| 56 | Materials | `ufCrm24_1707838248` |
| 62 | Max Copper (base OZ) | `ufCrm24_1707838441` |
| 66 | Shelf Life (Months) | `ufCrm24_1707838475` |
| 68 | Hard gold – Selective | `ufCrm24_1707838860` |
| 70 | Gold Fingers | `ufCrm24_1707838528` |
| 72 | Edge plating | `ufCrm24_1707839110` |
| 74 | Finish Type | `ufCrm24_1707768819` |
| 86 | Peelable SM | `ufCrm24_1707839629` |
| 108 | Flex Type | `ufCrm24_1707840178` |
| 110 | Cover layer | `ufCrm24_1707840205` |
| 160 | Production Unit | `ufCrm24_1707849863` |
| 188 | Type of Aluminum | `ufCrm24_1707850364` |
| 230 | Back Drill Holes | `ufCrm24_1707851410` |
| 232 | Coin | `ufCrm24_1707851442` |
| 234 | Embedded Components | `ufCrm24_1707851467` |
| 236 | Flex Layer Location | `ufCrm24_1707851507` |

---

## Режимы работы справочников

Система использует трёхуровневый fallback:

1. **БД (продакшен)** — если задана `DICTIONARIES_DB_URL`, справочники берутся из БД с TTL-кэшем 15 мин. Новые значения подхватываются автоматически.
2. **LLM нормализация (основной путь)** — Промпт 2 нормализует сырые значения к каноническим из `pcb_normalizer.py`.
3. **Статические словари (fallback)** — если Промпт 2 недоступен, используются словари из `bitrix24_dictionaries.py` с fuzzy-маппингом.

### Структура таблицы БД

```sql
CREATE TABLE bitrix24_dictionaries (
    param_id  INTEGER,   -- IBLOCK_ID справочника (напр. 56 для Materials)
    param     TEXT,      -- Название справочника
    item_id   INTEGER,   -- ID элемента (передаётся в API)
    item      TEXT       -- Текстовое значение (напр. "FR4 TG-180")
);
```
