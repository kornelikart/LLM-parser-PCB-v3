# LLM-Parser: Характеристики печатных плат

Интеллектуальный парсер для извлечения технических характеристик печатных плат из **Excel** и **Word** файлов с автоматической нормализацией значений и интеграцией с Битрикс24 (смарт-процесс, entityTypeId=182).

---

## Описание

Система использует **двухшаговый LLM-пайплайн** на базе **Mistral AI** для обработки листов технических требований ПП и бланков заказа (на русском и английском) из форматов `.xlsx`, `.xls`, `.docx`, `.doc`, `.txt`:

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
extract_document_data()        ← utils.py: Excel — только видимые листы/строки/колонки,
      │                          без листов-примеров и источников выпадающих списков;
      │                          Word — параграфы и таблицы; txt — как есть.
      │                          Рендер строк: «Метка | значение», лимит MAX_EXTRACT_CHARS
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
Интерфейс Gradio               ← Редактируемая таблица характеристик + таблица полей Битрикс24
      │                          (правки оператора → «Пересчитать поля Битрикс24» → повторная
      │                           нормализация и обновление полей/файлов выгрузки)
      ▼
crm.item.add (REST API)        ← bitrix24.py → создание элемента в смарт-процессе 182
```

Особенности интерфейса:

- **Таблица характеристик редактируется** до отправки: оператор может исправить ошибку
  распознавания и нажать **«Пересчитать поля Битрикс24»** — нормализация (Промпт 2) и маппинг
  выполняются заново по исправленным данным. После правки таблицы кнопка отправки блокируется
  до пересчёта, чтобы в CRM не ушли устаревшие поля.
- **Ошибка маппинга не скрывает результаты**: если поля Битрикс24 сформировать не удалось
  (например, в документе нет OEM PN или толщины платы), распознанные характеристики всё равно
  отображаются, а статус подсказывает, что исправить в таблице перед пересчётом.

---

## Извлекаемые характеристики (PCBCharacteristics)

| Поле | Описание |
|------|----------|
| `company_name` | Название компании |
| `board_name` | Обозначение платы (OEM PN) |
| `quantity` | Количество плат в заказе |
| `panel_size` | Габариты панели (Длина × Ширина) |
| `boards_per_panel` | Количество ПП в панели |
| `different_boards_per_panel` | Количество типов плат на панели |
| `technological_fields` | Наличие технологических полей (Yes/No) |
| `impedance_control` | Контроль импеданса (Yes/No) |
| `min_hole_size` | Мин. диаметр металлизированного отверстия, мм |
| `marking_side` | Сторона маркировки (TOP / BOTTOM / TOP+BOTTOM) |
| `serial_number` | Требуется серийный номер (Yes/No) |
| `date_code` | Требуется маркировка даты изготовления (Yes/No) |
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
| `ufCrm24_1708375852081` | Panel Length (mm) | Из `panel_size` (fallback: `panelization`) |
| `ufCrm24_1708375871512` | Panel Width (mm) | Из `panel_size` (fallback: `panelization`) |
| `ufCrm24_1708375895460` | Panel Size (sqr dec) | Panel L × W / 100 |
| `ufCrm24_1708375915545` | Board Per Panel | `boards_per_panel` |
| `ufCrm24_1708375925847` | Panel Usage (%) | Board Size × кол-во / Panel Size |

### Единица заказа: ea или pnl

`Order unit` и `Production Unit` автоматически переключаются с `ea` на **`pnl`**,
если плата поставляется в панелях. Признаки (достаточно одного):

- в панели больше одной платы (`boards_per_panel` > 1);
- есть технологические поля (`technological_fields` = Yes);
- задан размер панели (`panel_size`);
- в тексте панелизации сказано о поставке в панели/заготовке.

Формулировка «Панелизация: Нет» панельной поставкой **не** считается.

### Дополнительные характеристики

Этих полей нет в документации «Поля PCB». Коды трёх из них восстановлены по
выгрузкам справочников, прописаны в `pcb_normalizer.py` и заполняются
автоматически:

| Код поля | Название | Значения (item_id) |
|---|---|---|
| `ufCrm24_1707840096` | Controlled Impedance | No=6088, Yes=7204 |
| `ufCrm24_1707851328` | Serial Number | No=6516, Yes=6518 |
| `ufCrm24_1707841090` | Date code | on Top=6194, on Bottom=6196, Silkscreen-Top=6206, … |

**Date code** хранит сторону и способ нанесения маркировки. Сторона берётся из
`marking_side`, способ — из упоминания в тексте: «шелкографией» →
`Silkscreen-Top`, «паяльной маской» → `Soldermask-Top`, «медью» → `Copper-Top`;
если способ не указан — общее `on Top` / `on Bottom`.

Коды двух оставшихся полей неизвестны и задаются переменными окружения; при
заданном `BITRIX24_WEBHOOK_URL` они дополнительно ищутся на портале по названию:

| Характеристика | Значение | Переменная `.env` |
|---|---|---|
| No. Of Diff Boards | число (по умолчанию 1) | `B24_FIELD_NO_OF_DIFF_BOARDS` |
| Min. Hole size (mm) | число — само значение, не `item_id` | `B24_FIELD_MIN_HOLE_SIZE` |

Пока код не задан и не найден, характеристика распознаётся и видна в таблице
результатов и выгрузках, но в CRM не отправляется — в лог пишется предупреждение.

Автопоиск срабатывает, только когда кандидат **единственный**. Если под ключевые
слова подошло несколько полей, они пишутся в лог, а значение не отправляется —
чтобы не записать его в чужое поле. В этом случае задайте код явно в `.env`.

Если поле окажется списочным, текстовое значение автоматически переводится
в `item_id` по справочнику этого поля. Когда сопоставить значение не удалось,
поле **пропускается**: строка в списочном поле заставила бы Битрикс24
отклонить весь запрос `crm.item.add`.

`No. Of Diff Boards` заполняется вместе с `Board Per Panel`.

Посмотреть, что нашлось на портале:

```bash
python tools/fetch_bitrix24_fields.py
```

Флаг `--all` печатает все поля; полный ответ сохраняется в `tools/bitrix24_fields.json`.

### Допуски в размерах

Габариты разбираются с отбрасыванием допусков: квалитеты ISO (`h12`, `H7`, `js13`),
числовые (`±0.2`, `+0,3/-0,3`) и процентные (`±10%`). Строка бланка
`длина, мм | 114 | Допуск | h12` даёт длину **114**, а не 12.

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
    ├── bitrix24_api.py           # Живые справочники и метаданные полей с портала (TTL-кэш)
    ├── bitrix24_dictionaries.py  # Fallback-словари (без БД) + fuzzy-маппинг
    ├── db_dictionaries.py        # SQLAlchemy-клиент для справочников из БД (продакшен)
    ├── interface.py              # Gradio веб-интерфейс (загрузка, результаты, отправка)
    └── logger.py                 # Настройка логирования
```

---

## Eval-набор (контроль качества промптов)

В `eval/` лежит инфраструктура регрессионных проверок на реальных спецификациях
(сами файлы — в гитигнорируемой `eval/data/`):

```bash
python eval/run_extraction_eval.py   # извлечение текста: без LLM и API-ключа
python eval/run_llm_eval.py          # полный пайплайн: нужен MISTRAL_API_KEY
```

Подробности и формат эталонов — в [eval/README.md](eval/README.md).
Прогоняйте оба скрипта после любых правок экстракторов (`utils.py`) или промптов
(`utils.py`, `pcb_normalizer.py`).

Размер текста, передаваемого в LLM, ограничен переменной окружения
`MAX_EXTRACT_CHARS` (по умолчанию 60000 символов).

---

## Работа через VPN / proxy

`api.mistral.ai` во многих сетях доступен только через VPN, а Битрикс24 —
российский сервис, к которому обычно нужно ходить в обход VPN. Приложение
разводит эти два случая автоматически:

| Направление | Режим по умолчанию | Поведение при сбое |
|---|---|---|
| Mistral API | через системный proxy (`HTTP_PROXY` / `HTTPS_PROXY`) | проверяет прямое подключение и переключается на рабочий вариант |
| Битрикс24 | напрямую, мимо proxy | повтор запроса через proxy |

Выбранный режим для Mistral определяется один раз при первом запросе (проба ~1–2 сек)
и кэшируется. При сетевой ошибке кэш сбрасывается, и режим выбирается заново —
поэтому **включение или выключение VPN не требует перезапуска сервера**.

Принудительно задать режим (`1` — через proxy, `0` — напрямую):

```env
MISTRAL_TRUST_ENV=1
BITRIX24_TRUST_ENV=0
```

> Важно: оба LLM-шага обязаны использовать один и тот же сетевой режим.
> Клиент создаётся единственной функцией `utils.create_mistral_http_client()` —
> не создавайте `httpx.Client` для Mistral в других местах.

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

Источники значений справочников, по приоритету:

1. **Битрикс24 REST API (рекомендуется)** — если задан `BITRIX24_WEBHOOK_URL`,
   справочники читаются прямо с портала (`crm.item.fields` + `lists.element.get`)
   с TTL-кэшем 15 минут. Значение, добавленное в Битрикс24, сразу попадает
   и в промпт нормализации, и в маппинг — править код не нужно.
   Отключается: `USE_BITRIX24_API_DICTIONARIES=0`.
2. **БД** — если задана `DICTIONARIES_DB_URL`, справочники берутся из таблицы
   `bitrix24_dictionaries` с TTL-кэшем.
3. **Статические словари** — снимок в `pcb_normalizer.py` и
   `bitrix24_dictionaries.py`. Работают всегда, даже без сети.

Живые справочники **дополняют** статические, а не заменяют: значения с сервера
имеют приоритет, но локальные синонимы (`ENIG`, `Хим. Н.5 Зл.0.1`, `0,018` и др.),
важные для fuzzy-маппинга, сохраняются.

Поверх любого источника работает **LLM-нормализация** (Промпт 2): она приводит
сырой текст из документа к каноническому значению из справочника.

### Структура таблицы БД

```sql
CREATE TABLE bitrix24_dictionaries (
    param_id  INTEGER,   -- IBLOCK_ID справочника (напр. 56 для Materials)
    param     TEXT,      -- Название справочника
    item_id   INTEGER,   -- ID элемента (передаётся в API)
    item      TEXT       -- Текстовое значение (напр. "FR4 TG-180")
);
```
