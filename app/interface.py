import json
import os
import re

import gradio as gr
import pandas as pd

try:    # for running interface.py
    import utils
    from config import mistral_params, bitrix24_config
    from logger import setup_logger
    from model import PCBCharacteristics
    import bitrix24
except ImportError:  # for running main.py
    from . import utils
    from .config import mistral_params, bitrix24_config
    from .logger import setup_logger
    from .model import PCBCharacteristics
    from . import bitrix24

logger = setup_logger()

_EDIT_HINT = (
    "✏️ Характеристики изменены. Нажмите «Пересчитать поля Битрикс24», "
    "чтобы обновить поля перед отправкой."
)

# Типы файлов для загрузки. Браузерная часть Gradio 5 сверяет расширение со списком
# с учётом регистра, и «Бланк заказа.XLSX» отклонялся с «Invalid file type»;
# поэтому заглавные варианты перечислены явно. Сервер Gradio и extract_document_data
# регистр не учитывают.
_UPLOAD_EXTENSIONS = [".xlsx", ".xls", ".docx", ".doc", ".txt"]
UPLOAD_FILE_TYPES = _UPLOAD_EXTENSIONS + [ext.upper() for ext in _UPLOAD_EXTENSIONS]


def _file_basename(file):
    """Имя файла без расширения для сохранения результатов."""
    path = getattr(file, "name", None) or str(file)
    return path.rsplit(".", 1)[0] if "." in path else path


def _friendly_error_message(e) -> str:
    """Переводит технические ошибки (файл, Mistral, сеть) в сообщение для пользователя.

    Тип ошибки определяется по классу исключения и HTTP-статусу, а не по подстроке:
    в тексте ошибки бывают данные документа («ГОСТ Р 53429-2009» — это не HTTP 429).
    """
    if isinstance(e, (utils.MistralRateLimitError, utils.MistralNetworkError, utils.DocumentReadError)):
        return str(e)  # текст уже сформулирован для пользователя — без второй обёртки
    msg = str(e)
    status = utils.http_status_of(e)
    if status == 401 or (status is None and "unauthorized" in msg.lower()):
        return (
            "Mistral API вернул 401 Unauthorized. Проверьте, что переменная окружения "
            "`MISTRAL_API_KEY` задана и ключ действителен."
        )
    if utils.is_tier_error(e):
        return (
            "Модель недоступна на вашем тарифе Mistral (403 tier_not_allowed). "
            f"{utils._TIER_HINT}"
        )
    if utils.is_rate_limit_error(e):
        return "Превышен лимит запросов Mistral API (HTTP 429). Попробуйте позже."
    return f"Ошибка при обработке файла: {msg}"


def _table_to_pcb_dict(table_df: pd.DataFrame) -> dict:
    """Собирает словарь характеристик из (возможно отредактированной) таблицы Gradio."""
    valid_keys = set(PCBCharacteristics.model_fields.keys())
    data = {}
    for _, row in table_df.iterrows():
        if len(row) < 2:
            continue
        key = str(row.iloc[0]).strip()
        if key not in valid_keys:
            continue
        raw_val = row.iloc[1]
        if raw_val is None:
            val = ""
        else:
            try:
                val = "" if pd.isna(raw_val) else str(raw_val).strip()
            except (TypeError, ValueError):
                val = str(raw_val).strip()
        data[key] = val

    # layer_count в модели целочисленный: берём первое число из текста ("8 слоёв" -> 8)
    if "layer_count" in data:
        numbers = re.findall(r"\d+", str(data["layer_count"]))
        data["layer_count"] = int(numbers[0]) if numbers else 0

    try:
        return PCBCharacteristics(**data).model_dump()
    except Exception as e:
        logger.warning(
            "Отредактированные данные не прошли валидацию модели, используются как есть: %s", e
        )
        return data


def _map_fields_safely(pcb_data: dict, mistral_client):
    """Маппинг в поля Битрикс24. Возвращает (fields, error, note).

    error — маппинг не удался (результаты всё равно показываются);
    note — маппинг удался, но без LLM-нормализации (Промпт 2 не выполнился,
    сработал fallback на статические справочники) — пользователь должен это видеть.
    """
    diagnostics: dict = {}
    try:
        fields = bitrix24.map_pcb_to_bitrix24_fields(
            pcb_data, mistral_client=mistral_client, diagnostics=diagnostics
        )
    except Exception as e:
        logger.error("Не удалось сформировать поля Битрикс24: %s", e)
        return None, str(e), None
    note = None
    if diagnostics.get("normalization_error"):
        note = (
            "⚠️ LLM-нормализация (Промпт 2) не выполнена, использованы статические справочники — "
            "проверьте Materials / Finish Type / Copper. Причина: "
            f"{diagnostics['normalization_error']}"
        )
    return fields, None, note


def _write_result_files(base: str, parsed_dict: dict, b24_fields):
    """Сохраняет CSV/XLSX/JSON характеристик и JSON-заявку Битрикс24 (если поля сформированы)."""
    df = pd.DataFrame(list(parsed_dict.items()), columns=["Characteristic", "Value"])
    csv_path = f"{base}_pcb_parsed.csv"
    xlsx_path = f"{base}_pcb_parsed.xlsx"
    json_path = f"{base}_pcb_parsed.json"
    df.to_csv(csv_path, index=False)
    df.to_excel(xlsx_path, index=False)
    df.to_json(json_path, orient="records", force_ascii=False)

    b24_json_path = None
    if b24_fields:
        b24_json_path = f"{base}_bitrix24.json"
        payload = {"entityTypeId": bitrix24.ENTITY_TYPE_ID, "fields": b24_fields}
        with open(b24_json_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        logger.info("Создан JSON файл для Битрикс24: %s", b24_json_path)
    return df, csv_path, xlsx_path, json_path, b24_json_path


def _mapping_updates(b24_fields, map_error, b24_json_path, success_msg):
    """
    Обновления компонентов, зависящих от результата маппинга, в порядке:
    (bitrix24_mapped_fields, excel_download_bitrix24_json, b24_fields_state,
     bitrix24_status, bitrix24_send_btn).
    """
    if b24_fields:
        b24_df = pd.DataFrame([{"Field": k, "Value": v} for k, v in b24_fields.items()])
        if b24_json_path:
            json_upd = gr.update(value=b24_json_path, visible=True)
        else:
            json_upd = gr.update(visible=False)
        return (
            gr.update(value=b24_df, visible=True),
            json_upd,
            b24_fields,
            gr.update(value=success_msg, visible=True),
            gr.update(visible=True, interactive=True),
        )

    warn_msg = (
        f"⚠️ Характеристики распознаны, но поля Битрикс24 не сформированы: {map_error}\n"
        "Исправьте значения в таблице характеристик и нажмите «Пересчитать поля Битрикс24»."
    )
    return (
        gr.update(value=pd.DataFrame(), visible=False),
        gr.update(visible=False),
        None,
        gr.update(value=warn_msg, visible=True),
        gr.update(visible=True, interactive=False),
    )


# Поля, которые сами по себе не делают документ спецификацией платы: заказчика,
# обозначение и количество LLM находит и в письмах, счетах, сопроводительных записках.
_NON_TECHNICAL_FIELDS = {"company_name", "board_name", "quantity"}
# Значения, которыми LLM заполняет поле, когда данных нет.
_EMPTY_VALUES = {"", "0", "none", "null", "n/a", "-", "—", "no", "нет"}


def _has_pcb_characteristics(parsed: dict) -> bool:
    """Нашёл ли LLM в документе хоть одну техническую характеристику платы."""
    return any(
        str(value if value is not None else "").strip().casefold() not in _EMPTY_VALUES
        for key, value in parsed.items()
        if key not in _NON_TECHNICAL_FIELDS
    )


def parse_excel_pcb(file):
    """
    Извлекает данные из загруженного файла (Excel или Word), распознаёт характеристики
    через LLM и формирует поля Битрикс24. Ошибка маппинга не скрывает распознанное:
    таблица характеристик показывается всегда, а статус подсказывает, что исправить.
    Файл без текста или нечитаемый отклоняется до запроса к LLM; документ, в котором
    LLM не нашёл характеристик платы (письмо, счёт), — понятной ошибкой, а не пустой таблицей.
    """
    if isinstance(file, list) and file:
        file = file[0]
    if file is None:
        raise gr.Error("Сначала загрузите файл (Excel или Word).")

    logger.info("Starting to parse file for PCB characteristics: %s", getattr(file, "name", file))

    # Шаг 1: извлечение текста и распознавание характеристик (LLM промпт 1).
    try:
        doc_txt = utils.extract_document_data(file)
        logger.debug("Extracted document data")

        llm = utils.create_pcb_model(mistral_params)
        logger.debug("PCB model created successfully.")

        parsed_dict = utils.process_excel_pcb_with_retry(doc_txt, llm, model_params=mistral_params)
        logger.debug("Parsed PCB dictionary: %s", parsed_dict)
    except Exception as e:
        logger.error("An error occurred while parsing the file: %s", e)
        raise gr.Error(_friendly_error_message(e))

    # Документ прочитан, но это не спецификация: пустая таблица с «Характеристики
    # распознаны» только запутала бы пользователя.
    if not _has_pcb_characteristics(parsed_dict):
        name = os.path.basename(getattr(file, "name", None) or str(file))
        found = {k: v for k, v in parsed_dict.items() if v not in ("", 0, None)}
        logger.warning("В документе %s не найдено характеристик ПП (распознано: %s)", name, found)
        raise gr.Error(
            f"В документе «{name}» не найдено характеристик печатной платы. Проверьте, что "
            "загружен лист технических требований ПП или бланк заказа."
        )

    # Шаг 2: маппинг в поля Битрикс24 (LLM промпт 2 + справочники).
    b24_fields, map_error, map_note = _map_fields_safely(parsed_dict, llm)

    # Шаг 3: файлы выгрузки (характеристики — всегда, заявка Битрикс24 — если поля готовы).
    df, csv_path, xlsx_path, json_path, b24_json_path = _write_result_files(
        _file_basename(file), parsed_dict, b24_fields
    )

    success_msg = (
        "✅ Поля Битрикс24 сформированы. Проверьте данные: таблицу характеристик можно "
        "отредактировать и нажать «Пересчитать поля Битрикс24», затем отправить в CRM."
    )
    if map_note:
        success_msg = f"{success_msg}\n{map_note}"
    mapped_upd, b24_json_upd, state_val, status_upd, send_upd = _mapping_updates(
        b24_fields, map_error, b24_json_path, success_msg
    )

    return (
        gr.update(value=df, visible=True),         # excel_parsed_reports
        mapped_upd,                                # bitrix24_mapped_fields
        gr.update(value=csv_path, visible=True),   # excel_download_csv
        gr.update(value=xlsx_path, visible=True),  # excel_download_xlsx
        gr.update(value=json_path, visible=True),  # excel_download_json
        b24_json_upd,                              # excel_download_bitrix24_json
        state_val,                                 # b24_fields_state
        status_upd,                                # bitrix24_status
        send_upd,                                  # bitrix24_send_btn
        gr.update(visible=True),                   # recompute_btn
    )


def recompute_bitrix24_fields(edited_table, file):
    """
    Пересобирает поля Битрикс24 из отредактированной таблицы характеристик:
    повторная LLM-нормализация (промпт 2) + маппинг, обновление файлов выгрузки.
    """
    if edited_table is None or len(edited_table) == 0:
        raise gr.Error("Нет данных для пересчёта: сначала распознайте файл.")

    if isinstance(file, list) and file:
        file = file[0]

    pcb_data = _table_to_pcb_dict(edited_table)
    logger.info(
        "Пересчёт полей Битрикс24 по отредактированной таблице (%d характеристик)", len(pcb_data)
    )

    llm = None
    try:
        llm = utils.create_pcb_model(mistral_params)
    except Exception as e:
        # Без LLM-нормализации map_pcb_to_bitrix24_fields использует статические справочники.
        logger.warning("LLM недоступен для нормализации, используется fallback на справочники: %s", e)

    b24_fields, map_error, map_note = _map_fields_safely(pcb_data, llm)

    if file is not None:
        _, csv_path, xlsx_path, json_path, b24_json_path = _write_result_files(
            _file_basename(file), pcb_data, b24_fields
        )
        csv_upd = gr.update(value=csv_path, visible=True)
        xlsx_upd = gr.update(value=xlsx_path, visible=True)
        json_upd = gr.update(value=json_path, visible=True)
    else:
        b24_json_path = None
        csv_upd, xlsx_upd, json_upd = gr.update(), gr.update(), gr.update()

    success_msg = "✅ Поля Битрикс24 пересчитаны по отредактированным данным. Можно отправлять."
    if map_note:
        success_msg = f"{success_msg}\n{map_note}"
    mapped_upd, b24_json_upd, state_val, status_upd, send_upd = _mapping_updates(
        b24_fields, map_error, b24_json_path, success_msg
    )

    return (
        mapped_upd,    # bitrix24_mapped_fields
        csv_upd,       # excel_download_csv
        xlsx_upd,      # excel_download_xlsx
        json_upd,      # excel_download_json
        b24_json_upd,  # excel_download_bitrix24_json
        state_val,     # b24_fields_state
        status_upd,    # bitrix24_status
        send_upd,      # bitrix24_send_btn
    )


def on_table_edited():
    """Пользователь изменил таблицу характеристик: поля Битрикс24 устарели до пересчёта."""
    return (
        gr.update(value=_EDIT_HINT, visible=True),  # bitrix24_status
        gr.update(interactive=False),               # bitrix24_send_btn
    )


def send_to_bitrix24(b24_fields):
    """Отправляет подготовленные поля Битрикс24 (из состояния сессии) без повторного запуска LLM."""
    if not b24_fields:
        return (
            "Ошибка: нет подготовленных полей Битрикс24. Загрузите и распознайте файл, "
            "при необходимости нажмите «Пересчитать поля Битрикс24»."
        )

    webhook_url = bitrix24_config.get("webhook_url", "").strip()
    token       = bitrix24_config.get("token", "").strip()
    webhook_url_or_token = webhook_url if webhook_url else token

    if not webhook_url_or_token:
        return (
            "Ошибка: Webhook URL или токен Битрикс24 не задан.\n"
            "Установите переменную окружения BITRIX24_WEBHOOK_URL или BITRIX24_TOKEN.\n"
            "Формат webhook URL: https://fineline.bitrix24.ru/rest/6/<token>/crm.item.add"
        )

    try:
        logger.info("Отправка данных в Битрикс24...")
        result = bitrix24.create_bitrix24_item(webhook_url_or_token, b24_fields)

        item_id = result.get("result", {}).get("item", {}).get("id")
        if item_id:
            return f"✅ Успешно отправлено в Битрикс24! ID элемента: {item_id}"
        return f"⚠️ Данные отправлены, но ID не получен. Ответ: {result}"

    except Exception as e:
        logger.error("Ошибка при отправке в Битрикс24: %s", e)
        error_msg = str(e)
        if "401" in error_msg or "unauthorized" in error_msg.lower():
            return "❌ Ошибка авторизации в Битрикс24. Проверьте webhook URL или токен."
        return f"❌ Ошибка при отправке в Битрикс24: {error_msg}"


def _reset_results_updates():
    """Обновления, скрывающие все результаты и сбрасывающие состояние сессии."""
    return [
        gr.update(value=pd.DataFrame(), visible=False),  # excel_parsed_reports
        gr.update(value=pd.DataFrame(), visible=False),  # bitrix24_mapped_fields
        gr.update(value=None, visible=False),            # excel_download_csv
        gr.update(value=None, visible=False),            # excel_download_xlsx
        gr.update(value=None, visible=False),            # excel_download_json
        gr.update(value=None, visible=False),            # excel_download_bitrix24_json
        None,                                            # b24_fields_state
        gr.update(value="", visible=False),              # bitrix24_status
        gr.update(visible=False, interactive=False),     # bitrix24_send_btn
        gr.update(visible=False),                        # recompute_btn
    ]


def reset_before_parse():
    """Сбрасывает предыдущие результаты перед новым распознаванием (чтобы не отправить старые данные)."""
    return tuple(_reset_results_updates())


def hide_outputs():
    logger.debug("File was closed")
    return tuple(_reset_results_updates()) + (gr.update(visible=False),)  # + excel_process_btn


def create_interface(title: str = "gradio app"):
    interface = gr.Blocks(title=title)
    with interface:
        # Поля Битрикс24 текущей сессии (у каждого пользователя браузера — свои).
        b24_fields_state = gr.State(None)

        # Заголовок и краткое описание
        gr.Markdown("## LLM-Parser: Характеристики печатных плат")
        gr.Markdown(
            "Инструмент для распознавания технических требований ПП из файлов Excel / Word "
            "и формирования структурированных данных и заявки в Битрикс24."
        )

        with gr.Row():
            # Левая колонка: загрузка и запуск парсинга
            with gr.Column(scale=1):
                gr.Markdown("### Шаг 1. Загрузка файла")
                gr.Markdown(
                    "- **Поддерживаемые форматы**: `.xlsx`, `.xls`, `.docx`, `.doc`, `.txt`\n"
                    "- Подойдут лист технических требований ПП или бланк заказа "
                    "на русском и/или английском языке."
                )
                excel_input = gr.File(
                    label="Загрузить файл спецификации (Excel, Word или txt)",
                    file_types=UPLOAD_FILE_TYPES,
                    height=140
                )
                excel_process_btn = gr.Button(
                    value="Распознать характеристики",
                    visible=False,
                    variant="primary"
                )
                gr.Markdown(
                    "_Обработка может занять некоторое время из-за обращения к внешнему AI-сервису Mistral._"
                )

                # Интеграция с Битрикс24 (кнопка и статус)
                gr.Markdown("---")
                gr.Markdown("### Шаг 3. Отправка в Битрикс24")
                bitrix24_status = gr.Textbox(
                    label="Статус Битрикс24",
                    visible=False,
                    interactive=False,
                    lines=3,
                    max_lines=6,
                )
                bitrix24_send_btn = gr.Button(
                    value="Отправить распознанные данные в Битрикс24",
                    visible=False,
                    variant="secondary"
                )

            # Правая колонка: результаты и выгрузки
            with gr.Column(scale=2):
                gr.Markdown("### Шаг 2. Результаты распознавания")
                gr.Markdown(
                    "_Таблицу характеристик можно редактировать. После правок нажмите "
                    "«Пересчитать поля Битрикс24» — поля и файлы выгрузки обновятся._"
                )
                excel_parsed_reports = gr.DataFrame(
                    label="Распознанные характеристики печатной платы (редактируемые)",
                    show_copy_button=True,
                    visible=False,
                    min_width=10,
                    interactive=True,
                    col_count=(2, "fixed"),
                )
                recompute_btn = gr.Button(
                    value="Пересчитать поля Битрикс24",
                    visible=False,
                    variant="secondary",
                )
                bitrix24_mapped_fields = gr.DataFrame(
                    label="Поля для Битрикс24 (UF_CRM_24_* и значения/ID)",
                    show_copy_button=True,
                    visible=False,
                    min_width=10,
                    interactive=False,
                )
                with gr.Row():
                    excel_download_csv = gr.File(label="Скачать как CSV", visible=False)
                    excel_download_xlsx = gr.File(label="Скачать как XLSX", visible=False)
                    excel_download_json = gr.File(label="Скачать как JSON", visible=False)
                excel_download_bitrix24_json = gr.File(
                    label="Скачать JSON для Битрикс24",
                    visible=False,
                    file_types=[".json"]
                )

        # Компоненты результатов: порядок совпадает с reset_before_parse()/parse_excel_pcb().
        result_outputs = [
            excel_parsed_reports,
            bitrix24_mapped_fields,
            excel_download_csv,
            excel_download_xlsx,
            excel_download_json,
            excel_download_bitrix24_json,
            b24_fields_state,
            bitrix24_status,
            bitrix24_send_btn,
            recompute_btn,
        ]

        # Excel processing events
        excel_input.upload(lambda: gr.update(visible=True), None, excel_process_btn)
        excel_process_btn.click(
            # Сначала прячем старые результаты и сбрасываем состояние,
            # чтобы при ошибке распознавания нельзя было отправить прежние данные.
            reset_before_parse,
            None,
            result_outputs,
        ).then(
            parse_excel_pcb,
            [excel_input],
            result_outputs,
            queue=True,
        )
        excel_input.clear(
            hide_outputs,
            None,
            result_outputs + [excel_process_btn],
        )

        # Правка таблицы характеристик: поля Битрикс24 устаревают до пересчёта.
        excel_parsed_reports.input(
            on_table_edited,
            None,
            [bitrix24_status, bitrix24_send_btn],
        )
        recompute_btn.click(
            recompute_bitrix24_fields,
            [excel_parsed_reports, excel_input],
            [
                bitrix24_mapped_fields,
                excel_download_csv,
                excel_download_xlsx,
                excel_download_json,
                excel_download_bitrix24_json,
                b24_fields_state,
                bitrix24_status,
                bitrix24_send_btn,
            ],
            queue=True,
        )

        # Битрикс24 events
        bitrix24_send_btn.click(
            send_to_bitrix24,
            [b24_fields_state],
            bitrix24_status,
            queue=True
        )

    return interface

if __name__ == "__main__":
    create_interface().launch()
