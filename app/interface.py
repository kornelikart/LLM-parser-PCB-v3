import gradio as gr
import pandas as pd
import json
try:    # for running interface.py
    import utils
    from config import mistral_params, bitrix24_config
    from logger import setup_logger
    import bitrix24
except: # for running main.py
    from . import utils
    from .config import mistral_params, bitrix24_config
    from .logger import setup_logger
    from . import bitrix24

logger = setup_logger()


def _file_basename(file):
    """Имя файла без расширения для сохранения результатов."""
    path = getattr(file, "name", None) or str(file)
    return path.rsplit(".", 1)[0] if "." in path else path


def show_outputs():
    logger.info("Processing done")
    return gr.update(visible=True), gr.update(visible=True), \
        gr.update(visible=True), gr.update(visible=True), gr.update(visible=True)

def hide_outputs():
    logger.debug("File was closed")
    return gr.update(value=pd.DataFrame(), visible=False), gr.update(visible=False), \
        gr.update(visible=False), gr.update(visible=False), gr.update(visible=False), \
        gr.update(visible=False)

# Глобальная переменная для хранения распарсенных данных
_parsed_pcb_data = None

def parse_excel_pcb(file):
    """
    Извлекает данные из загруженного файла (Excel или Word), обрабатывает через LLM
    для извлечения характеристик ПП и сохраняет результаты в CSV, Excel и JSON.
    """
    global _parsed_pcb_data
    if isinstance(file, list) and file:
        file = file[0]
    logger.info("Starting to parse file for PCB characteristics: %s", getattr(file, "name", file))
    
    try:
        doc_txt = utils.extract_document_data(file)
        logger.debug("Extracted document data")

        llm = utils.create_pcb_model(mistral_params)
        logger.debug("PCB model created successfully.")

        parsed_dict = utils.process_excel_pcb(doc_txt, llm)
        logger.debug("Parsed PCB dictionary: %s", parsed_dict)
        
        _parsed_pcb_data = parsed_dict

        fn = _file_basename(file)
        path_ext = lambda ext: f"{fn}_pcb_parsed.{ext}"
        csv_path = path_ext("csv")
        xlsx_path = path_ext("xlsx")
        json_path = path_ext("json")
        bitrix24_json_path = f"{fn}_bitrix24.json"

        df = pd.DataFrame(list(parsed_dict.items()), columns=['Characteristic', 'Value'])
        df.to_csv(csv_path, index=False)
        df.to_excel(xlsx_path, index=False)
        df.to_json(json_path, index=False)
        
        # Создаем JSON файл в формате Битрикс24
        bitrix24_fields = bitrix24.map_pcb_to_bitrix24_fields(parsed_dict)
        bitrix24_payload = {
            "entityTypeId": 182,
            "fields": bitrix24_fields
        }
        with open(bitrix24_json_path, 'w', encoding='utf-8') as f:
            json.dump(bitrix24_payload, f, ensure_ascii=False, indent=2)
        logger.info(f"Создан JSON файл для Битрикс24: {bitrix24_json_path}")

    except Exception as e:
        logger.error("An error occurred while parsing the Excel file: %s", e)
        error_msg = str(e)
        if "401" in error_msg or "unauthorized" in error_msg.lower():
            raise Exception(
                "Mistral API вернул 401 Unauthorized. Проверьте, что переменная окружения "
                "`MISTRAL_API_KEY` задана и ключ действителен."
            )
        if "capacity exceeded" in error_msg.lower() or "429" in error_msg:
            raise Exception("Сервис временно недоступен из-за высокого спроса. Пожалуйста, попробуйте позже или обновите API ключ.")
        else:
            raise e
    return df, csv_path, xlsx_path, json_path, bitrix24_json_path


def send_to_bitrix24():
    """
    Отправляет распарсенные данные PCB в Битрикс24.
    
    Returns:
        str: Сообщение о результате отправки
    """
    global _parsed_pcb_data
    
    if not _parsed_pcb_data:
        return "Ошибка: Сначала необходимо загрузить и распарсить файл (Excel или Word)."
    
    # Приоритет: webhook_url > token
    webhook_url = bitrix24_config.get("webhook_url", "").strip()
    token = bitrix24_config.get("token", "").strip()
    
    webhook_url_or_token = webhook_url if webhook_url else token
    
    if not webhook_url_or_token:
        return (
            "Ошибка: Webhook URL или токен Битрикс24 не задан.\n"
            "Установите переменную окружения BITRIX24_WEBHOOK_URL или BITRIX24_TOKEN.\n"
            "Формат webhook URL: https://fineline.bitrix24.ru/rest/6/<token>/crm.item.add"
        )
    
    try:
        logger.info("Отправка данных в Битрикс24...")
        result = bitrix24.send_pcb_to_bitrix24(_parsed_pcb_data, webhook_url_or_token)
        
        item_id = result.get("result", {}).get("item", {}).get("id")
        if item_id:
            return f"✅ Успешно отправлено в Битрикс24! ID элемента: {item_id}"
        else:
            return f"⚠️ Данные отправлены, но ID не получен. Ответ: {result}"
            
    except Exception as e:
        logger.error(f"Ошибка при отправке в Битрикс24: {e}")
        error_msg = str(e)
        if "401" in error_msg or "unauthorized" in error_msg.lower():
            return "❌ Ошибка авторизации в Битрикс24. Проверьте webhook URL или токен."
        return f"❌ Ошибка при отправке в Битрикс24: {error_msg}"


def create_interface(title: str = "gradio app"):
    interface = gr.Blocks(title=title)
    with interface:
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
                    "- **Поддерживаемые форматы**: `.xlsx`, `.xls`, `.docx`, `.doc`\n"
                    "- Файлы формата Word могут быть в виде листа технических требований ПП "
                    "на русском и/или английском языке."
                )
                excel_input = gr.File(
                    label="Загрузить файл спецификации (Excel или Word)",
                    file_types=["file"],
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
                    label="Статус отправки в Битрикс24",
                    visible=False,
                    interactive=False
                )
                bitrix24_send_btn = gr.Button(
                    value="Отправить распознанные данные в Битрикс24",
                    visible=False,
                    variant="secondary"
                )

            # Правая колонка: результаты и выгрузки
            with gr.Column(scale=2):
                gr.Markdown("### Шаг 2. Результаты распознавания")
                excel_parsed_reports = gr.DataFrame(
                    label="Распознанные характеристики печатной платы",
                    show_copy_button=True,
                    visible=False,
                    min_width=10
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

        # Excel processing events
        excel_input.upload(lambda: gr.update(visible=True), None, excel_process_btn)
        excel_process_btn.click(parse_excel_pcb, [excel_input], 
                               [excel_parsed_reports, excel_download_csv, excel_download_xlsx, excel_download_json, excel_download_bitrix24_json], queue=True)
        excel_process_btn.click(show_outputs, None, [excel_parsed_reports, excel_download_csv, excel_download_xlsx, excel_download_json, excel_download_bitrix24_json], queue=True)
        excel_process_btn.click(
            lambda: (gr.update(visible=True), gr.update(visible=True)),
            None,
            [bitrix24_status, bitrix24_send_btn],
            queue=True
        )
        excel_input.clear(hide_outputs, None, [excel_parsed_reports, excel_download_csv, excel_download_xlsx, excel_download_json, excel_download_bitrix24_json, excel_process_btn])
        excel_input.clear(
            lambda: (gr.update(visible=False), gr.update(visible=False)),
            None,
            [bitrix24_status, bitrix24_send_btn]
        )
        
        # Битрикс24 events
        bitrix24_send_btn.click(
            send_to_bitrix24,
            None,
            bitrix24_status,
            queue=True
        )

    return interface

if __name__ == "__main__":
    create_interface().launch()
        