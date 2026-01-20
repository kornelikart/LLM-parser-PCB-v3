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
    This function extracts data from a given Excel file, processes it
    using a language model to extract PCB characteristics, and saves the results 
    into CSV, Excel, and JSON formats.

    Args:
        file (file-like object): The Excel file to be parsed.

    Returns:
        tuple: A tuple containing:
            - pd.DataFrame: A DataFrame with the parsed PCB characteristics.
            - str: The name of CSV file.
            - str: The name of Excel file.
            - str: The name of JSON file.

    Raises:
        Exception: If an error occurs during the parsing process.
    """
    global _parsed_pcb_data
    logger.info("Starting to parse Excel file for PCB characteristics: %s", file.name)
    
    try:
        excel_txt = utils.extract_excel_data(file)
        logger.debug("Extracted Excel data text")

        llm = utils.create_pcb_model(mistral_params)
        logger.debug("PCB model created successfully.")

        parsed_dict = utils.process_excel_pcb(excel_txt, llm)
        logger.debug("Parsed PCB dictionary: %s", parsed_dict)
        
        # Сохраняем данные для отправки в Битрикс24
        _parsed_pcb_data = parsed_dict

        fn = file.name.split(".")[0]
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
        return "Ошибка: Сначала необходимо распарсить Excel файл."
    
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
        gr.Markdown("# LLM-Parser: Характеристики печатных плат")
        
        gr.Markdown("## Парсинг характеристик печатных плат из Excel файлов")
        
        # Информационное сообщение о возможных задержках
        gr.Markdown("""
        **Примечание:** Обработка может занять некоторое время из-за использования внешнего AI сервиса. 
        При превышении лимитов запросов система автоматически повторит попытку.
        """)
        
        excel_input = gr.File(label="Загрузить Excel файл", file_types=[".xlsx", ".xls"], height=160)
        excel_process_btn = gr.Button(value="Парсить Excel данные", visible=False, variant="primary")
        excel_parsed_reports = gr.DataFrame(label="Распарсенные характеристики печатных плат", 
                                          show_copy_button=True, 
                                          visible=False, min_width=10)
        with gr.Row():
            excel_download_csv = gr.File(label="Скачать как CSV", visible=False)
            excel_download_xlsx = gr.File(label="Скачать как XLSX", visible=False)
            excel_download_json = gr.File(label="Скачать как JSON", visible=False)
        excel_download_bitrix24_json = gr.File(
            label="Скачать JSON для Битрикс24", 
            visible=False,
            file_types=[".json"]
        )
        
        # Интеграция с Битрикс24
        gr.Markdown("---")
        gr.Markdown("## Отправка в Битрикс24")
        bitrix24_status = gr.Textbox(
            label="Статус отправки",
            visible=False,
            interactive=False
        )
        bitrix24_send_btn = gr.Button(
            value="Отправить в Битрикс24",
            visible=False,
            variant="secondary"
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
        