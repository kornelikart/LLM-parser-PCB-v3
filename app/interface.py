import gradio as gr
import pandas as pd
try:    # for running interface.py
    import utils
    from config import mistral_params
    from logger import setup_logger
except: # for running main.py
    from . import utils
    from .config import mistral_params
    from .logger import setup_logger

logger = setup_logger()


def show_outputs():
    logger.info("Processing done")
    return gr.update(visible=True), gr.update(visible=True), \
        gr.update(visible=True), gr.update(visible=True)

def hide_outputs():
    logger.debug("File was closed")
    return gr.update(value=pd.DataFrame(), visible=False), gr.update(visible=False), \
        gr.update(visible=False), gr.update(visible=False), gr.update(visible=False)

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
    logger.info("Starting to parse Excel file for PCB characteristics: %s", file.name)
    
    try:
        excel_txt = utils.extract_excel_data(file)
        logger.debug("Extracted Excel data text")

        llm = utils.create_pcb_model(mistral_params)
        logger.debug("PCB model created successfully.")

        parsed_dict = utils.process_excel_pcb(excel_txt, llm)
        logger.debug("Parsed PCB dictionary: %s", parsed_dict)

        fn = file.name.split(".")[0]
        path_ext = lambda ext: f"{fn}_pcb_parsed.{ext}"
        csv_path = path_ext("csv")
        xlsx_path = path_ext("xlsx")
        json_path = path_ext("json")

        df = pd.DataFrame(list(parsed_dict.items()), columns=['Characteristic', 'Value'])
        df.to_csv(csv_path, index=False)
        df.to_excel(xlsx_path, index=False)
        df.to_json(json_path, index=False)

    except Exception as e:
        logger.error("An error occurred while parsing the Excel file: %s", e)
        error_msg = str(e)
        if "capacity exceeded" in error_msg.lower() or "429" in error_msg:
            raise Exception("Сервис временно недоступен из-за высокого спроса. Пожалуйста, попробуйте позже или обновите API ключ.")
        else:
            raise e
    return df, csv_path, xlsx_path, json_path


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

        # Excel processing events
        excel_input.upload(lambda: gr.update(visible=True), None, excel_process_btn)
        excel_process_btn.click(parse_excel_pcb, [excel_input], 
                               [excel_parsed_reports, excel_download_csv, excel_download_xlsx, excel_download_json], queue=True)
        excel_process_btn.click(show_outputs, None, [excel_parsed_reports, excel_download_csv, excel_download_xlsx, excel_download_json], queue=True)
        excel_input.clear(hide_outputs, None, [excel_parsed_reports, excel_download_csv, excel_download_xlsx, excel_download_json, excel_process_btn])

    return interface

if __name__ == "__main__":
    create_interface().launch()
        