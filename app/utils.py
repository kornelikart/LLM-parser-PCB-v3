from langchain_mistralai import ChatMistralAI
import httpx
import pandas as pd
import logging
import time
from typing import Optional, Union
import os
try:    # for running interface.py
    from model import PCBCharacteristics
    from logger import setup_logger
except ImportError:  # for running main.py
    from .model import PCBCharacteristics
    from .logger import setup_logger

logger = setup_logger(level=logging.INFO)


def _get_file_path(file) -> str:
    """Return path to file for reading. Supports path string or file-like with .name."""
    if isinstance(file, str):
        return file
    if hasattr(file, "name"):
        return file.name
    return str(file)


def extract_word_data(file) -> str:
    """Извлекает текст из документа Word (.docx): параграфы и таблицы."""
    logger.info("Извлечение данных из Word файла.")
    path = _get_file_path(file)
    try:
        from docx import Document
        doc = Document(path)
        parts = []
        for p in doc.paragraphs:
            text = (p.text or "").strip()
            if text:
                parts.append(text)
        for table in doc.tables:
            rows_text = []
            for row in table.rows:
                cells = [(cell.text or "").strip().replace("\n", " ") for cell in row.cells]
                if any(cells):
                    rows_text.append("\t".join(cells))
            if rows_text:
                parts.append("\n".join(rows_text))
        result = "\n\n".join(parts)
        logger.info("Word: извлечено символов %s, слов %s", len(result), len(result.split()))
        return result
    except Exception as e:
        logger.error("Ошибка чтения Word файла: %s", e)
        raise e


def _convert_doc_to_docx(doc_path: str, out_dir: str) -> str:
    """Конвертирует .doc в .docx через LibreOffice или doc2docx. Возвращает путь к .docx."""
    import os
    import subprocess
    import shutil
    base = os.path.splitext(os.path.basename(doc_path))[0]
    docx_path = os.path.join(out_dir, base + ".docx")
    # 1) LibreOffice (если установлен)
    soffice_candidates = [
        os.path.expandvars(r"%ProgramFiles%\LibreOffice\program\soffice.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\LibreOffice\program\soffice.exe"),
        "soffice",
        "libreoffice",
    ]
    for soffice in soffice_candidates:
        if soffice in ("soffice", "libreoffice"):
            exe = shutil.which(soffice)
            if not exe:
                continue
            soffice = exe
        elif not os.path.isfile(soffice):
            continue
        try:
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    "--convert-to", "docx",
                    "--outdir", out_dir,
                    doc_path,
                ],
                check=True,
                capture_output=True,
                timeout=60,
                creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
            )
            if os.path.isfile(docx_path):
                return docx_path
        except (subprocess.CalledProcessError, FileNotFoundError, OSError, subprocess.TimeoutExpired) as e:
            logger.debug("LibreOffice не сработал (%s): %s", soffice, e)
            continue
    # 2) doc2docx (требует установленный Word на Windows или LibreOffice)
    try:
        from doc2docx import convert
        convert(doc_path, docx_path)
        if os.path.isfile(docx_path):
            return docx_path
    except Exception as e:
        logger.debug("doc2docx не сработал: %s", e)
    raise RuntimeError(
        "Не удалось конвертировать .doc в .docx. Установите LibreOffice "
        "(https://www.libreoffice.org/) или сохраните документ как .docx в Word."
    )


def extract_word97_data(file) -> str:
    """Извлекает текст из Word 97–2003 (.doc). Пробует Win32 COM (Word), затем конвертацию в .docx."""
    import os
    import sys
    path = _get_file_path(file)
    path = os.path.abspath(path)
    logger.info("Извлечение данных из Word 97-2003 (.doc).")
    if sys.platform == "win32":
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(path, ReadOnly=True)
                try:
                    text = doc.Content.Text
                    if text and text.strip():
                        logger.info("Word 97-2003: текст извлечён через MS Word (COM).")
                        return text.strip()
                finally:
                    doc.Close(False)
            finally:
                word.Quit()
        except Exception as e:
            logger.debug("Извлечение .doc через Word COM не удалось: %s", e)
    import shutil
    import tempfile
    tmpdir = tempfile.mkdtemp()
    try:
        doc_copy = os.path.join(tmpdir, "input.doc")
        shutil.copy2(path, doc_copy)
        docx_path = _convert_doc_to_docx(doc_copy, tmpdir)
        return extract_word_data(docx_path)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def extract_document_data(file) -> str:
    """Извлекает текст из Excel или Word по расширению файла."""
    path = _get_file_path(file)
    path_lower = path.lower()
    if path_lower.endswith(".docx"):
        return extract_word_data(file)
    if path_lower.endswith(".doc"):
        return extract_word97_data(file)
    if path_lower.endswith((".xlsx", ".xls")):
        return extract_excel_data(file)
    raise ValueError(
        "Неподдерживаемый формат. Используйте файл .xlsx, .xls, .docx или .doc (Лист технических требований ПП)."
    )


def extract_excel_data(file) -> str:
    """Extracts data from an Excel file and converts it to a string.

    This function reads all sheets of the specified Excel file, processes the data by removing 
    empty rows and columns, and concatenates the non-empty data into a single string.

    Args:
        file: The path to the Excel file or a file-like object from which to extract data.

    Returns:
        str: A string representation of the extracted data, with each sheet's content 
             concatenated and formatted without indices or headers.
    """
    logger.info("Starting to extract data from the Excel file.")
    
    try:
        # Read all sheets from Excel file
        excel_file = pd.ExcelFile(file)
        logger.debug("Number of sheets in Excel file: %s", len(excel_file.sheet_names))
        
        all_data = []
        for sheet_name in excel_file.sheet_names:
            logger.debug("Processing sheet: %s", sheet_name)
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # Remove empty rows and columns
            df_clean = df.dropna(how='all').dropna(axis=1, how='all')
            
            if not df_clean.empty:
                # Convert DataFrame to string representation
                sheet_data = df_clean.to_string(index=False, header=True, na_rep="")
                all_data.append(f"Sheet: {sheet_name}\n{sheet_data}")
        
        # Combine all sheet data
        excel_txt = "\n\n".join(all_data)
        logger.info("Extracted text length: %s, Word count: %s", len(excel_txt), len(excel_txt.split()))
        return excel_txt
        
    except Exception as e:
        logger.error("Error reading Excel file: %s", e)
        raise e


def create_pcb_model(params: dict[str, str]) -> ChatMistralAI:
    """Creates and configures a ChatMistralAI model instance for PCB characteristics parsing.

    Args:
        params (dict[str, str]): A dictionary containing parameters for model configuration.
            Expected keys:
                - 'api_key': The API key for authenticating with the ChatMistralAI service.

    Returns:
        ChatMistralAI: An instance of the ChatMistralAI model configured for PCB characteristics parsing.
    """
    api_key = (params.get("api_key") or "").strip()
    # Логируем префикс ключа (безопасно) чтобы понять, какой ключ реально используется в сервере.
    key_prefix = (api_key[:6] + "...") if api_key else None
    logger.info("MISTRAL_API_KEY prefix: %s (len=%s)", key_prefix, len(api_key))
    if api_key == "mistral_api_key":
        logger.warning("Используется плейсхолдер MISTRAL_API_KEY. Проверьте загрузку .env.")
    if not api_key:
        raise ValueError(
            "Mistral API key is empty. Set environment variable `MISTRAL_API_KEY` "
            "before starting the app."
        )

    # Отключаем системные proxy-настройки (trust_env=True),
    # т.к. в некоторых окружениях они приводят к WinError 10061.
    # Важно: если передать client в ChatMistralAI, он НЕ выставляет base_url сам,
    # поэтому у клиента должен быть base_url.
    base_url = (os.getenv("MISTRAL_BASE_URL") or "").strip() or "https://api.mistral.ai/v1"
    http_client = httpx.Client(
        base_url=base_url,
        timeout=30.0,
        trust_env=False,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
    )
    llm = ChatMistralAI(
        model="mistral-medium-latest",
        temperature=0.1,
        api_key=api_key,
        client=http_client,
    )
    return llm.with_structured_output(PCBCharacteristics)


def process_excel_pcb_with_retry(excel_txt: str, llm_parser: ChatMistralAI, max_retries: int = 3, delay: float = 2.0) -> Optional[dict]:
    """Processes Excel data for PCB characteristics using a ChatMistralAI model with retry logic.

    Args:
        excel_txt (str): A string containing the Excel data to be processed.
        llm_parser (ChatMistralAI): An instance of the ChatMistralAI model used for parsing PCB data.
        max_retries (int): Maximum number of retry attempts.
        delay (float): Delay between retries in seconds.

    Returns:
        dict: A dictionary containing the processed PCB characteristics, or None if all retries failed.
    """
    messages = [
        (
            "system",
            "You are an experienced PCB engineer. Extract PCB characteristics from the provided data. "
            "Data may come from Excel or from a Word 'Technical Requirements Sheet PCB' / 'Лист технических требований ПП' (Russian/English). "
            "Extract at least the following fields into the structured PCBCharacteristics object: "
            "company_name (company name), board_name (board name / Identification / Обозначение), quantity (boards quantity, if present), "
            "base_material (Base material / Тип материала), "
            "board_thickness (finished PCB thickness with tolerance, from 'Finished thickness with tolerance, mm' / 'Толщина с допуском, мм'), "
            "foil_thickness (from CU foil thickness + electroplating / Толщина медной фольги / Thickness CU), "
            "layer_count (from Layer PCB / Слои топологии - count signal layers like Layer 1, Layer 2, etc.), "
            "board_size (Size with tolerance / (Длина x ширина) with tolerance), "
            "coverage_type (surface finish, from 'Surface Finish' / 'Финишное покрытие платы'), "
            "solder_mask_colour (solder mask colour from Material, color / Материал, цвет), "
            "solder_mark_colour (marking / legend colour), "
            "electrical_testing (information about electrical test / электроконтроль), "
            "panelization (panel / panelization info), edge_plating, contour_treatment and any other PCB fields defined in the schema. "
            "If any information is missing, use empty string or 0 as appropriate."
        ),
        ("human", excel_txt),
    ]
    
    for attempt in range(max_retries):
        try:
            logger.info("Attempting to process PCB data (attempt %d/%d)", attempt + 1, max_retries)
            answer = llm_parser.invoke(messages)
            logger.info("Successfully processed PCB data")
            return answer.model_dump()

        except Exception as e:
            error_msg = str(e)
            logger.warning("Attempt %d failed: %s", attempt + 1, error_msg)

            # Check if it's a rate limit error
            if "429" in error_msg or "capacity exceeded" in error_msg.lower():
                if attempt < max_retries - 1:
                    wait_time = delay * (2 ** attempt)  # Exponential backoff
                    logger.info("Rate limit exceeded. Waiting %.1f seconds before retry...", wait_time)
                    time.sleep(wait_time)
                    continue
                else:
                    logger.error("All retry attempts failed due to rate limiting")
                    raise Exception("Service temporarily unavailable due to high demand. Please try again later.")
            else:
                # For other errors, don't retry
                logger.error("Non-retryable error: %s", error_msg)
                raise e
    
    return None


