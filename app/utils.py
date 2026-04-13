from langchain_mistralai import ChatMistralAI
import httpx
import pandas as pd
import logging
import time
import struct
import re
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


def _extract_doc_text_ole(doc_path: str) -> str:
    """Извлекает текст из .doc через OLE-парсинг (без MS Word и LibreOffice).
    Реализует разбор piece table из Word Binary File Format (Word 97-2003).
    """
    try:
        import olefile
    except ImportError:
        logger.debug("olefile не установлен, OLE-извлечение недоступно.")
        return ""
    try:
        if not olefile.isOleFile(doc_path):
            return ""
        with olefile.OleFileIO(doc_path) as ole:
            if not ole.exists("WordDocument"):
                return ""
            wd = ole.openstream("WordDocument").read()
            if len(wd) < 0x01AA:
                return ""

            # FIB flags (offset 10): bit 9 (0x0200) = fWhichTblStm (0→0Table, 1→1Table)
            flags = struct.unpack_from("<H", wd, 10)[0]
            use_1table = bool(flags & 0x0200)

            # ccpText: количество символов основного текста (offset 0x004C)
            ccpText = struct.unpack_from("<I", wd, 0x004C)[0]
            if not ccpText:
                return ""

            # Параметры CLX (piece table) в Table-stream
            fcClx  = struct.unpack_from("<I", wd, 0x01A2)[0]
            lcbClx = struct.unpack_from("<I", wd, 0x01A6)[0]

            table_name = "1Table" if use_1table else "0Table"
            if not ole.exists(table_name):
                return ""
            table = ole.openstream(table_name).read()

            clx = table[fcClx: fcClx + lcbClx]

            # Пропускаем Prc-записи (тип 0x01)
            pos = 0
            while pos < len(clx) and clx[pos] == 0x01:
                cb = struct.unpack_from("<H", clx, pos + 1)[0]
                pos += 3 + cb

            # PlcPcd: тип 0x02, затем 4 байта размера, затем данные
            if pos >= len(clx) or clx[pos] != 0x02:
                return ""
            pcdt_size = struct.unpack_from("<I", clx, pos + 1)[0]
            pcdt = clx[pos + 5: pos + 5 + pcdt_size]

            # Количество кусков: (размер - 4) / 12 (n+1 CP по 4 байта + n PCD по 8 байт)
            n_pieces = (len(pcdt) - 4) // 12
            if n_pieces <= 0:
                return ""

            cps = [struct.unpack_from("<I", pcdt, i * 4)[0] for i in range(n_pieces + 1)]
            cp_base = (n_pieces + 1) * 4

            texts = []
            for i in range(n_pieces):
                pcd = pcdt[cp_base + i * 8: cp_base + i * 8 + 8]
                if len(pcd) < 8:
                    break
                fc_raw = struct.unpack_from("<I", pcd, 2)[0]
                is_ansi = bool(fc_raw & 0x40000000)
                fc = fc_raw & 0x3FFFFFFF
                char_count = cps[i + 1] - cps[i]
                if is_ansi:
                    raw = wd[fc >> 1: (fc >> 1) + char_count]
                    piece_text = raw.decode("cp1252", errors="replace")
                else:
                    raw = wd[fc: fc + char_count * 2]
                    piece_text = raw.decode("utf-16-le", errors="replace")
                texts.append(piece_text)

            raw_text = "".join(texts)
            # Убираем управляющие символы, оставляем переносы строк
            raw_text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", " ", raw_text)
            raw_text = raw_text.replace("\r", "\n")
            raw_text = re.sub(r" {3,}", "  ", raw_text)
            raw_text = re.sub(r"\n{3,}", "\n\n", raw_text)
            result = raw_text.strip()
            if result:
                logger.info(".doc OLE: извлечено %d символов.", len(result))
            return result
    except Exception as e:
        logger.debug("OLE-извлечение .doc не удалось: %s", e)
        return ""


def extract_word97_data(file) -> str:
    """Извлекает текст из Word 97–2003 (.doc).
    Порядок: Win32 COM → LibreOffice/doc2docx → OLE-парсинг (pure Python).
    """
    import os
    import sys
    path = _get_file_path(file)
    path = os.path.abspath(path)
    logger.info("Извлечение данных из Word 97-2003 (.doc).")

    # 1) Win32 COM через установленный MS Word
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
                        logger.info(".doc: текст извлечён через MS Word (COM).")
                        return text.strip()
                finally:
                    doc.Close(False)
            finally:
                word.Quit()
        except Exception as e:
            logger.debug("Извлечение .doc через Word COM не удалось: %s", e)

    # 2) LibreOffice / doc2docx — конвертация в .docx
    import shutil
    import tempfile
    try:
        tmpdir = tempfile.mkdtemp()
        try:
            doc_copy = os.path.join(tmpdir, "input.doc")
            shutil.copy2(path, doc_copy)
            docx_path = _convert_doc_to_docx(doc_copy, tmpdir)
            return extract_word_data(docx_path)
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)
    except Exception as e:
        logger.debug("Конвертация .doc не удалась: %s", e)

    # 3) Pure-Python OLE-парсинг (без внешних инструментов)
    text = _extract_doc_text_ole(path)
    if text:
        logger.info(".doc: текст извлечён через OLE-парсинг.")
        return text

    raise RuntimeError(
        "Не удалось прочитать .doc файл. Пожалуйста, пересохраните документ как .docx в Microsoft Word."
    )


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
        # Увеличен timeout: LLM-ответы для больших документов могут занимать > 30 сек.
        timeout=httpx.Timeout(connect=15.0, read=120.0, write=30.0, pool=15.0),
        # trust_env=True — разрешаем системные proxy (важно в корпоративных сетях).
        # WinError 10061 (отказ подключения) → проверьте MISTRAL_BASE_URL или proxy-настройки.
        trust_env=True,
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
            logger.warning("Attempt %d/%d failed: %s", attempt + 1, max_retries, error_msg)

            # Определяем тип ошибки
            is_rate_limit = "429" in error_msg or "capacity exceeded" in error_msg.lower()
            # WinError 10054 = connection reset, 10061 = connection refused — временные сетевые сбои
            is_network_error = any(x in error_msg for x in (
                "10054", "10061", "ReadError", "ConnectError",
                "TimeoutException", "RemoteProtocolError", "ConnectionReset",
                "Connection reset", "forcibly closed",
            ))

            if is_rate_limit or is_network_error:
                if attempt < max_retries - 1:
                    wait_time = delay * (2 ** attempt)  # Exponential backoff
                    reason = "Rate limit" if is_rate_limit else "Network error"
                    logger.info("%s. Retry in %.1f sec...", reason, wait_time)
                    time.sleep(wait_time)
                    continue
                else:
                    logger.error("All %d attempts failed.", max_retries)
                    if is_rate_limit:
                        raise Exception("Сервис временно недоступен (превышен лимит запросов). Попробуйте позже.")
                    raise Exception(
                        f"Ошибка сети при обращении к Mistral API: {error_msg}\n"
                        "Проверьте интернет-соединение и настройки proxy/антивируса."
                    )
            else:
                # Нереентерабельная ошибка (неверный ключ, невалидный запрос и т.д.)
                logger.error("Non-retryable error: %s", error_msg)
                raise e

    raise RuntimeError("Не удалось обработать данные PCB после всех попыток.")


