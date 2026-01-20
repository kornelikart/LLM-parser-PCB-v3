from langchain_mistralai import ChatMistralAI
import pandas as pd
import logging
import time
from typing import Optional
try:    # for running interface.py
    from model import PCBCharacteristics
    from logger import setup_logger
except: # for running main.py
    from .model import PCBCharacteristics
    from .logger import setup_logger

logger = setup_logger(level=logging.INFO)

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
        logger.debug(f"Number of sheets in Excel file: {len(excel_file.sheet_names)}")
        
        all_data = []
        for sheet_name in excel_file.sheet_names:
            logger.debug(f"Processing sheet: {sheet_name}")
            df = pd.read_excel(file, sheet_name=sheet_name)
            
            # Remove empty rows and columns
            df_clean = df.dropna(how='all').dropna(axis=1, how='all')
            
            if not df_clean.empty:
                # Convert DataFrame to string representation
                sheet_data = df_clean.to_string(index=False, header=True, na_rep="")
                all_data.append(f"Sheet: {sheet_name}\n{sheet_data}")
        
        # Combine all sheet data
        excel_txt = "\n\n".join(all_data)
        logger.info(f"Extracted text length: {len(excel_txt)}, Word count: {len(excel_txt.split())}")
        return excel_txt
        
    except Exception as e:
        logger.error(f"Error reading Excel file: {e}")
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
    if not api_key:
        raise ValueError(
            "Mistral API key is empty. Set environment variable `MISTRAL_API_KEY` "
            "before starting the app."
        )

    llm = ChatMistralAI(
        model="mistral-medium-latest",
        temperature=0.1,
        api_key=api_key,
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
        ("system", "You are an experienced PCB engineer. Extract PCB characteristics from the provided data including: company name, board name, quantity, base material, foil thickness, layer count, board size, and panelization. If any information is missing, use empty string or 0 as appropriate."),
        ("human", excel_txt)
    ]
    
    for attempt in range(max_retries):
        try:
            logger.info(f"Attempting to process PCB data (attempt {attempt + 1}/{max_retries})")
            answer = llm_parser.invoke(messages)
            logger.info("Successfully processed PCB data")
            return answer.model_dump()
            
        except Exception as e:
            error_msg = str(e)
            logger.warning(f"Attempt {attempt + 1} failed: {error_msg}")
            
            # Check if it's a rate limit error
            if "429" in error_msg or "capacity exceeded" in error_msg.lower():
                if attempt < max_retries - 1:
                    wait_time = delay * (2 ** attempt)  # Exponential backoff
                    logger.info(f"Rate limit exceeded. Waiting {wait_time} seconds before retry...")
                    time.sleep(wait_time)
                    continue
                else:
                    logger.error("All retry attempts failed due to rate limiting")
                    raise Exception("Service temporarily unavailable due to high demand. Please try again later.")
            else:
                # For other errors, don't retry
                logger.error(f"Non-retryable error occurred: {error_msg}")
                raise e
    
    return None


def process_excel_pcb(excel_txt: str, llm_parser: ChatMistralAI) -> dict:
    """Processes Excel data for PCB characteristics using a ChatMistralAI model.

    Args:
        excel_txt (str): A string containing the Excel data to be processed.
        llm_parser (ChatMistralAI): An instance of the ChatMistralAI model used for parsing PCB data.

    Returns:
        dict: A dictionary containing the processed PCB characteristics.
    """
    return process_excel_pcb_with_retry(excel_txt, llm_parser)
