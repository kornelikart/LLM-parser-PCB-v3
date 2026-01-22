# Configuration for Mistral AI API and Bitrix24
# Prefer environment variables so secrets are not stored in code.
import os

# Попытка загрузить .env файл, если он существует
try:
    from dotenv import load_dotenv
    load_dotenv()  # Загружает переменные из .env файла в корне проекта
except ImportError:
    # python-dotenv не установлен, используем только системные переменные окружения
    pass

# ⚠️ ВНИМАНИЕ: API ключ в коде - это небезопасно для продакшена!
# Приоритет: переменная окружения > значение по умолчанию в коде
# Для продакшена используйте переменные окружения или .env файл
_DEFAULT_MISTRAL_API_KEY = "z2bE6jVqaj6u1JlweQ3w29uUIwvyp2XA"

_env_api_key = os.getenv("MISTRAL_API_KEY", _DEFAULT_MISTRAL_API_KEY).strip()
_bitrix24_webhook_url = os.getenv("BITRIX24_WEBHOOK_URL", "").strip()
_bitrix24_token = os.getenv("BITRIX24_TOKEN", "").strip()

mistral_params = {
    "api_key": _env_api_key,
}

# Приоритет: webhook URL > token
bitrix24_config = {
    "webhook_url": _bitrix24_webhook_url,
    "token": _bitrix24_token,
    "base_url": "https://fineline.bitrix24.ru/rest/6",
    "entity_type_id": 182,
}

# Значения по умолчанию для обязательных полей Битрикс24 (если не найдены в данных)
# Эти значения должны существовать в соответствующих справочниках БД
bitrix24_defaults = {
    "order_unit_id": int(os.getenv("BITRIX24_DEFAULT_ORDER_UNIT_ID", "5256")),  # Order unit (IBLOCK_ID=50)
    "pcb_type_id": int(os.getenv("BITRIX24_DEFAULT_PCB_TYPE_ID", "6610")),  # PCB type (IBLOCK_ID=52)
    "peelable_sm_id": int(os.getenv("BITRIX24_DEFAULT_PEELABLE_SM_ID", "6014")),  # Peelable SM (IBLOCK_ID=86)
    "production_unit_id": int(os.getenv("BITRIX24_DEFAULT_PRODUCTION_UNIT_ID", "6270")),  # Production Unit (IBLOCK_ID=160)
}
