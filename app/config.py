# Configuration for Mistral AI API and Bitrix24
# Prefer environment variables so secrets are not stored in code.
import os

_env_api_key = os.getenv("MISTRAL_API_KEY", "").strip()
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
