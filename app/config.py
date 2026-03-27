# Configuration for Mistral AI API and Bitrix24
# Prefer environment variables so secrets are not stored in code.
import os
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# Попытка загрузить .env файл, если он существует
# (безопасно для разных cwd: путь считается относительно этого файла).
project_root = Path(__file__).resolve().parent.parent
dotenv_path = project_root / ".env"

def _load_env_file_manually(path: Path) -> None:
    if not path.exists():
        return
    try:
        raw = path.read_text(encoding="utf-8")
    except Exception:
        return

    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            continue
        k, v = line.split("=", 1)
        k = k.strip()
        v = v.strip().strip("'").strip('"')
        # Не перезаписываем уже заданные переменные окружения.
        if k and not os.getenv(k):
            os.environ[k] = v

try:
    from dotenv import load_dotenv  # type: ignore
    load_dotenv(dotenv_path=str(dotenv_path))
except ImportError:
    _load_env_file_manually(dotenv_path)

# ⚠️ ВНИМАНИЕ: API ключ в коде - это небезопасно для продакшена!
# Приоритет: переменная окружения > значение по умолчанию в коде
# Для продакшена используйте переменные окружения или .env файл
_DEFAULT_MISTRAL_API_KEY = "mistral_api_key"

_env_api_key = os.getenv("MISTRAL_API_KEY", _DEFAULT_MISTRAL_API_KEY).strip()
_bitrix24_webhook_url = os.getenv("BITRIX24_WEBHOOK_URL", "").strip()
_bitrix24_token = os.getenv("BITRIX24_TOKEN", "").strip()

mistral_params = {
    "api_key": _env_api_key,
}

if _env_api_key == _DEFAULT_MISTRAL_API_KEY:
    logger.warning(
        "MISTRAL_API_KEY не задан (используется плейсхолдер). Проверьте переменные окружения или файл .env в корне проекта."
    )

# Приоритет: webhook URL > token
bitrix24_config = {
    "webhook_url": _bitrix24_webhook_url,
    "token": _bitrix24_token,
    "base_url": "https://fineline.bitrix24.ru/rest/6",
    "entity_type_id": 182,
}

