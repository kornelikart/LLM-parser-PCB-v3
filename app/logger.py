"""
Настройка логирования приложения.

Логи пишутся в файл (с ротацией) и в консоль. Настройка выполняется один раз
на процесс, независимо от того, сколько модулей вызвало setup_logger().

Важно: хендлеры вешаются на КОРНЕВОЙ логгер, поэтому в общий лог попадают и
сообщения модулей, которые используют logging.getLogger(__name__) — например,
диффы нормализации из pcb_normalizer ('ENIG' → 'Imm. gold (chem.Ni/Au)').

Переменные окружения:
    LOG_FILE       путь к файлу лога (по умолчанию logs.log в корне проекта)
    LOG_LEVEL      DEBUG / INFO / WARNING / ERROR (по умолчанию INFO)
    LOG_MAX_BYTES  размер файла до ротации (по умолчанию 5 МБ)
"""
import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
_configured = False


def _resolve_level(level) -> int:
    """Уровень из аргумента или переменной LOG_LEVEL (она приоритетнее)."""
    env_level = (os.getenv("LOG_LEVEL") or "").strip().upper()
    if env_level:
        return getattr(logging, env_level, logging.INFO)
    return level if isinstance(level, int) else logging.INFO


def _configure_root(level: int) -> None:
    """Вешает файловый и консольный хендлеры на корневой логгер (один раз)."""
    log_file = (os.getenv("LOG_FILE") or "").strip() or str(_PROJECT_ROOT / "logs.log")
    try:
        max_bytes = int((os.getenv("LOG_MAX_BYTES") or "").strip() or 5 * 1024 * 1024)
    except ValueError:
        max_bytes = 5 * 1024 * 1024

    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    root = logging.getLogger()
    root.setLevel(level)

    # mode="a": история прошлых запусков не теряется, ротация не даёт файлу расти.
    # encoding="utf-8": в логах кириллица (названия материалов, покрытий).
    try:
        file_handler = RotatingFileHandler(
            log_file, mode="a", maxBytes=max_bytes, backupCount=3, encoding="utf-8"
        )
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        root.addHandler(file_handler)
    except OSError as e:  # каталог только для чтения — консоли достаточно
        logging.getLogger(__name__).warning("Файл лога %s недоступен: %s", log_file, e)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    root.addHandler(console_handler)


def setup_logger(name: str = "logs", level=logging.INFO) -> logging.Logger:
    """Возвращает логгер приложения, настроив логирование при первом вызове.

    Args:
        name: имя логгера.
        level: уровень; переопределяется переменной окружения LOG_LEVEL.
    Returns:
        logging.Logger
    """
    global _configured
    resolved = _resolve_level(level)

    if not _configured:
        _configure_root(resolved)
        _configured = True

    logger = logging.getLogger(name)
    logger.setLevel(resolved)
    return logger


# Example usage
if __name__ == "__main__":
    logger = setup_logger("my_logger")
    logger.info("This is an info message.")
    logger.error("This is an error message.")
