"""
Модуль для интеграции с Битрикс24 REST API.
Отправка данных о печатных платах в смарт-процесс Битрикс24.
"""
import os
import re
import httpx
import logging
from types import SimpleNamespace
from typing import Dict, Any, Optional
try:
    from logger import setup_logger
    import bitrix24_dictionaries as dicts
    from pcb_normalizer import normalize_and_get_ids
    from utils import create_mistral_http_client
    from bitrix24_api import get_bitrix24_api
except ImportError:
    from .logger import setup_logger
    from . import bitrix24_dictionaries as dicts
    from .pcb_normalizer import normalize_and_get_ids
    from .utils import create_mistral_http_client
    from .bitrix24_api import get_bitrix24_api

logger = setup_logger(level=logging.INFO)

# Базовый URL для REST API Битрикс24
BITRIX24_BASE_URL = "https://fineline.bitrix24.ru/rest/6"
ENTITY_TYPE_ID = 182  # ID смарт-процесса PCB


class _MistralChatAdapter:
    """
    Адаптер, чтобы `pcb_normalizer.normalize_pcb_data()` мог вызывать `mistral_client.chat(...)`.

    В этом проекте первый LLM-запрос делается через LangChain `ChatMistralAI.invoke(...)`,
    а у него нет публичного `.chat(...)`. Этот класс оборачивает LangChain вызов и возвращает
    объект в формате `response.choices[0].message.content`.
    """

    def __init__(self, mistral_client: Any):
        self._mistral_client = mistral_client
        self._api_key = (os.getenv("MISTRAL_API_KEY") or "").strip()
        # Если вдруг передали "сырой" ChatMistralAI — достанем ключ от него.
        if not self._api_key:
            self._api_key = getattr(mistral_client, "mistral_api_key", None)

    def chat(
        self,
        model: str,
        messages: list[dict[str, str]],
        response_format: Optional[dict] = None,
        temperature: float = 0.0,
    ) -> Any:
        if not self._api_key:
            raise ValueError("MISTRAL_API_KEY не задан; невозможно выполнить нормализацию PCB.")

        from langchain_mistralai import ChatMistralAI

        # LangChain ожидает формат сообщений как list[tuple[role, content]]
        # роль "user" заменяем на "human".
        lc_messages: list[tuple[str, str]] = []
        for m in messages:
            role = (m.get("role") or "").strip().lower()
            content = m.get("content") or ""
            lc_role = "system" if role == "system" else "human"
            lc_messages.append((lc_role, content))

        # Клиент создаётся той же функцией, что и для Промпта 1: настройки сети
        # (proxy, base_url, таймауты) обязаны совпадать, иначе нормализация виснет
        # в ConnectTimeout там, где Mistral доступен только через proxy/VPN.
        http_client = create_mistral_http_client(self._api_key, read_timeout=60.0)
        # response_format передаётся через model_kwargs, а не в invoke(),
        # т.к. LangChain ChatMistralAI не принимает его как kwarg вызова.
        model_kwargs = {}
        if response_format:
            model_kwargs["response_format"] = response_format
        llm = ChatMistralAI(
            model=model,
            temperature=temperature,
            api_key=self._api_key,
            client=http_client,
            model_kwargs=model_kwargs,
            # Нормализация — не критичный шаг: при сетевой ошибке есть fallback на
            # статические справочники. Ограничиваем ретраи LangChain, чтобы
            # пользователь не ждал минуты вместо быстрого перехода к fallback.
            max_retries=1,
        )
        try:
            ai_msg = llm.invoke(lc_messages)
        finally:
            http_client.close()
        content = getattr(ai_msg, "content", None) or str(ai_msg)

        # Снимаем markdown-обёртку (```json ... ```) если модель её добавила
        content = content.strip()
        if content.startswith("```"):
            content = re.sub(r"^```(?:json)?\s*\n?", "", content)
            content = re.sub(r"\n?```\s*$", "", content)

        return SimpleNamespace(
            choices=[SimpleNamespace(message=SimpleNamespace(content=content))]
        )


# ═══════════════════════════════════════════════════════════════════
# ПАРСИНГ ГЕОМЕТРИИ
# ═══════════════════════════════════════════════════════════════════

# Допуски, которые нельзя принимать за размер:
#   h12 / H7 / js13 / k6 — квалитеты ISO (ГОСТ 25347)
#   ±0,2  +0.3  -0,3  ±10%  — числовые допуски
_ISO_TOLERANCE_RE = re.compile(r"\b(?:h|H|js|JS|k|K|g|G|f|F|e|E|d|D|m|M|n|N)\d{1,2}\b")
_NUM_TOLERANCE_RE = re.compile(r"[±+\-−]\s*\d+(?:[.,]\d+)?\s*%?")
_PERCENT_RE = re.compile(r"\d+(?:[.,]\d+)?\s*%")
_DIM_PAIR_RE = re.compile(r"(\d+(?:\.\d+)?)\s*[x×хХ*]\s*(\d+(?:\.\d+)?)")


def _strip_tolerances(text: str, drop_numeric: bool = True) -> str:
    """Убирает допуски из строки размера, оставляя только сами габариты.

    'длина 114 Допуск h12, ширина 47 h12' → 'длина 114 Допуск , ширина 47'
    Без этого re.findall принимал '12' из 'h12' за ширину платы.

    drop_numeric=False оставляет ±/+/- допуски: нужно для запасного разбора
    строк вида '114-47', где дефис — разделитель размеров, а не знак допуска.
    """
    s = (text or "").replace(",", ".")
    s = _ISO_TOLERANCE_RE.sub(" ", s)
    s = _PERCENT_RE.sub(" ", s)
    if drop_numeric:
        s = _NUM_TOLERANCE_RE.sub(" ", s)
    return s


def _dimensions_from_text(cleaned: str, min_value: float, max_value: float):
    """Пара габаритов из уже очищенной строки: сначала 'A x B', затем первые два числа."""
    match = _DIM_PAIR_RE.search(cleaned)
    if match:
        a, b = float(match.group(1)), float(match.group(2))
        if min_value <= a <= max_value and min_value <= b <= max_value:
            return a, b

    nums = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", cleaned)]
    valid = [n for n in nums if min_value <= n <= max_value]
    if len(valid) >= 2:
        return valid[0], valid[1]
    return None


def _parse_dimensions(text: Optional[str], min_value: float, max_value: float = 2000.0):
    """Извлекает пару габаритов (длина, ширина) из текста.

    Сначала ищет явную пару 'A x B' (самый надёжный признак), затем — первые два
    правдоподобных числа. Допуски отбрасываются до разбора чисел.
    Возвращает (length, width) или None.
    """
    if not text or not str(text).strip():
        return None
    raw = str(text)

    dims = _dimensions_from_text(_strip_tolerances(raw), min_value, max_value)
    if dims:
        return dims

    # Запасной проход: в '114-47' дефис разделяет размеры, а не помечает допуск,
    # и обычная чистка съедает второе число вместе со знаком.
    return _dimensions_from_text(
        _strip_tolerances(raw, drop_numeric=False).replace("-", " "),
        min_value, max_value,
    )


def _parse_positive_int(value: Any) -> Optional[int]:
    """Первое целое положительное число из значения ('2 шт' → 2)."""
    if value is None:
        return None
    numbers = re.findall(r"\d+", str(value))
    if not numbers:
        return None
    n = int(numbers[0])
    return n if n > 0 else None


def _is_yes(value: Any) -> bool:
    """Признак утвердительного ответа в русских/английских бланках."""
    v = str(value or "").strip().lower()
    if not v:
        return False
    positives = ("yes", "да", "есть", "требуется", "true", "1", "+", "нужен", "нужно", "имеется")
    negatives = ("no", "нет", "не требуется", "отсутствует", "false", "0", "-", "—", "n/a")
    for neg in negatives:
        if v == neg or v.startswith(neg + " "):
            return False
    return any(p in v for p in positives)


def _is_panel_delivery(
    pcb_data: Dict[str, Any],
    panel_dims: Optional[tuple],
    boards_per_panel: Optional[int],
) -> bool:
    """Плата поставляется панелями (тогда Order unit = pnl, а не ea).

    Признаки: несколько плат в заготовке, наличие технологических полей,
    заданный размер панели или прямое указание в тексте.
    """
    if boards_per_panel and boards_per_panel > 1:
        return True
    if _is_yes(pcb_data.get("technological_fields")):
        return True
    if panel_dims:
        return True
    text = " ".join(
        str(pcb_data.get(k) or "") for k in ("panelization", "order_unit", "contour_treatment")
    ).lower()
    if any(m in text for m in ("панел", "заготовк", "pnl", "panel", "мультиз")):
        # «Панелизация: Нет» — это отсутствие панели, а не поставка панелями.
        if not any(m in text for m in ("нет", "без панел", "поштучно", "no panel")):
            return True
    return False


# ═══════════════════════════════════════════════════════════════════
# ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ БИТРИКС24
# Этих полей нет в документации «Поля PCB». Код поля берётся из .env,
# а при заданном webhook определяется автоматически по названию на портале.
# Если код найти не удалось, характеристика всё равно распознаётся и видна
# в таблице результатов, но в CRM не отправляется (пишется в лог).
# ═══════════════════════════════════════════════════════════════════

_OPTIONAL_FIELD_ENV = {
    "no_of_diff_boards": "B24_FIELD_NO_OF_DIFF_BOARDS",
    "impedance_control": "B24_FIELD_IMPEDANCE_CONTROL",
    "min_hole_size":     "B24_FIELD_MIN_HOLE_SIZE",
    "marking_side":      "B24_FIELD_MARKING_SIDE",
    "serial_number":     "B24_FIELD_SERIAL_NUMBER",
}

# Ключевые слова для автопоиска поля по названию в Битрикс24.
# exclude отсекает похожие, но другие поля (напр. «Holes per Board»
# не должно подхватиться как «минимальный диаметр отверстия»).
_OPTIONAL_FIELD_KEYWORDS: Dict[str, Dict[str, list]] = {
    "no_of_diff_boards": {
        "keywords": ["diff board", "different board", "no. of diff", "no of diff",
                     "тип плат", "типов плат", "разных плат"],
        "exclude": [],
    },
    "min_hole_size": {
        "keywords": ["min hole", "minimum hole", "smallest hole", "min. hole",
                     "min drill", "minimum drill", "hole size", "hole diameter",
                     "мин. отверст", "минимальный диаметр", "диаметр отверст"],
        "exclude": ["per board", "density", "aspect", "count", "back drill", "плотност"],
    },
    "impedance_control": {
        "keywords": ["imped", "импеданс", "волнов", "controlled imp"],
        "exclude": [],
    },
    "marking_side": {
        "keywords": ["marking side", "legend side", "silkscreen side", "side of marking",
                     "сторона маркировки", "маркировка сторона"],
        "exclude": ["colour", "color", "цвет"],
    },
    "serial_number": {
        "keywords": ["serial number", "serial no", "серийный номер", "порядковый номер",
                     "unique code", "barcode", "штрих-код", "штрихкод"],
        "exclude": [],
    },
}

OPTIONAL_FIELD_CODES: Dict[str, str] = {
    key: (os.getenv(env_name) or "").strip()
    for key, env_name in _OPTIONAL_FIELD_ENV.items()
}

# Результат автопоиска кэшируется на процесс: {ключ: код поля или ""}.
_autodetected_codes: Dict[str, str] = {}


def resolve_optional_field_code(key: str) -> str:
    """Код поля Битрикс24 для дополнительной характеристики.

    Приоритет: явное значение из .env → автопоиск по названию через REST API.
    Автопоиск срабатывает, только если кандидат единственный, иначе Bitrix24Api
    пишет кандидатов в лог и возвращает None — чтобы не записать значение
    в чужое поле.
    """
    explicit = OPTIONAL_FIELD_CODES.get(key) or ""
    if explicit:
        return explicit
    if key in _autodetected_codes:
        return _autodetected_codes[key]

    code = ""
    api = get_bitrix24_api()
    spec = _OPTIONAL_FIELD_KEYWORDS.get(key)
    if api and spec:
        try:
            meta = api.find_field(spec["keywords"], exclude=spec["exclude"])
        except Exception as e:
            # Портал временно недоступен — результат не кэшируем,
            # чтобы следующая попытка снова обратилась к API.
            logger.warning("Автопоиск поля '%s' не выполнен: %s", key, e)
            return ""
        if meta:
            code = meta.code
            logger.info(
                "Поле '%s' найдено в Битрикс24 автоматически: %s («%s», тип %s)",
                key, meta.code, meta.title, meta.type,
            )
    _autodetected_codes[key] = code
    return code


# Значение не удалось сопоставить со справочником поля — поле не отправляем.
_UNMAPPABLE = object()

# Синонимы канонических значений: справочник на портале может быть на русском
# ('Да'/'Нет'), а мы формируем 'Yes'/'No'. Без этого значение не находится,
# строка уходит в списочное поле, и Битрикс24 отклоняет ВЕСЬ crm.item.add.
_VALUE_SYNONYMS: Dict[str, tuple] = {
    "yes":        ("Yes", "Да", "да", "есть", "Есть", "Y", "True", "1", "+"),
    "no":         ("No", "Нет", "нет", "отсутствует", "N", "False", "0", "-"),
    "top":        ("TOP", "Top", "Верх", "верх", "сверху", "Сверху", "Top side"),
    "bottom":     ("BOTTOM", "Bottom", "Bot", "Низ", "низ", "снизу", "Снизу", "Bottom side"),
    "top+bottom": ("TOP+BOTTOM", "TOP/BOTTOM", "Both", "обе стороны", "Обе стороны",
                   "две стороны", "с двух сторон", "Top+Bot"),
    "none":       ("None", "Нет", "нет", "отсутствует", "No", "—"),
}


def _find_list_item_id(value: str, allowed: Dict[str, int]) -> Optional[int]:
    """Ищет item_id значения в справочнике поля, перебирая синонимы."""
    item_id = dicts.find_item_id(value, allowed)
    if item_id is not None:
        return item_id
    for variants in _VALUE_SYNONYMS.values():
        if not any(value.strip().casefold() == v.casefold() for v in variants):
            continue
        for variant in variants:
            item_id = dicts.find_item_id(variant, allowed)
            if item_id is not None:
                logger.debug("Значение '%s' сопоставлено как '%s'", value, variant)
                return item_id
    return None


def _coerce_for_field(code: str, value: Any) -> Any:
    """Приводит значение к тому, что ожидает поле.

    Для списочного поля (iblock_element / enumeration) Битрикс24 принимает
    item_id, а не текст. Если сопоставить не удалось, возвращается _UNMAPPABLE:
    поле пропускается. Отправить в списочное поле строку нельзя — портал
    отклонит весь запрос, и заявка не создастся вообще.
    """
    api = get_bitrix24_api()
    if not api or not isinstance(value, str):
        return value
    try:
        allowed = api.get_field_values(code)
    except Exception as e:
        logger.debug("Значения поля %s не получены: %s", code, e)
        return value
    if not allowed:
        return value  # поле не списочное — строка допустима

    item_id = _find_list_item_id(value, allowed)
    if item_id is not None:
        logger.debug("Поле %s: значение '%s' → item_id %s", code, value, item_id)
        return item_id
    logger.warning(
        "Поле %s: значение '%s' не найдено среди допустимых (%s) — поле пропущено, "
        "чтобы Битрикс24 не отклонил заявку целиком.",
        code, value, ", ".join(list(allowed)[:8]),
    )
    return _UNMAPPABLE


def _apply_optional_fields(fields: Dict[str, Any], pcb_data: Dict[str, Any]) -> None:
    """Заполняет дополнительные характеристики.

    Код поля берётся из .env либо определяется автоматически через REST API.
    Текстовые значения для списочных полей конвертируются в item_id.
    """
    prepared: Dict[str, Any] = {}

    raw_hole = pcb_data.get("min_hole_size")
    if raw_hole:
        nums = re.findall(r"\d+(?:[.,]\d+)?", str(raw_hole).replace(",", "."))
        if nums:
            hole = float(nums[0])
            if 0 < hole <= 20:
                prepared["min_hole_size"] = hole

    if str(pcb_data.get("impedance_control") or "").strip():
        prepared["impedance_control"] = "Yes" if _is_yes(pcb_data["impedance_control"]) else "No"

    if str(pcb_data.get("serial_number") or "").strip():
        prepared["serial_number"] = "Yes" if _is_yes(pcb_data["serial_number"]) else "No"

    side = str(pcb_data.get("marking_side") or "").strip()
    if side:
        low = side.lower()
        has_top = "top" in low or "верх" in low or "сверху" in low
        has_bot = "bot" in low or "низ" in low or "снизу" in low
        if "две стороны" in low or "2 сторон" in low or "обе" in low or (has_top and has_bot):
            prepared["marking_side"] = "TOP+BOTTOM"
        elif has_top:
            prepared["marking_side"] = "TOP"
        elif has_bot:
            prepared["marking_side"] = "BOTTOM"
        elif low in ("нет", "none", "отсутствует", "no"):
            prepared["marking_side"] = "None"

    for key, value in prepared.items():
        code = resolve_optional_field_code(key)
        if code:
            coerced = _coerce_for_field(code, value)
            if coerced is not _UNMAPPABLE:
                fields[code] = coerced
        else:
            logger.info(
                "Характеристика '%s' = %r распознана, но не отправлена: поле не найдено "
                "в Битрикс24 автоматически — задайте код в .env (%s).",
                key, value, _OPTIONAL_FIELD_ENV[key],
            )


def _mask_webhook_url(url: str) -> str:
    """Маскирует токен в webhook URL: он секретный и не должен попадать в логи."""
    return re.sub(r"(/rest/\d+/)[^/]+", r"\1***", url)


def create_bitrix24_item(
    webhook_url_or_token: str,
    fields: Dict[str, Any],
    entity_type_id: int = ENTITY_TYPE_ID
) -> Dict[str, Any]:
    """
    Создает элемент в смарт-процессе Битрикс24.
    
    Args:
        webhook_url_or_token: Webhook URL (полный) или токен для REST API Битрикс24
            - Webhook URL: https://fineline.bitrix24.ru/rest/6/<token>/crm.item.add
            - Токен: просто токен, будет использован для построения URL
        fields: Словарь с полями элемента (UF_CRM_24_*)
        entity_type_id: ID типа сущности (по умолчанию 182 для PCB)
    
    Returns:
        Dict с результатом создания элемента
    
    Raises:
        httpx.HTTPStatusError: При ошибке HTTP запроса
        ValueError: При отсутствии обязательных полей
    """
    if not webhook_url_or_token:
        raise ValueError(
            "Webhook URL или токен Битрикс24 не задан. "
            "Установите переменную окружения BITRIX24_WEBHOOK_URL или BITRIX24_TOKEN"
        )
    
    # Определяем, это полный URL или токен
    webhook_url_or_token = webhook_url_or_token.strip()
    if webhook_url_or_token.startswith("http://") or webhook_url_or_token.startswith("https://"):
        # Это полный webhook URL
        url = webhook_url_or_token
        if not url.endswith("/crm.item.add"):
            # Если передан базовый URL без метода, добавляем метод
            url = url.rstrip("/") + "/crm.item.add"
    else:
        # Это токен, строим URL
        url = f"{BITRIX24_BASE_URL}/{webhook_url_or_token}/crm.item.add"
    
    payload = {
        "entityTypeId": entity_type_id,
        "fields": fields
    }
    
    logger.info("Отправка данных в Битрикс24: %d полей", len(fields))
    logger.debug("URL: %s", _mask_webhook_url(url))
    logger.debug("Payload: %s", payload)

    # Битрикс24 — российский сервис: при включённом VPN/proxy (нужном для Mistral)
    # запрос к нему по умолчанию идёт НАПРЯМУЮ, мимо proxy, а при сетевой ошибке
    # повторяется через proxy. Принудительный режим — BITRIX24_TRUST_ENV=0/1.
    forced_trust = (os.getenv("BITRIX24_TRUST_ENV") or "").strip()
    trust_modes = [forced_trust == "1"] if forced_trust else [False, True]

    last_network_error: Optional[Exception] = None
    for trust_env in trust_modes:
        try:
            with httpx.Client(timeout=30.0, trust_env=trust_env) as client:
                response = client.post(
                    url,
                    json=payload,
                    headers={"Content-Type": "application/json"},
                )
                response.raise_for_status()
                result = response.json()

            if "error" in result:
                error_msg = result.get("error_description", result.get("error", "Unknown error"))
                logger.error("Ошибка Битрикс24 API: %s", error_msg)
                raise Exception(f"Ошибка Битрикс24: {error_msg}")

            item_id = result.get("result", {}).get("item", {}).get("id")
            logger.info("Успешно создан элемент в Битрикс24 с ID: %s", item_id)
            return result

        except httpx.HTTPStatusError as e:
            # Сервер ответил — сеть работает, повтор в другом режиме бессмыслен.
            logger.error("HTTP ошибка при отправке в Битрикс24: %s", e)
            try:
                error_response = e.response.json()
                error_detail = error_response.get(
                    "error_description", error_response.get("error", "")
                )
            except Exception:
                error_detail = str(e)
            raise Exception(f"Ошибка подключения к Битрикс24: {error_detail}")
        except httpx.RequestError as e:
            last_network_error = e
            logger.warning(
                "Битрикс24 недоступен (proxy=%s): %s", trust_env, e
            )
            continue

    logger.error("Ошибка запроса к Битрикс24: %s", last_network_error)
    raise Exception(f"Не удалось подключиться к Битрикс24: {last_network_error}")


def map_pcb_to_bitrix24_fields(pcb_data: Dict[str, Any], mistral_client: Any = None) -> Dict[str, Any]:
    """
    Преобразует данные PCB в формат полей Битрикс24.
    
    Маппинг полей с использованием справочников для iblock_element полей:
    - board_name -> UF_CRM_24_1709799376061 (OEM PN)
    - base_material -> UF_CRM_24_1707838248 (Materials) - через справочник 56
    - layer_count -> UF_CRM_24_1709815185 (No of Layers) - через справочник 54
    - coverage_type -> UF_CRM_24_1707768819 (Finish Type) - через справочник 74
    - foil_thickness -> UF_CRM_24_1707838441 (Max Copper) - через справочник 62
    - board_size -> парсится в Board Length/Width если возможно
    - panelization -> парсится в Panel Length/Width если возможно
    
    Args:
        pcb_data: Словарь с данными PCB (из PCBCharacteristics.model_dump())
    
    Returns:
        Словарь с полями для Битрикс24 (UF_CRM_24_*)
    """
    fields = {}
    
    # ========== ОБЯЗАТЕЛЬНЫЕ СТРОКОВЫЕ ПОЛЯ ==========
    
    # OEM PN (обязательное) - без него не отправляем данные
    board_name = (pcb_data.get("board_name") or "").strip()
    if not board_name:
        raise ValueError("Не задано обязательное поле 'board_name' (OEM PN) для заявки в Битрикс24.")
    fields["ufCrm24_1709799376061"] = board_name
    
    # OEM Description (обязательное) - дублируем OEM PN
    fields["ufCrm24_1709799393816"] = board_name
    
    # Rev. (обязательное)
    fields["ufCrm24_1709799420584"] = "."
    
    # Board Thickness (обязательное double) - без значения не отправляем данные
    # Берём из pcb_data["board_thickness"] (Finished thickness with tolerance, mm)
    board_thickness: Optional[float] = None
    thickness_src = (pcb_data.get("board_thickness") or "").strip()
    if thickness_src:
        # Заменяем запятую на точку и вытаскиваем первое число
        cleaned = thickness_src.replace(",", ".")
        # Берём первое число целое или с точкой, без захватывающих групп,
        # чтобы re.findall возвращал полное совпадение, а не только дробную часть.
        numbers = re.findall(r"\d+(?:\.\d+)?", cleaned)
        if numbers:
            try:
                value = float(numbers[0])
                if 0.1 <= value <= 10:
                    board_thickness = value
            except ValueError:
                board_thickness = None

    if board_thickness is None:
        raise ValueError(
            "Не удалось корректно определить обязательное поле 'board_thickness' "
            "(толщина платы). Проверьте исходный документ."
        )
    fields["ufCrm24_1708374728464"] = board_thickness
    
    # ========== ОБЯЗАТЕЛЬНЫЕ ПОЛЯ ТИПА IBLOCK_ELEMENT ==========

    normalization_used = False

    # ── НОВЫЙ ВАРИАНТ: нормализация через pcb_normalizer (LLM promt 2) ─────────
    if mistral_client:
        try:
            normalizer_client = mistral_client
            if not hasattr(mistral_client, "chat"):
                normalizer_client = _MistralChatAdapter(mistral_client)

            _, b24_ids = normalize_and_get_ids(pcb_data, normalizer_client)

            fields.update(b24_ids)
            normalization_used = True
        except Exception as e:
            logger.warning("Нормализация PCB не удалась, fallback на dicts.*: %s", e)
            normalization_used = False

    # ── Старый вариант: dicts.* (fallback) ────────────────────────────────
    if not normalization_used:
        # UF_CRM_24_1707838248: Materials (справочник 56)
        # Если материал не найден в справочнике — назначаем MIX/Others (ID 5646).
        if pcb_data.get("base_material"):
            material_id = dicts.get_material_id(pcb_data["base_material"]) or 5646
            fields["ufCrm24_1707838248"] = material_id
            if material_id == 5646:
                logger.warning(
                    "Материал '%s' не найден в справочнике — назначен MIX/Others (5646)",
                    pcb_data["base_material"],
                )
            else:
                logger.debug("Materials: '%s' -> %s", pcb_data["base_material"], material_id)

        # UF_CRM_24_1707768819: Finish Type (справочник 74)
        if pcb_data.get("coverage_type"):
            finish_id = dicts.get_finish_type_id(pcb_data["coverage_type"])
            if finish_id:
                fields["ufCrm24_1707768819"] = finish_id
                logger.debug("Finish Type: '%s' -> %s", pcb_data["coverage_type"], finish_id)
            else:
                logger.warning("Не найден ID для типа покрытия: '%s'", pcb_data["coverage_type"])

        # UF_CRM_24_1707838441: Max Copper (base OZ) (справочник 62)
        if pcb_data.get("foil_thickness"):
            copper_id = dicts.get_copper_thickness_id(pcb_data["foil_thickness"])
            if copper_id:
                fields["ufCrm24_1707838441"] = copper_id
                logger.debug("Copper thickness: '%s' -> %s", pcb_data["foil_thickness"], copper_id)
            else:
                logger.warning("Не найден ID для толщины меди: '%s'", pcb_data["foil_thickness"])

    # UF_CRM_24_1709815185: No of Layers — маппим только если pcb_normalizer не справился
    # (при normalization_used=True это поле уже есть в b24_ids → fields).
    if not normalization_used and pcb_data.get("layer_count"):
        layers_id = dicts.get_layers_id(str(pcb_data["layer_count"]))
        if layers_id:
            fields["ufCrm24_1709815185"] = layers_id
            logger.debug("Layers (fallback): '%s' -> %s", pcb_data["layer_count"], layers_id)
        else:
            logger.warning("Не найден ID для количества слоев: '%s'", pcb_data["layer_count"])

    # ========== ГЕОМЕТРИЧЕСКИЕ ПАРАМЕТРЫ (double) ==========

    # Размеры платы (Board Length / Board Width).
    board_dims = _parse_dimensions(pcb_data.get("board_size"), min_value=5.0)
    if board_dims:
        bl, bw = board_dims
        fields["ufCrm24_1708353384301"] = bl                          # Board Length (mm)
        fields["ufCrm24_1708353402068"] = bw                          # Board Width (mm)
        fields["ufCrm24_1708374692747"] = round(bl * bw / 100, 4)     # Board Size (sqr dec)
    elif pcb_data.get("board_size"):
        logger.warning("Не удалось определить размеры платы из '%s'", pcb_data.get("board_size"))

    # Размеры панели: сначала выделенное поле panel_size, затем общий panelization.
    # Порог 20 мм отсекает счётчики рядов/столбцов и зазоры.
    panel_dims = _parse_dimensions(pcb_data.get("panel_size"), min_value=20.0)
    if not panel_dims:
        panel_dims = _parse_dimensions(pcb_data.get("panelization"), min_value=20.0)
    if panel_dims:
        pl, pw = panel_dims
        fields["ufCrm24_1708375852081"] = pl                          # Panel Length (mm)
        fields["ufCrm24_1708375871512"] = pw                          # Panel Width (mm)
        fields["ufCrm24_1708375895460"] = round(pl * pw / 100, 4)     # Panel Size (sqr dec)

    # Количество плат на панели и количество РАЗНЫХ типов плат.
    boards_per_panel = _parse_positive_int(pcb_data.get("boards_per_panel"))
    if boards_per_panel:
        fields["ufCrm24_1708375915545"] = boards_per_panel            # Board Per Panel
        # No. Of Diff Boards заполняется вместе с Board Per Panel.
        # По умолчанию 1 — панель из повторяющейся одной платы.
        diff_boards = _parse_positive_int(pcb_data.get("different_boards_per_panel")) or 1
        diff_field = resolve_optional_field_code("no_of_diff_boards")
        if diff_field:
            fields[diff_field] = diff_boards
        else:
            logger.info(
                "No. Of Diff Boards = %s не отправлено: поле не найдено в Битрикс24 — "
                "задайте код в .env (B24_FIELD_NO_OF_DIFF_BOARDS).", diff_boards
            )
        # Использование площади панели, % — считаем, если известны обе площади.
        panel_area = fields.get("ufCrm24_1708375895460")
        board_area = fields.get("ufCrm24_1708374692747")
        if panel_area and board_area:
            usage = round(board_area * boards_per_panel / panel_area * 100, 2)
            if 0 < usage <= 100:
                fields["ufCrm24_1708375925847"] = usage                # Panel Usage (%)

    # ========== ЕДИНИЦА ЗАКАЗА: pnl при поставке в панелях ==========
    # Если плата поставляется в панели (несколько плат в заготовке, есть
    # технологические поля или задан размер панели) — заказ идёт панелями.
    if _is_panel_delivery(pcb_data, panel_dims, boards_per_panel):
        fields["ufCrm24_1707838030"] = dicts.ORDER_UNIT_DICT["pnl"]        # Order unit  = pnl
        fields["ufCrm24_1707849863"] = dicts.PRODUCTION_UNIT_DICT["pnl"]   # Production  = pnl
        logger.info("Поставка в панелях — Order unit / Production Unit = pnl")

    # ========== НОВЫЕ ХАРАКТЕРИСТИКИ (коды полей задаются в .env) ==========
    _apply_optional_fields(fields, pcb_data)

    # ========== ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ ==========

    # UF_CRM_24_1707839110: Edge plating — маппим только если pcb_normalizer не справился.
    # При normalization_used=True это поле уже есть в b24_ids → fields.
    # Solder Mask Color и Silkscreen Color: укажите реальные коды полей из Битрикс24
    # (ufCrm24_XXXXXXXXXX), чтобы активировать эти поля.
    if not normalization_used and pcb_data.get("edge_plating"):
        plating_id = dicts.get_edge_plating_id(pcb_data["edge_plating"])
        if plating_id:
            fields["ufCrm24_1707839110"] = plating_id
            logger.debug("Edge plating (fallback): '%s' -> %s", pcb_data["edge_plating"], plating_id)

    # Order unit / Production Unit обязательны: если ни нормализация, ни логика
    # панелей их не проставили (fallback-ветка) — ставим значение по умолчанию.
    fields.setdefault("ufCrm24_1707838030", dicts.ORDER_UNIT_DICT["ea"])
    fields.setdefault("ufCrm24_1707849863", dicts.PRODUCTION_UNIT_DICT["ea"])

    logger.info("Создано %d полей для Битрикс24", len(fields))
    logger.debug("Поля: %s", list(fields.keys()))
    return fields


def send_pcb_to_bitrix24(
    pcb_data: Dict[str, Any],
    webhook_url_or_token: str,
    entity_type_id: int = ENTITY_TYPE_ID,
    mistral_client: Any = None,
) -> Dict[str, Any]:
    """
    Отправляет данные PCB в Битрикс24.
    
    Args:
        pcb_data: Словарь с данными PCB
        webhook_url_or_token: Webhook URL (полный) или токен авторизации Битрикс24
        entity_type_id: ID типа сущности
        mistral_client: Mistral-клиент для нормализации (LLM prompt 2). Если None — используется fallback dicts.*
    
    Returns:
        Результат создания элемента в Битрикс24
    """
    fields = map_pcb_to_bitrix24_fields(pcb_data, mistral_client=mistral_client)
    return create_bitrix24_item(webhook_url_or_token, fields, entity_type_id)
