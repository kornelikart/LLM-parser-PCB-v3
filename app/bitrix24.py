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
except ImportError:
    from .logger import setup_logger
    from . import bitrix24_dictionaries as dicts
    from .pcb_normalizer import normalize_and_get_ids

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

        # Отключаем системные proxy-настройки (trust_env=True),
        # т.к. в некоторых окружениях они приводят к WinError 10061.
        # Важно: если передать client в ChatMistralAI, base_url не подставляется автоматически,
        # поэтому задаём base_url явно.
        base_url = (os.getenv("MISTRAL_BASE_URL") or "").strip() or "https://api.mistral.ai/v1"
        http_client = httpx.Client(
            base_url=base_url,
            timeout=30.0,
            trust_env=False,
            headers={
                "Authorization": f"Bearer {self._api_key}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
        )
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
    logger.debug("URL: %s", url)
    logger.debug("Payload: %s", payload)
    
    try:
        with httpx.Client(timeout=30.0) as client:
            response = client.post(
                url,
                json=payload,
                headers={"Content-Type": "application/json"}
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
        logger.error("HTTP ошибка при отправке в Битрикс24: %s", e)
        error_detail = ""
        try:
            error_response = e.response.json()
            error_detail = error_response.get("error_description", error_response.get("error", ""))
        except Exception:
            error_detail = str(e)
        raise Exception(f"Ошибка подключения к Битрикс24: {error_detail}")
    except httpx.RequestError as e:
        logger.error("Ошибка запроса к Битрикс24: %s", e)
        raise Exception(f"Не удалось подключиться к Битрикс24: {str(e)}")


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
    
    # Парсинг размеров платы (Board Length / Board Width) — извлекаем числа из "(253.0 x 140.0) ± 0.2 mm"
    if pcb_data.get("board_size"):
        try:
            size_str = (pcb_data["board_size"] or "").replace(",", ".")
            parts = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", size_str)]
            # Отсекаем допуски (< 5) и нереальные значения (> 2000 мм)
            valid = [p for p in parts if 5 <= p <= 2000]
            if len(valid) >= 2:
                bl, bw = valid[0], valid[1]
                fields["ufCrm24_1708353384301"] = bl                          # Board Length (mm)
                fields["ufCrm24_1708353402068"] = bw                          # Board Width (mm)
                fields["ufCrm24_1708374692747"] = round(bl * bw / 100, 4)    # Board Size (sqr dec)
        except Exception as e:
            logger.debug("Не удалось распарсить размер платы: %s", e)

    # Парсинг панелизации.
    # Строки вида "2 x 3 (200 x 150) ±0.3" содержат счётчики рядов/столбцов (2,3)
    # и реальные размеры панели (200, 150). Используем порог >= 20 мм, чтобы
    # счётчики (1–10) и допуски (0.x–1.x) не попали в размеры.
    if pcb_data.get("panelization"):
        try:
            panel_str = (pcb_data["panelization"] or "").replace(",", ".")
            parts = [float(n) for n in re.findall(r"\d+(?:\.\d+)?", panel_str)]
            valid = [p for p in parts if 20 <= p <= 2000]
            if len(valid) >= 2:
                pl, pw = valid[0], valid[1]
                fields["ufCrm24_1708375852081"] = pl                          # Panel Length (mm)
                fields["ufCrm24_1708375871512"] = pw                          # Panel Width (mm)
                fields["ufCrm24_1708375895460"] = round(pl * pw / 100, 4)    # Panel Size (sqr dec)
        except Exception as e:
            logger.debug("Не удалось распарсить панелизацию: %s", e)
    
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
