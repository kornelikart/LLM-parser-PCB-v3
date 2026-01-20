"""
Модуль для интеграции с Битрикс24 REST API.
Отправка данных о печатных платах в смарт-процесс Битрикс24.
"""
import httpx
import logging
from typing import Dict, Any, Optional
try:
    from logger import setup_logger
    import bitrix24_dictionaries as dicts
except:
    from .logger import setup_logger
    from . import bitrix24_dictionaries as dicts

logger = setup_logger(level=logging.INFO)

# Базовый URL для REST API Битрикс24
BITRIX24_BASE_URL = "https://fineline.bitrix24.ru/rest/6"
ENTITY_TYPE_ID = 182  # ID смарт-процесса PCB


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
    
    logger.info(f"Отправка данных в Битрикс24: {len(fields)} полей")
    logger.debug(f"URL: {url}")
    logger.debug(f"Payload: {payload}")
    
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
                logger.error(f"Ошибка Битрикс24 API: {error_msg}")
                raise Exception(f"Ошибка Битрикс24: {error_msg}")
            
            item_id = result.get("result", {}).get("item", {}).get("id")
            logger.info(f"Успешно создан элемент в Битрикс24 с ID: {item_id}")
            return result
            
    except httpx.HTTPStatusError as e:
        logger.error(f"HTTP ошибка при отправке в Битрикс24: {e}")
        error_detail = ""
        try:
            error_response = e.response.json()
            error_detail = error_response.get("error_description", error_response.get("error", ""))
        except:
            error_detail = str(e)
        raise Exception(f"Ошибка подключения к Битрикс24: {error_detail}")
    except httpx.RequestError as e:
        logger.error(f"Ошибка запроса к Битрикс24: {e}")
        raise Exception(f"Не удалось подключиться к Битрикс24: {str(e)}")


def map_pcb_to_bitrix24_fields(pcb_data: Dict[str, Any]) -> Dict[str, Any]:
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
    
    # OEM PN (обязательное)
    if pcb_data.get("board_name"):
        fields["ufCrm24_1709799376061"] = pcb_data["board_name"]
    
    # OEM Description (обязательное) - комбинируем несколько полей
    description_parts = []
    if pcb_data.get("base_material"):
        description_parts.append(f"Material: {pcb_data['base_material']}")
    if pcb_data.get("layer_count"):
        description_parts.append(f"Layers: {pcb_data['layer_count']}")
    if pcb_data.get("coverage_type"):
        description_parts.append(f"Finish: {pcb_data['coverage_type']}")
    if description_parts:
        fields["ufCrm24_1709799393816"] = ", ".join(description_parts)
    else:
        # Обязательное поле, должно быть заполнено
        fields["ufCrm24_1709799393816"] = "PCB"
    
    # Rev. (обязательное)
    fields["ufCrm24_1709799420584"] = ""
    
    # Board Thickness (обязательное double) - по умолчанию 1.6
    board_thickness = 1.6
    if pcb_data.get("board_size"):
        try:
            parts = pcb_data["board_size"].replace("x", " ").replace("X", " ").split()
            for part in parts:
                if "." in part:
                    thickness = float(part)
                    if 0.1 <= thickness <= 10:
                        board_thickness = thickness
                        break
        except:
            pass
    fields["ufCrm24_1708374728464"] = board_thickness
    
    # ========== ОБЯЗАТЕЛЬНЫЕ ПОЛЯ ТИПА IBLOCK_ELEMENT ==========
    
    # UF_CRM_24_1707838248: Materials (справочник 56)
    if pcb_data.get("base_material"):
        material_id = dicts.get_material_id(pcb_data["base_material"])
        if material_id:
            fields["ufCrm24_1707838248"] = material_id
            logger.debug(f"Materials: '{pcb_data['base_material']}' -> {material_id}")
        else:
            logger.warning(f"Не найден ID для материала: '{pcb_data['base_material']}'")
    
    # UF_CRM_24_1709815185: No of Layers (справочник 54)
    if pcb_data.get("layer_count"):
        layers_id = dicts.get_layers_id(str(pcb_data["layer_count"]))
        if layers_id:
            fields["ufCrm24_1709815185"] = layers_id
            logger.debug(f"Layers: '{pcb_data['layer_count']}' -> {layers_id}")
        else:
            logger.warning(f"Не найден ID для количества слоев: '{pcb_data['layer_count']}'")
    
    # UF_CRM_24_1707768819: Finish Type (справочник 74)
    if pcb_data.get("coverage_type"):
        finish_id = dicts.get_finish_type_id(pcb_data["coverage_type"])
        if finish_id:
            fields["ufCrm24_1707768819"] = finish_id
            logger.debug(f"Finish Type: '{pcb_data['coverage_type']}' -> {finish_id}")
        else:
            logger.warning(f"Не найден ID для типа покрытия: '{pcb_data['coverage_type']}'")
    
    # UF_CRM_24_1707838441: Max Copper (base OZ) (справочник 62)
    if pcb_data.get("foil_thickness"):
        copper_id = dicts.get_copper_thickness_id(pcb_data["foil_thickness"])
        if copper_id:
            fields["ufCrm24_1707838441"] = copper_id
            logger.debug(f"Copper thickness: '{pcb_data['foil_thickness']}' -> {copper_id}")
        else:
            logger.warning(f"Не найден ID для толщины меди: '{pcb_data['foil_thickness']}'")
    
    # UF_CRM_24_1707838030: Order unit (справочник 50) - обязательное, но нет в PCB данных
    # Можно установить значение по умолчанию или попросить пользователя указать
    # Пока пропускаем, но нужно будет добавить
    
    # UF_CRM_24_1707838074: PCB type (справочник 52) - обязательное
    # Можно попытаться определить по другим параметрам или установить значение по умолчанию
    # Пока пропускаем, но нужно будет добавить
    
    # UF_CRM_24_1707839629: Peelable SM (справочник 86) - обязательное
    # Нет в PCB данных, нужно установить значение по умолчанию или пропустить
    # Пока пропускаем
    
    # UF_CRM_24_1707849863: Production Unit (справочник 160) - обязательное
    # Нет в PCB данных, нужно установить значение по умолчанию
    # Пока пропускаем
    
    # ========== ГЕОМЕТРИЧЕСКИЕ ПАРАМЕТРЫ (double) ==========
    
    # Парсинг размеров платы
    if pcb_data.get("board_size"):
        try:
            size_str = pcb_data["board_size"].replace("x", " ").replace("X", " ").replace("×", " ")
            parts = [float(p) for p in size_str.split() if p.replace(".", "").isdigit()]
            if len(parts) >= 2:
                fields["ufCrm24_1708353384301"] = parts[0]  # Board Length (mm)
                fields["ufCrm24_1708353402068"] = parts[1]  # Board Width (mm)
                if len(parts) >= 3:
                    if 0.1 <= parts[2] <= 10:
                        fields["ufCrm24_1708374728464"] = parts[2]  # Board Thickness
        except Exception as e:
            logger.debug(f"Не удалось распарсить размер платы: {e}")
    
    # Парсинг панелизации
    if pcb_data.get("panelization"):
        try:
            panel_str = pcb_data["panelization"].replace("x", " ").replace("X", " ").replace("×", " ")
            parts = [float(p) for p in panel_str.split() if p.replace(".", "").isdigit()]
            if len(parts) >= 2:
                fields["ufCrm24_1708375852081"] = parts[0]  # Panel Length (mm)
                fields["ufCrm24_1708375871512"] = parts[1]  # Panel Width (mm)
        except Exception as e:
            logger.debug(f"Не удалось распарсить панелизацию: {e}")
    
    # ========== ДОПОЛНИТЕЛЬНЫЕ ПОЛЯ ==========
    
    # UF_CRM_24_*: Solder Mask Color (справочник 64)
    if pcb_data.get("solder_mask_colour"):
        color_id = dicts.get_solder_mask_color_id(pcb_data["solder_mask_colour"])
        if color_id:
            # Нужно узнать точный код поля для цвета паяльной маски
            # Пока пропускаем, так как нет точного кода поля в документации
            pass
    
    # UF_CRM_24_1707839110: Edge plating (справочник 72)
    if pcb_data.get("edge_plating"):
        plating_id = dicts.get_edge_plating_id(pcb_data["edge_plating"])
        if plating_id:
            fields["ufCrm24_1707839110"] = plating_id
            logger.debug(f"Edge plating: '{pcb_data['edge_plating']}' -> {plating_id}")
    
    logger.info(f"Создано {len(fields)} полей для Битрикс24")
    logger.debug(f"Поля: {list(fields.keys())}")
    return fields


def send_pcb_to_bitrix24(
    pcb_data: Dict[str, Any],
    webhook_url_or_token: str,
    entity_type_id: int = ENTITY_TYPE_ID
) -> Dict[str, Any]:
    """
    Отправляет данные PCB в Битрикс24.
    
    Args:
        pcb_data: Словарь с данными PCB
        webhook_url_or_token: Webhook URL (полный) или токен авторизации Битрикс24
        entity_type_id: ID типа сущности
    
    Returns:
        Результат создания элемента в Битрикс24
    """
    fields = map_pcb_to_bitrix24_fields(pcb_data)
    return create_bitrix24_item(webhook_url_or_token, fields, entity_type_id)
