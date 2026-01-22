"""
Справочники для маппинга текстовых значений на ID элементов Битрикс24.

ВАЖНО: Для работы с реальными данными необходимо заполнить справочники реальными ID из Битрикс24.

Формат справочника в Битрикс24:
- param_id - ID справочника (IBLOCK_ID)
- param - название параметра
- item_id - ID элемента в справочнике (это значение используется в полях)
- item - название элемента (текстовое значение)

Как получить реальные ID:
1. Используйте REST API Битрикс24: GET /rest/{user_id}/{token}/lists.element.get?IBLOCK_TYPE_ID=lists&IBLOCK_ID={IBLOCK_ID}
2. Или экспортируйте справочники из Битрикс24 в CSV/Excel
3. Заполните словари ниже реальными значениями item_id

Структура словарей: {текстовое_значение: item_id}
"""
from typing import Dict, Optional
import os

def _try_get_db():
    """
    Lazy import to avoid hard dependency on SQLAlchemy when DB mode is not used.
    Returns a DbDictionaries instance or None.
    """
    if (os.getenv("USE_DB_DICTIONARIES", "1").strip() == "0"):
        return None
    if not (os.getenv("DICTIONARIES_DB_URL") or "").strip():
        return None
    try:
        from .db_dictionaries import get_db_dictionaries  # type: ignore
        return get_db_dictionaries()
    except Exception:
        # If running as script from app/ or SQLAlchemy is not installed yet
        try:
            from db_dictionaries import get_db_dictionaries  # type: ignore
            return get_db_dictionaries()
        except Exception:
            return None

# Справочник 56: Materials (Материал основания платы)
# IBLOCK_ID = 56
MATERIALS_DICT: Dict[str, int] = {
    "FR4": 6281,
    "FR-4": 6281,  # альтернативное написание
    # Добавьте другие материалы по мере необходимости
}

# Справочник 74: Finish Type (Тип финишного покрытия)
FINISH_TYPE_DICT: Dict[str, int] = {
    "HASL": 7010,
    "HASL Lead-Free": 7010,
    "ENIG": 5256,  # пример, нужно уточнить ID
    "OSP": 6610,  # пример, нужно уточнить ID
    # Добавьте другие типы покрытия
}

# Справочник 54: No of Layers (Количество слоев)
LAYERS_DICT: Dict[str, int] = {
    "1": 6014,  # пример ID для 1 слоя
    "2": 6014,  # пример ID для 2 слоев
    "4": 6014,  # пример ID для 4 слоев
    "6": 6014,  # пример ID для 6 слоев
    "8": 6014,  # пример ID для 8 слоев
    # Нужно уточнить реальные ID для каждого количества слоев
}

# Справочник 62: Max Copper (base OZ) - Толщина меди
COPPER_THICKNESS_DICT: Dict[str, int] = {
    "0.5 OZ": 6002,
    "0.5": 6002,
    "1 OZ": 6002,
    "1": 6002,
    "1.5 OZ": 6002,
    "1.5": 6002,
    "2 OZ": 6002,
    "2": 6002,
    # Нужно уточнить реальные ID для каждой толщины
}

# Справочник 50: Order unit (Единица заказа)
ORDER_UNIT_DICT: Dict[str, int] = {
    "шт": 5256,  # пример
    "piece": 5256,
    "pcs": 5256,
    # Нужно уточнить реальные ID
}

# Справочник 52: PCB type (Тип платы)
PCB_TYPE_DICT: Dict[str, int] = {
    "Rigid": 5804,  # пример
    "Flex": 5804,
    "Rigid-Flex": 5804,
    # Нужно уточнить реальные ID
}

# Справочник 86: Peelable SM (Пилинг-маска)
PEELABLE_SM_DICT: Dict[str, int] = {
    "Yes": 6270,  # пример
    "No": 6270,
    # Нужно уточнить реальные ID
}

# Справочник 160: Production Unit (Производственный участок)
PRODUCTION_UNIT_DICT: Dict[str, int] = {
    # Нужно заполнить значениями из вашего справочника
}

# Справочник 64: Solder Mask Color (Цвет паяльной маски)
SOLDER_MASK_COLOR_DICT: Dict[str, int] = {
    "Green": 8002,
    "Red": 8002,  # пример, нужно уточнить ID
    "Blue": 8002,
    "Black": 8002,
    "White": 8002,
}

# Справочник 66: SilkScreen Color (Цвет маркировки)
SILKSCREEN_COLOR_DICT: Dict[str, int] = {
    "Green": 7002,
    "White": 7002,  # пример, нужно уточнить ID
    "Black": 7002,
}

# Справочник 72: Edge plating (Металлизация края)
EDGE_PLATING_DICT: Dict[str, int] = {
    "Yes": 5860,  # пример
    "No": 5860,
    # Нужно уточнить реальные ID
}

def normalize_text(text: str) -> str:
    """
    Нормализует текст для поиска в справочниках.
    Убирает пробелы, приводит к нижнему регистру, удаляет спецсимволы.
    """
    if not text:
        return ""
    return text.strip().lower().replace("-", "").replace("_", "").replace(" ", "")


def find_item_id(
    text_value: str,
    dictionary: Dict[str, int],
    fuzzy_match: bool = True
) -> Optional[int]:
    """
    Находит ID элемента в справочнике по текстовому значению.
    
    Args:
        text_value: Текстовое значение для поиска
        dictionary: Словарь маппинга {текст: item_id}
        fuzzy_match: Если True, использует нечеткое сравнение
    
    Returns:
        item_id или None, если не найдено
    """
    if not text_value:
        return None
    
    # Точное совпадение (регистронезависимое)
    normalized_input = normalize_text(text_value)
    for key, item_id in dictionary.items():
        if normalize_text(key) == normalized_input:
            return item_id
    
    # Нечеткое совпадение (если включено)
    if fuzzy_match:
        normalized_input = text_value.lower().strip()
        for key, item_id in dictionary.items():
            key_lower = key.lower().strip()
            # Проверка на вхождение подстроки
            if key_lower in normalized_input or normalized_input in key_lower:
                return item_id
            # Проверка на частичное совпадение слов
            input_words = set(normalized_input.split())
            key_words = set(key_lower.split())
            if input_words & key_words:  # пересечение множеств
                return item_id
    
    return None


def get_material_id(material_text: str) -> Optional[int]:
    """Получить ID материала из справочника 56"""
    db = _try_get_db()
    if db:
        return db.find_item_id(56, material_text)
    return find_item_id(material_text, MATERIALS_DICT)


def get_finish_type_id(finish_text: str) -> Optional[int]:
    """Получить ID типа финишного покрытия из справочника 74"""
    db = _try_get_db()
    if db:
        return db.find_item_id(74, finish_text)
    return find_item_id(finish_text, FINISH_TYPE_DICT)


def get_layers_id(layers_text: str) -> Optional[int]:
    """Получить ID количества слоев из справочника 54"""
    db = _try_get_db()
    if db:
        return db.find_item_id(54, str(layers_text))
    # Пытаемся извлечь число из текста
    import re
    numbers = re.findall(r'\d+', str(layers_text))
    if numbers:
        layers_count = numbers[0]
        return find_item_id(layers_count, LAYERS_DICT)
    return find_item_id(layers_text, LAYERS_DICT)


def get_copper_thickness_id(thickness_text: str) -> Optional[int]:
    """Получить ID толщины меди из справочника 62"""
    db = _try_get_db()
    if db:
        return db.find_item_id(62, thickness_text)
    return find_item_id(thickness_text, COPPER_THICKNESS_DICT)


def get_order_unit_id(unit_text: str) -> Optional[int]:
    """Получить ID единицы заказа из справочника 50"""
    db = _try_get_db()
    if db:
        return db.find_item_id(50, unit_text)
    return find_item_id(unit_text, ORDER_UNIT_DICT)


def get_pcb_type_id(pcb_type_text: str) -> Optional[int]:
    """Получить ID типа платы из справочника 52"""
    db = _try_get_db()
    if db:
        return db.find_item_id(52, pcb_type_text)
    return find_item_id(pcb_type_text, PCB_TYPE_DICT)


def get_peelable_sm_id(peelable_text: str) -> Optional[int]:
    """Получить ID пилинг-маски из справочника 86"""
    db = _try_get_db()
    if db:
        return db.find_item_id(86, peelable_text)
    return find_item_id(peelable_text, PEELABLE_SM_DICT)


def get_production_unit_id(unit_text: str) -> Optional[int]:
    """Получить ID производственного участка из справочника 160"""
    db = _try_get_db()
    if db:
        return db.find_item_id(160, unit_text)
    return find_item_id(unit_text, PRODUCTION_UNIT_DICT)


def get_solder_mask_color_id(color_text: str) -> Optional[int]:
    """Получить ID цвета паяльной маски из справочника 64"""
    db = _try_get_db()
    if db:
        return db.find_item_id(64, color_text)
    return find_item_id(color_text, SOLDER_MASK_COLOR_DICT)


def get_silkscreen_color_id(color_text: str) -> Optional[int]:
    """Получить ID цвета маркировки из справочника 66"""
    db = _try_get_db()
    if db:
        return db.find_item_id(66, color_text)
    return find_item_id(color_text, SILKSCREEN_COLOR_DICT)


def get_edge_plating_id(plating_text: str) -> Optional[int]:
    """Получить ID металлизации края из справочника 72"""
    db = _try_get_db()
    if db:
        return db.find_item_id(72, plating_text)
    return find_item_id(plating_text, EDGE_PLATING_DICT)
