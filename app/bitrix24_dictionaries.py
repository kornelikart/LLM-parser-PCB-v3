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
import re

_db_cache = None
_db_cache_initialized = False


def _try_get_db():
    """
    Lazy import to avoid hard dependency on SQLAlchemy when DB mode is not used.
    Returns a DbDictionaries instance or None. Result is cached after first call.
    """
    global _db_cache, _db_cache_initialized
    if _db_cache_initialized:
        return _db_cache
    _db_cache_initialized = True
    if os.getenv("USE_DB_DICTIONARIES", "1").strip() == "0":
        return None
    if not (os.getenv("DICTIONARIES_DB_URL") or "").strip():
        return None
    try:
        from .db_dictionaries import get_db_dictionaries  # type: ignore
        _db_cache = get_db_dictionaries()
    except ImportError:
        try:
            from db_dictionaries import get_db_dictionaries  # type: ignore
            _db_cache = get_db_dictionaries()
        except Exception:
            _db_cache = None
    except Exception:
        _db_cache = None
    return _db_cache

# Справочник 56: Materials (Материал основания платы)
# IBLOCK_ID = 56
# Ключи — текст, который может вернуть LLM или встретиться в спецификации.
MATERIALS_DICT: Dict[str, int] = {
    # Базовые FR‑4
    "fr4": 5774,
    "fr-4": 5774,
    "fr4 tg-135": 5804,
    "fr4 tg-150": 5806,
    "fr4 tg-170": 5808,
    "fr4 tg-180": 5774,
    "High Tg FR-4": 5774,
    "High Tg FR4": 5774,
    "High Tg FR-4 TG-180": 5774,
    "High Tg FR4 TG-180": 5774,
    # Варианты для формулировок из спецификаций (при необходимости замените ID на ваш из Битрикс24)
    "fr-4 (high tg reliability)": 5774,
    "fr4 (high tg reliability)": 5774,
    "FR-4 (high TG reliability)": 5774,
    "FR4 (high TG reliability)": 5774,
    # Часто встречающиеся материалы
    "aluminum": 5642,
    "mix/others": 5646,
    "polyimide": 5652,
    "rogers 4003": 5650,
    "rogers 4350": 5658,
    "rogers 3003": 5660,
    "rogers 5880": 5690,
    "rogers 6002": 5694,
    "rogers 6010": 5692,
    "megtron 4": 5716,
    "megtron 6": 5792,
    "isola fr408": 5664,
    "isola fr408hr": 5680,
    "isola itera mt40": 5688,
    "isola 370hr": 5742,
}

# Справочник 74: Finish Type (Тип финишного покрытия)
# IBLOCK_ID = 74
FINISH_TYPE_DICT: Dict[str, int] = {
    "none": 5926,
    "hasl (pbsn)": 5928,
    "hasl lf": 5946,
    "hasl lead-free": 5946,
    "hasl lead free": 5946,
    "enepig": 5930,
    "soft gold": 5932,
    "imm. gold (chem.ni/au)": 5934,
    # Точные значения (как в Битрикс24), для совместимости с pcb_normalizer.py
    "Imm. gold (chem.Ni/Au)": 5934,
    "ENIG": 5934,
    "Immersion Gold": 5934,
    "Хим. Н.5 Зл.0.1": 5934,
    "Immersion gold (Ni5 Au0.1)": 5934,
    "imm. silver (chem. ag)": 5950,
    "Imm. silver (chem. Ag)": 5950,
    "imm. tin (chem. sn)": 5952,
    "Imm. tin (chem. Sn)": 5952,
    "hard gold (galv. au)": 5954,
    "Hard gold (Galv. Au)": 5954,
    "flash gold": 5956,
    "osp": 5958,
    "mix": 5960,
    "no coating": 5948,
    "HASL (PbSn)": 5928,
    "HASL LF": 5946,
    "No Coating": 5948,
}

# Справочник 54: No of Layers (Количество слоев)
# IBLOCK_ID = 54
LAYERS_DICT: Dict[str, int] = {
    # Основной диапазон слоев (дублируем представление без ведущих нулей)
    "1": 6784,
    "01": 6784,
    "2": 6786,
    "02": 6786,
    "4": 6788,
    "04": 6788,
    "6": 6790,
    "06": 6790,
    "8": 6792,
    "08": 6792,
    "10": 6794,
    "12": 6796,
    "14": 6798,
    "16": 6800,
    "18": 6802,
    "20": 6804,
    "22": 6806,
    "24": 6808,
    "26": 6810,
    "28": 6812,
    "30": 6814,
    "32": 6816,
    "34": 6818,
    "36": 6820,
    "38": 6822,
    "40": 6824,
    "42": 6826,
    "44": 6828,
    "46": 6830,
    "48": 6832,
    "50": 6834,
    "52": 6836,
    "54": 6838,
    "56": 6840,
    "58": 6842,
    "60": 6844,
    "62": 6846,
    "64": 6848,
}

# Справочник 62: Max Copper (base OZ) - Толщина меди
# IBLOCK_ID = 62
COPPER_THICKNESS_DICT: Dict[str, int] = {
    "0": 5814,
    "0 oz": 5814,
    "0 OZ": 5814,
    "1/8 oz (4.375 um)": 5816,
    "1/8 OZ (4.375 um)": 5816,
    "0.125": 5816,
    "1/4 oz (8.75 um)": 5818,
    "1/4 OZ (8.75 um)": 5818,
    "0.25": 5818,
    "0.33 oz (12 um)": 5820,
    "0.33 OZ (12 um)": 5820,
    "0.33": 5820,
    "0.5 oz (17 um)": 5822,
    "0.5 OZ (17 um)": 5822,
    "0.5": 5822,
    "0,018": 5822,
    "1 oz (35 um)": 5824,
    "1 OZ (35 um)": 5824,
    "1": 5824,
    "0,035": 5824,
    "1.5 oz (52um)": 5826,
    "1.5 OZ (52um)": 5826,
    "1.5": 5826,
    "0,052": 5826,
    "2 oz (70 um)": 5832,
    "2 OZ (70 um)": 5832,
    "2": 5832,
    "0,070": 5832,
    "3 oz (105 um)": 5834,
    "3 OZ (105 um)": 5834,
    "3": 5834,
    "0,105": 5834,
    "4 oz (140 um)": 5836,
    "4 OZ (140 um)": 5836,
    "4": 5836,
    "0,140": 5836,
    "5 oz (175 um)": 5838,
    "5 OZ (175 um)": 5838,
    "5": 5838,
    "6 oz (210 um)": 5840,
    "6 OZ (210 um)": 5840,
    "6": 5840,
    "7 oz (245 um)": 5842,
    "7 OZ (245 um)": 5842,
    "7": 5842,
    "8 oz (280 um)": 5844,
    "8 OZ (280 um)": 5844,
    "8": 5844,
    "9 oz (315 um)": 5846,
    "9 OZ (315 um)": 5846,
    "9": 5846,
    "10 OZ (350 um)": 5830,
    "12 oz (400um)": 5828,
    "12 OZ (400um)": 5828,
    "12": 5828,
    "via migration": 5848,
    "Via Migration": 5848,
}

# Справочник 50: Order unit (Единица заказа)
# IBLOCK_ID = 50
ORDER_UNIT_DICT: Dict[str, int] = {
    "ea": 5256,
    "шт": 5256,
    "piece": 5256,
    "pcs": 5256,
    "pnl": 5258,
    "panel": 5258,
    "панель": 5258,
}

# Справочник 52: PCB type (Тип платы)
# IBLOCK_ID = 52
PCB_TYPE_DICT: Dict[str, int] = {
    "rigid": 6610,
    "rigid pcb": 6610,
    "flex": 6612,
    "flex pcb": 6612,
    "stiffener+flex": 6614,
    "stiffener+rigid+flex": 6616,
    "flex+rigid": 6618,
    "exotic": 6620,
    "semi-flex": 6622,
    "semiflex": 6622,
}

# Справочник 86: Peelable SM (Пилинг-маска)
# IBLOCK_ID = 86
PEELABLE_SM_DICT: Dict[str, int] = {
    "no": 6014,
    "нет": 6014,
    "none": 6014,
    "yes": 7174,
    "да": 7174,
}

# Справочник 160: Production Unit (Производственный участок)
# IBLOCK_ID = 160
PRODUCTION_UNIT_DICT: Dict[str, int] = {
    "ea": 6270,
    "шт": 6270,
    "pnl": 6272,
    "panel": 6272,
    "панель": 6272,
}

# Справочник 64: Solder Mask Color (Цвет паяльной маски). ID уточните в Битрикс24.
SOLDER_MASK_COLOR_DICT: Dict[str, int] = {
    "green": 8002,
    "Green": 8002,
    "red": 8002,
    "Red": 8002,
    "blue": 8002,
    "Blue": 8002,
    "black": 8002,
    "Black": 8002,
    "white": 8002,
    "White": 8002,
    "yellow": 8002,
    "Yellow": 8002,
}

# Справочник 66: SilkScreen Color (Цвет маркировки). ID уточните в Битрикс24.
SILKSCREEN_COLOR_DICT: Dict[str, int] = {
    "white": 7002,
    "White": 7002,
    "black": 7002,
    "Black": 7002,
    "green": 7002,
    "Green": 7002,
}

# Справочник 72: Edge plating (Металлизация края)
EDGE_PLATING_DICT: Dict[str, int] = {
    "yes": 5864,
    "да": 5864,
    "no": 5862,
    "нет": 5862,
    "none": 5862,
    "n/a": 5862,
    "—": 5862,
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
    При нечётком совпадении приоритет у самого длинного (наиболее специфичного) ключа,
    чтобы, например, "Immersion gold (Ni5 Au0.1)" не матчился как "soft gold".
    
    Args:
        text_value: Текстовое значение для поиска
        dictionary: Словарь маппинга {текст: item_id}
        fuzzy_match: Если True, использует нечеткое сравнение
    
    Returns:
        item_id или None, если не найдено
    """
    if not text_value:
        return None

    normalized_input = normalize_text(text_value)
    input_lower = text_value.lower().strip()

    # 1) Точное совпадение (нормализованное)
    for key, item_id in dictionary.items():
        if normalize_text(key) == normalized_input:
            return item_id

    # 2) Нечёткое совпадение: перебираем ключи от длинных к коротким,
    #    чтобы "Immersion gold (Ni5 Au0.1)" сработал раньше, чем "soft gold"
    if fuzzy_match:
        for key, item_id in sorted(dictionary.items(), key=lambda x: -len(x[0])):
            key_lower = key.lower().strip()
            key_norm = normalize_text(key)
            # Точное вхождение ключа в текст (или наоборот)
            if key_lower in input_lower or input_lower in key_lower:
                return item_id
            if key_norm and key_norm in normalized_input:
                return item_id
            # Совпадение по словам — только если ключ целиком "покрыт" входом
            # (вход содержит все значимые слова ключа), иначе пропускаем
            key_words = set(w for w in key_lower.split() if len(w) > 1)
            input_words = set(w for w in input_lower.split() if len(w) > 1)
            if key_words and key_words <= input_words:
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
    """Получить ID количества слоев из справочника 54. Из строк вроде '8' или '8 layers' берётся число."""
    db = _try_get_db()
    if db:
        return db.find_item_id(54, str(layers_text))
    numbers = re.findall(r'\d+', str(layers_text))
    if numbers:
        layers_count = numbers[0]
        return find_item_id(layers_count, LAYERS_DICT)
    return find_item_id(layers_text, LAYERS_DICT)


def _extract_primary_copper_thickness(thickness_text: str) -> str:
    """
    Из строк вроде "Top/Bot: 35 µm (1 oz), Inner: 18 µm (0.5 oz)" извлекает
    основную толщину (обычно внешние слои), чтобы не матчить "0" из "0.5 oz".
    """
    if not thickness_text or not thickness_text.strip():
        return thickness_text
    text = thickness_text.strip()
    # Берём первый сегмент до запятой (часто Top/Bot / внешние слои)
    first_part = text.split(",")[0].strip()
    # Ищем "X oz" или "X oz" — X может быть 0.125, 0.25, 0.5, 1, 1.5, 2 и т.д.
    oz_match = re.search(r"(\d+(?:\.\d+)?)\s*oz", first_part, re.IGNORECASE)
    if oz_match:
        val = oz_match.group(1)
        # Не используем "0" как основную толщину, если есть что-то вроде "0.5 oz"
        if val == "0":
            # В первом сегменте только 0 oz — попробуем весь текст
            oz_all = re.search(r"(\d+(?:\.\d+)?)\s*oz", text, re.IGNORECASE)
            if oz_all and oz_all.group(1) != "0":
                return oz_all.group(0).strip()  # e.g. "1 oz"
            return "0 oz"
        return oz_match.group(0).strip()
    # Альтернатива: "35 µm" / "35um" -> считаем 1 oz
    um_match = re.search(r"(\d+)\s*µm", first_part, re.IGNORECASE)
    if um_match:
        return first_part
    return thickness_text


def get_copper_thickness_id(thickness_text: str) -> Optional[int]:
    """Получить ID толщины меди из справочника 62. Для составных строк (Top/Bot: 1 oz, Inner: 0.5 oz) учитывается первая толщина."""
    db = _try_get_db()
    if db:
        return db.find_item_id(62, thickness_text)
    primary = _extract_primary_copper_thickness(thickness_text)
    return find_item_id(primary, COPPER_THICKNESS_DICT)


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
