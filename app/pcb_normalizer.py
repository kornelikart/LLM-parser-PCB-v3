from __future__ import annotations

import json
import logging
import re
from typing import Any, Optional

logger = logging.getLogger(__name__)

# ═══════════════════════════════════════════════════════════════════
# КАНОНИЧЕСКИЕ СПРАВОЧНИКИ
# Ключ  = NAME из Битрикс24 (verbatim)
# Значение = item_id (code из CSV)
# ═══════════════════════════════════════════════════════════════════

# IBLOCK_ID = 74  Finish Type
FINISH_TYPE: dict[str, int] = {
    "None":                   5926,
    "HASL (PbSn)":            5928,
    "ENEPIG":                 5930,
    "Soft gold":              5932,
    "Imm. gold (chem.Ni/Au)": 5934,
    "HASL LF":                5946,
    "No Coating":             5948,
    "Imm. silver (chem. Ag)": 5950,
    "Imm. tin (chem. Sn)":    5952,
    "Hard gold (Galv. Au)":   5954,
    "Flash gold":             5956,
    "OSP":                    5958,
    "MIX":                    5960,
}

# IBLOCK_ID = 62  Max Copper (base OZ)
COPPER: dict[str, int] = {
    "0":                 5814,
    "1/8 OZ (4.375 um)": 5816,
    "1/4 OZ (8.75 um)":  5818,
    "0.33 OZ (12 um)":   5820,
    "0.5 OZ (17 um)":    5822,
    "1 OZ (35 um)":      5824,
    "1.5 OZ (52um)":     5826,
    "12 OZ (400um)":     5828,
    "10 OZ (350 um)":    5830,
    "2 OZ (70 um)":      5832,
    "3 OZ (105 um)":     5834,
    "4 OZ (140 um)":     5836,
    "5 OZ (175 um)":     5838,
    "6 OZ (210 um)":     5840,
    "7 OZ (245 um)":     5842,
    "8 OZ (280 um)":     5844,
    "9 OZ (315 um)":     5846,
    "Via Migration":     5848,
}

# IBLOCK_ID = 56  Materials
MATERIALS: dict[str, int] = {
    "Aluminum":          5642,
    "MIX/Others":        5646,
    "Rogers 4003":       5650,
    "Polyimide":         5652,
    "Rogers 4350":       5658,
    "Rogers 3003":       5660,
    "Isola FR408":       5664,
    "Isola FR408HR":     5680,
    "Isola ITERA MT40":  5688,
    "Rogers 5880":       5690,
    "Rogers 6010":       5692,
    "Rogers 6002":       5694,
    "Megtron 4":         5716,
    "Isola 370HR":       5742,
    "FR4 TG-180":        5774,
    "Megtron 6":         5792,
    "FR4 TG-135":        5804,
    "FR4 TG-150":        5806,
    "FR4 TG-170":        5808,
}

# IBLOCK_ID = 52  PCB type
PCB_TYPE: dict[str, int] = {
    "Rigid":                6610,
    "Flex":                 6612,
    "Stiffener+Flex":       6614,
    "Stiffener+Rigid+Flex": 6616,
    "Flex+Rigid":           6618,
    "Exotic":               6620,
    "Semi-Flex":            6622,
}

# IBLOCK_ID = 50  Order unit
ORDER_UNIT: dict[str, int] = {
    "ea":  5256,
    "pnl": 5258,
}

# IBLOCK_ID = 86  Peelable SM
PEELABLE_SM: dict[str, int] = {
    "No":  6014,
    "Yes": 7174,
}

# IBLOCK_ID = 160  Production Unit
PRODUCTION_UNIT: dict[str, int] = {
    "ea":  6270,
    "pnl": 6272,
}

# IBLOCK_ID = 72  Edge plating
EDGE_PLATING: dict[str, int] = {
    "No":  5862,
    "Yes": 5864,
}

# IBLOCK_ID = 70  Gold Fingers
GOLD_FINGERS: dict[str, int] = {
    "No":  5860,
    "Yes": 7202,
}

# IBLOCK_ID = 230  Back Drill Holes
BACK_DRILL: dict[str, int] = {
    "No":  6528,
    "Yes": 6530,
}

# IBLOCK_ID = 232  Coin
COIN: dict[str, int] = {
    "No":  6532,
    "Yes": 6534,
}

# IBLOCK_ID = 234  Embedded Components
EMBEDDED_COMPONENTS: dict[str, int] = {
    "No":  6536,
    "Yes": 6538,
}

# IBLOCK_ID = 108  Flex Type
FLEX_TYPE: dict[str, int] = {
    "None":        6106,
    "Single Side": 6108,
    "double side": 6110,
    "Multilayer":  6112,
}

# IBLOCK_ID = 110  Cover layer
COVER_LAYER: dict[str, int] = {
    "No":  6114,
    "Yes": 6116,
}

# IBLOCK_ID = 236  Flex Layer Location
FLEX_LAYER_LOCATION: dict[str, int] = {
    "None":        6540,
    "Inner Layer": 6542,
    "Outer Layer": 6544,
}

# IBLOCK_ID = 66  Shelf Life (Months)
SHELF_LIFE: dict[str, int] = {
    "N/A": 5850,
    "3":   5852,
    "6":   5854,
    "12":  5856,
}

# ufCrm24_1707849930  Outer layers (Copper)
# Ключи совпадают с COPPER, чтобы одно LLM-значение маппилось на оба поля.
OUTER_COPPER_FROM_COPPER: dict[str, int] = {
    "1/8 OZ (4.375 um)": 6320,  # 0.125 OZ (4.375 um) base
    "1/4 OZ (8.75 um)":  6290,  # 1/4 OZ (8.75 um) base
    "0.33 OZ (12 um)":   6306,  # 0.33 OZ (11.5 um) base
    "0.5 OZ (17 um)":    6280,  # 0.5 OZ (17.5 um) base
    "1 OZ (35 um)":      6294,  # 1.0 OZ (35 um) base
    "1.5 OZ (52um)":     6314,  # 1.5 OZ (52um) base
    "2 OZ (70 um)":      6316,  # 2.0 OZ (70 um) base
    "3 OZ (105 um)":     6318,  # 3.0 OZ (105 um) base
    "4 OZ (140 um)":     6322,  # 4.0 OZ (140 um) base
    "5 OZ (175 um)":     6324,  # 5.0 OZ (175 um) base
    "6 OZ (210 um)":     6326,  # 6.0 OZ (210 um) base
    "7 OZ (245 um)":     6282,  # 7.0 OZ (245 um) base
    "8 OZ (280 um)":     6284,  # 8.0 OZ (280 um) base
    "9 OZ (315 um)":     6286,  # 9.0 OZ (315 um) base
    "10 OZ (350 um)":    6288,  # 10.0 OZ (350 um) base
    "12 OZ (400um)":     6292,  # 12 OZ (400um) base
}

# ufCrm24_1707841162  Class (IPC)
IPC_CLASS: dict[str, int] = {
    "IPC Class 2": 6212,
    "IPC Class 3": 6214,
}

# IBLOCK_ID = 54  No of Layers  (name → item_id)
LAYERS: dict[str, int] = {
    "01": 6784, "02": 6786, "04": 6788, "06": 6790,
    "08": 6792, "10": 6794, "12": 6796, "14": 6798,
    "16": 6800, "18": 6802, "20": 6804, "22": 6806,
    "24": 6808, "26": 6810, "28": 6812, "30": 6814,
    "32": 6816, "34": 6818, "36": 6820, "38": 6822,
    "40": 6824, "42": 6826, "44": 6828, "46": 6830,
    "48": 6832, "50": 6834, "52": 6836, "54": 6838,
    "56": 6840, "58": 6842, "60": 6844, "62": 6846,
    "64": 6848,
}

# ═══════════════════════════════════════════════════════════════════
# ПРОМПТ НОРМАЛИЗАЦИИ
# ═══════════════════════════════════════════════════════════════════

_NORMALIZATION_SYSTEM = (
    "You are a PCB specification normalizer. "
    "You map raw human text to exact canonical values from a database. "
    "Respond ONLY with valid JSON — no markdown, no explanation."
)

def _build_normalization_prompt(raw: dict[str, Any]) -> str:
    """
    Формирует промпт для второго вызова LLM.
    Передаёт только поля, требующие нормализации, и полный список допустимых значений.
    """
    relevant = {
        k: raw[k] for k in (
            "coverage_type", "foil_thickness", "base_material", "pcb_type",
            "edge_plating", "peelable_mask", "gold_fingers", "ipc_class",
        ) if raw.get(k)
    }

    finish_keys   = list(FINISH_TYPE.keys())
    copper_keys   = list(COPPER.keys())
    material_keys = list(MATERIALS.keys())
    pcb_type_keys = list(PCB_TYPE.keys())
    ipc_class_keys = list(IPC_CLASS.keys())

    return f"""Map the raw PCB data below to EXACTLY ONE canonical value per field.
Copy the value verbatim from the allowed list, or return null if nothing fits.

RAW DATA:
{json.dumps(relevant, ensure_ascii=False, indent=2)}

ALLOWED VALUES:

finish_type (field: coverage_type):
{json.dumps(finish_keys, ensure_ascii=False)}

Mapping hints for finish_type:
- "ENIG", "Immersion Gold", "Chem.Ni/Au", "ENiG"        → "Imm. gold (chem.Ni/Au)"
- "HASL", "Hot Air Solder", "SnPb HASL"                  → "HASL (PbSn)"
- "HASL LF", "Lead-free HASL", "RoHS HASL", "HASL(LF)"  → "HASL LF"
- "IAg", "Imm Silver", "Chem.Ag"                         → "Imm. silver (chem. Ag)"
- "ISn", "Imm Tin", "Chem.Sn"                            → "Imm. tin (chem. Sn)"
- "Hard Gold", "Galv. Gold", "Electrolytic Gold"         → "Hard gold (Galv. Au)"
- "Soft Gold", "Electroless Gold", "Au"                  → "Soft gold"
- "OSP", "Organic Coating"                               → "OSP"
- "ENEPIG"                                               → "ENEPIG"
- "Flash Gold"                                           → "Flash gold"
- "no coating", "bare copper", "uncoated"                → "No Coating"
- "none", "—", "N/A" (when no finish at all)             → "None"

copper_thickness (field: foil_thickness):
{json.dumps(copper_keys, ensure_ascii=False)}

Mapping hints for copper_thickness:
- "35 um", "35µm", "35 micron", "1oz", "1 oz"           → "1 OZ (35 um)"
- "70 um", "70µm", "2oz", "2 oz"                        → "2 OZ (70 um)"
- "17.5 um", "18 um", "0.5oz", "0.5 oz", "half oz"     → "0.5 OZ (17 um)"
- "105 um", "3oz", "3 oz"                               → "3 OZ (105 um)"
- "52 um", "52.5 um", "1.5oz", "1.5 oz"                → "1.5 OZ (52um)"
- "140 um", "4oz"                                        → "4 OZ (140 um)"
- "175 um", "5oz"                                        → "5 OZ (175 um)"
- "210 um", "6oz"                                        → "6 OZ (210 um)"
- "8.75 um", "0.25oz", "1/4 oz"                         → "1/4 OZ (8.75 um)"
- "4.375 um", "0.125oz", "1/8 oz"                       → "1/8 OZ (4.375 um)"
- "12 um", "0.33oz"                                      → "0.33 OZ (12 um)"
- For composite strings like "Top: 35µm, Inner: 17µm"   → use outer layer value

base_material (field: base_material):
{json.dumps(material_keys, ensure_ascii=False)}

Mapping hints for base_material:
- "FR4", "FR-4" (generic, no TG specified)               → "FR4 TG-180"
- "FR4 TG135", "FR-4 TG-135"                            → "FR4 TG-135"
- "FR4 TG150", "FR-4 TG-150"                            → "FR4 TG-150"
- "FR4 TG170", "FR-4 TG-170"                            → "FR4 TG-170"
- "FR4 TG180", "High Tg FR4", "FR-4 (High TG)"         → "FR4 TG-180"
- "Polyimide", "PI", "Kapton"                            → "Polyimide"
- "Rogers 4003C", "RO4003", "RO4003C"                   → "Rogers 4003"
- "Rogers 4350B", "RO4350", "RO4350B"                   → "Rogers 4350"
- "Aluminum", "Aluminium", "Al substrate"                → "Aluminum"

pcb_type (field: pcb_type):
{json.dumps(pcb_type_keys, ensure_ascii=False)}

Mapping hints for pcb_type:
- "Rigid", "standard", not specified                     → "Rigid"
- "Flex", "flexible", "FPC"                             → "Flex"
- "Rigid-Flex", "Rigid Flex"                            → "Flex+Rigid"
- "Semi-Flex", "semiflex"                               → "Semi-Flex"

ipc_class (field: ipc_class):
{json.dumps(ipc_class_keys, ensure_ascii=False)}

Mapping hints for ipc_class:
- "Class 2", "IPC-2", "IPC 2", "Class II"              → "IPC Class 2"
- "Class 3", "IPC-3", "IPC 3", "Class III"             → "IPC Class 3"

OUTPUT FORMAT (return exactly this JSON, no extra keys):
{{
  "finish_type":      "<canonical value or null>",
  "copper_thickness": "<canonical value or null>",
  "base_material":    "<canonical value or null>",
  "pcb_type":         "<canonical value or null>",
  "ipc_class":        "<canonical value or null>"
}}"""


# ═══════════════════════════════════════════════════════════════════
# НОРМАЛИЗАЦИЯ ЧЕРЕЗ LLM
# ═══════════════════════════════════════════════════════════════════

def normalize_pcb_data(raw: dict[str, Any], mistral_client: Any) -> dict[str, Any]:
    """
    Шаг 2 пайплайна: вызывает LLM для нормализации текстовых полей.

    Возвращает обогащённый словарь:
      - все исходные поля raw сохранены
      - добавлен ключ "_normalized": {finish_type, copper_thickness, base_material, pcb_type}

    После этого вызывайте map_to_bitrix24_ids() для получения item_id.
    """
    prompt = _build_normalization_prompt(raw)

    try:
        response = mistral_client.chat(
            model="mistral-small-latest",
            messages=[
                {"role": "system", "content": _NORMALIZATION_SYSTEM},
                {"role": "user",   "content": prompt},
            ],
            response_format={"type": "json_object"},
            temperature=0,        # детерминированный выбор
        )
        content = response.choices[0].message.content.strip()
        # Снимаем markdown-обёртку (```json ... ```) если модель её добавила
        if content.startswith("```"):
            content = re.sub(r"^```(?:json)?\s*\n?", "", content)
            content = re.sub(r"\n?```\s*$", "", content)
        normalized = json.loads(content)
    except json.JSONDecodeError as e:
        logger.error("Ошибка парсинга JSON от LLM нормализатора: %s | content: %.200s", e, locals().get("content", ""))
        normalized = {}
    except Exception as e:
        logger.error("Ошибка нормализации через LLM: %s", e)
        normalized = {}

    _validate_normalized(normalized)
    _log_normalization_diffs(raw, normalized)

    return {**raw, "_normalized": normalized}


def _validate_normalized(normalized: dict) -> None:
    """Проверяет, что LLM вернул только допустимые значения; обнуляет невалидные."""
    checks = {
        "finish_type":      FINISH_TYPE,
        "copper_thickness": COPPER,
        "base_material":    MATERIALS,
        "pcb_type":         PCB_TYPE,
        "ipc_class":        IPC_CLASS,
    }
    for field, canon_dict in checks.items():
        val = normalized.get(field)
        if val is not None and val not in canon_dict:
            logger.warning(
                "LLM вернул невалидное значение для %s: '%s' — обнулено", field, val
            )
            normalized[field] = None


def _log_normalization_diffs(raw: dict, normalized: dict) -> None:
    """Логирует изменения между сырым и нормализованным значением."""
    field_map = {
        "coverage_type":  "finish_type",
        "foil_thickness": "copper_thickness",
        "base_material":  "base_material",
        "pcb_type":       "pcb_type",
        "ipc_class":      "ipc_class",
    }
    for raw_key, norm_key in field_map.items():
        raw_val  = raw.get(raw_key)
        norm_val = normalized.get(norm_key)
        if not raw_val:
            continue
        if norm_val and raw_val != norm_val:
            logger.info("Нормализация %-20s  '%s'  →  '%s'", norm_key, raw_val, norm_val)
        elif not norm_val:
            logger.warning("Нормализация %-20s  '%s'  →  НЕ НАЙДЕНО", norm_key, raw_val)


# ═══════════════════════════════════════════════════════════════════
# МАППИНГ В ID БИТРИКС24
# ═══════════════════════════════════════════════════════════════════

def map_to_bitrix24_ids(enriched: dict[str, Any]) -> dict[str, int]:
    """
    Преобразует нормализованный словарь в {ufCrm24_*: item_id}.

    Принимает результат normalize_pcb_data().
    Для полей, которые не нормализуются через LLM (layer_count, edge_plating и т.д.),
    использует собственную логику.
    """
    normalized = enriched.get("_normalized", {})
    ids: dict[str, int] = {}

    # ── Поля из промпта 2 (LLM-нормализованные) ──────────────────

    if finish := normalized.get("finish_type"):
        if (v := FINISH_TYPE.get(finish)) is not None:
            ids["ufCrm24_1707768819"] = v

    if copper := normalized.get("copper_thickness"):
        if (v := COPPER.get(copper)) is not None:
            ids["ufCrm24_1707838441"] = v                    # Max Copper (base OZ)
        if (v := OUTER_COPPER_FROM_COPPER.get(copper)) is not None:
            ids["ufCrm24_1707849930"] = v                    # Outer layers (Copper)

    if material := normalized.get("base_material"):
        if (v := MATERIALS.get(material)) is not None:
            ids["ufCrm24_1707838248"] = v

    if pcb_type := normalized.get("pcb_type"):
        if (v := PCB_TYPE.get(pcb_type)) is not None:
            ids["ufCrm24_1707838074"] = v

    if ipc_cls := normalized.get("ipc_class"):
        if (v := IPC_CLASS.get(ipc_cls)) is not None:
            ids["ufCrm24_1707841162"] = v                    # Class

    # ── Количество слоёв (числовой маппинг, не через LLM) ────────

    if layer_raw := enriched.get("layer_count"):
        layer_id = _map_layers(str(layer_raw))
        if layer_id:
            ids["ufCrm24_1709815185"] = layer_id

    # ── Yes/No поля (прямой маппинг, не через LLM) ───────────────

    yes_no_fields: list[tuple[str, str, dict]] = [
        ("edge_plating",        "ufCrm24_1707839110", EDGE_PLATING),
        ("peelable_mask",       "ufCrm24_1707839629", PEELABLE_SM),
        ("gold_fingers",        "ufCrm24_1707838528", GOLD_FINGERS),        # Gold Fingers   iblock_id=70
        ("back_drill",          "ufCrm24_1707851410", BACK_DRILL),          # Back Drill     iblock_id=230
        ("coin",                "ufCrm24_1707851442", COIN),                # Coin           iblock_id=232
        ("embedded_components", "ufCrm24_1707851467", EMBEDDED_COMPONENTS), # Embedded Comp  iblock_id=234
        ("cover_layer",         "ufCrm24_1707840205", COVER_LAYER),         # Cover layer    iblock_id=110
    ]
    for src_key, field_code, canon in yes_no_fields:
        raw_val = enriched.get(src_key)
        if raw_val is None:
            continue
        mapped = _map_yes_no(str(raw_val), canon)
        if mapped is not None:
            ids[field_code] = mapped

    # ── Текстовые справочные поля (не Yes/No, прямой lookup) ─────
    lookup_fields: list[tuple[str, str, dict]] = [
        ("flex_type",          "ufCrm24_1707840178", FLEX_TYPE),          # Flex Type          iblock_id=108
        ("flex_layer_location","ufCrm24_1707851507", FLEX_LAYER_LOCATION), # Flex Layer Location iblock_id=236
    ]
    for src_key, field_code, canon in lookup_fields:
        raw_val = enriched.get(src_key)
        if not raw_val:
            continue
        mapped_v = canon.get(raw_val)
        if mapped_v is None:
            raw_lower = str(raw_val).strip().lower()
            for k, v in canon.items():
                if k.lower() == raw_lower:
                    mapped_v = v
                    break
        if mapped_v is not None:
            ids[field_code] = mapped_v

    # ── Единицы заказа (дефолт: ea) ──────────────────────────────

    order_unit_raw = (enriched.get("order_unit") or "ea").lower().strip()
    ids["ufCrm24_1707838030"] = ORDER_UNIT.get(
        "pnl" if "pnl" in order_unit_raw or "panel" in order_unit_raw else "ea",
        5256,
    )

    ids["ufCrm24_1707849863"] = PRODUCTION_UNIT.get(
        "pnl" if "pnl" in order_unit_raw else "ea",
        6270,
    )

    logger.info("Смаплировано %d полей Битрикс24", len(ids))
    return ids


# ═══════════════════════════════════════════════════════════════════
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ═══════════════════════════════════════════════════════════════════

def _map_layers(layer_text: str) -> Optional[int]:
    """
    Извлекает число из строки ("8 layers" → "08") и ищет в справочнике.
    Битрикс24 хранит значения с ведущим нулём для 1–9 ("01", "02" …).
    """
    numbers = re.findall(r"\d+", layer_text)
    if not numbers:
        return None
    n = int(numbers[0])
    # Битрикс24: 1-слойная → "01", 2 → "02", ... 10+ без нуля
    key = f"{n:02d}" if n < 10 else str(n)
    result = LAYERS.get(key)
    if result is None:
        logger.warning("Количество слоёв %s не найдено в справочнике", key)
    return result


def _map_yes_no(raw: str, canon: dict[str, int]) -> Optional[int]:
    """Маппинг булевых/Yes-No полей."""
    normalized = raw.strip().lower()
    if normalized in ("yes", "да", "1", "true", "y"):
        return canon.get("Yes")
    if normalized in ("no", "нет", "0", "false", "n", "none", "n/a", "—"):
        return canon.get("No")
    # Пробуем напрямую
    for key, val in canon.items():
        if key.lower() == normalized:
            return val
    return None


# ═══════════════════════════════════════════════════════════════════
# ПУБЛИЧНЫЙ API — единая точка входа
# ═══════════════════════════════════════════════════════════════════

def normalize_and_get_ids(
    raw_pcb_data: dict[str, Any],
    mistral_client: Any,
) -> tuple[dict[str, Any], dict[str, int]]:
    """
    Полный пайплайн нормализации:
      1. Вызывает LLM (промпт 2) для нормализации текстовых полей.
      2. Маппирует нормализованные значения на item_id Битрикс24.

    Возвращает:
      enriched  — raw_pcb_data + ключ "_normalized" (для логов/отладки)
      b24_ids   — {ufCrm24_*: item_id} для вставки в Битрикс24

    Пример использования в bitrix24_integration.py:
        enriched, b24_ids = normalize_and_get_ids(pcb_data, mistral_client)
        fields.update(b24_ids)
    """
    enriched = normalize_pcb_data(raw_pcb_data, mistral_client)
    b24_ids  = map_to_bitrix24_ids(enriched)
    return enriched, b24_ids

