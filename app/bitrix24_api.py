"""
Живые справочники и метаданные полей смарт-процесса из Битрикс24 REST API.

Зачем: статические словари в pcb_normalizer.py / bitrix24_dictionaries.py
устаревают, когда в Битрикс24 добавляют значения (например, новый тип покрытия),
а коды некоторых полей вообще отсутствуют в документации. Этот модуль читает
и то, и другое напрямую с портала.

Что даёт:
  * get_fields()            — {ufCrm24_*: FieldMeta(title, type, iblock_id, items)}
  * get_field_values(code)  — {название значения: item_id} для списочного поля
  * get_dictionary(iblock)  — {название значения: item_id} по IBLOCK_ID справочника
  * find_field(keywords)    — поиск поля по ключевым словам в названии

Всё кэшируется в процессе на BITRIX24_CACHE_TTL секунд (по умолчанию 15 минут).
Отключается переменной USE_BITRIX24_API_DICTIONARIES=0.
"""
from __future__ import annotations

import logging
import os
import re
import threading
import time
from dataclasses import dataclass, field as dc_field
from typing import Any, Dict, List, Optional

import httpx

try:
    from logger import setup_logger
except ImportError:
    from .logger import setup_logger

logger = setup_logger(level=logging.INFO)

DEFAULT_ENTITY_TYPE_ID = 182


@dataclass
class FieldMeta:
    """Метаданные пользовательского поля смарт-процесса."""
    code: str
    title: str
    type: str
    iblock_id: Optional[int] = None
    is_required: bool = False
    # Значения списочных полей типа enumeration приходят сразу в описании поля.
    inline_items: Dict[str, int] = dc_field(default_factory=dict)

    @property
    def is_list_like(self) -> bool:
        return bool(self.iblock_id) or bool(self.inline_items) or self.type in (
            "enumeration", "crm_status", "iblock_element", "iblock_section",
        )


def _extract_iblock_id(raw: dict) -> Optional[int]:
    """Достаёт IBLOCK_ID поля. Битрикс24 кладёт его в settings под разными именами."""
    settings = raw.get("settings") or {}
    if not isinstance(settings, dict):
        return None
    for key in ("IBLOCK_ID", "iblockId", "iblock_id", "IBLOCKID"):
        value = settings.get(key)
        if value in (None, "", 0, "0"):
            continue
        try:
            return int(value)
        except (TypeError, ValueError):
            continue
    return None


def _extract_inline_items(raw: dict) -> Dict[str, int]:
    """Значения enumeration/crm_status приходят списком прямо в описании поля."""
    items = raw.get("items")
    result: Dict[str, int] = {}
    if isinstance(items, list):
        for item in items:
            if not isinstance(item, dict):
                continue
            name = item.get("VALUE") or item.get("value") or item.get("NAME") or item.get("name")
            ident = item.get("ID") or item.get("id") or item.get("STATUS_ID")
            if name is None or ident is None:
                continue
            try:
                result[str(name).strip()] = int(ident)
            except (TypeError, ValueError):
                continue
    return result


class Bitrix24Api:
    """Клиент REST API Битрикс24 с TTL-кэшем справочников."""

    def __init__(
        self,
        webhook_url: str,
        entity_type_id: int = DEFAULT_ENTITY_TYPE_ID,
        cache_ttl: int = 900,
    ):
        self._base = self._normalize_base(webhook_url)
        self._entity_type_id = entity_type_id
        self._ttl = cache_ttl
        self._lock = threading.Lock()
        self._fields_cache: Optional[Dict[str, FieldMeta]] = None
        self._fields_expires_at: float = 0.0
        self._dict_cache: Dict[int, tuple] = {}   # iblock_id -> (expires_at, {name: id})

    @staticmethod
    def _normalize_base(webhook_url: str) -> str:
        """Из webhook любого вида делает базовый URL без имени метода."""
        base = (webhook_url or "").strip().rstrip("/")
        base = re.sub(r"/(crm|lists)\.[a-zA-Z.]+$", "", base)
        return base

    # ── низкоуровневый вызов ────────────────────────────────────────

    def _call(self, method: str, params: Optional[dict] = None) -> Any:
        """Вызов метода REST API.

        Битрикс24 — российский сервис: сначала пробуем напрямую (мимо VPN-proxy),
        при сетевой ошибке — через proxy. Логика та же, что в create_bitrix24_item.
        """
        if not self._base:
            raise ValueError("Не задан webhook Битрикс24.")
        url = f"{self._base}/{method}"
        forced = (os.getenv("BITRIX24_TRUST_ENV") or "").strip()
        trust_modes = [forced == "1"] if forced else [False, True]

        last_error: Optional[Exception] = None
        for trust_env in trust_modes:
            try:
                with httpx.Client(timeout=60.0, trust_env=trust_env) as client:
                    response = client.post(url, json=params or {})
                    response.raise_for_status()
                    data = response.json()
                break
            except httpx.RequestError as e:
                last_error = e
                continue
        else:
            raise ConnectionError(f"Битрикс24 недоступен: {last_error}")

        if isinstance(data, dict) and "error" in data:
            raise RuntimeError(
                f"Битрикс24 вернул ошибку для {method}: "
                f"{data.get('error_description') or data['error']}"
            )
        return data

    def _call_paged(self, method: str, params: dict, items_key: str) -> List[dict]:
        """Постраничный вызов: Битрикс24 отдаёт по 50 записей за раз."""
        collected: List[dict] = []
        start = 0
        while True:
            page = self._call(method, {**params, "start": start})
            result = page.get("result")
            if isinstance(result, dict):
                chunk = result.get(items_key) or []
            else:
                chunk = result or []
            if isinstance(chunk, dict):
                chunk = list(chunk.values())
            collected.extend(c for c in chunk if isinstance(c, dict))

            next_start = page.get("next")
            if next_start is None or not chunk:
                break
            start = int(next_start)
            if start > 10000:  # страховка от бесконечного цикла
                logger.warning("Пагинация %s прервана на %s записях", method, len(collected))
                break
        return collected

    # ── метаданные полей ────────────────────────────────────────────

    def get_fields(self, force: bool = False) -> Dict[str, FieldMeta]:
        """Все пользовательские поля смарт-процесса: {код: FieldMeta}."""
        with self._lock:
            if not force and self._fields_cache is not None and self._fields_expires_at > time.time():
                return self._fields_cache

        data = self._call("crm.item.fields", {"entityTypeId": self._entity_type_id})
        raw_fields = (data.get("result") or {}).get("fields") or {}

        fields: Dict[str, FieldMeta] = {}
        for code, raw in raw_fields.items():
            if not isinstance(raw, dict) or not code.lower().startswith("ufcrm"):
                continue
            meta = FieldMeta(
                code=code,
                title=str(raw.get("title") or raw.get("formLabel") or "").strip(),
                type=str(raw.get("type") or "").strip(),
                iblock_id=_extract_iblock_id(raw),
                is_required=bool(raw.get("isRequired")),
                inline_items=_extract_inline_items(raw),
            )
            fields[code] = meta

        with self._lock:
            self._fields_cache = fields
            self._fields_expires_at = time.time() + self._ttl
        logger.info("Битрикс24: загружено метаданных полей — %d", len(fields))
        return fields

    def find_field(self, keywords: List[str], exclude: Optional[List[str]] = None) -> Optional[FieldMeta]:
        """Поиск поля по ключевым словам в названии.

        Возвращает поле, только если совпадение однозначно (один кандидат);
        при нескольких кандидатах пишет их в лог и возвращает None, чтобы
        не отправить значение не в то поле.

        Бросает исключение, если портал недоступен: вызывающий код должен
        отличать «поле не найдено» от «список полей не получен», иначе
        временный сбой сети навсегда отключит характеристику.
        """
        exclude = [e.casefold() for e in (exclude or [])]
        fields = self.get_fields()

        matches = []
        for meta in fields.values():
            title = meta.title.casefold()
            if not title:
                continue
            if any(x in title for x in exclude):
                continue
            if any(k.casefold() in title for k in keywords):
                matches.append(meta)

        if not matches:
            return None
        if len(matches) > 1:
            logger.warning(
                "Неоднозначный поиск поля по %s — кандидаты: %s. "
                "Задайте код явно через .env.",
                keywords, ", ".join(f"{m.code} ({m.title})" for m in matches),
            )
            return None
        return matches[0]

    # ── справочники ─────────────────────────────────────────────────

    def get_dictionary(self, iblock_id: int, force: bool = False) -> Dict[str, int]:
        """Элементы справочника (универсального списка): {название: item_id}."""
        with self._lock:
            cached = self._dict_cache.get(iblock_id)
            if not force and cached and cached[0] > time.time():
                return cached[1]

        items: Dict[str, int] = {}
        last_error: Optional[Exception] = None
        # Справочники смарт-процессов лежат либо в обычных списках, либо в
        # списках бизнес-процессов — пробуем оба типа.
        for iblock_type in ("lists", "bitrix_processes"):
            try:
                rows = self._call_paged(
                    "lists.element.get",
                    {"IBLOCK_TYPE_ID": iblock_type, "IBLOCK_ID": iblock_id},
                    items_key="items",
                )
            except Exception as e:
                last_error = e
                continue
            for row in rows:
                name = row.get("NAME") or row.get("name")
                ident = row.get("ID") or row.get("id")
                if name is None or ident is None:
                    continue
                try:
                    items[str(name).strip()] = int(ident)
                except (TypeError, ValueError):
                    continue
            if items:
                break

        if not items:
            logger.warning(
                "Справочник iblock %s пуст или недоступен%s",
                iblock_id, f": {last_error}" if last_error else "",
            )

        with self._lock:
            self._dict_cache[iblock_id] = (time.time() + self._ttl, items)
        if items:
            logger.info("Битрикс24: справочник iblock %s — %d значений", iblock_id, len(items))
        return items

    def get_field_values(self, code: str) -> Dict[str, int]:
        """Допустимые значения списочного поля: {название: id}, независимо от типа."""
        fields = self.get_fields()
        meta = fields.get(code)
        if not meta:
            return {}
        if meta.inline_items:
            return dict(meta.inline_items)
        if meta.iblock_id:
            return self.get_dictionary(meta.iblock_id)
        return {}

    def clear_cache(self) -> None:
        with self._lock:
            self._fields_cache = None
            self._fields_expires_at = 0.0
            self._dict_cache.clear()
        logger.info("Кэш справочников Битрикс24 очищен.")


# ═══════════════════════════════════════════════════════════════════
# Синглтон
# ═══════════════════════════════════════════════════════════════════

_api_singleton: Optional[Bitrix24Api] = None
_singleton_lock = threading.Lock()
_singleton_initialized = False


def get_bitrix24_api() -> Optional[Bitrix24Api]:
    """Клиент API, если задан webhook и режим не отключён явно."""
    global _api_singleton, _singleton_initialized
    with _singleton_lock:
        if _singleton_initialized:
            return _api_singleton
        _singleton_initialized = True

        if (os.getenv("USE_BITRIX24_API_DICTIONARIES") or "1").strip() == "0":
            logger.info("Живые справочники Битрикс24 отключены (USE_BITRIX24_API_DICTIONARIES=0).")
            return None

        webhook = (os.getenv("BITRIX24_WEBHOOK_URL") or "").strip()
        if not webhook:
            token = (os.getenv("BITRIX24_TOKEN") or "").strip()
            base = (os.getenv("BITRIX24_BASE_URL") or "https://fineline.bitrix24.ru/rest/6").strip()
            webhook = f"{base}/{token}" if token else ""
        if not webhook:
            return None

        try:
            ttl = int((os.getenv("BITRIX24_CACHE_TTL") or "900").strip() or "900")
        except ValueError:
            ttl = 900
        try:
            entity_type_id = int((os.getenv("BITRIX24_ENTITY_TYPE_ID") or "").strip() or DEFAULT_ENTITY_TYPE_ID)
        except ValueError:
            entity_type_id = DEFAULT_ENTITY_TYPE_ID

        _api_singleton = Bitrix24Api(webhook, entity_type_id=entity_type_id, cache_ttl=ttl)
        return _api_singleton


def reset_bitrix24_api() -> None:
    """Сбрасывает синглтон (для тестов и смены webhook на лету)."""
    global _api_singleton, _singleton_initialized
    with _singleton_lock:
        _api_singleton = None
        _singleton_initialized = False
