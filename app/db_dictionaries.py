"""
DB-backed dictionaries for mapping human text values to Bitrix24 dictionary item IDs.

This module is intentionally lightweight:
- Reads DB connection URL from env var `DICTIONARIES_DB_URL`
- Queries a single table/view that stores dictionary mappings
- Adds an in-process TTL cache to avoid hammering the DB

Expected DB schema (can be a table or a VIEW):
  - param_id    (int)     : Bitrix24 list (dictionary) ID, e.g. 56 for Materials (same as iblock_id)
  - param       (text)    : Name of the dictionary, e.g. "Material"
  - item_id     (int)     : ID of element inside that list (value to send to Bitrix24)
  - item        (text)    : Human-readable value, e.g. "FR4"

If your schema differs, adjust `SQL_BY_IBLOCK_ID` or provide a compatible VIEW.
"""

from __future__ import annotations

import os
import time
from dataclasses import dataclass
from typing import Optional

def _get_sqlalchemy():
    """
    Import SQLAlchemy lazily so the project can run without DB dependencies
    when DICTIONARIES_DB_URL is not set.
    """
    try:
        from sqlalchemy import create_engine, text  # type: ignore
    except Exception as e:  # pragma: no cover
        raise ImportError(
            "SQLAlchemy is required for DB dictionaries. "
            "Install dependencies (pip install -r requirements.txt) "
            "or disable DB dictionaries (USE_DB_DICTIONARIES=0)."
        ) from e
    return create_engine, text


def _sql_by_iblock_id():
    _create_engine, _text = _get_sqlalchemy()
    return _text(
        """
        SELECT item_id, item
        FROM bitrix24_dictionaries
        WHERE param_id = :iblock_id
        """
    )


def normalize_text(text_value: str) -> str:
    if not text_value:
        return ""
    return text_value.strip().lower().replace("-", "").replace("_", "").replace(" ", "")


@dataclass(frozen=True)
class DbDictConfig:
    db_url: str
    cache_ttl_seconds: int = 900  # 15 minutes


class DbDictionaries:
    def __init__(self, cfg: DbDictConfig):
        self._cfg = cfg
        create_engine, _text = _get_sqlalchemy()
        # pool_pre_ping avoids stale connections on long-running servers
        self._engine = create_engine(cfg.db_url, pool_pre_ping=True)
        self._cache: dict[tuple[int, str], tuple[float, Optional[int]]] = {}

    def _get_cached(self, iblock_id: int, normalized_value: str) -> Optional[int] | None:
        key = (iblock_id, normalized_value)
        cached = self._cache.get(key)
        if not cached:
            return None
        expires_at, item_id = cached
        if expires_at < time.time():
            self._cache.pop(key, None)
            return None
        return item_id

    def _set_cached(self, iblock_id: int, normalized_value: str, item_id: Optional[int]) -> None:
        expires_at = time.time() + self._cfg.cache_ttl_seconds
        self._cache[(iblock_id, normalized_value)] = (expires_at, item_id)

    def find_item_id(self, iblock_id: int, text_value: str) -> Optional[int]:
        normalized_value = normalize_text(text_value)
        if not normalized_value:
            return None

        cached = self._get_cached(iblock_id, normalized_value)
        if cached is not None:
            return cached

        stmt = _sql_by_iblock_id()
        with self._engine.connect() as conn:
            rows = conn.execute(stmt, {"iblock_id": iblock_id}).fetchall()

        # Exact normalized match first
        for item_id, item in rows:
            if normalize_text(str(item)) == normalized_value:
                self._set_cached(iblock_id, normalized_value, int(item_id))
                return int(item_id)

        # Fuzzy-ish fallback: substring / shared words (matches previous behavior)
        input_raw = str(text_value).lower().strip()
        input_words = set(input_raw.split())
        for item_id, item in rows:
            key_raw = str(item).lower().strip()
            if key_raw in input_raw or input_raw in key_raw:
                self._set_cached(iblock_id, normalized_value, int(item_id))
                return int(item_id)
            key_words = set(key_raw.split())
            if input_words & key_words:
                self._set_cached(iblock_id, normalized_value, int(item_id))
                return int(item_id)

        self._set_cached(iblock_id, normalized_value, None)
        return None


_singleton: Optional[DbDictionaries] = None


def get_db_dictionaries() -> Optional[DbDictionaries]:
    """
    Returns a singleton DbDictionaries instance if DICTIONARIES_DB_URL is set.
    If env var is missing/empty, returns None (caller can use fallback dictionaries).
    """
    global _singleton
    db_url = (os.getenv("DICTIONARIES_DB_URL") or "").strip()
    if not db_url:
        return None
    if _singleton is None:
        ttl = int((os.getenv("DICTIONARIES_CACHE_TTL_SECONDS") or "900").strip() or "900")
        _singleton = DbDictionaries(DbDictConfig(db_url=db_url, cache_ttl_seconds=ttl))
    return _singleton

