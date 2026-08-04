# -*- coding: utf-8 -*-
"""
Запрашивает у Битрикс24 список полей смарт-процесса и ищет коды полей,
которых нет в документации «Поля PCB» (импеданс, сторона маркировки и т.д.).

Что делает:
  1. crm.item.fields?entityTypeId=182 — полный список полей с названиями и типами;
  2. ищет кандидатов по ключевым словам для каждой ненастроенной характеристики;
  3. печатает готовые строки для .env;
  4. для списочных полей (iblock_element) показывает, что нужна выгрузка справочника.

Требуется BITRIX24_WEBHOOK_URL в .env (или передайте webhook первым аргументом).

Запуск из корня проекта:
    python tools/fetch_bitrix24_fields.py
    python tools/fetch_bitrix24_fields.py https://fineline.bitrix24.ru/rest/6/<token>/
    python tools/fetch_bitrix24_fields.py --all      # напечатать вообще все поля
"""
import json
import re
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

import httpx  # noqa: E402

from app.bitrix24 import _mask_webhook_url  # noqa: E402
from app.config import bitrix24_config  # noqa: E402

ENTITY_TYPE_ID = bitrix24_config.get("entity_type_id", 182)

# Характеристика → (переменная .env, ключевые слова для поиска в названии поля)
TARGETS = {
    "No. Of Diff Boards": (
        "B24_FIELD_NO_OF_DIFF_BOARDS",
        ["diff board", "different board", "no. of diff", "тип плат", "разных плат"],
    ),
    "Мин. диаметр отверстия": (
        "B24_FIELD_MIN_HOLE_SIZE",
        ["min hole", "minimum hole", "smallest hole", "hole size", "hole dia",
         "min drill", "мин. отверст", "минимальн", "диаметр отверст"],
    ),
    "Контроль импеданса": (
        "B24_FIELD_IMPEDANCE_CONTROL",
        ["imped", "импеданс", "волнов", "controlled imp"],
    ),
    "Сторона маркировки": (
        "B24_FIELD_MARKING_SIDE",
        ["marking", "legend", "silkscreen", "silk screen", "маркиров", "шелкограф"],
    ),
    "Серийный номер": (
        "B24_FIELD_SERIAL_NUMBER",
        ["serial", "barcode", "bar code", "unique id", "серийн", "порядков", "штрих"],
    ),
}


def build_url(base: str, method: str) -> str:
    base = base.strip().rstrip("/")
    base = re.sub(r"/crm\.[a-z.]+$", "", base)
    return f"{base}/{method}"


def fetch(url: str) -> dict:
    """Запрос к Битрикс24: сначала мимо proxy, при сетевой ошибке — через него."""
    last_error = None
    for trust_env in (False, True):
        try:
            with httpx.Client(timeout=60.0, trust_env=trust_env) as client:
                response = client.get(url)
                response.raise_for_status()
                return response.json()
        except httpx.RequestError as e:
            last_error = e
            continue
    raise SystemExit(f"Не удалось подключиться к Битрикс24: {last_error}")


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    show_all = "--all" in sys.argv

    webhook = args[0] if args else (
        bitrix24_config.get("webhook_url") or bitrix24_config.get("token") or ""
    )
    if not webhook:
        print("Не задан webhook. Укажите BITRIX24_WEBHOOK_URL в .env или передайте\n"
              "его первым аргументом: python tools/fetch_bitrix24_fields.py <webhook_url>")
        return 2

    url = build_url(webhook, "crm.item.fields") + f"?entityTypeId={ENTITY_TYPE_ID}"
    print(f"Запрос: {_mask_webhook_url(url)}\n")

    data = fetch(url)
    if "error" in data:
        print(f"Ошибка Битрикс24: {data.get('error_description', data['error'])}")
        return 1

    fields = data.get("result", {}).get("fields", {})
    if not fields:
        print("Ответ получен, но список полей пуст. Проверьте права webhook и entityTypeId.")
        return 1

    uf_fields = {k: v for k, v in fields.items() if k.lower().startswith("ufcrm")}
    print(f"Всего полей: {len(fields)} | пользовательских (ufCrm24_*): {len(uf_fields)}\n")

    out_path = PROJECT_ROOT / "tools" / "bitrix24_fields.json"
    out_path.write_text(json.dumps(fields, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Полный список сохранён: {out_path}\n")

    if show_all:
        print("═" * 78)
        print("ВСЕ ПОЛЬЗОВАТЕЛЬСКИЕ ПОЛЯ")
        print("═" * 78)
        for code, info in sorted(uf_fields.items()):
            title = info.get("title") or info.get("formLabel") or ""
            print(f"  {code:26} {info.get('type', ''):16} {title}")
        print()

    print("═" * 78)
    print("КАНДИДАТЫ ДЛЯ НЕНАСТРОЕННЫХ ХАРАКТЕРИСТИК")
    print("═" * 78)
    env_lines = []
    for label, (env_name, keywords) in TARGETS.items():
        print(f"\n── {label}  ({env_name})")
        matches = []
        for code, info in uf_fields.items():
            title = str(info.get("title") or info.get("formLabel") or "")
            if any(k in title.lower() for k in keywords):
                matches.append((code, info.get("type", ""), title))
        if not matches:
            print("   совпадений по названию не найдено")
            continue
        for code, ftype, title in matches:
            note = ""
            if ftype == "crm_status" or "enumeration" in str(ftype):
                note = "  ← списочное: нужны ID вариантов (см. items в bitrix24_fields.json)"
            print(f"   {code:26} {ftype:16} {title}{note}")
        if len(matches) == 1:
            env_lines.append(f"{env_name}={matches[0][0]}")

    if env_lines:
        print("\n" + "═" * 78)
        print("ГОТОВЫЕ СТРОКИ ДЛЯ .env (однозначные совпадения)")
        print("═" * 78)
        for line in env_lines:
            print("  " + line)

    print("\nЕсли нужного поля нет — возможно, оно называется иначе. "
          "Запустите с --all и найдите его глазами.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
