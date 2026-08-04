# -*- coding: utf-8 -*-
"""
Юнит-тесты маппинга в поля Битрикс24. Работают офлайн: без API-ключа,
без webhook, без сети (обращения к порталу подменяются заглушкой).

Запуск из корня проекта:
    python tests/run_unit_tests.py

Код возврата 0 — все проверки пройдены. Прогоняйте после правок
pcb_normalizer.py, bitrix24.py и справочников.
"""
import os
import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

# Логи и артефакты — во временный каталог, проект не засоряем.
os.chdir(tempfile.mkdtemp())
os.environ["LOG_FILE"] = os.path.join(os.getcwd(), "logs.log")
os.environ.pop("BITRIX24_WEBHOOK_URL", None)
os.environ.pop("BITRIX24_TOKEN", None)

import httpx  # noqa: E402

from app import bitrix24 as b  # noqa: E402
from app import bitrix24_api, pcb_normalizer as pn  # noqa: E402

FAILS = []


def check(label, actual, expected):
    ok = actual == expected
    print(f"  {'✓' if ok else '✗'} {label}: {actual!r}" + ("" if ok else f"   ОЖИДАЛОСЬ {expected!r}"))
    if not ok:
        FAILS.append(label)


def section(title):
    print(f"\n═══ {title} ═══")


# ─────────────────────────────────────────────────────────────────
section("1. Габариты: допуски не должны попадать в размеры")
for label, raw, exp in [
    ("квалитет h12 (бланк ЗАСЛОН)", "114 h12 x 47 h12", (114.0, 47.0)),
    ("сырые строки формы", "длина 114 Допуск h12, ширина 47 h12", (114.0, 47.0)),
    ("квалитет js13", "90 js13 x 60 js13", (90.0, 60.0)),
    ("± допуск", "303±0.2 x 111.0±0.2 mm", (303.0, 111.0)),
    ("+/- допуск раздельно", "160 x 120 +0,3 -0,3", (160.0, 120.0)),
    ("процентный допуск", "100 x 75 ±10%", (100.0, 75.0)),
    ("дефис как разделитель", "114-47", (114.0, 47.0)),
    ("запятая как дробная", "160мм x 120мм ±0,2мм", (160.0, 120.0)),
    ("круглая плата — не пара", "Ø120,0 ±0,1", None),
    ("пустое значение", "", None),
]:
    check(label, b._parse_dimensions(raw, 5.0), exp)

# ─────────────────────────────────────────────────────────────────
section("2. Да/Нет в русских и английских бланках")
for raw, exp in [
    ("есть", True), ("Да", True), ("требуется", True),
    ("требуется в формате ХХХ", True), ("Yes", True), ("имеется", True),
    ("нет", False), ("Нет", False), ("не требуется", False),
    ("отсутствует", False), ("No", False), ("—", False),
    ("", None), ("на усмотрение производителя", None),
    # развёрнутые ответы LLM: ключевое слово внутри фразы
    ("Yes, формат ХХХ (уникальный код изготовителя)", True),
    ("Yes, формат ХХ (римские) ХХ (арабские)", True),
    ("No, не предусмотрено", False),
    # «да» не должно находиться внутри других слов («дата», «датчик»);
    # «не указана» — это отсутствие данных, а не отрицание: поле не отправляем
    ("дата изготовления не указана", None),
    ("датчик", None),
]:
    check(f"is_affirmative({raw!r})", pn.is_affirmative(raw), exp)

# ─────────────────────────────────────────────────────────────────
section("3. Справочники Controlled Impedance / Serial Number")
IMP, SER, DATE = "ufCrm24_1707840096", "ufCrm24_1707851328", "ufCrm24_1707841090"
for raw, exp in [("есть", 7204), ("Yes", 7204), ("нет", 6088), ("не требуется", 6088)]:
    check(f"импеданс {raw!r}", pn.map_to_bitrix24_ids({"impedance_control": raw,
                                                       "_normalized": {}}).get(IMP), exp)
for raw, exp in [("требуется в формате ХХХ", 6518), ("Yes", 6518), ("нет", 6516)]:
    check(f"серийный номер {raw!r}", pn.map_to_bitrix24_ids({"serial_number": raw,
                                                             "_normalized": {}}).get(SER), exp)

# ─────────────────────────────────────────────────────────────────
section("4. Date code: сторона и способ нанесения")
for data, exp in [
    ({"marking_side": "TOP"}, 6194),
    ({"marking_side": "BOTTOM"}, 6196),
    ({"marking_side": "сверху"}, 6194),
    ({"marking_side": "TOP, шелкография"}, 6206),
    ({"marking_side": "BOTTOM", "date_code": "Yes, шелкография"}, 6208),
    ({"marking_side": "TOP", "date_code": "Yes, паяльной маской"}, 6198),
    ({"marking_side": "TOP", "date_code": "Yes, медью"}, 6202),
    ({"marking_side": "нет"}, None),
    ({}, None),
]:
    label = f"{data.get('marking_side', '—')!r}" + (f" + {data['date_code']!r}" if data.get("date_code") else "")
    check(label, pn.map_to_bitrix24_ids({**data, "_normalized": {}}).get(DATE), exp)

# ─────────────────────────────────────────────────────────────────
section("5. Панель: геометрия, Board Per Panel, Order unit = pnl")
zaslon = {
    "board_name": "ИВУА.687261.013", "board_thickness": "2 ±0,3", "layer_count": 14,
    "base_material": "FR-4 Tg 170", "coverage_type": "ENIG",
    "board_size": "114 x 47", "panel_size": "134 x 74", "boards_per_panel": "1",
    "technological_fields": "Yes", "impedance_control": "есть",
    "serial_number": "требуется в формате ХХХ", "marking_side": "TOP",
    "min_hole_size": "0.2",
}
f = b.map_pcb_to_bitrix24_fields(zaslon, mistral_client=None)
check("Board Length", f.get("ufCrm24_1708353384301"), 114.0)
check("Board Width (был баг h12→12)", f.get("ufCrm24_1708353402068"), 47.0)
check("Panel Length", f.get("ufCrm24_1708375852081"), 134.0)
check("Panel Width", f.get("ufCrm24_1708375871512"), 74.0)
check("Board Per Panel", f.get("ufCrm24_1708375915545"), 1)
check("Panel Usage %", f.get("ufCrm24_1708375925847"), 54.03)
check("Order unit = pnl", f.get("ufCrm24_1707838030"), 5258)
check("Production Unit = pnl", f.get("ufCrm24_1707849863"), 6272)
check("Controlled Impedance", f.get(IMP), 7204)
check("Serial Number", f.get(SER), 6518)
check("Date code = on Top", f.get(DATE), 6194)
check("No of Layers = 14", f.get("ufCrm24_1709815185"), 6798)

section("6. Признаки поставки в панелях")
for label, args, exp in [
    ("несколько плат в панели", ({}, None, 2), True),
    ("технологические поля", ({"technological_fields": "Yes"}, None, 1), True),
    ("задан размер панели", ({}, (194.0, 74.0), 1), True),
    ("«Панелизация: Нет»", ({"panelization": "Нет"}, None, None), False),
    ("панелизации нет вовсе", ({}, None, None), False),
    ("текст «поставить в панели»", ({"panelization": "платы поставить в панели"}, None, None), True),
]:
    check(label, b._is_panel_delivery(*args), exp)

# ─────────────────────────────────────────────────────────────────
section("7. Min. Hole size — числовое поле (значение, не item_id)")
os.environ["B24_FIELD_MIN_HOLE_SIZE"] = "ufCrm24_HOLE"
import importlib  # noqa: E402
importlib.reload(b)
for raw, exp in [("0.2", 0.2), ("0,25 мм", 0.25), ("0.3 mm", 0.3)]:
    out = {}
    b._apply_optional_fields(out, {"min_hole_size": raw})
    check(f"{raw!r} → число", out.get("ufCrm24_HOLE"), exp)
out = {}
b._apply_optional_fields(out, {"min_hole_size": "0.2"})
check("тип значения", type(out.get("ufCrm24_HOLE")).__name__, "float")
os.environ.pop("B24_FIELD_MIN_HOLE_SIZE")

# ─────────────────────────────────────────────────────────────────
section("8. Живые справочники Битрикс24 (портал подменён заглушкой)")
FIELDS = {
    "ufCrm24_1707840096": {"title": "Controlled Impedance", "type": "iblock_element",
                           "settings": {"IBLOCK_ID": "300"}},
    "ufCrm24_1707838248": {"title": "Materials", "type": "iblock_element",
                           "settings": {"IBLOCK_ID": "56"}},
    "ufCrm24_DIFF": {"title": "No. of Diff Boards", "type": "double"},
}
LISTS = {
    300: [{"ID": 6088, "NAME": "No"}, {"ID": 7204, "NAME": "Yes"}],
    # на портале появился материал, которого нет в статическом словаре
    56: [{"ID": 5774, "NAME": "FR4 TG-180"}, {"ID": 9999, "NAME": "FR4 TG-155"}],
}


def fake_post(self, url, json=None, **kw):
    method = url.rsplit("/", 1)[-1]
    if method == "crm.item.fields":
        body = {"result": {"fields": FIELDS}}
    elif method == "lists.element.get":
        body = {"result": LISTS.get(int(json.get("IBLOCK_ID")), [])}
    else:
        body = {"result": []}
    return httpx.Response(200, json=body, request=httpx.Request("POST", url))


original_post = httpx.Client.post
httpx.Client.post = fake_post
os.environ["BITRIX24_WEBHOOK_URL"] = "https://x.bitrix24.ru/rest/6/TOKEN/crm.item.add"
bitrix24_api.reset_bitrix24_api()

materials = pn.get_dict("materials")
check("новое значение с портала видно", materials.get("FR4 TG-155"), 9999)
check("статические значения сохранены", materials.get("FR4 TG-180"), 5774)
check("автопоиск числового поля", b.resolve_optional_field_code("no_of_diff_boards"), "ufCrm24_DIFF")

section("9. Сбой портала не отключает поле навсегда")
httpx.Client.post = lambda self, url, json=None, **kw: (_ for _ in ()).throw(
    httpx.ConnectError("сеть недоступна"))
bitrix24_api.reset_bitrix24_api()
importlib.reload(b)
check("при сбое кода нет", b.resolve_optional_field_code("no_of_diff_boards"), "")
check("неудача НЕ закэширована", "no_of_diff_boards" in b._autodetected_codes, False)
httpx.Client.post = fake_post
bitrix24_api.reset_bitrix24_api()
check("после восстановления находит", b.resolve_optional_field_code("no_of_diff_boards"), "ufCrm24_DIFF")

httpx.Client.post = original_post
os.environ.pop("BITRIX24_WEBHOOK_URL", None)
bitrix24_api.reset_bitrix24_api()

# ─────────────────────────────────────────────────────────────────
section("10. Обязательные поля и офлайн-режим")
importlib.reload(b)
minimal = b.map_pcb_to_bitrix24_fields(
    {"board_name": "T-1", "board_thickness": "1.6", "base_material": "FR4",
     "board_size": "100 x 50"}, mistral_client=None)
check("OEM PN", minimal.get("ufCrm24_1709799376061"), "T-1")
check("Board Thickness", minimal.get("ufCrm24_1708374728464"), 1.6)
check("Order unit по умолчанию ea", minimal.get("ufCrm24_1707838030"), 5256)
try:
    b.map_pcb_to_bitrix24_fields({"board_thickness": "1.6"}, mistral_client=None)
    check("без OEM PN — ошибка", "не выброшена", "ValueError")
except ValueError:
    check("без OEM PN — ошибка", "ValueError", "ValueError")

section("11. Токен webhook маскируется в логах")
check("токен скрыт", b._mask_webhook_url("https://x.bitrix24.ru/rest/6/s3cr3t/crm.item.add"),
      "https://x.bitrix24.ru/rest/6/***/crm.item.add")

# ─────────────────────────────────────────────────────────────────
print("\n" + "=" * 64)
if FAILS:
    print(f"ПРОВАЛЕНО: {len(FAILS)}")
    for name in FAILS:
        print("  ✗", name)
else:
    print("ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ")
sys.exit(1 if FAILS else 0)
