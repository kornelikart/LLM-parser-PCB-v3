# -*- coding: utf-8 -*-
"""
Полный прогон LLM-пайплайна на eval-наборе. Требует MISTRAL_API_KEY (env или .env).

Для каждого файла eval/data/<имя> с ожиданиями <имя>.expected.json:
  1. извлечение текста (extract_document_data);
  2. Промпт 1 (structured output) → проверка секции "fields";
  3. Промпт 2 (нормализация) → проверка секции "normalized";
  4. map_pcb_to_bitrix24_fields без LLM → проверка "bitrix_required_ok"
     (валидность обязательных полей board_name / board_thickness).

Матчеры в "fields":
  {"equals": x} | {"contains": "s"} | {"contains_any": [...]} |
  {"contains_all": [...]} | {"not_empty": true}; + "soft": true — проверка
  информационная, в код возврата не входит.
В "normalized" значение — строка (точное совпадение), null (ожидаем отсутствие)
или {"value": ..., "soft": true}.

Запуск из корня проекта:
    python eval/run_llm_eval.py [--only ПОДСТРОКА_ИМЕНИ] [--sleep 2.0]

Фактические ответы LLM сохраняются в eval/results/llm_<метка времени>/.
"""
import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path

EVAL_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = EVAL_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app import bitrix24, utils  # noqa: E402
from app.config import mistral_params  # noqa: E402
from app.pcb_normalizer import normalize_pcb_data  # noqa: E402

DATA_DIR = EVAL_DIR / "data"
SPEC_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".txt"}


def iter_cases(only: str):
    for path in sorted(DATA_DIR.iterdir(), key=lambda p: p.name.lower()):
        if path.suffix.lower() not in SPEC_SUFFIXES:
            continue
        if only and only.lower() not in path.name.lower():
            continue
        expected_path = path.with_name(path.name + ".expected.json")
        if not expected_path.exists():
            print(f"⚠ {path.name}: нет .expected.json — пропущен")
            continue
        yield path, json.loads(expected_path.read_text(encoding="utf-8"))


def check_matcher(actual, matcher: dict) -> bool:
    a = "" if actual is None else str(actual).strip()
    al = a.casefold()
    if "equals" in matcher:
        exp = matcher["equals"]
        if isinstance(exp, (int, float)) and not isinstance(exp, bool):
            try:
                return float(a) == float(exp)
            except ValueError:
                return False
        return al == str(exp).strip().casefold()
    if "contains" in matcher:
        return str(matcher["contains"]).casefold() in al
    if "contains_any" in matcher:
        return any(str(x).casefold() in al for x in matcher["contains_any"])
    if "contains_all" in matcher:
        return all(str(x).casefold() in al for x in matcher["contains_all"])
    if "not_empty" in matcher:
        return bool(a) == bool(matcher["not_empty"])
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="LLM-eval пайплайна распознавания ПП")
    parser.add_argument("--only", default="", help="фильтр по подстроке имени файла")
    parser.add_argument("--sleep", type=float, default=2.0,
                        help="пауза между файлами, сек (бережём rate limit)")
    args = parser.parse_args()

    results_dir = EVAL_DIR / "results" / f"llm_{datetime.now():%Y%m%d_%H%M%S}"
    results_dir.mkdir(parents=True, exist_ok=True)

    llm = utils.create_pcb_model(mistral_params)
    adapter = bitrix24._MistralChatAdapter(llm)

    hard_failed = 0
    soft_failed = 0
    hard_total = 0

    for path, expected in iter_cases(args.only):
        print(f"\n══ {path.name}")
        try:
            text = utils.extract_document_data(str(path))
            parsed = utils.process_excel_pcb_with_retry(text, llm)
        except Exception as e:
            print(f"  ❌ пайплайн упал: {e}")
            hard_failed += 1
            hard_total += 1
            continue

        # ── fields (Промпт 1) ──
        for field, matcher in expected.get("fields", {}).items():
            soft = bool(matcher.get("soft"))
            ok = check_matcher(parsed.get(field), matcher)
            mark = "✓" if ok else ("⚠" if soft else "✗")
            if not soft:
                hard_total += 1
                hard_failed += 0 if ok else 1
            elif not ok:
                soft_failed += 1
            shown = {k: v for k, v in matcher.items() if k != "soft"}
            print(f"  {mark} fields.{field} = {parsed.get(field)!r}  ожидание {shown}")

        # ── normalized (Промпт 2) ──
        norm = {}
        norm_expected = expected.get("normalized", {})
        if norm_expected:
            try:
                enriched = normalize_pcb_data(parsed, adapter)
                norm = enriched.get("_normalized", {})
            except Exception as e:
                print(f"  ❌ нормализация упала: {e}")
            for key, exp in norm_expected.items():
                soft = isinstance(exp, dict) and bool(exp.get("soft"))
                val = exp.get("value") if isinstance(exp, dict) else exp
                actual = norm.get(key)
                ok = (actual == val) if val is not None else actual in (None, "")
                mark = "✓" if ok else ("⚠" if soft else "✗")
                if not soft:
                    hard_total += 1
                    hard_failed += 0 if ok else 1
                elif not ok:
                    soft_failed += 1
                print(f"  {mark} normalized.{key} = {actual!r}  ожидание {val!r}")

        # ── поля Битрикс24 (без LLM, fallback-справочники) ──
        b24_fields = {}
        if "bitrix_required_ok" in expected or "bitrix_fields" in expected:
            try:
                b24_fields = bitrix24.map_pcb_to_bitrix24_fields(dict(parsed), mistral_client=None)
                b24_ok, b24_err = True, ""
            except Exception as e:
                b24_ok, b24_err = False, str(e)

            if "bitrix_required_ok" in expected:
                ok = b24_ok == bool(expected["bitrix_required_ok"])
                hard_total += 1
                hard_failed += 0 if ok else 1
                mark = "✓" if ok else "✗"
                print(f"  {mark} bitrix_required_ok = {b24_ok} {('(' + b24_err + ')') if b24_err else ''}")

            # Проверка конкретных значений полей — ловит ошибки маппинга
            # (например, допуск h12, попавший в «ширину платы»).
            for code, exp in expected.get("bitrix_fields", {}).items():
                soft = isinstance(exp, dict) and bool(exp.get("soft"))
                want = exp.get("value") if isinstance(exp, dict) else exp
                actual = b24_fields.get(code)
                ok = (actual == want) if want is not None else actual is None
                mark = "✓" if ok else ("⚠" if soft else "✗")
                if not soft:
                    hard_total += 1
                    hard_failed += 0 if ok else 1
                elif not ok:
                    soft_failed += 1
                print(f"  {mark} bitrix.{code} = {actual!r}  ожидание {want!r}")

        out = {"parsed": parsed, "normalized": norm, "bitrix_fields": b24_fields}
        (results_dir / (path.name + ".actual.json")).write_text(
            json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        time.sleep(args.sleep)

    print(f"\nИтого: жёстких проверок {hard_total}, провалено {hard_failed}; "
          f"мягких провалов {soft_failed}.")
    print(f"Фактические ответы: {results_dir}")
    return 1 if hard_failed else 0


if __name__ == "__main__":
    sys.exit(main())
