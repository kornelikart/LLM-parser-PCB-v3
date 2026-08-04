# -*- coding: utf-8 -*-
"""
Оффлайн-проверка извлечения текста из спецификаций (без LLM и API-ключа).

Для каждого файла eval/data/<имя> с ожиданиями eval/data/<имя>.expected.json:
  - извлекает текст через app.utils.extract_document_data;
  - проверяет extract.must_contain (ключевые значения должны попасть в текст)
    и extract.must_not_contain (шум: скрытые колонки, листы-примеры, списки меню);
  - следит, чтобы текст укладывался в бюджет LLM (MAX_EXTRACT_CHARS).

Извлечённый текст каждого файла сохраняется в eval/results/extraction/
для ручного просмотра.

Запуск из корня проекта:
    python eval/run_extraction_eval.py [-v]

Код возврата 0 — все проверки пройдены.
"""
import argparse
import json
import sys
from pathlib import Path

EVAL_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = EVAL_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app.utils import extract_document_data, MAX_EXTRACT_CHARS  # noqa: E402

DATA_DIR = EVAL_DIR / "data"
RESULTS_DIR = EVAL_DIR / "results" / "extraction"
SPEC_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".txt"}


def iter_cases():
    if not DATA_DIR.is_dir():
        print(f"Нет директории с данными: {DATA_DIR}\n"
              f"Положите файлы спецификаций и <имя>.expected.json в eval/data/.")
        sys.exit(2)
    for path in sorted(DATA_DIR.iterdir(), key=lambda p: p.name.lower()):
        if path.suffix.lower() not in SPEC_SUFFIXES:
            continue
        expected_path = path.with_name(path.name + ".expected.json")
        expected = {}
        if expected_path.exists():
            expected = json.loads(expected_path.read_text(encoding="utf-8"))
        yield path, expected


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="печатать все проверки, а не только проваленные")
    args = parser.parse_args()

    RESULTS_DIR.mkdir(parents=True, exist_ok=True)

    failed_files = 0
    total_checks = 0
    total_failed = 0

    for path, expected in iter_cases():
        try:
            text = extract_document_data(str(path))
        except Exception as e:
            print(f"❌ {path.name}: ОШИБКА извлечения: {e}")
            failed_files += 1
            total_failed += 1
            continue

        (RESULTS_DIR / (path.name + ".extracted.txt")).write_text(text, encoding="utf-8")

        problems = []
        passed = []
        low = text.casefold()
        ex = expected.get("extract", {})
        for needle in ex.get("must_contain", []):
            total_checks += 1
            if needle.casefold() in low:
                passed.append(f"содержит {needle!r}")
            else:
                problems.append(f"НЕТ ожидаемой подстроки: {needle!r}")
        for needle in ex.get("must_not_contain", []):
            total_checks += 1
            if needle.casefold() in low:
                problems.append(f"ШУМ в тексте: {needle!r}")
            else:
                passed.append(f"нет шума {needle!r}")
        total_checks += 1
        if len(text) > MAX_EXTRACT_CHARS + 100:  # +маркер обрезки
            problems.append(f"текст больше бюджета: {len(text)} символов")
        else:
            passed.append(f"размер {len(text)} симв.")

        total_failed += len(problems)
        status = "✅" if not problems else "❌"
        if problems:
            failed_files += 1
        print(f"{status} {path.name}  ({len(text)} симв.)")
        for p in problems:
            print(f"     ✗ {p}")
        if args.verbose:
            for p in passed:
                print(f"     ✓ {p}")

    print(f"\nИтого: проверок {total_checks}, провалено {total_failed}, "
          f"файлов с проблемами {failed_files}.")
    print(f"Извлечённые тексты: {RESULTS_DIR}")
    return 1 if total_failed else 0


if __name__ == "__main__":
    sys.exit(main())
