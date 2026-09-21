# -*- coding: utf-8 -*-
"""
Офлайн-скоринг сохранённых ответов LLM (eval/results/<прогон>/*.actual.json)
по эталонам eval/data/*.expected.json — тем же матчером, что в run_llm_eval.py.

Позволяет сравнивать модели без повторных запросов к API:
    python eval/score_results.py eval/results/llm_20260805_141127 eval/results/llm_20260921_130000

Секции: fields (Промпт 1), normalized (Промпт 2), bitrix_fields (маппинг).
Проверка bitrix_required_ok в сохранённых ответах не восстанавливается и не считается.
"""
import json
import sys
from pathlib import Path

EVAL_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(EVAL_DIR))

from run_llm_eval import DATA_DIR, check_matcher  # noqa: E402


def score(run_dir: Path, verbose: bool = False) -> dict:
    hard_total = hard_failed = soft_failed = files = missing = 0
    failed_labels = []
    for expected_path in sorted(DATA_DIR.glob("*.expected.json"), key=lambda p: p.name.lower()):
        name = expected_path.name[: -len(".expected.json")]
        actual_path = run_dir / (name + ".actual.json")
        if not actual_path.exists():
            missing += 1
            continue
        files += 1
        expected = json.loads(expected_path.read_text(encoding="utf-8"))
        actual = json.loads(actual_path.read_text(encoding="utf-8"))
        parsed = actual.get("parsed") or {}
        norm = actual.get("normalized") or {}
        b24 = actual.get("bitrix_fields") or {}

        checks = []
        for field, matcher in expected.get("fields", {}).items():
            checks.append((f"fields.{field}", bool(matcher.get("soft")),
                           check_matcher(parsed.get(field), matcher), parsed.get(field)))
        for key, exp in expected.get("normalized", {}).items():
            soft = isinstance(exp, dict) and bool(exp.get("soft"))
            val = exp.get("value") if isinstance(exp, dict) else exp
            got = norm.get(key)
            ok = (got == val) if val is not None else got in (None, "")
            checks.append((f"normalized.{key}", soft, ok, got))
        for code, exp in expected.get("bitrix_fields", {}).items():
            soft = isinstance(exp, dict) and bool(exp.get("soft"))
            want = exp.get("value") if isinstance(exp, dict) else exp
            got = b24.get(code)
            ok = (got == want) if want is not None else got is None
            checks.append((f"bitrix.{code}", soft, ok, got))

        for label, soft, ok, got in checks:
            if soft:
                soft_failed += 0 if ok else 1
            else:
                hard_total += 1
                hard_failed += 0 if ok else 1
                if not ok:
                    failed_labels.append(f"{name}: {label} = {got!r}")
        if verbose:
            bad = [l for l, s, ok, _ in checks if not ok and not s]
            print(f"  {name}: жёстких провалов {len(bad)}" + (f" — {', '.join(bad)}" if bad else ""))

    return {"files": files, "missing": missing, "hard_total": hard_total,
            "hard_failed": hard_failed, "soft_failed": soft_failed, "failed": failed_labels}


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    verbose = "-v" in sys.argv
    if not args:
        print(__doc__)
        return 2
    for arg in args:
        run_dir = Path(arg)
        if not run_dir.is_absolute():
            run_dir = (EVAL_DIR.parent / run_dir).resolve()
        r = score(run_dir, verbose)
        passed = r["hard_total"] - r["hard_failed"]
        pct = 100.0 * passed / r["hard_total"] if r["hard_total"] else 0.0
        print(f"{run_dir.name}: файлов {r['files']} (нет ответа: {r['missing']}), "
              f"жёстких проверок {r['hard_total']}, пройдено {passed} ({pct:.0f}%), "
              f"мягких провалов {r['soft_failed']}")
        if verbose:
            for line in r["failed"]:
                print("    ✗", line)
    return 0


if __name__ == "__main__":
    sys.exit(main())
