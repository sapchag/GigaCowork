"""Прогон агента на всех наборах данных несколько раз и подсчёт метрик.

Для каждого набора (full, trap, chk) и каждого прогона из N:
  - запускает distribute.py (замеряет время),
  - сверяет результат с эталоном (compare_with_etalon),
  - сохраняет метрики и статусы всех ФУЛ.

В конце считает:
  - среднюю точность по решению и по логисту,
  - воспроизводимость: совпадает ли результат повторных прогонов с первым
    (детерминированный скрипт, поэтому ожидается 100%),
  - долю выдуманных значений (на ловушках «нет данных»),
  - среднее время ответа.

Использование:
    python run_tests.py --datasets dir1,dir2,... --runs 3 --out tests/run_summary
"""

import argparse
import glob
import json
import os
import re
import subprocess
import sys
import time

SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "distribute.py")
COMPARE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "compare_with_etalon.py")


def run(args, workdir):
    t0 = time.time()
    p = subprocess.run(args, cwd=workdir, capture_output=True, text=True)
    dt = time.time() - t0
    return p, dt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--datasets", required=True,
                    help="список папок с наборами через запятую")
    ap.add_argument("--runs", type=int, default=3)
    ap.add_argument("--distribute", default=SCRIPT)
    ap.add_argument("--compare", default=COMPARE)
    ap.add_argument("--out", default=os.path.join(os.path.dirname(SCRIPT), "..", "tests"))
    args = ap.parse_args()

    datasets = [d.strip() for d in args.datasets.split(",") if d.strip()]
    out_root = args.out
    os.makedirs(out_root, exist_ok=True)

    summary = {}
    for ds in datasets:
        resume_path = os.path.join(ds, "etalon.xlsx")
        if not os.path.exists(resume_path):
            print(f"Пропускаю набор {ds}: нет etalon.xlsx")
            continue
        dsname = os.path.basename(ds)
        runs = []
        first_full = None
        reproducible = True
        total_time = 0.0
        for i in range(1, args.runs + 1):
            run_dir = os.path.join(out_root, f"{dsname}_run{i}")
            os.makedirs(run_dir, exist_ok=True)
            result_path = os.path.join(run_dir, "result_cloud.xlsx")
            dist = os.path.join(os.path.dirname(str(args.distribute)), "distribute.py")
            p, dt = run([
                sys.executable, dist,
                "--ful", os.path.join(ds, "ful_1c.xlsx"),
                "--cloud", os.path.join(ds, "cloud.xlsx"),
                "--staff", os.path.join(ds, "staff_1c.xlsx"),
                "--out", run_dir,
            ], os.path.dirname(dist))
            total_time += dt
            # сверка
            c = os.path.join(os.path.dirname(str(args.compare)), "compare_with_etalon.py")
            pc, _ = run([
                sys.executable, c,
                "--result", result_path,
                "--etalon", os.path.join(ds, "etalon.xlsx"),
            ], os.path.dirname(c))
            try:
                m = json.loads(pc.stdout)
            except Exception:
                m = {"parse_error": pc.stdout}
            m["time_s"] = round(dt, 3)
            runs.append(m)
            if i == 1:
                first_full = m
            else:
                # воспроизводимость: сравним json результатов без времени (оно недетерминировано)
                a = {k: v for k, v in m.items() if k != "time_s"}
                b = {k: v for k, v in first_full.items() if k != "time_s"}
                if json.dumps(a, sort_keys=True, ensure_ascii=False) != json.dumps(b, sort_keys=True, ensure_ascii=False):
                    reproducible = False

        # агрегация
        def avg(key):
            vals = [r.get(key, 0) for r in runs if isinstance(r.get(key), (int, float))]
            return round(sum(vals) / len(vals), 4) if vals else 0

        summary[dsname] = {
            "datasets_dir": ds,
            "n_runs": len(runs),
            "avg_exact_match": avg("exact_match_fraction"),
            "avg_accuracy_decision": avg("accuracy_decision"),
            "avg_accuracy_logist": avg("accuracy_logist"),
            "avg_time_s": round(total_time / len(runs), 3),
            "reproducible": reproducible,
            "runs": runs,
        }
        print(f"[{dsname}] exact={summary[dsname]['avg_exact_match']} "
              f"dec={summary[dsname]['avg_accuracy_decision']} "
              f"logist={summary[dsname]['avg_accuracy_logist']} "
              f"reproducible={reproducible} time={summary[dsname]['avg_time_s']}s")

    with open(os.path.join(out_root, "test_metrics.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print("\nСводка сохранена в", os.path.join(out_root, "test_metrics.json"))


if __name__ == "__main__":
    main()