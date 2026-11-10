"""Сравнение результата агента с эталоном и подсчёт метрик.

Эталон: etalon.xlsx, ключ строки — номер ФУЛ (уникален, допускает дубли заказа
в пачке). Результат агента: result_cloud.xlsx, лист «Журнал решений» (по каждому
ФУЛ: номер, заказ, группа, логист, решение, обоснование) + листы «Новые строки
Cloud», «Повторные отправки», «Требует внимания», «Возможные пропуски».

Сверка идёт по номеру ФУЛ: решение (new/repeat_cloud/repeat_pack/attention),
для новых — группа и логист.

Используется стандартная библиотека + openpyxl.
"""

import argparse
import json
import os

import openpyxl


def read_rows(path, sheet=None):
    wb = openpyxl.load_workbook(path, data_only=True)
    target = sheet if sheet else wb.sheetnames[0]
    ws = wb[target]
    header = [c.value for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if any(v is not None and str(v).strip() for v in r):
            rows.append(dict(zip(header, r)))
    wb.close()
    return rows


def norm(v):
    if v is None:
        return ""
    return "".join(ch.lower() for ch in str(v) if ch.isalnum())


def key_of_ful_num(cell):
    """Нормализованный ключ номера ФУЛ."""
    return norm(str(cell))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--result", required=True, help="result_cloud.xlsx")
    ap.add_argument("--etalon", required=True, help="etalon.xlsx")
    ap.add_argument("--out", default=None, help="Файл журнала метрик (.json)")
    args = ap.parse_args()

    etalon_rows = read_rows(args.etalon, "Эталон")
    journal = read_rows(args.result, "Журнал решений")
    new_rows = read_rows(args.result, "Новые строки Cloud")
    repl_rows = read_rows(args.result, "Повторные отправки")
    attn_rows = read_rows(args.result, "Требует внимания")
    gap_rows = read_rows(args.result, "Возможные пропуски")

    # журнал агента по номеру ФУЛ
    journal_by_ful = {}
    for r in journal:
        f = r.get("Номер ФУЛ")
        if f:
            journal_by_ful[key_of_ful_num(f)] = r

    # новые строки по заказу (для сверки группы/логиста)
    new_by_order = {}
    for r in new_rows:
        o = r.get("Заказ")
        if o:
            new_by_order[norm(str(o))] = r

    # повторные по заказу
    repl_orders = {norm(str(r.get("Заказ"))) for r in repl_rows}

    # attention: сопоставим по номеру ФУЛ из листа (колонка «Номер ФУЛ»)
    attn_by_ful = {}
    for r in attn_rows:
        f = r.get("Номер ФУЛ")
        if f:
            attn_by_ful[key_of_ful_num(f)] = r

    # возможные пропуски — по заказу
    gap_orders = {norm(str(r.get("Заказ"))) for r in gap_rows}

    # ожидаемые группы/логисты для ловушки «вся группа недоступна» (мультимодальные)
    GROUP_OK_EVEN_IF_NO_LOGIST = ("мультимодальные",)

    results = []
    mismatch = 0
    decision_acc_n = 0
    decision_acc_ok = 0
    logist_acc_n = 0
    logist_acc_ok = 0

    def map_got_status(exp):
        for r in journal:
            pass
        return None

    for row in etalon_rows:
        key = row.get("Ключ(Заказ/ФУЛ)") or row.get("ФУЛ/Заказ")
        if key is None:
            continue
        exp_dec = (row.get("Решение") or "").strip()
        exp_group = row.get("Группа") or ""
        exp_logist = row.get("Логист") or ""
        exp_reason = row.get("Обоснование") or ""
        kn = key_of_ful_num(key)

        j = journal_by_ful.get(kn)
        got_dec = (j.get("Решение") if j else "").strip() if j else ""
        got_group = (j.get("Группа") or "") if j else ""
        got_logist = (j.get("Логист") or "") if j else ""

        def normalize(dec):
            d = norm(dec)
            if "пачк" in d:
                return "repeat_pack"
            if "повтор" in d or "cloud" in d:
                return "repeat_cloud"
            if "вниман" in d:
                return "attention"
            if "нов" in d:
                return "new"
            return norm(dec.replace("_", ""))

        gdec = normalize(got_dec)

        # определение верности решения
        exp_kind = norm(exp_dec.replace("_", ""))
        if exp_kind in ("new+gap", "newgap"):
            expected_status = "new"
        else:
            expected_status = exp_kind

        got_status = norm(gdec).replace("_", "")

        if exp_kind in ("new+gap", "newgap"):
            # заказ из журнала
            jorder = norm(str(j.get("Заказ") or "")) if j else ""
            in_gap = jorder in gap_orders
            ok_decision = (got_status == "new") and in_gap
        else:
            ok_decision = (got_status == expected_status)

        decision_acc_n += 1
        if ok_decision:
            decision_acc_ok += 1

        # проверка логиста для new
        log_ok = True
        if expected_status == "new" and exp_logist:
            logist_acc_n += 1
            if norm(got_logist) == norm(exp_logist):
                logist_acc_ok += 1
            else:
                log_ok = False
        elif expected_status == "attention":
            # ловушка «вся группа недоступна»: группа есть, логист должен быть пусто
            if exp_group in GROUP_OK_EVEN_IF_NO_LOGIST:
                log_ok = (not got_logist)

        full_ok = ok_decision and log_ok
        if not full_ok:
            mismatch += 1

        results.append({
            "key": str(key), "expected": exp_dec, "got_decision": gdec,
            "exp_group": exp_group, "got_group": got_group,
            "exp_logist": exp_logist, "got_logist": got_logist,
            "field_ok": full_ok, "exp_reason": exp_reason,
        })

    total = len(etalon_rows)
    exact = (total - mismatch) / total if total else 0
    acc_dec = decision_acc_ok / decision_acc_n if decision_acc_n else 1
    acc_log = logist_acc_ok / logist_acc_n if logist_acc_n else 1

    out = {
        "total_ful": total,
        "exact_match_fraction": round(exact, 4),
        "accuracy_decision": round(acc_dec, 4),
        "accuracy_logist": round(acc_log, 4),
        "mismatches": [r for r in results if not r["field_ok"]],
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))
    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()