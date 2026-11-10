"""Агент распределения заказов между логистами.

Детерминированный скрипт: читает три выгрузки Excel (ФУЛ из 1С, реестр Cloud,
сотрудники из 1С), применяет правила заказчика и выдаёт результирующий Excel:
новые строки Cloud по форме реестра, повторные отправки, возможные пропуски,
«Требует внимания» и журнал решений.

Запуск:
    python distribute.py --ful <файл ФУЛ .xlsx> --cloud <реестр .xlsx> \
        --staff <сотрудники .xlsx> --out <папка результата>

Используется только стандартная библиотека Python 3 и openpyxl.
Колонки ищутся по заголовкам.
"""

import argparse
import os
import sys

import openpyxl

from gigacowork_lib import (norm_key, norm_soft, find_column, parse_datetime,
                            fmt_date, extract_order_number, determine_group,
                            role_group, is_working, match_short_name,
                            norm_supplier)

# ---------------------------------------------------------------------------
# Чтение .xlsx с поиском колонок по заголовкам
# ---------------------------------------------------------------------------

def read_sheet(path, max_col=60):
    """Прочитать первый лист книги в список словарей (название колонки -> значение)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header = None
    start = 0
    # Ищем строку заголовков: первую строку, где есть хотя бы один непустой текст
    for i, row in enumerate(rows):
        cells = [c for c in row[:max_col]]
        if any(c is not None and str(c).strip() for c in cells):
            header = [str(c).strip() if c is not None else "" for c in cells]
            start = i + 1
            break
    if header is None:
        raise ValueError(f"Не найдена строка заголовков в {path}")
    records = []
    for row in rows[start:]:
        rec = {}
        if all(c is None or str(c).strip() == "" for c in row[:max_col]):
            continue
        for idx, h in enumerate(header):
            if not h:
                continue
            val = row[idx] if idx < len(row) else None
            rec[h] = val
        records.append(rec)
    return records, header


def col(rec, *keys):
    for k in keys:
        for h, v in rec.items():
            if norm_key(h) == norm_key(k):
                return v
        if k in rec:
            return rec[k]
    return None


# ---------------------------------------------------------------------------
# Разбор входных файлов
# ---------------------------------------------------------------------------

def parse_ful(path):
    records, header = read_sheet(path)
    out = []
    for rec in records:
        order = extract_order_number(col(rec, "Комментарий"))
        out.append({
            "raw": rec,
            "fyl_num": col(rec, "Номер"),
            "date": col(rec, "Дата"),
            "cutoff": col(rec, "Cut-off", "Cutoff", "Cut off"),
            "status": col(rec, "Статус"),
            "comment": col(rec, "Комментарий"),
            "delivery": col(rec, "Способ доставки", "Способ"),
            "brand": col(rec, "Бизнес единица", "Бренд", "Бизнес единицу"),
            "logist": col(rec, "Логист"),
            "supplier": col(rec, "Контрагент", "Поставщик"),
            "country": col(rec, "Страна"),
            "order": order,
        })
    return out


def parse_cloud(path):
    records, header = read_sheet(path)
    out = []
    for rec in records:
        out.append({
            "raw": rec,
            "order": col(rec, "Заказ"),
            "supplier": col(rec, "Поставщик"),
            "brand": col(rec, "Бренд"),
            "logist": col(rec, "Логист"),
            "status": col(rec, "Статус"),
            "confirm_date": col(rec, "Дата подтверждения", "Дата подтверж"),
            "cutoff": col(rec, "Cut-off", "Cutoff", "Cut off"),
            "date_ul": col(rec, "Дата получения УЛ", "Дата получения"),
            "country": col(rec, "Страна отправления", "Страна"),
            "delivery": col(rec, "Способ доставки", "Способ"),
            "docs_date": col(rec, "Документы получены", "Документы"),
            "comment": col(rec, "Комментарий"),
        })
    return out


def parse_staff(path):
    records, header = read_sheet(path)
    out = []
    for rec in records:
        out.append({
            "raw": rec,
            "name": col(rec, "Наименование", "ФИО", "Имя"),
            "state": col(rec, "Текущее состояние", "Состояние", "Статус"),
            "role": col(rec, "Должность"),
            "code": col(rec, "Код", "Табельный", "Таб"),
        })
    return out


# ---------------------------------------------------------------------------
# Основная логика
# ---------------------------------------------------------------------------

def build_staff_index(staff):
    """Индекс сотрудников: короткое имя (из Cloud) -> сотрудник. Также карта ФИО."""
    by_full = {}
    by_short = {}
    for s in staff:
        name = norm_soft(s["name"])
        if not name:
            continue
        by_full[name] = s
        s["_norm_name"] = name
        # короткая форма: Фамилия И.
        parts = [p for p in name.split()]
        if parts:
            short = parts[0] + " " + (parts[1][0] + "." if len(parts) > 1 else "")
            by_short[short] = s
    return by_full, by_short


def resolve_logist_short(name, staff):
    """Сопоставить короткое имя/ФИО логиста из Cloud с сотрудником 1С."""
    if not name:
        return None
    n = norm_soft(name)
    for s in staff:
        if match_short_name(n, s.get("_norm_name", "")):
            return s
    return None


def compute_loader(cloud_records, staff, active_only=True):
    """Загрузка: число заказов логиста в статусе «в работе».

    Возвращает словарь: (ФИО, ключ) -> счётчик. Ключ нормализованный.
    """
    count = {}
    for c in cloud_records:
        if c["status"] and norm_soft(c["status"]) != "в работе":
            continue
        logist = c["logist"]
        if not logist:
            continue
        s = resolve_logist_short(logist, staff)
        if s:
            key = s["_norm_name"]
        else:
            key = norm_soft(logist)
        count[key] = count.get(key, 0) + 1
    return count


def supplier_history(cloud_records, staff):
    """История: сопоставление (поставщик) -> список сотрудников по частоте.

    Возвращает dict: norm_supplier -> список кортежей (count, сотрудник).
    """
    history = {}
    for c in cloud_records:
        sup = c["supplier"]
        logist = c["logist"]
        if not sup or not logist:
            continue
        s = resolve_logist_short(logist, staff)
        if not s:
            continue
        key = norm_supplier(sup)
        hist = history.setdefault(key, {})
        hist[s["_norm_name"]] = hist.get(s["_norm_name"], 0) + 1
    return {k: sorted(v.items(), key=lambda x: -x[1]) for k, v in history.items()}


def available_in_group(group, staff):
    """Список сотрудников группы, работающих и пригодных к распределению."""
    result = []
    for s in staff:
        if role_group(s["role"]) != group:
            continue
        if not is_working(s["state"]):
            continue
        result.append(s)
    return result


def select_logist_for_supplier(group, supplier, staff, cloud_records):
    """Выбор логиста в группе.

    Приоритет:
      1. Предпочтительный логист из истории по поставщику (работающий, в группе).
      2. Иначе наименее загруженный (по заказам «в работе») работающий в группе.
    Возвращает (logist, reason).
    """
    avail = available_in_group(group, staff)
    if not avail:
        return None, "в группе нет доступных логистов (отпуск/увольнение/роль)"
    # сказано: ассистенты и руководители исключены автоматически (role_group).

    chosen = None
    reason = None
    if supplier:
        hist = supplier_history(cloud_records, staff)
        key = norm_supplier(supplier)
        supp_hist = hist.get(key)
        if supp_hist:
            for name, cnt in supp_hist:
                s = next((x for x in avail if x["_norm_name"] == name), None)
                if not s:
                    continue
                chosen = s
                reason = f"предпочтительный логист по истории с поставщиком ({cnt} заказов)"
                break

    loader = compute_loader(cloud_records, staff)
    if chosen is None:
        # наименее загруженный в группе
        best = None
        best_load = None
        for s in avail:
            load = loader.get(s["_norm_name"], 0)
            if best_load is None or load < best_load:
                best_load = load
                best = s
        chosen = best
        if supplier:
            reason = "новый поставщик или нет истории — наименее загруженный в группе"
        else:
            reason = "наименее загруженный в группе"

    return chosen, reason


def process(ful, cloud, staff):
    """Основной цикл: сопоставление, повторные, новые, пропуски, внимания.

    Возвращает структуру результата.
    """
    cloud_by_order = {}
    for c in cloud:
        k = norm_key(c["order"])
        if k:
            cloud_by_order.setdefault(k, c)

    by_full, by_short = build_staff_index(staff)
    for s in staff:
        if "_norm_name" not in s:
            s["_norm_name"] = norm_soft(s["name"])

    results = []          # строки Cloud по каждому новому заказу
    repeats = []          # повторные отправки
    attention = []        # требует внимания
    journal = []          # журнал решений по каждому ФУЛ
    new_keys = set()      # номера заказов, уже созданных в этой пачке (дубль внутри пачки)
    seen_ful_in_pack = [] # (norm supplier, order, ful_date) для поиска повторов в пачке

    ful_sorted = sorted(ful, key=lambda f: parse_datetime(f["date"]) or datetime.min())
    # приоритет: минимальная дата (сначала старые, чтобы «возможные пропуски» корректно)

    # Нормализация дат
    for f in ful:
        f["_dt"] = parse_datetime(f["date"]) or datetime.min()
        f["_norm_supplier"] = norm_supplier(f["supplier"])
        f["_order_key"] = norm_key(f["order"])

    # ==================== повтор в пачке: отслеживаем по номеру ФУЛ и по заказу ====================
    # Повтор в пачке = тот же номер ФУЛ, либо тот же заказ, уже встреченный в этой пачке.
    seen_fyl_in_pack = {}   # norm_key(номер ФУЛ) -> order_key первого вхождения
    seen_order_in_pack = {} # order_key -> status (первое вхождение заказа в пачке)

    for f in ful_sorted:
        entry = {
            "ful": f,
            "order": f["order"],
            "supplier": f["supplier"],
            "brand": f["brand"],
            "delivery": f["delivery"],
            "country": f["country"],
            "cutoff": f["cutoff"],
            "date": f["date"],
            "fyl_num": f["fyl_num"],
            "dup_in_pack": False,
            "group": None,
            "logist": None,
            "reason": None,
            "status": "новый",
        }

        # 1) Нет номера заказа -> требует внимания, строку не создаём
        if not f["order"]:
            entry["status"] = "attention"
            attention.append({
                "order": None, "supplier": f["supplier"], "ful": f["fyl_num"],
                "issue": "нет номера заказа в комментарии ФУЛ",
            })
            journal.append({
                "ful": f["fyl_num"], "order": None, "supplier": f["supplier"],
                "decision": "Требует внимания", "reason": "нет номера заказа",
            })
            results.append(entry)
            continue

        order_key = f["_order_key"]

        # 2) Дубль внутри пачки по номеру ФУЛ
        fyl_key = norm_key(f["fyl_num"])
        if fyl_key and fyl_key in seen_fyl_in_pack:
            entry["status"] = "repeat_pack"
            repeats.append({
                "order": f["order"], "ful": f["fyl_num"], "supplier": f["supplier"],
                "note": "Повторная отправка ФУЛа (дубль в пачке), " + fmt_date(f["date"]),
            })
            journal.append({
                "ful": f["fyl_num"], "order": f["order"], "supplier": f["supplier"],
                "decision": "Повтор в пачке", "reason": "тот же номер ФУЛ уже есть в пачке",
            })
            results.append(entry)
            continue

        # 3) Заказ уже есть в Cloud -> повторная отправка
        if order_key in cloud_by_order:
            entry["status"] = "repeat_cloud"
            repeats.append({
                "order": f["order"], "ful": f["fyl_num"], "supplier": f["supplier"],
                "note": "Повторная отправка ФУЛа, " + fmt_date(f["date"]),
            })
            journal.append({
                "ful": f["fyl_num"], "order": f["order"], "supplier": f["supplier"],
                "decision": "Повторная отправка", "reason": "заказ уже есть в Cloud",
            })
            results.append(entry)
            continue

        # 4) Дубль внутри пачки по заказу (другой номер ФУЛ на тот же заказ)
        if order_key in seen_order_in_pack:
            entry["status"] = "repeat_pack"
            repeats.append({
                "order": f["order"], "ful": f["fyl_num"], "supplier": f["supplier"],
                "note": "Повторная отправка ФУЛа (дубль в пачке), " + fmt_date(f["date"]),
            })
            journal.append({
                "ful": f["fyl_num"], "order": f["order"], "supplier": f["supplier"],
                "decision": "Повтор в пачке",
                "reason": f"заказ {f['order']} уже обработан в пачке",
            })
            results.append(entry)
            continue

        # регистрируем заказ и номер ФУЛ как обработанные в пачке
        seen_order_in_pack[order_key] = "seen"
        if fyl_key:
            seen_fyl_in_pack[fyl_key] = order_key

        # ------ новый заказ: определяем группу и логиста ------
        group, gr_reason = determine_group(f["delivery"], f["country"])
        entry["group"] = group
        if group is None:
            entry["status"] = "attention"
            attention.append({
                "order": f["order"], "supplier": f["supplier"], "ful": f["fyl_num"],
                "issue": f"не определена группа: {gr_reason}",
            })
            journal.append({
                "ful": f["fyl_num"], "order": f["order"], "supplier": f["supplier"],
                "decision": "Требует внимания", "reason": f"не определена группа: {gr_reason}",
            })
            results.append(entry)
            continue

        logist, l_reason = select_logist_for_supplier(group, f["supplier"], staff, cloud)
        entry["logist"] = logist
        if logist is None:
            entry["status"] = "attention"
            attention.append({
                "order": f["order"], "supplier": f["supplier"], "ful": f["fyl_num"],
                "issue": f"не найден логист в группе {group}: {l_reason}",
            })
            journal.append({
                "ful": f["fyl_num"], "order": f["order"], "supplier": f["supplier"],
                "group": group,
                "decision": "Требует внимания", "reason": l_reason,
            })
            results.append(entry)
            continue

        # ------ логист выбран: создаём строку Cloud ------
        entry["status"] = "new"
        entry["reason"] = l_reason

        def short_name(s):
            parts = s["_norm_name"].split()
            return (parts[0] + " " + (parts[1][0] + "." if len(parts) > 1 else "")).title() if parts else s["name"]

        cloud_row = {
            "order": f["order"],
            "supplier": f["supplier"],
            "brand": f["brand"],
            "logist": short_name(logist),
            "status": "в работе",
            "confirm_date": "",
            "cutoff": fmt_date(f["cutoff"]),
            "date_ul": fmt_date(f["date"]),
            "country": f["country"],
            "delivery": f["delivery"],
            "docs_date": "",
            "comment": "",
        }
        entry["cloud_row"] = cloud_row
        journal.append({
            "ful": f["fyl_num"], "order": f["order"], "supplier": f["supplier"],
            "group": group, "logist": short_name(logist),
            "decision": "Новая строка",
            "reason": f"группа {group}; {l_reason}",
        })
        results.append(entry)

    # ==================== возможные пропуски ====================
    # ФУЛ, которые пришли раньше последней записи в Cloud, но заказа в Cloud нет.
    if cloud:
        last_cloud = max((parse_datetime(c["date_ul"]) or datetime.min()) for c in cloud)
    else:
        last_cloud = None
    gaps = []
    if last_cloud:
        for f in ful_sorted:
            if f["_order_key"] in cloud_by_order:
                continue
            if f["order"] and f["_dt"] and f["_dt"] < last_cloud:
                # уже обработан как новый/внимание — но пометим как возможный пропуск
                gaps.append({
                    "order": f["order"], "supplier": f["supplier"], "ful": f["fyl_num"],
                    "date": fmt_date(f["date"]), "last_cloud": fmt_date(last_cloud),
                })

    return {
        "results": results,
        "repeats": repeats,
        "gaps": gaps,
        "attention": attention,
        "journal": journal,
    }


# ---------------------------------------------------------------------------
# Выгрузка результата в Excel
# ---------------------------------------------------------------------------

CLOUD_HEADERS = [
    "Заказ", "Поставщик", "Бренд", "Логист", "Статус", "Дата подтверждения",
    "Cut-off", "Дата получения УЛ", "Страна отправления", "Способ доставки",
    "Документы получены", "Комментарий",
]


def write_result(path, data, new_rows):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Новые строки Cloud"
    ws1.append(CLOUD_HEADERS)
    for r in new_rows:
        ws1.append([
            r.get("order", ""), r.get("supplier", ""), r.get("brand", ""),
            r.get("logist", ""), r.get("status", ""), r.get("confirm_date", ""),
            r.get("cutoff", ""), r.get("date_ul", ""), r.get("country", ""),
            r.get("delivery", ""), r.get("docs_date", ""), r.get("comment", ""),
        ])

    ws2 = wb.create_sheet("Повторные отправки")
    ws2.append(["Заказ", "Номер ФУЛ", "Поставщик", "Комментарий"])
    for r in data["repeats"]:
        ws2.append([r.get("order"), r.get("ful"), r.get("supplier"), r.get("note")])

    ws3 = wb.create_sheet("Возможные пропуски")
    ws3.append(["Заказ", "Поставщик", "Номер ФУЛ", "Дата ФУЛ", "Последняя запись Cloud"])
    for r in data["gaps"]:
        ws3.append([r.get("order"), r.get("supplier"), r.get("ful"),
                    r.get("date"), r.get("last_cloud")])

    ws4 = wb.create_sheet("Требует внимания")
    ws4.append(["Заказ", "Поставщик", "Номер ФУЛ", "Причина"])
    for r in data["attention"]:
        ws4.append([r.get("order"), r.get("supplier"), r.get("ful"), r.get("issue")])

    ws5 = wb.create_sheet("Журнал решений")
    ws5.append(["Номер ФУЛ", "Заказ", "Поставщик", "Группа", "Логист",
                "Решение", "Обоснование"])
    for r in data["journal"]:
        ws5.append([r.get("ful"), r.get("order"), r.get("supplier"), r.get("group"),
                    r.get("logist"), r.get("decision"), r.get("reason")])

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    wb.close()
    return path


def summarize(data):
    new_count = sum(1 for r in data["results"] if r["status"] == "new")
    att_count = len(data["attention"])
    # повторные отправки: отдельно повторы из Cloud и из пачки
    rep_cloud = sum(1 for r in data["repeats"] if "Cloud" in r.get("note", ""))
    rep_pack = len(data["repeats"]) - rep_cloud
    return {
        "new": new_count,
        "repeats": len(data["repeats"]),
        "rep_cloud": rep_cloud,
        "rep_pack": rep_pack,
        "attention": att_count,
        "gaps": len(data["gaps"]),
    }


# ---------------------------------------------------------------------------

def run(ful_path, cloud_path, staff_path, out_dir):
    ful = parse_ful(ful_path)
    cloud = parse_cloud(cloud_path)
    staff = parse_staff(staff_path)

    if not ful:
        return "Ошибка: файл ФУЛ пуст или не содержит строк", {}
    if not staff:
        return "Ошибка: файл сотрудников пуст или не содержит строк", {}

    data = process(ful, cloud, staff)
    new_rows = [r["cloud_row"] for r in data["results"] if r["status"] == "new"
                and r.get("cloud_row")]
    out_path = os.path.join(out_dir, "result_cloud.xlsx")
    write_result(out_path, data, new_rows)
    summary = summarize(data)
    return out_path, summary


def main(argv=None):
    ap = argparse.ArgumentParser(description="Агент распределения заказов между логистами")
    ap.add_argument("--ful", required=True, help="Выгрузка ФУЛ из 1С (.xlsx)")
    ap.add_argument("--cloud", required=True, help="Реестр заказов Cloud (.xlsx)")
    ap.add_argument("--staff", required=True, help="Список сотрудников из 1С (.xlsx)")
    ap.add_argument("--out", default=".", help="Папка для результата")
    args = ap.parse_args(argv)

    if not os.path.isdir(args.out):
        os.makedirs(args.out, exist_ok=True)

    import json
    result, summary = run(args.ful, args.cloud, args.staff, args.out)
    if isinstance(result, str) and result.startswith("Ошибка"):
        print(result)
        return 1
    print("Файл результата:", result)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())