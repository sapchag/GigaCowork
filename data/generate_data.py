"""Генератор синтетических данных для агента распределения заказов между логистами.

Создаёт три файла Excel (ФУЛ из 1С, реестр Cloud, сотрудники из 1С) с ловушками
и эталон правильного результата по каждому ФУЛ. Режимы:
  --mode full   — полный функциональный набор (все группы распределяются нормально)
  --mode trap   — набор-ловушка <вся группа недоступна>
  --mode chk    — тот же полный набор с другим сидом (проверка воспроизводимости)

Используется стандартная библиотека Python 3 и openpyxl. Все компании, бренды
и сотрудники вымышленные.

Запуск:
    python generate_data.py --out <папка> --mode full --seed 42
"""

import argparse
import os
import random
from collections import Counter
from datetime import datetime, timedelta

import openpyxl

# ---------------------------------------------------------------------------
# Сотрудники (вымышленные) — (ФИО, должность, состояние, код)
# ---------------------------------------------------------------------------

STAFF_BASE = [
    ("Виноградов Сергей Андреевич", "специалист по авиаперевозкам", "Работает", "A001"),
    ("Орлова Марина Викторовна",    "специалист по авиаперевозкам", "Работает", "A002"),
    ("Киселёв Дмитрий Павлович",    "специалист по авиаперевозкам", "Отпуск",   "A003"),
    ("Белова Наталья Игоревна",     "специалист по работе с ВЭД документацией", "Работает", "B001"),
    ("Смирнов Андрей Сергеевич",    "специалист по работе с ВЭД документацией", "Работает", "B002"),
    ("Тихонов Олег Николаевич",     "специалист по работе с ВЭД документацией", "Уволен",   "B003"),
    ("Зайцев Роман Владимирович",   "специалист по мультимодальным перевозкам", "Работает", "C001"),
    ("Морозова Алина Дмитриевна",   "специалист по мультимодальным перевозкам", "Работает", "C002"),
    ("Ефимов Григорий Алексеевич",  "специалист по мультимодальным перевозкам", "Отпуск по уходу за ребенком", "C003"),
    ("Громова Елена Станиславовна", "ассистент специалиста по авиаперевозкам", "Работает", "A900"),
    ("Соловьёв Игорь Максимович",   "руководитель группы ВЭД", "Работает", "B900"),
]

# Режим trap: все мультимодальные недоступны
STAFF_TRAP = [
    ("Виноградов Сергей Андреевич", "специалист по авиаперевозкам", "Работает", "A001"),
    ("Орлова Марина Викторовна",    "специалист по авиаперевозкам", "Работает", "A002"),
    ("Белова Наталья Игоревна",     "специалист по работе с ВЭД документацией", "Работает", "B001"),
    ("Зайцев Роман Владимирович",   "специалист по мультимодальным перевозкам", "Отпуск", "C001"),
    ("Морозова Алина Дмитриевна",   "специалист по мультимодальным перевозкам", "Отпуск", "C002"),
    ("Ефимов Григорий Алексеевич",  "специалист по мультимодальным перевозкам", "Уволен", "C003"),
    ("Громова Елена Станиславовна", "ассистент специалиста по авиаперевозкам", "Работает", "A900"),
]


def staff_short(name):
    parts = name.split()
    return parts[0] + " " + parts[1][0] + "."


def norm_sup(s):
    if not s:
        return ""
    return "".join(ch.lower() for ch in str(s) if ch.isalnum())


# Поставщики и варианты написания одного и того же
SUPPLIERS = [
    "Shenzhen Fashion Co.",
    "Гуанчжоу Текстиль Лтд",
    "Shanghai Apparel Group",
    "Ost-West Bekleidung GmbH",
    "Delhi Cotton House",
]
# намеренные вариации написания, которые нормализация должна свести к одной сущности
SUPPLIER_VARIANTS = {
    "Shenzhen Fashion Co.": ["Shenzhen Fashion Co.", "SHENZHEN FASHION CO.", "Shenzhen FashionCo"],
    "Гуанчжоу Текстиль Лтд": ["Гуанчжоу Текстиль Лтд", "  гуанчжоу текстиль лтд  ", "Гуанчжоу-Текстиль Лтд"],
}

BRANDS = ["Zara", "H&M", "Пальмира", "Детский Мир", "Северное Сияние", "Moda Rus"]

ROLE_BY_GROUP = {
    "авиаперевозки": "специалист по авиаперевозкам",
    "вэд": "специалист по работе с ВЭД документацией",
    "мультимодальные": "специалист по мультимодальным перевозкам",
}


def daterange(start, days):
    return start + timedelta(days=days)


def variant(base):
    """Вернуть вариант написания поставщика (смена регистра/пробелов/знаков)."""
    vs = SUPPLIER_VARIANTS.get(base, [base])
    return random.choice(vs)


# ---------------------------------------------------------------------------
# Построение реестра Cloud с историей по поставщикам
# ---------------------------------------------------------------------------

def build_cloud(seed):
    rng = random.Random(seed)
    start = datetime(2026, 5, 1)
    rows = []
    last_ul = start

    def add(order, supplier, brand, logist, status, date_ul, country, delivery,
            cutoff_days=12):
        nonlocal last_ul
        row = {
            "order": order, "supplier": supplier, "brand": brand, "logist": logist,
            "status": status, "date_ul": date_ul, "country": country,
            "delivery": delivery, "cutoff": date_ul + timedelta(days=cutoff_days),
            "comment": "",
        }
        rows.append(row)
        if date_ul > last_ul:
            last_ul = date_ul
        return row

    V = staff_short("Виноградов Сергей Андреевич")
    O = staff_short("Орлова Марина Викторовна")
    B = staff_short("Белова Наталья Игоревна")
    S = staff_short("Смирнов Андрей Сергеевич")
    Z = staff_short("Зайцев Роман Владимирович")
    M = staff_short("Морозова Алина Дмитриевна")

    # --- авиа: история Виноградов (Shenzhen), Орлова (Shanghai) ---
    for i in range(6):
        add("S0005" + "{:02d}".format(1 + i), "Shenzhen Fashion Co.", rng.choice(BRANDS), V,
            "в работе", daterange(start, i * 2), "Китай", "air")
    for i in range(3):
        add("S0005021", "Shenzhen Fashion Co.", rng.choice(BRANDS), O,
            "confirm", daterange(start, 2 + i), "Китай", "air")
    for i in range(3):
        add("S0005" + str(31 + i), "Shanghai Apparel Group", rng.choice(BRANDS), O,
            "в работе", daterange(start, 3 + i), "Китай", "air")

    # --- ВЭД: Белова чаще (Гуанчжоу) ---
    for i in range(5):
        add("S0005" + str(41 + i), "Гуанчжоу Текстиль Лтд", rng.choice(BRANDS),
            B if i % 2 == 0 else S, "в работе", daterange(start, i + 1), "Китай", "sea")
    add("S000550", "China Trading New", rng.choice(BRANDS), B, "confirm",
        daterange(start, 1), "Китай", "sea")

    # --- мультимодальные: Морозова (Delhi), Зайцев (Ost-West) ---
    for i in range(3):
        add("S0005" + str(61 + i), "Ost-West Bekleidung GmbH", rng.choice(BRANDS),
            Z, "в работе", daterange(start, 2 + i), "Германия", "truck")
    for i in range(4):
        add("S0005" + str(71 + i), "Delhi Cotton House", rng.choice(BRANDS), M,
            "в работе", daterange(start, i), "Индия", "sea")
    add("S000580", "Madrid Textiles", rng.choice(BRANDS), M, "confirm",
        daterange(start, 1), "Испания", "railway")

    # --- спец-случаи для ловушек ---
    # поставщик Air Moscow Cargo ведёт (в историю) Киселёв, который сейчас В ОТПУСКЕ
    add("S000590", "Air Moscow Cargo", rng.choice(BRANDS),
        staff_short("Киселёв Дмитрий Павлович"), "в работе",
        daterange(start, 4), "Россия", "air")
    # поставщик Rakuten Textiles ведёт Тихонов, который УВОЛЕН
    add("S000591", "Rakuten Textiles", rng.choice(BRANDS),
        staff_short("Тихонов Олег Николаевич"), "confirm",
        daterange(start, 3), "Китай", "sea")

    # заказ S000501 в Cloud (для повторной отправки)
    add("S000501", "Shenzhen Fashion Co.", rng.choice(BRANDS), V, "в работе",
        daterange(start, 2), "Китай", "air")

    return rows, last_ul


# ---------------------------------------------------------------------------
# Загрузка по заказам «в работе» (как считает агент)
# ---------------------------------------------------------------------------

def load_by_short(cloud, staff):
    """Занятость логиста (короткое имя) по заказам «в работе»."""
    cnt = Counter()
    for c in cloud:
        if c["status"] == "в работе" and c["logist"]:
            cnt[c["logist"]] += 1
    return cnt


def deciding_logist(group, cloud, staff, supplier):
    """Тот же выбор, что делает агент: предпочтительный по истории, иначе наименее загруженный."""
    # работники группы
    role = ROLE_BY_GROUP[group]
    avail = [s for s in staff
             if s[1] == role and s[2] == "Работает"]
    if not avail:
        return None, "в группе нет доступных логистов"
    # история по поставщику
    hist = Counter()
    sup = norm_sup(supplier)
    for c in cloud:
        if norm_sup(c["supplier"]) == sup and c["logist"]:
            # короткое имя -> ФИО
            for s in avail:
                if staff_short(s[0]) == c["logist"]:
                    hist[staff_short(s[0])] += 1
    pref_short = None
    if hist:
        pref_short = max(hist, key=hist.get)
        for s in avail:
            if staff_short(s[0]) == pref_short:
                return staff_short(s[0]), f"предпочтительный логист по истории ({hist[pref_short]} заказов)"
    # иначе наименее загруженный (по «в работе»), ничьи — по алфавиту для детерминизма
    load = load_by_short(cloud, staff)
    cands = sorted(avail, key=lambda s: (load.get(staff_short(s[0]), 0), staff_short(s[0])))
    best = cands[0]
    return staff_short(best[0]), f"новый поставщик/нет истории — наименее загруженный в группе"


def rule_group(delivery, country):
    if not delivery:
        return None, "нет способа доставки"
    if delivery.strip().lower() == "air":
        return "авиаперевозки", None
    if not country:
        return None, "нет страны отправления"
    if any(k in country for k in ("Китай", "кита", "Гонконг")):
        return "вэд", None
    return "мультимодальные", None


# ---------------------------------------------------------------------------
# Полный функциональный набор
# ---------------------------------------------------------------------------

def build_full(rng, seed):
    staff = STAFF_BASE[:]
    cloud, last_ul = build_cloud(seed)
    FUL_START = last_ul + timedelta(days=1)  # новые ФУЛ после последней записи Cloud

    def make_ful(order, supplier, brand, delivery, country, day, num):
        date = FUL_START + timedelta(days=day)
        comment = f"Номер: {order} от {date.strftime('%d.%m.%Y')}" if order else "Упаковочный лист без номера заказа"
        return {
            "Дата": date, "Номер": f"{num}-{order}" if order else num,
            "Статус": "Отправлен", "Комментарий": comment,
            "Способ доставки": delivery, "Бизнес единица": brand,
            "Логист": "", "Контрагент": supplier, "Страна": country,
        }, order, comment

    ful = []
    etalon = {}

    def new_f(order, supplier, brand, delivery, country, day=0, num="УЛ"):
        f, o, c = make_ful(order, variant(supplier), brand, delivery, country, day, num)
        # группа и логист по правилам (детерминированно)
        grp, r = rule_group(delivery, country)
        if grp:
            log, why = deciding_logist(grp, cloud, staff, variant(supplier))
        else:
            log, why = None, r
        ful.append(f)
        # ключ эталона — номер ФУЛ (уникален, допускает дубли заказа в пачке)
        etalon[f["Номер"]] = {"decision": "new" if log else "attention",
                              "group": grp, "logist": log, "reason": why}
        return f, o

    # --- новые заказы каждой группы ---
    # 1. авиа, поставщик с историей -> предпочтительный (Виноградов)
    new_f("S000601", "Shenzhen Fashion Co.", "Zara", "air", "Китай", 0, "УЛ601")
    # 2. авиа, Shanghai -> предпочтительный Орлова
    new_f("S000602", "Shanghai Apparel Group", "H&M", "air", "Китай", 0, "УЛ602")
    # 3. ВЭД, Гуанчжоу (вариант написания) -> Белова (чаще)
    new_f("S000603", "Гуанчжоу Текстиль Лтд", "Пальмира", "sea", "Китай", 1, "УЛ603")
    # 4. мультимодальные, Delhi -> Морозова
    new_f("S000604", "Delhi Cotton House", "Детский Мир", "sea", "Индия", 1, "УЛ604")
    # 5. ВЭД, новый поставщик (Китай, нет истории) -> наименее загруженный в ВЭД
    new_f("S000605", "Fujian Textiles", "Moda Rus", "sea", "Китай", 1, "УЛ605")
    # 6. авиа, новый поставщик (нет истории) -> наименее загруженный в авиа
    new_f("S000606", "Air Freight New", "Северное Сияние", "air", "Германия", 2, "УЛ606")
    # 7. предпочтительный логист (Киселёв) В ОТПУСКЕ -> наименее загруженный
    new_f("S000607", "Air Moscow Cargo", "H&M", "air", "Россия", 2, "УЛ607")
    # 8. предпочтительный логист (Тихонов) УВОЛЕН -> наименее загруженный
    new_f("S000608", "Rakuten Textiles", "H&M", "sea", "Китай", 2, "УЛ608")
    # 9. мультимодальные, новый поставщик -> наименее загруженный
    new_f("S000609", "Marvel Textiles", "Пальмира", "truck", "Турция", 3, "УЛ609")

    # --- повторная отправка заказа, уже есть в Cloud ---
    f, o, c = make_ful("S000501", "Shenzhen Fashion Co.", "Zara", "air", "Китай", 3, "УЛ501")
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "repeat_cloud", "group": None, "logist": None,
                          "reason": "заказ уже есть в Cloud — повторная отправка"}

    # --- повтор внутри пачки: два ФУЛ на один заказ (S000610) ---
    f1, o, c = make_ful("S000610", "Air Freight New", "Zara", "air", "Германия", 3, "УЛ610a")
    ful.append(f1)
    etalon[f1["Номер"]] = {"decision": "new", "group": "авиаперевозки",
                          "logist": deciding_logist("авиаперевозки", cloud, staff, "Air Freight New")[0],
                          "reason": "первый ФУЛ в пачке — новая строка"}
    f2, o2, c2 = make_ful("S000610", "Air Freight New", "Zara", "air", "Германия", 4, "УЛ610b")
    ful.append(f2)
    etalon[f2["Номер"]] = {"decision": "repeat_pack", "group": None,
                          "logist": None, "reason": "дубль в пачке: тот же заказ уже обработан"}

    # --- ловушки: нет данных ---
    # ФУЛ без номера заказа
    f, o, c = make_ful(None, "Shenzhen Fashion Co.", "Zara", "air", "Китай", 0, "УЛ-NOORDER")
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "attention", "group": None, "logist": None,
                          "reason": "нет номера заказа"}
    # ФУЛ без способа доставки
    f, o, c = make_ful("S000611", "Delhi Cotton House", "Детский Мир", None, "Индия", 4, "УЛ611")
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "attention", "group": None, "logist": None,
                          "reason": "нет способа доставки → группа не определена"}
    # ФУЛ без страны (для не-air)
    f, o, c = make_ful("S000612", "Marvel Textiles", "Пальмира", "railway", None, 5, "УЛ612")
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "attention", "group": None, "logist": None,
                          "reason": "нет страны отправления → группа не определена"}

    # --- возможный пропуск: ФУЛ с датой раньше последней записи Cloud, заказа нет в Cloud ---
    # (этот ФУЛ также является корректным новым заказом, но дата старая → попадает и в «пропуски»)
    gap_date = last_ul - timedelta(days=10)  # заметно раньше последней записи
    gap_comment = f"Номер: S000620 от {gap_date.strftime('%d.%m.%Y')}"
    f = {
        "Дата": gap_date, "Номер": "УЛ620-GAP", "Статус": "Отправлен",
        "Комментарий": gap_comment, "Способ доставки": "sea", "Бизнес единица": "Zara",
        "Логист": "", "Контрагент": "Old Delivery Co.", "Страна": "Китай",
    }
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "new+gap", "group": "вэд",
                          "logist": deciding_logist("вэд", cloud, staff, "Old Delivery Co.")[0],
                          "reason": "ФУЛ пришёл ранее последней записи Cloud → возможный пропуск, но распределяется как новый"}

    return staff, cloud, ful, etalon


# ---------------------------------------------------------------------------
# Набор-ловушка: вся группа (мультимодальные) недоступна
# ---------------------------------------------------------------------------

def build_trap(rng, seed):
    staff = STAFF_TRAP[:]
    cloud, last_ul = build_cloud(seed)
    FUL_START = last_ul + timedelta(days=1)

    def make_ful(order, supplier, brand, delivery, country, day, num, omit_order=False):
        date = FUL_START + timedelta(days=day)
        comment = (f"Номер: {order} от {date.strftime('%d.%m.%Y')}"
                   if order else "Упаковочный лист без номера заказа")
        return {
            "Дата": date, "Номер": f"{num}-{order}" if order else num,
            "Статус": "Отправлен", "Комментарий": comment,
            "Способ доставки": delivery, "Бизнес единица": brand,
            "Логист": "", "Контрагент": supplier, "Страна": country,
        }

    ful = []
    etalon = {}

    def new_f(order, supplier, brand, delivery, country, day=0, num="УЛ"):
        grp, r = rule_group(delivery, country)
        log, why = (None, r) if grp is None else deciding_logist(grp, cloud, staff, supplier)
        f = make_ful(order, supplier, brand, delivery, country, day, num)
        ful.append(f)
        etalon[f["Номер"]] = {"decision": "new" if log else "attention",
                             "group": grp, "logist": log, "reason": why}

    # нормальные распределяемые заказы (авиа и ВЭД доступны)
    for i, case in enumerate([
        ("S000701", "Shenzhen Fashion Co.", "Zara", "air", "Китай"),
        ("S000702", "Shanghai Apparel Group", "H&M", "air", "Китай"),
        ("S000703", "Гуанчжоу Текстиль Лтд", "Пальмира", "sea", "Китай"),
        ("S000704", "Shenzhen Fashion Co.", "Детский Мир", "air", "Китай"),
        ("S000705", "Shanghai Apparel Group", "Moda Rus", "air", "Китай"),
        ("S000706", "Гуанчжоу Текстиль Лтд", "Северное Сияние", "sea", "Китай"),
        ("S000707", "Fujian Textiles", "Zara", "sea", "Китай"),
        ("S000708", "Shenzhen Fashion Co.", "H&M", "air", "Китай"),
        ("S000709", "China Trading New", "Пальмира", "sea", "Китай"),
        ("S000710", "Shanghai Apparel Group", "Детский Мир", "air", "Китай"),
    ]):
        order, supplier, brand, delivery, country = case
        new_f(order, supplier, brand, delivery, country, i % 5, f"УЛ7{i:02d}")

    # мультимодальный заказ: вся группа в отпуске/уволена -> логист не определён
    f = make_ful("S000711", "Munich Cargo", "Zara", "truck", "Германия", 5, "УЛ711")
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "attention", "group": "мультимодальные",
                          "logist": None,
                          "reason": "вся группа недоступна (в отпуске/уволены) — логист не определён"}
    # ещё мультимодальные «в глухой группе»
    f = make_ful("S000712", "Krakow Transport", "H&M", "railway", "Польша", 6, "УЛ712")
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "attention", "group": "мультимодальные",
                          "logist": None,
                          "reason": "вся группа недоступна (в отпуске/уволены) — логист не определён"}
    # ещё один недоступный мультимодальный
    f = make_ful("S000713", "Vienna Freight", "Moda Rus", "sea", "Австрия", 6, "УЛ713")
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "attention", "group": "мультимодальные",
                          "logist": None,
                          "reason": "вся группа недоступна (в отпуске/уволены) — логист не определён"}
    # ловушка без данных
    f = make_ful(None, "Shenzhen Fashion Co.", "Zara", "air", "Китай", 0, "УЛ-NOORDER")
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "attention", "group": None, "logist": None,
                          "reason": "нет номера заказа"}
    # повторная отправка заказа уже в Cloud
    f = make_ful("S000501", "Shenzhen Fashion Co.", "Zara", "air", "Китай", 3, "УЛ501")
    ful.append(f)
    etalon[f["Номер"]] = {"decision": "repeat_cloud", "group": None, "logist": None,
                          "reason": "заказ уже есть в Cloud — повторная отправка"}

    return staff, cloud, ful, etalon


# ---------------------------------------------------------------------------

def write_set(out, staff, cloud, ful, etalon, kind):
    os.makedirs(out, exist_ok=True)

    # ФУЛ
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "ФУЛ"
    ful_h = ["Дата", "Cut-off date", "Номер", "Статус", "Комментарий",
             "Способ доставки", "Бизнес единица", "Логист", "Контрагент", "Страна"]
    ws.append(ful_h)
    for f in ful:
        ws.append([f.get(h, "") for h in ful_h])
    wb.save(os.path.join(out, "ful_1c.xlsx")); wb.close()

    # Cloud
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Реестр Cloud"
    cloud_h = ["Заказ", "Поставщик", "Бренд", "Логист", "Статус",
               "Дата получения УЛ", "Страна отправления", "Способ доставки",
               "Cut-off", "Комментарий"]
    ws.append(cloud_h)
    for c in cloud:
        ws.append([c.get(k, "") for k in
                   ["order", "supplier", "brand", "logist", "status", "date_ul",
                    "country", "delivery", "cutoff", "comment"]])
    wb.save(os.path.join(out, "cloud.xlsx")); wb.close()

    # Сотрудники
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Сотрудники"
    ws.append(["Наименование", "Текущее состояние", "Должность", "Код"])
    for name, role, state, code in staff:
        ws.append([name, state, role, code])
    wb.save(os.path.join(out, "staff_1c.xlsx")); wb.close()

    # Эталон
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Эталон"
    ws.append(["Ключ(Заказ/ФУЛ)", "Решение", "Группа", "Логист", "Обоснование"])
    for k, v in etalon.items():
        ws.append([k, v["decision"], v.get("group"), v.get("logist"), v.get("reason")])
    wb.save(os.path.join(out, "etalon.xlsx")); wb.close()

    print(f"[{kind}] ФУЛ: {len(ful)} | Cloud: {len(cloud)} | Сотрудники: {len(staff)} | "
          f"Эталон: {len(etalon)} -> {out}")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=".")
    ap.add_argument("--mode", default="full", choices=["full", "trap", "chk"])
    ap.add_argument("--seed", type=int, default=42)
    args = ap.parse_args(argv)

    rng = random.Random(args.seed)

    if args.mode == "full":
        staff, cloud, ful, etalon = build_full(rng, args.seed)
        write_set(args.out, staff, cloud, ful, etalon, "full")
    elif args.mode == "chk":
        staff, cloud, ful, etalon = build_full(rng, args.seed + 1000)
        write_set(args.out, staff, cloud, ful, etalon, "chk")
    else:  # trap
        staff, cloud, ful, etalon = build_trap(rng, args.seed)
        write_set(args.out, staff, cloud, ful, etalon, "trap")


if __name__ == "__main__":
    main()