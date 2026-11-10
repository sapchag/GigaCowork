"""Вспомогательные функции агента распределения заказов между логистами.

Только стандартная библиотека Python 3. Колонки ищутся по заголовкам,
а не по номеру, чтобы агент работал на файлах заказчика с другим порядком
колонок, лишними колонками и иначе названными листами.
"""

import re
from datetime import datetime

# ---------------------------------------------------------------------------
# Нормализация строк: регистр, пробелы, знаки препинания
# ---------------------------------------------------------------------------

def norm_key(value):
    """Ключевая нормализация для сопоставления: только буквы/цифры, нижний регистр."""
    if value is None:
        return ""
    s = str(value)
    return "".join(ch for ch in s.lower() if ch.isalnum())


def norm_soft(value):
    """Мягкая нормализация: нижний регистр и схлопывание пробелов."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().lower()


# ---------------------------------------------------------------------------
# Поиск колонки по заголовку
# ---------------------------------------------------------------------------

def _norm_header(h):
    h = norm_soft(h)
    h = re.sub(r"[^a-zа-яё0-9]", "", h)
    return h


def find_column(header_row, candidates, required_names=None):
    """Вернуть индекс колонки, чей заголовок совпадает с одним из кандидатов.

    header_row  — список заголовков (значения первой строки).
    candidates  — список нормализованных кандидатов (нижний регистр).
    required_names — если задан, заголовок должен содержать одно из этих
                     «ключевых» слов (в обычном виде, до нормализации).
    """
    normed = {}
    for idx, h in enumerate(header_row):
        key = _norm_header(h)
        if key not in normed:
            normed[key] = idx
    cand_norm = {_norm_header(c) for c in candidates}
    for key in sorted(normed):
        if key in cand_norm:
            return normed[key]
    if required_names:
        low = {_norm_header(c) for c in required_names}
        for key in sorted(normed):
            if key and any(rn and rn in key for rn in low):
                return normed[key]
    return None


# ---------------------------------------------------------------------------
# Даты
# ---------------------------------------------------------------------------

def parse_datetime(value):
    """Вернуть datetime из значения Excel (datetime / date / строка)."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if hasattr(value, "isoformat"):  # date
        return datetime(value.year, value.month, value.day)
    s = str(value).strip()
    if not s:
        return None
    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y",
                "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
                "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def fmt_date(value):
    """Дата в виде ДД.ММ.ГГГГ для комментариев и строк Cloud."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    try:
        return value.strftime("%d.%m.%Y")
    except Exception:
        return str(value)


# ---------------------------------------------------------------------------
# Номер заказа из комментария: 'Номер: S0001234 от ДД.ММ.ГГГГ'
# ---------------------------------------------------------------------------

ORDER_RE = re.compile(r"[НH]омер[:.]?\s*([A-Za-zА-Яа-я0-9_-]{3,})\s*(?:от\s+\S+)?", re.IGNORECASE)


def extract_order_number(comment):
    """Достать номер заказа из комментария ФУЛ или вернуть None."""
    if not comment:
        return None
    m = ORDER_RE.search(norm_soft(comment))
    if m:
        return m.group(1).upper()
    return None


# ---------------------------------------------------------------------------
# Группа логистов по способу доставки и стране
# ---------------------------------------------------------------------------

def is_air(delivery):
    return norm_soft(delivery) == "air"


def is_china(country):
    c = norm_soft(country)
    return any(k in c for k in ("кита", "гонконг", "hong", "hk"))


def determine_group(delivery, country):
    """Правило: air -> авиа; не air + Китай -> ВЭД; не air + не Китай -> мультимодальные.

    Возвращает (группа, причина) — причина заполняется для случая «не определено».
    """
    if delivery is None or norm_soft(delivery) == "":
        return None, "нет способа доставки"
    if is_air(delivery):
        return "авиаперевозки", None
    if country is None or norm_soft(country) == "":
        return None, "нет страны отправления"
    if is_china(country):
        return "вэд", None
    return "мультимодальные", None


GROUP_BY_ROLE = {
    "специалистпоавиаперевозкам": "авиаперевозки",
    "специалистпоработесвэддокументацией": "вэд",
    "специалистпомультимодальнымперевозкам": "мультимодальные",
}

EXCLUDE_ROLE_WORDS = ("ассистент", "руководитель", "ассистен", "помощник")


def role_group(role):
    """Группа по должности сотрудника; руководители/ассистенты -> None."""
    r = _norm_header(role)
    if not r:
        return None
    if any(w in r for w in EXCLUDE_ROLE_WORDS):
        return None
    return GROUP_BY_ROLE.get(r)


def is_working(state):
    s = norm_soft(state)
    return "работа" in s


# ---------------------------------------------------------------------------
# Сопоставление коротких имён Cloud с ФИО из 1С
# ---------------------------------------------------------------------------

def _tokens(name):
    parts = re.split(r"\s+", norm_soft(name))
    return [p for p in parts if p]


def match_short_name(short_name, full_name):
    """True, если короткое имя Cloud ('Иванов И.') относится к ФИО из 1С."""
    if not short_name or not full_name:
        return False
    st = [p for p in re.split(r"\s+", norm_soft(short_name)) if p]
    ft = [p for p in re.split(r"\s+", norm_soft(full_name)) if p]
    if not st or not ft:
        return False
    if st[0] != ft[0]:  # фамилии должны совпадать
        return False
    if len(st) == 1:
        return True  # только фамилия — считаем совпадением
    s_first = st[1].lstrip(".")
    f_first = ft[1]
    if s_first and f_first:
        return s_first[0] == f_first[0]
    return True


# ---------------------------------------------------------------------------
# Нормализация поставщика для «один поставщик — один логист»
# ---------------------------------------------------------------------------

def norm_supplier(value):
    return norm_key(value)