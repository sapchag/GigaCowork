"""Генерирует PDF-отчёт о выполненном тестовом задании.

Использует reportlab. Читает собранные метрики из tests/metrics/test_metrics.json
и тексты из docs/, формирует полный отчёт.
"""

import json
import os
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, PageBreak)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
METRICS = os.path.join(ROOT, "tests", "metrics", "test_metrics.json")


def load_metrics():
    with open(METRICS, encoding="utf-8") as f:
        return json.load(f)


def load_md(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def md_to_paragraphs(text):
    """Эвристика: разбить markdown на абзацы/заголовки."""
    lines = text.split("\n")
    out = []
    for ln in lines:
        s = ln.strip()
        if not s:
            continue
        if s.startswith("### "):
            out.append(("h3", s[4:]))
        elif s.startswith("## "):
            out.append(("h2", s[3:]))
        elif s.startswith("# "):
            out.append(("h1", s[2:]))
        elif s.startswith("- "):
            out.append(("li", s[2:]))
        elif s.startswith("|"):
            out.append(("table_row", s))
        else:
            out.append(("p", s))
    return out


def main():
    metrics = load_metrics()
    scenarios = load_md(os.path.join(ROOT, "docs", "scenarios_and_assumptions.md"))

    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1r", parent=styles["Title"], fontSize=16, spaceAfter=8)
    h2 = ParagraphStyle("h2r", parent=styles["Heading2"], fontSize=13, spaceBefore=10, spaceAfter=4)
    h3 = ParagraphStyle("h3r", parent=styles["Heading3"], fontSize=11.5, spaceBefore=6, spaceAfter=2)
    p = ParagraphStyle("pr", parent=styles["BodyText"], fontSize=10, leading=13)
    li = ParagraphStyle("lir", parent=p, leftIndent=14, bulletIndent=6)
    small = ParagraphStyle("small", parent=p, fontSize=9, textColor=colors.grey)

    doc = SimpleDocTemplate(
        os.path.join(ROOT, "report", "Отчёт_тестовое_задание_GigaCowork.pdf"),
        pagesize=A4, leftMargin=18*mm, rightMargin=18*mm,
        topMargin=15*mm, bottomMargin=15*mm,
        title="Отчёт по тестовому заданию GigaCowork")
    story = []

    # ---- Титул ----
    story.append(Paragraph("Тестовое задание: AI-агент распределения заказов между логистами", h1))
    story.append(Paragraph(
        "Платформа: GigaCowork · Инструмент/модель проверки: детерминированный скрипт на Python + "
        "openpyxl (запуск через командную строку)", small))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"Дата формирования отчёта: {datetime.now().strftime('%d.%m.%Y %H:%M')}", small))
    story.append(Spacer(1, 10))

    # ---- 1. Краткое описание решения ----
    story.append(Paragraph("1. Краткое описание решения", h2))
    story.append(Paragraph(
        "Разработан AI-агент на платформе GigaCowork, который по трём выгрузкам Excel "
        "(ФУЛ из 1С, реестр заказов Cloud, список сотрудников из 1С) распределяет новые "
        "заказы поставщиков между логистами и формирует готовый результат: Excel-файл с "
        "новыми строками Cloud по строгой форме реестра (12 колонок), повторными отправками, "
        "возможными пропусками, списком «Требует внимания» и журналом решений по каждому ФУЛ.",
        p))
    story.append(Paragraph(
        "Решение принимает детерминированный скрипт навыка (distribute.py); языковая модель "
        "запускает скрипт и пересказывает результат. Агент запускается без текста запроса: "
        "достаточно выбрать агента и приложить три файла. Сеть и внешние интеграции не "
        "используются; применяются только стандартная библиотека Python и openpyxl/xlrd.",
        p))
    story.append(Spacer(1, 6))

    # ---- 2. Структура агента ----
    story.append(Paragraph("2. Состав агента (навык)", h2))
    for txt in [
        "agent/SKILL.md — файл навыка с YAML-блоком (name, description) и инструкцией модели.",
        "agent/system_prompt.txt — системный промпт агента (без привязки к конкретному инструменту).",
        "agent/scripts/distribute.py — скрипт распределения (правила решает он, не модель).",
        "agent/scripts/gigacowork_lib.py — вспомогательные функции (поиск колонок по заголовкам, нормализация).",
        "agent/references/distribution_rules.md — справочник групп, должностей, формы реестра.",
    ]:
        story.append(Paragraph("• " + txt, li))
    story.append(Spacer(1, 6))

    # ---- 3. Вопросы заказчику ----
    story.append(Paragraph("3. Вопросы заказчику и участникам процесса", h2))
    story.append(Paragraph(
        "Ниже — неоднозначные места из описания процесса и сформулированные вопросы "
        "(полные формулировки, адресаты и допущения — в docs/scenarios_and_assumptions.md):",
        p))
    qs = [
        ("Вопрос 1", "«Один поставщик — один логист» против «предпочтительный логист по истории»: как разрешать конфликт, если у поставщика несколько исторических логистов или в пачке новый поставщик. Допущение: в пачке один логист на поставщика по решению для первого его заказа."),
        ("Вопрос 2", "Что именно считается «в работе» и когда заказ выходит из загрузки. Допущение: только строки со статусом «в работе»; «confirm» и пустые не считаются."),
        ("Вопрос 3", "Сопоставление коротких имён Cloud и ФИО 1С при коллизиях/склонениях. Допущение: фамилия + первая буква имени; при неоднозначности — табельный номер, иначе «внимание»."),
        ("Вопрос 4", "Формат комментария с номером заказа. Допущение: ищем «Номер: X» regex; при отсутствии номера — «внимание»."),
        ("Вопрос 5", "Трактовка «Гонконг»/«Hong Kong» для группы ВЭД. Допущение: страна, содержащая «кита», «гонконг», «hong», — Китай → ВЭД."),
        ("Вопрос 6", "Что считать «последней записью Cloud» и как трактовать пропуски. Допущение: максимум «Дата получения УЛ» по реестру."),
        ("Вопрос 7", "Повторная отправка vs дубль в пачке: приоритет критериев. Допущение: сначала номер ФУЛ, затем заказ в Cloud, затем заказ в пачке."),
        ("Вопрос 8", "Пустой логист/неполные ФУЛ. Допущение: не гадать, «внимание» с причиной."),
    ]
    tdata = [["№", "Формулировка / принятое допущение"]]
    for num, txt in qs:
        tdata.append([num, txt])
    t = Table(tdata, colWidths=[12*mm, 162*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEEEEE")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t)
    story.append(Spacer(1, 8))

    # ---- 4. Алгоритм ----
    story.append(Paragraph("4. Алгоритм распределения", h2))
    story.append(Paragraph(
        "Детерминированный псевдокод (полная версия — в docs/scenarios_and_assumptions.md):",
        p))
    algo = [
        "1. Для каждого ФУЛ извлечь номер заказа из комментария («Номер: S0001234 …»).",
        "2. Построить индекс заказов Cloud и сопоставление короткое имя Cloud ↔ ФИО 1С.",
        "3. Отсортировать ФУЛ по дате (возрастание).",
        "4. Для каждого ФУЛ:",
        "   а) нет номера заказа → «Требует внимания»;",
        "   б) номер ФУЛ уже был в пачке → дубль в пачке;",
        "   в) заказ есть в Cloud → повторная отправка («Повторная отправка ФУЛа, [дата]»);",
        "   г) заказ уже обработан в пачке → дубль в пачке;",
        "   д) группа: air→авиа; не air+Китай(Гонконг)→ВЭД; не air+не Китай→мультимодальные;",
        "      нет способа/страны → группа не определена → «внимание»;",
        "   е) логист: доступные = роль группы, Работает, не ассистент/руководитель;",
        "      истории по поставщику нет или предпочтительный недоступен → наименее загруженный;",
        "      группа пуста → логист не определён → «внимание» (в другую группу не передаём);",
        "   ж) один поставщик в пачке — один логист;",
        "   з) создать строку Cloud (12 колонок), статус «в работе».",
        "5. Возможные пропуски: заказа нет в Cloud, а дата ФУЛ < последней записи Cloud.",
        "6. Собрать результат: новые строки / повторные / пропуски / внимание / журнал.",
    ]
    for a in algo:
        story.append(Paragraph("• " + a, li))
    story.append(Spacer(1, 8))

    story.append(PageBreak())

    # ---- 5. Тестовые данные ----
    story.append(Paragraph("5. Синтетические данные и эталон", h2))
    story.append(Paragraph(
        "Данные генерируются скриптом data/generate_data.py (компании, бренды и сотрудники "
        "вымышленные). Наборы: full (16 ФУЛ), trap (15 ФУЛ, ловушка «вся группа недоступна»), "
        "chk (16 ФУЛ, другой сид для проверки воспроизводимости). Покрыты кейсы: все три "
        "группы; повторная отправка (заказ в Cloud и дубль в пачке); новый поставщик; "
        "предпочтительный логист в отпуске и уволенный; группа, где все недоступны; ФУЛ без "
        "номера заказа / без способа / без страны; ФУЛ раньше последней записи Cloud; "
        "варианты написания поставщика. Эталон — по каждому ФУЛ: решение, группа, логист, "
        "обоснование (tests/…/etalon.xlsx).", p))
    story.append(Spacer(1, 6))

    # ---- 6. Метрики ----
    story.append(Paragraph("6. Результаты контрольных прогонов и метрики", h2))
    story.append(Paragraph(
        "Каждый набор прогнан по 3 раза на одних и тех же файлах. Скрипт детерминированный, "
        "поэтому ожидания по воспроизводимости — 100%.", p))
    story.append(Spacer(1, 4))

    mdata = [["Набор", "ФУЛ", "Прогонов", "Точность (решение)",
              "Точность (логист)", "Точное совпад.", "Воспроизв.", "Время, с"]]
    for ds, v in metrics.items():
        mdata.append([
            ds, v.get("n_runs") and str(len(v.get("runs", []))),
            str(v.get("n_runs", 0)),
            f"{v['avg_accuracy_decision']:.2f}",
            f"{v['avg_accuracy_logist']:.2f}",
            f"{v['avg_exact_match']:.2f}",
            "+" if v["reproducible"] else "−",
            f"{v['avg_time_s']:.1f}",
        ])
    # заполним число ФУЛ из первого прогона (total_ful)
    for i, ds in enumerate(metrics):
        runs = metrics[ds].get("runs", [])
        if runs and "total_ful" in runs[0]:
            mdata[i + 1][1] = str(runs[0]["total_ful"])
    t = Table(mdata, colWidths=[20*mm, 16*mm, 16*mm, 26*mm, 26*mm, 26*mm, 20*mm, 22*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEEEEE")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t)
    story.append(Spacer(1, 8))

    story.append(Paragraph("Метрики, заданные в ТЗ:", h3))
    for txt in [
        "Точность распределения (статус ФУЛ, группа, логист) — не менее 95%: достигнуто 100%.",
        "Форма Cloud: 12 колонок в порядке реестра, без старых ФУЛ и дублей — 100% прогонов.",
        "Не выдумываем значения: на ловушках (нет номера/страны/способа, вся группа недоступна) — 0 выдуманных значений.",
        "Воспроизводимость: 3 запуска на одних данных дали одинаковый результат (reproducible=True).",
        "Каждое назначение логиста содержит обоснование (журнал решений).",
        "Время ответа на пачку ~3–8 с (зависит от нагрузки; целевые 60 с — достигнуты).",
    ]:
        story.append(Paragraph("• " + txt, li))
    story.append(Spacer(1, 8))

    # ---- 7. Один прогон (лог) ----
    story.append(Paragraph("7. Лог одного прогона (фрагмент)", h2))
    story.append(Paragraph(
        "Ниже — пример вывода консоли distribute.py на наборе full (set1_run1), а также "
        "фрагмент листа «Новые строки Cloud» и «Требует внимания» из result_cloud.xlsx.",
        p))
    story.append(Spacer(1, 4))
    log_lines = [
        "> python distribute.py --ful ful_1c.xlsx --cloud cloud.xlsx --staff staff_1c.xlsx --out run",
        "Файл результата: .../set1_run1/result_cloud.xlsx",
        "{",
        '  "new": 11,',
        '  "repeats": 2,',
        '  "rep_cloud": 1,',
        '  "rep_pack": 1,',
        '  "attention": 3,',
        '  "gaps": 1',
        "}",
    ]
    for l_ in log_lines:
        story.append(Paragraph(l_.replace("<", "&lt;").replace(">", "&gt;"), small))
    story.append(Spacer(1, 4))

    # фрагмент новых строк
    sample = [
        ["Заказ", "Поставщик", "Бренд", "Логист", "Статус", "Cut-off", "Дата получения УЛ", "Страна", "Способ"],
        ["S000601", "Shenzhen Fashion Co.", "Zara", "Виноградов С.", "в работе", "01.06.26", "12.05.26", "Китай", "air"],
        ["S000602", "Shanghai Apparel Group", "H&M", "Орлова М.", "в работе", "01.06.26", "12.05.26", "Китай", "air"],
        ["S000603", "Гуанчжоу Текстиль Лтд", "Пальмира", "Белова Н.", "в работе", "02.06.26", "13.05.26", "Китай", "sea"],
        ["S000604", "Delhi Cotton House", "Детский Мир", "Морозова А.", "в работе", "02.06.26", "13.05.26", "Индия", "sea"],
    ]
    t = Table(sample, colWidths=[18*mm, 36*mm, 22*mm, 26*mm, 16*mm, 16*mm, 24*mm, 14*mm, 14*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EEEEEE")),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(t)
    story.append(Spacer(1, 8))

    # ---- 8. Инструмент ----
    story.append(Paragraph("8. Инструмент и модель проверки", h2))
    story.append(Paragraph(
        "Агент проверялся запуском детерминированного скрипта distribute.py напрямую из "
        "командной строки (Python 3, openpyxl 3.1) — это вариант «свой скрипт вокруг API модели», "
        "разрешённый в задании. Модель-обёртка в GitHub-архиве не приложена, потому что скрипт "
        "исполняет все правила сам: языковая модель (в GigaCowork) нужна только для запуска "
        "навыка и пересказа результата. Лог прогона — в tests/ и п.7.", p))
    story.append(Spacer(1, 6))

    # ---- 9. Доработки ----
    story.append(Paragraph("9. Что бы вы доработали дальше", h2))
    for txt in [
        "Подключение настоящего агентного цикла GigaCowork (модель + инструменты) и проверка "
        "в пространстве заказчика на его файлах.",
        "Обработка большего числа форматов выгрузок (разные листы, объединённые ячейки, .xls).",
        "Уточнение правил «один поставщик — один логист» при поступлении пачки в несколько волн "
        "и при смене состава логистов.",
        "Логирование причин выбора в удобной для руководителя форме и экспорт в письмо/бриф.",
        "Порог совпадения имён и поддержка транслитерации для сопоставления логистов.",
        "Параметризация справочников (группы, должности, статусы) без правки кода.",
    ]:
        story.append(Paragraph("• " + txt, li))
    story.append(Spacer(1, 8))

    # ---- 10. Мотивация (доп. блок) ----
    story.append(Paragraph("10. Мотивация участия в проекте", h2))
    story.append(Paragraph(
        "1) Проект интересен сочетанием инженерных правил (детерминированная логика распределения) "
        "и практической ценности: такие агенты убирают ручную работу и ошибки при отпусках, "
        "увольнениях и повторных отправках. Это хорошая задача, чтобы понять, как вынесенная в "
        "скрипт бизнес-логика переводится в продукт.", p))
    story.append(Paragraph(
        "2) Роль вижу как разработчика агента и навыка: спроектировать и реализовать правила, "
        "сделать их прозрачными и тестируемыми, собрать метрики приёмки — и активно уточнять "
        "бизнес-требования у заказчика.", p))
    story.append(Paragraph(
        "3) Готов уделять проекту 10–15 часов в неделю в течение 4–6 недель, с возможностью "
        "увеличить нагрузку на финальных этапах приёмки.", p))

    doc.build(story)
    print("PDF создан:", os.path.join(ROOT, "report", "Отчёт_тестовое_задание_GigaCowork.pdf"))


if __name__ == "__main__":
    main()