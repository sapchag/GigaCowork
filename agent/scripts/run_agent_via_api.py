"""Запуск агента распределения заказов через GigaChat API.

Детерминированные правила (новый/повтор/дубль в пачке, группа, логист) всегда
исполняет скрипт distributes (distribute.py). Этот скрипт лишь добавляет к нему
языковую модель GigaChat: передаёт модели системный промпт агента и сводку
результата, а модель формирует ответ пользователю по шаблону.

Это вариант «свой скрипт вокруг API модели» из ТЗ (агентный цикл с вызовом
инструментов писать не обязательно). В GigaCowork ту же роль выполняет навык.

КЛЮЧИ НИКОГДА НЕ ПИШУТСЯ В КОД, ФАЙЛЫ, ЛОГИ И РЕПОЗИТОРИЙ.
Они читаются ТОЛЬКО из переменных окружения:
  GIGACHAT_API_KEY                       — base64 от "<client_id>:<auth_key>"
                                           (готовый заголовок Basic; так хранит
                                           classic-ключ разработчик на sber.ru), и/или
  GIGACHAT_CLIENT_ID + GIGACHAT_AUTH_KEY — пара для OAuth (другой вариант)

Требуются пакеты: langchain-gigachat (pip install langchain-gigachat).

Запуск (из папки agent/scripts):
  python run_agent_via_api.py \
      --ful ../tests/set1/ful_1c.xlsx \
      --cloud ../tests/set1/cloud.xlsx \
      --staff ../tests/set1/staff_1c.xlsx \
      --out ./out \
      --prompt ../system_prompt.txt \
      --model "GigaChat-2-Pro"

Модель по умолчанию: GigaChat-2-Pro (из списка get_models().data).
"""

import argparse
import base64
import os
import subprocess
import sys

from langchain_core.messages import HumanMessage, SystemMessage
from langchain_gigachat.chat_models import GigaChat  # pip install langchain-gigachat


def resolve_credentials():
    """Вернуть credentials для GigaChat: строка base64 (Basic) или (client_id, auth_key)."""
    raw = os.environ.get("GIGACHAT_API_KEY")
    if raw:
        cred = raw
        while len(cred) % 4 != 0:
            cred += "="  # выровнять длину base64
        base64.b64decode(cred, validate=True)  # проверка, что это валидный base64
        return cred, "api_key"

    cid = os.environ.get("GIGACHAT_CLIENT_ID")
    auth = os.environ.get("GIGACHAT_AUTH_KEY")
    if cid and auth:
        return (cid, auth), "oauth_pair"
    raise SystemExit(
        "Нет ключа. Задайте в переменных окружения GIGACHAT_API_KEY (base64 от "
        "client_id:auth_key) либо пару GIGACHAT_CLIENT_ID + GIGACHAT_AUTH_KEY."
    )


def main():
    ap = argparse.ArgumentParser(description="Агент распределения заказов + GigaChat")
    ap.add_argument("--ful", required=True, help="Выгрузка ФУЛ из 1С (.xlsx)")
    ap.add_argument("--cloud", required=True, help="Реестр заказов Cloud (.xlsx)")
    ap.add_argument("--staff", required=True, help="Список сотрудников из 1С (.xlsx)")
    ap.add_argument("--out", default=".", help="Папка результата")
    ap.add_argument("--script", default="distribute.py", help="Путь к distribute.py")
    ap.add_argument("--prompt", default="../system_prompt.txt", help="Системный промпт агента")
    ap.add_argument("--model", default="GigaChat-2-Pro", help="Модель GigaChat")
    ap.add_argument("--scope", default="GIGACHAT_API_PERS")
    args = ap.parse_args()

    credentials, mode = resolve_credentials()
    print(f"[ok] credentials: режим '{mode}' (ключ не печатается)", file=sys.stderr)

    # 1) детерминированные правила
    os.makedirs(args.out, exist_ok=True)
    p = subprocess.run(
        [sys.executable, args.script,
         "--ful", args.ful, "--cloud", args.cloud, "--staff", args.staff,
         "--out", args.out],
        capture_output=True, text=True)
    if p.returncode != 0:
        print("distribute.py упал:\n", p.stdout, p.stderr, file=sys.stderr)
        return 1
    out = p.stdout
    start, end = out.find("{"), out.rfind("}")
    summary = out[start:end + 1] if start != -1 else out.strip()
    print(f"[ok] distribute.py -> {os.path.join(args.out, 'result_cloud.xlsx')}", file=sys.stderr)
    print(f"[summary] {summary}", file=sys.stderr)

    # 2) модель GigaChat: системный промпт + сводка -> ответ пользователю
    model = GigaChat(
        credentials=credentials,
        scope=args.scope,
        model=args.model,
        verify_ssl_certs=False,  # самоподписанный сертификат GigaChat
        temperature=0.0,
        max_tokens=600,
    )
    with open(args.prompt, encoding="utf-8") as f:
        system = f.read()
    user = (
        "Набор заказов обработан детерминированным скриптом. Сводка:\n"
        f"{summary}\n"
        "Полный результат сохранён в result_cloud.xlsx (листы: Новые строки Cloud, "
        "Повторные отправки, Возможные пропуски, Требует внимания, Журнал решений).\n"
        "Дай короткий ответ пользователю по шаблону: 'Новых строк Cloud: N / Повторных "
        "отправок: M / Требует внимания: P', упомяни файл и что лежит на каждом листе. "
        "Не выдумывай значения."
    )
    resp = model.invoke([SystemMessage(content=system), HumanMessage(content=user)])
    print("\n===== ОТВЕТ МОДЕЛИ GigaChat =====\n")
    print(resp.content)
    return 0


if __name__ == "__main__":
    sys.exit(main())