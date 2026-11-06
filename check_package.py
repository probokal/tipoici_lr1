#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_package.py — самопроверка пакета нового бота (папка 18_).

    python3 tools/check_package.py
    python3 tools/check_package.py --path /путь/к/папке
    python3 tools/check_package.py --min 80

Что проверяет: файлы на месте, промпт не распух и не потерял ключевые правила,
скиллы целы, требования преподавателя записаны, чекер отчёта реально ловит
плохой .docx и пропускает хороший.

Коды возврата: 0 — пакет годен, 2 — есть провалы, 1 — не найдена папка.
"""

import argparse
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile

REQUIRED = {
    "00_ЧИТАЙ_МЕНЯ.md": "карта пакета",
    "01_ПРОМПТ_ЗАПУСК.md": "промпт запуска",
    "02_ДИАГНОЗ_СТАРОГО_БОТА.md": "разбор лога",
    "03_СИТУАЦИЯ_СЕЙЧАС.md": "состояние ЛР1/ЛР2",
    "04_СКИЛЛЫ.md": "8 скиллов",
    "05_ТРЕБОВАНИЯ_ПРЕПОДАВАТЕЛЯ.md": "требования",
    "06_ГДЕ_ОСТАНОВИЛИСЬ.md": "точка остановки",
    "tools/check_report.py": "чекер отчёта",
    "tools/make_fixtures.py": "фикстуры",
}


def read(path):
    if not os.path.isfile(path):
        return ""
    return io.open(path, encoding="utf-8", errors="replace").read()


def prompt_body(text):
    """Всё, что внутри первого ```…``` блока файла 01."""
    m = re.search(r"```[a-zA-Z]*\n(.*?)\n```", text, re.S)
    return m.group(1) if m else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--path", default=os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    ap.add_argument("--min", type=int, default=80)
    args = ap.parse_args()

    root = os.path.abspath(args.path)
    if not os.path.isdir(root):
        print("ОШИБКА: нет папки %s" % root)
        return 1

    checks = []

    def add(name, weight, ok, detail, critical=False):
        checks.append({"name": name, "weight": weight, "ok": bool(ok),
                       "detail": detail, "critical": critical})

    # 1. файлы
    missing = [f for f in REQUIRED if not os.path.isfile(os.path.join(root, f))]
    add("Все обязательные файлы на месте", 12, not missing,
        ("нет: " + ", ".join(missing)) if missing else "%d файлов" % len(REQUIRED),
        critical=True)

    p01 = read(os.path.join(root, "01_ПРОМПТ_ЗАПУСК.md"))
    body = prompt_body(p01)

    # 2. размер промпта
    n = len(body)
    add("Промпт влезает в одно сообщение", 10, 2500 <= n <= 6000,
        "%d символов, %d строк" % (n, len(body.splitlines())), critical=True)

    # 3. ключевые правила в промпте
    rules = {
        "одна команда за сообщение": r"одн\w*\s+(блок|команд)",
        "строка РЕСУРС": r"РЕСУРС",
        "критерий успеха": r"КРИТЕРИЙ",
        "ожидание ввода": r"ЖДУ",
        "журнал шагов": r"журнал",
        "реальные пути": r"D:\\учеба",
        "запрет повтора": r"повтор",
        "формат ≤15 строк": r"15\s*строк",
    }
    lost = [k for k, rx in rules.items() if not re.search(rx, body, re.I)]
    add("В промпте все правила-предохранители", 16, not lost,
        ("потеряно: " + ", ".join(lost)) if lost else "8 из 8", critical=True)

    # 4. скиллы
    p04 = read(os.path.join(root, "04_СКИЛЛЫ.md"))
    skills = re.findall(r"^##\s+S(\d+)\s*·", p04, re.M)
    blocks = re.split(r"^##\s+S\d+\s*·", p04, flags=re.M)[1:]
    thin = [i + 1 for i, b in enumerate(blocks)
            if len([w for w in ("Когда:", "Метод:", "Стоп:", "Тест:") if w in b]) < 4]
    add("Скиллов восемь, у каждого триггер и стоп", 10,
        len(skills) == 8 and not thin,
        "найдено S%s" % ",".join(skills) + ("; неполные: S%s" % thin if thin else ""))

    # 5. требования преподавателя
    p05 = read(os.path.join(root, "05_ТРЕБОВАНИЯ_ПРЕПОДАВАТЕЛЯ.md")).lower()
    need = {
        "TNR 14": "times new roman" in p05 and "14" in p05,
        "интервал 1,5": "1,5" in p05,
        "по ширине": "по ширине" in p05,
        "подписи 12": "12" in p05,
        "сквозная нумерация": "сквозн" in p05,
        "ссылка до объекта": "ссылк" in p05 and "до" in p05,
        "рамка": "рамк" in p05,
        "чек-лист": "чек-лист" in p05,
    }
    lost = [k for k, ok in need.items() if not ok]
    add("Требования записаны полностью", 10, not lost,
        ("нет: " + ", ".join(lost)) if lost else "8 из 8")

    # 6. ситуация реальная, а не выдуманная
    p03 = read(os.path.join(root, "03_СИТУАЦИЯ_СЕЙЧАС.md"))
    anchors = {
        "Рис. 22": "22" in p03,
        "скрин F9": "F9" in p03,
        "FTK Mount Image": "Mount Image" in p03,
        "RAID5 роли": "[0][1][2][4]" in p03,
        "путь к лабораторным": "типоици" in p03,
    }
    lost = [k for k, ok in anchors.items() if not ok]
    add("Ситуация опирается на факты из лога", 8, not lost,
        ("нет: " + ", ".join(lost)) if lost else "5 из 5 якорей")

    # 7. точка остановки и запрет повторных заходов
    p06 = read(os.path.join(root, "06_ГДЕ_ОСТАНОВИЛИСЬ.md"))
    dead_ok = "mdadm --examine" in p06 and ("×11" in p06 or "11" in p06)
    add("Описаны тупики старого бота", 6, dead_ok,
        "перечислены повторы и альтернатива" if dead_ok
        else "нет таблицы тупиков или команды mdadm --examine")

    # 8. чекер отчёта: живой прогон на двух фикстурах
    checker = os.path.join(root, "tools", "check_report.py")
    fixtures = os.path.join(root, "tools", "make_fixtures.py")
    detail, ok = "", False
    tmp = tempfile.mkdtemp(prefix="pkg18_")
    try:
        subprocess.run([sys.executable, fixtures, "--out", tmp],
                       capture_output=True, text=True, timeout=180)
        good = os.path.join(tmp, "fixture_good.docx")
        bad = os.path.join(tmp, "fixture_bad.docx")
        if not (os.path.isfile(good) and os.path.isfile(bad)):
            detail = "фикстуры не создались (нет python-docx/Pillow?)"
        else:
            rg = subprocess.run([sys.executable, checker, good], capture_output=True, text=True)
            rb = subprocess.run([sys.executable, checker, bad], capture_output=True, text=True)
            sg = re.search(r"ИТОГ: (\d+/\d+) — (.+)", rg.stdout)
            sb = re.search(r"ИТОГ: (\d+/\d+) — (.+)", rb.stdout)
            ok = rg.returncode == 0 and rb.returncode == 2
            detail = "хороший: %s — %s (код %d); плохой: %s — %s (код %d)" % (
                sg.group(1) if sg else "?", sg.group(2) if sg else "нет итога", rg.returncode,
                sb.group(1) if sb else "?", sb.group(2) if sb else "нет итога", rb.returncode)
    except Exception as e:  # noqa: BLE001
        detail = "прогон не удался: %s" % e
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    add("Чекер ловит плохой отчёт и пропускает хороший", 18, ok, detail, critical=True)

    # 9. README ссылается на все файлы
    readme = read(os.path.join(root, "00_ЧИТАЙ_МЕНЯ.md"))
    lost = [f for f in REQUIRED
            if os.path.basename(f) != "00_ЧИТАЙ_МЕНЯ.md"
            and os.path.basename(f) not in readme]
    add("Карта пакета ссылается на все файлы", 6, not lost,
        ("не упомянуты: " + ", ".join(lost)) if lost else "все, кроме самой карты")

    # 10. нет мусора и кэша
    junk = []
    for dirpath, dirnames, filenames in os.walk(root):
        for d in list(dirnames):
            if d == "__pycache__":
                junk.append(os.path.join(dirpath, d))
        for f in filenames:
            if f.endswith((".pyc", ".tmp", ".bak")):
                junk.append(os.path.join(dirpath, f))
    add("В пакете нет мусора и кэша", 4, not junk,
        ("лишнее: " + ", ".join(os.path.relpath(j, root) for j in junk[:5])) if junk else "чисто")

    return finish(checks, args)


def finish(checks, args):
    total = sum(c["weight"] for c in checks)
    score = round(sum(c["weight"] for c in checks if c["ok"]) * 100 / total) if total else 0
    crit = [c for c in checks if c["critical"] and not c["ok"]]
    passed = not crit and score >= args.min

    line = "-" * 74
    print(line)
    print("%-48s %3s %6s" % ("ПРОВЕРКА ПАКЕТА", "ВЕС", "ИТОГ"))
    print(line)
    for c in checks:
        print("%-48s %3d %6s" % (c["name"][:48], c["weight"],
                                 "да" if c["ok"] else ("КРИТ" if c["critical"] else "нет")))
    print(line)
    for c in checks:
        if c["ok"]:
            print("  · %s — %s" % (c["name"], c["detail"]))
    for c in checks:
        if not c["ok"]:
            print("  ! %s — %s" % (c["name"], c["detail"]))
    print(line)
    print("ИТОГ: %d/100 — %s" % (score, "ПАКЕТ ГОДЕН" if passed else "ЕСТЬ ПРОБЛЕМЫ"))
    if crit:
        print("КРИТИЧЕСКИЕ: %s" % ", ".join(c["name"] for c in crit))
    return 0 if passed else 2


if __name__ == "__main__":
    sys.exit(main())
