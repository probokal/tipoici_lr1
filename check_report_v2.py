#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_report.py v2 — проверка отчёта по требованиям Травкина.

Работает с .docx (только стандартная библиотека) и с .pdf (нужен pypdf).
Главное отличие от v1: ловит рисунки, взятые из другой работы, по хэшу.

    python3 tools/check_report.py ОТЧЕТ_ЛР1.docx
    python3 tools/check_report.py ОТЧЕТ_ЛР1.pdf --compare ОТЧЕТ_ЛР2.pdf
    python3 tools/check_report.py отчет.pdf --compare провал_бота/  --json
    python3 tools/check_report.py отчет.docx --min 85

--compare — файл или папка другой работы: любой рисунок, чей хэш там встречается,
считается чужим (критическая ошибка).

Коды возврата: 0 — принят, 2 — не готов, 1 — файл не читается.
"""

import argparse
import hashlib
import io
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def q(tag, ns=W):
    return "{%s}%s" % (ns, tag)


# ─────────────────────────────────────────────────── текстовые признаки

CAPTION_STD = re.compile(r"^\s*Рис\.\s*(\d+)\.", re.IGNORECASE)          # эталон
CAPTION_ALT = re.compile(r"^\s*Рисунок\s+(\d+)\s*[—–-]", re.IGNORECASE)  # допустимо, но без точки — ошибка
CAPTION_ANY = re.compile(r"^\s*(?:Рис\.|Рисунок)\s*(\d+)", re.IGNORECASE)
REF = re.compile(r"[Рр]ис(?:унке|унках|унок|унки|\.)\s*(\d+)(?:\s*[–—-]\s*(\d+))?", re.IGNORECASE)

PLACEHOLDERS = ["[скрин", "[рисунок", "вставить позже", "вставить рисунок",
                "вставить скрин", "обновить поле", "плейсхолдер", "placeholder",
                "todo", "xxx", "???"]

CONSOLE = re.compile(
    r"(?<![\w-])(?:"
    r"(?:sudo|apt-get|chkdsk|mmls|fls|icat)\s+[\w/]"          # команды, которые всегда команда
    r"|(?:mdadm|ewfmount|losetup|umount|blockdev|systemctl|fidentify)\s+(?:--?[a-z]|/dev|/mnt)"
    r"|dd\s+if="
    r"|mount\s+-o\s"
    r")(?![\w-])", re.IGNORECASE)

BANNED_PHRASES = [
    ("при помощи консольной команды", "Травкин запретил эту формулировку"),
    ("с помощью консольной команды", "Травкин запретил эту формулировку"),
    ("результат выглядит следующим образом", "анонс перед рисунком запрещён"),
    ("выглядит следующим образом", "анонс перед рисунком запрещён"),
    ("РЕСУРС", "служебный блок из чата не переносится в отчёт"),
    ("[ЗАПУСК]", "метка из чата не переносится в отчёт"),
    ("[НЕ ПРОВЕРЕНО]", "метка из чата не переносится в отчёт"),
    ("должно работать", "предположение вместо факта"),
    ("по идее", "предположение вместо факта"),
]

PERSONAL = re.compile(r"\b(я (открыл|запустил|сделал|получил|увидел|нашёл|нашел)|"
                      r"мне (удалось|не удалось)|у меня (получилось|не получилось))\b",
                      re.IGNORECASE)

SECTIONS = [
    ("ЦЕЛЬ И ЗАДАЧИ", r"ЦЕЛЬ\s+И\s+ЗАДАЧИ"),
    ("АППАРАТНОЕ ОБЕСПЕЧЕНИЕ", r"АППАРАТНОЕ\s+ОБЕСПЕЧЕНИЕ"),
    ("ПРОГРАММНОЕ ОБЕСПЕЧЕНИЕ", r"ПРОГРАММНОЕ\s+ОБЕСПЕЧЕНИЕ"),
    ("ВЫВОД", r"\bВЫВОД\b"),
]
TASK_HEAD = re.compile(r"ЗАДАНИЕ\s*(\d+)\s*[:.]", re.IGNORECASE)


# ─────────────────────────────────────────────────── чтение файлов

class Report:
    """Единое представление отчёта из docx или pdf."""

    def __init__(self, path):
        self.path = path
        self.kind = "pdf" if path.lower().endswith(".pdf") else "docx"
        self.blocks = []      # [{"kind": "para"|"table", "text", "runs", "drawings", "framed", "jc"}]
        self.images = []      # [(sha1, bytes_len, index)]
        self.styles_ok = False
        if self.kind == "docx":
            self._read_docx()
        else:
            self._read_pdf()

    # --- docx
    def _read_docx(self):
        with zipfile.ZipFile(self.path) as z:
            names = set(z.namelist())
            if "word/document.xml" not in names:
                raise ValueError("в архиве нет word/document.xml")
            self.style_map, self.default_style, self.docdefaults = self._styles(z, names)
            doc = ET.fromstring(z.read("word/document.xml"))
            for name in sorted(n for n in names if n.startswith("word/media/")
                               and not n.endswith("/")):
                data = z.read(name)
                self.images.append((hashlib.sha1(data).hexdigest(), len(data), name))
        body = doc.find(q("body"))
        for el in list(body):
            if el.tag == q("p"):
                self.blocks.append(self._para(el))
            elif el.tag == q("tbl"):
                self.blocks.append({"kind": "table", "text": "", "runs": [],
                                    "drawings": 0, "framed": 0, "jc": None})
        self.styles_ok = True

    def _styles(self, z, names):
        smap, default, dd = {}, None, {"font": None, "size": None, "line": None, "jc": None}
        if "word/styles.xml" in names:
            st = ET.fromstring(z.read("word/styles.xml"))
            for s in st.findall(q("style")):
                sid = s.get(q("styleId"))
                rpr, ppr = s.find(q("rPr")), s.find(q("pPr"))
                smap[sid] = {"font": self._font(rpr), "size": self._size(rpr),
                             "line": self._line(ppr), "jc": self._jc(ppr),
                             "basedOn": (s.find(q("basedOn")).get(q("val"))
                                         if s.find(q("basedOn")) is not None else None)}
                if s.get(q("type")) == "paragraph" and s.get(q("default")) in ("1", "true"):
                    default = sid
            d = st.find(q("docDefaults"))
            if d is not None:
                rpr = d.find("%s/%s" % (q("rPrDefault"), q("rPr")))
                ppr = d.find("%s/%s" % (q("pPrDefault"), q("pPr")))
                dd = {"font": self._font(rpr), "size": self._size(rpr),
                      "line": self._line(ppr), "jc": self._jc(ppr)}
        return smap, default or ("Normal" if "Normal" in smap else None), dd

    def _resolve(self, sid, key, depth=0):
        if not sid or sid not in self.style_map or depth > 5:
            return None
        s = self.style_map[sid]
        return s.get(key) or self._resolve(s.get("basedOn"), key, depth + 1)

    @staticmethod
    def _font(rpr):
        if rpr is None:
            return None
        f = rpr.find(q("rFonts"))
        return (f.get(q("ascii")) or f.get(q("hAnsi"))) if f is not None else None

    @staticmethod
    def _size(rpr):
        if rpr is None:
            return None
        s = rpr.find(q("sz"))
        try:
            return int(s.get(q("val"))) / 2.0 if s is not None else None
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _line(ppr):
        if ppr is None:
            return None
        sp = ppr.find(q("spacing"))
        if sp is None or sp.get(q("line")) is None:
            return None
        try:
            val = int(sp.get(q("line")))
        except ValueError:
            return None
        return round(val / 240.0, 2) if (sp.get(q("lineRule")) or "auto") == "auto" else None

    @staticmethod
    def _jc(ppr):
        if ppr is None:
            return None
        j = ppr.find(q("jc"))
        return j.get(q("val")) if j is not None else None

    def _para(self, p):
        ppr = p.find(q("pPr"))
        sid = self.default_style
        ps = ppr.find(q("pStyle")) if ppr is not None else None
        if ps is not None:
            sid = ps.get(q("val"))
        parts, runs = [], []
        for r in p.iter(q("r")):
            t = "".join((n.text or "") for n in r.iter(q("t")))
            if not t:
                continue
            parts.append(t)
            rpr = r.find(q("rPr"))
            runs.append({"text": t,
                         "font": self._font(rpr) or self._resolve(sid, "font") or self.docdefaults["font"],
                         "size": self._size(rpr) or self._resolve(sid, "size") or self.docdefaults["size"]})
        drawings = len(list(p.iter(q("drawing"))))
        framed = 0
        for d in p.iter(q("drawing")):
            for ln in d.iter(q("ln", A)):
                if ln.get("w") and ln.get("w") != "0":
                    framed += 1
                    break
        return {"kind": "para", "text": "".join(parts), "runs": runs,
                "drawings": drawings, "framed": framed,
                "jc": self._jc(ppr) or self._resolve(sid, "jc") or self.docdefaults["jc"],
                "line": self._line(ppr) or self._resolve(sid, "line") or self.docdefaults["line"]}

    # --- pdf
    def _read_pdf(self):
        try:
            import pypdf  # noqa: PLC0415
        except ImportError:
            raise RuntimeError(
                "для .pdf нужен pypdf: python -m pip install pypdf\n"
                "либо сохрани отчёт как .docx — тогда хватит стандартной библиотеки")
        r = pypdf.PdfReader(self.path)
        for page in r.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                self.blocks.append({"kind": "para", "text": line, "runs": [],
                                    "drawings": 0, "framed": 0, "jc": None, "line": None})
            try:
                for im in page.images:
                    data = im.data
                    self.images.append((hashlib.sha1(data).hexdigest(), len(data), im.name))
            except Exception:  # noqa: BLE001
                pass


# ─────────────────────────────────────────────────── анализ

CAPTION_SPAN = re.compile(r"Рис\.\s*(\d+)\.\s*([^.\n]{0,200}?)\s*\.", re.IGNORECASE)


def collect(rep):
    """Склеивает текст в одну строку: подписи в PDF рвутся переносом, а в docx
    каждая подпись — отдельный абзац. Единая строка лечит оба случая."""
    raw = "\n".join(b["text"] for b in rep.blocks)
    text_all = re.sub(r"\s+", " ", raw).strip()

    captions = []
    spans = []
    for m in CAPTION_SPAN.finditer(text_all):
        n = int(m.group(1))
        body = (m.group(2) or "").strip()
        captions.append({"n": n, "text": " ".join(["Рис.%d." % n, body + "."]),
                         "start": m.start(), "end": m.end(),
                         "std": True, "dot": True})
        spans.append((m.start(), m.end()))

    # подписи в альтернативном формате «Рисунок N — …» (без точки) — считаем отдельно
    for m in re.finditer(r"Рисунок\s+(\d+)\s*[—–-]\s*([^.\n]{0,200}?)(?=\.|$)", text_all, re.I):
        if any(a <= m.start() < b for a, b in spans):
            continue
        captions.append({"n": int(m.group(1)), "text": m.group(0).strip(),
                         "start": m.start(), "end": m.end(), "std": False,
                         "dot": (m.group(2) or "").strip().endswith(".")})
        spans.append((m.start(), m.end()))

    refs = {}
    for m in REF.finditer(text_all):
        if any(a <= m.start() < b for a, b in spans):
            continue
        a = int(m.group(1))
        z = int(m.group(2)) if m.group(2) else a
        if z < a or z - a > 25:
            z = a
        for k in range(a, z + 1):
            refs.setdefault(k, m.start())
    return text_all, captions, refs


def run_checks(rep, banned_hashes):
    text, captions, refs = collect(rep)
    checks = []

    def add(name, w, ok, detail, critical=False, na=False):
        checks.append({"name": name, "weight": w, "ok": bool(ok) or na,
                       "detail": detail, "critical": critical and not na, "na": na})

    low = text.lower()

    # 1. структура
    missing_sec = [n for n, rx in SECTIONS if not re.search(rx, text, re.IGNORECASE | re.MULTILINE)]
    add("Разделы отчёта на месте", 10, not missing_sec,
        ("нет: " + ", ".join(missing_sec)) if missing_sec else "цель, аппаратное, программное, вывод", critical=True)

    tasks = sorted({int(m.group(1)) for m in TASK_HEAD.finditer(text)})
    add("Есть разделы «ЗАДАНИЕ N»", 8, len(tasks) >= 2,
        ("найдено: %s" % tasks) if tasks else "ни одного — в эталоне это каркас отчёта", critical=True)

    add("Есть СОДЕРЖАНИЕ", 5, "содержание" in low,
        "оглавление найдено" if "содержание" in low else "нет раздела «СОДЕРЖАНИЕ»")

    # 2. заглушки
    found = []
    for ph in PLACEHOLDERS:
        if ph in low:
            found.append(ph)
    add("Нет заглушек и плейсхолдеров", 14, not found,
        ("найдено: " + ", ".join(found[:5])) if found else "чисто", critical=True)

    # 3. рисунки: подписи
    if not captions:
        add("Подписи к рисункам", 10, False, "подписей не найдено", critical=True)
    else:
        nodot = [c["n"] for c in captions if not c["dot"]]
        alt = [c["n"] for c in captions if not c["std"]]
        add("Точка в конце каждой подписи", 8, not nodot,
            ("без точки: %s" % nodot[:10]) if nodot else "все подписи с точкой", critical=True)
        add("Формат подписи «Рис.N.» как в эталоне", 5, not alt,
            ("не по эталону: %s" % alt[:8]) if alt else "все в формате Рис.N.")
        if rep.styles_ok:
            bad = [c["n"] for c in captions
                   if (c["size"] and c["size"] != 12.0) or c["jc"] not in ("center", None)]
            add("Подписи 12 пт по центру", 6, not bad,
                ("нарушения: %s" % bad[:6]) if bad else "все 12 пт по центру")
        else:
            add("Подписи 12 пт по центру", 6, True, "в PDF не проверяется", na=True)

    # 4. ссылки до рисунков
    # Травкин ссылку перед каждым рисунком не требует: в эталоне она есть у 1 подписи из 43.
    # Критично только одно — ссылка, стоящая ПОСЛЕ своего рисунка.
    after, before = [], 0
    for c in captions:
        pos = refs.get(c["n"])
        if pos is None:
            continue
        if pos > c["start"]:
            after.append(c["n"])
        else:
            before += 1
    share = before / len(captions) if captions else 0
    add("Нет ссылок после рисунка", 12, not after,
        ("ссылка после рисунка: %s" % after[:8]) if after
        else "ссылок до рисунка: %d из %d подписей" % (before, len(captions)),
        critical=bool(after))
    add("Текст ссылается на рисунки", 4, share >= 0.5,
        "ссылок до рисунка: %d из %d (%.0f%%)" % (before, len(captions), share * 100))

    # 5. картинки vs подписи
    n_img, n_cap = len(rep.images), len(captions)
    short = n_cap > n_img              # подписи без картинок — это заглушки
    add("Каждой подписи соответствует картинка", 10,
        n_img > 0 and not short,
        "картинок %d, подписей %d%s" % (n_img, n_cap,
                                        (" — %d рисунков без картинки" % (n_cap - n_img))
                                        if short else
                                        ("" if n_img == n_cap else
                                         " (лишние %d — служебные, не ошибка)" % (n_img - n_cap))),
        critical=True)

    # 6. дубли внутри документа
    dups = {}
    for h, n, name in rep.images:
        dups.setdefault(h, []).append(name)
    inner = [v for v in dups.values() if len(v) > 1]
    add("Внутри отчёта нет одинаковых картинок", 6, not inner,
        ("повторяются: %s" % inner[:3]) if inner else "все картинки различаются")

    # 7. чужие картинки из другой работы
    if banned_hashes is None:
        add("Нет рисунков из другой работы", 15, True, "проверка не запрошена (--compare)", na=True)
    else:
        stolen = [(h, n) for h, n, name in rep.images if h in banned_hashes]
        add("Нет рисунков из другой работы", 15, not stolen,
            ("совпало с другой работой: %d рисунков (хэши %s)"
             % (len(stolen), ", ".join(h[:8] for h, _ in stolen[:3])))
            if stolen else "ни один рисунок не совпал с другой работой", critical=True)

    # 8. консольные команды
    cmds = [m.group(0).strip() for m in CONSOLE.finditer(text)]
    add("В тексте нет консольных команд", 10, not cmds,
        ("найдено: " + ", ".join(sorted(set(cmds))[:8])) if cmds else "команд в тексте нет",
        critical=True)

    # 9. запрещённые формулировки
    hits = [(p, why) for p, why in BANNED_PHRASES if p.lower() in low]
    add("Нет запрещённых формулировок", 8, not hits,
        ("; ".join("%s (%s)" % (p, why) for p, why in hits[:4])) if hits else "чисто",
        critical=False)

    # 10. безличное изложение
    pers = sorted({m.group(0) for m in PERSONAL.finditer(text)})
    add("Изложение безличное", 5, not pers,
        ("личные обороты: %s" % pers[:5]) if pers else "личных оборотов нет")

    # 11. оформление (только docx)
    if rep.styles_ok:
        body = [b for b in rep.blocks if b["kind"] == "para" and len(b["text"]) > 40
                and not CAPTION_ANY.match(b["text"])]
        fonts = [r["font"] for b in body for r in b["runs"]]
        sizes = [r["size"] for b in body for r in b["runs"]]
        lines = [b["line"] for b in body if b["line"]]
        jcs = [b["jc"] for b in body if b["jc"]]

        def mode(vals):
            vals = [v for v in vals if v]
            if not vals:
                return None
            return max(set(vals), key=vals.count)

        f, s, ln, jc = mode(fonts), mode(sizes), mode(lines), mode(jcs)
        add("Шрифт Times New Roman", 8, bool(f) and "Times New Roman" in str(f),
            "основной шрифт: %s" % f, critical=True)
        add("Кегль 14", 8, s == 14.0, "кегль: %s" % s, critical=True)
        add("Интервал 1,5", 6, ln == 1.5, "междустрочный: %s" % ln, critical=True)
        add("Выравнивание по ширине", 5, jc in ("both", "justify"), "выравнивание: %s" % jc)
        framed = sum(b["framed"] for b in rep.blocks)
        add("Рисунки в тонкой рамке", 4, n_img > 0 and framed >= n_img,
            "рамка у %d из %d" % (framed, n_img))
    else:
        for name in ("Шрифт Times New Roman", "Кегль 14", "Интервал 1,5",
                     "Выравнивание по ширине", "Рисунки в тонкой рамке"):
            add(name, 0, True, "в PDF не проверяется — проверь в Word", na=True)

    return checks, {"images": n_img, "captions": n_cap, "tasks": tasks,
                    "refs": len(refs), "chars": len(text)}


def _refs_of(text):
    out = set()
    for m in REF.finditer(text):
        a = int(m.group(1))
        z = int(m.group(2)) if m.group(2) else a
        if z < a or z - a > 25:
            z = a
        out.update(range(a, z + 1))
    return out


def hashes_from(path):
    """Хэши картинок из файла или папки (для --compare)."""
    out = set()
    targets = []
    if os.path.isdir(path):
        for dp, _, fns in os.walk(path):
            targets += [os.path.join(dp, f) for f in fns
                        if f.lower().endswith((".docx", ".pdf"))]
    else:
        targets = [path]
    for t in targets:
        try:
            r = Report(t)
            out.update(h for h, _, _ in r.images)
        except Exception:  # noqa: BLE001
            pass
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--compare", action="append", default=[],
                    help="файл или папка другой работы: её рисунки в отчёте запрещены")
    ap.add_argument("--min", type=int, default=80)
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    try:
        rep = Report(args.path)
    except Exception as e:  # noqa: BLE001
        print("ОШИБКА: не читается %s — %s" % (args.path, e))
        return 1

    banned = None
    if args.compare:
        banned = set()
        for c in args.compare:
            banned |= hashes_from(c)

    checks, stat = run_checks(rep, banned)
    total = sum(c["weight"] for c in checks)
    score = round(sum(c["weight"] for c in checks if c["ok"]) * 100 / total) if total else 0
    crit = [c for c in checks if c["critical"] and not c["ok"]]
    passed = not crit and score >= args.min

    if args.json:
        print(json.dumps({"file": args.path, "score": score, "verdict":
                          "ПРИНЯТ" if passed else "НЕ ГОТОВ",
                          "critical": [c["name"] for c in crit],
                          "stat": stat, "checks": checks}, ensure_ascii=False, indent=2))
        return 0 if passed else 2

    line = "-" * 78
    print(line)
    print("ФАЙЛ: %s" % args.path)
    print("текст: %d симв. · картинок: %d · подписей: %d · задания: %s"
          % (stat["chars"], stat["images"], stat["captions"], stat["tasks"] or "нет"))
    print(line)
    for c in checks:
        mark = "н/д" if c["na"] else ("да" if c["ok"] else ("КРИТ" if c["critical"] else "нет"))
        print("%-44s %3d %6s" % (c["name"][:44], c["weight"], mark))
    print(line)
    for c in checks:
        if not c["ok"]:
            print(("✗ " if c["critical"] else "! ") + c["name"] + " — " + c["detail"])
    print(line)
    print("ИТОГ: %d/100 — %s" % (score, "ПРИНЯТ" if passed else "НЕ ГОТОВ"))
    if crit:
        print("КРИТИЧЕСКИЕ: " + ", ".join(c["name"] for c in crit))
        print("С этим отчёт не принимается, сколько бы баллов ни набрал.")
    return 0 if passed else 2


if __name__ == "__main__":
    raise SystemExit(main())
