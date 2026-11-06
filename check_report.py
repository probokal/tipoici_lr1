#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_report.py — проверка отчёта .docx по требованиям преподавателя (ТиПОиЦИ).

Только стандартная библиотека: pip не нужен.
    python3 tools/check_report.py ОТЧЕТ_ЛР2.docx
    python3 tools/check_report.py ОТЧЕТ_ЛР2.docx --min 80
    python3 tools/check_report.py ОТЧЕТ_ЛР2.docx --json

Коды возврата: 0 — принят, 2 — не готов, 1 — файл не читается.
"""

import argparse
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS = {"w": W, "a": A, "pic": PIC}


def q(tag, ns=W):
    return "{%s}%s" % (ns, tag)


# ---------------------------------------------------------------- разбор docx

class Doc:
    def __init__(self, path):
        self.path = path
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            if "word/document.xml" not in names:
                raise ValueError("в архиве нет word/document.xml — это не .docx")
            self.document = ET.fromstring(z.read("word/document.xml"))
            self.styles = (ET.fromstring(z.read("word/styles.xml"))
                           if "word/styles.xml" in names else ET.Element("styles"))
        self.style_map = self._read_styles()
        self.default_style = self._read_default_style()
        self.defaults = self._read_defaults()
        self.blocks = self._read_blocks()

    # --- стили
    def _read_styles(self):
        out = {}
        for st in self.styles.findall(q("style")):
            sid = st.get(q("styleId"))
            if not sid:
                continue
            rpr = st.find(q("rPr"))
            ppr = st.find(q("pPr"))
            base = st.find(q("basedOn"))
            out[sid] = {
                "font": self._font(rpr),
                "size": self._size(rpr),
                "line": self._line(ppr),
                "jc": self._jc(ppr),
                "basedOn": base.get(q("val")) if base is not None else None,
            }
        return out

    def _read_default_style(self):
        """Стиль абзаца по умолчанию (w:default="1", type="paragraph")."""
        for st in self.styles.findall(q("style")):
            if (st.get(q("type")) == "paragraph"
                    and st.get(q("default")) in ("1", "true")):
                return st.get(q("styleId"))
        return "Normal" if "Normal" in self.style_map else None

    def _read_defaults(self):
        d = {"font": None, "size": None, "line": None, "jc": None}
        dd = self.styles.find(q("docDefaults"))
        if dd is None:
            return d
        rpr = dd.find("%s/%s" % (q("rPrDefault"), q("rPr")))
        ppr = dd.find("%s/%s" % (q("pPrDefault"), q("pPr")))
        d["font"] = self._font(rpr)
        d["size"] = self._size(rpr)
        d["line"] = self._line(ppr)
        d["jc"] = self._jc(ppr)
        return d

    def _resolve(self, sid, key, depth=0):
        if not sid or sid not in self.style_map or depth > 5:
            return None
        st = self.style_map[sid]
        if st.get(key):
            return st[key]
        return self._resolve(st.get("basedOn"), key, depth + 1)

    # --- примитивы
    @staticmethod
    def _font(rpr):
        if rpr is None:
            return None
        f = rpr.find(q("rFonts"))
        if f is None:
            return None
        return f.get(q("ascii")) or f.get(q("hAnsi"))

    @staticmethod
    def _size(rpr):
        if rpr is None:
            return None
        s = rpr.find(q("sz"))
        if s is None:
            return None
        try:
            return int(s.get(q("val"))) / 2.0
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _line(ppr):
        if ppr is None:
            return None
        sp = ppr.find(q("spacing"))
        if sp is None:
            return None
        line = sp.get(q("line"))
        rule = sp.get(q("lineRule")) or "auto"
        if line is None:
            return None
        try:
            line = int(line)
        except ValueError:
            return None
        if rule == "auto":
            return round(line / 240.0, 2)
        return "%.1fтв" % (line / 20.0)

    @staticmethod
    def _jc(ppr):
        if ppr is None:
            return None
        j = ppr.find(q("jc"))
        return j.get(q("val")) if j is not None else None

    def _read_blocks(self):
        body = self.document.find(q("body"))
        blocks = []
        if body is None:
            return blocks
        for el in list(body):
            if el.tag == q("p"):
                blocks.append(self._para(el))
            elif el.tag == q("tbl"):
                blocks.append({"kind": "table", "text": "", "runs": [],
                               "drawings": 0, "framed": 0, "jc": None})
        return blocks

    def _para(self, p):
        ppr = p.find(q("pPr"))
        sid = None
        ps = ppr.find(q("pStyle")) if ppr is not None else None
        if ps is not None:
            sid = ps.get(q("val"))
        if sid is None:
            sid = self.default_style

        text_parts, runs = [], []
        for r in p.iter(q("r")):
            t = "".join((n.text or "") for n in r.iter(q("t")))
            if not t:
                continue
            text_parts.append(t)
            rpr = r.find(q("rPr"))
            font = self._font(rpr) or self._resolve(sid, "font") or self.defaults["font"]
            size = self._size(rpr) or self._resolve(sid, "size") or self.defaults["size"]
            runs.append({"text": t, "font": font, "size": size})

        drawings = list(p.iter(q("drawing"))) + list(p.iter(q("pict")))
        framed = 0
        for d in drawings:
            for ln in d.iter(q("ln", A)):
                if ln.get("w") and ln.get("w") != "0":
                    framed += 1
                    break

        return {
            "kind": "para",
            "text": "".join(text_parts).strip(),
            "runs": runs,
            "drawings": len(drawings),
            "framed": framed,
            "line": self._line(ppr) or self._resolve(sid, "line") or self.defaults["line"],
            "jc": self._jc(ppr) or self._resolve(sid, "jc") or self.defaults["jc"],
            "style": sid,
        }


# ---------------------------------------------------------------- проверки

CAPTION = re.compile(r"^\s*(?:Рисунок|Рис\.?)\s*(\d+)\s*[—–\-.:)]", re.IGNORECASE)
CAPTION_LOOSE = re.compile(r"^\s*(?:Рисунок|Рис\.?)\s*(\d+)", re.IGNORECASE)
REF = re.compile(r"рис(?:унке|унках|унок|унки|\.)\s*(\d+)(?:\s*[—–\-]\s*(\d+))?", re.IGNORECASE)
PLACEHOLDER = ["вставить позже", "плейсхолдер", "placeholder", "todo", "тодо",
               "[скрин", "[рисунок", "вставить рисунок", "вставить скрин",
               "xxx", "???", "____"]


def analyse(doc):
    paras = [b for b in doc.blocks if b["kind"] == "para"]
    index = {id(b): i for i, b in enumerate(doc.blocks)}

    # подписи и их номера
    captions, refs = [], []
    for i, b in enumerate(doc.blocks):
        if b["kind"] != "para":
            continue
        m = CAPTION.match(b["text"])
        if m:
            captions.append({"i": i, "n": int(m.group(1)), "b": b})
            continue
        for m in REF.finditer(b["text"]):
            a = int(m.group(1))
            z = int(m.group(2)) if m.group(2) else a
            if z < a or z - a > 20:
                z = a
            refs.extend(range(a, z + 1))

    # рисунки: блок с drawing
    drawings = [i for i, b in enumerate(doc.blocks) if b["kind"] == "para" and b["drawings"]]
    framed = sum(b["framed"] for b in doc.blocks if b["kind"] == "para" and b["drawings"])

    # шрифт / кегль / интервал / выравнивание по основному массиву текста
    def mode(values):
        vals = [v for v in values if v]
        if not vals:
            return None, 0
        counts = {}
        for v in vals:
            counts[v] = counts.get(v, 0) + 1
        best = max(counts.items(), key=lambda kv: kv[1])
        return best[0], best[1]

    body_paras = [b for b in doc.blocks
                  if b["kind"] == "para" and len(b["text"]) > 40
                  and not CAPTION_LOOSE.match(b["text"])]
    fonts = [r["font"] for b in body_paras for r in b["runs"]]
    sizes = [r["size"] for b in body_paras for r in b["runs"]]
    lines = [b["line"] for b in body_paras]
    jcs = [b["jc"] for b in body_paras]

    main_font, font_n = mode(fonts)
    main_size, size_n = mode(sizes)
    main_line, line_n = mode(lines)
    main_jc, jc_n = mode(jcs)

    return {
        "paras": paras, "blocks": doc.blocks, "captions": captions,
        "refs": refs, "drawings": drawings, "framed": framed,
        "main_font": main_font, "font_n": font_n, "font_total": len(fonts),
        "main_size": main_size, "size_n": size_n, "size_total": len(sizes),
        "main_line": main_line, "line_n": line_n, "line_total": len(lines),
        "main_jc": main_jc, "jc_n": jc_n, "jc_total": len(jcs),
    }


def run_checks(doc):
    a = analyse(doc)
    checks = []

    def add(name, weight, ok, detail, critical=False):
        checks.append({"name": name, "weight": weight,
                       "ok": bool(ok), "detail": detail, "critical": critical})

    # 1. шрифт
    ok = a["main_font"] and "Times New Roman" in a["main_font"]
    share = "%d/%d" % (a["font_n"], a["font_total"])
    add("Шрифт Times New Roman", 10, ok,
        ("основной шрифт: %s (%s)" % (a["main_font"], share)) if a["main_font"]
        else "шрифт не определён", critical=True)

    # 2. кегль
    ok = a["main_size"] == 14.0
    add("Кегль 14", 10, ok,
        "кегль: %s (%d/%d)" % (a["main_size"], a["size_n"], a["size_total"]), critical=True)

    # 3. интервал
    ok = a["main_line"] == 1.5
    add("Интервал 1,5", 8, ok,
        "междустрочный: %s (%d/%d)" % (a["main_line"], a["line_n"], a["line_total"]),
        critical=True)

    # 4. выравнивание
    ok = a["main_jc"] in ("both", "justify")
    add("Выравнивание по ширине", 6, ok,
        "выравнивание: %s (%d/%d)" % (a["main_jc"], a["jc_n"], a["jc_total"]))

    # 5. рисунки есть
    add("Рисунки найдены", 4, bool(a["drawings"]),
        "вставлено рисунков: %d" % len(a["drawings"]))

    # 6. у каждого рисунка подпись
    cap_after = {}
    for c in a["captions"]:
        for d in a["drawings"]:
            if 0 < c["i"] - d <= 3:
                cap_after[d] = c
                break
    missing = [d for d in a["drawings"] if d not in cap_after]
    detail = ("рисунков нет" if not a["drawings"] else
              "рисунков %d, подписей %d, без подписи: %d"
              % (len(a["drawings"]), len(a["captions"]), len(missing)))
    add("У каждого рисунка есть подпись", 10,
        bool(a["drawings"]) and not missing, detail, critical=True)

    # 7. сквозная нумерация
    nums = [c["n"] for c in a["captions"]]
    expected = list(range(1, len(nums) + 1))
    ok = nums == expected
    detail = "нумерация: %s" % (nums or "подписей нет")
    if nums and nums != expected:
        dup = sorted({n for n in nums if nums.count(n) > 1})
        gaps = [n for n in expected if n not in nums]
        if dup:
            detail += "; повторы: %s" % dup
        if gaps:
            detail += "; пропущены: %s" % gaps
        if nums != sorted(nums):
            detail += "; порядок нарушен"
    add("Сквозная нумерация рисунков", 10, bool(nums) and ok, detail, critical=True)

    # 8. подписи 12 по центру
    bad = []
    for c in a["captions"]:
        b = c["b"]
        size = None
        for r in b["runs"]:
            if r["size"]:
                size = r["size"]
                break
        centered = b["jc"] in ("center",)
        if size != 12.0 or not centered:
            bad.append("%d(кегль=%s, %s)" % (c["n"], size, b["jc"] or "нет выравн."))
    if not a["captions"]:
        add("Подписи 12 пт, по центру", 8, False, "подписей нет")
    else:
        add("Подписи 12 пт, по центру", 8, not bad,
            ("нарушения: " + "; ".join(bad[:6])) if bad else
            ("подписей %d, все 12 по центру" % len(a["captions"])))

    # 9. ссылка до объекта
    problems, unreferenced = [], []
    for c in a["captions"]:
        n = c["n"]
        hits = [i for i, b in enumerate(doc.blocks)
                if b["kind"] == "para" and not CAPTION.match(b["text"])
                and n in _refs_of(b["text"])]
        if not hits:
            unreferenced.append(n)
            continue
        if min(hits) > c["i"]:
            problems.append(
                "%d: первый раз упомянут в блоке %d, а подпись в %d — ссылка после объекта"
                % (n, min(hits), c["i"]))
    ok = bool(a["captions"]) and not problems and not unreferenced
    detail_parts = []
    if problems:
        detail_parts.append("; ".join(problems[:5]))
    if unreferenced:
        detail_parts.append("нет ссылки в тексте: рисунки %s" % unreferenced[:8])
    if not a["captions"]:
        detail = "подписей нет — проверять нечего"
    else:
        detail = " | ".join(detail_parts) if detail_parts else "все ссылки стоят до объектов"
    add("Ссылка на рисунок стоит ДО рисунка", 14, ok, detail, critical=True)

    # 10. названия таблиц сверху
    tbl_idx = [i for i, b in enumerate(doc.blocks) if b["kind"] == "table"]
    bad_tbl = []
    for t in tbl_idx:
        prev = doc.blocks[t - 1] if t > 0 else None
        if not prev or prev["kind"] != "para" or not re.search(r"таблиц", prev["text"], re.I):
            bad_tbl.append(t)
    if not tbl_idx:
        add("Название таблицы сверху справа", 8, True, "таблиц в документе нет")
    else:
        add("Название таблицы сверху справа", 8, not bad_tbl,
            ("без названия перед таблицей: %d" % len(bad_tbl)) if bad_tbl
            else "таблиц %d, у всех название сверху" % len(tbl_idx))

    # 11. плейсхолдеры
    found = []
    for b in doc.blocks:
        low = b["text"].lower()
        for ph in PLACEHOLDER:
            if ph in low:
                found.append(b["text"][:70])
                break
    add("Нет плейсхолдеров и «вставить позже»", 8, not found,
        ("найдено: " + " | ".join(found[:4])) if found else "чисто", critical=True)

    # 12. рамки
    add("Тонкие рамки у рисунков", 4,
        bool(a["drawings"]) and a["framed"] >= len(a["drawings"]),
        "с рамкой %d из %d" % (a["framed"], len(a["drawings"])))

    return checks, a


def _refs_of(text):
    out = set()
    for m in REF.finditer(text):
        a = int(m.group(1))
        z = int(m.group(2)) if m.group(2) else a
        if z < a or z - a > 20:
            z = a
        out.update(range(a, z + 1))
    return out


def main():
    ap = argparse.ArgumentParser(description="Проверка отчёта .docx")
    ap.add_argument("path")
    ap.add_argument("--min", type=int, default=80, help="порог баллов (по умолчанию 80)")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    try:
        doc = Doc(args.path)
    except (zipfile.BadZipFile, ValueError, KeyError, ET.ParseError) as e:
        print("ОШИБКА: не удалось прочитать %s — %s" % (args.path, e))
        print("Лечится так: открыть файл в Word и сохранить как .docx заново.")
        return 1

    checks, a = run_checks(doc)
    total = sum(c["weight"] for c in checks)
    score = sum(c["weight"] for c in checks if c["ok"])
    crit = [c for c in checks if c["critical"] and not c["ok"]]
    passed = (not crit) and score >= args.min
    verdict = "ПРИНЯТ" if passed else "НЕ ГОТОВ"

    if args.json:
        print(json.dumps({"file": args.path, "score": score, "total": total,
                          "verdict": verdict,
                          "critical_failed": [c["name"] for c in crit],
                          "checks": checks}, ensure_ascii=False, indent=2))
        return 0 if passed else 2

    line = "-" * 78
    print("ФАЙЛ : %s" % args.path)
    print("СТАТ : %d абзацев, %d рисунков, %d подписей, %d таблиц"
          % (len([b for b in doc.blocks if b["kind"] == "para"]),
             len(a["drawings"]), len(a["captions"]),
             len([b for b in doc.blocks if b["kind"] == "table"])))
    print(line)
    print("%-42s %3s %6s" % ("ПРОВЕРКА", "ВЕС", "ИТОГ"))
    print(line)
    for c in checks:
        mark = "да" if c["ok"] else ("КРИТ" if c["critical"] else "нет")
        print("%-42s %3d %6s" % (c["name"][:42], c["weight"], mark))
    print(line)
    print("ДЕТАЛИ")
    for c in checks:
        if not c["ok"]:
            print("  · %s — %s" % (c["name"], c["detail"]))
    print(line)
    print("ИТОГ: %d/%d — %s" % (score, total, verdict))
    if crit:
        print("КРИТИЧЕСКИЕ: %s" % ", ".join(c["name"] for c in crit))
        print("Отчёт с этим не принимается, сколько бы баллов ни набралось.")
    elif score < args.min:
        print("Порог %d не набран — доработать и прогнать заново." % args.min)
    return 0 if passed else 2


if __name__ == "__main__":
    sys.exit(main())
