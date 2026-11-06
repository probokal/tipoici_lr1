#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_fixtures.py — две учебные .docx для проверки tools/check_report.py.

    python3 tools/make_fixtures.py            # создаст хороший и плохой отчёт в /tmp
    python3 tools/make_fixtures.py --out .    # рядом с собой

Хороший: TNR 14, 1,5, по ширине, ссылки до рисунков, подписи 12 по центру, рамки.
Плохой: Calibri 12, одинарный, ссылки после объектов, нумерация 1,3,4, плейсхолдер.
Нужен python-docx и Pillow (только для генерации тестовых файлов).
"""

import argparse
import io
import os
import sys

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image, ImageDraw


def _png(path, size=(640, 360), label="screen"):
    img = Image.new("RGB", size, (245, 245, 245))
    d = ImageDraw.Draw(img)
    d.rectangle([8, 8, size[0] - 8, size[1] - 8], outline=(120, 120, 120), width=2)
    d.text((24, 24), label, fill=(30, 30, 30))
    img.save(path)
    return path


def _frame(pic_run):
    """Тонкая чёрная рамка вокруг рисунка (как ставит Word: pic:spPr/a:ln)."""
    for sp in pic_run._r.iter(qn("pic:spPr")):
        ln = sp.makeelement(qn("a:ln"), {"w": "9525"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        clr = fill.makeelement(qn("a:srgbClr"), {"val": "000000"})
        fill.append(clr)
        ln.append(fill)
        sp.append(ln)


def build(path, good=True):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman" if good else "Calibri"
    style.font.size = Pt(14) if good else Pt(12)
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rf)
    rf.set(qn("w:ascii"), style.font.name)
    rf.set(qn("w:hAnsi"), style.font.name)
    pf = style.paragraph_format
    pf.line_spacing = 1.5 if good else 1.0
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if good else WD_ALIGN_PARAGRAPH.LEFT

    def p(text, **kw):
        par = doc.add_paragraph(text)
        if not good:
            par.paragraph_format.line_spacing = 1.0
            par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return par

    tmp = os.path.join(os.path.dirname(path) or ".", "_tmp_fig.png")
    pic = _png(tmp)

    p("Отчёт по лабораторной работе. Раздел первый. " * 3)
    p("На рисунке 1 представлен общий вид окна программы. " + "Текст пояснения. " * 4)
    r = doc.add_paragraph().add_run()
    r.add_picture(pic, width=Pt(300))
    if good:
        _frame(r)
    cap = doc.add_paragraph("Рисунок 1 — Общий вид окна")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER if good else WD_ALIGN_PARAGRAPH.LEFT
    for run in cap.runs:
        run.font.size = Pt(12) if good else Pt(10)

    p("В таблице 1 приведены параметры объекта. " + "Текст. " * 5)
    from docx.enum.text import WD_ALIGN_PARAGRAPH as AL
    ttl = doc.add_paragraph("Таблица 1 — Параметры объекта")
    ttl.alignment = AL.RIGHT if good else AL.LEFT
    t = doc.add_table(rows=3, cols=3)
    t.cell(0, 0).text = "Параметр"
    t.cell(0, 1).text = "Значение"
    t.cell(0, 2).text = "Ед."

    if good:
        p("Далее на рисунке 2 показан результат поиска.")
        r2 = doc.add_paragraph().add_run()
        r2.add_picture(pic, width=Pt(300))
        _frame(r2)
        c2 = doc.add_paragraph("Рисунок 2 — Результат поиска")
        c2.alignment = AL.CENTER
        for run in c2.runs:
            run.font.size = Pt(12)

        p("Завершает работу рисунок 3 с итоговой сводкой.")
        r3 = doc.add_paragraph().add_run()
        r3.add_picture(pic, width=Pt(300))
        _frame(r3)
        c3 = doc.add_paragraph("Рисунок 3 — Итоговая сводка")
        c3.alignment = AL.CENTER
        for run in c3.runs:
            run.font.size = Pt(12)
    else:
        r2 = doc.add_paragraph().add_run()
        r2.add_picture(pic, width=Pt(300))
        c2 = doc.add_paragraph("Рисунок 2 — Результат поиска (вставить позже)")
        for run in c2.runs:
            run.font.size = Pt(10)
        p("Как видно на рисунке 2, поиск завершён.")
        r3 = doc.add_paragraph().add_run()
        r3.add_picture(pic, width=Pt(300))
        c3 = doc.add_paragraph("Рисунок 4 — Итоговая сводка")
        for run in c3.runs:
            run.font.size = Pt(10)

    doc.save(path)
    if os.path.exists(tmp):
        os.remove(tmp)
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="/tmp")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)
    a = build(os.path.join(args.out, "fixture_good.docx"), True)
    b = build(os.path.join(args.out, "fixture_bad.docx"), False)
    print("готово:\n  ", a, "\n  ", b)
    return 0


if __name__ == "__main__":
    sys.exit(main())
