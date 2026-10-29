"""Создание Word-протоколов: по шаблону с метками {{...}} или встроенный образец."""
import datetime
import os
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from .calcd import fmt

POINT_MARK = "{{т."  # строка таблицы с такими метками повторяется для каждой точки


def channel_fields(ch, digits=4):
    f = lambda v: fmt(v, digits)
    r = ch.result
    return {
        "№": ch.num, "KKS": ch.kks, "НАИМЕНОВАНИЕ": ch.name,
        "НИЖН": fmt(ch.low), "ВЕРХН": fmt(ch.high), "ЕД": ch.unit,
        "ДИАПАЗОН": f"{fmt(ch.low)} … {fmt(ch.high)} {ch.unit}".strip(),
        "ПИП": ch.pip, "ПОГР_ПИП": f(ch.pip_value), "МОДУЛЬ": ch.module,
        "ПОГР_МОДУЛЯ": f(ch.module_err), "ДОПУСТ": f(ch.allowed_value), "ЗАМЕТКИ": ch.notes,
        "a1": f(r.a1) if r else "", "a2": f(r.a2) if r else "",
        "b1": f(r.b1) if r else "", "b2": f(r.b2) if r else "",
        "gh": f(r.gh) if r else "", "gl": f(r.gl) if r else "",
        "ЭТ_ИК": f(ch.et_ik), "ОБЩАЯ": f(ch.total), "РЕЗУЛЬТАТ": ch.verdict,
        "ДАТА": datetime.date.today().strftime("%d.%m.%Y"),
    }


def point_fields(ch, i, digits=4):
    f = lambda v: fmt(v, digits)
    p = ch.points[i]
    rp = ch.result.points[i] if ch.result else None
    am = sum(p.measured) / len(p.measured) if p.measured else None
    return {
        "т.№": str(i + 1), "т.%": fmt(p.pct), "т.вход": fmt(p.input),
        "т.xd": fmt(p.xd), "т.изм": fmt(am),
        "т.ds": f(rp.ds) if rp else "", "т.dh": f(rp.dh) if rp else "",
        "т.dl": f(rp.dl) if rp else "",
        "т.верх": f(rp.upper) if rp else "", "т.низ": f(rp.lower) if rp else "",
    }


def _replace_in_paragraph(par, values):
    text = "".join(r.text for r in par.runs)
    if "{{" not in text:
        return
    new = re.sub(r"\{\{\s*([^{}]+?)\s*\}\}", lambda m: values.get(m.group(1), m.group(0)), text)
    if new == text:
        return
    # метка может быть разбита Word'ом на несколько фрагментов: текст в первый
    par.runs[0].text = new
    for r in par.runs[1:]:
        r.text = ""


def _iter_paragraphs(container):
    for p in container.paragraphs:
        yield p
    for t in container.tables:
        for row in t.rows:
            for cell in row.cells:
                yield from _iter_paragraphs(cell)


def _row_text(row):
    return "".join(c.text for c in row.cells)


def _expand_point_rows(doc, ch, digits):
    import copy
    tables = list(doc.tables)
    for t in tables:
        for row in list(t.rows):
            if POINT_MARK not in _row_text(row):
                continue
            tr = row._tr
            anchor = tr
            for i in range(len(ch.points)):
                new_tr = copy.deepcopy(tr)
                anchor.addnext(new_tr)
                anchor = new_tr
            tr.getparent().remove(tr)
            break
    # заполнить скопированные строки по порядку
    for t in doc.tables:
        idx = 0
        for row in t.rows:
            if POINT_MARK in _row_text(row) and idx < len(ch.points):
                vals = point_fields(ch, idx, digits)
                for cell in row.cells:
                    for p in _iter_paragraphs(cell):
                        _replace_in_paragraph(p, vals)
                idx += 1


def from_template(template, ch, digits=4):
    doc = Document(template)
    _expand_point_rows(doc, ch, digits)
    vals = channel_fields(ch, digits)
    for p in _iter_paragraphs(doc):
        _replace_in_paragraph(p, vals)
    for s in doc.sections:
        for p in _iter_paragraphs(s.header):
            _replace_in_paragraph(p, vals)
        for p in _iter_paragraphs(s.footer):
            _replace_in_paragraph(p, vals)
    return doc


def builtin(ch, digits=4):
    """Временный протокол, пока нет шаблона заказчика."""
    v = channel_fields(ch, digits)
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"
    st.font.size = Pt(12)
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h.add_run("ПРОТОКОЛ\nаттестации измерительного канала").bold = True

    info = doc.add_table(rows=0, cols=2)
    info.style = "Table Grid"
    for k, label in (("KKS", "Позиционное обозначение (KKS)"), ("НАИМЕНОВАНИЕ", "Наименование ИК"),
                     ("ДИАПАЗОН", "Диапазон измерения"), ("ПИП", "ПИП"), ("МОДУЛЬ", "Модуль"),
                     ("ДАТА", "Дата")):
        c = info.add_row().cells
        c[0].text, c[1].text = label, v[k]

    doc.add_paragraph()
    doc.add_paragraph().add_run("Результаты измерений и расчёта погрешности").bold = True
    cols = ("№", "КТ, %", "Действит. знач.", "Измер. знач.", "ds", "Δ верх.", "Δ нижн.")
    t = doc.add_table(rows=1, cols=len(cols))
    t.style = "Table Grid"
    for c, name in zip(t.rows[0].cells, cols):
        c.text = name
    for i in range(len(ch.points)):
        pf = point_fields(ch, i, digits)
        cells = t.add_row().cells
        for c, k in zip(cells, ("т.№", "т.%", "т.xd", "т.изм", "т.ds", "т.dh", "т.dl")):
            c.text = pf[k]

    doc.add_paragraph()
    res = doc.add_table(rows=0, cols=2)
    res.style = "Table Grid"
    for label, k in (("a1", "a1"), ("a2", "a2"), ("b1", "b1"), ("b2", "b2"),
                     ("gh (верхняя граница), %", "gh"), ("gl (нижняя граница), %", "gl"),
                     ("Приведённая погрешность ПИП, %", "ПОГР_ПИП"),
                     ("Приведённая погрешность ЭТ ИК, %", "ЭТ_ИК"),
                     ("Общая погрешность ИК, %", "ОБЩАЯ"),
                     ("Допустимая погрешность ИК, %", "ДОПУСТ")):
        c = res.add_row().cells
        c[0].text, c[1].text = label, v[k]
    doc.add_paragraph()
    p = doc.add_paragraph("Заключение: ")
    p.add_run(v["РЕЗУЛЬТАТ"] or "—").bold = True
    return doc


def safe_name(s):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", s).strip(" .")
    return s or "канал"


def generate(channels, out_dir, template=None, digits=4):
    """Создать по протоколу на канал. Возвращает список путей."""
    os.makedirs(out_dir, exist_ok=True)
    paths = []
    for ch in channels:
        doc = from_template(template, ch, digits) if template else builtin(ch, digits)
        base = safe_name(ch.kks or f"канал_{ch.num}")
        path = os.path.join(out_dir, base + ".docx")
        n = 2
        while os.path.exists(path) and path in paths:
            path = os.path.join(out_dir, f"{base}_{n}.docx")
            n += 1
        doc.save(path)
        paths.append(path)
    return paths
