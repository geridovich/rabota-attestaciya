"""Сделать шаблоны Word из образцов заказчика (Протокол.docx, Перечень.docx).

Из образцов берётся оформление: один протокол и по одной строке-образцу в
таблицах. Конкретные значения заменяются метками {{...}}, которые потом
заполняет attest/docgen.py. Запуск:

    python tools/make_templates.py Протокол.docx Перечень.docx
"""
import copy
import os
import sys

from docx import Document

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
OUT = os.path.join(os.path.dirname(__file__), "..", "attest", "templates")


def text(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def set_text(el, value, like=None):
    """Текст в первый фрагмент первого абзаца, остальные фрагменты очистить.

    В пустой ячейке фрагмент создаётся с оформлением шрифта из ячейки like.
    """
    ts = list(el.iter(W + "t"))
    if not ts:
        p = el if el.tag == W + "p" else el.find(".//" + W + "p")
        r = p.makeelement(W + "r", {})
        src = like.find(".//" + W + "r/" + W + "rPr") if like is not None else None
        if src is not None:
            r.append(copy.deepcopy(src))
        t = r.makeelement(W + "t", {})
        r.append(t)
        p.append(r)
        ts = [t]
    ts[0].text = value
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for t in ts[1:]:
        t.text = ""


def cells(tr):
    return tr.findall(W + "tc")


def rows(tbl):
    return tbl.findall(W + "tr")


def remove(el):
    el.getparent().remove(el)


def protocol(src):
    doc = Document(src)
    body = doc.element.body
    kids = [e for e in body if e.tag != W + "sectPr"]
    starts = [i for i, e in enumerate(kids) if e.tag == W + "p" and text(e).startswith("Протокол №")]
    for e in kids[:starts[0]] + kids[starts[1]:]:
        remove(e)
    kids = kids[starts[0]:starts[1]]
    paras = [e for e in kids if e.tag == W + "p"]
    tbls = [e for e in kids if e.tag == W + "tbl"]
    set_text(kids[0], "Протокол № {{№}} от {{ДАТА}}")
    title = next(e for e in paras if text(e).startswith("Результатов"))
    set_text(title, "Результатов экспериментальных исследований при проведении "
                       "метрологической аттестации ИК {{СТАНЦИЯ}}")

    info = rows(tbls[0])
    set_text(cells(info[0])[-1], "{{НАИМЕНОВАНИЕ}}")
    set_text(cells(info[1])[-1], "{{KKS}}")
    set_text(cells(info[2])[-1], "{{ПИП}}")
    set_text(cells(info[3])[-1], "{{ВТОРИЧНЫЙ}}", like=cells(info[2])[-1])
    set_text(cells(info[4])[-1], "{{МОДУЛЬ}}")
    set_text(cells(info[5])[-1], "{{ДИАПАЗОН}}")

    # таблица измерений: шапка 3 строки, затем пары Меньшее/Большее по точкам
    meas = rows(tbls[2])
    for tr in meas[5:]:
        remove(tr)
    low, high = meas[3], meas[4]
    lc = cells(low)
    set_text(lc[0], "{{т.%}}")
    set_text(lc[1], "{{т.вход}}")
    set_text(lc[2], "{{т.xd}}")
    for k, c in enumerate(lc[4:]):
        set_text(c, "{{т.м%d}}" % (k + 1))
    for k, c in enumerate(cells(high)[4:]):
        set_text(c, "{{т.б%d}}" % (k + 1))

    for c, key in zip(cells(rows(tbls[3])[1]), ("a1", "a2", "b1", "b2", "gh", "gl")):
        set_text(c, "{{%s}}" % key)

    sign = cells(rows(tbls[4])[0])
    set_text(sign[1], "{{ДОЛЖНОСТЬ}}")
    set_text(sign[3], "__{{ФИО}}_")
    return doc


def perechen(src):
    doc = Document(src)
    body = doc.element.body
    kids = [e for e in body if e.tag != W + "sectPr"]
    tbls = [e for e in kids if e.tag == W + "tbl"]
    title = next(e for e in kids if e.tag == W + "p" and text(e).startswith("Перечень"))
    set_text(title, "Перечень измерительных каналов {{ПЕРЕЧЕНЬ}}")

    # перечень: шапка 2 строки (повторяется на каждой странице) + пара строк на канал
    main = tbls[0]
    tr = rows(main)
    for r in tr[:2]:
        trpr = r.find(W + "trPr")
        if trpr is None:
            trpr = r.makeelement(W + "trPr", {})
            r.insert(0, trpr)
        if trpr.find(W + "tblHeader") is None:
            trpr.append(trpr.makeelement(W + "tblHeader", {}))
    for r in tr[4:]:
        remove(r)
    a, b = cells(tr[2]), cells(tr[3])
    for c, key in zip(a, ("п.№", "п.KKS", "п.наим", "п.диап", "п.пип", "п.пип_погр", "п.ик")):
        set_text(c, "{{%s}}" % key)
    set_text(b[4], "{{п.модуль}}")
    set_text(b[5], "{{п.модуль_погр}}")
    # продолжение перечня было отдельной таблицей после разрыва страницы - убрать
    i = kids.index(tbls[1])
    for e in kids[kids.index(main) + 1:i + 1]:
        remove(e)

    # градуировочная таблица: подпись + шапка 2 строки + первая строка диапазона
    # + строка продолжения. Таблиц столько, сколько видов сигнала (мА, Ом, мВ),
    # их размножает docgen, поэтому в шаблоне остаётся одна, вторая удаляется.
    tbl = tbls[2]
    tr = rows(tbl)
    for r in tr[4:]:
        remove(r)
    first, nxt = cells(tr[2]), cells(tr[3])
    for c, key in zip(first, ("диап", "ед", "%", "знач", "вх")):
        set_text(c, "{{г.%s}}" % key, like=first[2])
    for c, key in zip(nxt[2:], ("%", "знач", "вх")):
        set_text(c, "{{г2.%s}}" % key, like=nxt[2])
    hdr = cells(tr[1])  # первые две ячейки - продолжение объединения
    set_text(hdr[3], "в ед. изм. вел.")
    set_text(hdr[4], "в ед. вх. воздействия, {{ВХ}}")
    cap1 = next(e for e in kids if e.tag == W + "p" and text(e).startswith("Таблица 1"))
    set_text(cap1, "{{ТАБЛИЦА}}")
    cap2 = next(e for e in kids if e.tag == W + "p" and text(e).startswith("Таблица 2"))
    for e in kids[kids.index(tbl) + 1:]:
        if e is not cap2 and e is not tbls[3] and not (e.tag == W + "p" and not text(e).strip()):
            continue
        if kids.index(e) >= kids.index(cap2) - 3:
            remove(e)
    return doc


def main(proto_src, perechen_src):
    os.makedirs(OUT, exist_ok=True)
    protocol(proto_src).save(os.path.join(OUT, "protocol.docx"))
    perechen(perechen_src).save(os.path.join(OUT, "perechen.docx"))
    print("готово:", os.path.abspath(OUT))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
