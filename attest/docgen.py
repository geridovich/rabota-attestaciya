"""Word-документы по шаблонам заказчика: протоколы и перечень с градуировочными таблицами.

Шаблоны attest/templates/*.docx сделаны из образцов заказчика
(tools/make_templates.py). В них метки {{...}}; строки таблиц с метками
точки или канала размножаются по количеству точек или каналов.
"""
import copy
import os
import re
import sys

from docx import Document

from .calcd import fmt
from .data import expand

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MARK_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
OBSERVATIONS = 20  # столбцов «Число наблюдений» в протоколе

EXECUTORS = (
    ("Руководитель группы", "Жулдыбин Д.А."),
    ("Ведущий инженер", "Калекин В.В."),
    ("Ведущий инженер", "Детчуев А.А."),
    ("Инженер I кат.", "Геридович Н.В."),
    ("Инженер", "Лобанов Н.В."),
)

SIGNALS = {
    "мА": "градуировочная таблица ИК сигналов в диапазоне 4-20 мА",
    "Ом": "градуировочная таблица ИК сигналов термосопротивления",
}


def template_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    if getattr(sys, "frozen", False):
        base = os.path.join(base, "attest")
    return os.path.join(base, "templates", name)


def num(v):
    """Число из Excel без хвостов вида 0,44999999999999996, с запятой."""
    if v is None:
        return ""
    return fmt(float("%.10g" % v))


def err_num(v):
    """Погрешность в перечне: не больше 3 знаков после запятой (0,574216858 -> 0,574)."""
    if v is None:
        return ""
    return fmt(float("%.3f" % v))


def rng_text(ch):
    return "%s…%s" % (num(ch.low), num(ch.high))


# ---------- работа с XML ----------
def _xml_text(el):
    return "".join(t.text or "" for t in el.iter(W + "t"))


def fill(el, values):
    """Подставить метки во все абзацы элемента."""
    for p in el.iter(W + "p"):
        ts = list(p.iter(W + "t"))
        text = "".join(t.text or "" for t in ts)
        if "{{" not in text:
            continue
        new = MARK_RE.sub(lambda m: values.get(m.group(1), m.group(0)), text)
        if new == text:
            continue
        # метка может быть разбита Word'ом на несколько фрагментов: текст в первый
        ts[0].text = new
        ts[0].set(XML_SPACE, "preserve")
        for t in ts[1:]:
            t.text = ""


def _rows_with(tbl, prefix):
    return [tr for tr in tbl.iter(W + "tr") if "{{" + prefix in _xml_text(tr)]


def _expand_group(tbl, prefix, items, fields):
    """Строки с метками prefix - образец группы; повторить для каждого элемента."""
    group = _rows_with(tbl, prefix)
    if not group:
        return
    for item in items:
        for tr in group:
            new = copy.deepcopy(tr)
            fill(new, fields(item))
            group[0].addprevious(new)
    for tr in group:
        tr.getparent().remove(tr)


def _page_break_before(el):
    if el.tag != W + "p":
        return
    ppr = el.find(W + "pPr")
    if ppr is None:
        ppr = el.makeelement(W + "pPr", {})
        el.insert(0, ppr)
    if ppr.find(W + "pageBreakBefore") is None:
        ppr.insert(0, ppr.makeelement(W + "pageBreakBefore", {}))


def _body_parts(doc):
    body = doc.element.body
    sect = body.sectPr
    proto = [el for el in body.iterchildren() if el is not sect]
    for el in proto:
        body.remove(el)
    return body, sect, proto


def _append(body, sect, el):
    if sect is not None:
        sect.addprevious(el)
    else:
        body.append(el)


# ---------- протоколы ----------
def date_for(ch, ranges, default=""):
    """ranges: [(с №, по №, дата)]; дата по номеру протокола (№ канала из Excel)."""
    try:
        n = int(float(ch.num))
    except (TypeError, ValueError):
        return default
    for lo, hi, date in ranges:
        if lo is not None and hi is not None and lo <= n <= hi and date:
            return date
    return default


def observations(p):
    """Меньшее / Большее по 20 наблюдений: первая и вторая половина измерений."""
    if not p.measured:
        return [""] * OBSERVATIONS, [""] * OBSERVATIONS
    series = expand(p.measured, OBSERVATIONS * 2)
    return series[:OBSERVATIONS], series[OBSERVATIONS:]


def protocol_fields(ch, ranges, station, executor):
    r = ch.result
    f = lambda v: fmt(v) if r else ""
    return {
        "№": ch.num, "ДАТА": date_for(ch, ranges, "__.__.____"), "СТАНЦИЯ": station,
        "НАИМЕНОВАНИЕ": ch.name, "KKS": ch.kks, "ПИП": ch.pip, "ВТОРИЧНЫЙ": "",
        "МОДУЛЬ": ch.module, "ДИАПАЗОН": "%s %s" % (rng_text(ch), ch.unit),
        "a1": f(r.a1) if r else "", "a2": f(r.a2) if r else "", "b1": f(r.b1) if r else "",
        "b2": f(r.b2) if r else "", "gh": f(r.gh) if r else "", "gl": f(r.gl) if r else "",
        "ДОЛЖНОСТЬ": executor[0], "ФИО": executor[1],
    }


def point_fields(p):
    low, high = observations(p)
    vals = {"т.%": num(p.pct), "т.вход": num(p.input), "т.xd": num(p.xd)}
    for k in range(OBSERVATIONS):
        vals["т.м%d" % (k + 1)] = num(low[k]) if low[k] != "" else ""
        vals["т.б%d" % (k + 1)] = num(high[k]) if high[k] != "" else ""
    return vals


def _number_from_one(sect):
    """Нумерация страниц с 1 (в образце заказчика стояло «начать с 29»)."""
    if sect is None:
        return
    pg = sect.find(W + "pgNumType")
    if pg is None:
        pg = sect.makeelement(W + "pgNumType", {})
        sect.append(pg)
    pg.set(W + "start", "1")


def build_protocols(channels, ranges, station, executor, template=None):
    doc = Document(template or template_path("protocol.docx"))
    body, sect, proto = _body_parts(doc)
    _number_from_one(sect)
    for n, ch in enumerate(channels):
        els = [copy.deepcopy(el) for el in proto]
        for el in els:
            for tbl in [el] if el.tag == W + "tbl" else []:
                _expand_group(tbl, "т.", ch.points, point_fields)
            fill(el, protocol_fields(ch, ranges, station, executor))
        if n and els:
            _page_break_before(els[0])
        for el in els:
            _append(body, sect, el)
    return doc


# ---------- перечень и градуировочные таблицы ----------
THERMOCOUPLE_RE = re.compile(r"ТХА|ТХК|ТПП|ТПР|ТП-|ТП\s|КТХА|\(K\)|\(L\)|термопар", re.I)


def is_temperature(ch):
    u = ch.unit.replace("˚", "°").upper()
    return u.endswith("С") or u.endswith("C")


def signal_kind(ch):
    """Вид входного воздействия для градуировочной таблицы.

    Температура: термосопротивление - «Ом», термопара - None (в образце
    заказчика для термопар таблиц нет). Остальное - токовый сигнал «мА».
    """
    if is_temperature(ch):
        return None if THERMOCOUPLE_RE.search(ch.pip or "") else "Ом"
    return "мА"


# НСХ термосопротивлений: медь как в образце заказчика (50М, alpha 0,00426),
# платина по ГОСТ 6651-2009 (100П alpha 0,00391, Pt100 alpha 0,00385)
PLATINUM = {"П": (3.9690e-3, -5.841e-7, -4.330e-12), "Pt": (3.9083e-3, -5.775e-7, -4.183e-12)}


def rtd_type(pip):
    s = (pip or "").replace(" ", "")
    m = re.search(r"(50|100|500|1000)(М|M|П|P)", s)
    if re.search(r"Pt\s*(100|500|1000)", pip or "", re.I):
        r0 = float(re.search(r"Pt\s*(\d+)", pip, re.I).group(1))
        return r0, "Pt"
    if m:
        return float(m.group(1)), "П" if m.group(2) in "ПP" else "М"
    if re.search(r"ТСП", s):
        return 100.0, "П"
    return 50.0, "М"


def resistance(pip, t):
    r0, kind = rtd_type(pip)
    if kind == "М":
        return r0 * (1 + 4.26e-3 * t)
    a, b, c = PLATINUM[kind]
    r = 1 + a * t + b * t * t
    if t < 0:
        r += c * (t - 100) * t ** 3
    return r0 * r


def calibration_input(ch, kind, p):
    if kind == "Ом":
        return round(resistance(ch.pip, p.xd), 3)
    # 4-20 мА по % диапазона: в Excel в столбце H бывает и ток, и значение в ед. изм.
    pct = p.pct if p.pct is not None else (p.xd - ch.low) / (ch.high - ch.low) * 100.0
    return round(4 + 16 * pct / 100.0, 6)


def calibration_groups(channels):
    """Уникальные диапазоны по видам сигнала, в порядке появления."""
    groups = {}
    for ch in channels:
        kind = signal_kind(ch)
        if kind is None or not ch.points:
            continue
        key = (rng_text(ch), ch.unit,
               tuple((num(p.pct), num(p.xd), num(calibration_input(ch, kind, p))) for p in ch.points))
        lst = groups.setdefault(kind, [])
        if key not in [k for k, _ in lst]:
            lst.append((key, ch))
    return [(kind, [ch for _, ch in groups[kind]]) for kind in ("мА", "Ом") if kind in groups]


def perechen_fields(ch):
    pip_err = ch.pip_err if ch.pip_err is not None else ch.pip_ik
    return {
        "п.№": ch.num, "п.KKS": ch.kks, "п.наим": ch.name,
        "п.диап": "%s %s" % (rng_text(ch), ch.unit), "п.пип": ch.pip, "п.пип_погр": err_num(pip_err),
        "п.ик": err_num(ch.allowed_value), "п.модуль": ch.module, "п.модуль_погр": err_num(ch.module_err),
    }


def _calibration_table(proto_tbl, kind, chans):
    tbl = copy.deepcopy(proto_tbl)
    first = _rows_with(tbl, "г.")
    nxt = _rows_with(tbl, "г2.")
    anchor = first[0]
    for ch in chans:
        for i, p in enumerate(ch.points):
            src = first[0] if i == 0 else nxt[0]
            new = copy.deepcopy(src)
            vals = {"г.диап": rng_text(ch), "г.ед": ch.unit}
            pre = "г." if i == 0 else "г2."
            vals.update({pre + "%": num(p.pct), pre + "знач": num(p.xd),
                         pre + "вх": num(calibration_input(ch, kind, p))})
            fill(new, vals)
            anchor.addprevious(new)
    for tr in first + nxt:
        tr.getparent().remove(tr)
    return tbl


def build_perechen(channels, title, template=None):
    doc = Document(template or template_path("perechen.docx"))
    body = doc.element.body
    fill(body, {"ПЕРЕЧЕНЬ": title})
    tables = list(body.iter(W + "tbl"))
    _expand_group(tables[0], "п.", channels, perechen_fields)

    cap = next(p for p in body.iter(W + "p") if "{{ТАБЛИЦА}}" in _xml_text(p))
    proto_tbl = tables[1]
    anchor = cap
    for n, (kind, chans) in enumerate(calibration_groups(channels)):
        c = copy.deepcopy(cap)
        fill(c, {"ТАБЛИЦА": "Таблица %d – %s" % (n + 1, SIGNALS[kind])})
        if n:
            _page_break_before(c)
        t = _calibration_table(proto_tbl, kind, chans)
        fill(t, {"ВХ": kind})
        anchor.addprevious(c)
        anchor.addprevious(t)
        spacer = cap.makeelement(W + "p", {})
        anchor.addprevious(spacer)
    cap.getparent().remove(cap)
    proto_tbl.getparent().remove(proto_tbl)
    return doc


# ---------- сохранение ----------
def safe_name(s):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", s).strip(" .")
    return s or "документ"


def unique_path(out_dir, name):
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.join(out_dir, safe_name(name))
    path, n = base + ".docx", 2
    while os.path.exists(path):
        path, n = "%s (%d).docx" % (base, n), n + 1
    return path


def save_protocols(channels, out_dir, ranges, station, executor, template=None):
    first, last = channels[0], channels[-1]
    name = "Протоколы №%s" % first.num + ("-%s" % last.num if last is not first else "")
    path = unique_path(out_dir, name)
    build_protocols(channels, ranges, station, executor, template).save(path)
    return path


def save_perechen(channels, out_dir, title, template=None):
    path = unique_path(out_dir, "Перечень")
    build_perechen(channels, title, template).save(path)
    return path
