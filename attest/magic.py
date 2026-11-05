"""Magic: заполнить пустые «Измеренное значение выходного сигнала» (J..M) в Excel.

Формулы листа аттестации (одинаковые во всех файлах заказчика):
  I = (E-D)*%+D                       действительное значение точки
  N = |СРЗНАЧ(J:M) - I| / (E-D) * 100  приведённая погрешность точки, %
  O = МАКС(N по точкам)                погрешность ЭТ ИК
  Q = КОРЕНЬ(O² + P²)                  погрешность ИК (P - погрешность ПИП)
  R = 1,1 * КОРЕНЬ(P² + W²)            допустимая (W - погрешность модуля, обычно 0,15)
  S = «Годен», если Q <= R

Канал годен, пока O <= КОРЕНЬ(R² - P²). Заданная доля каналов получает
случайную ненулевую погрешность ниже этой границы, остальные - точные значения.
"""
import math
import os
import random

from . import layout
from .data import is_kks, to_float

_rng = random.SystemRandom()  # каждый запуск - новый случайный набор


def _cached(ws_vals, col, r):
    if not col:
        return None
    v = ws_vals["%s%d" % (col, r)].value
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else to_float(v)


def read_channels(ws_vals, lay):
    """Каналы листа с номерами строк Excel: [{'row', 'kks', 'low', 'high', 'P', 'R', 'points': [...]}]."""
    g = lambda key, r: _cached(ws_vals, lay.get(key), r)
    hidden = lambda r: bool(ws_vals.row_dimensions[r].hidden) if r in ws_vals.row_dimensions else False
    channels, cur = [], None
    for r in range(4, ws_vals.max_row + 1):
        kks = ws_vals["%s%d" % (lay["kks"], r)].value
        num = ws_vals["%s%d" % (lay["num"], r)].value if lay.get("num") else None
        has_num = num is not None and str(num).strip() != ""
        if cur is not None and not has_num and (kks is None or not str(kks).strip()):
            cur["hidden"] = cur["hidden"] and hidden(r)  # скрыт, только если скрыты все строки
        if is_kks(kks, ws_vals["%s%d" % (lay["num"], r)].value if lay.get("num") else None,
                  ws_vals["%s%d" % (lay["name"], r)].value if lay.get("name") else None):
            low, high = g("low", r), g("high", r)
            if "low" not in lay and lay.get("range"):  # диапазон текстом «0…1,02»
                low, high = layout.parse_range(ws_vals["%s%d" % (lay["range"], r)].value or "")
            p = g("pip_ik", r)
            if p is None:
                p = g("pip_err", r)
            w = g("module_err", r)
            allowed = g("allowed", r)
            if allowed is None:
                allowed = g("allowed_ik", r)
            if allowed is None and p is not None:
                allowed = 1.1 * math.sqrt(p * p + (w if w is not None else 0.15) ** 2)
            cur = {"row": r, "kks": str(kks).strip(), "low": low or 0.0, "high": high,
                   "P": p, "R": allowed, "points": [], "hidden": hidden(r)}
            channels.append(cur)
        elif has_num and cur is not None and cur["row"] != r:
            cur = None  # № без KKS - заготовка; её строки не относятся к предыдущему каналу (как в data)
        if cur is None or cur["high"] is None:
            continue
        pct = g("pct", r)
        xd = g("output", r)
        if xd is None and pct is not None:  # формула I = (E-D)*%+D без сохранённого значения
            xd = (cur["high"] - cur["low"]) * pct / 100.0 + cur["low"]
        if pct is None and xd is None:
            continue
        cells = {c: ws_vals["%s%d" % (c, r)].value for c in lay["m"]}
        blank = [c for c, v in cells.items() if v is None or str(v).strip() == ""]
        nums = [to_float(v) for c, v in cells.items() if c not in blank]
        cur["points"].append({"row": r, "xd": xd, "empty": len(blank) == len(cells),
                              "blank": blank, "nums": [v for v in nums if v is not None]})
    # скрытые в Excel каналы не трогаем
    return [c for c in channels if not c["hidden"]]


def step_for(span):
    """Шаг округления измерений: примерно 0,05 % диапазона, степень десяти."""
    if span <= 0:
        return 1.0
    return 10.0 ** math.floor(math.log10(span * 0.0005))


def _round(v, step):
    digits = max(0, -int(math.floor(math.log10(step))))
    return round(round(v / step) * step, digits)


def max_error(ch):
    """Наибольшая ЭТ ИК, при которой канал ещё годен (с запасом 5 %)."""
    p, r = ch["P"], ch["R"]
    if r is None:
        return 0.1
    p = p or 0.0
    if r <= p:
        return 0.0
    return math.sqrt(r * r - p * p) * 0.95


def generate_channel(ch, nonzero, rng=_rng):
    """{строка Excel: измеренное значение} для точек канала."""
    span = abs(ch["high"] - ch["low"])  # бывает и обратный диапазон (1,6…0)
    step = step_for(span)
    out = {p["row"]: p["xd"] for p in ch["points"]}
    limit = max_error(ch)
    # первая точка (0 %) всегда точная; погрешность только у остальных
    pts = sorted(ch["points"], key=lambda p: p["row"])[1:]
    if not nonzero or limit <= 0 or span <= 0 or not pts:
        return out
    # Доли взяты из заполненных Excel заказчика (~1000 каналов с погрешностью):
    # отклонения в одну сторону; на точках 2..5 нулевых 0/1/2/3 - 34/24/25/17 %;
    # разных величин отклонения 1/2/3/4 - 59/22/14/5 %; наибольшее отклонение
    # обычно 0,1-0,4 от границы годности.
    sign = rng.choice((-1, 1))
    frac = _pick(rng, ((0.05, 0.15, 30), (0.15, 0.25, 25), (0.25, 0.35, 15),
                       (0.35, 0.45, 13), (0.45, 0.95, 17)))
    top = _round(frac * limit / 100.0 * span, step)
    while top / span * 100.0 > limit and top >= step:  # после округления - не за границу
        top = _round(top - step, step)
    if top < step / 2:  # шаг крупный - минимальное заметное отклонение
        top = step
        if step / span * 100.0 > limit:
            return out
    n_zero = min(len(pts) - 1, _choice_w(rng, (0, 1, 2, 3), (34, 24, 25, 17)))
    active = rng.sample(pts, len(pts) - n_zero)
    levels = [top]
    for _ in range(_choice_w(rng, (1, 2, 3, 4), (59, 22, 14, 5)) - 1):
        lvl = _round(rng.uniform(0.3, 1.0) * top, step)
        if step / 2 <= lvl < top and lvl not in levels:
            levels.append(lvl)
    levels = levels[:len(active)]  # наибольшее отклонение (top) всегда остаётся
    mags = levels + [rng.choice(levels) for _ in range(len(active) - len(levels))]
    rng.shuffle(mags)
    for p, m in zip(active, mags):
        out[p["row"]] = _round(p["xd"] + sign * m, step)
    return out


def _choice_w(rng, items, weights):
    x = rng.uniform(0, sum(weights))
    for item, w in zip(items, weights):
        x -= w
        if x <= 0:
            return item
    return items[-1]


def _pick(rng, bands):
    lo, hi, _w = _choice_w(rng, bands, [b[2] for b in bands])
    return rng.uniform(lo, hi)


def point_error(ch, values):
    """ЭТ ИК канала по сгенерированным значениям (формулы N, O листа)."""
    span = abs(ch["high"] - ch["low"])
    if span == 0:
        return 0.0
    return max(abs(values[p["row"]] - p["xd"]) / span * 100.0 for p in ch["points"])


def fill_partial(ch):
    """Канал, где часть измерений уже есть: дописать только пустые ячейки.

    Введённые значения не меняются. Пустая точка целиком получает точное
    действительное значение, недописанная строка (J есть, K..M пусто) -
    среднее уже введённых, чтобы погрешность точки (N) не изменилась.
    Возвращает {строка Excel: {столбец: значение}}.
    """
    out = {}
    for p in ch["points"]:
        if not p["blank"]:
            continue
        if p["empty"]:
            if p["xd"] is None:
                continue
            v = p["xd"]
        elif p["nums"]:
            v = sum(p["nums"]) / len(p["nums"])
        else:
            continue  # в строке только текст/ошибки - это покажет окно «Ошибки»
        out[p["row"]] = {c: round(v, 10) for c in p["blank"]}
    return out


def _out_path(path, out_dir, suffix):
    src = os.path.join(out_dir, os.path.basename(path)) if out_dir else path
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(src)[0]
    out, n = "%s (%s).xlsx" % (base, suffix), 2
    while os.path.exists(out):
        out, n = "%s (%s %d).xlsx" % (base, suffix, n), n + 1
    return out


def make_good(path, head_rows, percent, rng=_rng, out_dir=None):
    """Негодные каналы (строки Excel с KKS) - заново заполнить J..M по правилам magic.

    Измерения этих каналов заменяются целиком, остальные каналы не трогаются.
    Возвращает (путь копии, {строка: значение}, [строки каналов, которые годными
    сделать нельзя: допуск R не больше погрешности ПИП P]).
    """
    import openpyxl
    wb = openpyxl.load_workbook(path)
    vals = openpyxl.load_workbook(path, data_only=True)
    ws_vals, lay = layout.pick_sheet(vals)
    ws = wb[ws_vals.title]
    heads = set(head_rows)
    channels = [c for c in read_channels(ws_vals, lay) if c["row"] in heads]
    impossible = [c["row"] for c in channels if max_error(c) <= 0]
    channels = [c for c in channels if max_error(c) > 0 and all(p["xd"] is not None for p in c["points"])]
    k = int(round(len(channels) * percent / 100.0))
    chosen = set(id(c) for c in rng.sample(channels, min(k, len(channels))))
    written = {}
    for c in channels:
        for row, v in generate_channel(c, id(c) in chosen, rng).items():
            for col in lay["m"]:
                ws["%s%d" % (col, row)] = v
            written[row] = v
    wb.calculation.fullCalcOnLoad = True
    out = _out_path(path, out_dir, "годные")
    wb.save(out)
    return out, written, impossible


def run(path, percent, out_path=None, rng=_rng, out_dir=None):
    """Заполнить пустые J..M. Возвращает (путь нового файла, {строка: значение}, статистика).

    Каналы без единого измерения заполняются целиком (с долей погрешности),
    в частично заполненных дописываются только пустые ячейки (fill_partial).
    В {строка: значение} для таких строк - {столбец: значение}.
    """
    import openpyxl
    wb = openpyxl.load_workbook(path)
    vals = openpyxl.load_workbook(path, data_only=True)
    ws_vals, lay = layout.pick_sheet(vals)
    ws = wb[ws_vals.title]
    all_channels = read_channels(ws_vals, lay)
    channels = [c for c in all_channels
                if c["points"] and all(p["empty"] for p in c["points"])
                and all(p["xd"] is not None for p in c["points"])]
    partial = [c for c in all_channels if c["points"] and not all(p["empty"] for p in c["points"])]
    # погрешность можно дать только каналам с диапазоном и запасом до допуска
    eligible = [c for c in channels if c["high"] != c["low"] and max_error(c) > 0]
    k = min(len(eligible), int(round(len(channels) * percent / 100.0)))
    chosen = set(id(c) for c in rng.sample(eligible, k))
    written = {}
    nonzero = 0
    for c in channels:
        values = generate_channel(c, id(c) in chosen, rng)
        if point_error(c, values) > 0:
            nonzero += 1
        for row, v in values.items():
            for col in lay["m"]:
                ws["%s%d" % (col, row)] = v
            written[row] = v
    partial_channels = partial_cells = 0
    for c in partial:
        cells = fill_partial(c)
        if cells:
            partial_channels += 1
        for row, cols in cells.items():
            for col, v in cols.items():
                ws["%s%d" % (col, row)] = v
                partial_cells += 1
            written[row] = cols
    wb.calculation.fullCalcOnLoad = True  # Excel пересчитает формулы N, O, Q, S при открытии
    if out_path is None:
        src = os.path.join(out_dir, os.path.basename(path)) if out_dir else path
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(src)[0]
        out_path, n = base + " (magic).xlsx", 2
        while os.path.exists(out_path):
            out_path, n = "%s (magic %d).xlsx" % (base, n), n + 1
    wb.save(out_path)
    return out_path, written, {"sheet": ws.title, "channels": len(channels), "nonzero": nonzero,
                               "partial_channels": partial_channels, "partial_cells": partial_cells}
