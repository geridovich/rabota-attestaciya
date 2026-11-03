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

from .data import to_float

MEASURE_COLS = ("J", "K", "L", "M")
HEADER_MARK = "Измеренное значение"
_rng = random.SystemRandom()  # каждый запуск - новый случайный набор


def pick_sheet(wb):
    """Лист с измерениями: первый, где в шапке есть «Измеренное значение…»."""
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=2, values_only=True):
            if any(isinstance(v, str) and HEADER_MARK in v for v in row):
                return ws
    return wb.worksheets[0]


def _cached(ws_vals, cell):
    v = ws_vals[cell].value
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else to_float(v)


def _formula_value(ws_vals, r, col, low, high, pct):
    """Значение ячейки: сохранённое Excel, иначе посчитанное по формуле листа."""
    v = _cached(ws_vals, "%s%d" % (col, r))
    if v is not None:
        return v
    if col == "I" and pct is not None:
        return (high - low) * pct / 100.0 + low
    return None


def read_channels(ws, ws_vals):
    """Каналы листа с номерами строк Excel: [{'row', 'kks', 'low', 'high', 'P', 'R', 'points': [...]}]."""
    channels, cur = [], None
    for r in range(4, ws.max_row + 1):
        kks = ws_vals["B%d" % r].value
        if kks is not None and str(kks).strip():
            low = _cached(ws_vals, "D%d" % r) or 0.0
            high = _cached(ws_vals, "E%d" % r)
            p = _cached(ws_vals, "P%d" % r)
            if p is None:
                p = _cached(ws_vals, "U%d" % r)
            w = _cached(ws_vals, "W%d" % r)
            allowed = _cached(ws_vals, "R%d" % r)
            if allowed is None:
                allowed = _cached(ws_vals, "X%d" % r)
            if allowed is None and p is not None:
                allowed = 1.1 * math.sqrt(p * p + (w if w is not None else 0.15) ** 2)
            cur = {"row": r, "kks": str(kks).strip(), "low": low, "high": high,
                   "P": p, "R": allowed, "points": []}
            channels.append(cur)
        if cur is None or cur["high"] is None:
            continue
        pct = _cached(ws_vals, "G%d" % r)
        xd = _formula_value(ws_vals, r, "I", cur["low"], cur["high"], pct)
        if pct is None and xd is None:
            continue
        filled = [ws_vals["%s%d" % (c, r)].value for c in MEASURE_COLS]
        cur["points"].append({"row": r, "xd": xd,
                              "empty": all(v is None or str(v).strip() == "" for v in filled)})
    return channels


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
    span = ch["high"] - ch["low"]
    step = step_for(span)
    out = {p["row"]: p["xd"] for p in ch["points"]}
    limit = max_error(ch)
    if not nonzero or limit <= 0 or span <= 0:
        return out
    target = rng.uniform(0.15, 0.85) * limit
    pts = ch["points"]
    loud = rng.randrange(len(pts))  # хотя бы одна точка с заметной погрешностью
    for i, p in enumerate(pts):
        err = rng.uniform(0.5, 1.0) * target if i == loud else rng.uniform(0.0, 1.0) * target
        if i != loud and rng.random() < 0.35:
            err = 0.0
        val = _round(p["xd"] + rng.choice((-1, 1)) * err / 100.0 * span, step)
        # после округления погрешность не должна выйти за границу годности
        while abs(val - p["xd"]) / span * 100.0 > limit and abs(val - p["xd"]) >= step:
            val = _round(val - math.copysign(step, val - p["xd"]), step)
        out[p["row"]] = val
    if all(abs(out[p["row"]] - p["xd"]) < step / 2 for p in pts):
        p = pts[loud]
        val = _round(p["xd"] + rng.choice((-1, 1)) * step, step)
        if abs(val - p["xd"]) / span * 100.0 <= limit:
            out[p["row"]] = val
    return out


def point_error(ch, values):
    """ЭТ ИК канала по сгенерированным значениям (формулы N, O листа)."""
    span = ch["high"] - ch["low"]
    return max(abs(values[p["row"]] - p["xd"]) / span * 100.0 for p in ch["points"])


def run(path, percent, out_path=None, rng=_rng):
    """Заполнить пустые J..M. Возвращает (путь нового файла, {строка: значение}, статистика)."""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    vals = openpyxl.load_workbook(path, data_only=True)
    ws = pick_sheet(wb)
    ws_vals = vals[ws.title]
    channels = [c for c in read_channels(ws, ws_vals)
                if c["points"] and all(p["empty"] for p in c["points"])
                and all(p["xd"] is not None for p in c["points"])]
    k = int(round(len(channels) * percent / 100.0))
    chosen = set(id(c) for c in rng.sample(channels, k)) if channels else set()
    written = {}
    nonzero = 0
    for c in channels:
        values = generate_channel(c, id(c) in chosen, rng)
        if point_error(c, values) > 0:
            nonzero += 1
        for row, v in values.items():
            for col in MEASURE_COLS:
                ws["%s%d" % (col, row)] = v
            written[row] = v
    wb.calculation.fullCalcOnLoad = True  # Excel пересчитает формулы N, O, Q, S при открытии
    if out_path is None:
        base, ext = os.path.splitext(path)
        out_path, n = base + " (magic).xlsx", 2
        while os.path.exists(out_path):
            out_path, n = "%s (magic %d).xlsx" % (base, n), n + 1
    wb.save(out_path)
    return out_path, written, {"sheet": ws.title, "channels": len(channels), "nonzero": nonzero}
