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
    channels, cur = [], None
    for r in range(4, ws_vals.max_row + 1):
        kks = ws_vals["%s%d" % (lay["kks"], r)].value
        if is_kks(kks):
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
                   "P": p, "R": allowed, "points": []}
            channels.append(cur)
        if cur is None or cur["high"] is None:
            continue
        pct = g("pct", r)
        xd = g("output", r)
        if xd is None and pct is not None:  # формула I = (E-D)*%+D без сохранённого значения
            xd = (cur["high"] - cur["low"]) * pct / 100.0 + cur["low"]
        if pct is None and xd is None:
            continue
        filled = [ws_vals["%s%d" % (c, r)].value for c in lay["m"]]
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
    span = abs(ch["high"] - ch["low"])  # бывает и обратный диапазон (1,6…0)
    step = step_for(span)
    out = {p["row"]: p["xd"] for p in ch["points"]}
    limit = max_error(ch)
    # первая точка (0 %) всегда точная; погрешность только у остальных
    pts = sorted(ch["points"], key=lambda p: p["row"])[1:]
    if not nonzero or limit <= 0 or span <= 0 or not pts:
        return out
    # одна и та же погрешность на всех точках канала, в одну сторону
    err = rng.uniform(0.15, 0.85) * limit
    delta = _round(rng.choice((-1, 1)) * err / 100.0 * span, step)
    # после округления погрешность не должна выйти за границу годности
    while abs(delta) / span * 100.0 > limit and abs(delta) >= step:
        delta = _round(delta - math.copysign(step, delta), step)
    if abs(delta) < step / 2:  # шаг крупный - минимальное заметное отклонение
        delta = math.copysign(step, delta or 1)
        if step / span * 100.0 > limit:
            return out
    for p in pts:
        out[p["row"]] = _round(p["xd"] + delta, step)
    return out


def point_error(ch, values):
    """ЭТ ИК канала по сгенерированным значениям (формулы N, O листа)."""
    span = abs(ch["high"] - ch["low"])
    if span == 0:
        return 0.0
    return max(abs(values[p["row"]] - p["xd"]) / span * 100.0 for p in ch["points"])


def run(path, percent, out_path=None, rng=_rng):
    """Заполнить пустые J..M. Возвращает (путь нового файла, {строка: значение}, статистика)."""
    import openpyxl
    wb = openpyxl.load_workbook(path)
    vals = openpyxl.load_workbook(path, data_only=True)
    ws_vals, lay = layout.pick_sheet(vals)
    ws = wb[ws_vals.title]
    channels = [c for c in read_channels(ws_vals, lay)
                if c["points"] and all(p["empty"] for p in c["points"])
                and all(p["xd"] is not None for p in c["points"])]
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
    wb.calculation.fullCalcOnLoad = True  # Excel пересчитает формулы N, O, Q, S при открытии
    if out_path is None:
        base, ext = os.path.splitext(path)
        out_path, n = base + " (magic).xlsx", 2
        while os.path.exists(out_path):
            out_path, n = "%s (magic %d).xlsx" % (base, n), n + 1
    wb.save(out_path)
    return out_path, written, {"sheet": ws.title, "channels": len(channels), "nonzero": nonzero}
