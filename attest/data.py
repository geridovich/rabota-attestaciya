"""Разбор данных из Excel (вставка из буфера или файл .xlsx) в список каналов."""
import math
import re
from dataclasses import dataclass, field

from . import calcd

# Столбцы листа аттестации (как в К-8.xlsx)
COLUMNS = {
    "num": "A", "kks": "B", "name": "C", "low": "D", "high": "E", "unit": "F",
    "pct": "G", "input": "H", "output": "I",
    "m1": "J", "m2": "K", "m3": "L", "m4": "M",
    "pip_ik": "P", "allowed": "R", "pip": "T", "pip_err": "U",
    "module": "V", "module_err": "W", "allowed_ik": "X", "notes": "Y",
}


def col_index(letter):
    n = 0
    for ch in letter.upper():
        n = n * 26 + ord(ch) - 64
    return n - 1


def col_letter(idx):
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def parse_columns_spec(spec):
    """'A-Y' или 'B,C,D,H-M' -> список букв в порядке вставки."""
    out = []
    for part in re.split(r"[,\s;]+", spec.strip().upper()):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out += [col_letter(i) for i in range(col_index(a), col_index(b) + 1)]
        else:
            out.append(part)
    return out


def to_float(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(" ", "").replace("−", "-").replace(" ", "").replace(",", ".")
    if s.endswith("%"):
        s = s[:-1]
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def to_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


@dataclass
class Point:
    pct: float
    input: float
    xd: float
    measured: list


@dataclass
class Channel:
    num: str
    kks: str
    name: str
    low: float
    high: float
    unit: str
    pip: str = ""
    pip_err: float = None
    pip_ik: float = None
    module: str = ""
    module_err: float = None
    allowed: float = None
    notes: str = ""
    points: list = field(default_factory=list)
    result: object = None
    error: str = ""
    pip_fallback: float = None      # значения из настроек, если в данных пусто
    allowed_fallback: float = None

    # --- итоговые величины ---
    @property
    def et_ik(self):
        return self.result.et_ik if self.result else None

    @property
    def pip_value(self):
        """Погрешность ПИП для общей погрешности: столбец P, иначе U."""
        for v in (self.pip_ik, self.pip_err, self.pip_fallback):
            if v is not None:
                return v
        return None

    @property
    def allowed_value(self):
        return self.allowed if self.allowed is not None else self.allowed_fallback

    @property
    def total(self):
        if self.et_ik is None:
            return None
        pip = self.pip_value or 0.0
        return math.sqrt(pip ** 2 + self.et_ik ** 2)

    @property
    def verdict(self):
        if self.total is None or self.allowed_value is None:
            return ""
        return "Не годен" if self.total > self.allowed_value else "Годен"

    @property
    def excel_err(self):
        """Средняя приведённая погрешность по точкам, как столбец N/O в Excel."""
        rng = self.high - self.low
        if not rng:
            return None
        errs = [abs((sum(p.measured) / len(p.measured) - p.xd) / rng * 100)
                for p in self.points if p.measured]
        return max(errs) if errs else None


FULL_ROW_MIN = 13  # вставка целых строк листа (A..M и дальше)
SHORT_KEYS = ("num", "kks", "name", "low", "high", "unit", "pct")


def _is_kks(s):
    return bool(re.search(r"[A-Za-z]", s) and re.search(r"\d", s))


def _close(a, b):
    return a is not None and b is not None and abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))


def _has_pct_column(lines):
    """Есть ли во вставке столбец «%» контрольной точки.

    Жёлтые столбцы листа идут без «%» и входного сигнала: №, KKS, Наим.,
    Нижн., Верхн., Ед., действительное, измерения. Если «%» есть, то за ним
    стоит действительное значение (или вход 4..20 мА, а потом действительное),
    и оно совпадает с нижн. + (верхн. - нижн.) * % / 100.
    """
    votes = total = 0
    low = high = None
    for cells in lines:
        if len(cells) > 1 and (cells[0] or cells[1]):
            low, high = to_float(cells[3]) if len(cells) > 3 else None, to_float(cells[4]) if len(cells) > 4 else None
        if low is None or high is None or len(cells) < 8:
            continue
        pct = to_float(cells[6])
        rest = [to_float(c) for c in cells[7:9]]
        if pct is None or rest[0] is None:
            continue
        expected = (high - low) * pct / 100.0 + low
        total += 1
        if any(_close(v, expected) for v in rest):
            votes += 1
    return total > 0 and votes * 2 > total


def rows_from_text(text):
    """Вставка из Excel (табуляция) -> строки {буква столбца: значение}.

    Поддерживаются два вида вставки:
      * целые строки листа (A..Y или с B) - столбцы берутся по своим местам;
      * выделенные (жёлтые) столбцы: №, KKS, Наименование, Нижн., Верхн., Ед.,
        [%, входной сигнал], действительное значение, измерения 1..N.
        Входной сигнал (4..20 мА) определяется и пропускается автоматически.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.split("\t") for ln in text.split("\n")]
    lines = [[c.strip() for c in ln] for ln in lines if any(c.strip() for c in ln)]
    if not lines:
        return []
    shift = 1 if _is_kks(lines[0][0]) else 0  # вставка начата с KKS, без №
    width = max(len(ln) for ln in lines) + shift
    rows = []
    if width >= FULL_ROW_MIN + 6:
        for ln in lines:
            rows.append({col_letter(i + shift): c for i, c in enumerate(ln)})
        return rows

    with_pct = _has_pct_column([[""] * shift + ln for ln in lines])
    keys = SHORT_KEYS if with_pct else SHORT_KEYS[:-1]
    low = high = None
    for ln in lines:
        cells = [""] * shift + ln
        row = {COLUMNS[k]: cells[i] for i, k in enumerate(keys) if i < len(cells)}
        row["_short"] = True
        if row.get(COLUMNS["kks"]) or row.get(COLUMNS["num"]):
            low = to_float(row.get(COLUMNS["low"])) or 0.0
            high = to_float(row.get(COLUMNS["high"])) or 0.0
        rest = [to_float(c) for c in cells[len(keys):]]
        while rest and rest[-1] is None:
            rest.pop()
        pct = to_float(row.get(COLUMNS["pct"]))
        if rest:
            expected = None if pct is None or low is None else (high - low) * pct / 100.0 + low
            if len(rest) >= 2 and not _close(rest[0], expected) and _close(rest[1], expected):
                row[COLUMNS["input"]] = rest.pop(0)
            row[COLUMNS["output"]] = rest.pop(0)
            for i, v in enumerate(rest):
                row[col_letter(col_index(COLUMNS["m1"]) + i)] = v
        if not with_pct and low is not None and high != low and to_float(row.get(COLUMNS["output"])) is not None:
            row[COLUMNS["pct"]] = (to_float(row[COLUMNS["output"]]) - low) / (high - low) * 100.0
        rows.append(row)
    return rows


def rows_from_xlsx(path, first_row=4):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    rows = []
    for r in ws.iter_rows(min_row=first_row, values_only=True):
        if not any(v is not None and str(v).strip() for v in r):
            continue
        rows.append({col_letter(i): v for i, v in enumerate(r)})
    wb.close()
    return rows


def measured_values(row):
    """Измерения: J..M в листе, во вставке жёлтых столбцов - все после действительного."""
    last = col_index(COLUMNS["m4"]) if not row.get("_short") else 200
    out = []
    for i in range(col_index(COLUMNS["m1"]), last + 1):
        v = to_float(row.get(col_letter(i)))
        if v is not None:
            out.append(v)
    return out


def channels_from_rows(rows):
    """Группировка строк в каналы: строка с № или KKS начинает новый канал.

    Строки точек без измеренных значений пропускаются (у температурных
    каналов первая строка канала - пустая точка 0 %).
    """
    channels = []
    cur = None
    g = lambda row, key: row.get(COLUMNS[key])
    for row in rows:
        kks = to_text(g(row, "kks"))
        num = to_text(g(row, "num"))
        if kks or num:
            allowed = to_float(g(row, "allowed"))
            cur = Channel(
                num=num, kks=kks, name=to_text(g(row, "name")),
                low=to_float(g(row, "low")) or 0.0, high=to_float(g(row, "high")) or 0.0,
                unit=to_text(g(row, "unit")),
                pip=to_text(g(row, "pip")), pip_err=to_float(g(row, "pip_err")),
                pip_ik=to_float(g(row, "pip_ik")), module=to_text(g(row, "module")),
                module_err=to_float(g(row, "module_err")),
                allowed=allowed if allowed is not None else to_float(g(row, "allowed_ik")),
                notes=to_text(g(row, "notes")),
            )
            channels.append(cur)
        if cur is None:
            continue
        measured = measured_values(row)
        if not measured:
            continue
        pct = to_float(g(row, "pct"))
        xd = to_float(g(row, "output"))
        if xd is None and pct is not None:  # формула не посчитана - считаем сами
            xd = (cur.high - cur.low) * pct / 100.0 + cur.low
        cur.points.append(Point(pct, to_float(g(row, "input")), xd, measured))
    # пустые заготовки без KKS и без измерений не нужны
    return [c for c in channels if c.kks or c.points]


def expand(measured, qm):
    """Значения точки -> ряд из qm измерений (как вводили в CalcD: «0» 20 раз)."""
    if not measured:
        raise calcd.CalcError("нет измеренных значений")
    if qm % len(measured) == 0:
        rep = qm // len(measured)
        return [m for m in measured for _ in range(rep)]
    return [measured[i % len(measured)] for i in range(qm)]


def calc_channel(ch, qm=20):
    """Ряд собирается «по точкам подряд» (режим 0 CalcD), как при ручном вводе."""
    ch.result, ch.error = None, ""
    try:
        if len(ch.points) < 2:
            raise calcd.CalcError("меньше двух точек")
        if any(p.xd is None for p in ch.points):
            raise calcd.CalcError("нет действительного значения (столбец I)")
        values = []
        for p in ch.points:
            values += expand(p.measured, qm)
        ch.result = calcd.calc(values, [p.xd for p in ch.points], qm, 0)
    except (calcd.CalcError, ValueError, ZeroDivisionError) as e:
        ch.error = str(e)
    return ch
