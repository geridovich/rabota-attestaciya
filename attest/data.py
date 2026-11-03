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
    """Число из ячейки; пусто, текст, NaN и бесконечность -> None."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(v) else None
    s = str(v).strip().replace(" ", "").replace("−", "-").replace(" ", "").replace(",", ".")
    if s.endswith("%"):
        s = s[:-1]
    if not s:
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return f if math.isfinite(f) else None


def is_blank(v):
    return v is None or not str(v).strip()


def bad_number(v):
    """В ячейке что-то есть, но это не число («ОШИБКА», NaN, #ДЕЛ/0!)."""
    return not is_blank(v) and to_float(v) is None


def to_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    # перенос строки внутри ячейки Excel (Alt+Enter) - в пробел
    return re.sub(r"\s*\n\s*", " ", str(v)).strip()


def is_kks(s, num=None, name=None):
    """Начинает ли строка канал по столбцу KKS.

    Обычно позиция - одно «слово» (10HBK50CQ004). Бывает и с пробелом
    («AA00TU03 LCA30CP001» в ТГ-36) - тогда канал, если есть № или наименование.
    Заметки вида «НЕТ НА МНЕМОСХЕМЕ» (ТА-3) стоят без № и без наименования.
    """
    s = to_text(s)
    return bool(s) and (" " not in s or bool(to_text(num)) or bool(to_text(name)))


def split_clipboard(text):
    """Текст из буфера Excel -> строки ячеек.

    Ячейку с переносом строки Excel берёт в кавычки ("...""..."), поэтому
    простое деление по строкам рвёт такой канал на две строки.
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    rows, row, i, n = [], [], 0, len(text)
    while True:
        quoted = False
        if i < n and text[i] == '"':
            j, buf = i + 1, []
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(text[j])
                j += 1
            # ячейка в кавычках, только если закрывающая кавычка стоит в конце ячейки;
            # иначе кавычка - просто часть текста («"Б" датчик»)
            if j < n and (j + 1 == n or text[j + 1] in "\t\n"):
                row.append("".join(buf))
                i = j + 1
                quoted = True
        if not quoted:
            k = i
            while k < n and text[k] not in "\t\n":
                k += 1
            row.append(text[i:k])
            i = k
        if i >= n:
            rows.append(row)
            return rows
        if text[i] == "\t":
            i += 1
        else:  # перевод строки - конец строки таблицы
            rows.append(row)
            row = []
            i += 1
            if i >= n:
                return rows


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
    bad_cells: list = field(default_factory=list)  # ячейки с не-числом вместо числа

    # --- итоговые величины ---
    @property
    def et_ik(self):
        """Погрешность ЭТ ИК, как в Excel (столбец O = макс. приведённая погрешность по точкам).

        Коэффициенты gh/gl из расчёта идут в протокол, а годность определяется
        так же, как в Excel заказчика.
        """
        return self.excel_err if self.result else None

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
        total, allowed = self.total, self.allowed_value
        if self.error or total is None or allowed is None:
            return ""
        if not (math.isfinite(total) and math.isfinite(allowed)):
            return ""  # NaN в сравнении даёт False - без проверки вышло бы «Годен»
        return "Не годен" if total > allowed else "Годен"

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
    lines = [[to_text(c) for c in ln] for ln in split_clipboard(text) if any(c.strip() for c in ln)]
    if not lines:
        return []
    shift = 1 if _is_kks(lines[0][0]) else 0  # вставка начата с KKS, без №
    width = max(len(ln) for ln in lines) + shift
    rows = []
    if width >= FULL_ROW_MIN + 6:
        for ln in lines:
            rows.append({col_letter(i + shift): c for i, c in enumerate(ln)})
        return rows

    # есть ли столбец «%», решаем отдельно для каждого канала
    blocks, with_pct = [], []
    for ln in lines:
        cells = [""] * shift + ln
        if not blocks or (len(cells) > 1 and (cells[0] or cells[1])):
            blocks.append([])
        blocks[-1].append(cells)
    for block in blocks:
        flag = _has_pct_column(block)
        with_pct += [flag] * len(block)

    low = high = None
    for ln, has_pct in zip(lines, with_pct):
        keys = SHORT_KEYS if has_pct else SHORT_KEYS[:-1]
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
        if not has_pct and low is not None and high != low and to_float(row.get(COLUMNS["output"])) is not None:
            row[COLUMNS["pct"]] = (to_float(row[COLUMNS["output"]]) - low) / (high - low) * 100.0
        rows.append(row)
    return rows


def rows_from_xlsx(path, first_row=4, overrides=None):
    """Строки листа с измерениями. overrides: {строка Excel: значение для J..M} (magic)."""
    import openpyxl
    from . import layout
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws, lay = layout.pick_sheet(wb)
    keys = [k for k in COLUMNS if k not in ("m1", "m2", "m3", "m4")]
    m_std = [col_letter(col_index(COLUMNS["m1"]) + k) for k in range(len(lay["m"]))]
    rows = []
    for n, r in enumerate(ws.iter_rows(min_row=first_row, values_only=True), start=first_row):
        if not any(v is not None and str(v).strip() for v in r):
            continue
        raw = {col_letter(i): v for i, v in enumerate(r)}
        # раскладка листа -> стандартные столбцы (как в ТА-6), дальше всё одинаково
        row = {COLUMNS[k]: raw.get(lay[k]) for k in keys if k in lay}
        for std, src in zip(m_std, lay["m"]):
            row[std] = overrides[n] if overrides and n in overrides else raw.get(src)
        if "low" not in lay and lay.get("range") and raw.get(lay["range"]) is not None:
            low, high = layout.parse_range(raw[lay["range"]])
            if low is None and high is None and "unit" not in lay:
                row[COLUMNS["unit"]] = raw[lay["range"]]  # ТА-3: единица строкой ниже диапазона
            else:
                row[COLUMNS["low"]], row[COLUMNS["high"]] = low, high
        rows.append(row)
    wb.close()
    return rows


def measured_values(row, bad=None):
    """Измерения: J..M в листе, во вставке жёлтых столбцов - все после действительного.

    Пустые ячейки пропускаются, а не-числа записываются в bad (буква столбца),
    чтобы не считать канал по неполному ряду молча.
    """
    last = col_index(COLUMNS["m4"]) if not row.get("_short") else 200
    out = []
    for i in range(col_index(COLUMNS["m1"]), last + 1):
        raw = row.get(col_letter(i))
        v = to_float(raw)
        if v is not None:
            out.append(v)
        elif bad is not None and bad_number(raw):
            bad.append(col_letter(i))
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
        if not is_kks(kks, g(row, "num"), g(row, "name")):  # заметки «НЕТ НА МНЕМОСХЕМЕ»
            kks = ""
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
            cur.bad_cells += [COLUMNS[k] for k in ("low", "high", "pip_ik", "pip_err", "allowed")
                              if bad_number(g(row, k))]
        if cur is None:
            continue
        if not cur.unit and to_text(g(row, "unit")):
            cur.unit = to_text(g(row, "unit"))
        bad = []
        measured = measured_values(row, bad)
        if bad_number(g(row, "output")):
            bad.append(COLUMNS["output"])
        if bad:
            cur.bad_cells.append("%s (точка %d)" % (", ".join(bad), len(cur.points) + 1))
        if not measured:
            continue
        pct = to_float(g(row, "pct"))
        xd = to_float(g(row, "output"))
        if xd is None and pct is not None:  # формула не посчитана - считаем сами
            xd = (cur.high - cur.low) * pct / 100.0 + cur.low
        cur.points.append(Point(pct, to_float(g(row, "input")), xd, measured))
    # заготовки без KKS и пустые строки-шаблоны (без названия и диапазона) не нужны
    return [c for c in channels if c.kks and (c.name or c.high != c.low)]


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
        if ch.bad_cells:
            raise calcd.CalcError("не число в ячейках: " + "; ".join(ch.bad_cells))
        if ch.high == ch.low:
            raise calcd.CalcError("нулевой диапазон (нижняя граница = верхней)")
        if len(ch.points) < 2:
            raise calcd.CalcError("нет измеренных значений" if not ch.points else "меньше двух точек")
        if any(p.xd is None for p in ch.points):
            raise calcd.CalcError("нет действительного значения (столбец I)")
        values = []
        for p in ch.points:
            values += expand(p.measured, qm)
        ch.result = calcd.calc(values, [p.xd for p in ch.points], qm, 0)
    except (calcd.CalcError, ValueError, ZeroDivisionError) as e:
        ch.error = str(e)
    return ch
