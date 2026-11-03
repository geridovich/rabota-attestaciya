"""«Полечить файл»: привести Excel аттестации к одному виду (как ТА-6).

Что встречалось в файлах заказчика и что здесь исправляется:
  * столбцы сдвинуты или идут в другом порядке (ГЭС-6 без «№», ТА-3) -
    данные переносятся в стандартные столбцы A..Y;
  * диапазон текстом «0…1,02», единица измерения строкой ниже;
  * заметки в столбце KKS («НЕТ НА МНЕМОСХЕМЕ») - переносятся в «Заметки»;
  * пустые строки-заготовки без названия и диапазона - удаляются;
  * нет «№» - каналы нумеруются по порядку;
  * числа записаны текстом («0,5»), переносы строк в названиях;
  * нет действительного значения (I) или формул N, O, Q, S - формулы ставятся заново;
  * нет допустимой погрешности (R) - формула 1,1·√(P² + W²);
  * измерение заполнено не во всех столбцах J..M - одинаковое значение дописывается.
Исходный файл не меняется, результат - копия «… (вылечено).xlsx».
"""
import os
import re

from .data import COLUMNS, is_kks, rows_from_xlsx, to_float, to_text

HEAD1 = {
    "A": "№", "B": "KKS", "C": "Наименование параметра", "D": "Диапазон измерения ИК",
    "G": "Контрольная точка", "J": "Измеренное значение выходного сигнала",
    "N": "Среднее значение погрешности %", "O": "Погрешность ИК",
    "R": "Значение допустимой погрешности ИК, %", "S": "Результат калибровки", "T": "ПИП",
    "U": "Погрешность ПИП, %", "V": "Модуль", "W": "Погрешность модуля, %",
    "X": "Допустимая погрешность ИК, %", "Y": "Заметки",
}
HEAD2 = {
    "D": "Нижняя граница", "E": "Верхняя граница", "F": "Ед.изм.", "G": "% ",
    "H": "Значение входного сигнала ", "I": "Значение выходного сигнала",
    "J": 1, "K": 2, "L": 3, "M": 4, "O": "ЭТ ИК", "P": "ПИП ИК", "Q": "ИК",
}
WIDTHS = {"A": 5, "B": 16, "C": 40, "D": 9, "E": 9, "F": 8, "G": 6, "H": 10, "I": 10, "J": 9,
          "K": 9, "L": 9, "M": 9, "N": 10, "O": 9, "P": 8, "Q": 9, "R": 11, "S": 10,
          "T": 16, "U": 10, "V": 12, "W": 10, "X": 11, "Y": 24}
MEAS = ("J", "K", "L", "M")
CHANNEL_COLS = ("A", "B", "C", "D", "E", "F", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y")


def _num(v, fixes, what):
    """Число; текстовое число («0,5») считается исправлением."""
    if v is None or isinstance(v, (int, float)) and not isinstance(v, bool):
        return v
    f = to_float(v)
    if f is not None:
        fixes[what] = fixes.get(what, 0) + 1
    return f


def _group(rows, fixes):
    """Строки -> каналы {'head': строка-заголовок, 'points': [...], 'notes': [...]}."""
    channels, cur, auto_num = [], None, 0
    c = COLUMNS
    for row in rows:
        kks_raw = row.get(c["kks"])
        kks = to_text(kks_raw)
        num = to_text(row.get(c["num"]))
        if kks and not is_kks(kks):
            # заметка в столбце KKS - к текущему каналу
            if cur is not None:
                cur["notes"].append(kks)
            fixes["Заметки из столбца KKS перенесены в «Заметки»"] = \
                fixes.get("Заметки из столбца KKS перенесены в «Заметки»", 0) + 1
            kks = ""
            row = dict(row)
            row[c["kks"]] = None
            if not num and not any(row.get(c[k]) is not None for k in ("pct", "m1")):
                continue
        if isinstance(kks_raw, str) and "\n" in kks_raw or \
                isinstance(row.get(c["name"]), str) and "\n" in row.get(c["name"]):
            fixes["Переносы строк в названиях убраны"] = fixes.get("Переносы строк в названиях убраны", 0) + 1
        if kks or num:
            cur = {"head": row, "points": [], "notes": []}
            channels.append(cur)
        if cur is None:
            continue
        if any(row.get(c[k]) is not None and str(row.get(c[k])).strip() != ""
               for k in ("pct", "output", "m1", "m2", "m3", "m4")):
            cur["points"].append(row)
        unit = to_text(row.get(c["unit"]))
        if row is not cur["head"] and unit and not to_text(cur["head"].get(c["unit"])):
            # ТА-3: «°C» строкой ниже диапазона
            cur["head"] = dict(cur["head"])
            cur["head"][c["unit"]] = unit
            fixes["Единица измерения взята из строки ниже"] = \
                fixes.get("Единица измерения взята из строки ниже", 0) + 1
    kept = []
    for ch in channels:
        h = ch["head"]
        name = to_text(h.get(c["name"]))
        low, high = to_float(h.get(c["low"])), to_float(h.get(c["high"]))
        if not to_text(h.get(c["kks"])) or (not name and (low is None or low == high)):
            fixes["Пустые строки-заготовки удалены"] = fixes.get("Пустые строки-заготовки удалены", 0) + 1
            continue
        kept.append(ch)
    for ch in kept:
        if not to_text(ch["head"].get(c["num"])):
            auto_num += 1
    if auto_num:
        fixes["Каналы без «№» пронумерованы"] = auto_num
    return kept


def heal(path, out_path=None, out_dir=None):
    """Вылечить файл. Возвращает (путь копии, {что исправлено: сколько}, каналов)."""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, Side

    fixes = {}
    rows = rows_from_xlsx(path)
    channels = _group(rows, fixes)
    c = COLUMNS

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Аттестация"
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col, text in HEAD1.items():
        ws[col + "1"] = text
    for col, text in HEAD2.items():
        ws[col + "2"] = text
    for i, col in enumerate(sorted(WIDTHS, key=lambda x: (len(x), x))):
        ws[col + "3"] = i + 1
        ws.column_dimensions[col].width = WIDTHS[col]
    for r in (1, 2, 3):
        for col in WIDTHS:
            ws[col + str(r)].alignment = center
            ws[col + str(r)].border = border
            ws[col + str(r)].font = Font(bold=r < 3)

    r = 4
    used_nums = set()
    seq = 0
    for ch in channels:
        h = ch["head"]
        first = r
        pts = ch["points"] or [h]
        last = first + len(pts) - 1
        num = to_text(h.get(c["num"]))
        seq += 1
        if not num or num in used_nums:
            num = str(seq)
        used_nums.add(num)
        ws["A%d" % first] = int(num) if num.isdigit() else num
        ws["B%d" % first] = to_text(h.get(c["kks"]))
        ws["C%d" % first] = to_text(h.get(c["name"]))
        ws["D%d" % first] = _num(h.get(c["low"]), fixes, "Числа, записанные текстом, исправлены")
        ws["E%d" % first] = _num(h.get(c["high"]), fixes, "Числа, записанные текстом, исправлены")
        ws["F%d" % first] = to_text(h.get(c["unit"]))
        for k, col in (("pip_ik", "P"), ("pip", "T"), ("pip_err", "U"), ("module", "V"),
                       ("module_err", "W"), ("allowed_ik", "X")):
            v = h.get(c[k])
            ws["%s%d" % (col, first)] = to_text(v) if k in ("pip", "module") else \
                _num(v, fixes, "Числа, записанные текстом, исправлены")
        allowed = _num(h.get(c["allowed"]), fixes, "Числа, записанные текстом, исправлены")
        if allowed is None:
            allowed = ws["X%d" % first].value
        if allowed is None and ws["P%d" % first].value is not None:
            allowed = "=1.1*SQRT(P{0}*P{0}+{1}*{1})".format(first, ws["W%d" % first].value or 0.15)
            fixes["Допустимая погрешность посчитана формулой"] = \
                fixes.get("Допустимая погрешность посчитана формулой", 0) + 1
        ws["R%d" % first] = allowed
        notes = [to_text(h.get(c["notes"]))] + ch["notes"]
        ws["Y%d" % first] = "; ".join(n for n in notes if n) or None

        for i, p in enumerate(pts):
            rr = first + i
            pct = _num(p.get(c["pct"]), fixes, "Числа, записанные текстом, исправлены")
            ws["G%d" % rr] = pct
            ws["H%d" % rr] = _num(p.get(c["input"]), fixes, "Числа, записанные текстом, исправлены")
            if pct is not None:
                ws["I%d" % rr] = "=(E{0}-D{0})*{1}%+D{0}".format(first, _fmt(pct))
            else:
                ws["I%d" % rr] = _num(p.get(c["output"]), fixes, "Числа, записанные текстом, исправлены")
            vals = [_num(p.get(c[m]), fixes, "Числа, записанные текстом, исправлены")
                    for m in ("m1", "m2", "m3", "m4")]
            filled = [v for v in vals if v is not None]
            if filled and len(filled) < 4 and len(set(filled)) == 1:
                vals = [filled[0]] * 4
                fixes["Измерения дописаны в пустые столбцы J..M"] = \
                    fixes.get("Измерения дописаны в пустые столбцы J..M", 0) + 1
            for col, v in zip(MEAS, vals):
                ws["%s%d" % (col, rr)] = v
            ws["N%d" % rr] = "=ABS(((AVERAGE(J{0}:M{0})-I{0})/(E{1}-D{1}))*100)".format(rr, first)
        ws["O%d" % first] = "=MAXA(N%d:N%d)" % (first, last)
        ws["Q%d" % first] = "=SQRT(O{0}*O{0}+P{0}*P{0})".format(first)
        ws["S%d" % first] = '=IF(Q{0}>R{0},"Не годен","Годен")'.format(first)

        for rr in range(first, last + 1):
            for col in WIDTHS:
                cell = ws["%s%d" % (col, rr)]
                cell.border = border
                cell.alignment = center
        if last > first:
            for col in CHANNEL_COLS:
                ws.merge_cells("%s%d:%s%d" % (col, first, col, last))
        r = last + 1

    ws.freeze_panes = "D4"
    wb.calculation.fullCalcOnLoad = True
    if out_path is None:
        src = os.path.join(out_dir, os.path.basename(path)) if out_dir else path
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(src)[0]
        base = re.sub(r" \(вылечено( \d+)?\)$", "", base)
        out_path, n = base + " (вылечено).xlsx", 2
        while os.path.exists(out_path):
            out_path, n = "%s (вылечено %d).xlsx" % (base, n), n + 1
    wb.save(out_path)
    return out_path, fixes, len(channels)


def _fmt(v):
    return ("%g" % v).replace(",", ".")
