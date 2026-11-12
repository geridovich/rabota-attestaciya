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
  * остальные листы файла (например «Перечень ИК») переносятся в копию как есть;
  * измерение заполнено не во всех столбцах J..M - одинаковое значение дописывается.
Исходный файл не меняется, результат - копия «… (вылечено).xlsx».
"""
import copy
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
        if kks and not is_kks(kks, num, row.get(c["name"])):
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
            cur = {"head": row, "points": [], "notes": [], "hidden": bool(row.get("_hidden"))}
            channels.append(cur)
        if cur is None:
            continue
        if row is not cur["head"]:  # канал скрыт, только если скрыты все его строки
            cur["hidden"] = cur["hidden"] and bool(row.get("_hidden"))
        if row.get("_hidden"):
            if any(row.get(c[k]) is not None and str(row.get(c[k])).strip() != ""
                   for k in ("pct", "output", "m1", "m2", "m3", "m4")):
                fixes["Скрытые строки точек убраны"] = fixes.get("Скрытые строки точек убраны", 0) + 1
        elif any(row.get(c[k]) is not None and str(row.get(c[k])).strip() != ""
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
        if ch["hidden"]:  # скрытые в Excel каналы не учитываются и в копию не идут
            fixes["Скрытые каналы пропущены"] = fixes.get("Скрытые каналы пропущены", 0) + 1
            continue
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


_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")


def _source_formula(src_ws, h, std_col, first):
    """Формула ячейки исходного листа, переписанная на столбцы и строку копии.

    Файлы, сохранённые не Excel (например, заполненные скриптом), хранят формулы без
    посчитанных значений: «=P4» читается как пусто, и столбец терялся бы. Ссылки
    переводятся по таблице столбцов исходник -> копия; если ссылка ведёт в столбец,
    которого в копии нет, формула не переносится."""
    src = h.get("_src") or {}
    sc, sr = src.get(std_col), h.get("_row")
    if not sc or not sr:
        return None
    f = src_ws["%s%d" % (sc, sr)].value
    if not isinstance(f, str) or not f.startswith("="):
        return None
    back = {v: k for k, v in src.items()}
    ok = [True]

    def sub(m):
        col = back.get(m.group(2))
        if col is None:
            ok[0] = False
            return m.group(0)
        row = int(m.group(4))
        return "%s%s%s%d" % (m.group(1), col, m.group(3), row if m.group(3) else row - sr + first)
    out = _REF.sub(sub, f)
    return out if ok[0] else None


def heal(path, out_path=None, out_dir=None, progress=None):
    """Вылечить файл. Возвращает (путь копии, {что исправлено: сколько}, каналов)."""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, Side
    from . import xlfast
    xlfast.install()

    fixes = {}
    tick = progress or (lambda f: None)  # доля сделанного 0..1 - для процентов в окне
    tick(0.02)
    rows = rows_from_xlsx(path)
    tick(0.15)
    channels = _group(rows, fixes)
    c = COLUMNS

    # Копия строится прямо в исходной книге: новый лист аттестации ставится на место
    # старого, остальные листы остаются как есть - их не нужно копировать по ячейкам.
    try:
        from . import layout
        wb = openpyxl.load_workbook(path)
        old_ws, _lay = layout.pick_sheet(wb)
        ws = wb.create_sheet("Аттестация (новая)", wb.index(old_ws))
    except Exception:  # noqa: BLE001 - не xlsx или нет листа аттестации - новая книга
        wb, old_ws = openpyxl.Workbook(), None
        ws = wb.active
        ws.title = "Аттестация"
    tick(0.35)
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
    base_style = None
    row_map = {1: 1, 2: 2, 3: 3}  # строка копии -> строка исходного листа (для оформления)
    # номер протокола = № канала, поэтому номера должны быть разными;
    # новый номер не должен совпасть ни с уже выданным, ни с чьим-то ещё в файле
    taken = {to_text(ch["head"].get(c["num"])) for ch in channels}
    used_nums = set()
    seq = 0
    for chn, ch in enumerate(channels):
        if chn % 20 == 0:
            tick(0.35 + 0.25 * chn / max(1, len(channels)))
        h = ch["head"]
        first = r
        pts = ch["points"] or [h]
        last = first + len(pts) - 1
        num = to_text(h.get(c["num"]))
        seq += 1
        if not num or num in used_nums:
            if num:
                fixes["Повторяющиеся № заменены свободными"] = fixes.get("Повторяющиеся № заменены свободными", 0) + 1
            n = seq
            while str(n) in used_nums or (str(n) in taken and str(n) != num):
                n += 1
            num = str(n)
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
            if v is None and old_ws is not None:  # формула без сохранённого значения (=P4)
                f = _source_formula(old_ws, h, c[k], first)
                if f:
                    ws["%s%d" % (col, first)] = f
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
            row_map[rr] = p.get("_row") or h.get("_row")
            pct = _num(p.get(c["pct"]), fixes, "Числа, записанные текстом, исправлены")
            ws["G%d" % rr] = pct
            ws["H%d" % rr] = _num(p.get(c["input"]), fixes, "Числа, записанные текстом, исправлены")
            if pct is not None:
                ws["I%d" % rr] = "=ROUND((E{0}-D{0})*{1}%+D{0},10)".format(first, _fmt(pct))
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
            # ROUND: иначе Excel показывает шум двоичных дробей (0,6*75% = 0,44999…) как 9,25E-15
            ws["N%d" % rr] = "=ROUND(ABS(((AVERAGE(J{0}:M{0})-I{0})/(E{1}-D{1}))*100),9)".format(rr, first)
        ws["O%d" % first] = "=MAXA(N%d:N%d)" % (first, last)
        ws["Q%d" % first] = "=SQRT(O{0}*O{0}+P{0}*P{0})".format(first)
        ws["S%d" % first] = '=IF(Q{0}>R{0},"Не годен","Годен")'.format(first)

        for rr in range(first, last + 1):
            for col in WIDTHS:
                cell = ws["%s%d" % (col, rr)]
                if base_style is None:  # рамка + выравнивание один раз, дальше - копия набора стилей
                    cell.border = border
                    cell.alignment = center
                    base_style = copy.copy(cell._style)
                else:
                    cell._style = copy.copy(base_style)
        if last > first:
            for col in CHANNEL_COLS:
                _merge(ws, "%s%d:%s%d" % (col, first, col, last))
        r = last + 1

    ws.freeze_panes = "D4"
    tick(0.6)
    if old_ws is not None:
        if _apply_source_look(ws, path, row_map, wb, lambda f: tick(0.6 + 0.3 * f)):
            fixes["Ширина столбцов, высота строк и оформление взяты из исходного файла"] = 1
        wb.remove(old_ws)
        ws.title = "Аттестация" if "Аттестация" not in wb.sheetnames else "Аттестация (вылечено)"
        wb.active = wb.index(ws)
        others = [sh.title for sh in wb.worksheets if sh is not ws]
        if others:
            fixes["Остальные листы перенесены без изменений (%s)" % ", ".join(others)] = len(others)
    tick(0.9)
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
    tick(1.0)
    return out_path, fixes, len(channels)


def _apply_source_look(ws, path, row_map, src=None, tick=None):
    """Ширина столбцов, высота строк, оформление ячеек и объединения шапки - как в исходном
    листе. Столбцы сопоставляются по заголовкам, поэтому работает и для сдвинутых раскладок.
    Протоколам и перечню оформление не нужно: они читают только значения."""
    import copy
    import openpyxl
    from openpyxl.utils import column_index_from_string, get_column_letter
    from . import layout

    try:
        src = src or openpyxl.load_workbook(path)
        sws, lay = layout.pick_sheet(src, exclude=ws)
    except Exception:  # noqa: BLE001 - не xlsx или нет листа аттестации
        return False
    # стандартный столбец -> столбец исходного листа
    std_keys = {"A": "num", "B": "kks", "C": "name", "D": "low", "E": "high", "F": "unit", "G": "pct",
                "H": "input", "I": "output", "O": "et", "P": "pip_ik", "Q": "q", "R": "allowed",
                "S": "verdict", "T": "pip", "U": "pip_err", "V": "module", "W": "module_err",
                "X": "allowed_ik", "Y": "notes"}
    cmap = {std: lay[k] for std, k in std_keys.items() if lay.get(k)}
    if "D" not in cmap and lay.get("range"):  # ТА-3: диапазон текстом в одном столбце
        cmap["D"] = lay["range"]
    for std, srcc in zip(MEAS, lay.get("m", [])):
        cmap[std] = srcc
    if lay.get("et"):  # «Среднее значение погрешности» - столбец перед «ЭТ ИК»
        prev = column_index_from_string(lay["et"]) - 1
        if prev >= 1 and get_column_letter(prev) not in cmap.values():
            cmap["N"] = get_column_letter(prev)

    # ширины: у исходника бывают группы столбцов (min..max) и ширина «по умолчанию»
    widths = {}
    for dim in sws.column_dimensions.values():
        if dim.width:
            lo = dim.min or column_index_from_string(dim.index)
            for i in range(lo, (dim.max or lo) + 1):
                widths[i] = dim.width
    fmt = sws.sheet_format
    ws.sheet_format.defaultColWidth = fmt.defaultColWidth
    ws.sheet_format.baseColWidth = fmt.baseColWidth
    ws.sheet_format.defaultRowHeight = fmt.defaultRowHeight
    for std, srcc in cmap.items():
        w = widths.get(column_index_from_string(srcc))
        if w:
            ws.column_dimensions[std].width = w
        elif std in ws.column_dimensions:  # у исходника ширина по умолчанию - и у копии
            del ws.column_dimensions[std]
    looks = {}  # одинаковые стили копируются один раз - иначе тысячи строк идут минутами
    same_book = sws.parent is ws.parent
    total = max(1, len(row_map))
    for k, (dst_row, src_row) in enumerate(row_map.items()):
        if tick and k % 200 == 0:
            tick(k / total)
        if not src_row:
            continue
        h = sws.row_dimensions[src_row].height if src_row in sws.row_dimensions else None
        ws.row_dimensions[dst_row].height = h
        for std, srcc in cmap.items():
            sc = sws["%s%d" % (srcc, src_row)]
            if not sc.has_style:
                continue
            d = ws["%s%d" % (std, dst_row)]
            formula = isinstance(d.value, str) and d.value.startswith("=")
            if same_book:  # та же книга - набор стилей переносится как есть
                d._style = copy.copy(sc._style)
                if formula and sc.number_format == "@":  # «текстовый» формат ломает формулу
                    d.number_format = "General"
                continue
            key = (tuple(sc._style), formula, tuple(d._style) if formula else None)
            done = looks.get(key)
            if done is not None:  # такой же стиль уже переносили - копия готового набора
                d._style = copy.copy(done)
                continue
            d.font, d.border, d.fill = copy.copy(sc.font), copy.copy(sc.border), copy.copy(sc.fill)
            d.alignment = copy.copy(sc.alignment)
            if not formula:
                d.number_format = sc.number_format
            looks[key] = copy.copy(d._style)
    # объединения шапки (строки 1-2): «Диапазон измерения ИК» над D..F и т.п.
    inv = {}
    for std, srcc in cmap.items():
        inv[column_index_from_string(srcc)] = column_index_from_string(std)
    for rng in sws.merged_cells.ranges:
        if rng.max_row > 2:
            continue
        a, b = inv.get(rng.min_col), inv.get(rng.max_col)
        if a and b and b - a == rng.max_col - rng.min_col:
            try:
                _merge(ws, start_row=rng.min_row, start_column=a, end_row=rng.max_row, end_column=b)
            except ValueError:
                pass
    return True


def _copy_other_sheets(path, wb, src=None, tick=None):
    """Листы исходного файла, кроме листа аттестации, - в копию как есть.

    Раньше копия содержала только лист «Аттестация», и, например, «Перечень ИК» пропадал.
    Переносятся значения и формулы, оформление ячеек, объединения, ширина столбцов и высота строк.
    """
    import copy
    import openpyxl
    from . import layout

    try:
        src = src or openpyxl.load_workbook(path)
        att, _lay = layout.pick_sheet(src)
    except Exception:  # не xlsx или нет листа аттестации - копировать нечего
        return []
    copied = []
    looks = {}  # одинаковые стили копируются один раз
    sheets = list(src.worksheets)
    for si, sws in enumerate(sheets):
        if tick:
            tick(si / max(1, len(sheets)))
        if sws is att:
            continue
        title = sws.title
        while title in wb.sheetnames:
            title += " (исх.)"
        dws = wb.create_sheet(title)
        for row in sws.iter_rows():
            for cell in row:
                if cell.__class__.__name__ == "MergedCell":
                    continue
                d = dws.cell(row=cell.row, column=cell.column, value=cell.value)
                if cell.has_style:  # индексы стилей у книг разные - копируем сами стили
                    key = tuple(cell._style)
                    done = looks.get(key)
                    if done is not None:  # копия готового набора стилей - без поиска по таблицам
                        d._style = copy.copy(done)
                        continue
                    d.font, d.border, d.fill = copy.copy(cell.font), copy.copy(cell.border), copy.copy(cell.fill)
                    d.alignment, d.protection = copy.copy(cell.alignment), copy.copy(cell.protection)
                    d.number_format = cell.number_format
                    looks[key] = copy.copy(d._style)
        for rng in sws.merged_cells.ranges:
            _merge(dws, str(rng))
        for key, dim in sws.column_dimensions.items():
            dws.column_dimensions[key].width = dim.width
            dws.column_dimensions[key].hidden = dim.hidden
        for key, dim in sws.row_dimensions.items():
            dws.row_dimensions[key].height = dim.height
            dws.row_dimensions[key].hidden = dim.hidden
        dws.freeze_panes = sws.freeze_panes
        copied.append(title)
    return copied


def _merge(ws, *args, **kw):
    """Объединить ячейки без перерисовки рамок (openpyxl делает это очень медленно;
    рамки и так переносятся из исходного файла по ячейкам)."""
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.merge import MergedCellRange
    coord = args[0] if args else "%s%d:%s%d" % (get_column_letter(kw["start_column"]), kw["start_row"],
                                                 get_column_letter(kw["end_column"]), kw["end_row"])
    # прямо в множество: merged_cells.add() перебирает все уже объединённые
    # диапазоны, и на тысячах каналов это минуты
    ws.merged_cells.ranges.add(MergedCellRange(ws, coord))


def _fmt(v):
    return ("%g" % v).replace(",", ".")
