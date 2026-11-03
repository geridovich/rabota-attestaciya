"""Раскладка листа аттестации по шапке (строки 1-2).

Файлы заказчика похожи, но столбцы бывают сдвинуты: в ГЭС-6 нет «№», в ТА-3
нет границ диапазона (диапазон текстом «0…1,02»). Поэтому столбцы ищутся по
заголовкам, а не по буквам.
"""
from .data import col_letter

# ключ -> (строка шапки, проверка текста)
ROW1 = (
    ("num", lambda t: t == "№"),
    ("kks", lambda t: t == "KKS"),
    ("name", lambda t: t.startswith("Наименование")),
    ("range", lambda t: t.startswith("Диапазон измерения")),
    ("measured", lambda t: t.startswith("Измеренное значение")),
    ("allowed", lambda t: t.startswith("Значение допустимой погрешности")),
    ("verdict", lambda t: t.startswith("Результат")),
    ("pip", lambda t: t == "ПИП"),
    ("pip_err", lambda t: t.startswith("Погрешность ПИП")),
    ("module", lambda t: t == "Модуль"),
    ("module_err", lambda t: t.startswith("Погрешность модуля")),
    ("allowed_ik", lambda t: t.startswith("Допустимая погрешность ИК")),
    ("notes", lambda t: t.startswith("Заметки")),
)
ROW2 = (
    ("low", lambda t: t.startswith("Нижняя")),
    ("high", lambda t: t.startswith("Верхняя")),
    ("unit", lambda t: t.startswith("Ед")),
    ("pct", lambda t: t.strip() == "%"),
    ("input", lambda t: t.startswith("Значение входного")),
    ("output", lambda t: t.startswith("Значение выходного")),
    ("et", lambda t: t == "ЭТ ИК"),
    ("pip_ik", lambda t: t == "ПИП ИК"),
    ("q", lambda t: t == "ИК"),
)


def _text(v):
    return " ".join(str(v).split()) if v is not None else ""


def detect(header1, header2):
    """header1/header2 - значения строк 1 и 2. Возвращает {ключ: буква столбца}."""
    lay = {}
    for i, v in enumerate(header1):
        t = _text(v)
        for key, ok in ROW1:
            if key not in lay and t and ok(t):
                lay[key] = col_letter(i)
    for i, v in enumerate(header2):
        t = _text(v)
        for key, ok in ROW2:
            if key not in lay and t and ok(t):
                lay[key] = col_letter(i)
    # столбцы измерений: под «Измеренное значение…» номера 1, 2, 3, 4
    if "measured" in lay:
        start = _index(lay["measured"])
        cols = []
        for i in range(start, len(header2)):
            if not isinstance(header2[i], (int, float)) or isinstance(header2[i], bool):
                break
            cols.append(col_letter(i))
        lay["m"] = cols or [col_letter(start + k) for k in range(4)]
    # «ПИП ИК» иногда заменено числом - тогда это столбец сразу после «ЭТ ИК»
    if "pip_ik" not in lay and "et" in lay:
        lay["pip_ik"] = col_letter(_index(lay["et"]) + 1)
    return lay


def _index(letter):
    n = 0
    for ch in letter:
        n = n * 26 + ord(ch) - 64
    return n - 1


def is_attestation(lay):
    return "kks" in lay and "m" in lay


def parse_range(text):
    """«0…1,02» / «-50...200» -> (нижняя, верхняя)."""
    from .data import to_float
    s = str(text).replace("...", "…").replace("..", "…")
    if "…" not in s:
        return None, None
    a, b = s.split("…", 1)
    return to_float(a), to_float(b.split()[0] if b.split() else b)


def pick_sheet(wb):
    """Первый лист, похожий на лист аттестации; (лист, раскладка)."""
    for ws in wb.worksheets:
        head = [list(r) for r in ws.iter_rows(min_row=1, max_row=2, values_only=True)]
        while len(head) < 2:
            head.append([])
        lay = detect(head[0], head[1])
        if is_attestation(lay):
            return ws, lay
    raise ValueError("В файле нет листа с «KKS» и «Измеренное значение выходного сигнала» в шапке")
