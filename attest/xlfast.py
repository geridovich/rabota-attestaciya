"""Ускорение openpyxl при открытии файлов с тысячами объединённых ячеек.

openpyxl при чтении каждого объединения (в ТЭЦ-5 их ~57 тысяч) переносит рамки на
края и защиту на все ячейки, и каждый раз ищет такой же стиль в списке книги через
сравнение объектов. Это больше половины времени лечения и магии. Здесь то же самое
делается через номера стилей с запоминанием: результат тот же, но в разы быстрее.
"""
from openpyxl.cell.cell import MergedCell
from openpyxl.styles.borders import Border
from openpyxl.styles.cell_style import StyleArray
from openpyxl.worksheet.merge import MergedCellRange

_SIDES = ("top", "left", "right", "bottom")


def _memo(wb):
    m = getattr(wb, "_plusultra_border_memo", None)
    if m is None:
        m = {}
        wb._plusultra_border_memo = m
    return m


def _style(cell):
    if not cell._style:
        cell._style = StyleArray()
    return cell._style


def _add_border(wb, memo, base_id, key, extra):
    """Номер рамки «рамка base_id + extra» (как cell.border += extra)."""
    k = (base_id, key)
    bid = memo.get(k)
    if bid is None:
        bid = wb._borders.add(wb._borders[base_id] + extra())
        memo[k] = bid
    return bid


def _get_borders(self):
    ws = self.ws
    self.start_cell = ws._cells.get((self.min_row, self.min_col))
    if self.start_cell is None:
        self.start_cell = ws.cell(row=self.min_row, column=self.min_col)
    end_cell = ws._cells.get((self.max_row, self.max_col))
    if end_cell is not None:
        wb = ws.parent
        st, en = _style(self.start_cell), _style(end_cell)
        eb = en.borderId

        def extra():
            b = wb._borders[eb]
            return Border(right=b.right, bottom=b.bottom)
        st.borderId = _add_border(wb, _memo(wb), st.borderId, ("end", eb), extra)


def _format(self):
    ws = self.ws
    wb = ws.parent
    memo = _memo(wb)
    cells = ws._cells
    st = _style(self.start_cell)
    sb = st.borderId
    start_border = wb._borders[sb]
    for name in _SIDES:
        side = getattr(start_border, name)
        if side and side.style is None:
            continue
        key = ("side", sb, name)
        for coord in getattr(self, name):
            cell = cells.get(coord)
            if cell is None:
                cell = MergedCell(ws, row=coord[0], column=coord[1])
                cells[coord] = cell
            cs = _style(cell)
            cs.borderId = _add_border(wb, memo, cs.borderId, key,
                                      lambda: Border(**{name: getattr(start_border, name)}))
    # защита: у openpyxl она есть всегда - копия защиты левой верхней ячейки во все
    pid = st.protectionId
    for coord in self.cells:
        cell = cells.get(coord)
        if cell is None:
            cell = MergedCell(ws, row=coord[0], column=coord[1])
            cells[coord] = cell
        _style(cell).protectionId = pid


def install():
    if getattr(MergedCellRange, "_plusultra_fast", False):
        return
    MergedCellRange._get_borders = _get_borders
    MergedCellRange.format = _format
    MergedCellRange._plusultra_fast = True
