"""«Полечить файл»: сдвинутые столбцы, заметки в KKS, диапазон текстом, заготовки."""
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from attest import data, heal  # noqa: E402


def make_messy(path):
    """Как ТА-3: без «Нижняя/Верхняя», диапазон текстом, единица строкой ниже, заметки в KKS."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    h1 = {"A": "№", "B": "KKS", "C": "Наименование параметра", "D": "Диапазон измерения ИК",
          "E": "Контрольная точка", "H": "Измеренное значение выходного сигнала"}
    h2 = {"E": "% ", "F": "Значение входного сигнала ", "G": "Значение выходного сигнала",
          "H": 1, "I": 2, "J": 3, "K": 4}
    for c, v in h1.items():
        ws[c + "1"] = v
    for c, v in h2.items():
        ws[c + "2"] = v
    r = 4
    for n, (kks, rng) in enumerate((("SG23P001", "0…1,02"), ("SC21T001", "0…100")), start=1):
        ws["A%d" % r], ws["B%d" % r], ws["C%d" % r], ws["D%d" % r] = n, kks, "Канал\n%d" % n, rng
        ws["D%d" % (r + 1)] = "°C"
        for k, pct in enumerate((0, 25, 50, 75, 100)):
            ws["E%d" % (r + k)] = pct
            ws["H%d" % (r + k)] = "0,5" if k == 2 else 0.1 * k
        ws["B%d" % (r + 5)] = "НЕТ НА МНЕМОСХЕМЕ"
        r += 6
    ws["A%d" % r] = 99  # пустая заготовка
    wb.save(path)


class HealTest(unittest.TestCase):
    def test_heal_messy(self):
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "messy.xlsx")
            make_messy(src)
            out, fixes, n = heal.heal(src)
            self.assertEqual(n, 2)
            self.assertIn("Заметки из столбца KKS перенесены в «Заметки»", fixes)
            self.assertIn("Единица измерения взята из строки ниже", fixes)
            chs = data.channels_from_rows(data.rows_from_xlsx(out))
            self.assertEqual([c.kks for c in chs], ["SG23P001", "SC21T001"])
            c = chs[0]
            self.assertEqual((c.low, c.high, c.unit, c.name), (0.0, 1.02, "°C", "Канал 1"))
            self.assertEqual(c.notes, "НЕТ НА МНЕМОСХЕМЕ")
            self.assertEqual(c.points[2].measured, [0.5, 0.5, 0.5, 0.5])
            import openpyxl
            ws = openpyxl.load_workbook(out).active
            self.assertEqual(ws["I5"].value, "=(E4-D4)*25%+D4")
            self.assertTrue(str(ws["S4"].value).startswith("=IF(Q4>R4"))


if __name__ == "__main__":
    unittest.main()
