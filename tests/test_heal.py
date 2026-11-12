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
            self.assertEqual(ws["I5"].value, "=ROUND((E4-D4)*25%+D4,10)")
            self.assertTrue(str(ws["S4"].value).startswith("=IF(Q4>R4"))

    def test_other_sheets_kept(self):
        """Лист «Перечень ИК» рядом с листом аттестации не пропадает после лечения."""
        import openpyxl
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "two.xlsx")
            make_messy(src)
            wb = openpyxl.load_workbook(src)
            ws = wb.create_sheet("Перечень ИК")
            ws["A1"], ws["B1"], ws["A4"], ws["B4"] = "№", "KKS", 1, "SG23P001"
            ws.merge_cells("A1:A2")
            ws.column_dimensions["B"].width = 17
            wb.save(src)
            out, fixes, _n = heal.heal(src)
            res = openpyxl.load_workbook(out)
            self.assertEqual(res.sheetnames, ["Аттестация", "Перечень ИК"])
            p = res["Перечень ИК"]
            self.assertEqual((p["A4"].value, p["B4"].value), (1, "SG23P001"))
            self.assertIn("A1:A2", [str(r) for r in p.merged_cells.ranges])
            self.assertEqual(p.column_dimensions["B"].width, 17)
            self.assertTrue(any(k.startswith("Остальные листы перенесены") for k in fixes))

    def test_input_signal(self):
        """H: температуры без 4-20 мА - действительные значения, пустой H у остальных - 4..20."""
        import openpyxl
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "inp.xlsx")
            wb = openpyxl.Workbook()
            ws = wb.active
            for col, text in heal.HEAD1.items():
                ws[col + "1"] = text
            for col, text in heal.HEAD2.items():
                ws[col + "2"] = text
            chans = (  # KKS, низ, верх, ед., ПИП, модуль, H
                ("T1CT001", 0, 600, "°C", "Метран-ТХА", "LIG16", (4, 8, 12, 16, 20)),
                ("T2GW001", -40, 80, "°C", "Метран-ТСМ-243", "AIG16", None),
                ("T3CT001", -50, 150, "°С", None, "МФК 1500", None),
                ("P1CP001", 0, 1.6, "МПа", "Метран 100ДИ", "AIG16", None),
                ("P2CP001", 0, 1.6, "МПа", "Метран 100ДИ", "AIG16", (0, 0.4, 0.8, 1.2, 1.6)),
            )
            r = 4
            for n, (kks, lo, hi, u, pip, mod, h) in enumerate(chans, start=1):
                ws["A%d" % r], ws["B%d" % r], ws["C%d" % r] = n, kks, "Канал %d" % n
                ws["D%d" % r], ws["E%d" % r], ws["F%d" % r] = lo, hi, u
                ws["T%d" % r], ws["V%d" % r], ws["P%d" % r], ws["W%d" % r] = pip, mod, 0.5, 0.15
                for k, pct in enumerate((0, 25, 50, 75, 100)):
                    ws["G%d" % (r + k)] = pct
                    if h:
                        ws["H%d" % (r + k)] = h[k]
                r += 5
            wb.save(src)
            out, fixes, n = heal.heal(src)
            self.assertEqual(n, 5)
            res = openpyxl.load_workbook(out).active
            got = [[res["H%d" % (4 + 5 * i + k)].value for k in range(5)] for i in range(5)]
            self.assertEqual(got[0], [0, 150, 300, 450, 600])
            self.assertEqual(got[1], [4, 8, 12, 16, 20])
            self.assertEqual(got[2], [-50, 0, 50, 100, 150])
            self.assertEqual(got[3], [4, 8, 12, 16, 20])
            self.assertEqual(got[4], [0, 0.4, 0.8, 1.2, 1.6])
            self.assertEqual(fixes["Входной сигнал температур = действительные значения"], 2)
            self.assertEqual(fixes["Пустой входной сигнал заполнен 4…20 мА"], 2)


if __name__ == "__main__":
    unittest.main()
