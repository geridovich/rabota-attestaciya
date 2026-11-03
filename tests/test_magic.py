"""Magic: заполнение пустых измерений так, чтобы все каналы были годны."""
import os
import random
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from attest import data, magic  # noqa: E402


def make_sheet(path, channels):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    head1 = {"A": "№", "B": "KKS", "C": "Наименование параметра", "D": "Диапазон измерения ИК",
             "G": "Контрольная точка", "J": "Измеренное значение выходного сигнала",
             "O": "Погрешность ИК", "R": "Значение допустимой погрешности ИК, %"}
    head2 = {"D": "Нижняя граница", "E": "Верхняя граница", "F": "Ед.изм.", "G": "% ",
             "H": "Значение входного сигнала ", "I": "Значение выходного сигнала",
             "J": 1, "K": 2, "L": 3, "M": 4, "O": "ЭТ ИК", "P": "ПИП ИК", "Q": "ИК"}
    for c, v in head1.items():
        ws[c + "1"] = v
    for c, v in head2.items():
        ws[c + "2"] = v
    r = 4
    for n, (low, high, p, allowed) in enumerate(channels, start=1):
        ws["A%d" % r], ws["B%d" % r], ws["C%d" % r] = n, "KKS%03d" % n, "Канал %d" % n
        ws["D%d" % r], ws["E%d" % r], ws["F%d" % r] = low, high, "мм"
        ws["P%d" % r], ws["R%d" % r] = p, allowed
        for k, pct in enumerate((0, 25, 50, 75, 100)):
            ws["G%d" % (r + k)] = pct
            ws["I%d" % (r + k)] = (high - low) * pct / 100.0 + low
        r += 5
    wb.save(path)


class MagicTest(unittest.TestCase):
    def test_all_good_and_share(self):
        rng = random.Random(1)
        chans = [(rng.choice((0, -50)), rng.choice((1, 6, 160, 600, 1600, 5000)), 0.2, 0.321)
                 for _ in range(200)]
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "src.xlsx")
            make_sheet(src, chans)
            out, written, st = magic.run(src, 30, rng=random.Random(7))
            self.assertEqual(st["channels"], 200)
            self.assertEqual(st["nonzero"], 60)
            chs = data.channels_from_rows(data.rows_from_xlsx(src, overrides=written))
            for c in chs:
                data.calc_channel(c)
            self.assertTrue(all(c.verdict == "Годен" for c in chs))
            self.assertEqual(sum(1 for c in chs if c.excel_err > 0), 60)
            # в новом файле J..M заполнены одинаково
            import openpyxl
            ws = openpyxl.load_workbook(out).active
            self.assertEqual(len({ws["%s5" % c].value for c in "JKLM"}), 1)

    def test_first_point_exact_and_one_sign(self):
        rng = random.Random(5)
        for _ in range(300):
            ch = {"low": 0.0, "high": 1600.0, "P": 0.2, "R": 0.321,
                  "points": [{"row": 4 + k, "xd": 400.0 * k, "empty": True} for k in range(5)]}
            v = magic.generate_channel(ch, True, rng)
            self.assertEqual(v[4], 0.0)
            signs = {(v[r] > ch["points"][r - 4]["xd"]) - (v[r] < ch["points"][r - 4]["xd"]) for r in v}
            self.assertLessEqual(len(signs - {0}), 1)
            self.assertGreater(magic.point_error(ch, v), 0)

    def test_random_each_run(self):
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "src.xlsx")
            make_sheet(src, [(0, 1600, 0.2, 0.321)] * 50)
            _, w1, _ = magic.run(src, 40)
            _, w2, _ = magic.run(src, 40)
            self.assertNotEqual(w1, w2)

    def test_shifted_layout_like_ges6(self):
        """ГЭС-6: нет столбца «№», всё сдвинуто на один влево."""
        import openpyxl
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "src.xlsx")
            make_sheet(src, [(-50, 200, 0.52, 0.612)] * 10)
            wb = openpyxl.load_workbook(src)
            wb.active.delete_cols(1)
            wb.save(src)
            _, written, st = magic.run(src, 50, rng=random.Random(3))
            self.assertEqual((st["channels"], st["nonzero"]), (10, 5))
            ws = openpyxl.load_workbook(_).active
            self.assertIsNotNone(ws["I5"].value)  # измерения легли в I..L
            self.assertIsNone(ws["M5"].value)
            chs = data.channels_from_rows(data.rows_from_xlsx(src, overrides=written))
            for c in chs:
                data.calc_channel(c)
            self.assertEqual([c.verdict for c in chs], ["Годен"] * 10)

    def test_filled_channels_untouched(self):
        import openpyxl
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "src.xlsx")
            make_sheet(src, [(0, 100, 0.2, 0.321)] * 3)
            wb = openpyxl.load_workbook(src)
            for r in range(4, 9):
                wb.active["J%d" % r] = 1.5
            wb.save(src)
            _, written, st = magic.run(src, 60)
            self.assertEqual(st["channels"], 2)
            self.assertFalse(any(r in written for r in range(4, 9)))


if __name__ == "__main__":
    unittest.main()
