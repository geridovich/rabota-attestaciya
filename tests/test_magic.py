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
    ws["J1"] = "Измеренное значение выходного сигнала"
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

    def test_random_each_run(self):
        with tempfile.TemporaryDirectory() as d:
            src = os.path.join(d, "src.xlsx")
            make_sheet(src, [(0, 1600, 0.2, 0.321)] * 50)
            _, w1, _ = magic.run(src, 40)
            _, w2, _ = magic.run(src, 40)
            self.assertNotEqual(w1, w2)

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
