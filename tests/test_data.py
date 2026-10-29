"""Разбор вставки из Excel."""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from attest import data  # noqa: E402

# Жёлтые столбцы: №, KKS, наименование, диапазон, ед., %, действительное, измерения
PASTE_NEG = (
    "41\t01HNB02CT102\tТемпература металла подшипника №2 эл. Двиг. ДС\t-50\t200\t°С\t0\t-50\t-50\t-50\n"
    "\t\t\t\t\t\t25\t12,5\t12,5\t12,52\n"
    "\t\t\t\t\t\t50\t75\t75\t75,3\n"
    "\t\t\t\t\t\t75\t137,5\t137,5\t137,3\n"
    "\t\t\t\t\t\t100\t200\t200\t200,3\n"
)
# Строки A..M целиком (с входным сигналом 4..20 мА в H)
PASTE_AM = (
    "1\t08HHA01CP001\tДавление опрессовки\t0\t1\tкгс/см²\t0\t4\t0\t0\t\t\t\n"
    "\t\t\t\t\t\t25\t8\t0,25\t0,26\t\t\t\n"
    "\t\t\t\t\t\t50\t12\t0,5\t0,5\t\t\t\n"
    "\t\t\t\t\t\t75\t16\t0,75\t0,75\t\t\t\n"
    "\t\t\t\t\t\t100\t20\t1\t1,01\t\t\t\n"
)


class PasteTest(unittest.TestCase):
    def test_negative_range_several_measurements(self):
        chs = data.channels_from_rows(data.rows_from_text(PASTE_NEG))
        self.assertEqual(len(chs), 1)
        ch = chs[0]
        self.assertEqual((ch.kks, ch.low, ch.high), ("01HNB02CT102", -50.0, 200.0))
        self.assertEqual([p.xd for p in ch.points], [-50, 12.5, 75, 137.5, 200])
        self.assertEqual(ch.points[2].measured, [75, 75.3])
        data.calc_channel(ch)
        self.assertEqual(ch.error, "")
        self.assertIsNotNone(ch.result.gh)

    def test_input_signal_skipped(self):
        ch = data.channels_from_rows(data.rows_from_text(PASTE_AM))[0]
        self.assertEqual([p.xd for p in ch.points], [0, 0.25, 0.5, 0.75, 1])
        self.assertEqual([p.input for p in ch.points], [4, 8, 12, 16, 20])
        self.assertEqual(ch.points[1].measured, [0.26])

    def test_paste_without_number(self):
        text = "\n".join(ln.split("\t", 1)[1] for ln in PASTE_NEG.strip("\n").split("\n"))
        ch = data.channels_from_rows(data.rows_from_text(text))[0]
        self.assertEqual(ch.kks, "01HNB02CT102")
        self.assertEqual(len(ch.points), 5)


if __name__ == "__main__":
    unittest.main()
