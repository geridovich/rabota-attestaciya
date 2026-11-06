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

# Ровно жёлтые столбцы листа (скрин 3): A..F, I, J - без «%» и входного сигнала;
# объединённые ячейки A..F копируются только в первую строку канала
PASTE_YELLOW = (
    "1\t08HHA01CP001\tДавление опрессовки гор. №IА\t0\t1\tкгс/см²\t0\t0\n"
    "\t\t\t\t\t\t0,25\t0,26\n"
    "\t\t\t\t\t\t0,5\t0,5\n"
    "\t\t\t\t\t\t0,75\t0,75\n"
    "\t\t\t\t\t\t1\t1,01\n"
    "3\t08HHL01DP001\tДавление воздуха к  гор. №IА\t0\t160\tкгс/м²\t0\t0\n"
    "\t\t\t\t\t\t40\t40\n"
    "\t\t\t\t\t\t80\t80\n"
    "\t\t\t\t\t\t120\t120\n"
    "\t\t\t\t\t\t160\t160\n"
)


class PasteTest(unittest.TestCase):
    def test_yellow_columns_without_percent(self):
        chs = data.channels_from_rows(data.rows_from_text(PASTE_YELLOW))
        self.assertEqual([c.kks for c in chs], ["08HHA01CP001", "08HHL01DP001"])
        self.assertEqual([p.xd for p in chs[0].points], [0, 0.25, 0.5, 0.75, 1])
        self.assertEqual([p.measured for p in chs[0].points], [[0], [0.26], [0.5], [0.75], [1.01]])
        self.assertEqual([p.pct for p in chs[1].points], [0, 25, 50, 75, 100])

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

    def test_protocols_one_document(self):
        from attest import docgen
        chs = data.channels_from_rows(data.rows_from_text(PASTE_YELLOW))
        for c in chs:
            data.calc_channel(c)
        doc = docgen.build_protocols(chs, [(1, 2, "07.07.2026"), (3, 3, "08.07.2026")],
                                     "САУГ КА-4 ТЭЦ-17", docgen.EXECUTORS[3])
        heads = [p.text for p in doc.paragraphs if p.text.startswith("Протокол №")]
        self.assertEqual(heads, ["Протокол № 1 от 07.07.2026", "Протокол № 3 от 08.07.2026"])
        cells = [c.text for t in doc.tables for r in t.rows for c in r.cells]
        self.assertNotIn("{{", "".join(cells))
        self.assertIn("Геридович Н.В.", cells)
        # 5 точек x (Меньшее + Большее) + 3 строки шапки
        self.assertEqual(len(doc.tables[2].rows), 13)
        self.assertEqual(doc.tables[2].rows[5].cells[4].text, "0,26")

    def test_perechen_and_calibration_tables(self):
        from attest import docgen
        chs = data.channels_from_rows(data.rows_from_text(PASTE_YELLOW + PASTE_NEG))
        doc = docgen.build_perechen(chs, "ПТК САУГ КА-4 Выборгской ТЭЦ (ТЭЦ-17)")
        text = "\n".join(p.text for p in doc.paragraphs)
        self.assertIn("Перечень измерительных каналов ПТК САУГ КА-4 Выборгской ТЭЦ (ТЭЦ-17)", text)
        self.assertIn("Таблица 1 – градуировочная таблица ИК сигналов в диапазоне 4-20 мА", text)
        self.assertIn("Таблица 2 – градуировочная таблица ИК сигналов термосопротивления", text)
        # перечень: 3 строки шапки + по строке на канал; границы и ед. изм. отдельно
        self.assertEqual(len(doc.tables[0].rows), 3 + 3)
        self.assertEqual([c.text for c in doc.tables[0].rows[5].cells][:6],
                         ["41", "01HNB02CT102", "Температура металла подшипника №2 эл. Двиг. ДС", "-50", "200", "°С"])
        # 4-20 мА считаются по % диапазона, даже если в Excel нет входного сигнала
        ma = doc.tables[1]
        self.assertEqual([[c.text for c in r.cells][-1] for r in ma.rows[2:7]], ["4", "8", "12", "16", "20"])
        # градуировочная: -50...200 °С, медь 50М: -50 -> 39,35 Ом
        cal = doc.tables[2]
        self.assertEqual([c.text for c in cal.rows[2].cells][-1], "39,35")
        self.assertNotIn("{{", "".join(c.text for t in doc.tables for r in t.rows for c in r.cells))

    def test_mixed_layouts_in_one_paste(self):
        chs = data.channels_from_rows(data.rows_from_text(PASTE_YELLOW + PASTE_NEG))
        self.assertEqual([p.xd for p in chs[2].points], [-50, 12.5, 75, 137.5, 200])
        self.assertEqual([p.xd for p in chs[0].points], [0, 0.25, 0.5, 0.75, 1])

    def test_cell_with_line_breaks(self):
        """Ячейка с Alt+Enter приходит из Excel в кавычках - канал не должен рваться."""
        text = ('281\t26MAD02CT126\t"Температура колодок\nопорно-упорного\nподшипника (сторона\n'
                'генератора)"\t-50\t200\t˚C\t0\t-50\t-50\t-50\n'
                '\t\t\t\t\t\t25\t12,5\t12,5\t12,5\n'
                '\t\t\t\t\t\t50\t75\t75\t75\n'
                '\t\t\t\t\t\t75\t137,5\t137,5\t137,5\n'
                '\t\t\t\t\t\t100\t200\t200\t200\n'
                '282\t26MAJ10CT001\tДатчик "Б" и ""кавычки""\t-50\t350\t˚C\t0\t-50\t-50\t-50\n')
        chs = data.channels_from_rows(data.rows_from_text(text))
        self.assertEqual([c.kks for c in chs], ["26MAD02CT126", "26MAJ10CT001"])
        self.assertEqual(chs[0].name, "Температура колодок опорно-упорного подшипника (сторона генератора)")
        self.assertEqual(len(chs[0].points), 5)
        self.assertEqual(chs[1].name, 'Датчик "Б" и ""кавычки""')

    def test_kks_with_space_vs_note(self):
        """«AA00TU03 LCA30CP001» с № - канал (ТГ-36); «НЕТ НА МНЕМОСХЕМЕ» без № - заметка."""
        self.assertTrue(data.is_kks("AA00TU03 LCA30CP001", 348, "Давление конд."))
        self.assertFalse(data.is_kks("НЕТ НА МНЕМОСХЕМЕ", None, None))
        self.assertTrue(data.is_kks("10HBK50CQ004"))

    def test_allowed_from_formula_without_cache(self):
        """R = 1,1·√(P²+W²) формулой без сохранённого значения - считаем сами."""
        c = data.Channel(num="1", kks="K", name="N", low=0, high=1, unit="", pip_ik=0.5, module_err=0.15)
        self.assertAlmostEqual(c.allowed_value, 1.1 * (0.25 + 0.0225) ** 0.5)

    def test_paste_without_number(self):
        text = "\n".join(ln.split("\t", 1)[1] for ln in PASTE_NEG.strip("\n").split("\n"))
        ch = data.channels_from_rows(data.rows_from_text(text))[0]
        self.assertEqual(ch.kks, "01HNB02CT102")
        self.assertEqual(len(ch.points), 5)


def sheet_channel(m1="0", allowed=1.0, low=0, high=100):
    """Канал из строк листа: 3 точки, допуск в R."""
    rows = [{"A": 1, "B": "10HBK50CQ004", "C": "Давление", "D": low, "E": high, "F": "кПа",
             "I": 0, "J": m1, "K": 0, "R": allowed},
            {"I": 50, "J": 50, "K": 50},
            {"I": 100, "J": 100, "K": 100}]
    ch = data.channels_from_rows(rows)[0]
    return data.calc_channel(ch, 20)


class BadValues(unittest.TestCase):
    """Отчёт тестирования PU-01, PU-05: не-числа не должны давать «Годен»."""

    def test_good_channel(self):
        ch = sheet_channel()
        self.assertEqual(ch.error, "")
        self.assertEqual(ch.verdict, "Годен")

    def test_nan_measurement(self):
        for v in ("NaN", float("nan"), "inf", "ОШИБКА", "#ДЕЛ/0!"):
            ch = sheet_channel(m1=v)
            self.assertIn("не число", ch.error, v)
            self.assertEqual(ch.verdict, "", v)

    def test_nan_allowed(self):
        ch = sheet_channel(allowed="NaN")
        self.assertIn("допустимая погрешность", ch.error)
        self.assertEqual(ch.verdict, "")

    def test_zero_range(self):
        ch = sheet_channel(low=0, high=0)
        self.assertIn("нулевой диапазон", ch.error)
        self.assertIn("на 0", ch.problems()[0].why)

    def test_blank_lower_bound(self):
        ch = sheet_channel(low=None)
        self.assertIn("нижняя граница", ch.problems()[0].why)
        self.assertIsNone(ch.result)

    def test_bad_allowed_fallback_x(self):
        rows = [{"A": 1, "B": "10HBK50CQ004", "C": "Давление", "D": 0, "E": 100, "I": 0, "J": 0, "X": "#VALUE!"},
                {"I": 50, "J": 50}, {"I": 100, "J": 100}]
        ch = data.calc_channel(data.channels_from_rows(rows)[0], 20)
        self.assertIn("не число", ch.error)

    def test_same_xd_everywhere(self):
        rows = [{"A": 1, "B": "10HBK50CQ004", "C": "Давление", "D": 0, "E": 100, "I": 50, "J": 50, "R": 1},
                {"I": 50, "J": 50}, {"I": 50, "J": 50}]
        ch = data.calc_channel(data.channels_from_rows(rows)[0], 20)
        self.assertIn("одно действительное", ch.error)
        self.assertEqual(ch.verdict, "")

    def test_paste_0_100_keeps_first_measurement(self):
        """Диапазон 0..100 без столбца «%»: первое измерение не должно уйти во «вход»."""
        text = ("1\t10HBK50CQ004\tУровень\t0\t100\t%\t0\t4\t0\t0\t0\n"
                "\t\t\t\t\t\t50\t54\t50\t50\t50\n"
                "\t\t\t\t\t\t100\t104\t100\t100\t100\n")
        ch = data.channels_from_rows(data.rows_from_text(text))[0]
        self.assertEqual(ch.points[0].measured, [4, 0, 0, 0])
        self.assertAlmostEqual(data.calc_channel(ch, 20).et_ik, 1.0)

    def test_paste_error_cell_is_reported(self):
        text = ("1\t10HBK50CQ004\tУровень\t0\t10\tм\t0\t#DIV/0!\t0\t0\t0\n"
                "\t\t\t\t\t\t5\t#DIV/0!\t5\t5\t5\n")
        ch = data.calc_channel(data.channels_from_rows(data.rows_from_text(text))[0], 20)
        self.assertIn("не число", ch.error)

    def test_problem_points_to_cell(self):
        ch = sheet_channel(allowed="#VALUE!")
        p = ch.problems()[0]
        self.assertIn("ошибка формулы", p.why.lower())
        self.assertEqual(p.value, "#VALUE!")


if __name__ == "__main__":
    unittest.main()
