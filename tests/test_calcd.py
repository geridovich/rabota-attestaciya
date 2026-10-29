"""Сверка с отчётом CalcD из папки input (файл «Ргаза»)."""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from attest.calcd import calc  # noqa: E402

VALUES = [
    0.004, 0.005, 0.006, 0.004, 0.004, 0.005, 0.006, 0.004,
    3.208, 3.207, 3.207, 3.206, 3.208, 3.207, 3.207, 3.206,
    6.409, 6.41, 6.409, 6.408, 6.409, 6.41, 6.409, 6.408,
    9.61, 9.611, 9.611, 9.612, 9.61, 9.611, 9.611, 9.612,
    12.812, 12.812, 12.813, 12.811, 12.812, 12.812, 12.813, 12.811,
    16.014, 16.014, 16.013, 16.015, 16.014, 16.014, 16.013, 16.015,
]
XD = [0.0, 3.2, 6.4, 9.6, 12.8, 16.0]


class CalcDTest(unittest.TestCase):
    def test_reference_report(self):
        r = calc(VALUES, XD, 8, 0)
        self.assertAlmostEqual(r.a1, 0.015467200230161902, places=15)
        self.assertAlmostEqual(r.a2, 0.012818514055551095, places=15)
        self.assertAlmostEqual(r.b1, 0.0005788409115429329, places=16)
        self.assertAlmostEqual(r.b2, 0.0005506233741712683, places=16)
        self.assertAlmostEqual(r.gh, 0.09667000143851188, places=14)
        self.assertAlmostEqual(r.gl, 0.08011571284719435, places=14)
        p = r.points[0]
        self.assertAlmostEqual(p.am, 0.00475, places=12)
        self.assertAlmostEqual(p.D, 1.6071428571e-7, places=16)
        self.assertAlmostEqual(p.upper, 0.0056981092, places=9)
        self.assertAlmostEqual(p.lower, 0.0038018907, places=9)
        self.assertAlmostEqual(r.points[3].dh, 0.021024072980974057, places=15)
        self.assertAlmostEqual(r.points[5].dl, 0.021628488042291387, places=15)
        # протокол: ПИП 0,1111 -> общая 0,14726947809414512
        self.assertAlmostEqual((0.1111 ** 2 + r.et_ik ** 2) ** 0.5, 0.14726947809414512, places=14)

    def test_constant_values(self):
        vals = []
        for v in (0, 0.26, 0.5, 0.75, 1.01):
            vals += [v] * 20
        r = calc(vals, [0, 0.25, 0.5, 0.75, 1.0], 20, 0)
        self.assertAlmostEqual(r.points[1].ds, 0.01, places=12)


if __name__ == "__main__":
    unittest.main()
