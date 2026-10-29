"""Расчёт как в CalcD.jar (2012), перенесён 1:1 по байткоду.

Обозначения как в отчёте CalcD:
  x   - измерение (i-е значение в j-й точке)
  xd  - действительное значение точки
  am  - среднее арифметическое
  S   - СКО ряда
  U   - |am - x| / S (сравнивается с Umax из V1)
  am+ / am- - средние прямого / обратного хода (первая / вторая половина ряда)
  ds  - систематическая составляющая
  D   - дисперсия, СКО = sqrt(D)
  H   - вариация |am+ - am-|
  dh / dl - верхняя / нижняя граница погрешности по линейной регрессии
"""
import math
from dataclasses import dataclass, field

# vocs/V1.csv: n; h(0.05); h(0.01)  - критерий грубой погрешности
V1 = """4; 1,42 ;  1,46
5; 1,6  ;  1,67
6; 1,73 ;  1,82
7; 1,83 ;  1,94
8; 1,91 ;  2,03
9; 1,98 ;  2,11
10; 2,03 ;  2,18
11; 2,09 ;  2,23
12; 2,13 ;  2,29
13; 2,17 ;  2,33
14; 2,21 ;  2,37
15; 2,25 ;  2,41
16; 2,28 ;  2,44
17; 2,31 ;  2,48
18; 2,34 ;  2,5
19; 2,36 ;  2,53
20; 2,38 ;  2,56
25; 2,635;  2,87
30; 2,696;  2,928
40; 2,792;  3,015
50; 2,86 ;  3,082
100; 3,076;  3,285
250; 3,339;  3.534"""

# vocs/V2.csv: k; t(0.90); t(0.95)  - коэффициент Стьюдента
V2 = """4;2,132;2,776
5;2,015;2,571
6;1,943;2,447
7;1,895;2,365
8;1,86 ;2,306
9;1,833;2,262
10;1,812;2,228
11;1,796;2,201
12;1,782;2,179
13;1,771;2,16
14;1,761;2,145
15;1,753;2,131
16;1,746;2,12
17;1,74 ;2,11
18;1,734;2,101
19;1,729;2,093
20;1,725;2,086
21;1,721;2,08
22;1,717;2,074
23;1,714;2,069
24;1,711;2,064
25;1,708;2,06
26;1,706;2,056
27;1,703;2,052
28;1,701;2,048
29;1,699;2,045
30;1,697;2,042
40;1,68 ;2,02
оо;1,645;1,96"""

MODES = [
    "0 - по точкам подряд",
    "1 - прямой ход по всем точкам, затем обратный",
    "2 - парами по точкам",
    "3 - поочерёдно по точкам",
]


class CalcError(ValueError):
    pass


def _div(a, b):
    """Деление как в Java (без исключения при делении на 0)."""
    if b == 0:
        if a == 0 or math.isnan(a):
            return math.nan
        return math.copysign(math.inf, a) * (math.copysign(1, b))
    return a / b


def jsum(xs):
    """Сумма как в Java: последовательно. sum() в Python 3.12+ компенсирует
    ошибки округления и даёт другие последние знаки."""
    total = 0.0
    for x in xs:
        total += x
    return total


def get_from_v(table, n, col):
    """func.getfromV: строка, начинающаяся с n, иначе предыдущая перед первой большей."""
    prev = ""
    line = None
    for raw in table.splitlines():
        line = raw.strip()
        if line.startswith(str(n)):
            break
        try:
            key = int(line.split(";")[0])
        except ValueError:
            break  # строка "оо" - бесконечность
        if key > n:
            line = prev
            break
        prev = line
    if not line:
        raise CalcError(f"Нет значения в таблице для n={n} (нужно не меньше 4 измерений в точке)")
    return float(line.split(";")[col].strip().replace(",", "."))


@dataclass
class PointResult:
    xd: float
    am: float
    S: float
    am_plus: float
    am_minus: float
    ds: float
    D: float
    sko: float
    H: float
    upper: float
    lower: float
    dh: float = 0.0
    dl: float = 0.0
    rows: list = field(default_factory=list)  # (x, xc, x/am, (x-am)/am, U)


@dataclass
class CalcResult:
    points: list
    a1: float
    a2: float
    b1: float
    b2: float
    gh: float
    gl: float
    umax: float
    kstd: float

    @property
    def et_ik(self):
        """Погрешность ЭТ ИК для протокола CalcD: max(|gh|, |gl|)."""
        return max(abs(self.gh), abs(self.gl))


def split_by_mode(values, qpoints, qm, mode):
    """Frame2.input_to_array: раскладка ряда значений по точкам."""
    if len(values) != qpoints * qm:
        raise CalcError(f"Ожидалось {qpoints * qm} значений, получено {len(values)}")
    pmx = [[0.0] * qm for _ in range(qpoints)]
    it = iter(values)
    half = qm // 2
    if mode == 0:
        for j in range(qpoints):
            for i in range(qm):
                pmx[j][i] = next(it)
    elif mode == 1:
        for j in range(qpoints):
            for i in range(half):
                pmx[j][i] = next(it)
        for j in range(qpoints):
            for i in range(half, qm):
                pmx[j][i] = next(it)
    elif mode == 2:
        for i in range(half):
            for j in range(qpoints):
                pmx[j][i * 2] = next(it)
                pmx[j][i * 2 + 1] = next(it)
    elif mode == 3:
        for i in range(qm):
            for j in range(qpoints):
                pmx[j][i] = next(it)
    else:
        raise CalcError(f"Неизвестный режим последовательности {mode}")
    return pmx


def calc(values, xd, qm, mode=0):
    """Полный расчёт CalcD. values - все измерения подряд, xd - действительные значения точек."""
    qpoints = len(xd)
    if qpoints < 2:
        raise CalcError("Нужно минимум 2 точки")
    if qm % 4:
        raise CalcError("Количество измерений в точке должно быть кратно 4")
    pmx = split_by_mode(values, qpoints, qm, mode)
    umax = get_from_v(V1, qm, 2)
    kstd = get_from_v(V2, qm - 1, 2)

    points = []
    for j in range(qpoints):
        xs = pmx[j]
        xc = list(xs)

        def calc_dj():
            am = jsum(xs) / qm
            s2 = 0.0
            for i in range(qm):
                s2 += (xs[i] - am) ** 2
                xc[i] = xs[i]
            S = math.sqrt(s2 / (qm - 1))
            u = [_div(abs(x - am), S) for x in xs]
            return am, S, u

        am, S, u = calc_dj()
        # changeUs: замена грубых промахов. В CalcD повторный CalcDj сразу
        # возвращает xc = x, поэтому замена на результат не влияет.
        for i in range(qm):
            if u[i] > umax:
                xc[i] = am if i in (0, qm - 1) else (xs[i - 1] + xs[i + 1]) / 2
        am, S, u = calc_dj()

        x0 = xd[j]
        half = qm // 2
        am_p = jsum(xc[:half]) * 2.0 / qm
        am_m = jsum(xc[half:]) * 2.0 / qm
        ds = ((am_p - x0) + (am_m - x0)) / 2.0
        # В CalcD сумма квадратов не накапливается: берутся только последние
        # элементы каждой половины. Повторяем как есть.
        d = ((xc[half - 1] - am_p) ** 2 + (xc[qm - 1] - am_m) ** 2) / (qm - 1)
        sko = math.sqrt(d)
        H = abs(am_p - am_m)
        w = kstd * math.sqrt(sko ** 2 + H ** 2 / 12.0)
        rows = [(xs[i], xc[i], _div(xs[i], am), _div(xs[i] - am, am), u[i]) for i in range(qm)]
        points.append(PointResult(x0, am, S, am_p, am_m, ds, d, sko, H, ds + w, ds - w, rows=rows))

    def regress(ys):
        sx = sy = sxy = sx2 = 0.0
        for p, y in zip(points, ys):
            sx += p.xd
            sy += y
            sxy += y * p.xd
            sx2 += p.xd ** 2
        sxy *= qpoints
        sx2 *= qpoints
        b = _div(sx * sy - sxy, sx ** 2 - sx2)
        return (sy + b * sx) / qpoints, b

    a1, b1 = regress([p.upper for p in points])
    a2, b2 = regress([p.lower for p in points])
    xmax = max(0.0, *[p.xd for p in points])
    xmin = min(0.0, *[p.xd for p in points])
    rng = xmax - xmin
    gh = _div(a1 * 100.0, rng)
    gl = _div(a2 * 100.0, rng)
    for p in points:
        p.dh = gh * rng / 100.0 + b1 * p.xd
        p.dl = gl * rng / 100.0 + b2 * p.xd
    return CalcResult(points, a1, a2, b1, b2, gh, gl, umax, kstd)


def fmt(v, digits=None):
    """Число с десятичной запятой. digits=None - все знаки, как в CalcD."""
    if v is None:
        return ""
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return "-"
    if digits is None:
        if float(v).is_integer():
            return str(int(v))
        s = repr(float(v))
        if "e" in s or "E" in s:
            s = f"{v:.20f}".rstrip("0").rstrip(".")
    else:
        s = f"{v:.{digits}f}"
    if s in ("-0", "-0.0") or (s.startswith("-") and float(s) == 0):
        s = s[1:]
    return s.replace(".", ",")


def report(res, digits=None):
    """Текстовый отчёт в духе окна «Результаты расчёта» CalcD."""
    f = lambda v: fmt(v, digits)
    lines = ["РАСЧЕТ ТОЧЕК", "",
             "xd; am; S; am+; am-; ds; D; СКО; H; Граница верх; Граница низ; delta H; delta L"]
    for p in res.points:
        lines.append("; ".join(f(v) for v in (p.xd, p.am, p.S, p.am_plus, p.am_minus, p.ds, p.D,
                                               p.sko, p.H, p.upper, p.lower, p.dh, p.dl)))
    lines += ["",
              f"a1={f(res.a1)}", f"a2={f(res.a2)}", f"b1={f(res.b1)}", f"b2={f(res.b2)}", "",
              f"gh={f(res.gh)}", f"gl={f(res.gl)}"]
    return "\n".join(lines)
