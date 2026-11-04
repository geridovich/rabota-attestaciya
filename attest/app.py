"""Plus Ultra: каналы из Excel -> расчёт коэффициентов -> протоколы и перечень в Word."""
import datetime
import json
import os
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import calcd, data, dnd, docgen, heal, magic, themes
from .music import Sound, asset_path

APP_NAME = "Plus Ultra"
TITLE = "GERA | ATTESTATION SYSTEM"  # по центру заголовка окна
APP_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else sys.argv[0]))
SETTINGS = os.path.join(APP_DIR, "settings.json")
MAX_BATCH = 3000  # протоколов в одном документе
DATE_RANGES = 15  # строк «с № по № - дата»
HEALED_DIR = "Вылеченные файлы Excel"
MAGIC_DIR = "Заполненные измерения (magic)"
DATE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
QM_MIN, QM_MAX = 8, 200  # таблица Стьюдента CalcD начинается с n-1 = 4, кратно 4 -> не меньше 8
DIGITS_MAX = 15
DEFAULTS = {
    "pip": "", "allowed": "", "qm": 20, "digits": 4, "batch": 20,
    "out_dir": os.path.join(APP_DIR, "Документы"), "theme": themes.DEFAULT,
    "station": "САУГ КА-4 ТЭЦ-17", "perechen": "ПТК САУГ КА-4 Выборгской ТЭЦ (ТЭЦ-17)",
    "executor": 0, "dates": [["", "", ""] for _ in range(DATE_RANGES)],
    "click_sound": True, "magic_sound": True,
}

COLS = (("num", "№", 35), ("kks", "KKS", 100), ("name", "Наименование", 160), ("range", "Диапазон", 80),
        ("gh", "gh", 55), ("gl", "gl", 55), ("et", "ЭТ ИК, %", 60), ("total", "Общая, %", 60),
        ("allowed", "Допуст., %", 60), ("verdict", "Результат", 65))


def load_settings():
    s = dict(DEFAULTS)
    try:
        with open(SETTINGS, encoding="utf-8") as f:
            s.update(json.load(f))
    except (OSError, ValueError):
        pass
    dates = list(s.get("dates") or [])[:DATE_RANGES]
    s["dates"] = dates + [["", "", ""]] * (DATE_RANGES - len(dates))
    return s


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.scale = max(1.0, self.winfo_fpixels("1i") / 96.0)  # масштаб Windows (125 %, 150 %...)
        self.settings = load_settings()
        self.channels = []
        self.click_sound = Sound("click.mp3", "plusultra_click", bool(self.settings["click_sound"]))
        self.magic_sound = Sound("magic.mp3", "plusultra_magic", bool(self.settings["magic_sound"]))
        self.xlsx_path = None  # Excel, открытый кнопкой «Открыть Excel…» (нужен для magic)
        self._set_icon()
        self.texts = []  # tk.Text - их цвета задаются вручную

        self.style = ttk.Style(self)
        self._fit_to_screen()
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        self._build_main_tab(nb)
        self._build_settings_tab(nb)
        self._build_help_tab(nb)
        self.nb = nb
        self.apply_theme(self.settings["theme"])
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._title_width = 0
        self.bind("<Configure>", self._center_title, add="+")
        # щелчок на все кнопки программы (и в окнах-диалогах ttk)
        self.bind_class("TButton", "<Button-1>", lambda _e: self.click_sound.play(), add="+")
        try:
            dnd.enable(self, self.on_drop)  # перетаскивание Excel в окно
        except (OSError, AttributeError, ValueError):
            pass

    def _set_icon(self):
        """Значок окна и панели задач."""
        try:
            self.iconbitmap(default=asset_path("window.ico"))
        except tk.TclError:
            pass
        try:
            self._icon_img = tk.PhotoImage(file=asset_path("window.png"))
            self.iconphoto(True, self._icon_img)
        except tk.TclError:
            pass

    def _center_title(self, _e=None):
        """«GERA | ATTESTATION SYSTEM» по центру заголовка окна.

        Windows пишет заголовок от левого края, поэтому после «Plus Ultra»
        ставятся пробелы - столько, чтобы вторая надпись встала в центр.
        """
        w = self.winfo_width()
        if w == self._title_width or w < 50:
            return
        self._title_width = w
        import tkinter.font as tkfont
        font = tkfont.Font(family="Segoe UI", size=9)  # шрифт заголовка Windows
        left = self.px(36)  # значок окна и отступ до текста
        gap = (w - font.measure(TITLE)) / 2 - left - font.measure(APP_NAME)
        n = max(3, int(gap / max(1, font.measure(" "))))
        self.title(APP_NAME + " " * n + TITLE)

    def px(self, n):
        return int(n * self.scale)

    def _fit_to_screen(self):
        """Окно по размеру экрана: на ноутбуке 1366x768 открывается развёрнутым."""
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(self.px(1320), sw - 20), min(self.px(840), sh - 80)
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")
        self.minsize(min(self.px(900), sw), min(self.px(560), sh - 80))
        if w >= sw - 20 or h >= sh - 80:
            try:
                self.state("zoomed")
            except tk.TclError:
                pass

    # ---------- главная ----------
    def _build_main_tab(self, nb):
        tab = ttk.Frame(nb, padding=6)
        nb.add(tab, text="Главная")

        bar = ttk.Frame(tab)
        bar.pack(fill="x")
        ttk.Button(bar, text="Вставить из буфера", command=self.paste_clipboard).pack(side="left")
        ttk.Button(bar, text="Открыть Excel…", command=self.open_xlsx).pack(side="left", padx=(4, 0))
        ttk.Button(bar, text="Полечить файл", command=self.heal_file).pack(side="left", padx=4)
        ttk.Button(bar, text="Очистить", command=self.clear).pack(side="left")
        ttk.Button(bar, text="Рассчитать", command=self.recalc).pack(side="left", padx=4)
        ttk.Button(bar, text="Копировать результаты", command=self.copy_results).pack(side="left")
        ttk.Button(bar, text="Ошибки…", command=self.show_errors).pack(side="left", padx=4)
        self.click_btn = ttk.Button(bar, width=8, command=self.toggle_click_sound)
        self.click_btn.pack(side="right")
        self._click_btn_text()
        self.status = ttk.Label(bar, text="")
        self.status.pack(side="right", padx=8)

        body = ttk.PanedWindow(tab, orient="horizontal")
        body.pack(fill="both", expand=True, pady=(6, 0))

        left = ttk.PanedWindow(body, orient="vertical")
        top = ttk.Frame(left)
        self.tree = ttk.Treeview(top, columns=[c[0] for c in COLS], show="headings", selectmode="extended")
        for key, title, w in COLS:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=self.px(w), minwidth=self.px(30),
                             anchor="w" if key in ("kks", "name") else "center")
        ys = ttk.Scrollbar(top, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ys.set)
        self.tree.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.show_detail)
        left.add(top, weight=3)
        bottom = ttk.Frame(left)
        self.detail = tk.Text(bottom, width=40, height=4, font=("Courier New", 9), wrap="none")
        self.texts.append(self.detail)
        dx = ttk.Scrollbar(bottom, orient="horizontal", command=self.detail.xview)
        self.detail.configure(xscrollcommand=dx.set)
        self._build_magic(bottom)
        self.detail.pack(fill="both", expand=True)
        dx.pack(fill="x")
        left.add(bottom, weight=1)
        body.add(left, weight=3)

        body.add(self._build_side_panel(body), weight=0)

    def _build_side_panel(self, parent):
        side = ttk.Frame(parent, padding=(8, 0, 0, 0))
        s = self.settings

        doc = ttk.LabelFrame(side, text="Документы", padding=6)
        doc.pack(fill="x")
        self.station = tk.StringVar(value=s["station"])
        self.perechen = tk.StringVar(value=s["perechen"])
        self.executor = ttk.Combobox(doc, state="readonly", width=34,
                                     values=["%d. %s %s" % (i + 1, d, f) for i, (d, f) in enumerate(docgen.EXECUTORS)])
        self.executor.current(min(max(0, int(s["executor"])), len(docgen.EXECUTORS) - 1))
        for r, (label, widget) in enumerate((
                ("Аттестация ИК:", ttk.Entry(doc, textvariable=self.station, width=36)),
                ("Перечень ИК:", ttk.Entry(doc, textvariable=self.perechen, width=36)),
                ("Исполнитель:", self.executor))):
            ttk.Label(doc, text=label).grid(row=r, column=0, sticky="w", padx=(0, 6), pady=1)
            widget.grid(row=r, column=1, sticky="we", pady=1)
        doc.columnconfigure(1, weight=1)

        dates = ttk.LabelFrame(side, text="Даты протоколов (дата в виде 07.07.2026)", padding=6)
        dates.pack(fill="x", pady=(6, 0))
        for col in (0, 4):
            for k, h in enumerate(("с №", "по №", "дата")):
                ttk.Label(dates, text=h, style="Hint.TLabel").grid(row=0, column=col + k)
        self.date_vars = []
        for i in range(DATE_RANGES):
            col = 0 if i < 8 else 4
            row = (i if i < 8 else i - 8) + 1
            lo, hi, dt = (tk.StringVar(value=v) for v in s["dates"][i])
            ttk.Entry(dates, textvariable=lo, width=5).grid(row=row, column=col, padx=(0 if col == 0 else 12, 1), pady=0)
            ttk.Entry(dates, textvariable=hi, width=5).grid(row=row, column=col + 1, padx=1)
            ttk.Entry(dates, textvariable=dt, width=10).grid(row=row, column=col + 2, padx=1)
            self.date_vars.append((lo, hi, dt))

        gen = ttk.LabelFrame(side, text="Создать", padding=6)
        gen.pack(fill="x", pady=(6, 0))
        row = ttk.Frame(gen)
        row.pack(fill="x")
        ttk.Label(row, text="Протоколов в документе (1–3000):").pack(side="left")
        self.batch = tk.IntVar(value=s["batch"])
        ttk.Spinbox(row, from_=1, to=MAX_BATCH, width=6, textvariable=self.batch).pack(side="right")
        row = ttk.Frame(gen)
        row.pack(fill="x", pady=(2, 0))
        ttk.Label(row, text="начиная со строки списка:").pack(side="left")
        self.start = tk.IntVar(value=1)
        ttk.Spinbox(row, from_=1, to=100000, width=6, textvariable=self.start).pack(side="right")
        btns = ttk.Frame(gen)
        btns.pack(fill="x", pady=(6, 0))
        ttk.Button(btns, text="Протоколы", command=self.generate).grid(row=0, column=0, sticky="we")
        ttk.Button(btns, text="Только выделенные", command=self.generate_selected).grid(row=0, column=1, sticky="we", padx=(4, 0))
        ttk.Button(btns, text="Перечень", command=self.generate_perechen).grid(row=1, column=0, sticky="we", pady=(4, 0))
        ttk.Button(btns, text="Открыть папку", command=self.open_out_dir).grid(row=1, column=1, sticky="we", padx=(4, 0), pady=(4, 0))
        btns.columnconfigure(0, weight=1)
        btns.columnconfigure(1, weight=1)
        return side

    def _build_magic(self, parent):
        """Круглая чёрная кнопка magic и ползунок «% каналов с погрешностью»."""
        box = ttk.Frame(parent, padding=(8, 2, 2, 2))
        box.pack(side="right", fill="y")
        d = self.px(150)
        self.magic_btn = tk.Canvas(box, width=d + 4, height=d + 4, highlightthickness=0, cursor="hand2")
        self.magic_btn.create_oval(2, 2, d + 2, d + 2, fill="#000000", outline="#444444", width=3, tags="disc")
        self.magic_btn.create_text(d // 2 + 2, d // 2 + 2, text="magic", fill="#ffffff",
                                   font=("Segoe UI", 24, "bold italic"), tags="label")
        self.magic_btn.bind("<Button-1>", lambda _e: self.run_magic())
        self.magic_btn.bind("<Enter>", lambda _e: self.magic_btn.itemconfig("disc", outline="#d4af37"))
        self.magic_btn.bind("<Leave>", lambda _e: self.magic_btn.itemconfig("disc", outline="#444444"))
        self.magic_btn.pack()
        self.magic_pct = tk.IntVar(value=int(self.settings.get("magic_pct", 20)))
        self.magic_label = ttk.Label(box, text="", style="Hint.TLabel")
        scale = ttk.Scale(box, from_=5, to=60, orient="horizontal", length=self.px(170),
                          command=lambda v: self._magic_pct_changed(v))
        scale.set(self.magic_pct.get())
        scale.pack(pady=(4, 0))
        self.magic_label.pack()
        self._magic_pct_changed(self.magic_pct.get())

    def _magic_pct_changed(self, v):
        self.magic_pct.set(int(round(float(v))))
        self.magic_label.config(text="с погрешностью: %d %%" % self.magic_pct.get())

    def heal_file(self):
        """Привести открытый Excel к одному виду (копия «… (вылечено).xlsx») и работать с ней."""
        if not self.xlsx_path or not os.path.exists(self.xlsx_path):
            messagebox.showwarning("Полечить файл", "Сначала откройте файл Excel кнопкой «Открыть Excel…».")
            return
        if getattr(self, "_busy", False):
            return
        self._busy = True
        self.config(cursor="watch")
        self.status.config(text="Лечу файл…")
        path, box, out_dir = self.xlsx_path, {}, self.sub_dir(HEALED_DIR)

        def work():  # в потоке - без обращений к Tk
            try:
                box["result"] = heal.heal(path, out_dir=out_dir)
            except Exception as e:  # noqa: BLE001
                box["error"] = e

        threading.Thread(target=work, daemon=True).start()
        self._wait(box, self._heal_done)

    def _wait(self, box, done):
        if not box:
            self.after(200, self._wait, box, done)
            return
        self._busy = False
        self.config(cursor="")
        done(box)

    def _heal_done(self, box):
        if "error" in box:
            self.refresh()
            messagebox.showerror("Полечить файл", "Не получилось: %s" % box["error"])
            return
        out, fixes, n = box["result"]
        self.xlsx_path = out
        self.channels = []
        self.add_channels(data.channels_from_rows(data.rows_from_xlsx(out)))
        lines = ["• %s: %d" % (k, v) for k, v in fixes.items()] or ["• косяков не нашлось, файл и так в порядке"]
        text = "Каналов: %d. Что исправлено:\n%s\n\nДальше работа идёт с копией:\n%s\n\nОткрыть её?" % (
            n, "\n".join(lines), out)
        if messagebox.askyesno("Полечить файл", text):
            os.startfile(out)

    def run_magic(self):
        if not self.xlsx_path or not os.path.exists(self.xlsx_path):
            messagebox.showwarning("Magic", "Сначала откройте файл Excel кнопкой «Открыть Excel…».")
            return
        if getattr(self, "_magic_busy", False):
            return
        self.magic_sound.play()
        self._magic_busy = True
        self.config(cursor="watch")
        self.status.config(text="Magic: заполняю измеренные значения…")
        path, pct, box = self.xlsx_path, self.magic_pct.get(), {}
        out_dir = self.sub_dir(MAGIC_DIR)  # в потоке Tk трогать нельзя

        def work():
            try:
                box["result"] = magic.run(path, pct, out_dir=out_dir)
            except Exception as e:  # noqa: BLE001
                box["error"] = e

        threading.Thread(target=work, daemon=True).start()
        self._magic_wait(box, path)

    def _magic_wait(self, box, path):
        if not box:
            self.after(200, self._magic_wait, box, path)
            return
        self._magic_busy = False
        self.config(cursor="")
        if self.xlsx_path != path:
            return  # пока шла обработка, данные очистили или открыли другой файл
        if "error" in box:
            self.refresh()
            messagebox.showerror("Magic", "Не получилось: %s" % box["error"])
            return
        out, written, st = box["result"]
        if not st["channels"]:
            self.refresh()
            messagebox.showinfo("Magic", "Пустых «Измеренных значений» на листе «%s» нет — всё уже заполнено."
                                % st["sheet"])
            return
        # каналы в программе - с новыми значениями, протоколы можно делать сразу
        self.channels = data.channels_from_rows(data.rows_from_xlsx(path, overrides=written))
        self.start.set(1)
        self.batch.set(min(len(self.channels), MAX_BATCH))
        self.recalc()
        if messagebox.askyesno(
                "Magic", "Заполнено каналов: %d (с погрешностью: %d) на листе «%s».\n%s\n\nОткрыть файл?"
                % (st["channels"], st["nonzero"], st["sheet"], out)):
            os.startfile(out)

    # ---------- настройки ----------
    def _build_settings_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)
        nb.add(tab, text="Настройки")
        s = self.settings
        r = 0

        def row(label, widget):
            nonlocal r
            ttk.Label(tab, text=label).grid(row=r, column=0, sticky="w", pady=4, padx=(0, 10))
            widget.grid(row=r, column=1, sticky="w", pady=4)
            r += 1

        self.theme = ttk.Combobox(tab, state="readonly", values=list(themes.PALETTES), width=18)
        self.theme.set(s["theme"] if s["theme"] in themes.PALETTES else themes.DEFAULT)
        self.theme.bind("<<ComboboxSelected>>", lambda _e: self.apply_theme(self.theme.get()))
        row("Цветовая палитра:", self.theme)
        self.qm = tk.IntVar(value=s["qm"])
        row("Измерений в точке:", ttk.Spinbox(tab, from_=QM_MIN, to=QM_MAX, increment=4, width=6,
                                              textvariable=self.qm))
        self.digits = tk.IntVar(value=s["digits"])
        row("Знаков после запятой в таблице:", ttk.Spinbox(tab, from_=0, to=15, width=6, textvariable=self.digits))
        self.pip_default = tk.StringVar(value=s["pip"])
        row("ПИП, % (если нет в данных):", ttk.Entry(tab, width=8, textvariable=self.pip_default))
        self.allowed_default = tk.StringVar(value=s["allowed"])
        row("Допустимая, % (если нет в данных):", ttk.Entry(tab, width=8, textvariable=self.allowed_default))
        self.magic_sound_on = tk.BooleanVar(value=self.magic_sound.enabled)
        row("Звук кнопки magic:", ttk.Checkbutton(
            tab, text="включён", variable=self.magic_sound_on,
            command=lambda: setattr(self.magic_sound, "enabled", self.magic_sound_on.get())))
        out = ttk.Frame(tab)
        self.out_dir = tk.StringVar(value=s["out_dir"])
        ttk.Entry(out, textvariable=self.out_dir, width=50).pack(side="left")
        ttk.Button(out, text="…", width=3, command=self.pick_out_dir).pack(side="left", padx=4)
        row("Папка для документов:", out)
        ttk.Button(tab, text="Пересчитать", command=self.recalc).grid(row=r, column=1, sticky="w", pady=(10, 0))

    # ---------- помощь ----------
    HELP_STEPS = (
        ("1", "Открыть Excel", "Нажмите «Открыть Excel…» или просто перетащите файл Excel мышью в окно. "
                               "Можно и скопировать строки каналов в Excel (Ctrl+C) и нажать «Вставить из буфера»."),
        ("2", "Полечить файл", "Если файл «кривой» (сдвинутые столбцы, заметки в KKS, пустые заготовки, "
                               "числа текстом), нажмите «Полечить файл». Программа сделает исправленную копию "
                               "в папке «Вылеченные файлы Excel» и дальше будет работать с ней."),
        ("3", "Magic и ползунок", "Если в Excel пусты «Измеренные значения выходного сигнала», нажмите "
                                  "большую чёрную кнопку magic: программа заполнит их сама. "
                                  "Ползунок под кнопкой (от 5 до 60 %) задаёт, у какой доли каналов будет "
                                  "небольшая погрешность. Например, 20 % из 400 каналов - у 80 каналов "
                                  "значения немного отличаются от действительных, у остальных 320 совпадают "
                                  "точно. Подпись под ползунком показывает выбранный процент. "
                                  "Погрешность всегда в пределах допуска - все каналы остаются годными; "
                                  "первая точка (0 %) всегда без погрешности. Каждое нажатие даёт новый "
                                  "случайный набор. Готовая копия - в папке «Заполненные измерения (magic)»."),
        ("4", "Рассчитать", "Пересчитывает все каналы заново. При открытии Excel расчёт идёт сам, "
                            "поэтому кнопка нужна, когда вы поменяли что-то в «Настройках» (число "
                            "измерений, ПИП или допуск по умолчанию) или исправили Excel и хотите "
                            "увидеть новый результат."),
        ("5", "Проверить таблицу", "В таблице видно каждый канал: погрешности, допуск и «Годен / Не годен». "
                                   "Щелчок по каналу показывает расчёт снизу. Серые строки - каналы "
                                   "с ошибкой, красные - «Не годен»."),
        ("6", "Ошибки…", "Открывает список всех каналов с ошибками: № и KKS канала, ячейка Excel "
                         "(например, P5104), что в ней стоит и почему это ошибка - текст вместо числа, "
                         "ошибка формулы #VALUE!, нулевой диапазон, нет измерений. Щелчок по строке "
                         "показывает внизу, что исправить, двойной щелчок находит канал в главной таблице. "
                         "Каналы с ошибками в протоколы не попадают."),
        ("7", "Копировать результаты", "Копирует таблицу результатов всех каналов (№, KKS, наименование, "
                                       "a1, a2, b1, b2, gh, gl, ЭТ ИК, ПИП, общая и допустимая погрешность, "
                                       "результат). Дальше откройте Excel и вставьте (Ctrl+V)."),
        ("8", "Очистить", "Убирает все каналы из программы, чтобы начать с другим файлом. "
                          "Сам Excel при этом не меняется."),
        ("9", "Документы", "Впишите окончание заголовка протокола («…аттестации ИК САУГ КА-4 ТЭЦ-17»), "
                           "название перечня и выберите, кто делал протоколы."),
        ("10", "Даты протоколов", "До 15 строк «с № по № - дата», например 1-10 и 07.07.2026. "
                                  "Номер протокола - это № канала из Excel."),
        ("11", "Создать", "«Протоколы» - один Word со всеми протоколами (число в документе ставится само), "
                          "«Только выделенные» - по выделенным строкам таблицы, «Перечень» - перечень и "
                          "градуировочные таблицы, «Открыть папку» - папка с документами и подпапками."),
        ("12", "Звук кнопок", "Кнопка ♪ в правом верхнем углу включает и выключает щелчок, который "
                              "слышно при нажатии на любую кнопку программы. «♪ вкл» - щелчок есть, "
                              "«♪ выкл» - кнопки нажимаются молча. Выбор запоминается до следующего запуска."),
        ("13", "Настройки", "Цветовая палитра, число измерений в точке, папка для документов, "
                            "звук кнопки magic (галочку можно снять, чтобы magic работал без звука)."),
    )

    def _build_help_tab(self, nb):
        outer = ttk.Frame(nb)
        nb.add(outer, text="Помощь")
        canvas = tk.Canvas(outer, highlightthickness=0)
        ys = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=ys.set)
        ys.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        page = ttk.Frame(canvas, padding=14)
        win = canvas.create_window(0, 0, window=page, anchor="nw")
        page.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        self._help_texts = []

        def resize(e):
            canvas.itemconfigure(win, width=e.width)
            for lbl in self._help_texts:  # перенос строк по ширине окна
                lbl.configure(wraplength=max(200, e.width - self.px(90)))

        canvas.bind("<Configure>", resize)
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units")
                        if self.nb.index("current") == 2 else None)
        self.help_canvas = canvas
        self._help_imgs = []

        ttk.Label(page, text="Как пользоваться Plus Ultra", style="Title.TLabel").pack(anchor="w")
        ttk.Label(page, text="Номера на картинке совпадают с шагами ниже.", style="Hint.TLabel").pack(anchor="w", pady=(0, 8))
        self._help_image(page, "help_main.png")
        for num, title, text in self.HELP_STEPS:
            row = ttk.Frame(page)
            row.pack(fill="x", pady=4)
            ttk.Label(row, text=num, style="Step.TLabel", width=3, anchor="center").pack(side="left", anchor="n")
            col = ttk.Frame(row)
            col.pack(side="left", fill="x", expand=True, padx=(8, 0))
            ttk.Label(col, text=title, style="Bold.TLabel").pack(anchor="w")
            lbl = ttk.Label(col, text=text, wraplength=self.px(820), justify="left")
            lbl.pack(anchor="w")
            self._help_texts.append(lbl)
        ttk.Label(page, text="Настройки (14 - цветовая палитра, 15 - звук кнопки magic)", style="Bold.TLabel").pack(anchor="w", pady=(12, 4))
        self._help_image(page, "help_settings.png")
        ttk.Label(page, text="Обычный порядок: 1 → (2) → (3) → 5 → (6) → 9 → 10 → 11. "
                                    "В скобках - если нужно.", style="Bold.TLabel").pack(anchor="w", pady=(12, 0))

    def _help_image(self, parent, name):
        path = asset_path(name)
        if not os.path.exists(path):
            return
        try:
            img = tk.PhotoImage(file=path)
        except tk.TclError:
            return
        self._help_imgs.append(img)
        tk.Label(parent, image=img, borderwidth=1, relief="solid").pack(anchor="w", pady=(0, 8))

    # ---------- темы и звук ----------
    def apply_theme(self, name):
        pal = themes.apply(self.style, name, self)
        for t in self.texts:
            t.configure(background=pal["field"], foreground=pal["fg"], insertbackground=pal["fg"],
                        selectbackground=pal["select"], selectforeground=pal["select_fg"])
        self.tree.tag_configure("bad", foreground=pal["bad"])
        self.tree.tag_configure("err", foreground=pal["muted"])
        self.magic_btn.configure(background=pal["bg"])
        self.help_canvas.configure(background=pal["bg"])
        self.settings["theme"] = name

    def toggle_click_sound(self):
        """Кнопка ♪: звук нажатия кнопок вкл/выкл."""
        self.click_sound.enabled = not self.click_sound.enabled
        self._click_btn_text()
        self.click_sound.play()  # сразу слышно, что звук включился

    def _click_btn_text(self):
        self.click_btn.config(text="♪ вкл" if self.click_sound.enabled else "♪ выкл")

    # ---------- данные ----------
    def paste_clipboard(self):
        try:
            text = self.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("Буфер пуст", "Скопируйте в Excel нужные строки (Ctrl+C).")
            return
        self.xlsx_path = None
        self.add_channels(data.channels_from_rows(data.rows_from_text(text)))

    def open_xlsx(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if path:
            self.open_path(path)

    def on_drop(self, files):
        """Файлы, перетащенные в окно: берётся первый Excel."""
        xl = [f for f in files if f.lower().endswith((".xlsx", ".xlsm"))]
        if not xl:
            messagebox.showwarning("Перетаскивание", "Перетащите файл Excel (.xlsx). Старый формат .xls "
                                   "сначала пересохраните в Excel как .xlsx.")
            return
        self.open_path(xl[0])

    def open_path(self, path):
        """Открыть Excel: каналы из него заменяют список в программе."""
        self.config(cursor="watch")
        self.update()
        try:
            rows = data.rows_from_xlsx(path)
        except Exception as e:  # noqa: BLE001 - показать любую ошибку чтения
            messagebox.showerror("Не удалось открыть", str(e))
            return
        finally:
            self.config(cursor="")
        chs = data.channels_from_rows(rows)
        if not chs:  # прежний файл и таблица остаются как были
            self.add_channels(chs)
            return
        self.xlsx_path = path
        self.channels = []
        self.add_channels(chs)

    def add_channels(self, chs):
        if not chs:
            messagebox.showwarning("Нет данных", "Не найдено ни одного канала (нужны KKS и строки точек).")
            return
        self.channels += chs
        # по умолчанию - протоколы на все вставленные каналы одним документом
        self.start.set(1)
        self.batch.set(min(len(self.channels), MAX_BATCH))
        self.recalc()

    def clear(self):
        self.channels = []
        self.xlsx_path = None
        self.start.set(1)
        self.recalc()

    def recalc(self):
        qm = self._int(self.qm, 20)
        if qm < QM_MIN or qm > QM_MAX or qm % 4:
            qm = max(QM_MIN, min(QM_MAX, qm - qm % 4))
            self.qm.set(qm)
            messagebox.showwarning("Измерений в точке", "Число измерений должно быть кратно 4, от %d до %d. "
                                   "Поставлено %d." % (QM_MIN, QM_MAX, qm))
        pip = data.to_float(self.pip_default.get())
        allowed = data.to_float(self.allowed_default.get())
        for ch in self.channels:
            ch.pip_fallback, ch.allowed_fallback = pip, allowed
            data.calc_channel(ch, qm)
        self.refresh()

    def _int(self, var, default):
        try:
            return int(var.get())
        except (tk.TclError, ValueError):
            return default

    def _digits(self):
        d = self._int(self.digits, 4)
        if not 0 <= d <= DIGITS_MAX:
            d = max(0, min(DIGITS_MAX, d))
            self.digits.set(d)
        return d

    def refresh(self):
        d = self._digits()
        f = lambda v: calcd.fmt(v, d)
        self.tree.delete(*self.tree.get_children())
        for i, ch in enumerate(self.channels):
            r = ch.result
            vals = (ch.num, ch.kks, ch.name, "%s … %s" % (calcd.fmt(ch.low), calcd.fmt(ch.high)),
                    f(r.gh) if r else "", f(r.gl) if r else "", f(ch.et_ik) if r else ch.error,
                    f(ch.total), f(ch.allowed_value), ch.verdict)
            tag = "err" if ch.error else ("bad" if ch.verdict == "Не годен" else "")
            self.tree.insert("", "end", iid=str(i), values=vals, tags=(tag,))
        bad = sum(1 for c in self.channels if c.error)
        self.status.config(text="Каналов: %d" % len(self.channels) + (", с ошибками: %d" % bad if bad else ""))

    def show_detail(self, _=None):
        sel = self.tree.selection()
        self.detail.delete("1.0", "end")
        if not sel or int(sel[0]) >= len(self.channels):
            return
        ch = self.channels[int(sel[0])]
        head = "%s  %s  (%s … %s %s)\n" % (ch.kks, ch.name, calcd.fmt(ch.low), calcd.fmt(ch.high), ch.unit)
        body = calcd.report(ch.result, self._digits()) if ch.result else "Ошибка: " + ch.error
        self.detail.insert("1.0", head + "\n" + body)

    def show_errors(self):
        """Окно: в каких каналах ошибки, в какой ячейке, почему и что сделать."""
        items = [(i, ch, p) for i, ch in enumerate(self.channels) for p in ch.problems()]
        if not items:
            messagebox.showinfo("Ошибки", "Ошибок нет." if self.channels else "Сначала откройте Excel.")
            return
        win = tk.Toplevel(self)
        win.title("Ошибки в данных: %d" % len({i for i, _c, _p in items}))
        win.geometry("%dx%d" % (self.px(1000), self.px(520)))
        win.transient(self)
        sheet = next((ch.sheet for _i, ch, _p in items if ch.sheet), "")
        ttk.Label(win, text=("Лист «%s». " % sheet if sheet else "") +
                  "Щелчок по строке - подробности ниже, двойной щелчок - показать канал в таблице.",
                  style="Hint.TLabel").pack(anchor="w", padx=8, pady=(8, 4))
        info = tk.Text(win, height=6, wrap="word", font=("Segoe UI", 10))
        info.pack(side="bottom", fill="x", padx=8, pady=8)  # до таблицы, чтобы не вытеснялось
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=8)
        cols = (("num", "№", 60), ("kks", "KKS", 140), ("where", "Где", 150),
                ("value", "Что в ячейке", 130), ("why", "Почему это ошибка", 480))
        tree = ttk.Treeview(frame, columns=[c[0] for c in cols], show="headings", height=8)
        for key, title, w in cols:
            tree.heading(key, text=title)
            tree.column(key, width=self.px(w), anchor="w")
        ys = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=ys.set)
        tree.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        for k, (i, ch, p) in enumerate(items):
            tree.insert("", "end", iid=str(k), values=(ch.num, ch.kks, p.where, p.value, p.why))
        pal = themes.PALETTES.get(self.settings["theme"], {})
        if pal:
            info.configure(background=pal["field"], foreground=pal["fg"])

        def select(_e=None):
            sel = tree.selection()
            if not sel:
                return
            _i, ch, p = items[int(sel[0])]
            info.delete("1.0", "end")
            info.insert("1.0", "Канал № %s  %s  %s\nГде: %s%s\nПочему: %s\n%s" % (
                ch.num, ch.kks, ch.name, p.where, ("   (в ячейке: %s)" % p.value) if p.value else "",
                p.why, ("Что сделать: " + p.fix) if p.fix else ""))

        def goto(_e=None):
            sel = tree.selection()
            if not sel:
                return
            iid = str(items[int(sel[0])][0])
            if self.tree.exists(iid):
                self.tree.selection_set(iid)
                self.tree.see(iid)

        tree.bind("<<TreeviewSelect>>", select)
        tree.bind("<Double-1>", goto)
        tree.selection_set("0")

    def copy_results(self):
        d = self._digits()
        lines = ["№\tKKS\tНаименование\ta1\ta2\tb1\tb2\tgh\tgl\tЭТ ИК\tПИП\tОбщая\tДопуст.\tРезультат"]
        for ch in self.channels:
            r = ch.result
            nums = [r.a1, r.a2, r.b1, r.b2, r.gh, r.gl] if r else [None] * 6
            nums += [ch.et_ik, ch.pip_value, ch.total, ch.allowed_value]
            lines.append("\t".join([ch.num, ch.kks, ch.name] + [calcd.fmt(v, d) for v in nums] + [ch.verdict]))
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        self.status.config(text="Результаты скопированы — вставьте в Excel (Ctrl+V)")

    # ---------- документы ----------
    def pick_out_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.out_dir.set(p)

    def sub_dir(self, name):
        """Подпапка в папке документов (вылеченные Excel, файлы magic)."""
        return os.path.join(self.out_dir.get(), name)

    def open_out_dir(self):
        d = self.out_dir.get()
        for name in (HEALED_DIR, MAGIC_DIR):
            os.makedirs(os.path.join(d, name), exist_ok=True)
        os.startfile(d)

    def date_ranges(self):
        """[(с, по, дата)] из полей; ошибки - списком строк."""
        ranges, errors = [], []
        for i, (lo, hi, dt) in enumerate(self.date_vars):
            a, b, d = lo.get().strip(), hi.get().strip(), dt.get().strip()
            if not (a or b or d):
                continue
            try:
                a, b = int(a), int(b or a)
            except ValueError:
                errors.append("строка %d: номера «с» и «по» должны быть числами" % (i + 1))
                continue
            if not DATE_RE.match(d):
                errors.append("строка %d: дата «%s» не в виде 07.07.2026" % (i + 1, d))
                continue
            try:
                datetime.datetime.strptime(d, "%d.%m.%Y")
            except ValueError:
                errors.append("строка %d: такой даты нет в календаре: %s" % (i + 1, d))
                continue
            ranges.append((min(a, b), max(a, b), d))
        return ranges, errors

    def generate(self):
        batch = max(1, min(MAX_BATCH, self._int(self.batch, 20)))
        start = max(1, self._int(self.start, 1)) - 1
        self._generate(self.channels[start:start + batch], start + batch)

    def generate_selected(self):
        sel = [self.channels[int(i)] for i in self.tree.selection() if int(i) < len(self.channels)]
        if len(sel) > MAX_BATCH:
            messagebox.showwarning("Слишком много", "Выделите не больше %d каналов." % MAX_BATCH)
            return
        self._generate(sel, None)

    def _generate(self, chs, next_start):
        chs = [c for c in chs if c.result]
        if not chs:
            messagebox.showwarning("Нет каналов", "Нет рассчитанных каналов для протоколов.")
            return
        ranges, errors = self.date_ranges()
        if errors:
            messagebox.showerror("Даты протоколов", "\n".join(errors))
            return
        no_date = [c.num for c in chs if not docgen.date_for(c, ranges)]
        if no_date and not messagebox.askyesno(
                "Даты протоколов",
                "Без даты протоколов: %d (например, № %s).\nДата будет «__.__.____». Продолжить?"
                % (len(no_date), ", ".join(no_date[:5]))):
            return
        executor = docgen.EXECUTORS[self.executor.current()]
        path = self._run(lambda: docgen.save_protocols(chs, self.out_dir.get(), ranges,
                                                       self.station.get().strip(), executor))
        if not path:
            return
        if next_start is not None and next_start < len(self.channels):
            self.start.set(next_start + 1)
        self._done("Протоколов в документе: %d" % len(chs), path)

    def generate_perechen(self):
        if not self.channels:
            messagebox.showwarning("Нет каналов", "Сначала вставьте каналы из Excel.")
            return
        path = self._run(lambda: docgen.save_perechen(self.channels, self.out_dir.get(),
                                                      self.perechen.get().strip()))
        if path:
            self._done("Каналов в перечне: %d" % len(self.channels), path)

    def _run(self, fn):
        self.config(cursor="watch")
        self.update()
        try:
            return fn()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Ошибка создания документа", str(e))
            return None
        finally:
            self.config(cursor="")

    def _done(self, what, path):
        if messagebox.askyesno("Готово", "%s\n%s\n\nОткрыть документ?" % (what, path)):
            os.startfile(path)

    def on_close(self):
        s = {"pip": self.pip_default.get(), "allowed": self.allowed_default.get(),
             "qm": self._int(self.qm, 20), "digits": self._int(self.digits, 4),
             "batch": self._int(self.batch, 20), "out_dir": self.out_dir.get(),
             "theme": self.settings["theme"], "station": self.station.get(),
             "perechen": self.perechen.get(), "executor": self.executor.current(),
             "dates": [[v.get() for v in row] for row in self.date_vars],
             "magic_pct": self.magic_pct.get(),
             "click_sound": self.click_sound.enabled, "magic_sound": self.magic_sound.enabled}
        try:
            with open(SETTINGS, "w", encoding="utf-8") as f:
                json.dump(s, f, ensure_ascii=False, indent=1)
        except OSError:
            pass
        self.click_sound.close()
        self.magic_sound.close()
        self.destroy()


def enable_dpi_awareness():
    """Чёткий текст при масштабе Windows 125–150 % (на Vista/7 - старый API)."""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def main():
    enable_dpi_awareness()
    App().mainloop()
