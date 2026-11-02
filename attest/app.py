"""Plus Ultra: каналы из Excel -> расчёт коэффициентов -> протоколы и перечень в Word."""
import json
import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import calcd, data, docgen, themes
from .music import Music

APP_NAME = "Plus Ultra"
APP_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else sys.argv[0]))
SETTINGS = os.path.join(APP_DIR, "settings.json")
MAX_BATCH = 3000  # протоколов в одном документе
DATE_RANGES = 15  # строк «с № по № - дата»
DATE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
DEFAULTS = {
    "pip": "", "allowed": "", "qm": 20, "digits": 4, "batch": 20,
    "out_dir": os.path.join(APP_DIR, "Документы"), "theme": themes.DEFAULT,
    "station": "САУГ КА-4 ТЭЦ-17", "perechen": "ПТК САУГ КА-4 Выборгской ТЭЦ (ТЭЦ-17)",
    "executor": 0, "dates": [["", "", ""] for _ in range(DATE_RANGES)],
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
        self.music = Music()
        self.texts = []  # tk.Text - их цвета задаются вручную

        self.style = ttk.Style(self)
        self._fit_to_screen()
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        self._build_main_tab(nb)
        self._build_settings_tab(nb)
        self.apply_theme(self.settings["theme"])
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(2000, self._music_tick)

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
        ttk.Button(bar, text="Открыть Excel…", command=self.open_xlsx).pack(side="left", padx=4)
        ttk.Button(bar, text="Очистить", command=self.clear).pack(side="left")
        ttk.Button(bar, text="Рассчитать", command=self.recalc).pack(side="left", padx=4)
        ttk.Button(bar, text="Копировать результаты", command=self.copy_results).pack(side="left")
        self.music_btn = ttk.Button(bar, text="♪ выкл", width=8, command=self.toggle_music)
        self.music_btn.pack(side="right")
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
        row("Измерений в точке:", ttk.Spinbox(tab, from_=4, to=200, increment=4, width=6, textvariable=self.qm))
        self.digits = tk.IntVar(value=s["digits"])
        row("Знаков после запятой в таблице:", ttk.Spinbox(tab, from_=0, to=15, width=6, textvariable=self.digits))
        self.pip_default = tk.StringVar(value=s["pip"])
        row("ПИП, % (если нет в данных):", ttk.Entry(tab, width=8, textvariable=self.pip_default))
        self.allowed_default = tk.StringVar(value=s["allowed"])
        row("Допустимая, % (если нет в данных):", ttk.Entry(tab, width=8, textvariable=self.allowed_default))
        out = ttk.Frame(tab)
        self.out_dir = tk.StringVar(value=s["out_dir"])
        ttk.Entry(out, textvariable=self.out_dir, width=50).pack(side="left")
        ttk.Button(out, text="…", width=3, command=self.pick_out_dir).pack(side="left", padx=4)
        row("Папка для документов:", out)
        ttk.Button(tab, text="Пересчитать", command=self.recalc).grid(row=r, column=1, sticky="w", pady=(10, 0))

    # ---------- темы и музыка ----------
    def apply_theme(self, name):
        pal = themes.apply(self.style, name, self)
        for t in self.texts:
            t.configure(background=pal["field"], foreground=pal["fg"], insertbackground=pal["fg"],
                        selectbackground=pal["select"], selectforeground=pal["select_fg"])
        self.tree.tag_configure("bad", foreground=pal["bad"])
        self.tree.tag_configure("err", foreground=pal["muted"])
        self.settings["theme"] = name

    def toggle_music(self):
        on = self.music.toggle()
        self.music_btn.config(text="♪ вкл" if on else "♪ выкл")
        if not on and not os.path.exists(self.music.path):
            messagebox.showwarning("Нет музыки", "Файл музыки не найден.")

    def _music_tick(self):
        self.music.keep_looping()
        self.after(2000, self._music_tick)

    # ---------- данные ----------
    def paste_clipboard(self):
        try:
            text = self.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("Буфер пуст", "Скопируйте в Excel нужные строки (Ctrl+C).")
            return
        self.add_channels(data.channels_from_rows(data.rows_from_text(text)))

    def open_xlsx(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        self.config(cursor="watch")
        self.update()
        try:
            rows = data.rows_from_xlsx(path)
        except Exception as e:  # noqa: BLE001 - показать любую ошибку чтения
            messagebox.showerror("Не удалось открыть", str(e))
            return
        finally:
            self.config(cursor="")
        self.add_channels(data.channels_from_rows(rows))

    def add_channels(self, chs):
        if not chs:
            messagebox.showwarning("Нет данных", "Не найдено ни одного канала (нужны KKS и строки точек).")
            return
        self.channels += chs
        self.recalc()

    def clear(self):
        self.channels = []
        self.start.set(1)
        self.recalc()

    def recalc(self):
        qm = self._int(self.qm, 20)
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

    def refresh(self):
        d = self._int(self.digits, 4)
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
        if not sel:
            return
        ch = self.channels[int(sel[0])]
        head = "%s  %s  (%s … %s %s)\n" % (ch.kks, ch.name, calcd.fmt(ch.low), calcd.fmt(ch.high), ch.unit)
        body = calcd.report(ch.result, self._int(self.digits, 4)) if ch.result else "Ошибка: " + ch.error
        self.detail.insert("1.0", head + "\n" + body)

    def copy_results(self):
        d = self._int(self.digits, 4)
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

    def open_out_dir(self):
        d = self.out_dir.get()
        os.makedirs(d, exist_ok=True)
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
            ranges.append((min(a, b), max(a, b), d))
        return ranges, errors

    def generate(self):
        batch = max(1, min(MAX_BATCH, self._int(self.batch, 20)))
        start = max(1, self._int(self.start, 1)) - 1
        self._generate(self.channels[start:start + batch], start + batch)

    def generate_selected(self):
        sel = [self.channels[int(i)] for i in self.tree.selection()]
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
             "dates": [[v.get() for v in row] for row in self.date_vars]}
        try:
            with open(SETTINGS, "w", encoding="utf-8") as f:
                json.dump(s, f, ensure_ascii=False, indent=1)
        except OSError:
            pass
        self.music.close()
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
