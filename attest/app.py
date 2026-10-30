"""Окно программы: вставка из Excel, расчёт CalcD по каналам, Word-протоколы."""
import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import calcd, data, protocol

APP_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else sys.argv[0]))
SETTINGS = os.path.join(APP_DIR, "settings.json")
DEFAULTS = {"pip": "", "allowed": "", "qm": 20, "digits": 4, "batch": 20, "template": "", "out_dir": os.path.join(APP_DIR, "Протоколы")}

MAX_BATCH = 3000  # протоколов в одном документе

COLS = (("num", "№", 35), ("kks", "KKS", 110), ("name", "Наименование", 200), ("range", "Диапазон", 90),
        ("gh", "gh", 65), ("gl", "gl", 65), ("et", "ЭТ ИК, %", 65), ("pip", "ПИП, %", 55),
        ("total", "Общая, %", 65), ("allowed", "Допуст., %", 65), ("verdict", "Результат", 70))


def load_settings():
    s = dict(DEFAULTS)
    try:
        with open(SETTINGS, encoding="utf-8") as f:
            s.update(json.load(f))
    except (OSError, ValueError):
        pass
    return s


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Аттестация ИК — расчёт CalcD и протоколы")
        self.scale = max(1.0, self.winfo_fpixels("1i") / 96.0)  # масштаб Windows (125 %, 150 %...)
        self._fit_to_screen()
        self.settings = load_settings()
        self.channels = []
        self.next_index = 0

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        self._build_channels_tab(nb)
        self._build_manual_tab(nb)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def px(self, n):
        return int(n * self.scale)

    def _fit_to_screen(self):
        """Окно по размеру экрана: на ноутбуке 1366x768 открывается развёрнутым."""
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(self.px(1280), sw - 20), min(self.px(820), sh - 80)
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")
        self.minsize(min(self.px(800), sw), min(self.px(500), sh - 80))
        if w >= sw - 20 or h >= sh - 80:
            try:
                self.state("zoomed")
            except tk.TclError:
                pass

    # ---------- вкладка «Каналы из Excel» ----------
    def _build_channels_tab(self, nb):
        tab = ttk.Frame(nb, padding=6)
        nb.add(tab, text="Каналы из Excel")

        bar = ttk.Frame(tab)
        bar.pack(fill="x")
        bar2 = ttk.Frame(tab)
        bar2.pack(fill="x", pady=(4, 0))
        ttk.Button(bar, text="Вставить из буфера", command=self.paste_clipboard).pack(side="left")
        ttk.Button(bar, text="Открыть Excel…", command=self.open_xlsx).pack(side="left", padx=4)
        ttk.Button(bar, text="Очистить", command=self.clear).pack(side="left")
        ttk.Button(bar, text="Рассчитать", command=self.recalc).pack(side="left", padx=8)
        ttk.Button(bar, text="Копировать результаты", command=self.copy_results).pack(side="left")
        self.status = ttk.Label(bar, text="")
        self.status.pack(side="right")
        bar = bar2
        ttk.Label(bar, text="Измерений в точке:").pack(side="left")
        self.qm = tk.IntVar(value=self.settings["qm"])
        ttk.Spinbox(bar, from_=4, to=200, increment=4, width=5, textvariable=self.qm).pack(side="left")
        ttk.Label(bar, text="  Знаков после запятой:").pack(side="left")
        self.digits = tk.IntVar(value=self.settings["digits"])
        ttk.Spinbox(bar, from_=0, to=15, width=4, textvariable=self.digits).pack(side="left")
        ttk.Label(bar, text="  Если нет в данных: ПИП, %").pack(side="left")
        self.pip_default = tk.StringVar(value=self.settings["pip"])
        ttk.Entry(bar, width=6, textvariable=self.pip_default).pack(side="left")
        ttk.Label(bar, text=" допуст., %").pack(side="left")
        self.allowed_default = tk.StringVar(value=self.settings["allowed"])
        ttk.Entry(bar, width=6, textvariable=self.allowed_default).pack(side="left")

        paned = ttk.PanedWindow(tab, orient="vertical")
        paned.pack(fill="both", expand=True, pady=6)

        top = ttk.Frame(paned)
        self.tree = ttk.Treeview(top, columns=[c[0] for c in COLS], show="headings", selectmode="extended")
        for key, title, w in COLS:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=self.px(w), minwidth=self.px(30), anchor="w" if key in ("kks", "name") else "center")
        ys = ttk.Scrollbar(top, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ys.set)
        self.tree.pack(side="left", fill="both", expand=True)
        ys.pack(side="right", fill="y")
        self.tree.tag_configure("bad", foreground="#b00020")
        self.tree.tag_configure("err", foreground="#888888")
        self.tree.bind("<<TreeviewSelect>>", self.show_detail)
        paned.add(top, weight=3)

        bottom = ttk.Frame(paned)
        self.detail = tk.Text(bottom, width=40, height=8, font=("Courier New", 10), wrap="none")
        dx = ttk.Scrollbar(bottom, orient="horizontal", command=self.detail.xview)
        self.detail.configure(xscrollcommand=dx.set)
        self.detail.pack(fill="both", expand=True)
        dx.pack(fill="x")
        paned.add(bottom, weight=2)

        gen = ttk.LabelFrame(tab, text="Word-протоколы", padding=6)
        gen.pack(fill="x", side="bottom", before=paned)  # всегда видна, таблица ужимается
        ttk.Label(gen, text="Шаблон:").grid(row=0, column=0, sticky="w")
        self.template = tk.StringVar(value=self.settings["template"])
        ttk.Entry(gen, textvariable=self.template, width=40).grid(row=0, column=1, sticky="we")
        ttk.Button(gen, text="…", width=3, command=self.pick_template).grid(row=0, column=2)
        ttk.Label(gen, text="(пусто — временный образец)").grid(row=0, column=3, sticky="w")
        ttk.Label(gen, text="Папка:").grid(row=1, column=0, sticky="w")
        self.out_dir = tk.StringVar(value=self.settings["out_dir"])
        ttk.Entry(gen, textvariable=self.out_dir, width=40).grid(row=1, column=1, sticky="we")
        ttk.Button(gen, text="…", width=3, command=self.pick_out_dir).grid(row=1, column=2)
        row2 = ttk.Frame(gen)
        row2.grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))
        ttk.Label(row2, text="Протоколов в документе (1–3000):").pack(side="left")
        self.batch = tk.IntVar(value=self.settings["batch"])
        ttk.Spinbox(row2, from_=1, to=3000, width=6, textvariable=self.batch).pack(side="left")
        ttk.Label(row2, text="  со строки списка:").pack(side="left")
        self.start = tk.IntVar(value=1)
        ttk.Spinbox(row2, from_=1, to=100000, width=7, textvariable=self.start).pack(side="left")
        ttk.Button(row2, text="Создать протоколы", command=self.generate).pack(side="left", padx=8)
        ttk.Button(row2, text="Только выделенные", command=self.generate_selected).pack(side="left")
        ttk.Button(row2, text="Открыть папку", command=self.open_out_dir).pack(side="left", padx=8)
        gen.columnconfigure(1, weight=1)

    # ---------- вкладка «Ручной ввод (как CalcD)» ----------
    def _build_manual_tab(self, nb):
        tab = ttk.Frame(nb, padding=6)
        nb.add(tab, text="Ручной ввод (как CalcD)")
        left = ttk.Frame(tab)
        left.pack(side="left", fill="y")
        ttk.Label(left, text="Измерения (по одному в строке):").pack(anchor="w")
        self.m_values = tk.Text(left, width=18, height=10)
        self.m_values.pack(fill="y", expand=True)
        mid = ttk.Frame(tab)
        mid.pack(side="left", fill="y", padx=8)
        ttk.Label(mid, text="Действительные значения точек:").pack(anchor="w")
        self.m_xd = tk.Text(mid, width=18, height=10)
        self.m_xd.pack()
        ttk.Label(mid, text="Измерений в точке:").pack(anchor="w", pady=(8, 0))
        self.m_qm = tk.IntVar(value=20)
        ttk.Spinbox(mid, from_=4, to=200, increment=4, width=6, textvariable=self.m_qm).pack(anchor="w")
        ttk.Label(mid, text="Последовательность:").pack(anchor="w", pady=(8, 0))
        self.m_mode = ttk.Combobox(mid, values=calcd.MODES, state="readonly", width=44)
        self.m_mode.current(0)
        self.m_mode.pack(anchor="w")
        ttk.Button(mid, text="Расчёт", command=self.manual_calc).pack(anchor="w", pady=8)
        self.m_out = tk.Text(tab, width=40, font=("Courier New", 10), wrap="none")
        self.m_out.pack(side="left", fill="both", expand=True)

    def manual_calc(self):
        try:
            vals = [data.to_float(s) for s in self.m_values.get("1.0", "end").split()]
            xd = [data.to_float(s) for s in self.m_xd.get("1.0", "end").split()]
            if None in vals or None in xd:
                raise calcd.CalcError("Есть нечисловые значения")
            res = calcd.calc(vals, xd, int(self.m_qm.get()), self.m_mode.current())
        except (calcd.CalcError, ValueError, tk.TclError) as e:
            messagebox.showerror("Ошибка", str(e))
            return
        self.m_out.delete("1.0", "end")
        self.m_out.insert("1.0", calcd.report(res))

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
        try:
            rows = data.rows_from_xlsx(path)
        except Exception as e:  # noqa: BLE001 - показать любую ошибку чтения
            messagebox.showerror("Не удалось открыть", str(e))
            return
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
            vals = (ch.num, ch.kks, ch.name, f"{calcd.fmt(ch.low)} … {calcd.fmt(ch.high)}",
                    f(r.gh) if r else "", f(r.gl) if r else "", f(ch.et_ik) if r else ch.error,
                    f(ch.pip_value), f(ch.total), f(ch.allowed_value), ch.verdict)
            tag = "err" if ch.error else ("bad" if ch.verdict == "Не годен" else "")
            self.tree.insert("", "end", iid=str(i), values=vals, tags=(tag,))
        bad = sum(1 for c in self.channels if c.error)
        self.status.config(text=f"Каналов: {len(self.channels)}" + (f", с ошибками: {bad}" if bad else ""))

    def show_detail(self, _=None):
        sel = self.tree.selection()
        self.detail.delete("1.0", "end")
        if not sel:
            return
        ch = self.channels[int(sel[0])]
        head = f"{ch.kks}  {ch.name}  ({calcd.fmt(ch.low)} … {calcd.fmt(ch.high)} {ch.unit})\n"
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

    # ---------- протоколы ----------
    def pick_template(self):
        p = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if p:
            self.template.set(p)

    def pick_out_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.out_dir.set(p)

    def open_out_dir(self):
        d = self.out_dir.get()
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    def generate(self):
        batch = max(1, min(MAX_BATCH, self._int(self.batch, 20)))
        start = max(1, self._int(self.start, 1)) - 1
        self._generate(self.channels[start:start + batch], start + batch)

    def generate_selected(self):
        sel = [self.channels[int(i)] for i in self.tree.selection()]
        if len(sel) > MAX_BATCH:
            messagebox.showwarning("Слишком много", f"Выделите не больше {MAX_BATCH} каналов.")
            return
        self._generate(sel, None)

    def _generate(self, chs, next_start):
        chs = [c for c in chs if c.result]
        if not chs:
            messagebox.showwarning("Нет каналов", "Нет рассчитанных каналов для протоколов.")
            return
        tpl = self.template.get().strip() or None
        if tpl and not os.path.exists(tpl):
            messagebox.showerror("Нет шаблона", tpl)
            return
        self.config(cursor="watch")
        self.update()
        try:
            path = protocol.generate(chs, self.out_dir.get(), tpl, self._int(self.digits, 4))
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Ошибка создания протоколов", str(e))
            return
        finally:
            self.config(cursor="")
        if next_start is not None and next_start < len(self.channels):
            self.start.set(next_start + 1)
        if messagebox.askyesno("Готово", f"Протоколов в документе: {len(chs)}\n{path}\n\nОткрыть документ?"):
            os.startfile(path)

    def on_close(self):
        s = {"pip": self.pip_default.get(), "allowed": self.allowed_default.get(), "qm": self._int(self.qm, 20), "digits": self._int(self.digits, 4),
             "batch": self._int(self.batch, 20), "template": self.template.get(), "out_dir": self.out_dir.get()}
        try:
            with open(SETTINGS, "w", encoding="utf-8") as f:
                json.dump(s, f, ensure_ascii=False, indent=1)
        except OSError:
            pass
        self.destroy()


def enable_dpi_awareness():
    """Чёткий текст при масштабе Windows 125–150 % (на Windows 7 - старый API)."""
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
