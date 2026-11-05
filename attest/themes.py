"""Цветовые палитры окна."""

PALETTES = {
    "Batman": {  # ночной Готэм: графит и жёлтый сигнал
        "bg": "#16181d", "panel": "#23262e", "field": "#2c3039", "fg": "#e8e8e8", "muted": "#8a8f99",
        "accent": "#f2c200", "accent_fg": "#111111", "accent_hover": "#ffd84d", "hover_fg": "#111111",
        "select": "#f2c200", "select_fg": "#111111", "bad": "#ff6b6b", "border": "#3a3f4a",
    },
    "Superman": {  # синий костюм, красный плащ
        "bg": "#e9f0fb", "panel": "#ffffff", "field": "#ffffff", "fg": "#0d1b3d", "muted": "#7b869c",
        "accent": "#1f4fbf", "accent_fg": "#ffffff", "accent_hover": "#c8102e", "hover_fg": "#ffffff",
        "select": "#c8102e", "select_fg": "#ffffff", "bad": "#c8102e", "border": "#b8c6e3",
    },
    "Spider-Man": {  # красный с синим
        "bg": "#fcecec", "panel": "#ffffff", "field": "#ffffff", "fg": "#1a1a2e", "muted": "#857a80",
        "accent": "#b11226", "accent_fg": "#ffffff", "accent_hover": "#1d3f9c", "hover_fg": "#ffffff",
        "select": "#1d3f9c", "select_fg": "#ffffff", "bad": "#b11226", "border": "#e2b5b9",
    },
    "Iron Man": {  # тёмно-красная броня и золото
        "bg": "#2a0f0f", "panel": "#3a1616", "field": "#4a1d1d", "fg": "#f5e6c8", "muted": "#b39a8a",
        "accent": "#e0a526", "accent_fg": "#2a0f0f", "accent_hover": "#f5c04a", "hover_fg": "#2a0f0f",
        "select": "#b8860b", "select_fg": "#ffffff", "bad": "#ff8a65", "border": "#5c2626",
    },
    "Hulk": {  # зелёный и фиолетовый
        "bg": "#e8f4e5", "panel": "#ffffff", "field": "#ffffff", "fg": "#1f2a1a", "muted": "#7d8a77",
        "accent": "#2e7d32", "accent_fg": "#ffffff", "accent_hover": "#6a2c91", "hover_fg": "#ffffff",
        "select": "#6a2c91", "select_fg": "#ffffff", "bad": "#c62828", "border": "#b9d6b2",
    },
}
DEFAULT = "Superman"


def apply(style, name, root):
    p = PALETTES.get(name, PALETTES[DEFAULT])
    style.theme_use("clam")
    style.configure(".", background=p["bg"], foreground=p["fg"], fieldbackground=p["field"],
                    bordercolor=p["border"], lightcolor=p["bg"], darkcolor=p["bg"],
                    troughcolor=p["panel"], selectbackground=p["select"], selectforeground=p["select_fg"],
                    insertcolor=p["fg"])
    style.configure("TFrame", background=p["bg"])
    style.configure("TLabel", background=p["bg"], foreground=p["fg"])
    style.configure("Hint.TLabel", background=p["bg"], foreground=p["muted"])
    style.configure("Bold.TLabel", background=p["bg"], foreground=p["fg"], font=("Segoe UI", 10, "bold"))
    style.configure("Title.TLabel", background=p["bg"], foreground=p["accent"], font=("Segoe UI", 16, "bold"))
    style.configure("Link.TLabel", background=p["bg"], foreground=p["accent"],
                    font=("Segoe UI", 10, "bold underline"))
    style.configure("Step.TLabel", background=p["accent"], foreground=p["accent_fg"], font=("Segoe UI", 12, "bold"))
    style.configure("TLabelframe", background=p["bg"], bordercolor=p["border"])
    style.configure("TLabelframe.Label", background=p["bg"], foreground=p["accent"])
    style.configure("TButton", background=p["accent"], foreground=p["accent_fg"], bordercolor=p["accent"],
                    focusthickness=1, padding=(8, 3))
    style.map("TButton", background=[("pressed", p["accent_hover"]), ("active", p["accent_hover"])],
              foreground=[("active", p["hover_fg"])])
    for w in ("TEntry", "TSpinbox", "TCombobox"):
        style.configure(w, fieldbackground=p["field"], foreground=p["fg"], background=p["panel"],
                        arrowcolor=p["fg"], insertcolor=p["fg"])
        style.map(w, fieldbackground=[("readonly", p["field"])], foreground=[("readonly", p["fg"])],
                  selectbackground=[("readonly", p["field"])], selectforeground=[("readonly", p["fg"])])
    style.configure("Treeview", background=p["panel"], fieldbackground=p["panel"], foreground=p["fg"],
                    bordercolor=p["border"])
    style.map("Treeview", background=[("selected", p["select"])], foreground=[("selected", p["select_fg"])])
    style.configure("Treeview.Heading", background=p["accent"], foreground=p["accent_fg"], relief="flat")
    style.map("Treeview.Heading", background=[("active", p["accent_hover"])])
    style.configure("TNotebook", background=p["bg"], bordercolor=p["border"])
    style.configure("TNotebook.Tab", background=p["panel"], foreground=p["fg"], padding=(12, 4))
    style.map("TNotebook.Tab", background=[("selected", p["accent"])], foreground=[("selected", p["accent_fg"])])
    style.configure("TPanedwindow", background=p["bg"])
    style.configure("TScrollbar", background=p["panel"], arrowcolor=p["fg"], bordercolor=p["border"])
    root.configure(background=p["bg"])
    # выпадающий список Combobox - обычный tk.Listbox
    root.option_add("*TCombobox*Listbox.background", p["field"])
    root.option_add("*TCombobox*Listbox.foreground", p["fg"])
    root.option_add("*TCombobox*Listbox.selectBackground", p["select"])
    root.option_add("*TCombobox*Listbox.selectForeground", p["select_fg"])
    return p
