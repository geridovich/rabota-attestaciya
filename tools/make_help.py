"""Скриншоты для вкладки «Помощь»: окно с примером данных и номерами шагов.

Запуск (нужны Pillow и пример Excel):
    python tools/make_help.py пример.xlsx
Картинки сохраняются в attest/assets/help_*.png.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from PIL import Image, ImageDraw, ImageFont, ImageGrab  # noqa: E402

from attest import app as appmod  # noqa: E402

OUT = os.path.join(os.path.dirname(__file__), "..", "attest", "assets")
RED = (220, 30, 40)


def find(widget, text):
    """Виджет по тексту (кнопки, подписи, рамки)."""
    for w in widget.winfo_children():
        try:
            if w.cget("text") == text:
                return w
        except Exception:  # noqa: BLE001 - у части виджетов нет text
            pass
        r = find(w, text)
        if r is not None:
            return r
    return None


def box(root, w):
    x, y = w.winfo_rootx() - root.winfo_rootx(), w.winfo_rooty() - root.winfo_rooty()
    return x, y, x + w.winfo_width(), y + w.winfo_height()


def tab_box(root, index):
    """Ярлык вкладки блокнота: ищем по x, где nb.index("@x,y") == index."""
    nb = root.nb
    y = int(10 * root.scale)
    xs = []
    for x in range(0, nb.winfo_width(), 2):
        try:
            if nb.index("@%d,%d" % (x, y)) == index:
                xs.append(x)
        except Exception:  # noqa: BLE001 - не над ярлыком
            pass
    ox, oy = nb.winfo_rootx() - root.winfo_rootx(), nb.winfo_rooty() - root.winfo_rooty()
    return ox + xs[0] - 3, oy, ox + xs[-1] + 3, oy + int(26 * root.scale)


def mark(img, marks, scale):
    """Рамки и красные номера. Номер ставится снаружи рамки на поле вокруг скриншота,
    чтобы не обрезался краем картинки и не закрывал соседние кнопки."""
    r = int(15 * scale)
    m = 3 * r  # поле вокруг скриншота
    out = Image.new("RGB", (img.width + 2 * m, img.height + 2 * m), "white")
    out.paste(img, (m, m))
    d = ImageDraw.Draw(out)
    try:
        font = ImageFont.truetype("arialbd.ttf", int(19 * scale))
    except OSError:
        font = ImageFont.load_default()
    for n, (x0, y0, x1, y1), side in marks:
        x0, y0, x1, y1 = x0 + m, y0 + m, x1 + m, y1 + m
        d.rectangle((x0, y0, x1, y1), outline=RED, width=max(2, int(3 * scale)))
        gap = r + int(4 * scale)
        cx, cy = {
            "l": (x0 - gap, (y0 + y1) / 2), "r": (x1 + gap, (y0 + y1) / 2),
            "t": ((x0 + x1) / 2, y0 - gap), "b": ((x0 + x1) / 2, y1 + gap),
            "tl": (x0 - gap, y0 - gap), "tr": (x1 + gap, y0 - gap),
            "in": (x0 + gap, y0 + gap),
        }[side]
        d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=RED, outline="white", width=2)
        t = str(n)
        tw, th = d.textbbox((0, 0), t, font=font)[2:]
        d.text((cx - tw / 2, cy - th / 2 - 2), t, fill="white", font=font)
    return out


def shot(root):
    root.update()
    x, y = root.winfo_rootx(), root.winfo_rooty()
    return ImageGrab.grab((x, y, x + root.winfo_width(), y + root.winfo_height()))


def main(xlsx):
    appmod.enable_dpi_awareness()
    a = appmod.App()
    a.state("normal")
    a.geometry("1366x700+0+0")
    a.apply_theme("Superman")
    a.attributes("-topmost", True)
    a.open_path(xlsx)
    a.date_vars[0][0].set("1")
    a.date_vars[0][1].set("200")
    a.date_vars[0][2].set("07.07.2026")
    a.out_dir.set(r"C:\Аттестация\Документы")
    a.tree.selection_set("0")
    a.show_detail()
    a.update()
    a.after(800)
    a.update()
    s = a.scale

    def b(text, pad=4):
        x0, y0, x1, y1 = box(a, find(a, text))
        return x0 - pad, y0 - pad, x1 + pad, y1 + pad

    tree = box(a, a.tree)
    magic_btn = box(a, a.magic_btn)
    magic_box = (magic_btn[0] - 30, magic_btn[1] - 6, magic_btn[2] + 30, magic_btn[3] + int(60 * s))
    excel = b("Открыть Excel…")
    marks = [
        (1, (b("Вставить из буфера")[0], excel[1], b("Полечить файл")[0] - 2, excel[3]), "l"),
        (2, b("Полечить файл"), "t"),
        (3, magic_box, "in"),
        (4, b("Рассчитать"), "t"),
        (5, (tree[0], tree[1], tree[2], tree[3]), "l"),
        (6, b("Ошибки…"), "t"),
        (7, b("Очистить"), "t"),
        (8, box(a, find(a, "Документы")), "r"),
        (9, box(a, find(a, "Даты протоколов (дата в виде 07.07.2026)")), "r"),
        (10, box(a, find(a, "Создать")), "r"),
        (11, box(a, a.click_btn), "t"),
        (12, tab_box(a, 1), "t"),
    ]
    img = mark(shot(a).convert("RGB"), marks, s)
    img.save(os.path.join(OUT, "help_main.png"))

    a.nb.select(1)
    a.update()
    a.after(500)
    a.update()
    theme, snd = box(a, a.theme), b("включён")
    img = mark(shot(a).convert("RGB"), [(13, (theme[0] - 4, theme[1] - 4, theme[2] + 4, theme[3] + 4), "r"),
                                        (14, snd, "r")], s)
    img.crop((0, 0, int(img.width * 0.55), int(img.height * 0.5))).save(os.path.join(OUT, "help_settings.png"))
    for f in ("help_main.png", "help_settings.png"):
        im = Image.open(os.path.join(OUT, f))
        if im.width > 1000:  # компактнее для exe
            im = im.resize((1000, int(im.height * 1000 / im.width)), Image.LANCZOS)
        im.save(os.path.join(OUT, f), optimize=True)
    a.click_sound.close()
    a.magic_sound.close()
    a.destroy()


if __name__ == "__main__":
    main(sys.argv[1])
