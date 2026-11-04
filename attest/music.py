"""Короткие звуки через Windows MCI (без сторонних библиотек, работает с Vista)."""
import ctypes
import os
import sys


def asset_path(name):
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return os.path.join(base, "attest", "assets", name)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", name)


class Sound:
    """Короткий звук один раз (щелчок кнопки, кнопка magic). enabled=False - молчит."""

    def __init__(self, name, alias, enabled=True):
        self.path = asset_path(name)
        self.alias = alias
        self.enabled = enabled
        self.opened = False

    def _mci(self, cmd):
        if sys.platform != "win32":
            return 1
        buf = ctypes.create_unicode_buffer(128)
        return ctypes.windll.winmm.mciSendStringW(cmd, buf, 127, 0)

    def play(self):
        if not self.enabled or not os.path.exists(self.path):
            return False
        if not self.opened:
            if self._mci('open "%s" type mpegvideo alias %s' % (self.path, self.alias)) != 0:
                return False
            self.opened = True
        return self._mci("play %s from 0" % self.alias) == 0

    def close(self):
        if self.opened:
            self._mci("close %s" % self.alias)
            self.opened = False
