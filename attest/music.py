"""Фоновая музыка по кругу через Windows MCI (без сторонних библиотек, работает с Vista)."""
import ctypes
import os
import sys

ALIAS = "plusultra_music"


def asset_path(name):
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return os.path.join(base, "attest", "assets", name)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", name)


class Music:
    def __init__(self, path=None):
        self.path = path or asset_path("music.mp3")
        self.playing = False
        self.opened = False

    def _mci(self, cmd):
        if sys.platform != "win32":
            return 1
        buf = ctypes.create_unicode_buffer(128)
        return ctypes.windll.winmm.mciSendStringW(cmd, buf, 127, 0)

    def status(self):
        if sys.platform != "win32":
            return ""
        buf = ctypes.create_unicode_buffer(128)
        ctypes.windll.winmm.mciSendStringW("status %s mode" % ALIAS, buf, 127, 0)
        return buf.value

    def play(self):
        if not os.path.exists(self.path):
            return False
        if not self.opened:
            if self._mci('open "%s" type mpegvideo alias %s' % (self.path, ALIAS)) != 0:
                return False
            self.opened = True
        # repeat поддерживается устройством mpegvideo; на всякий случай есть keep_looping
        if self._mci("play %s from 0 repeat" % ALIAS) != 0:
            self._mci("play %s from 0" % ALIAS)
        self.playing = True
        return True

    def keep_looping(self):
        """Вызывать периодически: если трек доиграл (repeat не сработал) - запустить снова."""
        if self.playing and self.status() == "stopped":
            self._mci("play %s from 0" % ALIAS)

    def stop(self):
        if self.opened:
            self._mci("stop %s" % ALIAS)
        self.playing = False

    def toggle(self):
        if self.playing:
            self.stop()
        else:
            self.play()
        return self.playing

    def close(self):
        if self.opened:
            self._mci("close %s" % ALIAS)
            self.opened = False
        self.playing = False


class Sound:
    """Короткий звук один раз (кнопка magic)."""

    def __init__(self, name, alias):
        self.path = asset_path(name)
        self.alias = alias
        self.opened = False

    def _mci(self, cmd):
        if sys.platform != "win32":
            return 1
        buf = ctypes.create_unicode_buffer(128)
        return ctypes.windll.winmm.mciSendStringW(cmd, buf, 127, 0)

    def play(self):
        if not os.path.exists(self.path):
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
