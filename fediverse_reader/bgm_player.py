import ctypes
import os
from pathlib import Path


class BgmPlayerError(RuntimeError):
    pass


class BgmPlayer:
    _ALIAS = "fediverse_bgm"

    def __init__(self) -> None:
        self._supported = os.name == "nt"
        self._loaded_path = ""
        self._is_playing = False

    @property
    def loaded_path(self) -> str:
        return self._loaded_path

    @property
    def is_playing(self) -> bool:
        return self._is_playing

    def _send(self, command: str) -> None:
        if not self._supported:
            raise BgmPlayerError("このOSではBGM再生に対応していません。")
        error_code = ctypes.windll.winmm.mciSendStringW(command, None, 0, None)
        if error_code == 0:
            return
        buf = ctypes.create_unicode_buffer(256)
        ctypes.windll.winmm.mciGetErrorStringW(error_code, buf, len(buf))
        message = buf.value or f"mci error code={error_code}"
        raise BgmPlayerError(message)

    @staticmethod
    def _quote(value: str) -> str:
        return '"' + value.replace('"', '""') + '"'

    def load(self, path: str) -> None:
        resolved = str(Path(path).expanduser().resolve())
        if not Path(resolved).exists():
            raise BgmPlayerError(f"ファイルが見つかりません: {resolved}")

        self.close()
        self._send(f"open {self._quote(resolved)} type mpegvideo alias {self._ALIAS}")
        self._loaded_path = resolved
        self._is_playing = False

    def play(self, loop: bool = True) -> None:
        if not self._loaded_path:
            raise BgmPlayerError("先にMP3ファイルを選択してください。")
        command = f"play {self._ALIAS} repeat" if loop else f"play {self._ALIAS}"
        self._send(command)
        self._is_playing = True

    def set_volume(self, percent: int) -> None:
        if not self._loaded_path:
            return
        clamped = max(0, min(100, int(percent)))
        mci_volume = clamped * 10  # 0-1000
        self._send(f"setaudio {self._ALIAS} volume to {mci_volume}")

    def stop(self) -> None:
        if not self._loaded_path:
            self._is_playing = False
            return
        try:
            self._send(f"stop {self._ALIAS}")
        except BgmPlayerError:
            pass
        self._is_playing = False

    def close(self) -> None:
        if not self._loaded_path:
            self._is_playing = False
            return
        try:
            self._send(f"close {self._ALIAS}")
        except BgmPlayerError:
            pass
        self._loaded_path = ""
        self._is_playing = False
