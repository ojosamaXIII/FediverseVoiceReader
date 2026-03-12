from __future__ import annotations

from collections.abc import Callable

from .common import apply_window_icon, ensure_tk_env


ensure_tk_env()

import tkinter as tk

from .app import App


def create_root() -> tk.Tk:
    root = tk.Tk()
    apply_window_icon(root)
    return root


def _close_handler(root: tk.Tk, app: App) -> Callable[[], None]:
    def on_close() -> None:
        app.shutdown()
        root.destroy()

    return on_close


def run_application() -> None:
    root = create_root()
    app = App(root)
    root.protocol("WM_DELETE_WINDOW", _close_handler(root, app))
    root.mainloop()
