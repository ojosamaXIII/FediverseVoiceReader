import queue
import threading
from datetime import datetime

import tkinter as tk

from .app_config_mixin import AppConfigMixin
from .app_events_mixin import AppEventsMixin
from .app_ui_mixin import AppUiMixin
from .common import OAuthClient, logs_dir_path
from .timeline_speaker import TimelineSpeaker


class App(AppUiMixin, AppConfigMixin, AppEventsMixin):
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Fediverse ローカルタイムライン読み上げ")
        self.root.geometry("860x680")

        self.oauth_client: OAuthClient | None = None
        self.access_token: str | None = None
        self.worker: TimelineSpeaker | None = None
        self.worker_thread: threading.Thread | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.speaker_label_to_id: dict[str, int] = {}
        self.selected_speaker_id = 3
        self.accounts: dict[str, dict[str, str]] = {}
        self.dictionary_entries: list[dict[str, str]] = []
        self.ng_words: list[str] = []
        self.muted_accounts: list[str] = []
        self.log_file_path = logs_dir_path() / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

        self._build_ui()
        self._load_from_config()
        self._bind_runtime_setting_watchers()
        self._drain_log_queue()
        self.root.after(150, self._run_startup_sequence)

