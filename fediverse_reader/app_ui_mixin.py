from typing import Any

import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

from .common import ReaderStartSettings, choose_tts_backend, normalize_voicevox_url, timeline_kind_to_label, timeline_label_to_kind


class AppUiMixin:
    def _build_ui(self) -> None:
        frm = ttk.Frame(self.root, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="FediverseサーバーURL").grid(row=0, column=0, sticky="w")
        self.instance_var = tk.StringVar(value="https://mastodon.social")
        ttk.Entry(frm, textvariable=self.instance_var, width=64).grid(row=0, column=1, sticky="we")

        ttk.Label(frm, text="保存済みアカウント").grid(row=1, column=0, sticky="w")
        self.account_combo_var = tk.StringVar(value="")
        self.account_combo = ttk.Combobox(
            frm,
            textvariable=self.account_combo_var,
            width=54,
            state="readonly",
            values=[],
        )
        self.account_combo.grid(row=1, column=1, sticky="we")
        self.account_combo.bind("<<ComboboxSelected>>", self.on_account_selected)

        ttk.Label(frm, text="VOICEVOX URL").grid(row=2, column=0, sticky="w")
        self.voicevox_var = tk.StringVar(value="http://127.0.0.1:50021")
        ttk.Entry(frm, textvariable=self.voicevox_var, width=64).grid(row=2, column=1, sticky="we")

        ttk.Label(frm, text="話者").grid(row=3, column=0, sticky="w")
        self.speaker_combo_var = tk.StringVar(value="")
        self.speaker_combo = ttk.Combobox(
            frm,
            textvariable=self.speaker_combo_var,
            width=54,
            state="readonly",
            values=[],
        )
        self.speaker_combo.grid(row=3, column=1, sticky="we")

        speaker_btn_row = ttk.Frame(frm)
        speaker_btn_row.grid(row=4, column=1, sticky="w", pady=(4, 0))
        ttk.Button(speaker_btn_row, text="話者一覧を更新", command=self.on_refresh_speakers).pack(side=tk.LEFT)

        ttk.Label(frm, text="取得間隔(秒)").grid(row=5, column=0, sticky="w")
        self.poll_var = tk.StringVar(value="1")
        ttk.Entry(frm, textvariable=self.poll_var, width=10).grid(row=5, column=1, sticky="w")

        ttk.Label(frm, text="取得件数").grid(row=6, column=0, sticky="w")
        self.limit_var = tk.StringVar(value="10")
        ttk.Entry(frm, textvariable=self.limit_var, width=10).grid(row=6, column=1, sticky="w")

        ttk.Label(frm, text="タイムライン種別").grid(row=7, column=0, sticky="w")
        self.timeline_kind_var = tk.StringVar(value="ローカル")
        self.timeline_combo = ttk.Combobox(
            frm,
            textvariable=self.timeline_kind_var,
            width=20,
            state="readonly",
            values=["ローカル", "ホーム"],
        )
        self.timeline_combo.grid(row=7, column=1, sticky="w")

        ttk.Label(frm, text="読み上げ速度").grid(row=8, column=0, sticky="w")
        self.speech_rate_var = tk.StringVar(value="1.0")
        ttk.Entry(frm, textvariable=self.speech_rate_var, width=10).grid(row=8, column=1, sticky="w")

        ttk.Label(frm, text="読み上げ音量").grid(row=9, column=0, sticky="w")
        self.speech_volume_var = tk.StringVar(value="1.0")
        ttk.Entry(frm, textvariable=self.speech_volume_var, width=10).grid(row=9, column=1, sticky="w")

        ttk.Label(frm, text="読み上げピッチ").grid(row=10, column=0, sticky="w")
        self.speech_pitch_var = tk.StringVar(value="0.0")
        ttk.Entry(frm, textvariable=self.speech_pitch_var, width=10).grid(row=10, column=1, sticky="w")

        ttk.Label(frm, text="長文省略しきい値(N文字)").grid(row=11, column=0, sticky="w")
        self.omit_long_threshold_var = tk.StringVar(value="140")
        ttk.Entry(frm, textvariable=self.omit_long_threshold_var, width=10).grid(row=11, column=1, sticky="w")

        self.read_username_var = tk.BooleanVar(value=True)
        self.read_cw_var = tk.BooleanVar(value=False)
        self.skip_boosts_var = tk.BooleanVar(value=False)
        self.skip_replies_var = tk.BooleanVar(value=False)
        self.omit_body_when_cw_var = tk.BooleanVar(value=False)
        self.skip_posts_with_url_var = tk.BooleanVar(value=False)
        self.skip_posts_with_hashtag_var = tk.BooleanVar(value=False)
        self.auto_start_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm, text="ユーザー名を読む", variable=self.read_username_var).grid(
            row=12, column=1, sticky="w"
        )
        ttk.Checkbutton(frm, text="CW文も読む", variable=self.read_cw_var).grid(row=13, column=1, sticky="w")
        ttk.Checkbutton(frm, text="ブースト除外", variable=self.skip_boosts_var).grid(row=14, column=1, sticky="w")
        ttk.Checkbutton(frm, text="返信除外", variable=self.skip_replies_var).grid(row=15, column=1, sticky="w")
        ttk.Checkbutton(frm, text="CWありは本文を読まない", variable=self.omit_body_when_cw_var).grid(
            row=16, column=1, sticky="w"
        )
        ttk.Checkbutton(frm, text="URL入り投稿を除外", variable=self.skip_posts_with_url_var).grid(
            row=17, column=1, sticky="w"
        )
        ttk.Checkbutton(frm, text="#入り投稿を除外", variable=self.skip_posts_with_hashtag_var).grid(
            row=18, column=1, sticky="w"
        )
        ttk.Checkbutton(frm, text="起動時に自動読み上げ開始", variable=self.auto_start_var).grid(
            row=19, column=1, sticky="w"
        )

        step_row = ttk.Frame(frm)
        step_row.grid(row=20, column=0, columnspan=2, pady=(8, 8), sticky="we")
        ttk.Button(step_row, text="辞書設定", command=self.on_open_dictionary_editor).pack(side=tk.LEFT, padx=4)
        ttk.Button(step_row, text="ミュート/NG設定", command=self.on_open_filter_editor).pack(side=tk.LEFT, padx=4)
        ttk.Button(step_row, text="1) VOICEVOX確認", command=self.on_check_voicevox).pack(side=tk.LEFT, padx=4)
        ttk.Button(step_row, text="2) ログイン開始", command=self.on_start_login).pack(side=tk.LEFT, padx=4)

        ttk.Label(frm, text="認可コード").grid(row=21, column=0, sticky="w")
        self.auth_code_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.auth_code_var, width=64).grid(row=21, column=1, sticky="we")
        ttk.Button(frm, text="ログイン完了", command=self.on_complete_login).grid(
            row=22, column=1, sticky="w", pady=(4, 8)
        )

        run_row = ttk.Frame(frm)
        run_row.grid(row=23, column=0, columnspan=2, sticky="we", pady=(4, 8))
        self.start_btn = ttk.Button(run_row, text="3) 読み上げ開始", command=self.on_start_reading)
        self.start_btn.pack(side=tk.LEFT, padx=4)
        self.stop_btn = ttk.Button(run_row, text="停止", command=self.on_stop_reading, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=4)
        ttk.Button(run_row, text="ログ保存", command=self.on_export_log).pack(side=tk.LEFT, padx=4)
        ttk.Button(run_row, text="VOICEVOXサイトを開く", command=self.on_open_voicevox_site).pack(
            side=tk.LEFT, padx=4
        )

        self.status_var = tk.StringVar(value="未ログイン")
        ttk.Label(frm, textvariable=self.status_var).grid(row=24, column=0, columnspan=2, sticky="w")

        self.log_widget = scrolledtext.ScrolledText(frm, height=20, wrap=tk.WORD, state=tk.DISABLED)
        self.log_widget.grid(row=25, column=0, columnspan=2, sticky="nsew")

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(25, weight=1)

    def _bind_runtime_setting_watchers(self) -> None:
        watched: list[tuple[tk.Variable, str]] = [
            (self.voicevox_var, "VOICEVOX URL"),
            (self.speaker_combo_var, "読み上げモデル"),
            (self.poll_var, "取得間隔"),
            (self.limit_var, "取得件数"),
            (self.timeline_kind_var, "タイムライン種別"),
            (self.speech_rate_var, "読み上げ速度"),
            (self.speech_volume_var, "読み上げ音量"),
            (self.speech_pitch_var, "読み上げピッチ"),
            (self.omit_long_threshold_var, "長文省略しきい値"),
            (self.read_username_var, "ユーザー名読み上げ"),
            (self.read_cw_var, "CW読み上げ"),
            (self.skip_boosts_var, "ブースト除外"),
            (self.skip_replies_var, "返信除外"),
            (self.omit_body_when_cw_var, "CW本文省略"),
            (self.skip_posts_with_url_var, "URL投稿除外"),
            (self.skip_posts_with_hashtag_var, "#投稿除外"),
            (self.auto_start_var, "起動時自動読み上げ"),
            (self.account_combo_var, "選択アカウント"),
        ]
        for var, setting_name in watched:
            var.trace_add(
                "write",
                lambda *args, name=setting_name: self._on_runtime_setting_changed(name),
            )

    def _on_runtime_setting_changed(self, setting_name: str) -> None:
        self._save_current_config()
        if not self.worker:
            return
        if not (self.worker_thread and self.worker_thread.is_alive()):
            return
        if self.worker.stop_event.is_set():
            return
        self.log(f"設定変更を検知したため停止: {setting_name}")
        self.on_stop_reading()

    def _current_speaker_id(self) -> int:
        label = self.speaker_combo_var.get()
        resolved = self.speaker_label_to_id.get(label)
        if resolved is not None:
            self.selected_speaker_id = resolved
            return resolved
        return self.selected_speaker_id

    def _set_speaker_by_id(self, speaker_id: int) -> None:
        self.selected_speaker_id = speaker_id
        for label, sid in self.speaker_label_to_id.items():
            if sid == speaker_id:
                self.speaker_combo_var.set(label)
                return
        values = list(self.speaker_combo.cget("values"))
        if values:
            self.speaker_combo_var.set(values[0])
            fallback_id = self.speaker_label_to_id.get(values[0])
            if fallback_id is not None:
                self.selected_speaker_id = fallback_id

    def _resolve_timeline_kind(self) -> str:
        timeline_kind = timeline_label_to_kind(self.timeline_kind_var.get().strip())
        if timeline_kind in ("local", "home"):
            return timeline_kind
        self.timeline_kind_var.set(timeline_kind_to_label("local"))
        return "local"

    def _parse_start_settings(self) -> ReaderStartSettings | None:
        voicevox_url = normalize_voicevox_url(self.voicevox_var.get())
        self.voicevox_var.set(voicevox_url)
        use_voicevox, message = choose_tts_backend(voicevox_url)
        self.log(message)
        self.status_var.set(message)

        try:
            poll_interval = int(self.poll_var.get().strip())
            fetch_limit = int(self.limit_var.get().strip())
        except ValueError:
            messagebox.showerror("入力エラー", "取得間隔と取得件数は整数で入力してください。")
            return None

        try:
            speech_rate = float(self.speech_rate_var.get().strip())
            speech_volume = float(self.speech_volume_var.get().strip())
            speech_pitch = float(self.speech_pitch_var.get().strip())
        except ValueError:
            messagebox.showerror("入力エラー", "読み上げ速度・音量・ピッチは数値で入力してください。")
            return None

        try:
            omit_long_threshold = int(self.omit_long_threshold_var.get().strip())
        except ValueError:
            messagebox.showerror("入力エラー", "長文省略しきい値は整数で入力してください。")
            return None
        omit_long_threshold = max(1, omit_long_threshold)
        self.omit_long_threshold_var.set(str(omit_long_threshold))

        return ReaderStartSettings(
            voicevox_url=voicevox_url,
            use_voicevox=use_voicevox,
            speaker_id=self._current_speaker_id(),
            poll_interval_sec=max(3, poll_interval),
            fetch_limit=max(1, min(40, fetch_limit)),
            timeline_kind=self._resolve_timeline_kind(),
            speech_rate=speech_rate,
            speech_volume=speech_volume,
            speech_pitch=speech_pitch,
            omit_long_threshold=omit_long_threshold,
        )

