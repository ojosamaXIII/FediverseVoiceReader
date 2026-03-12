import queue
import threading
import time
import webbrowser
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from .common import (
    build_authorize_url,
    check_voicevox,
    exchange_code_for_token,
    fetch_voicevox_speakers,
    normalize_instance_url,
    normalize_voicevox_url,
    register_app,
    verify_account_with_backend,
)
from .timeline_speaker import TimelineSpeaker


class AppEventsMixin:
    def on_account_selected(self, _event: Any = None) -> None:
        key = self.account_combo_var.get().strip()
        account = self.accounts.get(key)
        if not account:
            return
        self.instance_var.set(account["instance_url"])
        self.access_token = account["access_token"]
        acct = account.get("acct", "")
        backend = account.get("backend", "mastodon")
        if acct:
            self.status_var.set(f"アカウント選択 ({backend}): @{acct}")
        else:
            self.status_var.set("アカウントを選択しました。")
        self._save_current_config()

    @staticmethod
    def _dictionary_entries_to_text(entries: list[dict[str, str]]) -> str:
        lines: list[str] = []
        for entry in entries:
            mode = str(entry.get("mode", "plain")).strip().lower()
            src = str(entry.get("from", ""))
            dst = str(entry.get("to", ""))
            if not src:
                continue
            if mode == "regex":
                lines.append(f"re:{src}=>{dst}")
            else:
                lines.append(f"{src}={dst}")
        return "\n".join(lines)

    @staticmethod
    def _dictionary_text_to_entries(text: str) -> list[dict[str, str]]:
        entries: list[dict[str, str]] = []
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("re:") and "=>" in line:
                src, dst = line[3:].split("=>", 1)
                src = src.strip()
                if not src:
                    continue
                entries.append({"mode": "regex", "from": src, "to": dst})
                continue
            if "=" not in line:
                continue
            src, dst = line.split("=", 1)
            src = src.strip()
            if not src:
                continue
            entries.append({"mode": "plain", "from": src, "to": dst})
        return entries

    @staticmethod
    def _line_list_to_text(items: list[str]) -> str:
        return "\n".join(items)

    @staticmethod
    def _text_to_line_list(text: str) -> list[str]:
        out: list[str] = []
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            out.append(line)
        return out

    def on_open_dictionary_editor(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("辞書設定")
        dialog.geometry("640x420")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(
            dialog,
            text="通常: 変換前=変換後 / 正規表現: re:検索パターン=>置換後。先頭#の行はコメントです。",
        ).pack(anchor="w", padx=10, pady=(10, 4))

        editor = scrolledtext.ScrolledText(dialog, wrap=tk.WORD)
        editor.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        editor.insert("1.0", self._dictionary_entries_to_text(self.dictionary_entries))

        btn_row = ttk.Frame(dialog)
        btn_row.pack(fill=tk.X, padx=10, pady=(0, 10))

        def on_save() -> None:
            new_entries = self._dictionary_text_to_entries(editor.get("1.0", tk.END))
            self.dictionary_entries = new_entries
            self._save_current_config()
            self.log(f"辞書設定を保存しました ({len(new_entries)}件)。")
            dialog.destroy()

        ttk.Button(btn_row, text="保存", command=on_save).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="キャンセル", command=dialog.destroy).pack(side=tk.LEFT, padx=4)

    def on_open_filter_editor(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("ミュート/NG設定")
        dialog.geometry("760x520")
        dialog.transient(self.root)
        dialog.grab_set()

        frame = ttk.Frame(dialog, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="NGワード（1行に1件。部分一致した投稿をスキップ）").grid(
            row=0, column=0, sticky="w"
        )
        ng_editor = scrolledtext.ScrolledText(frame, wrap=tk.WORD, height=10)
        ng_editor.grid(row=1, column=0, sticky="nsew", pady=(4, 10))
        ng_editor.insert("1.0", self._line_list_to_text(self.ng_words))

        ttk.Label(frame, text="ミュートアカウント（1行に1件。ユーザーID/ユーザー名/表示名）").grid(
            row=2, column=0, sticky="w"
        )
        mute_editor = scrolledtext.ScrolledText(frame, wrap=tk.WORD, height=10)
        mute_editor.grid(row=3, column=0, sticky="nsew", pady=(4, 10))
        mute_editor.insert("1.0", self._line_list_to_text(self.muted_accounts))

        btn_row = ttk.Frame(frame)
        btn_row.grid(row=4, column=0, sticky="w")

        def on_save() -> None:
            self.ng_words = self._text_to_line_list(ng_editor.get("1.0", tk.END))
            self.muted_accounts = self._text_to_line_list(mute_editor.get("1.0", tk.END))
            self._save_current_config()
            self.log(
                f"ミュート/NG設定を保存しました (NG: {len(self.ng_words)}件, ミュート: {len(self.muted_accounts)}件)。"
            )
            dialog.destroy()

        ttk.Button(btn_row, text="保存", command=on_save).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="キャンセル", command=dialog.destroy).pack(side=tk.LEFT, padx=4)

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        frame.rowconfigure(3, weight=1)

    def log(self, msg: str) -> None:
        timestamped = f"{time.strftime('%H:%M:%S')} {msg}"
        self.log_queue.put(msg)
        try:
            path = self.log_file_path
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("a", encoding="utf-8") as f:
                f.write(timestamped + "\n")
        except OSError:
            pass

    def _drain_log_queue(self) -> None:
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_widget.configure(state=tk.NORMAL)
            self.log_widget.insert(tk.END, f"{time.strftime('%H:%M:%S')} {msg}\n")
            self.log_widget.see(tk.END)
            self.log_widget.configure(state=tk.DISABLED)
        self.root.after(200, self._drain_log_queue)

    def on_export_log(self) -> None:
        default_name = f"FediverseVoiceReader_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        target = filedialog.asksaveasfilename(
            title="ログ保存先を選択",
            defaultextension=".log",
            initialfile=default_name,
            filetypes=[("Log files", "*.log"), ("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not target:
            return
        try:
            content = self.log_widget.get("1.0", tk.END)
            Path(target).write_text(content, encoding="utf-8")
            self.log(f"ログを保存しました: {target}")
        except OSError as exc:
            messagebox.showerror("ログ保存エラー", f"ログ保存に失敗しました。\n{exc}")

    def on_refresh_speakers(self) -> None:
        voicevox_url = normalize_voicevox_url(self.voicevox_var.get())
        self.voicevox_var.set(voicevox_url)
        if not voicevox_url:
            messagebox.showinfo("話者一覧", "VOICEVOX URLが空のため、Windows標準読み上げを使用します。")
            self.log("VOICEVOX URL未設定。話者一覧は更新しません。")
            return

        preferred_id = self._current_speaker_id()
        try:
            options = fetch_voicevox_speakers(voicevox_url)
        except requests.RequestException as exc:
            messagebox.showerror("話者取得エラー", f"VOICEVOX話者一覧の取得に失敗しました。\n{exc}")
            self.log(f"話者一覧の取得失敗: {exc}")
            return

        self.speaker_label_to_id = {label: sid for label, sid in options}
        labels = [label for label, _ in options]
        self.speaker_combo.configure(values=labels)
        self._set_speaker_by_id(preferred_id)
        self.log(f"話者一覧を更新しました ({len(labels)}件)。")
        if labels and not self.speaker_combo_var.get():
            self.speaker_combo_var.set(labels[0])

    def on_check_voicevox(self) -> None:
        voicevox_url = normalize_voicevox_url(self.voicevox_var.get())
        self.voicevox_var.set(voicevox_url)
        ok, message = check_voicevox(voicevox_url)
        self.status_var.set(message)
        self.log(message)
        if ok:
            self.on_refresh_speakers()
        elif voicevox_url:
            messagebox.showwarning("VOICEVOX未検出", message)

    def on_open_voicevox_site(self) -> None:
        webbrowser.open("https://voicevox.hiroshiba.jp/")
        webbrowser.open("https://github.com/VOICEVOX/voicevox_engine")
        self.log("VOICEVOXの公式ページを開きました。")

    def on_start_login(self) -> None:
        raw_instance = self.instance_var.get().strip()
        instance = normalize_instance_url(raw_instance)
        if not instance:
            messagebox.showerror("入力エラー", "FediverseサーバーURLを入力してください。")
            return
        self.instance_var.set(instance)
        self.log(f"OAuthログイン開始: 入力='{raw_instance}' 正規化='{instance}'")
        try:
            self.oauth_client = register_app(instance)
            auth_url = build_authorize_url(self.oauth_client)
            webbrowser.open(auth_url)
            if self.oauth_client.backend == "misskey":
                self.status_var.set("ブラウザで許可後、「ログイン完了」を押してください。")
                self.log("Misskey認可セッション作成完了。認可ページを開きました。")
            else:
                self.status_var.set("ブラウザでログインし、認可コードを貼り付けてください。")
                self.log("OAuthアプリ登録完了。認可ページを開きました。")
        except requests.RequestException as exc:
            messagebox.showerror("ログイン開始エラー", f"OAuthログインを開始できませんでした。\n{exc}")
            self.log(f"OAuthログイン開始失敗: {exc}")

    def on_complete_login(self) -> None:
        if not self.oauth_client:
            messagebox.showerror("ログインエラー", "先に「ログイン開始」を押してください。")
            return
        code = self.auth_code_var.get().strip()
        if self.oauth_client.backend == "mastodon" and not code:
            messagebox.showerror("ログインエラー", "認可コードを貼り付けてください。")
            return
        try:
            token = exchange_code_for_token(self.oauth_client, code)
            acct = verify_account_with_backend(
                self.oauth_client.instance_url, token, self.oauth_client.backend
            )
            self.access_token = token
            account_key = f"@{acct} ({self.oauth_client.instance_url})"
            self.accounts[account_key] = {
                "instance_url": self.oauth_client.instance_url,
                "backend": self.oauth_client.backend,
                "access_token": token,
                "acct": acct,
            }
            self._refresh_account_combo(account_key)
            self.instance_var.set(self.oauth_client.instance_url)
            self._save_current_config()
            self.status_var.set(f"ログイン完了 ({self.oauth_client.backend}): @{acct}")
            self.log(f"ログイン完了 ({self.oauth_client.backend}): @{acct}")
        except requests.RequestException as exc:
            messagebox.showerror("ログインエラー", f"認可コードの交換に失敗しました。\n{exc}")
            self.log(f"OAuthログイン完了失敗: {exc}")

    def on_start_reading(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return
        selected_key = self.account_combo_var.get().strip()
        selected = self.accounts.get(selected_key)
        if not selected:
            messagebox.showerror("未ログイン", "保存済みアカウントを選択するか、新規ログインしてください。")
            return
        self.access_token = selected["access_token"]
        instance = selected["instance_url"]
        self.instance_var.set(instance)

        settings = self._parse_start_settings()
        if not settings:
            return

        self.worker = TimelineSpeaker(
            instance_url=instance,
            backend=selected.get("backend", "mastodon"),
            access_token=selected["access_token"],
            voicevox_url=settings.voicevox_url if settings.use_voicevox else "",
            speaker_id=settings.speaker_id,
            poll_interval_sec=settings.poll_interval_sec,
            fetch_limit=settings.fetch_limit,
            timeline_kind=settings.timeline_kind,
            speech_rate=settings.speech_rate,
            speech_volume=settings.speech_volume,
            speech_pitch=settings.speech_pitch,
            omit_long_threshold=settings.omit_long_threshold,
            dictionary_entries=[dict(x) for x in self.dictionary_entries],
            ng_words=list(self.ng_words),
            muted_accounts=list(self.muted_accounts),
            skip_boosts=self.skip_boosts_var.get(),
            skip_replies=self.skip_replies_var.get(),
            omit_body_when_cw=self.omit_body_when_cw_var.get(),
            read_username=self.read_username_var.get(),
            read_cw=self.read_cw_var.get(),
            skip_posts_with_url=self.skip_posts_with_url_var.get(),
            skip_posts_with_hashtag=self.skip_posts_with_hashtag_var.get(),
            logger=self.log,
        )
        self.worker_thread = threading.Thread(target=self.worker.run, daemon=True)
        self.worker_thread.start()
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self.status_var.set("読み上げを開始しました。")
        self._save_current_config()
        if settings.use_voicevox:
            self.log(f"読み上げワーカー起動 (VOICEVOX, 話者ID={settings.speaker_id})。")
        else:
            self.log("読み上げワーカー起動 (Windows標準読み上げ)。")

    def on_stop_reading(self) -> None:
        if self.worker:
            self.worker.stop()
        self.start_btn.configure(state=tk.NORMAL)
        self.stop_btn.configure(state=tk.DISABLED)
        self.status_var.set("停止中...")
        self.log("停止要求を受け付けました。")
        self._save_current_config()

    def shutdown(self) -> None:
        if self.worker:
            self.worker.stop()
        self._save_current_config()

