from typing import Any

import requests

from .common import (
    load_config,
    normalize_instance_url,
    normalize_voicevox_url,
    save_config,
    timeline_kind_to_label,
    timeline_label_to_kind,
    verify_account,
    verify_account_with_backend,
)


class AppConfigMixin:
    @staticmethod
    def _normalize_dictionary_entries(raw_entries: Any) -> list[dict[str, str]]:
        entries: list[dict[str, str]] = []
        if not isinstance(raw_entries, list):
            return entries
        for item in raw_entries:
            if not isinstance(item, dict):
                continue
            mode = str(item.get("mode", "plain")).strip().lower()
            src = str(item.get("from", "")).strip()
            dst = str(item.get("to", ""))
            if src:
                entries.append({"mode": mode, "from": src, "to": dst})
        return entries

    @staticmethod
    def _normalize_line_entries(raw_items: Any) -> list[str]:
        if not isinstance(raw_items, list):
            return []
        out: list[str] = []
        for item in raw_items:
            value = str(item).strip()
            if value:
                out.append(value)
        return out

    @staticmethod
    def _normalize_accounts(raw_accounts: Any) -> dict[str, dict[str, str]]:
        normalized: dict[str, dict[str, str]] = {}
        if not isinstance(raw_accounts, dict):
            return normalized
        for key, item in raw_accounts.items():
            if not isinstance(item, dict):
                continue
            account_key = str(key).strip()
            instance_url = str(item.get("instance_url", "")).strip().rstrip("/")
            access_token = str(item.get("access_token", "")).strip()
            acct = str(item.get("acct", "")).strip()
            backend = str(item.get("backend", "mastodon")).strip().lower() or "mastodon"
            if account_key and instance_url and access_token:
                normalized[account_key] = {
                    "instance_url": normalize_instance_url(instance_url),
                    "backend": backend,
                    "access_token": access_token,
                    "acct": acct,
                }
        return normalized

    @staticmethod
    def _load_legacy_dictionary_entries(raw_entries: Any) -> list[dict[str, str]]:
        entries: list[dict[str, str]] = []
        if not isinstance(raw_entries, list):
            return entries
        for item in raw_entries:
            if isinstance(item, (list, tuple)) and len(item) >= 2:
                src = str(item[0]).strip()
                dst = str(item[1])
                if src:
                    entries.append({"mode": "plain", "from": src, "to": dst})
        return entries

    def _build_config_payload(self) -> dict[str, Any]:
        dictionary_entries = []
        for entry in self.dictionary_entries:
            mode = str(entry.get("mode", "plain")).strip().lower()
            src = str(entry.get("from", ""))
            dst = str(entry.get("to", ""))
            if src:
                dictionary_entries.append({"mode": mode, "from": src, "to": dst})
        ng_word_entries = [w for w in self.ng_words if w]
        muted_account_entries = [a for a in self.muted_accounts if a]
        return {
            "instance_url": normalize_instance_url(self.instance_var.get().strip()),
            "voicevox_url": normalize_voicevox_url(self.voicevox_var.get().strip()),
            "speaker_id": self._current_speaker_id(),
            "poll_interval_sec": self.poll_var.get().strip(),
            "fetch_limit": self.limit_var.get().strip(),
            "timeline_kind": timeline_label_to_kind(self.timeline_kind_var.get().strip()),
            "speech_rate": self.speech_rate_var.get().strip(),
            "speech_volume": self.speech_volume_var.get().strip(),
            "speech_pitch": self.speech_pitch_var.get().strip(),
            "omit_long_threshold": self.omit_long_threshold_var.get().strip(),
            "dictionary_entries": dictionary_entries,
            "ng_words": ng_word_entries,
            "muted_accounts": muted_account_entries,
            "skip_boosts": self.skip_boosts_var.get(),
            "skip_replies": self.skip_replies_var.get(),
            "omit_body_when_cw": self.omit_body_when_cw_var.get(),
            "skip_posts_with_url": self.skip_posts_with_url_var.get(),
            "skip_posts_with_hashtag": self.skip_posts_with_hashtag_var.get(),
            "auto_start_on_launch": self.auto_start_var.get(),
            "read_username": self.read_username_var.get(),
            "read_cw": self.read_cw_var.get(),
            "selected_account_key": self.account_combo_var.get().strip(),
            "accounts": self.accounts,
        }

    def _save_current_config(self) -> None:
        save_config(self._build_config_payload())

    def _load_from_config(self) -> None:
        cfg = load_config()
        if not cfg:
            return
        try:
            self.selected_speaker_id = int(cfg.get("speaker_id", self.selected_speaker_id))
        except (TypeError, ValueError):
            self.selected_speaker_id = 3
        self.instance_var.set(normalize_instance_url(str(cfg.get("instance_url", self.instance_var.get()))))
        self.voicevox_var.set(normalize_voicevox_url(str(cfg.get("voicevox_url", self.voicevox_var.get()))))
        self.poll_var.set(str(cfg.get("poll_interval_sec", self.poll_var.get())))
        self.limit_var.set(str(cfg.get("fetch_limit", self.limit_var.get())))
        loaded_timeline_kind = str(cfg.get("timeline_kind", "local")).strip().lower()
        self.timeline_kind_var.set(timeline_kind_to_label(loaded_timeline_kind))
        self.speech_rate_var.set(str(cfg.get("speech_rate", self.speech_rate_var.get())))
        self.speech_volume_var.set(str(cfg.get("speech_volume", self.speech_volume_var.get())))
        self.speech_pitch_var.set(str(cfg.get("speech_pitch", self.speech_pitch_var.get())))
        self.omit_long_threshold_var.set(
            str(cfg.get("omit_long_threshold", self.omit_long_threshold_var.get()))
        )
        self.skip_boosts_var.set(bool(cfg.get("skip_boosts", False)))
        self.skip_replies_var.set(bool(cfg.get("skip_replies", False)))
        self.omit_body_when_cw_var.set(bool(cfg.get("omit_body_when_cw", False)))
        self.skip_posts_with_url_var.set(bool(cfg.get("skip_posts_with_url", False)))
        self.skip_posts_with_hashtag_var.set(bool(cfg.get("skip_posts_with_hashtag", False)))
        self.auto_start_var.set(bool(cfg.get("auto_start_on_launch", False)))
        self.read_username_var.set(bool(cfg.get("read_username", True)))
        self.read_cw_var.set(bool(cfg.get("read_cw", False)))
        self.dictionary_entries = self._normalize_dictionary_entries(cfg.get("dictionary_entries", []))

        # backward-compat: old tuple-style dictionary rules
        if not self.dictionary_entries:
            self.dictionary_entries = self._load_legacy_dictionary_entries(cfg.get("dictionary_rules", []))

        self.ng_words = self._normalize_line_entries(cfg.get("ng_words", []))
        self.muted_accounts = self._normalize_line_entries(cfg.get("muted_accounts", []))
        self.accounts = self._normalize_accounts(cfg.get("accounts", {}))

        # backward-compat: old single-account format
        legacy_token = str(cfg.get("access_token", "")).strip()
        legacy_instance = normalize_instance_url(str(cfg.get("instance_url", "")).strip())
        if legacy_token and legacy_instance and not self.accounts:
            try:
                acct = verify_account(legacy_instance, legacy_token)
                key = f"@{acct} ({legacy_instance})"
                self.accounts[key] = {
                    "instance_url": legacy_instance,
                    "backend": "mastodon",
                    "access_token": legacy_token,
                    "acct": acct,
                }
            except requests.RequestException:
                pass

        self._refresh_account_combo(str(cfg.get("selected_account_key", "")).strip())
        self.log("設定を読み込みました。")

    def _run_startup_sequence(self) -> None:
        self._try_restore_login()
        if self.auto_start_var.get():
            selected_key = self.account_combo_var.get().strip()
            if selected_key and selected_key in self.accounts:
                self.log("起動時自動読み上げを実行します。")
                self.root.after(200, self.on_start_reading)
            else:
                self.log("自動読み上げは有効ですが、保存済みアカウントがないため開始しません。")

    def _try_restore_login(self) -> None:
        selected_key = self.account_combo_var.get().strip()
        if not selected_key or selected_key not in self.accounts:
            self.log("保存済みログイン情報はありません。")
            return
        selected = self.accounts[selected_key]
        instance = selected["instance_url"]
        backend = selected.get("backend", "mastodon")
        token = selected["access_token"]
        if not instance or not token:
            return
        try:
            acct = verify_account_with_backend(instance, token, backend)
            self.access_token = token
            self.instance_var.set(instance)
            self.status_var.set(f"保存済みログインを復元: @{acct}")
            self.log(f"保存済みログインを復元しました: @{acct}")
        except requests.RequestException:
            self.accounts.pop(selected_key, None)
            self._refresh_account_combo("")
            self.access_token = None
            self._save_current_config()
            self.status_var.set("保存済みログインは無効でした。再ログインしてください。")
            self.log("保存済みログインは無効でした。ログイン一覧から削除しました。")

    def _refresh_account_combo(self, preferred_key: str) -> None:
        keys = sorted(self.accounts.keys())
        self.account_combo.configure(values=keys)
        if preferred_key in self.accounts:
            self.account_combo_var.set(preferred_key)
        elif keys:
            self.account_combo_var.set(keys[0])
        else:
            self.account_combo_var.set("")

