import os
import re
import subprocess
import tempfile
import threading
import time
import wave
from typing import Any, Callable

import requests

from .common import (
    HTTP_USER_AGENT,
    QUOTE_STATUS_URL_RE,
    clamp_float,
    clean_text,
    contains_url_like,
    strip_known_custom_emoji_shortcodes,
)


class TimelineSpeaker:
    def __init__(
        self,
        instance_url: str,
        backend: str,
        access_token: str,
        voicevox_url: str,
        speaker_id: int,
        poll_interval_sec: int,
        fetch_limit: int,
        timeline_kind: str,
        speech_rate: float,
        speech_volume: float,
        speech_pitch: float,
        omit_long_threshold: int,
        dictionary_entries: list[dict[str, str]],
        ng_words: list[str],
        muted_accounts: list[str],
        skip_boosts: bool,
        skip_replies: bool,
        omit_body_when_cw: bool,
        read_username: bool,
        read_cw: bool,
        skip_posts_with_url: bool,
        skip_posts_with_hashtag: bool,
        logger: Callable[[str], None],
    ) -> None:
        self.instance_url = instance_url.rstrip("/")
        self.backend = backend
        self.access_token = access_token
        self.voicevox_url = voicevox_url.rstrip("/")
        self.use_voicevox = bool(self.voicevox_url)
        self.speaker_id = speaker_id
        self.poll_interval_sec = poll_interval_sec
        self.fetch_limit = fetch_limit
        self.timeline_kind = timeline_kind
        self.speech_rate = speech_rate
        self.speech_volume = speech_volume
        self.speech_pitch = speech_pitch
        self.omit_long_threshold = omit_long_threshold
        self.dictionary_plain_rules: list[tuple[str, str]] = []
        self.dictionary_regex_rules: list[tuple[re.Pattern[str], str]] = []
        self._build_dictionary_rules(dictionary_entries)
        self.ng_words = [w.lower() for w in ng_words if w]
        self.muted_accounts = {a.lower() for a in muted_accounts if a}
        self.skip_boosts = skip_boosts
        self.skip_replies = skip_replies
        self.omit_body_when_cw = omit_body_when_cw
        self.read_username = read_username
        self.read_cw = read_cw
        self.skip_posts_with_url = skip_posts_with_url
        self.skip_posts_with_hashtag = skip_posts_with_hashtag
        self.logger = logger
        self.seen_ids: set[str] = set()
        self.stop_event = threading.Event()
        self.playback_lock = threading.Lock()
        self.current_tts_process: subprocess.Popen[str] | None = None

    def _build_dictionary_rules(self, entries: list[dict[str, str]]) -> None:
        self.dictionary_plain_rules = []
        self.dictionary_regex_rules = []
        for entry in entries:
            mode = str(entry.get("mode", "plain")).strip().lower()
            src = str(entry.get("from", ""))
            dst = str(entry.get("to", ""))
            if not src:
                continue
            if mode == "regex":
                try:
                    self.dictionary_regex_rules.append((re.compile(src), dst))
                except re.error as exc:
                    self.logger(f"正規表現が無効です: {src} ({exc})")
            else:
                self.dictionary_plain_rules.append((src, dst))

    def apply_dictionary(self, text: str) -> str:
        output = text
        for src, dst in self.dictionary_plain_rules:
            output = output.replace(src, dst)
        for pattern, dst in self.dictionary_regex_rules:
            output = pattern.sub(dst, output)
        return output

    def headers(self) -> dict[str, str]:
        if self.backend == "misskey":
            return {"Accept": "application/json", "User-Agent": HTTP_USER_AGENT}
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json",
            "User-Agent": HTTP_USER_AGENT,
        }

    @staticmethod
    def is_quote_post(src_status: dict[str, Any], raw_content: str) -> bool:
        # Newer/extended implementations may expose quote metadata.
        for key in ("quote", "quoted_status", "quoted_status_id", "quote_id", "quote_url"):
            if src_status.get(key):
                return True

        # Some servers expose quoted status URL via card.
        card = src_status.get("card")
        if isinstance(card, dict):
            card_url = str(card.get("url", ""))
            if QUOTE_STATUS_URL_RE.search(card_url):
                return True

        # Fallback: quoted status link embedded in content.
        return bool(QUOTE_STATUS_URL_RE.search(raw_content or ""))

    def fetch_timeline(self) -> list[dict[str, Any]]:
        if self.backend == "misskey":
            if self.timeline_kind == "home":
                url = f"{self.instance_url}/api/notes/timeline"
            else:
                url = f"{self.instance_url}/api/notes/local-timeline"
            res = requests.post(
                url,
                headers=self.headers(),
                json={"i": self.access_token, "limit": self.fetch_limit},
                allow_redirects=False,
                timeout=20,
            )
        else:
            if self.timeline_kind == "home":
                url = f"{self.instance_url}/api/v1/timelines/home"
                params = {"limit": str(self.fetch_limit)}
            else:
                url = f"{self.instance_url}/api/v1/timelines/public"
                params = {"local": "true", "limit": str(self.fetch_limit)}
            res = requests.get(
                url,
                headers=self.headers(),
                params=params,
                allow_redirects=False,
                timeout=20,
            )
        if 300 <= res.status_code < 400:
            location = str(res.headers.get("Location", "")).strip()
            raise requests.RequestException(
                f"タイムラインAPIがリダイレクトを返しました ({res.status_code}): "
                f"{location or '(Locationなし)'}"
            )
        res.raise_for_status()
        content_type = str(res.headers.get("Content-Type", "")).lower()
        if "application/json" not in content_type:
            snippet = res.text[:180].replace("\r", " ").replace("\n", " ")
            raise requests.RequestException(
                f"タイムラインAPI応答がJSONではありません。status={res.status_code} body={snippet}"
            )
        payload = res.json()
        return payload if isinstance(payload, list) else []

    def build_message(self, status: dict[str, Any]) -> str:
        if self.backend == "misskey":
            src = status.get("renote") or status
            user_info = src.get("user") or {}
            user = user_info.get("name") or user_info.get("username") or "unknown"
            user = clean_text(str(user))

            raw_content = str(src.get("text", "") or "")
            content = clean_text(raw_content)
            spoiler = clean_text(str(src.get("cw", "") or ""))

            user = self.apply_dictionary(user)
            content = self.apply_dictionary(content)
            spoiler = self.apply_dictionary(spoiler)

            if self.omit_long_threshold > 0 and len(content) > self.omit_long_threshold:
                content = "以下省略"
            if spoiler and self.omit_body_when_cw:
                content = "本文省略"

            parts: list[str] = []
            if self.read_username:
                parts.append(user)
            if status.get("renote"):
                parts.append("リノート")
            if spoiler and self.read_cw:
                parts.append(f"コンテンツ警告 {spoiler}")
            parts.append(content if content else "本文なし")
            return "、".join(parts)

        src = status.get("reblog") or status
        account = src.get("account") or {}
        user = account.get("display_name") or account.get("username") or "unknown"
        user = clean_text(str(user))
        user = strip_known_custom_emoji_shortcodes(user, account.get("emojis") or [])

        raw_content = str(src.get("content", "") or "")
        is_quote = self.is_quote_post(src, raw_content)
        content = clean_text(raw_content)
        spoiler = clean_text(src.get("spoiler_text", ""))
        content = strip_known_custom_emoji_shortcodes(content, src.get("emojis") or [])
        spoiler = strip_known_custom_emoji_shortcodes(spoiler, src.get("emojis") or [])

        user = self.apply_dictionary(user)
        content = self.apply_dictionary(content)
        spoiler = self.apply_dictionary(spoiler)

        if self.omit_long_threshold > 0 and len(content) > self.omit_long_threshold:
            content = "以下省略"
        if spoiler and self.omit_body_when_cw:
            content = "本文省略"

        parts: list[str] = []
        if self.read_username:
            parts.append(user)
        if is_quote:
            parts.append("引用投稿")
        if spoiler and self.read_cw:
            parts.append(f"コンテンツ警告 {spoiler}")
        parts.append(content if content else "本文なし")
        return "、".join(parts)

    def should_skip_status(self, status: dict[str, Any]) -> tuple[bool, str]:
        if self.backend == "misskey":
            if self.skip_boosts and status.get("renote"):
                return True, "リノート除外"

            src = status.get("renote") or status
            if self.skip_replies and src.get("replyId"):
                return True, "返信除外"

            user_info = src.get("user") or {}
            username = str(user_info.get("username", "")).strip()
            host = str(user_info.get("host", "")).strip()
            full_acct = f"{username}@{host}" if username and host else username
            account_candidates = {
                full_acct.lower(),
                username.lower(),
                clean_text(str(user_info.get("name", ""))).strip().lower(),
            }
            account_candidates = {x for x in account_candidates if x}
            if self.muted_accounts and (account_candidates & self.muted_accounts):
                return True, "ミュートアカウント"

            if self.ng_words:
                content = clean_text(str(src.get("text", "") or ""))
                spoiler = clean_text(str(src.get("cw", "") or ""))
                target = f"{content} {spoiler}".lower()
                for word in self.ng_words:
                    if word and word in target:
                        return True, f"NGワード一致: {word}"
            raw_content = str(src.get("text", "") or "")
            raw_spoiler = str(src.get("cw", "") or "")
            raw_target = f"{raw_content}\n{raw_spoiler}"
            if self.skip_posts_with_url and contains_url_like(raw_target):
                return True, "URLを含む投稿除外"
            if self.skip_posts_with_hashtag and "#" in raw_target:
                return True, "#を含む投稿除外"
            return False, ""

        if self.skip_boosts and status.get("reblog"):
            return True, "ブースト除外"

        src = status.get("reblog") or status
        if self.skip_replies and src.get("in_reply_to_id"):
            return True, "返信除外"

        account = src.get("account") or {}

        account_candidates = {
            str(account.get("acct", "")).strip().lower(),
            str(account.get("username", "")).strip().lower(),
            clean_text(str(account.get("display_name", ""))).strip().lower(),
        }
        account_candidates = {x for x in account_candidates if x}
        if self.muted_accounts and (account_candidates & self.muted_accounts):
            return True, "ミュートアカウント"

        if self.ng_words:
            content = clean_text(str(src.get("content", "") or ""))
            spoiler = clean_text(str(src.get("spoiler_text", "") or ""))
            content = strip_known_custom_emoji_shortcodes(content, src.get("emojis") or [])
            spoiler = strip_known_custom_emoji_shortcodes(spoiler, src.get("emojis") or [])
            target = f"{content} {spoiler}".lower()
            for word in self.ng_words:
                if word and word in target:
                    return True, f"NGワード一致: {word}"
        raw_content = str(src.get("content", "") or "")
        raw_spoiler = str(src.get("spoiler_text", "") or "")
        raw_target = f"{raw_content}\n{raw_spoiler}"
        if self.skip_posts_with_url and contains_url_like(raw_target):
            return True, "URLを含む投稿除外"
        if self.skip_posts_with_hashtag and "#" in raw_target:
            return True, "#を含む投稿除外"

        return False, ""

    def speak(self, text: str) -> None:
        if self.stop_event.is_set():
            return
        if not self.use_voicevox:
            rate_scale = clamp_float(self.speech_rate, 0.5, 2.0)
            volume_scale = clamp_float(self.speech_volume, 0.0, 2.0)
            # System.Speech Rate is -10..10, Volume is 0..100.
            win_rate = int(round((rate_scale - 1.0) * 10))
            win_rate = max(-10, min(10, win_rate))
            win_volume = int(round(volume_scale * 100))
            win_volume = max(0, min(100, win_volume))
            script = (
                "Add-Type -AssemblyName System.Speech; "
                "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
                f"$s.Rate = {win_rate}; "
                f"$s.Volume = {win_volume}; "
                "$t = [Console]::In.ReadToEnd(); "
                "$s.Speak($t);"
            )
            proc = subprocess.Popen(
                ["powershell.exe", "-NoProfile", "-Command", script],
                stdin=subprocess.PIPE,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                text=True,
            )
            with self.playback_lock:
                self.current_tts_process = proc
            try:
                if proc.stdin:
                    proc.stdin.write(text)
                    proc.stdin.close()
                while proc.poll() is None:
                    if self.stop_event.is_set():
                        self._interrupt_current_playback()
                        break
                    time.sleep(0.05)
            finally:
                with self.playback_lock:
                    if self.current_tts_process is proc:
                        self.current_tts_process = None
            return

        query_res = requests.post(
            f"{self.voicevox_url}/audio_query",
            params={"text": text, "speaker": self.speaker_id},
            timeout=20,
        )
        query_res.raise_for_status()
        if self.stop_event.is_set():
            return
        query_payload = query_res.json()
        query_payload["speedScale"] = clamp_float(self.speech_rate, 0.5, 2.0)
        query_payload["volumeScale"] = clamp_float(self.speech_volume, 0.0, 2.0)
        query_payload["pitchScale"] = clamp_float(self.speech_pitch, -0.15, 0.15)
        synth_res = requests.post(
            f"{self.voicevox_url}/synthesis",
            params={"speaker": self.speaker_id},
            json=query_payload,
            timeout=30,
        )
        synth_res.raise_for_status()
        if self.stop_event.is_set():
            return
        wav_bytes = synth_res.content
        with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
            f.write(wav_bytes)
            wav_path = f.name
        try:
            import winsound

            duration_sec = 0.0
            with wave.open(wav_path, "rb") as wav_file:
                frame_rate = wav_file.getframerate()
                frame_count = wav_file.getnframes()
                if frame_rate > 0:
                    duration_sec = frame_count / frame_rate
            winsound.PlaySound(wav_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
            deadline = time.monotonic() + max(0.1, duration_sec + 0.3)
            while time.monotonic() < deadline:
                if self.stop_event.is_set():
                    self._interrupt_current_playback()
                    break
                time.sleep(0.05)
        finally:
            try:
                os.remove(wav_path)
            except OSError:
                pass

    def _interrupt_current_playback(self) -> None:
        with self.playback_lock:
            proc = self.current_tts_process
            self.current_tts_process = None
        if proc and proc.poll() is None:
            try:
                proc.terminate()
                proc.wait(timeout=0.5)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
        try:
            import winsound

            winsound.PlaySound(None, winsound.SND_PURGE)
        except Exception:
            pass

    def seed_seen(self) -> None:
        for status in self.fetch_timeline():
            sid = str(status.get("id", ""))
            if sid:
                self.seen_ids.add(sid)

    def run(self) -> None:
        self.logger("タイムライン状態を初期化中...")
        self.seed_seen()
        kind = "ホーム" if self.timeline_kind == "home" else "ローカル"
        self.logger(f"読み上げ開始。新着の{kind}タイムライン投稿のみ対象です。")
        while not self.stop_event.is_set():
            try:
                statuses = self.fetch_timeline()
                new_items: list[dict[str, Any]] = []
                for status in statuses:
                    sid = str(status.get("id", ""))
                    if not sid or sid in self.seen_ids:
                        continue
                    self.seen_ids.add(sid)
                    should_skip, reason = self.should_skip_status(status)
                    if should_skip:
                        self.logger(f"スキップ: {reason}")
                        continue
                    new_items.append(status)
                for status in reversed(new_items):
                    if self.stop_event.is_set():
                        break
                    message = self.build_message(status)
                    self.logger(f"読み上げ: {message}")
                    self.speak(message)
            except requests.RequestException as exc:
                self.logger(f"通信エラー: {exc}")
            except Exception as exc:  # noqa: BLE001
                self.logger(f"エラー: {exc}")
            self.stop_event.wait(self.poll_interval_sec)
        self.logger("読み上げを停止しました。")

    def stop(self) -> None:
        self.stop_event.set()
        self._interrupt_current_playback()


