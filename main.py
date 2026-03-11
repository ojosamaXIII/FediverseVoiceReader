import html
import json
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable
from urllib.parse import urlencode, urlsplit

import requests
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


APP_DIR_NAME = "FediverseVoiceReader"
TAG_RE = re.compile(r"<[^>]+>")
SPACE_RE = re.compile(r"\s+")
URL_RE = re.compile(r"https?://[^\s]+", flags=re.IGNORECASE)
URL_LIKE_RE = re.compile(r"\b(?:www\.)[^\s]+", flags=re.IGNORECASE)
URL_LEADING_LINE_RE = re.compile(r"^\s*https?://\S.*$", flags=re.IGNORECASE | re.MULTILINE)
QUOTE_STATUS_URL_RE = re.compile(
    r"https?://[^\s\"'<>]+(?:/@[^/\s\"'<>]+/\d+|/users/[^/\s\"'<>]+/statuses/\d+)",
    flags=re.IGNORECASE,
)
CUSTOM_EMOJI_RE = re.compile(r":[a-zA-Z0-9_+-]+:")
EMOJI_RE = re.compile(
    "["
    "\U0001F1E0-\U0001F1FF"  # flags
    "\U0001F300-\U0001F5FF"  # symbols & pictographs
    "\U0001F600-\U0001F64F"  # emoticons
    "\U0001F680-\U0001F6FF"  # transport & map
    "\U0001F700-\U0001F77F"  # alchemical symbols
    "\U0001F780-\U0001F7FF"  # geometric extended
    "\U0001F800-\U0001F8FF"  # supplemental arrows-c
    "\U0001F900-\U0001F9FF"  # supplemental symbols and pictographs
    "\U0001FA00-\U0001FAFF"  # chess, symbols, pictographs extended-a
    "\U00002700-\U000027BF"  # dingbats
    "\U00002600-\U000026FF"  # misc symbols
    "]+",
    flags=re.UNICODE,
)
EMOJI_JOINER_RE = re.compile(r"[\u200d\ufe0f]")
REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob"
HTTP_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/134.0.0.0 Safari/537.36 FediverseVoiceReader/1.0"
)


def clean_text(raw_html: str) -> str:
    text = TAG_RE.sub(" ", raw_html or "")
    text = html.unescape(text)
    # Omit whole lines that start with URL so they are not read aloud at all.
    text = URL_LEADING_LINE_RE.sub(" ", text)
    text = URL_RE.sub(" URL省略 ", text)
    text = URL_LIKE_RE.sub(" URL省略 ", text)
    text = CUSTOM_EMOJI_RE.sub(" ", text)
    text = EMOJI_RE.sub(" ", text)
    text = EMOJI_JOINER_RE.sub("", text)
    return SPACE_RE.sub(" ", text).strip()


def normalize_url(raw: str, default_scheme: str) -> str:
    value = (raw or "").strip()
    if not value:
        return ""
    if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", value):
        return value.rstrip("/")
    return f"{default_scheme}://{value}".rstrip("/")


def normalize_instance_url(raw: str) -> str:
    value = normalize_url(raw, "https")
    if not value:
        return ""
    parts = urlsplit(value)
    host = (parts.hostname or "").strip().lower()
    if not host:
        return value

    # mstdn.jp では media サブドメインのURLが共有されることがあり、
    # それをOAuth先に使うと接続失敗するため本体ドメインへ補正する。
    if host == "media.mstdn.jp":
        host = "mstdn.jp"

    port = f":{parts.port}" if parts.port else ""
    return f"{parts.scheme or 'https'}://{host}{port}".rstrip("/")


def normalize_voicevox_url(raw: str) -> str:
    value = (raw or "").strip()
    if not value:
        return ""
    if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", value):
        return value.rstrip("/")
    if value.startswith(("localhost", "127.", "[::1]")):
        return f"http://{value}".rstrip("/")
    return f"http://{value}".rstrip("/")


def timeline_kind_to_label(kind: str) -> str:
    return "ホーム" if kind == "home" else "ローカル"


def timeline_label_to_kind(label: str) -> str:
    return "home" if label == "ホーム" else "local"


def strip_known_custom_emoji_shortcodes(text: str, emoji_defs: Any) -> str:
    if not text or not isinstance(emoji_defs, list):
        return text
    output = text
    for item in emoji_defs:
        if not isinstance(item, dict):
            continue
        shortcode = str(item.get("shortcode", "")).strip()
        if not shortcode:
            continue
        output = output.replace(f":{shortcode}:", " ")
    return SPACE_RE.sub(" ", output).strip()


def resource_path(name: str) -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS")) / name
    return Path(__file__).resolve().parent / name


def apply_window_icon(root: tk.Tk) -> None:
    # Ensure Windows taskbar groups this app with its own icon.
    if os.name == "nt":
        try:
            import ctypes

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "OjosamaXIII.FediverseVoiceReader"
            )
        except Exception:
            pass

    # On Windows, iconbitmap(.ico) is the most reliable for taskbar/window icon.
    ico_path = resource_path("FediverseVoiceReader.ico")
    if ico_path.exists():
        try:
            root.iconbitmap(default=str(ico_path))
        except tk.TclError:
            pass

    png_path = resource_path("FediverseVoiceReader.png")
    if png_path.exists():
        try:
            icon = tk.PhotoImage(file=str(png_path))
            root.iconphoto(True, icon)
            root._icon_photo = icon  # type: ignore[attr-defined]
            return
        except tk.TclError:
            pass


def config_path() -> Path:
    appdata = os.getenv("APPDATA")
    base_dir = Path(appdata) if appdata else (Path.home() / ".config")
    return base_dir / APP_DIR_NAME / "config.json"


def logs_dir_path() -> Path:
    return config_path().parent / "logs"


def load_config() -> dict[str, Any]:
    path = config_path()
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_config(payload: dict[str, Any]) -> None:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


@dataclass
class OAuthClient:
    instance_url: str
    client_id: str
    client_secret: str


def register_app(instance_url: str) -> OAuthClient:
    base = normalize_instance_url(instance_url).rstrip("/")
    if not base:
        raise ValueError("インスタンスURLが空です。")
    res = requests.post(
        f"{base}/api/v1/apps",
        data={
            "client_name": "Fediverse Timeline Reader",
            "redirect_uris": REDIRECT_URI,
            "scopes": "read",
            "website": "",
        },
        headers={"Accept": "application/json", "User-Agent": HTTP_USER_AGENT},
        allow_redirects=False,
        timeout=20,
    )
    if 300 <= res.status_code < 400:
        location = str(res.headers.get("Location", "")).strip()
        raise requests.RequestException(
            f"OAuthアプリ登録がリダイレクトされました ({res.status_code}): {location or '(Locationなし)'}"
        )
    res.raise_for_status()
    try:
        payload = res.json()
    except ValueError as exc:
        snippet = res.text[:180].replace("\r", " ").replace("\n", " ")
        raise requests.RequestException(
            f"OAuthアプリ登録の応答がJSONではありません。status={res.status_code} body={snippet}"
        ) from exc
    return OAuthClient(
        instance_url=base,
        client_id=payload["client_id"],
        client_secret=payload["client_secret"],
    )


def build_authorize_url(client: OAuthClient) -> str:
    params = {
        "client_id": client.client_id,
        "scope": "read",
        "redirect_uri": REDIRECT_URI,
        "response_type": "code",
    }
    return f"{client.instance_url}/oauth/authorize?{urlencode(params)}"


def exchange_code_for_token(client: OAuthClient, code: str) -> str:
    res = requests.post(
        f"{client.instance_url}/oauth/token",
        data={
            "grant_type": "authorization_code",
            "code": code.strip(),
            "client_id": client.client_id,
            "client_secret": client.client_secret,
            "redirect_uri": REDIRECT_URI,
            "scope": "read",
        },
        headers={"Accept": "application/json", "User-Agent": HTTP_USER_AGENT},
        timeout=20,
    )
    res.raise_for_status()
    payload = res.json()
    return payload["access_token"]


def verify_account(instance_url: str, access_token: str) -> str:
    res = requests.get(
        f"{instance_url.rstrip('/')}/api/v1/accounts/verify_credentials",
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
            "User-Agent": HTTP_USER_AGENT,
        },
        timeout=20,
    )
    res.raise_for_status()
    account = res.json()
    return account.get("acct") or account.get("username") or "unknown"


def check_voicevox(voicevox_url: str) -> tuple[bool, str]:
    base = normalize_voicevox_url(voicevox_url)
    if not base:
        return False, "VOICEVOX URL未設定。Windows標準読み上げを使用します。"

    try:
        res = requests.get(f"{base}/version", timeout=5)
        if res.ok:
            return True, f"VOICEVOX Engine接続OK (version={res.text.strip()})"
    except requests.RequestException:
        pass
    try:
        res = requests.get(f"{base}/speakers", timeout=5)
        if res.ok:
            return True, "VOICEVOX Engine接続OK"
    except requests.RequestException:
        pass
    return False, "VOICEVOX Engineに接続できません。Windows標準読み上げに自動切替します。"


def choose_tts_backend(voicevox_url: str) -> tuple[bool, str]:
    ok, message = check_voicevox(voicevox_url)
    if ok:
        return True, message
    return False, message


def clamp_float(value: float, min_value: float, max_value: float) -> float:
    return max(min_value, min(max_value, value))


def speak_with_windows_tts(text: str, rate_scale: float, volume_scale: float) -> None:
    rate_scale = clamp_float(rate_scale, 0.5, 2.0)
    volume_scale = clamp_float(volume_scale, 0.0, 2.0)
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
    subprocess.run(
        ["powershell.exe", "-NoProfile", "-Command", script],
        input=text,
        text=True,
        check=True,
        timeout=180,
    )


def fetch_voicevox_speakers(voicevox_url: str) -> list[tuple[str, int]]:
    base = normalize_voicevox_url(voicevox_url)
    res = requests.get(f"{base}/speakers", timeout=10)
    res.raise_for_status()
    payload = res.json()
    options: list[tuple[str, int]] = []
    if not isinstance(payload, list):
        return options
    for speaker in payload:
        name = str(speaker.get("name", "Unknown"))
        styles = speaker.get("styles") or []
        for style in styles:
            sid = style.get("id")
            style_name = style.get("name") or "Default"
            if isinstance(sid, int):
                label = f"{name} ({style_name}) [ID:{sid}]"
                options.append((label, sid))
    options.sort(key=lambda x: x[1])
    return options


class TimelineSpeaker:
    def __init__(
        self,
        instance_url: str,
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
        logger: Callable[[str], None],
    ) -> None:
        self.instance_url = instance_url.rstrip("/")
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
        self.logger = logger
        self.seen_ids: set[str] = set()
        self.stop_event = threading.Event()

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
                    self.logger(f"霎樊嶌regex無効: {src} ({exc})")
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

        return False, ""

    def speak(self, text: str) -> None:
        if not self.use_voicevox:
            speak_with_windows_tts(
                text=text,
                rate_scale=self.speech_rate,
                volume_scale=self.speech_volume,
            )
            return

        query_res = requests.post(
            f"{self.voicevox_url}/audio_query",
            params={"text": text, "speaker": self.speaker_id},
            timeout=20,
        )
        query_res.raise_for_status()
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
        wav_bytes = synth_res.content
        with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
            f.write(wav_bytes)
            wav_path = f.name
        try:
            import winsound

            winsound.PlaySound(wav_path, winsound.SND_FILENAME)
        finally:
            try:
                os.remove(wav_path)
            except OSError:
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


class App:
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
        self.accounts: dict[str, dict[str, str]] = {}
        self.dictionary_entries: list[dict[str, str]] = []
        self.ng_words: list[str] = []
        self.muted_accounts: list[str] = []
        self.log_file_path = logs_dir_path() / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

        self._build_ui()
        self._load_from_config()
        self._drain_log_queue()
        self.root.after(150, self._run_startup_sequence)

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
        ttk.Checkbutton(frm, text="起動時に自動読み上げ開始", variable=self.auto_start_var).grid(
            row=17, column=1, sticky="w"
        )

        step_row = ttk.Frame(frm)
        step_row.grid(row=18, column=0, columnspan=2, pady=(8, 8), sticky="we")
        ttk.Button(step_row, text="辞書設定", command=self.on_open_dictionary_editor).pack(side=tk.LEFT, padx=4)
        ttk.Button(step_row, text="ミュート/NG設定", command=self.on_open_filter_editor).pack(side=tk.LEFT, padx=4)
        ttk.Button(step_row, text="1) VOICEVOX確認", command=self.on_check_voicevox).pack(side=tk.LEFT, padx=4)
        ttk.Button(step_row, text="2) ログイン開始", command=self.on_start_login).pack(side=tk.LEFT, padx=4)

        ttk.Label(frm, text="認可コード").grid(row=19, column=0, sticky="w")
        self.auth_code_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.auth_code_var, width=64).grid(row=19, column=1, sticky="we")
        ttk.Button(frm, text="ログイン完了", command=self.on_complete_login).grid(
            row=20, column=1, sticky="w", pady=(4, 8)
        )

        run_row = ttk.Frame(frm)
        run_row.grid(row=21, column=0, columnspan=2, sticky="we", pady=(4, 8))
        self.start_btn = ttk.Button(run_row, text="3) 読み上げ開始", command=self.on_start_reading)
        self.start_btn.pack(side=tk.LEFT, padx=4)
        self.stop_btn = ttk.Button(run_row, text="停止", command=self.on_stop_reading, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=4)
        ttk.Button(run_row, text="ログ保存", command=self.on_export_log).pack(side=tk.LEFT, padx=4)
        ttk.Button(run_row, text="VOICEVOXサイトを開く", command=self.on_open_voicevox_site).pack(
            side=tk.LEFT, padx=4
        )

        self.status_var = tk.StringVar(value="未ログイン")
        ttk.Label(frm, textvariable=self.status_var).grid(row=22, column=0, columnspan=2, sticky="w")

        self.log_widget = scrolledtext.ScrolledText(frm, height=20, wrap=tk.WORD, state=tk.DISABLED)
        self.log_widget.grid(row=23, column=0, columnspan=2, sticky="nsew")

        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(23, weight=1)

    def _current_speaker_id(self) -> int:
        label = self.speaker_combo_var.get()
        return self.speaker_label_to_id.get(label, 3)

    def _set_speaker_by_id(self, speaker_id: int) -> None:
        for label, sid in self.speaker_label_to_id.items():
            if sid == speaker_id:
                self.speaker_combo_var.set(label)
                return
        values = list(self.speaker_combo.cget("values"))
        if values:
            self.speaker_combo_var.set(values[0])

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
        self.auto_start_var.set(bool(cfg.get("auto_start_on_launch", False)))
        self.read_username_var.set(bool(cfg.get("read_username", True)))
        self.read_cw_var.set(bool(cfg.get("read_cw", False)))
        loaded_dictionary = cfg.get("dictionary_entries", [])
        self.dictionary_entries = []
        if isinstance(loaded_dictionary, list):
            for item in loaded_dictionary:
                if not isinstance(item, dict):
                    continue
                mode = str(item.get("mode", "plain")).strip().lower()
                src = str(item.get("from", "")).strip()
                dst = str(item.get("to", ""))
                if src:
                    self.dictionary_entries.append({"mode": mode, "from": src, "to": dst})

        # backward-compat: old tuple-style dictionary rules
        if not self.dictionary_entries:
            legacy_dict = cfg.get("dictionary_rules", [])
            if isinstance(legacy_dict, list):
                for item in legacy_dict:
                    if isinstance(item, (list, tuple)) and len(item) >= 2:
                        src = str(item[0]).strip()
                        dst = str(item[1])
                        if src:
                            self.dictionary_entries.append({"mode": "plain", "from": src, "to": dst})
        loaded_ng_words = cfg.get("ng_words", [])
        self.ng_words = []
        if isinstance(loaded_ng_words, list):
            for item in loaded_ng_words:
                word = str(item).strip()
                if word:
                    self.ng_words.append(word)
        loaded_muted_accounts = cfg.get("muted_accounts", [])
        self.muted_accounts = []
        if isinstance(loaded_muted_accounts, list):
            for item in loaded_muted_accounts:
                account = str(item).strip()
                if account:
                    self.muted_accounts.append(account)
        raw_accounts = cfg.get("accounts", {})
        if isinstance(raw_accounts, dict):
            self.accounts = {}
            for key, item in raw_accounts.items():
                if not isinstance(item, dict):
                    continue
                instance_url = str(item.get("instance_url", "")).strip().rstrip("/")
                access_token = str(item.get("access_token", "")).strip()
                acct = str(item.get("acct", "")).strip()
                if key and instance_url and access_token:
                    self.accounts[str(key)] = {
                        "instance_url": normalize_instance_url(instance_url),
                        "access_token": access_token,
                        "acct": acct,
                    }

        # backward-compat: old single-account format
        legacy_token = str(cfg.get("access_token", "")).strip()
        legacy_instance = normalize_instance_url(str(cfg.get("instance_url", "")).strip())
        if legacy_token and legacy_instance and not self.accounts:
            try:
                acct = verify_account(legacy_instance, legacy_token)
                key = f"@{acct} ({legacy_instance})"
                self.accounts[key] = {
                    "instance_url": legacy_instance,
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
        token = selected["access_token"]
        if not instance or not token:
            return
        try:
            acct = verify_account(instance, token)
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

    def on_account_selected(self, _event: Any = None) -> None:
        key = self.account_combo_var.get().strip()
        account = self.accounts.get(key)
        if not account:
            return
        self.instance_var.set(account["instance_url"])
        self.access_token = account["access_token"]
        acct = account.get("acct", "")
        if acct:
            self.status_var.set(f"アカウント選択: @{acct}")
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
            text="通常: 変換前=変換後 / 正規表現: re:pattern=>replace。先頭#はコメントです。",
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

        ttk.Label(frame, text="NGワード (1行1件。部分一致で投稿をスキップ)").grid(
            row=0, column=0, sticky="w"
        )
        ng_editor = scrolledtext.ScrolledText(frame, wrap=tk.WORD, height=10)
        ng_editor.grid(row=1, column=0, sticky="nsew", pady=(4, 10))
        ng_editor.insert("1.0", self._line_list_to_text(self.ng_words))

        ttk.Label(frame, text="ミュートアカウント (1行1件。acct/username/display_name)").grid(
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
        if not code:
            messagebox.showerror("ログインエラー", "認可コードを貼り付けてください。")
            return
        try:
            token = exchange_code_for_token(self.oauth_client, code)
            acct = verify_account(self.oauth_client.instance_url, token)
            self.access_token = token
            account_key = f"@{acct} ({self.oauth_client.instance_url})"
            self.accounts[account_key] = {
                "instance_url": self.oauth_client.instance_url,
                "access_token": token,
                "acct": acct,
            }
            self._refresh_account_combo(account_key)
            self.instance_var.set(self.oauth_client.instance_url)
            self._save_current_config()
            self.status_var.set(f"ログイン完了: @{acct}")
            self.log(f"ログイン完了: @{acct}")
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
            return
        try:
            speech_rate = float(self.speech_rate_var.get().strip())
            speech_volume = float(self.speech_volume_var.get().strip())
            speech_pitch = float(self.speech_pitch_var.get().strip())
        except ValueError:
            messagebox.showerror("入力エラー", "読み上げ速度・音量・ピッチは数値で入力してください。")
            return
        try:
            omit_long_threshold = int(self.omit_long_threshold_var.get().strip())
        except ValueError:
            messagebox.showerror("入力エラー", "長文省略しきい値は整数で入力してください。")
            return
        omit_long_threshold = max(1, omit_long_threshold)
        self.omit_long_threshold_var.set(str(omit_long_threshold))

        timeline_kind = timeline_label_to_kind(self.timeline_kind_var.get().strip())
        if timeline_kind not in ("local", "home"):
            timeline_kind = "local"
            self.timeline_kind_var.set(timeline_kind_to_label(timeline_kind))

        speaker_id = self._current_speaker_id()
        self.worker = TimelineSpeaker(
            instance_url=instance,
            access_token=selected["access_token"],
            voicevox_url=voicevox_url if use_voicevox else "",
            speaker_id=speaker_id,
            poll_interval_sec=max(3, poll_interval),
            fetch_limit=max(1, min(40, fetch_limit)),
            timeline_kind=timeline_kind,
            speech_rate=speech_rate,
            speech_volume=speech_volume,
            speech_pitch=speech_pitch,
            omit_long_threshold=omit_long_threshold,
            dictionary_entries=[dict(x) for x in self.dictionary_entries],
            ng_words=list(self.ng_words),
            muted_accounts=list(self.muted_accounts),
            skip_boosts=self.skip_boosts_var.get(),
            skip_replies=self.skip_replies_var.get(),
            omit_body_when_cw=self.omit_body_when_cw_var.get(),
            read_username=self.read_username_var.get(),
            read_cw=self.read_cw_var.get(),
            logger=self.log,
        )
        self.worker_thread = threading.Thread(target=self.worker.run, daemon=True)
        self.worker_thread.start()
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self.status_var.set("読み上げを開始しました。")
        self._save_current_config()
        if use_voicevox:
            self.log(f"読み上げワーカー起動 (VOICEVOX, 話者ID={speaker_id})。")
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


def main() -> None:
    root = tk.Tk()
    apply_window_icon(root)
    app = App(root)

    def on_close() -> None:
        if app.worker:
            app.worker.stop()
        app._save_current_config()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.mainloop()


if __name__ == "__main__":
    main()

