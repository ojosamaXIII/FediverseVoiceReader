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
import wave
from datetime import datetime
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable
from urllib.parse import urlencode, urlsplit

import requests


def ensure_tk_env() -> None:
    # Some Windows environments fail to locate Tcl/Tk when launched from
    # non-ASCII working directories. Set explicit library paths early.
    base_prefix = Path(sys.base_prefix)
    tcl_root = base_prefix / "tcl"
    tcl_lib = tcl_root / "tcl8.6"
    tk_lib = tcl_root / "tk8.6"
    dll_dir = base_prefix / "DLLs"
    if "TCL_LIBRARY" not in os.environ and tcl_lib.exists():
        os.environ["TCL_LIBRARY"] = str(tcl_lib)
    if "TK_LIBRARY" not in os.environ and tk_lib.exists():
        os.environ["TK_LIBRARY"] = str(tk_lib)
    if dll_dir.exists():
        os.environ["PATH"] = f"{dll_dir};{os.environ.get('PATH', '')}"


ensure_tk_env()

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


APP_DIR_NAME = "FediverseVoiceReader"
TAG_RE = re.compile(r"<[^>]+>")
SPACE_RE = re.compile(r"\s+")
URL_RE = re.compile(r"https?://[^\s]+", flags=re.IGNORECASE)
URL_LIKE_RE = re.compile(r"\b(?:www\.)[^\s]+", flags=re.IGNORECASE)
URL_LEADING_LINE_RE = re.compile(
    r"^\s*(?:https?://|www\.)\S.*$",
    flags=re.IGNORECASE | re.MULTILINE,
)
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


def contains_url_like(text: str) -> bool:
    if not text:
        return False
    return bool(URL_RE.search(text) or URL_LIKE_RE.search(text))


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
    backend: str
    client_id: str
    client_secret: str
    session_token: str = ""


@dataclass
class ReaderStartSettings:
    voicevox_url: str
    use_voicevox: bool
    speaker_id: int
    poll_interval_sec: int
    fetch_limit: int
    timeline_kind: str
    speech_rate: float
    speech_volume: float
    speech_pitch: float
    omit_long_threshold: int


def detect_backend(instance_url: str) -> str:
    base = normalize_instance_url(instance_url).rstrip("/")
    if not base:
        return "mastodon"
    try:
        res = requests.post(
            f"{base}/api/meta",
            json={"detail": False},
            headers={"Accept": "application/json", "User-Agent": HTTP_USER_AGENT},
            timeout=10,
        )
        if res.ok:
            payload = res.json()
            if isinstance(payload, dict) and "maintainerName" in payload:
                return "misskey"
    except requests.RequestException:
        pass
    return "mastodon"


def register_mastodon_app(instance_url: str) -> OAuthClient:
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
        backend="mastodon",
        client_id=payload["client_id"],
        client_secret=payload["client_secret"],
    )


def register_misskey_app(instance_url: str) -> OAuthClient:
    base = normalize_instance_url(instance_url).rstrip("/")
    if not base:
        raise ValueError("インスタンスURLが空です。")
    app_res = requests.post(
        f"{base}/api/app/create",
        json={
            "name": "Fediverse Timeline Reader",
            "description": "Read timeline by TTS",
            "permission": ["read:account", "read:notes"],
        },
        headers={"Accept": "application/json", "User-Agent": HTTP_USER_AGENT},
        timeout=20,
    )
    app_res.raise_for_status()
    app_payload = app_res.json()
    app_secret = str(app_payload.get("secret", "")).strip()
    if not app_secret:
        raise requests.RequestException("Misskeyアプリ作成に失敗しました。secretが取得できません。")

    sess_res = requests.post(
        f"{base}/api/auth/session/generate",
        json={"appSecret": app_secret},
        headers={"Accept": "application/json", "User-Agent": HTTP_USER_AGENT},
        timeout=20,
    )
    sess_res.raise_for_status()
    sess_payload = sess_res.json()
    session_token = str(sess_payload.get("token", "")).strip()
    authorize_url = str(sess_payload.get("url", "")).strip()
    if not session_token or not authorize_url:
        raise requests.RequestException("Misskey認可セッション生成に失敗しました。")

    return OAuthClient(
        instance_url=base,
        backend="misskey",
        client_id=authorize_url,
        client_secret=app_secret,
        session_token=session_token,
    )


def register_app(instance_url: str) -> OAuthClient:
    backend = detect_backend(instance_url)
    if backend == "misskey":
        return register_misskey_app(instance_url)
    return register_mastodon_app(instance_url)


def build_authorize_url(client: OAuthClient) -> str:
    if client.backend == "misskey":
        return client.client_id
    params = {
        "client_id": client.client_id,
        "scope": "read",
        "redirect_uri": REDIRECT_URI,
        "response_type": "code",
    }
    return f"{client.instance_url}/oauth/authorize?{urlencode(params)}"


def exchange_code_for_token(client: OAuthClient, code: str) -> str:
    if client.backend == "misskey":
        if not client.session_token:
            raise requests.RequestException("Misskey認可セッション情報がありません。")
        res = requests.post(
            f"{client.instance_url}/api/auth/session/userkey",
            json={"appSecret": client.client_secret, "token": client.session_token},
            headers={"Accept": "application/json", "User-Agent": HTTP_USER_AGENT},
            timeout=20,
        )
        res.raise_for_status()
        payload = res.json()
        access_token = str(payload.get("accessToken", "")).strip()
        if not access_token:
            raise requests.RequestException("Misskeyアクセストークン取得に失敗しました。")
        return access_token

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
    return verify_account_with_backend(instance_url, access_token, "mastodon")


def verify_account_with_backend(instance_url: str, access_token: str, backend: str) -> str:
    if backend == "misskey":
        res = requests.post(
            f"{instance_url.rstrip('/')}/api/i",
            json={"i": access_token},
            headers={"Accept": "application/json", "User-Agent": HTTP_USER_AGENT},
            timeout=20,
        )
        res.raise_for_status()
        account = res.json()
        username = str(account.get("username", "")).strip()
        host = str(account.get("host", "")).strip()
        if username and host:
            return f"{username}@{host}"
        return username or "unknown"

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


