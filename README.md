# FediverseVoiceReader

Fediverse（Mastodon系）のタイムラインを取得し、VOICEVOX または Windows 標準読み上げで音声化する Windows 向けアプリです。

## 配布方針

- リポジトリ: ソースコードを公開
- GitHub Releases: `FediverseVoiceReader.exe` を配布

## 主な機能

- OAuthログイン（Fediverseアカウント）
- 複数アカウント保存・選択
- ローカル / ホームタイムライン切替
- VOICEVOX話者一覧取得・選択
- VOICEVOX未設定/未接続時は Windows 標準読み上げへ自動切替
- 設定でURL入り投稿自体を除外可能
- `#` を含む投稿を読み上げ対象から除外可能
- 絵文字（Unicode / カスタム絵文字）を読み上げ対象から除外
- ブースト除外 / 返信除外 / CW時本文省略
- 読み上げ速度・音量・ピッチ調整
- 辞書置換（通常 / 正規表現）
- ミュート/NG設定
- 長文しきい値による「以下省略」
- 引用投稿の通知読み上げ
- 起動時自動読み上げ
- ログ保存/エクスポート
- 作業用BGMとしてMP3ファイルを再生可能（ループ再生対応）

## 利用者向け要件（.exe版）

- 必須: Windows 10/11
- 必須: インターネット接続
- 任意: VOICEVOX（未導入でも Windows 標準読み上げで動作）
- 不要: Python

## 開発者向け要件（ソース実行）

- 必須: Python 3.12（`tkinter` 同梱版）
- 必須: `pip install -r requirements.txt`
- 任意: VOICEVOX

## ローカル実行（ソース）

```powershell
cd "C:\path\to\Fediverse読み上げ"
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py
```

## EXEビルド（PyInstaller）

```powershell
cd "C:\path\to\Fediverse読み上げ"
.\.venv\Scripts\python.exe -m pip install pyinstaller
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile --windowed --name FediverseVoiceReader --icon "FediverseVoiceReader.ico" --add-data "FediverseVoiceReader.ico;." --add-data "FediverseVoiceReader.png;." main.py
```

出力:

- `dist\FediverseVoiceReader.exe`

## 使い方（.exe版）

1. `FediverseVoiceReader.exe` を起動
2. `1) VOICEVOX確認` を実行
3. 必要なら `話者一覧を更新` で話者を選択
4. `2) ログイン開始` でブラウザ認証し、認可コードを貼り付けて `ログイン完了`
5. アカウントを選択
6. 必要に応じて読み上げ設定を調整
7. `3) 読み上げ開始`

## 保存先と機密情報

- 設定: `%APPDATA%\FediverseVoiceReader\config.json`
- ログ: `%APPDATA%\FediverseVoiceReader\logs\`

`config.json` にはアクセストークンが含まれるため、公開しないでください。

## Windows の警告（SmartScreen）について

本アプリは現在 **コード署名なし** で配布しています。  
そのため、初回実行時に Windows Defender SmartScreen の警告が表示される場合があります。

### 実行手順（警告が出た場合）

1. `FediverseVoiceReader.exe` を起動  
2. 「Windows によって PC が保護されました」と表示されたら **`詳細情報`** をクリック  
3. 表示された **`実行`** ボタンをクリック

### 注意

- 配布元はこのリポジトリ / Releases のみです  
- 不審な改変版を避けるため、必ず公式ページからダウンロードしてください  
- 心配な場合は、ソースコードから自分でビルドして利用してください

## クレジット・ライセンス

このソフトは音声合成に VOICEVOX を利用します。  
VOICEVOX本体・VOICEVOX ENGINE は同梱していません（ユーザー環境のものを使用）。

- VOICEVOX: https://voicevox.hiroshiba.jp/
- VOICEVOX ENGINE: https://github.com/VOICEVOX/voicevox_engine

音声・キャラクターの利用条件は、VOICEVOXおよび各キャラクターの規約に従ってください。
