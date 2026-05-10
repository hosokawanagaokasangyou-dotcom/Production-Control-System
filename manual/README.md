# マニュアル HTML（静的生成）

## 前提

- Python 3.10 以上
- 連続スクショを含むフルパイプラインは **Windows デスクトップ**で実行すること（JavaFX が表示できること）

## セットアップ

推奨（Windows PowerShell、リポジトリ直下）:

```text
python -m venv manual\.venv
manual\.venv\Scripts\pip install -r manual\requirements.txt
```

またはグローバル環境へ:

```text
pip install -r manual/requirements.txt
```

## 一括実行（推奨）

リポジトリ直下で:

```text
.\Publish-Manual.ps1
```

生成物は `manual/html/`（ブラウザで `manual/html/index.html` を開く）。

## 設定

`manual/pipeline-manifest.yaml` で撮影タブ（`MainShellTabId.key()`）・注入先 Markdown を編集する。

## フェーズだけ省略する場合

```text
.\Publish-Manual.ps1 -SkipSnap
```

など（`-SkipDataPrep` `-SkipSnap` `-SkipSync` `-SkipInject` `-SkipHtml`）。
