# マニュアル HTML（静的生成）

## 前提

- Python 3.10 以上
- 画面は **手動**でキャプチャし、`manual/src/images` に **PNG** を置く（ファイル名は `pipeline-manifest.yaml` の `injections[].tab_key` と一致、例: `run.png`）。

## セットアップ

例（Windows PowerShell、リポジトリ直下）:

```text
python -m venv manual\.venv
manual\.venv\Scripts\pip install -r manual\requirements.txt
```

またはグローバル環境へ:

```text
pip install -r manual/requirements.txt
```

## 一括生成（通常）

リポジトリ直下で:

```text
.\Publish-Manual.ps1
```

出力は `manual/html/`（ブラウザで `manual/html/index.html` を開く）。

## 設定

`manual/pipeline-manifest.yaml` でプレースホルダ `<!-- MANUAL_SNAP:key -->` を画像参照へ置換する対応付けを行う（`tab_key` は `MainShellTabId.key()` に合わせる）。

## フェーズだけ個別にスキップする場合

```text
.\Publish-Manual.ps1 -SkipDataPrep
```

など。`-SkipDataPrep` `-SkipInject` `-SkipHtml`。
