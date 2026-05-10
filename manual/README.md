# マニュアル HTML の生成方法

## 何が起きるか

1. **画像** … `manual/src/images/` に、`tab_key` 名の PNG を置く（例: `run.png`, `env.png`）。アプリの該当画面を手動でスクショして保存する。
2. **公開パイプライン** … `Publish-Manual.ps1`（または同等の Python）が次を順に実行する。
   - （任意）`data_prep` でファイルコピー
   - **inject** … Markdown 内の `<!-- MANUAL_SNAP:キー -->` を `![](../images/キー.png)` に置き換え
   - **HTML** … `manual/src` 以下の Markdown から `manual/html/` に静的サイトを出力

完成したらブラウザで **`manual/html/index.html`** を開く。

## 前提

- Python **3.10 以上**
- 依存: `pip install -r manual/requirements.txt`（`markdown`, `PyYAML`）

## Windows（PowerShell）— リポジトリ直下で

初回だけ仮想環境の例:

```text
python -m venv manual\.venv
manual\.venv\Scripts\pip install -r manual\requirements.txt
```

生成:

```text
.\Publish-Manual.ps1
```

一部だけスキップする例:

```text
.\Publish-Manual.ps1 -SkipDataPrep
```

（`-SkipInject` … 置換を飛ばす／`-SkipHtml` … HTML 生成だけ飛ばす）

## WSL / Linux（bash）— リポジトリ直下で

```text
pip install -r manual/requirements.txt
python3 scripts/manual_publish.py
```

マニフェストを変える場合:

```text
python3 scripts/manual_publish.py --manifest manual/pipeline-manifest.yaml
```

## 設定の場所

- **`manual/pipeline-manifest.yaml`** … どの Markdown のどのプレースホルダを、どのキャプションで画像にするか（`injections`）。
- **`manual/src/chapters/*.md`** … 本文と `<!-- MANUAL_SNAP:run -->` などのマーカー。

## inject について

初回はプレースホルダが **`![](画像)` に書き換わります**。すでに置換済みでマーカーが無い章は、inject はその行をスキップします（再びマーカーを書けば再度置換対象になります）。
