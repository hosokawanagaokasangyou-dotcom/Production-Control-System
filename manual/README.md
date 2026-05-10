# マニュアル HTML の生成方法

## 何が起きるか

1. **画像** … `manual/src/images/` に、`tab_key` 名の PNG を置く（例: `run.png`, `env.png`）。アプリの該当画面を手動でスクショして保存する。
2. **公開パイプライン** … `Publish-Manual.ps1`（または同等の Python）が次を順に実行する。
   - （任意）`data_prep` でファイルコピー
   - **inject** … Markdown 内の `<!-- MANUAL_SNAP:キー -->` を画像参照に置換
   - **HTML** … `manual/src` 以下の Markdown から `manual/html/` に静的サイトを出力

完成したらブラウザで **`manual/html/index.html`** を開く。

## 前提

- Python **3.10 以上**
- 依存: `pip install -r manual/requirements.txt`（`markdown`, `PyYAML`）

## Windows（PowerShell）— リポジトリ直下

初回だけ仮想環境の例:

```text
python -m venv manual\.venv
manual\.venv\Scripts\pip install -r manual\requirements.txt
```

生成:

```text
.\Publish-Manual.ps1
```

## WSL / Linux — `pip install` が externally-managed で拒まれるとき

Ubuntu 等ではシステムの Python に直接 `pip install` できません（PEP 668）。**必ず venv を使う**。

```text
python3 -m venv manual/.venv
source manual/.venv/bin/activate
pip install -r manual/requirements.txt
python3 scripts/manual_publish.py
```

（Windows と同じく、`manual/.venv` を使うなら `Publish-Manual.ps1` は使わず上記 `python3` で問題ありません。）

## 設定の場所

- **`manual/pipeline-manifest.yaml`** … `injections` でプレースホルダと Markdown を対応付ける。
- **`manual/src/chapters/*.md`** … `<!-- MANUAL_SNAP:run -->` などのマーカー。

## トラブル: `SyntaxError: Non-UTF-8 code ... manual_publish.py`

リポジトリのスクリプトは **UTF-8** です。エディタを「UTF-8 で保存」に固定してください。`/mnt/c` 経由で CP932 保存すると文字化けし、Python が起動できなくなります。

## inject について

初回はプレースホルダが画像参照に書き換わります。マーカーを残したい場合は Git で元に戻すか、マーカーを手で再度書いてください。
