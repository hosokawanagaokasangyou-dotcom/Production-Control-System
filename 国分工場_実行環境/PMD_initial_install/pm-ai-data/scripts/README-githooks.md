# Git フック（Markdown UTF-8 検証と version.txt 自動バンプ）

`pre-commit` は次を実行します。

1. **ステージされた `*.md` が UTF-8 として解釈できるか**（`scripts/check_staged_markdown_utf8.py`）。Shift-JIS 等の場合はコミットを拒否する。
2. リポジトリ直下の **`version.txt` をコミットのたびに +0.01** する。

## 有効化（各クローンで一度）

リポジトリルートで:

```bash
git config core.hooksPath scripts/git-hooks
```

`chmod +x scripts/git-hooks/pre-commit` が必要な環境では実行権限を付与してください。

## 無効化

```bash
git config --unset core.hooksPath
```

## スキップしてコミット（フックを動かさない）

```bash
git commit --no-verify ...
```

Python 3 と `scripts/bump_version_txt.py` / `scripts/check_staged_markdown_utf8.py` が必要です。
