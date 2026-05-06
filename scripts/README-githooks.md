# Git フック（version.txt 自動バンプ）

リポジトリ直下の `version.txt` を、**コミットのたびに +0.01** する `pre-commit` を同梱しています。

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

Python 3 と `scripts/bump_version_txt.py` が必要です。
