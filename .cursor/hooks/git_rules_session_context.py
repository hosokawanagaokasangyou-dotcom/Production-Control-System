# -*- coding: utf-8 -*-
"""sessionStart hook: inject repo Git workflow summary into agent context."""

import json
import sys


def main() -> None:
    sys.stdin.read()
    text = (
        "【このリポジトリの Git 運用（エージェント向け・必須の要約）】\n"
        "版管理対象を変更したターンでは .cursor/rules/git-commit-push-after-code-changes.mdc に従い、"
        "その依頼で触ったファイルを git add → commit → push し、応答で結果をユーザーに報告する。"
        "依頼外の差分が混在する場合はコミットを分けるかユーザーに確認。編集前のローカルコミットは不要。"
        "AI がコミット文を生成する場合は日本語（.cursor/rules/git-commit-message-japanese.mdc）。\n"
    )
    sys.stdout.write(json.dumps({"additional_context": text}, ensure_ascii=False))


if __name__ == "__main__":
    main()
