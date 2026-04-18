# -*- coding: utf-8 -*-
"""sessionStart hook: inject repo Git workflow summary into agent context."""

import json
import sys


def main() -> None:
    sys.stdin.read()
    text = (
        "【このリポジトリの Git 運用（エージェント向け・必須の要約）】\n"
        "版管理対象（ソース、*.md / *.html、設定、ルール等）を編集ツールで変えるターンでは、"
        "編集の前に必ず git status。未コミットがあれば .cursor/rules/pre-edit-git-commit-push.mdc "
        "の手順でローカルコミット（メッセージ先頭は固定の "
        "「(下記、変更を加える前の直前コミットです)」＋そのターンのユーザー入力先頭300文字相当・"
        "改行タブは半角スペース）→ コミットしたら git push → 結果をユーザーに報告。"
        "作業ツリーがクリーンならコミット省略。無関係な変更が多いときは git add の範囲に注意。"
        "AI が別途コミット文を生成する場合は日本語（.cursor/rules/git-commit-message-japanese.mdc）。\n"
    )
    sys.stdout.write(json.dumps({"additional_context": text}, ensure_ascii=False))


if __name__ == "__main__":
    main()
