# -*- coding: utf-8 -*-
"""sessionStart hook: inject repo Git workflow summary into agent context."""

import json
import sys


def main() -> None:
    sys.stdin.read()
    text = (
        "【このリポジトリの Git 運用（エージェント向け・必須の要約）】\n"
        "正本: .cursor/rules/git-commit-push-after-code-changes.mdc\n"
        "関連: pm-ai-cache-network-source-tracking.mdc 、version-txt-bump-on-commit.mdc\n"
        "フック一覧: .cursor/rules/git-commit-hooks-persistence.mdc（.cursor/hooks.json）\n"
        "版管理対象を変更したターンでは git add → commit → push し、応答で結果をユーザーに報告する。"
        "依頼外の差分が混在する場合はコミットを分けるかユーザーに確認。編集前のローカルコミットは不要。"
        "AI がコミット文を生成する場合は日本語のみ（規約は正本「コミットメッセージ規約」節）。"
        "コミット前に core.hooksPath=scripts/git-hooks を確認。version.txt は pre-commit で +0.01（version-txt-bump 参照）。"
        ".pm-ai-cache/ は追跡対象。Write / StrReplace / エディタ直編集後は git_commit_post_edit_reminder が注意喚起。\n"
    )
    sys.stdout.write(json.dumps({"additional_context": text}, ensure_ascii=False))


if __name__ == "__main__":
    main()
