# -*- coding: utf-8 -*-
"""sessionStart: NDJSON デバッグ計測の正本 API と WSL ミラー規約をエージェントへ注入。"""

from __future__ import annotations

import json
import sys


def main() -> None:
    sys.stdin.read()
    text = (
        "【Cursor デバッグ NDJSON（計測追加・ログ取得時は必須）】\n"
        "正本ルール: .cursor/rules/agent-debug-ndjson-logging.mdc 、"
        ".cursor/rules/agent-debug-wsl-windows-mirror.mdc 、"
        ".cursor/rules/agent-debug-hooks-persistence.mdc\n"
        "Java: AgentDebugLog.appendStructured(shell.snapshotUiEnv(), sessionId, …) を "
        "// #region agent log … // #endregion で囲む。禁止: ソースに /mnt/c/... や OS 固定絶対パスでログファイル直書き。\n"
        "Python 子プロセス: planning_core.agent_debug_ndjson.append_structured(...) のみ。"
        "独自の _agent_debug_log_* や debug-<id>.log への open 直書きは作らない。\n"
        "sessionId はチャット指定の debug-<sessionId>.log に揃える。"
        "親 JVM が PM_AI_DEBUG_LOG / PM_AI_AGENT_DEBUG_SESSION を渡す配台試行ではそれに従う。\n"
        "WSL ワークスペースでログが見えないときは Java 側 AgentDebugLog の UNC ミラー"
        "（PM_AI_WSL_DISTRO / PM_AI_DEBUG_LOG_WSL_UNC）を疑う。修正確証後は #region agent log ブロックを削除。\n"
    )
    sys.stdout.write(json.dumps({"additional_context": text}, ensure_ascii=False))


if __name__ == "__main__":
    main()
