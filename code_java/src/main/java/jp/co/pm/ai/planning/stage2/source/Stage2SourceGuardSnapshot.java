package jp.co.pm.ai.planning.stage2.source;

import java.util.Map;

/** 非同期固定ソースガード開始時の実行条件。値は比較にのみ使い、ログへ出さない。 */
public record Stage2SourceGuardSnapshot(
        boolean todayDispatch,
        boolean planInputDirty,
        long planInputDirtyGeneration,
        boolean dispatchTableDirty,
        long dispatchTableDirtyGeneration,
        boolean pipelineRunning,
        Map<String, String> environment) {

    public Stage2SourceGuardSnapshot {
        environment = environment == null ? Map.of() : Map.copyOf(environment);
    }

    public boolean matches(Stage2SourceGuardSnapshot current) {
        return current != null
                && todayDispatch == current.todayDispatch
                && planInputDirty == current.planInputDirty
                && planInputDirtyGeneration == current.planInputDirtyGeneration
                && dispatchTableDirty == current.dispatchTableDirty
                && dispatchTableDirtyGeneration == current.dispatchTableDirtyGeneration
                && pipelineRunning == current.pipelineRunning
                && environment.equals(current.environment);
    }

    public String mismatchMessage(Stage2SourceGuardSnapshot current) {
        if (current == null) {
            return "実行条件を再取得できませんでした。";
        }
        if (todayDispatch != current.todayDispatch) {
            return "当日配台の設定が変更されました。";
        }
        if (planInputDirty != current.planInputDirty
                || planInputDirtyGeneration != current.planInputDirtyGeneration
                || dispatchTableDirty != current.dispatchTableDirty
                || dispatchTableDirtyGeneration != current.dispatchTableDirtyGeneration) {
            return "入力表の保存状態が変更されました。";
        }
        if (pipelineRunning != current.pipelineRunning) {
            return "別の段階処理が開始されました。";
        }
        if (!environment.equals(current.environment)) {
            return "実行環境が変更されました。";
        }
        return "";
    }
}
