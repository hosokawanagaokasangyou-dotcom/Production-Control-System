package jp.co.pm.ai.planning.stage2;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.function.Consumer;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * 段階2 Java エンジンへの入力。{@code collectUiEnv()} 互換マップと、マスタ解決に使うタスク入力ブックパス、ログシンクを保持する。
 */
public final class Stage2RunContext {

    private final Map<String, String> uiEnv;
    private final String taskInputWorkbookPath;
    private final Consumer<String> logLine;

    public Stage2RunContext(
            Map<String, String> uiEnv, String taskInputWorkbookPath, Consumer<String> logLine) {
        this.uiEnv = Map.copyOf(uiEnv != null ? new HashMap<>(uiEnv) : Map.of());
        this.taskInputWorkbookPath = taskInputWorkbookPath != null ? taskInputWorkbookPath.strip() : "";
        this.logLine = logLine != null ? logLine : (s -> {});
    }

    public Map<String, String> uiEnv() {
        return uiEnv;
    }

    public String taskInputWorkbookPath() {
        return taskInputWorkbookPath;
    }

    public void log(String line) {
        logLine.accept(Objects.requireNonNullElse(line, ""));
    }

    public boolean stage2WriteExcel() {
        return Stage2EnvParsing.stage2WriteExcel(uiEnv);
    }

    public boolean mirrorPlanWorkbookJson() {
        return Stage2EnvParsing.envEnabled(AppPaths.KEY_PM_AI_PLAN_WORKBOOK_JSON, uiEnv, true);
    }

    public boolean mirrorMemberScheduleJson() {
        return Stage2EnvParsing.envEnabled(AppPaths.KEY_PM_AI_MEMBER_SCHEDULE_JSON, uiEnv, true);
    }
}
