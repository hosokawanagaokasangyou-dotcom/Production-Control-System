package jp.co.pm.ai.planning.stage2.input;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.debug.AgentDebugLog;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.SkillsSheetMemberReader;
import jp.co.pm.ai.planning.stage2.Stage2RunContext;

/** 段階2 Java エンジン用の入力解決（マスタ・計画入力・除外 JSON）。 */
public final class Stage2InputLoader {

    public static final String ENV_TASK_PLAN_SHEET = "TASK_PLAN_SHEET";
    private static final String DEFAULT_PLAN_SHEET = "配台計画_タスク入力";

    private Stage2InputLoader() {}

    public static Stage2InputSnapshot load(Stage2RunContext ctx) throws IOException {
        Map<String, String> ui = ctx.uiEnv();
        Path master =
                AppPaths.resolveMasterWorkbookPathResolved(ui, ctx.taskInputWorkbookPath()).normalize();
        if (!Files.isRegularFile(master)) {
            throw new IOException("マスタブックが見つかりません: " + master);
        }
        List<String> members = SkillsSheetMemberReader.readMemberDisplayNames(master);
        Stage2MasterFactoryHoursReader.FactoryHours fh =
                Stage2MasterFactoryHoursReader.read(master);

        Path excludePath = resolveExcludeRulesPath(ui);
        int excludeCount = Stage2ExcludeRulesReader.countRules(excludePath);

        String planPathStr = trim(ui.get(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH));
        if (planPathStr.isEmpty()) {
            throw new IOException("PM_AI_PLAN_INPUT_PATH が空です。");
        }
        Path planPath = Path.of(planPathStr).toAbsolutePath().normalize();
        if (!Files.isRegularFile(planPath)) {
            throw new IOException("配台計画入力が見つかりません: " + planPath);
        }
        String sheet = trim(ui.get(ENV_TASK_PLAN_SHEET));
        if (sheet.isEmpty()) {
            sheet = DEFAULT_PLAN_SHEET;
        }
        // #region agent log
        try {
            LinkedHashMap<String, Object> d = new LinkedHashMap<>();
            d.put("envTaskPlanSheetRaw", ui.get(ENV_TASK_PLAN_SHEET));
            d.put("resolvedSheet", sheet);
            d.put("looksLikeHaigoTypo", sheet.contains("配合") && !sheet.contains("配台"));
            d.put("equalsOfficialDefault", DEFAULT_PLAN_SHEET.equals(sheet));
            AgentDebugLog.appendStructured(
                    ui,
                    "b59b51",
                    "H1-H3",
                    "Stage2InputLoader.load",
                    "task plan sheet before PlanInputTabularIo.read",
                    d);
        } catch (Throwable ignored) {
            // debug-only
        }
        // #endregion
        PlanInputTabularIo.TabularSheet tab = PlanInputTabularIo.read(planPath, sheet);

        return new Stage2InputSnapshot(
                master,
                members,
                fh.start(),
                fh.end(),
                excludeCount,
                planPath,
                sheet,
                tab);
    }

    private static Path resolveExcludeRulesPath(Map<String, String> ui) {
        String p = trim(ui.get(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON));
        if (!p.isEmpty()) {
            Path ap = Path.of(p).toAbsolutePath().normalize();
            if (Files.isRegularFile(ap)) {
                return ap;
            }
        }
        return AppPaths.resolveDefaultExcludeRulesJsonPath(ui).orElse(null);
    }

    private static String trim(String s) {
        return s != null ? s.strip() : "";
    }

    public static void logSummary(Stage2RunContext ctx, Stage2InputSnapshot snap) {
        ctx.log(
                "[stage2-java] 入力: master="
                        + snap.masterPath()
                        + " | members="
                        + snap.memberDisplayNames().size()
                        + " | exclude_rules="
                        + snap.excludeRuleCount()
                        + " ("
                        + Stage2ExcludeRulesReader.fileLabel(resolveExcludeRulesPath(ctx.uiEnv()))
                        + ")"
                        + " | plan_rows="
                        + snap.planningTasksSheet().rows().size()
                        + " | plan_sheet="
                        + snap.planSheetName());
        if (snap.factoryStart().isPresent() && snap.factoryEnd().isPresent()) {
            ctx.log(
                    "[stage2-java] 工場稼働枠(master メイン A12/B12): "
                            + snap.factoryStart().get()
                            + " ～ "
                            + snap.factoryEnd().get());
        } else {
            ctx.log("[stage2-java] 工場稼働枠: master メイン A12/B12 は未採用（既定日内枠を想定）。");
        }
    }
}
