package jp.co.pm.ai.planning.stage2.parity;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import java.util.function.Consumer;

import jp.co.pm.ai.desktop.PlanInputTabController;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner;
import jp.co.pm.ai.desktop.bridge.PythonProcessRunner.RunRequest;
import jp.co.pm.ai.desktop.bridge.StagePythonExecutable;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.Stage2OutputNaming;
import jp.co.pm.ai.desktop.ipc.IpcStdoutTap;
import jp.co.pm.ai.planning.stage2.Stage2EnvParsing;
import jp.co.pm.ai.planning.stage2.Stage2ExitCodes;
import jp.co.pm.ai.planning.stage2.Stage2JavaEngine;
import jp.co.pm.ai.planning.stage2.Stage2RunContext;

/**
 * JavaFX を介さず Python 段階2 → Java 段階2を順に実行し、成果物を {@link Stage2ParityBundle} でまとめる。
 *
 * <p>呼び出し元は {@link jp.co.pm.ai.desktop.MainShellController#childEnvForPython} と同等のマージ済み env（{@code
 * childBase}）を渡すこと。ヘッドレス CLI は {@link jp.co.pm.ai.desktop.bridge.Stage2PythonChildEnv#headlessBaseFromSystemEnv()} で近似する。
 */
public final class Stage2HeadlessParityRunner {

    private static final String SCRIPT_STAGE2 = "plan_simulation_stage2.py";

    private Stage2HeadlessParityRunner() {}

    public record Outcome(Throwable fatalError, Stage2ParityBundle bundle) {

        public static Outcome fatal(Throwable t) {
            return new Outcome(t, null);
        }

        public static Outcome compared(Stage2ParityBundle b) {
            return new Outcome(null, b);
        }
    }

    /**
     * @param childBaseMerged {@code childEnvForPython} 相当（ネットワークソース解決済み）
     * @param taskInputWorkbookPath 段階2 Python 子へ渡すタスク入力ブックパス（UI と同じ解決値）
     * @param scriptDirFallback {@code PM_AI_CODE_PYTHON_DIR} が空のときの cwd 候補
     * @param uiTabularSnapshot 非 null なら配台計画タブとファイルの照合。null ならファイルを二重読みして常に一致扱い（CLI 用）
     */
    public static Outcome run(
            Map<String, String> childBaseMerged,
            String taskInputWorkbookPath,
            String scriptDirFallback,
            PlanInputTabularIo.TabularSheet uiTabularSnapshot,
            Consumer<String> log) {
        Objects.requireNonNull(childBaseMerged, "childBaseMerged");
        Consumer<String> lg = log != null ? log : s -> {};
        try {
            String parityPlanPath =
                    firstNonBlank(childBaseMerged.get(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH), "").strip();
            if (parityPlanPath.isEmpty()) {
                return Outcome.fatal(
                        new IllegalStateException(
                                "PM_AI_PLAN_INPUT_PATH が空です。環境でタスク入力パスを設定してください。"));
            }
            Path outDir = AppPaths.defaultPlanningOutputDir(childBaseMerged);
            String paritySheet =
                    firstNonBlank(
                            childBaseMerged.get(PlanInputTabController.ENV_TASK_PLAN_SHEET),
                            PlanInputTabController.DEFAULT_PLAN_INPUT_SHEET_NAME);
            lg.accept("[stage2-parity] 入力照合対象（子 env の PM_AI_PLAN_INPUT_PATH）: " + parityPlanPath);

            Stage2ParityCheckResult planInputCmp;
            Path planPath = Path.of(parityPlanPath);
            if (uiTabularSnapshot != null) {
                planInputCmp =
                        Stage2PlanInputUiParity.compareUiToDisk(uiTabularSnapshot, planPath, paritySheet);
            } else {
                PlanInputTabularIo.TabularRead disk =
                        PlanInputTabularIo.readWithResolvedSheet(planPath, paritySheet);
                planInputCmp =
                        Stage2PlanInputUiParity.compareUiToDisk(disk.tabular(), planPath, paritySheet);
            }
            lg.accept(
                    planInputCmp.identical()
                            ? "[stage2-parity] 配台計画タブの表と入力ファイル: 一致（またはヘッドレス照合）"
                            : "[stage2-parity] 配台計画タブの表と入力ファイル: 不一致（Python/Java は入力ファイルを読みます）");

            long floorBeforePy = Stage2OutputNaming.maxPrimaryPlanJsonLastModifiedMillis(outDir);
            long floorMemJsonPy = Stage2OutputNaming.maxPrimaryMemberJsonLastModifiedMillis(outDir);
            long floorPlanXlsxPy = Stage2OutputNaming.maxPrimaryPlanXlsxLastModifiedMillis(outDir);
            long floorMemXlsxPy = Stage2OutputNaming.maxPrimaryMemberXlsxLastModifiedMillis(outDir);

            Map<String, String> childPy = new HashMap<>(childBaseMerged);
            childPy.put(AppPaths.KEY_PM_AI_STAGE2_ENGINE, "python");
            lg.accept("[stage2-parity] (1) Python 段階2 を実行します…");
            int pyCode = runPythonStage2Blocking(childPy, taskInputWorkbookPath, scriptDirFallback, lg);
            lg.accept("[stage2-parity] Python exitCode=" + pyCode + " " + Stage2ExitCodes.hint(pyCode));
            if (pyCode != 0) {
                return Outcome.fatal(
                        new IllegalStateException(
                                "Python 段階2が失敗したため比較を中止しました (exit=" + pyCode + ")."));
            }
            Path pyJson = Stage2OutputNaming.newestPrimaryPlanJsonAfter(outDir, floorBeforePy);
            if (pyJson == null || !Files.isRegularFile(pyJson)) {
                return Outcome.fatal(
                        new IOException("Python 実行後も新しい計画 JSON が見つかりません: " + outDir));
            }
            lg.accept("[stage2-parity] Python 計画JSON: " + pyJson);

            long floorBeforeJava = Stage2OutputNaming.maxPrimaryPlanJsonLastModifiedMillis(outDir);
            long floorMemJsonJava = Stage2OutputNaming.maxPrimaryMemberJsonLastModifiedMillis(outDir);
            long floorPlanXlsxJava = Stage2OutputNaming.maxPrimaryPlanXlsxLastModifiedMillis(outDir);
            long floorMemXlsxJava = Stage2OutputNaming.maxPrimaryMemberXlsxLastModifiedMillis(outDir);

            Map<String, String> childJava = new HashMap<>(childBaseMerged);
            childJava.put(AppPaths.KEY_PM_AI_STAGE2_ENGINE, "java");
            lg.accept("[stage2-parity] (2) Java 段階2 を実行します…");
            lg.accept("[stage2-java] PM_AI_STAGE2_ENGINE=java — JVM 内 jp.co.pm.ai.planning.stage2（同一検証）");
            Stage2RunContext jctx = new Stage2RunContext(childJava, taskInputWorkbookPath, lg);
            int javaCode = Stage2JavaEngine.run(jctx);
            lg.accept("[stage2-parity] Java exitCode=" + javaCode + " " + Stage2ExitCodes.hint(javaCode));
            if (javaCode != 0) {
                return Outcome.fatal(
                        new IllegalStateException(
                                "Java 段階2が失敗したため比較を中止しました (exit=" + javaCode + ")."));
            }
            Path javaJson = Stage2OutputNaming.newestPrimaryPlanJsonAfter(outDir, floorBeforeJava);
            if (javaJson == null || !Files.isRegularFile(javaJson)) {
                return Outcome.fatal(
                        new IOException("Java 実行後も新しい計画 JSON が見つかりません: " + outDir));
            }
            lg.accept("[stage2-parity] Java 計画JSON: " + javaJson);

            Stage2ParityCheckResult planJsonCmp =
                    Stage2ProductionPlanJsonParity.compareFiles(pyJson, javaJson);

            Stage2ParityCheckResult memberJsonCmp = null;
            if (Stage2EnvParsing.envEnabled(AppPaths.KEY_PM_AI_MEMBER_SCHEDULE_JSON, childBaseMerged, true)) {
                Path pyMemberJson =
                        Stage2OutputNaming.newestPrimaryMemberJsonAfter(outDir, floorMemJsonPy);
                Path javaMemberJson =
                        Stage2OutputNaming.newestPrimaryMemberJsonAfter(outDir, floorMemJsonJava);
                if (pyMemberJson == null
                        || !Files.isRegularFile(pyMemberJson)
                        || javaMemberJson == null
                        || !Files.isRegularFile(javaMemberJson)) {
                    return Outcome.fatal(
                            new IOException(
                                    "人員 JSON が見つかりません（PM_AI_MEMBER_SCHEDULE_JSON 有効時）。dir="
                                            + outDir));
                }
                memberJsonCmp =
                        Stage2ProductionPlanJsonParity.compareFiles(pyMemberJson, javaMemberJson);
            }

            Stage2ParityCheckResult planWorkbookCmp = null;
            Stage2ParityCheckResult memberWorkbookCmp = null;
            if (Stage2EnvParsing.stage2WriteExcel(childBaseMerged)) {
                Path pyPlanXlsx =
                        Stage2OutputNaming.newestPrimaryPlanXlsxAfter(outDir, floorPlanXlsxPy);
                Path javaPlanXlsx =
                        Stage2OutputNaming.newestPrimaryPlanXlsxAfter(outDir, floorPlanXlsxJava);
                if (pyPlanXlsx == null
                        || !Files.isRegularFile(pyPlanXlsx)
                        || javaPlanXlsx == null
                        || !Files.isRegularFile(javaPlanXlsx)) {
                    return Outcome.fatal(
                            new IOException(
                                    "計画 xlsx が見つかりません（PM_AI_STAGE2_WRITE_EXCEL 有効時）。dir="
                                            + outDir));
                }
                planWorkbookCmp = Stage2WorkbookSemanticParity.compareXlsx(pyPlanXlsx, javaPlanXlsx);

                Path pyMemberXlsx =
                        Stage2OutputNaming.newestPrimaryMemberXlsxAfter(outDir, floorMemXlsxPy);
                Path javaMemberXlsx =
                        Stage2OutputNaming.newestPrimaryMemberXlsxAfter(outDir, floorMemXlsxJava);
                if (pyMemberXlsx == null
                        || !Files.isRegularFile(pyMemberXlsx)
                        || javaMemberXlsx == null
                        || !Files.isRegularFile(javaMemberXlsx)) {
                    return Outcome.fatal(
                            new IOException(
                                    "人員 xlsx が見つかりません（PM_AI_STAGE2_WRITE_EXCEL 有効時）。dir="
                                            + outDir));
                }
                memberWorkbookCmp =
                        Stage2WorkbookSemanticParity.compareXlsx(pyMemberXlsx, javaMemberXlsx);
            }

            return Outcome.compared(
                    new Stage2ParityBundle(
                            planInputCmp,
                            planJsonCmp,
                            memberJsonCmp,
                            planWorkbookCmp,
                            memberWorkbookCmp));
        } catch (Throwable ex) {
            return Outcome.fatal(ex);
        }
    }

    private static int runPythonStage2Blocking(
            Map<String, String> childEnv,
            String taskInputWorkbookPath,
            String scriptDirFallback,
            Consumer<String> log)
            throws IOException, InterruptedException {
        Path py = StagePythonExecutable.resolve(childEnv);
        Path dir =
                Path.of(
                        firstNonBlank(
                                childEnv.get(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR), scriptDirFallback));
        RunRequest req =
                new RunRequest(py, dir, SCRIPT_STAGE2, taskInputWorkbookPath != null ? taskInputWorkbookPath : "", childEnv);
        return PythonProcessRunner.runBlockingSameThread(
                req, line -> IpcStdoutTap.handleLine(line, log));
    }

    private static String firstNonBlank(String... parts) {
        if (parts == null) {
            return "";
        }
        for (String p : parts) {
            if (p != null && !p.isBlank()) {
                return p.trim();
            }
        }
        return "";
    }
}
