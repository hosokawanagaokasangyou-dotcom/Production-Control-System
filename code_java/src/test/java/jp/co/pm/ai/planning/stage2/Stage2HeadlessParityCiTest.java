package jp.co.pm.ai.planning.stage2;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIfEnvironmentVariable;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.bridge.Stage2PythonChildEnv;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.planning.stage2.parity.Stage2HeadlessParityRunner;

/**
 * {@code PM_AI_STAGE2_HEADLESS_CI=1} のときのみ実行。リポジトリの {@code code/python} と Python 3.14+ を前提に、
 * {@link Stage2HeadlessParityRunner} で Python→Java を順に起動する。
 *
 * <p>計画 primary JSON の厳密一致は現状要求しない（Java 側が PassThrough のため）。入力照合と両エンジンの正常終了のみ検証する。
 */
@EnabledIfEnvironmentVariable(named = AppPaths.KEY_PM_AI_STAGE2_HEADLESS_CI, matches = "1")
class Stage2HeadlessParityCiTest {

    @Test
    void headlessRunner_pythonThenJava_exitsOk(@TempDir Path root) throws Exception {
        Path repo = resolveRepoRootContainingCodePython();
        Path pyDir = repo.resolve("code").resolve("python");
        assertTrue(Files.isRegularFile(pyDir.resolve("plan_simulation_stage2.py")));

        Path out = root.resolve("out");
        Files.createDirectories(out);
        Path master = root.resolve("master.xlsx");
        writeMinimalMaster(master);
        Path plan = root.resolve("plan.csv");
        Files.writeString(plan, "依頼NO,工程名\nT1,加工A\n", StandardCharsets.UTF_8);

        Map<String, String> ui = new HashMap<>();
        ui.put(AppPaths.KEY_PM_AI_OUTPUT_DIR, out.toString());
        ui.put(AppPaths.KEY_PM_AI_REPO_ROOT, repo.toString());
        ui.put(AppPaths.KEY_PM_AI_CODE_PYTHON_DIR, pyDir.toString());
        ui.put(AppPaths.KEY_PM_AI_MASTER_WORKBOOK, master.toString());
        ui.put(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH, plan.toString());
        ui.put(AppPaths.KEY_PM_AI_PLAN_WORKBOOK_JSON, "1");
        ui.put(AppPaths.KEY_PM_AI_MEMBER_SCHEDULE_JSON, "0");
        ui.put(AppPaths.KEY_PM_AI_STAGE2_WRITE_EXCEL, "0");
        ui.put(AppPaths.KEY_PM_AI_STAGE2_JAVA_DELEGATE_PYTHON_DISPATCH, "0");

        String pyExe = firstNonBlank(System.getenv(AppPaths.KEY_PM_AI_PYTHON), detectPlanningCorePythonOrNull());
        Assumptions.assumeTrue(
                pyExe != null && !pyExe.isBlank(),
                "planning_core を import できる Python（3.14+ 推奨）を PATH または PM_AI_PYTHON で指定してください");
        ui.put(AppPaths.KEY_PM_AI_PYTHON, pyExe);

        Map<String, String> child = new HashMap<>(ui);
        Stage2PythonChildEnv.stripLegacyWorkbookKeys(child);
        Stage2PythonChildEnv.ensureSkipWorkbookEnvSheetDefault(child);
        Stage2PythonChildEnv.applyNetworkSourceAndChildPause(child, true, true);

        List<String> log = new ArrayList<>();
        Stage2HeadlessParityRunner.Outcome o =
                Stage2HeadlessParityRunner.run(child, "", "", null, log::add);
        assertNull(o.fatalError(), () -> String.join("\n", log));
        assertNotNull(o.bundle());
        assertTrue(
                o.bundle().planInputUiVsDisk().identical(),
                () -> o.bundle().planInputUiVsDisk().summary());
    }

    private static String firstNonBlank(String a, String b) {
        if (a != null && !a.isBlank()) {
            return a.strip();
        }
        if (b != null && !b.isBlank()) {
            return b.strip();
        }
        return "";
    }

    /**
     * {@code python -c "import planning_core"} が成功する最初のコマンドを返す。見つからなければ {@code null}（テストを
     * {@link Assumptions#assumeTrue} でスキップ）。
     */
    private static String detectPlanningCorePythonOrNull() {
        String[] candidates = {"/usr/bin/python3.14", "/usr/local/bin/python3.14", "python3.14", "python3"};
        for (String c : candidates) {
            try {
                Process p =
                        new ProcessBuilder(c, "-c", "import planning_core")
                                .redirectErrorStream(true)
                                .start();
                if (p.waitFor(90, TimeUnit.SECONDS) && p.exitValue() == 0) {
                    return c;
                }
            } catch (@SuppressWarnings("unused") Exception ignored) {
                /* try next */
            }
        }
        return null;
    }

    private static Path resolveRepoRootContainingCodePython() {
        Path c = Path.of("").toAbsolutePath().normalize();
        if (Files.isDirectory(c.resolve("code").resolve("python"))) {
            return c;
        }
        Path p = c.getParent();
        if (p != null && Files.isDirectory(p.resolve("code").resolve("python"))) {
            return p;
        }
        throw new IllegalStateException(
                "code/python が見つかりません（JUnit の cwd を code_java にして実行してください）: " + c);
    }

    private static void writeMinimalMaster(Path path) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var skills = wb.createSheet("skills");
            skills.createRow(0).createCell(0).setCellValue("メンバー");
            skills.createRow(0).createCell(1).setCellValue("工程テスト");
            skills.createRow(1).createCell(1).setCellValue("機械テスト");
            skills.createRow(2).createCell(0).setCellValue("テストオペレータ");
            var main = wb.createSheet("メイン");
            var r12 = main.createRow(11);
            r12.createCell(0).setCellValue("08:00");
            r12.createCell(1).setCellValue("17:00");
            try (OutputStream os = Files.newOutputStream(path)) {
                wb.write(os);
            }
        }
    }
}
