package jp.co.pm.ai.desktop.io;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import jp.co.pm.ai.desktop.config.AppPaths;

/**
 * リモートデスクトップ接続前に、共有フォルダ（UNC）上の起動ファイル実在を確認する。
 */
public final class RdpLaunchSharedPathValidator {

    /** 1 件のパス検証結果。 */
    public record Issue(String label, String path, String detail) {}

    /** 検証の結果。 */
    public record Result(boolean ok, List<Issue> issues, List<String> checkedPaths) {

        public Result(boolean ok, List<Issue> issues) {
            this(ok, issues, List.of());
        }

        public String formatBlockingMessage() {
            if (ok || issues.isEmpty()) {
                return "";
            }
            StringBuilder sb = new StringBuilder();
            sb.append(
                    "接続前に共有フォルダ（UNC）上のファイルを確認しましたが、次が見つかりません。\n\n");
            sb.append(formatIssueList());
            sb.append(
                    "\nRAP設定でパスを修正するか、ファイルを共有 DATA に配置してから再実行してください。");
            return sb.toString();
        }

        /** 存在確認ボタン向け（問題あり）。 */
        public String formatExistenceNgMessage() {
            if (ok || issues.isEmpty()) {
                return "";
            }
            StringBuilder sb = new StringBuilder();
            sb.append("次のファイルが見つかりません。\n\n");
            sb.append(formatIssueList());
            return sb.toString();
        }

        /** 存在確認ボタン向け（問題なし）。 */
        public String formatExistenceOkMessage() {
            if (!ok || checkedPaths.isEmpty()) {
                return "ファイルを確認しました。";
            }
            StringBuilder sb = new StringBuilder("次のファイルを確認しました。\n\n");
            for (String path : checkedPaths) {
                sb.append("・").append(path).append("\n");
            }
            return sb.toString();
        }

        private String formatIssueList() {
            StringBuilder sb = new StringBuilder();
            for (Issue issue : issues) {
                sb.append("・").append(issue.label()).append("\n  ").append(issue.path());
                if (issue.detail() != null && !issue.detail().isBlank()) {
                    sb.append("\n  （").append(issue.detail()).append("）");
                }
                sb.append("\n");
            }
            return sb.toString();
        }
    }

    private RdpLaunchSharedPathValidator() {}

    /** RPA プログラム（exe）パスの存在確認。パス修復後に {@link Files#isRegularFile} で判定。 */
    public static Result validateProgramPath(String program) {
        String repaired = UncPathSegmentRepair.repair(trimToEmpty(program));
        if (repaired.isBlank()) {
            return new Result(
                    false,
                    List.of(new Issue("RPA プログラム", "(未入力)", "パスが空です")),
                    List.of());
        }
        List<Issue> issues = new ArrayList<>();
        checkRegularFile(issues, "RPA プログラム", repaired);
        return new Result(
                issues.isEmpty(), List.copyOf(issues), List.of(repaired));
    }

    /** RPA 引数内のシナリオ（.ardrpa）パスの存在確認。シナリオ未指定なら ok。 */
    public static Result validateScenarioArguments(String arguments) {
        String repairedArgs =
                RpaScenarioArgumentSupport.repairScenarioArguments(trimToEmpty(arguments));
        List<String> scenarioPaths =
                RpaScenarioArgumentSupport.extractScenarioPaths(repairedArgs);
        if (scenarioPaths.isEmpty()) {
            return new Result(true, List.of(), List.of());
        }
        List<Issue> issues = new ArrayList<>();
        List<String> checked = new ArrayList<>();
        for (String scenarioPath : scenarioPaths) {
            String repaired = UncPathSegmentRepair.repair(scenarioPath);
            checked.add(repaired);
            checkRegularFile(issues, "RPA シナリオ", repaired);
        }
        return new Result(issues.isEmpty(), List.copyOf(issues), List.copyOf(checked));
    }

    /**
     * 起動プロファイルと共有ランチャー exe の UNC パスを検証する。
     */
    public static Result validateBeforeConnect(
            String rpaProgram, String rpaArguments, Path sharedLauncherExe) {
        List<Issue> issues = new ArrayList<>();
        List<String> checked = new ArrayList<>();

        Result program = validateProgramPath(rpaProgram);
        issues.addAll(program.issues());
        checked.addAll(program.checkedPaths());

        Result scenarios = validateScenarioArguments(rpaArguments);
        issues.addAll(scenarios.issues());
        checked.addAll(scenarios.checkedPaths());

        if (sharedLauncherExe != null) {
            String launcherPath =
                    UncPathSegmentRepair.repair(sharedLauncherExe.toString());
            if (isUncPath(launcherPath)) {
                checked.add(launcherPath);
                checkRegularFile(issues, AppPaths.RDP_LAUNCHER_EXE_BASENAME, launcherPath);
            }
        }
        return new Result(issues.isEmpty(), List.copyOf(issues), List.copyOf(checked));
    }

    static boolean isUncPath(String path) {
        if (path == null || path.isBlank()) {
            return false;
        }
        String trimmed = path.strip();
        return trimmed.startsWith("\\\\") || trimmed.startsWith("//");
    }

    private static void checkRegularFile(List<Issue> issues, String label, String path) {
        if (path.isBlank()) {
            return;
        }
        try {
            Path file = Path.of(path);
            if (Files.isRegularFile(file)) {
                return;
            }
            if (Files.exists(file)) {
                issues.add(new Issue(label, path, "ファイルではありません"));
                return;
            }
            issues.add(new Issue(label, path, "見つかりません"));
        } catch (Exception ex) {
            issues.add(new Issue(label, path, "参照できません: " + ex.getMessage()));
        }
    }

    private static String trimToEmpty(String value) {
        return value != null ? value.strip() : "";
    }
}
