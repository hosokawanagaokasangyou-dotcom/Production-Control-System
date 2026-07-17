package jp.co.pm.ai.desktop.config;

import java.nio.file.Path;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

/**
 * 段階1／段階2のコア成果物はローカルディスクのみに置く方針。
 *
 * <p>工場共有（UNC / 共有DATA）への直書きを拒否し、リポジトリ {@code output/} へフォールバックする。
 * 段階2後のアラジン入力用 Excel・サマリ Excel は共有のままとし、本クラスの対象外。
 */
public final class PipelineLocalResultsPolicy {

    private PipelineLocalResultsPolicy() {}

    /** UNC または工場共有 DATA（湖南／国分）配下なら true。 */
    public static boolean isSharedOrUncPath(Path path) {
        if (path == null) {
            return false;
        }
        String raw;
        try {
            raw = path.toAbsolutePath().normalize().toString();
        } catch (RuntimeException ex) {
            raw = path.toString();
        }
        return isSharedOrUncPathText(raw);
    }

    public static boolean isSharedOrUncPathText(String pathText) {
        if (pathText == null || pathText.isBlank()) {
            return false;
        }
        String s = pathText.strip().replace('/', '\\');
        if (s.startsWith("\\\\") || s.startsWith("//")) {
            return true;
        }
        String lower = s.toLowerCase(Locale.ROOT);
        if (lower.startsWith("m:\\湖南工場") || lower.startsWith("m:/湖南工場")) {
            return true;
        }
        String konan = AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR.replace('/', '\\');
        String kokubu = AppPaths.DEFAULT_KOKUBU_SHARED_DATA_DIR.replace('/', '\\');
        String konanM = AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR_M.replace('/', '\\');
        if (startsWithIgnoreCase(s, konan)
                || startsWithIgnoreCase(s, kokubu)
                || startsWithIgnoreCase(s, konanM)) {
            return true;
        }
        /*
         * 工場共有 DATA フォルダ名のみをヒューリスティックにする。
         * 「●配台AIシステム」単体はポータブル導入先（例: C:\●配台AIシステム\PMD_initial_install\pm-ai-data\output）
         * にも含まれるため共有扱いしない（段階2直前クリーンアップが plan_input_tasks.xlsx を消す事故を防ぐ）。
         */
        return s.contains("\\共有DATA");
    }

    /** 段階1／2 コア成果物のローカル出力ディレクトリ（{@code {repo}/output}）。 */
    public static Path localPipelineOutputDir(Map<String, String> ui) {
        return AppPaths.resolveRepoRoot(ui != null ? ui : Map.of())
                .resolve("output")
                .toAbsolutePath()
                .normalize();
    }

    /**
     * 候補パスが共有／UNC なら {@link #localPipelineOutputDir} を返す。それ以外は候補を正規化して返す。
     */
    public static Path requireLocalDirectory(Path candidate, Map<String, String> ui) {
        if (candidate == null || isSharedOrUncPath(candidate)) {
            return localPipelineOutputDir(ui);
        }
        return candidate.toAbsolutePath().normalize();
    }

    /**
     * 子プロセス／解決用の env マップで、段階成果物系キーが共有を指していればローカルへ書き換える。
     *
     * @return 変更があったとき true
     */
    public static boolean rewritePipelineOutputEnvToLocal(Map<String, String> ui) {
        if (ui == null) {
            return false;
        }
        boolean changed = false;
        Path localOut = localPipelineOutputDir(ui);
        String localOutStr = localOut.toString();

        if (rewriteIfShared(ui, AppPaths.KEY_PM_AI_OUTPUT_DIR, localOutStr)) {
            changed = true;
        }
        if (rewriteIfShared(ui, AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR, localOutStr)) {
            changed = true;
        }

        String plan = trim(ui.get(AppPaths.KEY_PM_AI_PLAN_INPUT_PATH));
        if (!plan.isEmpty() && isSharedOrUncPathText(plan)) {
            ui.put(
                    AppPaths.KEY_PM_AI_PLAN_INPUT_PATH,
                    localOut.resolve(AppPaths.STAGE1_PLAN_TASKS_FILENAME).toString());
            changed = true;
        }

        String exclude = trim(ui.get(AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON));
        if (!exclude.isEmpty() && isSharedOrUncPathText(exclude)) {
            ui.put(
                    AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
                    localOut.resolve(AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME).toString());
            changed = true;
        }
        return changed;
    }

    /** {@link #rewritePipelineOutputEnvToLocal} 用の可変コピーを返す。 */
    public static Map<String, String> copyWithLocalPipelineOutputs(Map<String, String> ui) {
        Map<String, String> copy = new HashMap<>(ui != null ? ui : Map.of());
        rewritePipelineOutputEnvToLocal(copy);
        return copy;
    }

    private static boolean rewriteIfShared(Map<String, String> ui, String key, String localValue) {
        String current = trim(ui.get(key));
        if (current.isEmpty()) {
            return false;
        }
        if (isSharedOrUncPathText(current)) {
            ui.put(key, localValue);
            return true;
        }
        return false;
    }

    private static boolean startsWithIgnoreCase(String value, String prefix) {
        if (value == null || prefix == null || prefix.isEmpty()) {
            return false;
        }
        return value.regionMatches(true, 0, prefix, 0, prefix.length());
    }

    private static String trim(String s) {
        return s != null ? s.strip() : "";
    }
}
