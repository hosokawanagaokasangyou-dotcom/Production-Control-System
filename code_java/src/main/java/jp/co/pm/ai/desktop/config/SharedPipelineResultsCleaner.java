package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import jp.co.pm.ai.desktop.io.Stage2OutputNaming;

/**
 * 工場共有 DATA 上に残った段階1／段階2コア成果物を削除する。
 *
 * <p>削除しないもの:
 * <ul>
 *   <li>アラジン入力用配台計画（段階2後の共有出力）</li>
 *   <li>サマリ_AI配台.xlsx およびマスタ類</li>
 *   <li>タスク入力／実績明細のソース本体（ネットワーク原本）</li>
 * </ul>
 */
public final class SharedPipelineResultsCleaner {

    private static final Logger LOG = Logger.getLogger(SharedPipelineResultsCleaner.class.getName());

    private SharedPipelineResultsCleaner() {}

    /**
     * 既知の工場共有 DATA ルートから段階1／2成果物を削除する。
     *
     * @return 削除に成功したパス一覧（存在しなかったものは含めない）
     */
    public static List<Path> deletePipelineArtifactsFromShared(Map<String, String> ui) {
        List<Path> deleted = new ArrayList<>();
        for (Path root : sharedDataRoots(ui)) {
            if (root == null || !Files.isDirectory(root)) {
                continue;
            }
            deleted.addAll(deleteUnderSharedRoot(root));
        }
        return deleted;
    }

    static List<Path> sharedDataRoots(Map<String, String> ui) {
        Set<Path> roots = new LinkedHashSet<>();
        Map<String, String> u = ui != null ? ui : Map.of();
        addIfPresent(roots, Path.of(AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR));
        addIfPresent(roots, Path.of(AppPaths.DEFAULT_KONAN_SHARED_DATA_DIR_M));
        addIfPresent(roots, Path.of(AppPaths.DEFAULT_KOKUBU_SHARED_DATA_DIR));
        Path summaryParent = AppPaths.summarySharedDataDir(u);
        if (PipelineLocalResultsPolicy.isSharedOrUncPath(summaryParent)) {
            roots.add(summaryParent.toAbsolutePath().normalize());
        }
        String outOverride = trim(u.get(AppPaths.KEY_PM_AI_OUTPUT_DIR));
        if (!outOverride.isEmpty()) {
            Path p = Path.of(outOverride);
            if (PipelineLocalResultsPolicy.isSharedOrUncPath(p)) {
                roots.add(p.toAbsolutePath().normalize());
            }
        }
        String resultDir = trim(u.get(AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR));
        if (!resultDir.isEmpty()) {
            Path p = Path.of(resultDir);
            if (PipelineLocalResultsPolicy.isSharedOrUncPath(p)) {
                roots.add(p.toAbsolutePath().normalize());
            }
        }
        return List.copyOf(roots);
    }

    private static List<Path> deleteUnderSharedRoot(Path root) {
        List<Path> deleted = new ArrayList<>();
        deleteExact(root.resolve(AppPaths.STAGE1_PLAN_TASKS_FILENAME), deleted);
        deleteExact(root.resolve(AppPaths.STAGE1_TASK_INPUT_PREVIEW_FILENAME), deleted);
        deleteExact(root.resolve(AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME), deleted);
        deleteExact(root.resolve(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME), deleted);
        deleteExact(root.resolve("結果_配台表.xlsx"), deleted);
        deleteExact(root.resolve("result_dispatch_table.json"), deleted);
        deleteExact(root.resolve("result_dispatch_table.xlsx"), deleted);

        try (DirectoryStream<Path> stream = Files.newDirectoryStream(root)) {
            for (Path child : stream) {
                if (!Files.isRegularFile(child)) {
                    continue;
                }
                String name = child.getFileName().toString();
                if (shouldDeleteStageArtifactFileName(name)) {
                    deleteExact(child, deleted);
                }
            }
        } catch (IOException ex) {
            LOG.log(Level.FINE, "共有ルート列挙失敗: " + root, ex);
        }
        return deleted;
    }

    /** 単体テスト用: ファイル名が段階1／2成果物として削除対象か。 */
    static boolean shouldDeleteStageArtifactFileName(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            return false;
        }
        String n = fileName.strip();
        String lower = n.toLowerCase(Locale.ROOT);
        if (lower.contains("アラジン入力用") || lower.startsWith("サマリ_")) {
            return false;
        }
        if (n.equals(AppPaths.STAGE1_PLAN_TASKS_FILENAME)
                || n.equals(AppPaths.STAGE1_TASK_INPUT_PREVIEW_FILENAME)
                || n.equals(AppPaths.STAGE1_EXCLUDE_RULES_JSON_FILENAME)
                || n.equals(AppPaths.RESULT_DISPATCH_TABLE_JSON_BASENAME)
                || n.equals("結果_配台表.xlsx")
                || n.equals("result_dispatch_table.json")
                || n.equals("result_dispatch_table.xlsx")) {
            return true;
        }
        Path fake = Path.of(n);
        if (Stage2OutputNaming.acceptsPrimaryPlanXlsx(fake)
                || Stage2OutputNaming.acceptsPrimaryMemberXlsx(fake)
                || Stage2OutputNaming.acceptsPrimaryMemberJson(fake)) {
            return true;
        }
        if (n.startsWith("計画") && (lower.endsWith(".xlsx") || lower.endsWith(".json"))) {
            return true;
        }
        if (n.startsWith("人員") && (lower.endsWith(".xlsx") || lower.endsWith(".json"))) {
            return true;
        }
        if (lower.startsWith("plan_") && (lower.endsWith(".xlsx") || lower.endsWith(".json"))) {
            return true;
        }
        if (lower.startsWith("member_") && (lower.endsWith(".xlsx") || lower.endsWith(".json"))) {
            return true;
        }
        if (lower.contains("shaped") && lower.endsWith(".json")) {
            return true;
        }
        if (lower.startsWith("stage1_") && (lower.endsWith(".xlsx") || lower.endsWith(".json"))) {
            return true;
        }
        return false;
    }

    private static void deleteExact(Path file, List<Path> deleted) {
        if (file == null || !Files.isRegularFile(file)) {
            return;
        }
        try {
            Files.delete(file);
            deleted.add(file.toAbsolutePath().normalize());
            LOG.info("共有上の段階成果物を削除: " + file);
        } catch (IOException ex) {
            LOG.log(Level.WARNING, "共有上の段階成果物削除失敗: " + file, ex);
        }
    }

    private static void addIfPresent(Set<Path> roots, Path path) {
        if (path == null) {
            return;
        }
        try {
            Path n = path.toAbsolutePath().normalize();
            if (Files.isDirectory(n)) {
                roots.add(n);
            }
        } catch (RuntimeException ignored) {
            // 到達不能ドライブ等
        }
    }

    private static String trim(String s) {
        return s != null ? s.strip() : "";
    }
}
