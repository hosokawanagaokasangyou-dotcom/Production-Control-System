package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;

/**
 * 段階2で使用した shaped アラジン加工計画（{@link AppPaths#resolveShapedAladdinPlanJsonPath}）と、
 * タスク入力ソース最新ファイルの再読込結果が同一かを判定する。
 */
public final class AladdinProcessingPlanSourceFreshness {

    public record Result(boolean identicalToNewestSource, Optional<Path> newestSourceFile) {}

    private AladdinProcessingPlanSourceFreshness() {}

    /**
     * 保存済み shaped JSON とソース最新の加工計画が同一内容なら {@code true}。
     * shaped JSON が無い・空・ソース未解決のときは {@code false}（再読込ボタンは有効のまま）。
     */
    public static boolean isSavedShapedPlanIdenticalToNewestSource(Map<String, String> ui) {
        try {
            return evaluate(ui).identicalToNewestSource();
        } catch (IOException ex) {
            return false;
        }
    }

    public static Result evaluate(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path shapedJson = AppPaths.resolveShapedAladdinPlanJsonPath(u);
        if (!Files.isRegularFile(shapedJson)) {
            return new Result(false, Optional.empty());
        }
        PlanInputTabularIo.TabularSheet saved = loadSavedShapedTabular(shapedJson);
        if (saved.headers().isEmpty() && saved.rows().isEmpty()) {
            return new Result(false, Optional.empty());
        }

        Path dir = AppPaths.resolveTaskInputSourceDir(u);
        if (dir == null || !Files.isDirectory(dir)) {
            return new Result(false, Optional.empty());
        }
        Optional<Path> newest = NetworkSourceDirResolver.newestTaskInputFileInDirectory(dir, u);
        if (newest.isEmpty()) {
            return new Result(false, Optional.empty());
        }
        Path file = newest.get().toAbsolutePath().normalize();
        String low = file.getFileName().toString().toLowerCase(Locale.ROOT);
        if (low.endsWith(".pq") || low.endsWith(".parquet")) {
            return new Result(false, Optional.of(file));
        }

        PlanInputTabularIo.TabularSheet fromSource =
                AladdinProcessingPlanSourceReloader.readNewestAladdinTabularFromDisk(file);
        return new Result(tabularSheetsEqual(saved, fromSource), Optional.of(file));
    }

    static boolean tabularSheetsEqual(
            PlanInputTabularIo.TabularSheet a, PlanInputTabularIo.TabularSheet b) {
        if (a == null || b == null) {
            return false;
        }
        return normalizeHeaders(a.headers()).equals(normalizeHeaders(b.headers()))
                && normalizeRows(a.rows()).equals(normalizeRows(b.rows()));
    }

    private static PlanInputTabularIo.TabularSheet loadSavedShapedTabular(Path shapedJson)
            throws IOException {
        JsonTableIo.ArrayTable table = JsonTableIo.loadArrayTable(shapedJson);
        return new PlanInputTabularIo.TabularSheet(table.columns(), table.rows());
    }

    private static List<String> normalizeHeaders(List<String> headers) {
        if (headers == null) {
            return List.of();
        }
        return headers.stream().map(h -> h != null ? h : "").toList();
    }

    private static List<List<String>> normalizeRows(List<List<String>> rows) {
        if (rows == null) {
            return List.of();
        }
        return rows.stream()
                .map(
                        row ->
                                row == null
                                        ? List.<String>of()
                                        : row.stream().map(c -> c != null ? c : "").toList())
                .toList();
    }
}
