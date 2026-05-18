package jp.co.pm.ai.desktop.io;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/**
 * ネットワークソース（タスク入力／加工実績明細）の再読込高速化用。
 *
 * <p>直前に読み込んだファイルと<strong>ファイル名が同一</strong>のとき、POI 等のディスク読込を省略し、
 * メモリ上の表データを再利用する。
 */
public final class NetworkSourceFileReloadCache {

    /** 整形済み表＋ Excel シートメタ（不変スナップショット）。 */
    public record Snapshot(
            String fileName,
            boolean excel,
            List<String> sheetNames,
            int selectedSheetIndex,
            List<String> headers,
            List<List<String>> rows) {

        public Snapshot {
            fileName = fileName != null ? fileName : "";
            sheetNames = List.copyOf(sheetNames != null ? sheetNames : List.of());
            headers = List.copyOf(headers != null ? headers : List.of());
            rows = deepCopyRows(rows);
        }

        public PlanInputTabularIo.TabularSheet toTabularSheet() {
            return new PlanInputTabularIo.TabularSheet(new ArrayList<>(headers), deepCopyRows(rows));
        }
    }

    private static volatile Snapshot aladdinSnapshot;
    private static volatile Snapshot actualsSnapshot;

    private NetworkSourceFileReloadCache() {}

    public static Optional<Snapshot> matchAladdin(Path file) {
        return match(file, aladdinSnapshot);
    }

    public static Optional<Snapshot> matchActuals(Path file) {
        return match(file, actualsSnapshot);
    }

    public static void storeAladdin(
            Path file,
            boolean excel,
            List<String> sheetNames,
            int selectedSheetIndex,
            PlanInputTabularIo.TabularSheet shaped) {
        aladdinSnapshot = snapshotFrom(file, excel, sheetNames, selectedSheetIndex, shaped);
    }

    public static void storeActuals(
            Path file,
            boolean excel,
            List<String> sheetNames,
            int selectedSheetIndex,
            PlanInputTabularIo.TabularSheet shaped) {
        actualsSnapshot = snapshotFrom(file, excel, sheetNames, selectedSheetIndex, shaped);
    }

    /** 段階1キャッシュクリアやテストで、同一ファイル名の再読込省略を無効化する。 */
    public static void clearAll() {
        aladdinSnapshot = null;
        actualsSnapshot = null;
    }

    private static Optional<Snapshot> match(Path file, Snapshot cached) {
        if (file == null || cached == null || file.getFileName() == null) {
            return Optional.empty();
        }
        String name = file.getFileName().toString();
        if (name.isEmpty() || !name.equals(cached.fileName())) {
            return Optional.empty();
        }
        return Optional.of(cached);
    }

    private static Snapshot snapshotFrom(
            Path file,
            boolean excel,
            List<String> sheetNames,
            int selectedSheetIndex,
            PlanInputTabularIo.TabularSheet shaped) {
        String fileName =
                file != null && file.getFileName() != null
                        ? file.getFileName().toString()
                        : "";
        if (shaped == null) {
            return new Snapshot(fileName, excel, sheetNames, selectedSheetIndex, List.of(), List.of());
        }
        return new Snapshot(
                fileName,
                excel,
                sheetNames,
                selectedSheetIndex,
                shaped.headers(),
                shaped.rows());
    }

    private static List<List<String>> deepCopyRows(List<List<String>> src) {
        if (src == null || src.isEmpty()) {
            return List.of();
        }
        List<List<String>> out = new ArrayList<>(src.size());
        for (List<String> row : src) {
            out.add(row != null ? List.copyOf(row) : List.of());
        }
        return List.copyOf(out);
    }
}
