package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.List;
import java.util.function.Function;

/**
 * 検証実行前のソース到達判定。Excel が開いていても読み取り専用で開けば可。
 */
public final class VerifySourceAccess {

    private VerifySourceAccess() {}

    public static boolean canReadAtLeastReadOnly(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return false;
        }
        try (SeekableByteChannel ch = Files.newByteChannel(path, StandardOpenOption.READ)) {
            return true;
        } catch (IOException e) {
            return false;
        }
    }

    public static boolean canWriteFile(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return false;
        }
        try (RandomAccessFile raf = new RandomAccessFile(path.toFile(), "rw")) {
            return true;
        } catch (IOException e) {
            return false;
        }
    }

    /** 検出行／実ファイルの読取・書込表示。欠落は {@code —}。 */
    public record FileAccess(boolean present, boolean readable, boolean writable) {

        public static FileAccess of(Path path) {
            if (path == null || !Files.isRegularFile(path)) {
                return new FileAccess(false, false, false);
            }
            return new FileAccess(true, canReadAtLeastReadOnly(path), canWriteFile(path));
        }

        public static FileAccess ofRow(KouchinDiscovery.Row row) {
            if (row == null || row.missing()) {
                return new FileAccess(false, false, false);
            }
            String full = row.fullPath();
            return of(full == null || full.isBlank() ? null : Path.of(full));
        }

        public String readLabel() {
            return present ? (readable ? "可" : "不可") : "—";
        }

        public String writeLabel() {
            return present ? (writable ? "可" : "不可") : "—";
        }

        public String readCss() {
            return css(present, readable, true);
        }

        public String writeCss() {
            return css(present, writable, false);
        }

        public String readHint() {
            if (!present) {
                return "ファイルなし";
            }
            return readable ? "検証に必要。読取可" : "検証不可。読取できない";
        }

        public String writeHint() {
            if (!present) {
                return "ファイルなし";
            }
            return writable ? "書込可（検証には不要）" : "Excelで開いている等。検証は読取できれば可";
        }

        private static String css(boolean present, boolean ok, boolean readColumn) {
            if (!present) {
                return "pm-kouchin-access-na";
            }
            if (ok) {
                return "pm-kouchin-access-ok";
            }
            return readColumn ? "pm-kouchin-access-ng" : "pm-kouchin-access-warn";
        }
    }

    public static boolean factorySourcesReady(List<KouchinDiscovery.Row> rows) {
        return blockReason(rows) == null;
    }

    public static boolean factorySourcesReady(
            List<KouchinDiscovery.Row> rows, Function<KouchinDiscovery.Row, FileAccess> accessOf) {
        return blockReason(rows, accessOf) == null;
    }

    public static boolean bothFactoriesReady(
            List<KouchinDiscovery.Row> kokubu, List<KouchinDiscovery.Row> konan) {
        return factorySourcesReady(kokubu) && factorySourcesReady(konan);
    }

    /** 検証不可の理由。許可するときは {@code null}。 */
    public static String blockReason(List<KouchinDiscovery.Row> rows) {
        return blockReason(rows, FileAccess::ofRow);
    }

    public static String blockReason(
            List<KouchinDiscovery.Row> rows, Function<KouchinDiscovery.Row, FileAccess> accessOf) {
        if (rows == null || rows.isEmpty()) {
            return "検出未完了";
        }
        Function<KouchinDiscovery.Row, FileAccess> fn = accessOf == null ? FileAccess::ofRow : accessOf;
        for (KouchinDiscovery.Row row : rows) {
            if (optionalSource(row == null ? null : row.role())) {
                continue;
            }
            if (row == null || row.missing()) {
                String role = row == null || row.role() == null || row.role().isBlank() ? "ファイル" : row.role();
                return "見つかりません: " + role;
            }
            FileAccess acc = fn.apply(row);
            if (acc == null || !acc.present() || !acc.readable()) {
                return "読み取れません: " + row.role();
            }
        }
        boolean has1 = false;
        boolean has2 = false;
        boolean has3 = false;
        for (KouchinDiscovery.Row row : rows) {
            String role = row.role() == null ? "" : row.role();
            if (role.startsWith("①")) {
                has1 = true;
            } else if (role.startsWith("②")) {
                has2 = true;
            } else if (role.startsWith("③")) {
                has3 = true;
            }
        }
        if (!has1 || !has2 || !has3) {
            return "検出未完了";
        }
        return null;
    }

    static boolean optionalSource(String role) {
        return role != null && role.contains("月次処理");
    }
}
