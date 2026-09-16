package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.channels.SeekableByteChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.util.List;

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
            return css(present, readable);
        }

        public String writeCss() {
            return css(present, writable);
        }

        private static String css(boolean present, boolean ok) {
            if (!present) {
                return "pm-kouchin-access-na";
            }
            return ok ? "pm-kouchin-access-ok" : "pm-kouchin-access-ng";
        }
    }

    public static boolean factorySourcesReady(List<KouchinDiscovery.Row> rows) {
        return blockReason(rows) == null;
    }

    public static boolean bothFactoriesReady(
            List<KouchinDiscovery.Row> kokubu, List<KouchinDiscovery.Row> konan) {
        return factorySourcesReady(kokubu) && factorySourcesReady(konan);
    }

    /** 検証不可の理由。許可するときは {@code null}。 */
    public static String blockReason(List<KouchinDiscovery.Row> rows) {
        if (rows == null || rows.isEmpty()) {
            return "検出未完了";
        }
        for (KouchinDiscovery.Row row : rows) {
            if (row == null || row.missing()) {
                String role = row == null || row.role() == null || row.role().isBlank() ? "ファイル" : row.role();
                return "見つかりません: " + role;
            }
            String full = row.fullPath();
            Path path = full == null || full.isBlank() ? null : Path.of(full);
            if (!canReadAtLeastReadOnly(path)) {
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
}
