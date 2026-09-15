package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;

/** 検査表索引 CSV の共有フォルダへの公開と、起動時のローカル取り込み。 */
public final class InspectionSheetIndexShare {

    public static final double SIZE_DIFF_RATIO = 0.10;
    public static final long SIZE_DIFF_MIN_BYTES = 64L * 1024L;

    public enum PullResult {
        COPIED,
        SKIPPED,
        SHARE_MISSING
    }

    private InspectionSheetIndexShare() {}

    public static boolean shouldPull(long localBytes, long sharedBytes) {
        if (sharedBytes <= 0L) {
            return false;
        }
        if (localBytes <= 0L) {
            return true;
        }
        if (localBytes >= sharedBytes) {
            return false;
        }
        long deficit = sharedBytes - localBytes;
        long threshold =
                Math.max(SIZE_DIFF_MIN_BYTES, (long) (sharedBytes * SIZE_DIFF_RATIO));
        return deficit > threshold;
    }

    public static void publish(Path localCsv, Path shareCsv) throws IOException {
        if (localCsv == null || !Files.isRegularFile(localCsv)) {
            throw new IOException("本索引がありません。先に検査表索引を完走してください。");
        }
        copyAtomic(localCsv, shareCsv);
    }

    public static PullResult pullIfNeeded(Path localCsv, Path shareCsv) throws IOException {
        long shared = fileSize(shareCsv);
        if (shared <= 0L) {
            return PullResult.SHARE_MISSING;
        }
        long local = fileSize(localCsv);
        if (!shouldPull(local, shared)) {
            return PullResult.SKIPPED;
        }
        copyAtomic(shareCsv, localCsv);
        return PullResult.COPIED;
    }

    public static void publish(Map<String, String> ui) throws IOException {
        FactorySite site = InspectionSheetOpenService.factorySite(ui);
        publish(InspectionSheetIndexStore.indexFile(site), shareIndexFile(ui, site));
    }

    public static PullResult pullIfNeeded(Map<String, String> ui) throws IOException {
        FactorySite site = InspectionSheetOpenService.factorySite(ui);
        return pullIfNeeded(InspectionSheetIndexStore.indexFile(site), shareIndexFile(ui, site));
    }

    public static Path shareIndexFile(Map<String, String> ui, FactorySite site) {
        return resolveShareDir(ui, site).resolve(indexFileName(site));
    }

    public static Path resolveShareDir(Map<String, String> ui, FactorySite site) {
        Map<String, String> u = ui != null ? ui : Map.of();
        String override = u.getOrDefault(AppPaths.KEY_PM_AI_INSPECTION_SHEET_INDEX_SHARE_DIR, "");
        if (override != null && !override.isBlank()) {
            return Path.of(override.strip()).toAbsolutePath().normalize();
        }
        return Path.of(AppPaths.defaultInspectionSheetIndexShareDirForFactory(site))
                .toAbsolutePath()
                .normalize();
    }

    static String indexFileName(FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        if (effective == FactorySite.RDP_LAUNCHER) {
            effective = FactorySite.KONAN;
        }
        return effective.name() + ".csv";
    }

    static long fileSize(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return 0L;
        }
        try {
            return Files.size(path);
        } catch (IOException ex) {
            return 0L;
        }
    }

    static void copyAtomic(Path src, Path dest) throws IOException {
        if (src == null || dest == null) {
            throw new IOException("複写先または複写元が空です");
        }
        if (dest.getParent() != null) {
            Files.createDirectories(dest.getParent());
        }
        Path tmp = dest.resolveSibling(dest.getFileName().toString() + ".tmp");
        Files.copy(src, tmp, StandardCopyOption.REPLACE_EXISTING);
        try {
            Files.move(
                    tmp, dest, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
        } catch (AtomicMoveNotSupportedException ex) {
            Files.move(tmp, dest, StandardCopyOption.REPLACE_EXISTING);
        }
    }
}
