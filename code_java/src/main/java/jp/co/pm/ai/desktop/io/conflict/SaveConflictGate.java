package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

/** 保存競合ゲートの補助（ディスク読込・要約フォールバック）。 */
public final class SaveConflictGate {

    private SaveConflictGate() {}

    /** 欠落は null、空ファイルは length 0。 */
    public static Map<Path, byte[]> readDiskBytes(FingerprintBaseline baseline) throws IOException {
        Map<Path, byte[]> disk = new LinkedHashMap<>();
        for (Path p : baseline.paths()) {
            try {
                if (Files.isRegularFile(p)) {
                    disk.put(p, Files.readAllBytes(p));
                } else {
                    disk.put(p, null);
                }
            } catch (NoSuchFileException e) {
                disk.put(p, null);
            }
        }
        return disk;
    }

    public static String summarizeOrFallback(
            ConflictDiffSummarizer summarizer,
            FingerprintBaseline baseline,
            Map<Path, byte[]> diskBytes,
            ConflictCheckResult check) {
        try {
            return summarizer.summarize(baseline.snapshots(), diskBytes, check.mismatchedPaths());
        } catch (Exception e) {
            return fallbackSummary(diskBytes, check);
        }
    }

    public static String fallbackSummary(
            Map<Path, byte[]> diskBytes, ConflictCheckResult check) {
        StringBuilder sb = new StringBuilder("詳細差分を生成できませんでした。\n");
        for (Path p : check.mismatchedPaths()) {
            byte[] bytes = diskBytes != null ? diskBytes.get(p) : null;
            String hex =
                    bytes == null
                            ? FileContentFingerprint.ABSENT
                            : FileContentFingerprint.sha256Hex(bytes);
            String shortHex = hex.length() <= 12 ? hex : hex.substring(0, 12);
            sb.append("・")
                    .append(p.getFileName())
                    .append(" hash=")
                    .append(shortHex)
                    .append("…\n");
        }
        return sb.toString().trim();
    }
}
