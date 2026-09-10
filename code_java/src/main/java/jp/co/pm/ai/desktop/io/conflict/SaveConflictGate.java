package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

/** 保存競合ゲートの補助（ディスク読込・要約フォールバック）。 */
public final class SaveConflictGate {

    private SaveConflictGate() {}

    public static Map<Path, byte[]> readDiskBytes(FingerprintBaseline baseline) throws IOException {
        Map<Path, byte[]> disk = new LinkedHashMap<>();
        for (Path p : baseline.paths()) {
            if (Files.isRegularFile(p)) {
                disk.put(p, Files.readAllBytes(p));
            } else {
                disk.put(p, new byte[0]);
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
            StringBuilder sb = new StringBuilder("詳細差分を生成できませんでした。\n");
            for (Path p : check.mismatchedPaths()) {
                sb.append("・").append(p.getFileName()).append('\n');
            }
            return sb.toString().trim();
        }
    }
}
