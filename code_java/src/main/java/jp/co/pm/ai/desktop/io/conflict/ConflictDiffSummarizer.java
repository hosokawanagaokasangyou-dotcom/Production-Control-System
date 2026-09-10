package jp.co.pm.ai.desktop.io.conflict;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;

@FunctionalInterface
public interface ConflictDiffSummarizer {

    /**
     * @param baselineSnapshots 読込時バイト（欠落は空配列）
     * @param diskBytes 現ディスクバイト（欠落は空配列）
     * @param mismatched 不一致パス
     * @return 業務語の要約（改行可）
     */
    String summarize(
            Map<Path, byte[]> baselineSnapshots,
            Map<Path, byte[]> diskBytes,
            List<Path> mismatched);
}
