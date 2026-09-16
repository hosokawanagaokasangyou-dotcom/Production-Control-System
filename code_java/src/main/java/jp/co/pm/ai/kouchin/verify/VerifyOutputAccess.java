package jp.co.pm.ai.kouchin.verify;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * 検証結果 Excel の出力先が書けるか。全先が書けないときだけ実行を止める。
 */
public final class VerifyOutputAccess {

    static final String PROBE_NAME = ".pm-ai-kouchin-write-probe.tmp";

    private VerifyOutputAccess() {}

    public static boolean canWriteDir(Path dir) {
        if (dir == null) {
            return false;
        }
        Path probe = null;
        try {
            Files.createDirectories(dir);
            probe = dir.resolve(PROBE_NAME);
            Files.write(probe, new byte[] {1});
            return true;
        } catch (IOException e) {
            return false;
        } finally {
            if (probe != null) {
                try {
                    Files.deleteIfExists(probe);
                } catch (IOException ignored) {
                    // 判定は書き込み成否が正
                }
            }
        }
    }

    public static boolean anyOutputWritable(List<Path> dirs) {
        if (dirs == null || dirs.isEmpty()) {
            return false;
        }
        for (Path dir : dirs) {
            if (canWriteDir(dir)) {
                return true;
            }
        }
        return false;
    }

    /** 全出力先が書けないときの理由。書ける先が1つでもあれば {@code null}。 */
    public static String writeBlockReason(List<Path> dirs) {
        return anyOutputWritable(dirs) ? null : "結果Excelを書き込めません";
    }
}
