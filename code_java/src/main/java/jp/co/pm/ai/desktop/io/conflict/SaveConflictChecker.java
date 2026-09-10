package jp.co.pm.ai.desktop.io.conflict;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** baseline とディスク上の指紋を比較する。 */
public final class SaveConflictChecker {

    private SaveConflictChecker() {}

    public static ConflictCheckResult check(FingerprintBaseline baseline) {
        Objects.requireNonNull(baseline, "baseline");
        List<Path> mismatched = new ArrayList<>();
        try {
            for (Path path : baseline.paths()) {
                String current = FileContentFingerprint.sha256Hex(path);
                String expected = baseline.hashes().get(path);
                if (!Objects.equals(current, expected)) {
                    mismatched.add(path);
                }
            }
        } catch (IOException e) {
            return ConflictCheckResult.ioError(
                    e.getMessage() != null ? e.getMessage() : e.toString());
        }
        if (mismatched.isEmpty()) {
            return ConflictCheckResult.ok();
        }
        return ConflictCheckResult.conflict(mismatched);
    }
}
