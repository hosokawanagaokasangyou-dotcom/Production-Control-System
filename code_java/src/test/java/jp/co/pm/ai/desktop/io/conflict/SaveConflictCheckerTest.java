package jp.co.pm.ai.desktop.io.conflict;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class SaveConflictCheckerTest {

    @TempDir Path dir;

    @Test
    void unchangedSet_isOk() throws Exception {
        Path json = dir.resolve("attendance-data.json");
        Path xlsx = dir.resolve("勤怠・機械カレンダー.xlsx");
        Files.writeString(json, "{\"members\":[]}", StandardCharsets.UTF_8);
        Files.write(xlsx, new byte[] {1, 2, 3});
        FingerprintBaseline baseline = FingerprintBaseline.capture(List.of(json, xlsx));
        ConflictCheckResult r = SaveConflictChecker.check(baseline);
        assertEquals(ConflictCheckResult.Kind.OK, r.kind());
    }

    @Test
    void changedRelatedFile_isConflict() throws Exception {
        Path json = dir.resolve("attendance-data.json");
        Path xlsx = dir.resolve("勤怠・機械カレンダー.xlsx");
        Files.writeString(json, "{\"members\":[]}", StandardCharsets.UTF_8);
        Files.write(xlsx, new byte[] {1, 2, 3});
        FingerprintBaseline baseline = FingerprintBaseline.capture(List.of(json, xlsx));
        Files.write(xlsx, new byte[] {9, 9, 9});
        ConflictCheckResult r = SaveConflictChecker.check(baseline);
        assertEquals(ConflictCheckResult.Kind.CONFLICT, r.kind());
        assertTrue(r.mismatchedPaths().contains(xlsx.toAbsolutePath().normalize()));
    }

    @Test
    void absentThenAppears_isConflict() throws Exception {
        Path json = dir.resolve("attendance-data.json");
        FingerprintBaseline baseline = FingerprintBaseline.capture(List.of(json));
        Files.writeString(json, "{}", StandardCharsets.UTF_8);
        assertEquals(ConflictCheckResult.Kind.CONFLICT, SaveConflictChecker.check(baseline).kind());
    }

    @Test
    void presentThenMissing_isConflict() throws Exception {
        Path json = dir.resolve("attendance-data.json");
        Files.writeString(json, "{}", StandardCharsets.UTF_8);
        FingerprintBaseline baseline = FingerprintBaseline.capture(List.of(json));
        Files.delete(json);
        assertEquals(ConflictCheckResult.Kind.CONFLICT, SaveConflictChecker.check(baseline).kind());
    }
}
