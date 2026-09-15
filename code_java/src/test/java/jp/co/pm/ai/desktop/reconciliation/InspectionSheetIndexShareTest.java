package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class InspectionSheetIndexShareTest {

    @Test
    void shouldPull_whenLocalMissingOrMuchSmaller() {
        assertTrue(InspectionSheetIndexShare.shouldPull(0L, 200_000L));
        long shared = 5_000_000L;
        long threshold =
                Math.max(
                        InspectionSheetIndexShare.SIZE_DIFF_MIN_BYTES,
                        (long) (shared * InspectionSheetIndexShare.SIZE_DIFF_RATIO));
        assertTrue(InspectionSheetIndexShare.shouldPull(shared - threshold - 1, shared));
        assertFalse(InspectionSheetIndexShare.shouldPull(shared - 1, shared));
        assertFalse(InspectionSheetIndexShare.shouldPull(shared + 100_000L, shared));
        assertFalse(InspectionSheetIndexShare.shouldPull(100L, 0L));
    }

    @Test
    void publish_copiesCompleteCsvAndPullRestoresWhenLocalMissing(@TempDir Path tmp)
            throws Exception {
        Path local = tmp.resolve("local").resolve("KONAN.csv");
        Path share = tmp.resolve("share").resolve("KONAN.csv");
        Files.createDirectories(local.getParent());
        Files.writeString(local, "irai_no,x\nA8-1,1\n", StandardCharsets.UTF_8);

        InspectionSheetIndexShare.publish(local, share);
        assertTrue(Files.isRegularFile(share));
        assertEquals(
                Files.readString(local, StandardCharsets.UTF_8),
                Files.readString(share, StandardCharsets.UTF_8));

        Files.delete(local);
        assertEquals(
                InspectionSheetIndexShare.PullResult.COPIED,
                InspectionSheetIndexShare.pullIfNeeded(local, share));
        assertEquals(
                Files.readString(share, StandardCharsets.UTF_8),
                Files.readString(local, StandardCharsets.UTF_8));
    }

    @Test
    void publish_failsWhenLocalCompleteMissing(@TempDir Path tmp) {
        Path local = tmp.resolve("missing.csv");
        Path share = tmp.resolve("share").resolve("KONAN.csv");
        IOException ex =
                assertThrows(
                        IOException.class, () -> InspectionSheetIndexShare.publish(local, share));
        assertTrue(ex.getMessage().contains("本索引"));
    }

    @Test
    void pullIfNeeded_skipsWhenLocalLarger(@TempDir Path tmp) throws Exception {
        Path local = tmp.resolve("KONAN.csv");
        Path share = tmp.resolve("share").resolve("KONAN.csv");
        Files.createDirectories(share.getParent());
        Files.writeString(share, "small\n", StandardCharsets.UTF_8);
        Files.writeString(local, "0123456789".repeat(20), StandardCharsets.UTF_8);
        assertEquals(
                InspectionSheetIndexShare.PullResult.SKIPPED,
                InspectionSheetIndexShare.pullIfNeeded(local, share));
        assertTrue(Files.readString(local, StandardCharsets.UTF_8).startsWith("0123456789"));
    }
}
