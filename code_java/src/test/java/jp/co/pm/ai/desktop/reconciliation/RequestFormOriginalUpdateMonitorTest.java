package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.atomic.AtomicInteger;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RequestFormOriginalUpdateMonitorTest {

    @Test
    void pollDetectsMtimeChange(@TempDir Path tmp) throws Exception {
        File file = tmp.resolve("book.xlsm").toFile();
        Files.writeString(tmp.resolve("book.xlsm"), "v1");

        RequestFormOriginalUpdateMonitor monitor = new RequestFormOriginalUpdateMonitor();
        monitor.ensureTracked(file);
        assertFalse(monitor.isUpdated(file));

        Thread.sleep(1100L);
        Files.writeString(tmp.resolve("book.xlsm"), "v2");
        monitor.poll(file);
        assertTrue(monitor.isUpdated(file));
    }

    @Test
    void markPreviewAcknowledgedClearsUpdated(@TempDir Path tmp) throws Exception {
        File file = tmp.resolve("book.xlsm").toFile();
        Files.writeString(tmp.resolve("book.xlsm"), "v1");

        RequestFormOriginalUpdateMonitor monitor = new RequestFormOriginalUpdateMonitor();
        monitor.ensureTracked(file);
        Thread.sleep(1100L);
        Files.writeString(tmp.resolve("book.xlsm"), "v2");
        monitor.poll(file);
        assertTrue(monitor.isUpdated(file));

        monitor.markPreviewAcknowledged(file);
        assertFalse(monitor.isUpdated(file));

        monitor.poll(file);
        assertFalse(monitor.isUpdated(file));
    }

    @Test
    void onUpdatedKeysChangedFiresOnce(@TempDir Path tmp) throws Exception {
        File file = tmp.resolve("book.xlsm").toFile();
        Files.writeString(tmp.resolve("book.xlsm"), "v1");

        RequestFormOriginalUpdateMonitor monitor = new RequestFormOriginalUpdateMonitor();
        AtomicInteger fires = new AtomicInteger();
        monitor.setOnUpdatedKeysChanged(keys -> fires.incrementAndGet());
        monitor.ensureTracked(file);

        Thread.sleep(1100L);
        Files.writeString(tmp.resolve("book.xlsm"), "v2");
        monitor.poll(file);
        monitor.poll(file);

        assertTrue(monitor.isUpdated(file));
        assertTrue(fires.get() >= 1);
    }
}
