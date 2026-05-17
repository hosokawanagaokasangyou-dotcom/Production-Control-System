package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.List;
import java.util.Optional;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

class NetworkSourceFileReloadCacheTest {

    @AfterEach
    void tearDown() {
        NetworkSourceFileReloadCache.clearAll();
    }

    @Test
    void matchByFileNameOnly_notFullPath() {
        Path a = Path.of("/data/a/plan.xlsx");
        Path b = Path.of("/other/plan.xlsx");
        NetworkSourceFileReloadCache.storeAladdin(
                a,
                true,
                List.of("Sheet1"),
                0,
                new PlanInputTabularIo.TabularSheet(List.of("h"), List.of(List.of("1"))));

        assertTrue(NetworkSourceFileReloadCache.matchAladdin(b).isPresent());
        assertEquals("plan.xlsx", NetworkSourceFileReloadCache.matchAladdin(b).orElseThrow().fileName());
    }

    @Test
    void differentFileNameDoesNotMatch() {
        NetworkSourceFileReloadCache.storeActuals(
                Path.of("/x/foo.csv"),
                false,
                List.of(),
                0,
                new PlanInputTabularIo.TabularSheet(List.of("c"), List.of()));

        assertTrue(NetworkSourceFileReloadCache.matchActuals(Path.of("/y/bar.csv")).isEmpty());
    }

    @Test
    void snapshotIsDefensiveCopy() {
        NetworkSourceFileReloadCache.storeAladdin(
                Path.of("t.xlsx"),
                false,
                List.of(),
                0,
                new PlanInputTabularIo.TabularSheet(List.of("h"), List.of(List.of("v"))));
        NetworkSourceFileReloadCache.Snapshot snap =
                NetworkSourceFileReloadCache.matchAladdin(Path.of("t.xlsx")).orElseThrow();
        PlanInputTabularIo.TabularSheet tab = snap.toTabularSheet();
        assertEquals(List.of("v"), tab.rows().get(0));
    }
}
