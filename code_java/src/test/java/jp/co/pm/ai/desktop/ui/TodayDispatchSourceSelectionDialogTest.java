package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicReference;

import javafx.application.Platform;
import javafx.scene.control.ToggleGroup;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.planning.stage2.source.NetworkSourceExtractionCatalog;
import jp.co.pm.ai.planning.stage2.source.NetworkSourceExtractionTimeSupport;
import jp.co.pm.ai.planning.stage2.source.Stage1SourcePairMatcher;

class TodayDispatchSourceSelectionDialogTest {
    @TempDir Path temp;

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @Test
    void rowExposesPlanAndDailyTextsForTableBinding() throws Exception {
        CountDownLatch done = new CountDownLatch(1);
        AtomicReference<AssertionError> failure = new AtomicReference<>();
        Platform.runLater(
                () -> {
                    try {
                        var plan =
                                new NetworkSourceExtractionCatalog.SourceEntry(
                                        temp.resolve("加工計画DATA_20260716_080500.xlsx"),
                                        LocalDateTime.of(2026, 7, 16, 8, 5),
                                        NetworkSourceExtractionTimeSupport.SourceKind.FILENAME,
                                        "加工計画DATA_20260716_080500.xlsx");
                        var daily =
                                new NetworkSourceExtractionCatalog.SourceEntry(
                                        temp.resolve("加工日報発行問合せ_20260716_164200.csv"),
                                        LocalDateTime.of(2026, 7, 16, 16, 42),
                                        NetworkSourceExtractionTimeSupport.SourceKind.FILENAME,
                                        "加工日報発行問合せ_20260716_164200.csv");
                        var pair =
                                new Stage1SourcePairMatcher.MatchedPair(
                                        plan, daily, 517L, false, List.of(daily));
                        var row = new TodayDispatchSourceSelectionDialog.Row(pair, new ToggleGroup());
                        assertEquals("08:05", row.getPlanTime());
                        assertEquals("加工計画DATA_20260716_080500.xlsx", row.getPlanFile());
                        assertEquals("16:42", row.getDailyTime());
                        assertEquals("517分", row.getDelta());
                        assertEquals("08:05", row.planTimeProperty().get());
                    } catch (AssertionError error) {
                        failure.set(error);
                    } finally {
                        done.countDown();
                    }
                });
        assertTrue(done.await(5, TimeUnit.SECONDS));
        if (failure.get() != null) {
            throw failure.get();
        }
    }

    @Test
    void missingSameDayRequiresManualSelectionAndCannotConfirm() {
        var pair = new Stage1SourcePairMatcher.MatchedPair(entry(temp.resolve("plan.xlsx")), null, Long.MAX_VALUE, true, List.of());
        assertTrue(TodayDispatchSourceSelectionDialog.requiresManualDailySelection(pair));
        assertFalse(TodayDispatchSourceSelectionDialog.canConfirmSelection(pair));
    }

    @Test
    void manualCsvBuildsConfirmableOverrideAndCancelDoesNot() throws Exception {
        var pair = new Stage1SourcePairMatcher.MatchedPair(entry(temp.resolve("plan.xlsx")), null, Long.MAX_VALUE, true, List.of());
        Path csv = temp.resolve("manual.csv");
        Files.writeString(csv, "a\nb\nc\n依頼NO,工程名\n");
        var selected = TodayDispatchSourceSelectionDialog.selectManualDailyReport(pair, csv);
        assertTrue(selected.isPresent());
        assertTrue(TodayDispatchSourceSelectionDialog.canConfirmSelection(selected.orElseThrow()));
        assertTrue(TodayDispatchSourceSelectionDialog.selectManualDailyReport(pair, null).isEmpty());
    }

    private static NetworkSourceExtractionCatalog.SourceEntry entry(Path path) {
        return new NetworkSourceExtractionCatalog.SourceEntry(path, LocalDateTime.of(2026, 7, 10, 8, 0), NetworkSourceExtractionTimeSupport.SourceKind.FILENAME, path.getFileName().toString());
    }
}
