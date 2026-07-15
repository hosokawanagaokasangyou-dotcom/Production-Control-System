package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.planning.stage2.source.NetworkSourceExtractionCatalog;
import jp.co.pm.ai.planning.stage2.source.NetworkSourceExtractionTimeSupport;
import jp.co.pm.ai.planning.stage2.source.Stage1SourcePairMatcher;

class TodayDispatchSourceSelectionDialogTest {
    @TempDir Path temp;

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
