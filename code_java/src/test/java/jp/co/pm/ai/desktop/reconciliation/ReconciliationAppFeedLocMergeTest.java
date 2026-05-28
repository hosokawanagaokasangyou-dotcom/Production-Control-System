package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class ReconciliationAppFeedLocMergeTest {

    @Test
    void mergeFeedLocOptionsFromPlanning_appendsDistinctValues() {
        ReconciliationApp app = new ReconciliationApp();
        int added =
                app.mergeFeedLocOptionsFromPlanning(
                        java.util.List.of("EC機　湖南", "SEC", "スリット機1　湖南"));
        assertTrue(added >= 2);
        assertTrue(
                app.snapshotComboChoices()
                        .optionsFor(RequestFormComboChoices.KEY_FEED_LOC)
                        .contains("EC機　湖南"));
        assertTrue(
                app.snapshotComboChoices()
                        .optionsFor(RequestFormComboChoices.KEY_FEED_LOC)
                        .contains("スリット機1　湖南"));
        assertEquals(0, app.mergeFeedLocOptionsFromPlanning(java.util.List.of("EC機　湖南")));
    }
}
