package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;
import java.time.temporal.ChronoUnit;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class AladdinEntryOpenLatestPolicyTest {

    @TempDir
    Path tempDir;

    @Test
    void resolve_missingFile_disallowsOpenWithNotGeneratedBadge() throws Exception {
        Path missing = tempDir.resolve("missing.xlsx");

        AladdinEntryOpenLatestPolicy.State state =
                AladdinEntryOpenLatestPolicy.resolve(missing, Instant.now());

        assertFalse(state.openAllowed());
        assertEquals(AladdinEntryOpenLatestPolicy.BADGE_NOT_GENERATED, state.badgeText());
        assertFalse(state.highlightGenerationsButton());
    }

    @Test
    void resolve_withinFifteenMinutes_allowsOpenWithCountdownBadge() throws Exception {
        Path latest = tempDir.resolve("latest.xlsx");
        Files.writeString(latest, "x");
        Instant now = Instant.parse("2026-07-30T12:00:00Z");
        Files.setLastModifiedTime(
                latest, FileTime.from(now.minus(14, ChronoUnit.MINUTES)));

        AladdinEntryOpenLatestPolicy.State state =
                AladdinEntryOpenLatestPolicy.resolve(latest, now);

        assertTrue(state.openAllowed());
        assertEquals("あと 60秒", state.badgeText());
        assertFalse(state.highlightGenerationsButton());
    }

    @Test
    void resolve_afterFifteenMinutes_disallowsOpenAndHighlightsGenerations() throws Exception {
        Path latest = tempDir.resolve("latest.xlsx");
        Files.writeString(latest, "x");
        Instant now = Instant.parse("2026-07-30T12:00:00Z");
        Files.setLastModifiedTime(
                latest, FileTime.from(now.minus(15, ChronoUnit.MINUTES)));

        AladdinEntryOpenLatestPolicy.State state =
                AladdinEntryOpenLatestPolicy.resolve(latest, now);

        assertFalse(state.openAllowed());
        assertEquals(AladdinEntryOpenLatestPolicy.BADGE_EXPIRED, state.badgeText());
        assertTrue(state.highlightGenerationsButton());
    }

    @Test
    void formatCountdownBadge_neverShowsNegativeSeconds() {
        assertEquals("あと 0秒", AladdinEntryOpenLatestPolicy.formatCountdownBadge(-5));
        assertEquals("あと 900秒", AladdinEntryOpenLatestPolicy.formatCountdownBadge(900));
    }
}
