package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.Map;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class AladdinDataCoverTest {

    private static final YearMonthKey AUG = new YearMonthKey(2026, 8);

    @Test
    @DisplayName("対象年月が一致し加工金額があるときは対象月データを含む")
    void coversWhenYmMatchesAndHasRows() {
        AladdinData data = new AladdinData(Map.of("C8-1", 1000.0), "2026年08月", AUG);
        assertTrue(data.coversTargetMonth(AUG));
    }

    @Test
    @DisplayName("加工金額が0件なら対象月データを含まない")
    void emptyRowsDoNotCover() {
        AladdinData data = new AladdinData(Map.of(), "2026年08月", AUG);
        assertFalse(data.coversTargetMonth(AUG));
    }

    @Test
    @DisplayName("対象年月が違うならデータを含まない")
    void mismatchedYmDoesNotCover() {
        AladdinData data = new AladdinData(Map.of("C7-1", 1000.0), "2026年07月", new YearMonthKey(2026, 7));
        assertFalse(data.coversTargetMonth(AUG));
    }

    @Test
    @DisplayName("対象年月が読めないなら対象月データを含まない")
    void unreadableYmDoesNotCover() {
        AladdinData data = new AladdinData(Map.of("C8-1", 1000.0), "", null);
        assertFalse(data.coversTargetMonth(AUG));
    }

    @Test
    @DisplayName("①の対象月が不明で行があるときは月整合を要求しない")
    void unknownTargetYmWithRowsCovers() {
        AladdinData data = new AladdinData(Map.of("C8-1", 1000.0), "2026年08月", AUG);
        assertTrue(data.coversTargetMonth(null));
    }

    @Test
    @DisplayName("nullデータは対象月データを含まない")
    void nullDataDoesNotCover() {
        assertFalse(AladdinData.coversTargetMonth(null, AUG));
    }
}
