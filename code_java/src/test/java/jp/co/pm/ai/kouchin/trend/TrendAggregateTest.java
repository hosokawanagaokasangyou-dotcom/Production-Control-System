package jp.co.pm.ai.kouchin.trend;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.kouchin.verify.YearMonthKey;

class TrendAggregateTest {

    private static final YearMonthKey YM_JUL = new YearMonthKey(2026, 7);
    private static final YearMonthKey YM_AUG = new YearMonthKey(2026, 8);

    @Test
    @DisplayName("国分 スライス1/3 と スライス を同一グループにする")
    void processGroupMapsSliceAliases() {
        assertEquals("スライス", TrendAggregate.processGroup("スライス1"));
        assertEquals("スライス", TrendAggregate.processGroup("スライス3"));
        assertEquals("スライス", TrendAggregate.processGroup("スライス１"));
        assertEquals("スライス", TrendAggregate.processGroup("スライス"));
        assertEquals("スライス", TrendAggregate.processGroup("スライス1/3"));
    }

    @Test
    @DisplayName("グループ横断でスライス1+3を合算する")
    void groupedCrossSumsSlice1And3() {
        TrendMonthData kokubu =
                month(
                        YM_AUG,
                        Map.of(
                                "スライス1", new TrendMetric(100, 10),
                                "スライス3", new TrendMetric(50, 5)));
        TrendMonthData konan =
                month(YM_AUG, Map.of("スライス", new TrendMetric(80, 8)));
        TrendCrossMatrix g =
                TrendAggregate.buildGroupedCross(
                        List.of(kokubu), List.of(konan), List.of(YM_AUG), "wage");
        assertTrue(g.processes().contains("スライス"));
        assertEquals(150.0, g.values().get("国分").get("スライス").get(YM_AUG), 1e-9);
        assertEquals(80.0, g.values().get("湖南").get("スライス").get(YM_AUG), 1e-9);
    }

    @Test
    @DisplayName("前月差の異符号かつ近い大きさは強い相殺")
    void shiftBadgeStrongOffset() {
        assertEquals("強い相殺", TrendAggregate.shiftBadge(100.0, -80.0));
        assertEquals("同方向", TrendAggregate.shiftBadge(100.0, 50.0));
        assertNull(TrendAggregate.shiftBadge(100.0, null));
    }

    @Test
    @DisplayName("振替疑いは相殺量の大きい工程を先にする")
    void rankShiftSuspectsOrdersByOffset() {
        TrendCrossMatrix grouped =
                new TrendCrossMatrix(
                        List.of("スライス", "スリット"),
                        List.of(YM_JUL, YM_AUG),
                        "wage",
                        Map.of(
                                "国分",
                                Map.of(
                                        "スライス", Map.of(YM_JUL, 100.0, YM_AUG, 180.0),
                                        "スリット", Map.of(YM_JUL, 50.0, YM_AUG, 55.0)),
                                "湖南",
                                Map.of(
                                        "スライス", Map.of(YM_JUL, 200.0, YM_AUG, 120.0),
                                        "スリット", Map.of(YM_JUL, 40.0, YM_AUG, 42.0))),
                        true);
        List<TrendShiftRow> rows = TrendAggregate.rankShiftSuspects(grouped, "wage");
        assertEquals("スライス", rows.get(0).process());
        assertEquals("強い相殺", rows.get(0).badge());
    }

    private static TrendMonthData month(YearMonthKey ym, Map<String, TrendMetric> kinds) {
        return new TrendMonthData(ym, null, kinds, Map.of(), List.of());
    }
}
