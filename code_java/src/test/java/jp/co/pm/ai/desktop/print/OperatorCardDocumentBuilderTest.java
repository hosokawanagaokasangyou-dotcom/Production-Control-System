package jp.co.pm.ai.desktop.print;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;

class OperatorCardDocumentBuilderTest {

    @Test
    void mergeTimeRange_joins_first_start_and_last_end() {
        assertEquals(
                "08:00-09:30",
                OperatorCardDocumentBuilder.mergeTimeRange(
                        "08:00-08:10", "09:20-09:30"));
    }

    @Test
    void resolveThreeDayColumns_finds_mm_dd_headers() {
        List<String> cols =
                List.of(
                        "時間帯",
                        "05/07 (Thu)",
                        "05/08 (Fri)",
                        "05/09 (Sat)");
        LocalDate start = LocalDate.of(2026, 5, 7);
        List<String> three = OperatorCardDocumentBuilder.resolveThreeDayColumns(cols, start);
        assertEquals(3, three.size());
        assertEquals("05/07 (Thu)", three.get(0));
        assertEquals("05/08 (Fri)", three.get(1));
        assertEquals("05/09 (Sat)", three.get(2));
    }

    @Test
    void parseColumnDate_reads_month_day() {
        LocalDate d =
                OperatorCardDocumentBuilder.parseColumnDate("05/08 (Fri)", 2026);
        assertNotNull(d);
        assertEquals(LocalDate.of(2026, 5, 8), d);
    }

    @Test
    void formatDaySectionTitle_weekday_is_japanese_short() {
        LocalDate thu = LocalDate.of(2026, 5, 7);
        assertEquals(
                "2026-05-07  05/07（木）",
                OperatorCardPreviewFactory.formatDaySectionTitle(thu));
    }

    @Test
    void canonicalTeamCellKey_ignores_whitespace_between_tokens() {
        String a =
                "[Y5-1] \u30b9\u30e9\u30a4\u30b9+\u30b9\u30e9\u30a4\u30b9\u6a5f1\u3000\u6e56\u5357(\u88dc)";
        String b =
                "[Y5-1] \u30b9\u30e9\u30a4\u30b9+\u30b9\u30e9\u30a4\u30b9\u6a5f1 \u6e56\u5357(\u88dc)";
        assertEquals(
                OperatorCardDocumentBuilder.canonicalTeamCellKey(a),
                OperatorCardDocumentBuilder.canonicalTeamCellKey(b));
    }

    @Test
    void parseDispatchTableDay_accepts_iso_and_slash() {
        assertEquals(
                LocalDate.of(2026, 5, 7),
                OperatorCardDocumentBuilder.parseDispatchTableDay("2026-05-07"));
        assertEquals(
                LocalDate.of(2026, 5, 7),
                OperatorCardDocumentBuilder.parseDispatchTableDay("2026/05/07"));
    }

    @Test
    void findDispatchRow_matches_json_dispatch_day_with_slashes() {
        List<Map<String, String>> rows =
                List.of(
                        Map.ofEntries(
                                Map.entry("\u914d\u53f0\u65e5", "2026/05/07"),
                                Map.entry("\u4f9d\u983cNO", "Y5-1"),
                                Map.entry("\u5de5\u7a0b\u540d", "\u30b9\u30e9\u30a4\u30b9"),
                                Map.entry(
                                        "\u6a5f\u68b0\u540d",
                                        "\u30b9\u30e9\u30a4\u30b9\u6a5f1\u3000\u6e56\u5357"),
                                Map.entry("\u5f53\u65e5\u914d\u53f0\u6570\u91cf", "3600"),
                                Map.entry("\u63db\u7b97\u6570\u91cf", "9000")));
        Map<String, String> hit =
                OperatorCardDocumentBuilder.findDispatchRow(
                        rows,
                        LocalDate.of(2026, 5, 7),
                        "Y5-1",
                        "\u30b9\u30e9\u30a4\u30b9",
                        "\u30b9\u30e9\u30a4\u30b9\u6a5f1\u3000\u6e56\u5357");
        assertNotNull(hit);
        assertEquals("3600", hit.get("\u5f53\u65e5\u914d\u53f0\u6570\u91cf"));
        assertEquals("9000", hit.get("\u63db\u7b97\u6570\u91cf"));
    }
}
