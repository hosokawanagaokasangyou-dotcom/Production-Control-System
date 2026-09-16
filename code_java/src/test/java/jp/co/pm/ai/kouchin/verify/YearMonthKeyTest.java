package jp.co.pm.ai.kouchin.verify;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class YearMonthKeyTest {

    @Test
    @DisplayName("ファイル名の「yyyy年m月度」から対象年月を取得する")
    void parsesGatsudoFromFileName() {
        Optional<YearMonthKey> ym = YearMonthKey.parseGatsudo("後加工工賃明細（2026年7月度)V02.xlsx");
        assertTrue(ym.isPresent());
        assertEquals(new YearMonthKey(2026, 7), ym.get());
    }

    @Test
    @DisplayName("全角数字・全角スペースを含む表記も NFKC 正規化して読める")
    void parsesFullWidthGatsudo() {
        Optional<YearMonthKey> ym = YearMonthKey.parseGatsudo("２０２６年　１２月度　加工賃試算");
        assertEquals(Optional.of(new YearMonthKey(2026, 12)), ym);
    }

    @Test
    @DisplayName("RVSHEETyyyymm から対象年月を取得する")
    void parsesRvsheet() {
        assertEquals(Optional.of(new YearMonthKey(2026, 7)),
                YearMonthKey.parseRvsheet("RVSHEET202607.csv"));
        assertEquals(Optional.of(new YearMonthKey(2025, 1)),
                YearMonthKey.parseRvsheet("rvsheet202501.CSV"));
        assertTrue(YearMonthKey.parseRvsheet("RVSHEET.csv").isEmpty());
    }

    @Test
    @DisplayName("「対象年月 : yyyy年mm月」から対象年月を取得する")
    void parsesTaishoYearMonth() {
        assertEquals(Optional.of(new YearMonthKey(2026, 7)),
                YearMonthKey.parseYearMonth("対象年月 : 2026年07月"));
    }

    @Test
    @DisplayName("13月など不正な月は取り込まない")
    void rejectsInvalidMonth() {
        assertTrue(YearMonthKey.parseYearMonth("2026年13月").isEmpty());
        assertThrows(IllegalArgumentException.class, () -> new YearMonthKey(2026, 0));
    }

    @Test
    @DisplayName("年をまたぐ月加算ができる")
    void addsMonthsAcrossYears() {
        assertEquals(new YearMonthKey(2027, 1), new YearMonthKey(2026, 12).plusMonths(1));
        assertEquals(new YearMonthKey(2025, 12), new YearMonthKey(2026, 1).plusMonths(-1));
        assertEquals(new YearMonthKey(2026, 7), new YearMonthKey(2026, 1).plusMonths(6));
    }

    @Test
    @DisplayName("年月の大小比較と月差の判定ができる")
    void comparesByMonthIndex() {
        YearMonthKey june = new YearMonthKey(2026, 6);
        YearMonthKey july = new YearMonthKey(2026, 7);
        YearMonthKey lastJuly = new YearMonthKey(2025, 7);

        assertTrue(june.compareTo(july) < 0);
        assertTrue(july.compareTo(june) > 0);
        assertEquals(0, july.compareTo(new YearMonthKey(2026, 7)));
        assertEquals(1, july.monthIndex() - june.monthIndex());
        assertEquals(12, july.monthIndex() - lastJuly.monthIndex());
    }

    @Test
    @DisplayName("ラベル文字列を生成する")
    void buildsLabels() {
        YearMonthKey ym = new YearMonthKey(2026, 7);
        assertEquals("2026年7月度", ym.gatsudoLabel());
        assertEquals("2026年7月", ym.ymLabel());
        assertEquals("7月", ym.monthLabel());
    }
}
