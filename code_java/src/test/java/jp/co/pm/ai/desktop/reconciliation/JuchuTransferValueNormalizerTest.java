package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;

import java.time.LocalDate;

import org.junit.jupiter.api.Test;

class JuchuTransferValueNormalizerTest {

    @Test
    void parseLocalDate_monthDayUsesReferenceYear() {
        assertEquals(
                LocalDate.of(2026, 7, 15),
                JuchuTransferValueNormalizer.parseLocalDate(
                        "7/15", LocalDate.of(2026, 7, 9)));
    }

    @Test
    void parseLocalDate_monthDayWithWeekdaySuffix() {
        assertEquals(
                LocalDate.of(2026, 6, 10),
                JuchuTransferValueNormalizer.parseLocalDate(
                        "6/10（水）", LocalDate.of(2026, 6, 1)));
    }

    @Test
    void normalizeDateVal_monthDayBecomesIso() {
        assertEquals(
                "2026-07-15",
                JuchuTransferValueNormalizer.normalizeDateVal("7/15"));
    }

    @Test
    void parseLocalDate_japaneseMonthDayUsesReferenceYear() {
        assertEquals(
                LocalDate.of(2026, 7, 7),
                JuchuTransferValueNormalizer.parseLocalDate(
                        "7月7日", LocalDate.of(2026, 7, 7)));
    }

    @Test
    void parseLocalDate_japaneseMonthDayWithWeekdaySuffix() {
        assertEquals(
                LocalDate.of(2026, 7, 7),
                JuchuTransferValueNormalizer.parseLocalDate(
                        "7月7日（月）", LocalDate.of(2026, 7, 1)));
    }

    @Test
    void parseLocalDate_japaneseFullDate() {
        assertEquals(
                LocalDate.of(2026, 7, 7),
                JuchuTransferValueNormalizer.parseLocalDate("2026年7月7日", null));
    }

    @Test
    void normalizeDateVal_japaneseMonthDayBecomesIso() {
        assertEquals(
                "2026-07-07",
                JuchuTransferValueNormalizer.normalizeDateVal("7月7日"));
    }

    @Test
    void parseMonthDayValue_blankReturnsNull() {
        assertNull(JuchuTransferValueNormalizer.parseMonthDayValue("", LocalDate.now()));
    }

    @Test
    void toHalfWidthAscii_convertsFullwidthLettersAndDigits() {
        assertEquals(
                "A123-1",
                JuchuTransferValueNormalizer.toHalfWidthAscii("Ａ１２３－１"));
        assertEquals(
                "20010-H600-1180X250",
                JuchuTransferValueNormalizer.toHalfWidthAscii("２００１０－Ｈ６００－１１８０Ｘ２５０"));
    }

    @Test
    void toHalfWidthAscii_preservesHalfwidthKatakana() {
        assertEquals(
                "ｽﾗｲｽ面",
                JuchuTransferValueNormalizer.toHalfWidthAscii("ｽﾗｲｽ面"));
    }

    @Test
    void normalizeNumeric_acceptsFullwidthDigits() {
        assertEquals(1230.0, JuchuTransferValueNormalizer.normalizeNumeric("１,２３０"));
        assertEquals(250.0, JuchuTransferValueNormalizer.normalizeNumeric("２５０"));
    }
}
