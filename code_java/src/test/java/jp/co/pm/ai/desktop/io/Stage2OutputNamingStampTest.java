package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDateTime;

import org.junit.jupiter.api.Test;

class Stage2OutputNamingStampTest {

    @Test
    void formatStage2Stamp_matchesLengthAndDigitSuffix() {
        LocalDateTime base = LocalDateTime.of(2026, 5, 13, 8, 36, 24, 123_000_000);
        LocalDateTime run = LocalDateTime.of(2026, 5, 13, 8, 36, 24, 456_789_000);
        String stamp = Stage2OutputNaming.formatStage2Stamp(base, run);
        assertEquals(Stage2OutputNaming.STAMP_DIGITS, stamp.length());
        assertTrue(stamp.startsWith("260513083624"));
        assertTrue(stamp.chars().allMatch(Character::isDigit));
    }
}
