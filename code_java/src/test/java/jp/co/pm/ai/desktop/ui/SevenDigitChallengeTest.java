package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class SevenDigitChallengeTest {

    @Test
    void generate_returnsSevenDigitsWithoutLeadingZero() {
        String code = SevenDigitChallenge.generate(bound -> 0);
        assertEquals("1000000", code);
        assertEquals(7, code.length());

        String max = SevenDigitChallenge.generate(bound -> bound - 1);
        assertEquals("9999999", max);
    }

    @Test
    void matches_acceptsOnlyExactDigits() {
        assertTrue(SevenDigitChallenge.matches("1234567", "1234567"));
        assertFalse(SevenDigitChallenge.matches("1234567", "1234568"));
        assertFalse(SevenDigitChallenge.matches("1234567", "123456"));
        assertTrue(SevenDigitChallenge.matches("1234567", " 1234567 "));
        assertFalse(SevenDigitChallenge.matches("1234567", null));
    }
}
