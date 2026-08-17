package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class SevenDigitChallengeTest {

    @Test
    void generate_returnsTwelveDigitsWithoutLeadingZero() {
        String code = SevenDigitChallenge.generate(bound -> 0L);
        assertEquals("100000000000", code);
        assertEquals(12, code.length());

        String max = SevenDigitChallenge.generate(bound -> bound - 1);
        assertEquals("999999999999", max);
    }

    @Test
    void matches_acceptsOnlyExactDigits() {
        assertTrue(SevenDigitChallenge.matches("123456789012", "123456789012"));
        assertFalse(SevenDigitChallenge.matches("123456789012", "123456789013"));
        assertFalse(SevenDigitChallenge.matches("123456789012", "12345678901"));
        assertTrue(SevenDigitChallenge.matches("123456789012", " 123456789012 "));
        assertFalse(SevenDigitChallenge.matches("123456789012", null));
    }
}
