package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

class SevenDigitChallengeTest {

    @Test
    void generate_returnsFourDigitsWithoutLeadingZero() {
        String code = SevenDigitChallenge.generate(bound -> 0L);
        assertEquals("1000", code);
        assertEquals(4, code.length());

        String max = SevenDigitChallenge.generate(bound -> bound - 1);
        assertEquals("9999", max);
    }

    @Test
    void matches_acceptsOnlyExactDigits() {
        assertTrue(SevenDigitChallenge.matches("1234", "1234"));
        assertFalse(SevenDigitChallenge.matches("1234", "1235"));
        assertFalse(SevenDigitChallenge.matches("1234", "123"));
        assertTrue(SevenDigitChallenge.matches("1234", " 1234 "));
        assertFalse(SevenDigitChallenge.matches("1234", null));
    }
}
