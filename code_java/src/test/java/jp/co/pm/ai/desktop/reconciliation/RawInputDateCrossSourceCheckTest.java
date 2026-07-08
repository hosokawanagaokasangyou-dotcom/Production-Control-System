package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class RawInputDateCrossSourceCheckTest {

    @Test
    void allFourEqual_isMatch() {
        RawInputDateCrossSourceCheck.CrossSourceResult r =
                RawInputDateCrossSourceCheck.evaluate(
                        "2026/7/5", "2026/7/5", "2026/7/5", "2026/7/5", true);
        assertEquals(RawInputDateCrossSourceCheck.STATUS_MATCH, r.status());
    }

    @Test
    void oneMissing_isMismatch() {
        RawInputDateCrossSourceCheck.CrossSourceResult r =
                RawInputDateCrossSourceCheck.evaluate("2026/7/5", "2026/7/5", "", "2026/7/5", true);
        assertEquals(RawInputDateCrossSourceCheck.STATUS_MISMATCH, r.status());
    }

    @Test
    void differentDate_isMismatch() {
        RawInputDateCrossSourceCheck.CrossSourceResult r =
                RawInputDateCrossSourceCheck.evaluate(
                        "2026/7/5", "2026/7/6", "2026/7/5", "2026/7/5", true);
        assertEquals(RawInputDateCrossSourceCheck.STATUS_MISMATCH, r.status());
    }

    @Test
    void monthDayFormEquivalentToFullDate_isMatch() {
        RawInputDateCrossSourceCheck.CrossSourceResult r =
                RawInputDateCrossSourceCheck.evaluate(
                        "2026/7/5", "7/5", "2026-07-05", "2026/7/5", true);
        assertEquals(RawInputDateCrossSourceCheck.STATUS_MATCH, r.status());
    }

    @Test
    void multiLineSameUniqueDate_isMatch() {
        RawInputDateCrossSourceCheck.CrossSourceResult r =
                RawInputDateCrossSourceCheck.evaluate(
                        "2026/7/5\n2026/7/5", "2026/7/5", "2026/7/5", "7/5", true);
        assertEquals(RawInputDateCrossSourceCheck.STATUS_MATCH, r.status());
    }

    @Test
    void allBlank_isNa() {
        RawInputDateCrossSourceCheck.CrossSourceResult r =
                RawInputDateCrossSourceCheck.evaluate("", "", "", "", true);
        assertEquals(RawInputDateCrossSourceCheck.STATUS_NA, r.status());
    }

    @Test
    void aladdinNotLoadedAndOthersPresent_isMismatch() {
        RawInputDateCrossSourceCheck.CrossSourceResult r =
                RawInputDateCrossSourceCheck.evaluate("", "2026/7/5", "2026/7/5", "2026/7/5", false);
        assertEquals(RawInputDateCrossSourceCheck.STATUS_MISMATCH, r.status());
    }
}
