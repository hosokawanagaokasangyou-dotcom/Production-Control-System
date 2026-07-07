package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class TaskIdLeadingAlphaPrefixTest {

    @Test
    void extractsTwoLetterPrefix() {
        assertEquals("JR", TaskIdLeadingAlphaPrefix.extract("JR260703"));
        assertEquals("GB", TaskIdLeadingAlphaPrefix.extract("gb6064"));
        assertEquals("TT", TaskIdLeadingAlphaPrefix.extract("TT-1"));
    }

    @Test
    void extractsOneLetterWhenSecondIsNotLetter() {
        assertEquals("C", TaskIdLeadingAlphaPrefix.extract("C7-10"));
        assertEquals("W", TaskIdLeadingAlphaPrefix.extract("W5-13"));
        assertEquals("E", TaskIdLeadingAlphaPrefix.extract("E123"));
    }

    @Test
    void blankOrNonAlphaReturnsOther() {
        assertEquals(TaskIdLeadingAlphaPrefix.OTHER, TaskIdLeadingAlphaPrefix.extract(""));
        assertEquals(TaskIdLeadingAlphaPrefix.OTHER, TaskIdLeadingAlphaPrefix.extract("123"));
    }
}
