package jp.co.pm.ai.desktop.io.gantt;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

class PersonNameHeuristicsTest {

    @Test
    void acceptsJapaneseNames() {
        Assertions.assertTrue(PersonNameHeuristics.looksLikePersonName("田中一郎"));
        Assertions.assertTrue(PersonNameHeuristics.looksLikePersonName("山田 太郎"));
        Assertions.assertTrue(PersonNameHeuristics.looksLikePersonName("サトウ"));
    }

    @Test
    void rejectsMachineCodesAndQuantities() {
        Assertions.assertFalse(PersonNameHeuristics.looksLikePersonName("[Y5-135]"));
        Assertions.assertFalse(PersonNameHeuristics.looksLikePersonName("[Y"));
        Assertions.assertFalse(PersonNameHeuristics.looksLikePersonName("Y5-135"));
        Assertions.assertFalse(PersonNameHeuristics.looksLikePersonName("LAC"));
        Assertions.assertFalse(PersonNameHeuristics.looksLikePersonName("2000"));
        Assertions.assertFalse(PersonNameHeuristics.looksLikePersonName("2000m"));
    }
}
