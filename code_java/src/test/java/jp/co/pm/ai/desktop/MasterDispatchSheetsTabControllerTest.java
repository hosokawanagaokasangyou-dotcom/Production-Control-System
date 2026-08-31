package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import org.junit.jupiter.api.Test;

class MasterDispatchSheetsTabControllerTest {

    @Test
    void focusKeysForMissingPairs_containsOnlyAddedEquipmentPairs() {
        Set<String> actual =
                MasterDispatchSheetsTabController.focusKeysForMissingPairs(
                        List.of(
                                new PlanTasksMissingSkillsColumnPrompt.MissingPair(
                                        "巻返し", "機1", "T-1"),
                                new PlanTasksMissingSkillsColumnPrompt.MissingPair(
                                        "分割", "スライス機1", "T-2")));

        assertEquals(
                new LinkedHashSet<>(Set.of("巻返し+機1", "分割+スライス機1")),
                actual);
    }
}
