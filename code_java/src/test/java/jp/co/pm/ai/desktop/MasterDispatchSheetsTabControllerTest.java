package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.ArrayList;
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

    @Test
    void combinationDocumentIndexesFromModelRows_mapsGridModelRowNotCompressedViewRow() {
        assertEquals(
                Set.of(624),
                MasterDispatchSheetsTabController.combinationDocumentIndexesFromModelRows(
                        List.of(624), 1));
        assertEquals(
                Set.of(40),
                MasterDispatchSheetsTabController.combinationDocumentIndexesFromModelRows(
                        List.of(40), 1));
    }

    @Test
    void combinationDeleteLabels_followModelMappedDocumentIndexes() {
        List<List<String>> combo = new ArrayList<>();
        combo.add(List.of("組み合わせ行ID", "工程名", "機械名", "編集ロック"));
        for (int id = 1; id <= 624; id++) {
            if (id == 40) {
                combo.add(List.of(String.valueOf(id), "スリット", "スリット機3", ""));
            } else if (id == 624) {
                combo.add(List.of(String.valueOf(id), "分割", "LAC/EC機", ""));
            } else {
                combo.add(List.of(String.valueOf(id), "他", "機" + id, ""));
            }
        }
        Set<Integer> fromViewMistakenAsModel =
                MasterDispatchSheetsTabController.combinationDocumentIndexesFromModelRows(
                        List.of(40), 1);
        Set<Integer> fromModelRow624 =
                MasterDispatchSheetsTabController.combinationDocumentIndexesFromModelRows(
                        List.of(624), 1);
        assertEquals(
                List.of("・スリット × スリット機3"),
                MasterDispatchSheetsTabController.combinationDeleteLabels(
                        combo, fromViewMistakenAsModel));
        assertEquals(
                List.of("・分割 × LAC/EC機"),
                MasterDispatchSheetsTabController.combinationDeleteLabels(
                        combo, fromModelRow624));
    }
}

