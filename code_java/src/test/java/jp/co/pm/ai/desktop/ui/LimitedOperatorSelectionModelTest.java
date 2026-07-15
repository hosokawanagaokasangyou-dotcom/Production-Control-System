package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;

import org.junit.jupiter.api.Test;

class LimitedOperatorSelectionModelTest {

    @Test
    void initialSelectionIsRestoredAndLaterChecksAppendInSelectionOrder() {
        LimitedOperatorSelectionModel model =
                new LimitedOperatorSelectionModel(
                        List.of("山田", "佐藤", "鈴木"), List.of("佐藤", "山田"));

        model.setSelected("佐藤", false);
        model.setSelected("鈴木", true);
        model.setSelected("佐藤", true);

        assertEquals(List.of("山田", "鈴木", "佐藤"), model.selectedNames());
    }

    @Test
    void searchAndSelectAllOperateOnVisibleCandidates() {
        LimitedOperatorSelectionModel model =
                new LimitedOperatorSelectionModel(
                        List.of("山田 太郎", "佐藤 花子", "山本 次郎"), List.of());

        assertEquals(List.of("山田 太郎", "山本 次郎"), model.filteredCandidates("山"));
        model.selectAll(model.filteredCandidates("山"));
        assertEquals(List.of("山田 太郎", "山本 次郎"), model.selectedNames());
        model.clearAll();
        assertEquals(List.of(), model.selectedNames());
    }

    @Test
    void existingOutOfCandidateNameRemainsVisibleAndSelectedUntilExplicitlyCleared() {
        LimitedOperatorSelectionModel model =
                new LimitedOperatorSelectionModel(
                        List.of("山田"), List.of("候補外の旧担当", "山田"));

        assertEquals(List.of("候補外の旧担当", "山田"), model.selectedNames());
        assertEquals(
                List.of("山田", "候補外の旧担当"),
                model.filteredDisplayNames(""));
        assertFalse(model.isCandidate("候補外の旧担当"));
        assertTrue(model.hasSelectedOutOfCandidateNames());
        assertThrows(IllegalStateException.class, model::validateConfirmable);

        model.setSelected("候補外の旧担当", false);

        assertEquals(List.of("山田"), model.selectedNames());
        assertFalse(model.hasSelectedOutOfCandidateNames());
        model.validateConfirmable();
    }

    @Test
    void outOfCandidateNameCannotBeNewlySelectedAfterItIsCleared() {
        LimitedOperatorSelectionModel model =
                new LimitedOperatorSelectionModel(
                        List.of("山田"), List.of("候補外の旧担当"));

        model.setSelected("候補外の旧担当", false);
        model.setSelected("候補外の旧担当", true);

        assertEquals(List.of(), model.selectedNames());
    }
}
