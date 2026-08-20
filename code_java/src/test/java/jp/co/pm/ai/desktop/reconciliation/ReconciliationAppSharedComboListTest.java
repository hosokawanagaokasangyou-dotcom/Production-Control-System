package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.List;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.junit.jupiter.api.Test;

class ReconciliationAppSharedComboListTest {

    @Test
    void observableListSetAllSelfClearsItems() {
        ObservableList<String> shared = FXCollections.observableArrayList("1", "2", "3");
        shared.setAll(shared);
        assertTrue(shared.isEmpty());
    }

    @Test
    void replaceItemsUnlessSameList_skipsWhenTargetIsSource() {
        ObservableList<String> shared = FXCollections.observableArrayList("EC", "SEC", "ｽﾗｲｽ");
        ReconciliationApp.replaceItemsUnlessSameList(shared, shared);
        assertEquals(List.of("EC", "SEC", "ｽﾗｲｽ"), List.copyOf(shared));
    }

    @Test
    void replaceItemsUnlessSameList_copiesWhenListsDiffer() {
        ObservableList<String> target = FXCollections.observableArrayList("old");
        ObservableList<String> source = FXCollections.observableArrayList("有", "無", "-");
        ReconciliationApp.replaceItemsUnlessSameList(target, source);
        assertEquals(List.of("有", "無", "-"), List.copyOf(target));
    }
}
