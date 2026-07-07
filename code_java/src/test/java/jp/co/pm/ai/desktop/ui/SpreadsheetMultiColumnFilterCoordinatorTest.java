package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.junit.jupiter.api.Test;

class SpreadsheetMultiColumnFilterCoordinatorTest {

    @Test
    void retainSelectionToSearchVisible_removesItemsNotInFilteredList() {
        Set<String> copy = new HashSet<>(Set.of("W7-7", "W7-8", "W8-1"));
        SpreadsheetMultiColumnFilterCoordinator.retainSelectionToSearchVisible(
                copy, List.of("W7-7"), "W7-7");
        assertEquals(Set.of("W7-7"), copy);
    }

    @Test
    void retainSelectionToSearchVisible_noOpWhenSearchEmpty() {
        Set<String> copy = new HashSet<>(Set.of("W7-7", "W7-8"));
        SpreadsheetMultiColumnFilterCoordinator.retainSelectionToSearchVisible(
                copy, List.of("W7-7"), "");
        assertEquals(Set.of("W7-7", "W7-8"), copy);
    }
}
