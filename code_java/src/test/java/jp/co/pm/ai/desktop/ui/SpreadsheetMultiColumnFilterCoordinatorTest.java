package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

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

    @Test
    void retainSelectionToSearchVisible_keepsLockedTitleValues() {
        Set<String> copy = new HashSet<>(Set.of("スライス", "OP3", "AS4"));
        SpreadsheetMultiColumnFilterCoordinator.retainSelectionToSearchVisible(
                copy, List.of("OP3"), "OP", Set.of("スライス"));
        assertEquals(Set.of("OP3", "スライス"), copy);
    }

    @Test
    void lastAlwaysVisibleGridRow_consecutiveFixedAfterFilter() {
        assertEquals(
                2,
                SpreadsheetMultiColumnFilterCoordinator.lastAlwaysVisibleGridRow(
                        0, List.of(0, 1, 2)));
        assertEquals(
                0, SpreadsheetMultiColumnFilterCoordinator.lastAlwaysVisibleGridRow(0, List.of(0)));
        assertEquals(
                0,
                SpreadsheetMultiColumnFilterCoordinator.lastAlwaysVisibleGridRow(
                        0, List.of(0, 5)));
    }

    @Test
    void lockedColumnFilterValues_nonBlankTitleRowsOnly() {
        List<String> col = List.of("", "スライス", "スライス機３", "OP3", "");
        Set<String> locked =
                SpreadsheetMultiColumnFilterCoordinator.lockedColumnFilterValues(col, 0, 2);
        assertEquals(Set.of("スライス", "スライス機３"), locked);
    }

    @Test
    void shouldHideGridRow_neverHidesAlwaysVisibleTitleRows() {
        assertFalse(
                SpreadsheetMultiColumnFilterCoordinator.shouldHideGridRow(1, 2, true, true, true));
        assertTrue(
                SpreadsheetMultiColumnFilterCoordinator.shouldHideGridRow(3, 2, true, false, false));
        assertFalse(
                SpreadsheetMultiColumnFilterCoordinator.shouldHideGridRow(3, 2, false, false, false));
    }
}
