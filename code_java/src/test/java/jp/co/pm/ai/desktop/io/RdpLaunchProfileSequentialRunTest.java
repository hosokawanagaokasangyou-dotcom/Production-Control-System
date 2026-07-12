package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.ArrayList;
import java.util.List;

import org.junit.jupiter.api.Test;

class RdpLaunchProfileSequentialRunTest {

    @Test
    void toggleSelection_addsThenRemovesInClickOrder() {
        List<Integer> order = new ArrayList<>();
        order = new ArrayList<>(RdpLaunchProfileSequentialRun.toggleSelection(order, 2));
        order = new ArrayList<>(RdpLaunchProfileSequentialRun.toggleSelection(order, 5));
        assertEquals(List.of(2, 5), order);

        order = new ArrayList<>(RdpLaunchProfileSequentialRun.toggleSelection(order, 2));
        assertEquals(List.of(5), order);
    }

    @Test
    void quickButtonLabel_showsOrderMarkerWhenSelected() {
        assertEquals(
                "② 2: 問合せ",
                RdpLaunchProfileSequentialRun.quickButtonLabel("2: 問合せ", 2));
        assertEquals(
                "2: 問合せ",
                RdpLaunchProfileSequentialRun.quickButtonLabel("2: 問合せ", -1));
    }

    @Test
    void launchButtonTextIdle_reflectsSelectionCount() {
        assertEquals(
                "連続実行するタスクを選択",
                RdpLaunchProfileSequentialRun.launchButtonTextIdle(0));
        assertEquals(
                "連続実行を開始（3件）",
                RdpLaunchProfileSequentialRun.launchButtonTextIdle(3));
    }

    @Test
    void progressStatusText_includesProfileLabel() {
        assertEquals(
                "連続実行 2/5: アラジン起動のみ",
                RdpLaunchProfileSequentialRun.progressStatusText(
                        2, 5, "アラジン起動のみ"));
    }

    @Test
    void normalizeSelection_preservesOrderWithoutDuplicates() {
        List<Integer> normalized =
                RdpLaunchProfileSequentialRun.normalizeSelection(List.of(1, 2, 2, 3, null, 0));
        assertEquals(List.of(1, 2, 3), normalized);
    }

    @Test
    void toggleSelection_signOutOnlyAllowedOnlyWhenEmpty() {
        List<Integer> order = new ArrayList<>();
        order = new ArrayList<>(RdpLaunchProfileSequentialRun.toggleSelection(order, 99));
        assertEquals(List.of(99), order);

        order = new ArrayList<>(RdpLaunchProfileSequentialRun.toggleSelection(order, 2));
        assertEquals(List.of(99, 2), order);

        List<Integer> withRpa = List.of(2);
        assertFalse(RdpLaunchProfileSequentialRun.canAddProfileToSelection(withRpa, 99));
        assertEquals(
                withRpa,
                RdpLaunchProfileSequentialRun.toggleSelection(withRpa, 99));
    }

    @Test
    void validateSignOutOnlyAtHead_rejectsWhenNotFirst() {
        assertTrue(
                RdpLaunchProfileSequentialRun.validateSignOutOnlyAtHead(List.of(2, 99))
                        .isPresent());
        assertTrue(
                RdpLaunchProfileSequentialRun.validateSignOutOnlyAtHead(List.of(99, 2))
                        .isEmpty());
    }

    @Test
    void selectionRequiresAladdinCredentials_skipsSignOutOnly() {
        assertFalse(
                RdpLaunchProfileSequentialRun.selectionRequiresAladdinCredentials(List.of(99)));
        assertTrue(
                RdpLaunchProfileSequentialRun.selectionRequiresAladdinCredentials(
                        List.of(99, 2)));
    }

    @Test
    void selectionOrderIndex_isOneBased() {
        List<Integer> order = List.of(3, 1, 5);
        assertEquals(2, RdpLaunchProfileSequentialRun.selectionOrderIndex(order, 1));
        assertEquals(-1, RdpLaunchProfileSequentialRun.selectionOrderIndex(order, 4));
    }
}
