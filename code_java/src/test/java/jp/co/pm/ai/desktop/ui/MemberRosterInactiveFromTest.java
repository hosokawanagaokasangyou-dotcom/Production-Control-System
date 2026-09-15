package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.Optional;

import org.junit.jupiter.api.Test;

class MemberRosterInactiveFromTest {

    @Test
    void parseKeepsIsoDateAndDropsBlank() {
        assertEquals(Optional.of(LocalDate.of(2026, 9, 15)), MemberRosterInactiveFrom.parse("2026-09-15"));
        assertTrue(MemberRosterInactiveFrom.parse("").isEmpty());
        assertTrue(MemberRosterInactiveFrom.parse("not-a-date").isEmpty());
        assertTrue(MemberRosterInactiveFrom.parse(null).isEmpty());
    }

    @Test
    void activeOnIsFalseFromTransferDate() {
        LocalDate from = LocalDate.of(2026, 9, 15);
        assertTrue(MemberRosterInactiveFrom.activeOn(from, LocalDate.of(2026, 9, 14)));
        assertFalse(MemberRosterInactiveFrom.activeOn(from, LocalDate.of(2026, 9, 15)));
        assertTrue(MemberRosterInactiveFrom.activeOn(null, LocalDate.of(2026, 9, 15)));
    }

    @Test
    void visibleInMonthUsesFirstOfMonth() {
        LocalDate mid = LocalDate.of(2026, 9, 15);
        LocalDate first = LocalDate.of(2026, 9, 1);
        assertTrue(MemberRosterInactiveFrom.visibleInMonth(mid, 2026, 8));
        assertTrue(MemberRosterInactiveFrom.visibleInMonth(mid, 2026, 9));
        assertFalse(MemberRosterInactiveFrom.visibleInMonth(mid, 2026, 10));
        assertTrue(MemberRosterInactiveFrom.visibleInMonth(first, 2026, 8));
        assertFalse(MemberRosterInactiveFrom.visibleInMonth(first, 2026, 9));
        assertTrue(MemberRosterInactiveFrom.visibleInMonth(null, 2026, 10));
    }
}
