package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeParseException;
import java.util.Optional;

/** 名簿の {@code inactive_from} / {@code returned_on}（職場異動・復帰）判定。Python 側と同規則。 */
public final class MemberRosterInactiveFrom {

    private MemberRosterInactiveFrom() {}

    public static Optional<LocalDate> parse(String raw) {
        if (raw == null) {
            return Optional.empty();
        }
        String text = raw.strip();
        if (text.isEmpty()) {
            return Optional.empty();
        }
        try {
            return Optional.of(LocalDate.parse(text));
        } catch (DateTimeParseException ex) {
            return Optional.empty();
        }
    }

    /** 異動日が空、または {@code d < inactiveFrom} のとき在籍。 */
    public static boolean activeOn(LocalDate inactiveFrom, LocalDate d) {
        return activeOn(inactiveFrom, null, d);
    }

    /** 復帰日以降は再び在籍。{@code returnedOn} が異動日より前なら無視。 */
    public static boolean activeOn(LocalDate inactiveFrom, LocalDate returnedOn, LocalDate d) {
        if (inactiveFrom == null || d == null) {
            return true;
        }
        LocalDate returned = effectiveReturnedOn(inactiveFrom, returnedOn);
        if (returned != null && !d.isBefore(returned)) {
            return true;
        }
        return d.isBefore(inactiveFrom);
    }

    /** 異動日が空、または異動日が当該月1日より後ならその月のグリッドに出す。 */
    public static boolean visibleInMonth(LocalDate inactiveFrom, int year, int month) {
        return visibleInMonth(inactiveFrom, null, year, month);
    }

    /** 当該月に在籍日が1日でもあれば出す。 */
    public static boolean visibleInMonth(
            LocalDate inactiveFrom, LocalDate returnedOn, int year, int month) {
        if (inactiveFrom == null) {
            return true;
        }
        if (inactiveFrom.isAfter(LocalDate.of(year, month, 1))) {
            return true;
        }
        LocalDate returned = effectiveReturnedOn(inactiveFrom, returnedOn);
        if (returned != null) {
            LocalDate last = YearMonth.of(year, month).atEndOfMonth();
            return !returned.isAfter(last);
        }
        return false;
    }

    private static LocalDate effectiveReturnedOn(LocalDate inactiveFrom, LocalDate returnedOn) {
        if (inactiveFrom == null || returnedOn == null) {
            return null;
        }
        if (returnedOn.isBefore(inactiveFrom)) {
            return null;
        }
        return returnedOn;
    }
}
