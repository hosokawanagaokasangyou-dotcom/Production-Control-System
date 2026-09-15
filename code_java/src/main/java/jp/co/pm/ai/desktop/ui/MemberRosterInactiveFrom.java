package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.time.format.DateTimeParseException;
import java.util.Optional;

/** 名簿の {@code inactive_from}（職場異動日）判定。Python 側と同規則。 */
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
        if (inactiveFrom == null || d == null) {
            return true;
        }
        return d.isBefore(inactiveFrom);
    }

    /** 異動日が空、または異動日が当該月1日より後ならその月のグリッドに出す。 */
    public static boolean visibleInMonth(LocalDate inactiveFrom, int year, int month) {
        if (inactiveFrom == null) {
            return true;
        }
        return inactiveFrom.isAfter(LocalDate.of(year, month, 1));
    }
}
