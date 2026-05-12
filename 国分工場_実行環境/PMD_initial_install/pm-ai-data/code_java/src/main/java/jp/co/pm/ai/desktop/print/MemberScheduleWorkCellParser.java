package jp.co.pm.ai.desktop.print;

import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Parses member_schedule grid cells ({@code [req] process+machine...}).
 */
public final class MemberScheduleWorkCellParser {

    private static final Pattern TASK_LINE =
            Pattern.compile("^\\[([^\\]]+)\\]\\s*([^+]+)\\+(.+)$");

    private MemberScheduleWorkCellParser() {}

    public record ParsedWorkCell(String requestNo, String processName, String machineName, String rawLabel) {}

    public static ParsedWorkCell parse(String cell) {
        String raw = cell != null ? cell.trim() : "";
        if (raw.isEmpty()) {
            return new ParsedWorkCell("", "", "", "");
        }
        Matcher m = TASK_LINE.matcher(raw);
        if (!m.matches()) {
            return new ParsedWorkCell("", raw, "", raw);
        }
        String req = Objects.requireNonNullElse(m.group(1), "").trim();
        String proc = Objects.requireNonNullElse(m.group(2), "").trim();
        String mach = Objects.requireNonNullElse(m.group(3), "").trim();
        mach = stripTrailingParenRole(mach).trim();
        return new ParsedWorkCell(req, proc, mach, raw);
    }

    /** Removes a trailing parenthetical suffix for matching dispatch-table machine names. */
    static String stripTrailingParenRole(String machineWithOptionalSuffix) {
        String s = machineWithOptionalSuffix != null ? machineWithOptionalSuffix.trim() : "";
        if (s.endsWith(")") && s.contains("(")) {
            int open = s.lastIndexOf('(');
            if (open > 0 && open < s.length() - 1) {
                return s.substring(0, open).trim();
            }
        }
        return s;
    }
}
