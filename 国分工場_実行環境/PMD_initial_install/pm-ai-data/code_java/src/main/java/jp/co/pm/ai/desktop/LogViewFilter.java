package jp.co.pm.ai.desktop;

import java.util.function.Predicate;

/** Run-tab log list filter mode (applied to each line via {@link LogLineKind}). */
enum LogViewFilter {
    ALL("すべて", s -> true),
    ERRORS_ONLY(
            "エラーのみ", s -> LogLineKind.classify(s) == LogLineKind.ERROR),
    WARNS_ONLY(
            "警告のみ", s -> LogLineKind.classify(s) == LogLineKind.WARN),
    NORMAL_ONLY(
            "通常のみ", s -> LogLineKind.classify(s) == LogLineKind.NORMAL);

    private final String label;
    private final Predicate<String> match;

    LogViewFilter(String label, Predicate<String> match) {
        this.label = label;
        this.match = match;
    }

    String getLabel() {
        return label;
    }

    boolean test(String line) {
        return match.test(line);
    }

    /** Restores filter from {@link #name()} stored in session; unknown or blank yields ALL. */
    static LogViewFilter fromStoredName(String name) {
        if (name == null || name.isBlank()) {
            return ALL;
        }
        try {
            return LogViewFilter.valueOf(name.trim());
        } catch (IllegalArgumentException e) {
            return ALL;
        }
    }
}
