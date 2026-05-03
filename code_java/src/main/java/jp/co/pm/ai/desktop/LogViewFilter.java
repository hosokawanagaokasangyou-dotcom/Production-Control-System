package jp.co.pm.ai.desktop;

import java.util.function.Predicate;

/** Run-tab log list filter mode (applied to each line via {@link LogLineKind}). */
enum LogViewFilter {
    ALL("\u3059\u3079\u3066", s -> true),
    ERRORS_ONLY(
            "\u30a8\u30e9\u30fc\u306e\u307f", s -> LogLineKind.classify(s) == LogLineKind.ERROR),
    WARNS_ONLY(
            "\u8b66\u544a\u306e\u307f", s -> LogLineKind.classify(s) == LogLineKind.WARN),
    NORMAL_ONLY(
            "\u901a\u5e38\u306e\u307f", s -> LogLineKind.classify(s) == LogLineKind.NORMAL);

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
}
