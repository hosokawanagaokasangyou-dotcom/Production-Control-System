package jp.co.pm.ai.desktop.ui;

import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** 段階1／2 実行中の子プロセスログ行からモーダル進捗の詳細文言を抽出する。 */
public final class StageRunLogProgressParser {

    private static final String PREFIX_CHILD = "[child] ";
    private static final Pattern LOG_TIMESTAMP =
            Pattern.compile(
                    "^\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2},\\d{3} (?:INFO|WARNING|ERROR|DEBUG) ");
    private static final Pattern PM_AI_PROGRESS = Pattern.compile("PM_AI_PROGRESS\\s+(\\d+)");

    private StageRunLogProgressParser() {}

    public static String stripChildPrefix(String line) {
        if (line == null) {
            return "";
        }
        String t = line.strip();
        if (t.startsWith(PREFIX_CHILD)) {
            return t.substring(PREFIX_CHILD.length()).strip();
        }
        return t;
    }

    public static Optional<String> extractDetail(String line) {
        String t = stripChildPrefix(line);
        if (t.isEmpty()) {
            return Optional.empty();
        }
        if (t.startsWith("--- start:")
                || t.startsWith("[end]")
                || t.startsWith("[busy]")
                || t.startsWith("[interrupt]")
                || t.startsWith("[error]")
                || t.startsWith("[run]")
                || t.startsWith("[policy]")
                || t.startsWith("[cleanup]")
                || t.startsWith("[stage1]")
                || t.startsWith("[stage2]")) {
            return Optional.empty();
        }
        Matcher progress = PM_AI_PROGRESS.matcher(t);
        if (progress.find()) {
            return Optional.of("進捗 " + progress.group(1) + "%");
        }
        String cleaned = LOG_TIMESTAMP.matcher(t).replaceFirst("").strip();
        if (cleaned.isEmpty()) {
            return Optional.empty();
        }
        if (cleaned.length() > 120) {
            cleaned = cleaned.substring(0, 117) + "…";
        }
        return Optional.of(cleaned);
    }
}
