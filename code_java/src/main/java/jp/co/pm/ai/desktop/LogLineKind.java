package jp.co.pm.ai.desktop;

import java.util.Locale;

/** stdout/stderr one-line classification for run-tab log highlighting. */
enum LogLineKind {
    NORMAL,
    WARN,
    ERROR;

    static LogLineKind classify(String line) {
        if (line == null || line.isEmpty()) {
            return NORMAL;
        }
        String s = line.toLowerCase(Locale.ROOT);
        if (matchesError(s)) {
            return ERROR;
        }
        if (matchesWarn(s)) {
            return WARN;
        }
        return NORMAL;
    }

    private static boolean matchesError(String s) {
        if (s.contains("[error]")
                || s.contains("traceback")
                || s.contains("planningvalidationerror")
                || s.contains("失敗")
                || s.contains("エラー")) {
            return true;
        }
        if (s.contains("exception")) {
            return true;
        }
        if (s.contains("fatal")) {
            return true;
        }
        return s.contains("error:") || s.contains("error :");
    }

    private static boolean matchesWarn(String s) {
        return s.contains("[warn")
                || s.contains("warning")
                || s.contains("warn:")
                || s.contains("警告")
                || s.contains("deprecated")
                || s.contains("userwarning");
    }
}
