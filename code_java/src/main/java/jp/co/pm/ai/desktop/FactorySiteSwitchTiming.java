package jp.co.pm.ai.desktop;

import java.util.List;

/** 工場切り替えの工程時間を、ログ出力用の固定形式へ変換する。 */
final class FactorySiteSwitchTiming {

    private static final long NANOS_PER_MILLISECOND = 1_000_000L;

    private static final List<String> CORE_PHASE_NAMES =
            List.of(
                    "connect",
                    "save-old-workspace",
                    "load-new-workspace",
                    "restore-env-session",
                    "refresh-request-form",
                    "refresh-pipeline",
                    "refresh-remote-toolbar",
                    "stabilize-env",
                    "match-env",
                    "finish");

    private static final List<String> POST_PHASE_NAMES =
            List.of(
                    "attendance-company",
                    "attendance-member",
                    "attendance-machine",
                    "attendance-master",
                    "background-load");

    private final long startedAtNanos;

    FactorySiteSwitchTiming() {
        this(System.nanoTime());
    }

    FactorySiteSwitchTiming(long startedAtNanos) {
        this.startedAtNanos = startedAtNanos;
    }

    static List<String> corePhaseNames() {
        return CORE_PHASE_NAMES;
    }

    static List<String> postPhaseNames() {
        return POST_PHASE_NAMES;
    }

    static String corePhaseName(int step) {
        return phaseName(CORE_PHASE_NAMES, step);
    }

    static String postPhaseName(int unit) {
        return phaseName(POST_PHASE_NAMES, unit);
    }

    String phaseLine(String phase, long workStartedAtNanos, long workEndedAtNanos) {
        if (!isValidPhase(phase)
                || workStartedAtNanos < startedAtNanos
                || workEndedAtNanos < workStartedAtNanos) {
            return "";
        }
        long workNanos = workEndedAtNanos - workStartedAtNanos;
        long elapsedNanos = workEndedAtNanos - startedAtNanos;
        if (workNanos < 0L || elapsedNanos < 0L) {
            return "";
        }
        return "[factory-timing] phase="
                + phase
                + " workMs="
                + toMillis(workNanos)
                + " elapsedMs="
                + toMillis(elapsedNanos);
    }

    String totalLine(long endedAtNanos) {
        if (endedAtNanos < startedAtNanos) {
            return "";
        }
        long elapsedNanos = endedAtNanos - startedAtNanos;
        if (elapsedNanos < 0L) {
            return "";
        }
        return "[factory-timing] totalMs=" + toMillis(elapsedNanos);
    }

    private static boolean isValidPhase(String phase) {
        if (phase == null || phase.isBlank()) {
            return false;
        }
        return phase.chars().noneMatch(Character::isWhitespace);
    }

    private static String phaseName(List<String> phaseNames, int oneBasedIndex) {
        int index = oneBasedIndex - 1;
        if (index < 0 || index >= phaseNames.size()) {
            return "";
        }
        return phaseNames.get(index);
    }

    private static long toMillis(long nanos) {
        return nanos / NANOS_PER_MILLISECOND;
    }
}
