package jp.co.pm.ai.planning.stage2.source;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

import jp.co.pm.ai.desktop.dispatch.RawInputMorningDispatchRateAnalyzer;

/**
 * 加工計画の取得時刻から {@code PM_AI_STAGE2_SKIP_TODAY_DISPATCH} を自動判定する。
 *
 * <p>抽出日が基準暦日かつ時刻が始業開始より前 → 当日配台（skip_today OFF）。それ以外 → ON。
 */
public final class Stage2SkipTodayDispatchPolicy {

    /** Python {@code DEFAULT_START_TIME} と同一。 */
    public static final LocalTime DEFAULT_SHIFT_START = RawInputMorningDispatchRateAnalyzer.MORNING_WINDOW_START;

    private Stage2SkipTodayDispatchPolicy() {}

    /**
     * @param planExtractionTime 加工計画の取得時刻（日報は使わない）
     * @param referenceDate 判定基準の暦日（通常は {@link LocalDate#now()}）
     * @param shiftStart 始業開始（master A15 相当。未設定時は {@link #DEFAULT_SHIFT_START}）
     * @return {@code true} = 当日は配台しない（skip_today ON）
     */
    public static boolean shouldSkipTodayDispatch(
            LocalDateTime planExtractionTime, LocalDate referenceDate, LocalTime shiftStart) {
        if (planExtractionTime == null) {
            return true;
        }
        LocalDate ref = referenceDate != null ? referenceDate : LocalDate.now();
        LocalTime start = shiftStart != null ? shiftStart : DEFAULT_SHIFT_START;
        if (!planExtractionTime.toLocalDate().equals(ref)) {
            return true;
        }
        return !planExtractionTime.toLocalTime().isBefore(start);
    }

    public static boolean shouldSkipTodayDispatch(LocalDateTime planExtractionTime) {
        return shouldSkipTodayDispatch(planExtractionTime, LocalDate.now(), DEFAULT_SHIFT_START);
    }
}
