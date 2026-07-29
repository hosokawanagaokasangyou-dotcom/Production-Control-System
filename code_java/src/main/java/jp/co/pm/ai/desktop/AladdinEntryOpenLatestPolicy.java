package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;

/** 「最新を開く」ボタンの有効期限（生成からの経過時間）判定。 */
public final class AladdinEntryOpenLatestPolicy {

    /** 生成から「最新を開く」を有効にする最大経過時間。 */
    public static final Duration MAX_AGE = Duration.ofMinutes(15);

    static final String BADGE_NOT_GENERATED = "未生成（先にExcel出力を実行してください）";

    static final String BADGE_EXPIRED = "生成から15分経過（世代を開く…をご利用ください）";

    public record State(
            boolean openAllowed, String badgeText, boolean highlightGenerationsButton) {}

    private AladdinEntryOpenLatestPolicy() {}

    /** 残り秒数のカウントダウン表示（例: {@code あと 842秒}）。 */
    public static String formatCountdownBadge(long remainingSeconds) {
        return "あと " + Math.max(0, remainingSeconds) + "秒";
    }

    /**
     * 最新固定 xlsx の更新日時を生成時刻とみなし、開けるかどうかを返す。
     *
     * @param latestXlsx {@link jp.co.pm.ai.desktop.config.AppPaths#aladdinEntryDispatchPlanLocalXlsxPath}
     * @param now 判定基準時刻（テスト用に注入）
     */
    public static State resolve(Path latestXlsx, Instant now) throws IOException {
        if (latestXlsx == null || !Files.isRegularFile(latestXlsx)) {
            return new State(false, BADGE_NOT_GENERATED, false);
        }
        long generatedAtMillis = Files.getLastModifiedTime(latestXlsx).toMillis();
        long ageMs = now.toEpochMilli() - generatedAtMillis;
        long maxMs = MAX_AGE.toMillis();
        if (ageMs >= maxMs) {
            return new State(false, BADGE_EXPIRED, true);
        }
        long remainingSeconds = Math.max(0, (maxMs - ageMs) / 1000);
        return new State(true, formatCountdownBadge(remainingSeconds), false);
    }
}
