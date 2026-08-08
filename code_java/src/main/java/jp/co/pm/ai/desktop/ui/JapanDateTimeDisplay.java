package jp.co.pm.ai.desktop.ui;

import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Locale;

/** ISO 8601 保存日時を日本国内向けに表示する。 */
public final class JapanDateTimeDisplay {

    public static final ZoneId JST = ZoneId.of("Asia/Tokyo");

    private static final DateTimeFormatter SAVED_AT =
            DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss", Locale.JAPAN);

    private JapanDateTimeDisplay() {}

    /**
     * UTC 等のタイムゾーン付き ISO を日本時間に変換して表示する。
     * タイムゾーン無しの旧形式は UTC として解釈する。
     */
    public static String formatSavedAtForDisplay(String raw) {
        if (raw == null || raw.isBlank()) {
            return "";
        }
        String trimmed = raw.strip();
        try {
            return SAVED_AT.format(Instant.parse(trimmed).atZone(JST));
        } catch (DateTimeParseException ignored) {
            // fall through
        }
        try {
            LocalDateTime local = LocalDateTime.parse(trimmed, DateTimeFormatter.ISO_LOCAL_DATE_TIME);
            return SAVED_AT.format(local.atZone(ZoneId.of("UTC")).withZoneSameInstant(JST));
        } catch (DateTimeParseException ignored) {
            return trimmed;
        }
    }
}
