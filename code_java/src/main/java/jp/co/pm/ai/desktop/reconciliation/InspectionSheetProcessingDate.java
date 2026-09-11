package jp.co.pm.ai.desktop.reconciliation;

import java.time.LocalDate;
import java.util.Calendar;
import java.util.Locale;
import java.util.Optional;
import java.util.TimeZone;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.DateUtil;

/** 検査表ヘッダの加工日パース。 */
public final class InspectionSheetProcessingDate {

    private static final Pattern DATE =
            Pattern.compile("(\\d{4})[/.-](\\d{1,2})[/.-](\\d{1,2})");

    private InspectionSheetProcessingDate() {}

    public static boolean isProcessingDateLabel(String text) {
        if (text == null || text.isBlank()) {
            return false;
        }
        return InspectionSheetIraiNo.toHalfWidth(text).contains("加工日");
    }

    public static Optional<LocalDate> parse(String text) {
        if (text == null || text.isBlank()) {
            return Optional.empty();
        }
        Matcher m = DATE.matcher(InspectionSheetIraiNo.toHalfWidth(text));
        if (!m.find()) {
            return Optional.empty();
        }
        try {
            return Optional.of(
                    LocalDate.of(
                            Integer.parseInt(m.group(1)),
                            Integer.parseInt(m.group(2)),
                            Integer.parseInt(m.group(3))));
        } catch (RuntimeException ex) {
            return Optional.empty();
        }
    }

    public static Optional<LocalDate> fromExcelSerial(double serial) {
        if (!DateUtil.isValidExcelDate(serial) || serial < 20000) {
            return Optional.empty();
        }
        Calendar cal = DateUtil.getJavaCalendar(serial, false, TimeZone.getTimeZone("UTC"));
        if (cal == null) {
            return Optional.empty();
        }
        return Optional.of(
                LocalDate.of(
                        cal.get(Calendar.YEAR),
                        cal.get(Calendar.MONTH) + 1,
                        cal.get(Calendar.DAY_OF_MONTH)));
    }

    public static String yearMonth(LocalDate date) {
        if (date == null) {
            return "";
        }
        return String.format(Locale.ROOT, "%04d-%02d", date.getYear(), date.getMonthValue());
    }
}
