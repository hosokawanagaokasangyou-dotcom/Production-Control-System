package jp.co.pm.ai.desktop.io;

import java.time.LocalDate;
import java.util.Calendar;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.DateUtil;

/**
 * 国分アラジン／実績明細の未書式日付セル（Excel シリアル）を暦日に変換する。
 */
public final class ExcelSerialDates {

    /** 1954年付近より前の値は数量と区別できないため日付とみなさない。 */
    static final double MIN_INCLUSIVE = 20000d;

    /** 2064年付近より後は日付として扱わない。 */
    static final double MAX_INCLUSIVE = 60000d;

    private ExcelSerialDates() {}

    public static LocalDate parse(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String t = raw.strip().replace(",", "");
        if (t.isEmpty()) {
            return null;
        }
        try {
            double serial = Double.parseDouble(t);
            return fromSerial(serial);
        } catch (NumberFormatException ex) {
            return null;
        }
    }

    public static LocalDate fromSerial(double serial) {
        if (!Double.isFinite(serial) || serial < MIN_INCLUSIVE || serial > MAX_INCLUSIVE) {
            return null;
        }
        if (!DateUtil.isValidExcelDate(serial)) {
            return null;
        }
        Calendar cal = DateUtil.getJavaCalendar(serial, false, TimeZone.getTimeZone("UTC"));
        if (cal == null) {
            return null;
        }
        return LocalDate.of(
                cal.get(Calendar.YEAR), cal.get(Calendar.MONTH) + 1, cal.get(Calendar.DAY_OF_MONTH));
    }
}
