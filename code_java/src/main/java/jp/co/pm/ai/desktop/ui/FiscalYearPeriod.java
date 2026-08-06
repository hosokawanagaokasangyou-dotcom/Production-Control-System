package jp.co.pm.ai.desktop.ui;

import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.List;

/** 会社カレンダー等で使う年度期間（開始日から翌年同日の前日まで）。 */
public final class FiscalYearPeriod {

    public static final FiscalYearPeriod DEFAULT_APRIL_MARCH = new FiscalYearPeriod(4, 1);

    private final int startMonth;
    private final int startDay;

    public FiscalYearPeriod(int startMonth, int startDay) {
        if (startMonth < 1 || startMonth > 12) {
            throw new IllegalArgumentException("startMonth: " + startMonth);
        }
        if (startDay < 1 || startDay > 31) {
            throw new IllegalArgumentException("startDay: " + startDay);
        }
        this.startMonth = startMonth;
        this.startDay = startDay;
    }

    public int startMonth() {
        return startMonth;
    }

    public int startDay() {
        return startDay;
    }

    /** 月の日数に合わせてクランプした開始日。 */
    public LocalDate startDate(int fiscalYearLabel) {
        return clampDay(fiscalYearLabel, startMonth, startDay);
    }

    public LocalDate endDateInclusive(int fiscalYearLabel) {
        LocalDate nextStart = startDate(fiscalYearLabel + 1);
        return nextStart.minusDays(1);
    }

    public static LocalDate clampDay(int year, int month, int day) {
        int max = YearMonth.of(year, month).lengthOfMonth();
        int d = Math.min(day, max);
        try {
            return LocalDate.of(year, month, d);
        } catch (DateTimeException e) {
            return LocalDate.of(year, month, max);
        }
    }

    public List<YearMonth> monthsInOrder(int fiscalYearLabel) {
        List<YearMonth> out = new ArrayList<>();
        YearMonth cur = YearMonth.from(startDate(fiscalYearLabel));
        YearMonth last = YearMonth.from(endDateInclusive(fiscalYearLabel));
        while (!cur.isAfter(last)) {
            out.add(cur);
            cur = cur.plusMonths(1);
        }
        return out;
    }

    public boolean contains(LocalDate date, int fiscalYearLabel) {
        return !date.isBefore(startDate(fiscalYearLabel))
                && !date.isAfter(endDateInclusive(fiscalYearLabel));
    }

    public static int fiscalYearLabelFor(LocalDate date, FiscalYearPeriod period) {
        LocalDate boundary = clampDay(date.getYear(), period.startMonth(), period.startDay());
        if (date.isBefore(boundary)) {
            return date.getYear() - 1;
        }
        return date.getYear();
    }

    public String rangeLabel(int fiscalYearLabel) {
        LocalDate start = startDate(fiscalYearLabel);
        LocalDate end = endDateInclusive(fiscalYearLabel);
        return fiscalYearLabel
                + "年度（"
                + start.getYear()
                + "/"
                + start.getMonthValue()
                + "/"
                + start.getDayOfMonth()
                + "〜"
                + end.getYear()
                + "/"
                + end.getMonthValue()
                + "/"
                + end.getDayOfMonth()
                + "）";
    }
}
