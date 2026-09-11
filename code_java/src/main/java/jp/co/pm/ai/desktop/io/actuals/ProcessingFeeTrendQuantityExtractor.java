package jp.co.pm.ai.desktop.io.actuals;

import java.text.Normalizer;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.dispatch.ResultDispatchPlanningStageSupport;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;
import jp.co.pm.ai.desktop.io.actuals.ProcessingFeeTrendAggregator.QuantityLine;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.ActualSource;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.Filter;
import jp.co.pm.ai.desktop.io.actuals.ProcessingTrendAggregator.PlanSource;

/**
 * 実績・予定スナップショットから加工賃集計用の数量行（日付・依頼No・m）を抽出する。
 *
 * <p>日報ソースの実績 m は「実製品出来高」を正とする（無ければ実加工量→実加工数へフォールバック）。
 */
public final class ProcessingFeeTrendQuantityExtractor {

    private static final double EPS = 1e-9;
    private static final String COL_MACHINE = "機械名";
    private static final String COL_PROCESS = "工程名";
    private static final String COL_ACTUAL_QTY_DETAIL = "実加工数";
    private static final String COL_ACTUAL_QTY_DAILY = "実加工量";
    /** 加工賃用: 加工日報の製品出来高（加工量トレンドの実加工量とは別）。 */
    private static final String COL_PRODUCT_OUTPUT_DAILY = "実製品出来高";
    private static final String COL_ACTUAL_DATE_DAILY = "加工日付";
    private static final String COL_ACTUAL_DATE = "加工日";
    private static final String COL_ACTUAL_START_DT = "加工開始日時";
    /** 加工日報の終了時刻（加工日付と合成して最終工程判定に使う）。 */
    private static final String COL_END_TIME = "終了時間";
    private static final String COL_WAREHOUSE = "倉庫";
    private static final String COL_TASK_ID = "依頼NO";
    private static final String COL_CONVERSION_QTY = "換算数量";
    private static final String COL_UNPROCESSED = "未加工";
    private static final String COL_COMPLETION_FLAG = "加工完了区分";
    private static final String TOTAL_ROW_PREFIX = "[合計]";
    private static final Pattern DATE_HEADER = Pattern.compile("\\d{4}/\\d{2}/\\d{2}");
    private static final Pattern WHITESPACE = Pattern.compile("\\s+");
    private static final Pattern TIME_HM =
            Pattern.compile("^(\\d{1,2}):(\\d{2})(?::(\\d{2}))?$");
    private static final Pattern TIME_COMPACT = Pattern.compile("^(\\d{1,2})(\\d{2})$");

    private ProcessingFeeTrendQuantityExtractor() {}

    public static List<QuantityLine> extractActual(
            ActualsSnapshot dailyReport,
            ActualsSnapshot detail,
            Filter filter) {
        return extractActual(dailyReport, detail, filter, true);
    }

    /**
     * @param boundToPeriod true なら Filter の from〜to 内のみ。false なら日付がある行は期間外も採用
     *     （依頼NO別の受注残＝AO 正本用。日次棒は呼び出し側で期間に絞る）
     */
    public static List<QuantityLine> extractActual(
            ActualsSnapshot dailyReport,
            ActualsSnapshot detail,
            Filter filter,
            boolean boundToPeriod) {
        boolean daily = filter.actualSource() == ActualSource.DAILY_REPORT;
        ActualsSnapshot src = daily ? dailyReport : detail;
        return extractActualFrom(src, filter, daily, boundToPeriod);
    }

    public static List<QuantityLine> extractPlan(
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            Filter filter,
            LocalDate today) {
        if (filter.planSource() == PlanSource.DISPATCH) {
            return extractDispatch(dispatch, filter, today != null ? today : LocalDate.now());
        }
        return extractAladdin(aladdin, filter, today != null ? today : LocalDate.now());
    }

    private static List<QuantityLine> extractActualFrom(
            ActualsSnapshot actuals, Filter f, boolean preferDaily, boolean boundToPeriod) {
        List<QuantityLine> out = new ArrayList<>();
        if (actuals == null || actuals.headers() == null || actuals.rows() == null) {
            return out;
        }
        List<String> headers = actuals.headers();
        int iQty =
                preferDaily
                        ? firstCol(
                                headers,
                                COL_PRODUCT_OUTPUT_DAILY,
                                COL_ACTUAL_QTY_DAILY,
                                COL_ACTUAL_QTY_DETAIL)
                        : firstCol(headers, COL_ACTUAL_QTY_DETAIL, COL_ACTUAL_QTY_DAILY, COL_PRODUCT_OUTPUT_DAILY);
        if (iQty < 0) {
            return out;
        }
        int iMachine = colIdx(headers, COL_MACHINE);
        int iProcess = colIdx(headers, COL_PROCESS);
        int iTask = colIdx(headers, COL_TASK_ID);
        int iStartDt = colIdx(headers, COL_ACTUAL_START_DT);
        int iDailyDate = colIdx(headers, COL_ACTUAL_DATE_DAILY);
        int iKakouDate = colIdx(headers, COL_ACTUAL_DATE);
        int iEndTime = colIdx(headers, COL_END_TIME);
        String mk = normKey(f.machine());
        String pk = normKey(f.process());
        LocalDate from = f.from();
        LocalDate to = f.to();
        for (List<String> row : actuals.rows()) {
            if (row == null) {
                continue;
            }
            if (!matches(mk, cellAt(row, iMachine)) || !matches(pk, cellAt(row, iProcess))) {
                continue;
            }
            LocalDate d = rowActualDate(row, iStartDt, iDailyDate, iKakouDate);
            if (d == null) {
                continue;
            }
            if (boundToPeriod && (d.isBefore(from) || d.isAfter(to))) {
                continue;
            }
            double qty = parseDouble(cellAt(row, iQty));
            if (Math.abs(qty) <= EPS) {
                continue;
            }
            String task =
                    iTask >= 0 ? ProcessingTrendAggregator.normKey(cellAt(row, iTask)) : "";
            String process = iProcess >= 0 ? cellAt(row, iProcess).strip() : "";
            LocalDateTime finishedAt =
                    iEndTime >= 0 ? composeFinishedAt(d, cellAt(row, iEndTime)) : null;
            out.add(new QuantityLine(d, task, qty, process, finishedAt));
        }
        return out;
    }

    private static List<QuantityLine> extractAladdin(
            AladdinSnapshot aladdin, Filter f, LocalDate today) {
        List<QuantityLine> out = new ArrayList<>();
        if (aladdin == null || aladdin.headers() == null || aladdin.rows() == null) {
            return out;
        }
        List<String> headers = aladdin.headers();
        int iMachine = colIdx(headers, COL_MACHINE);
        int iProcess = colIdx(headers, COL_PROCESS);
        int iWarehouse = colIdx(headers, COL_WAREHOUSE);
        int iTask = colIdx(headers, COL_TASK_ID);
        int iConv = colIdx(headers, COL_CONVERSION_QTY);
        int iDone = colIdx(headers, COL_ACTUAL_QTY_DETAIL);
        int iUnprocessed = colIdx(headers, COL_UNPROCESSED);
        int iCompletion = colIdx(headers, COL_COMPLETION_FLAG);

        TreeMap<LocalDate, Integer> allDateCols = new TreeMap<>();
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (h != null && DATE_HEADER.matcher(h.strip()).matches()) {
                LocalDate d = parseDate(h);
                if (d != null) {
                    allDateCols.putIfAbsent(d, i);
                }
            }
        }
        List<LocalDate> futureDates = new ArrayList<>(allDateCols.tailMap(today, true).keySet());
        String mk = normKey(f.machine());
        String pk = normKey(f.process());
        LocalDate from = f.from();
        LocalDate to = f.to();
        Map<LocalDate, Double> rowValues = new LinkedHashMap<>();
        for (List<String> row : aladdin.rows()) {
            if (row == null || isAladdinTotalRow(row, iWarehouse, iMachine, iTask)) {
                continue;
            }
            if (!matches(mk, cellAt(row, iMachine)) || !matches(pk, cellAt(row, iProcess))) {
                continue;
            }
            String task = iTask >= 0 ? cellAt(row, iTask).strip() : "";
            String process = iProcess >= 0 ? cellAt(row, iProcess).strip() : "";
            rowValues.clear();
            for (Map.Entry<LocalDate, Integer> e : allDateCols.entrySet()) {
                double v = parseDouble(cellAt(row, e.getValue()));
                if (Math.abs(v) > EPS) {
                    rowValues.put(e.getKey(), v);
                }
            }
            if (rowValues.isEmpty()) {
                continue;
            }
            capRemainingPlan(row, rowValues, futureDates, iConv, iDone, iUnprocessed, iCompletion);
            for (Map.Entry<LocalDate, Double> e : rowValues.entrySet()) {
                LocalDate day = e.getKey();
                if (day == null || day.isBefore(from) || day.isAfter(to)) {
                    continue;
                }
                if (Math.abs(e.getValue()) <= EPS) {
                    continue;
                }
                out.add(new QuantityLine(day, task, e.getValue(), process));
            }
        }
        return out;
    }

    private static List<QuantityLine> extractDispatch(
            DispatchSnapshot dispatch, Filter f, LocalDate today) {
        List<QuantityLine> out = new ArrayList<>();
        if (dispatch == null || dispatch.headers() == null || dispatch.rows() == null) {
            return out;
        }
        List<String> headers = dispatch.headers();
        int iMachine = colIdx(headers, ResultDispatchSchema.COL_MACHINE);
        int iProcess = colIdx(headers, ResultDispatchSchema.COL_PROCESS);
        int iDate = colIdx(headers, ResultDispatchSchema.COL_DISPATCH_DATE);
        int iQty = colIdx(headers, ResultDispatchSchema.COL_DISPATCH_QTY);
        int iTask = colIdx(headers, "依頼NO");
        if (iDate < 0 || iQty < 0) {
            return out;
        }
        List<List<String>> rows =
                ResultDispatchPlanningStageSupport.normalizeLegacyDispatchRowsForAggregation(
                        headers, dispatch.rows());
        String mk = normKey(f.machine());
        String pk = normKey(f.process());
        LocalDate from = f.from();
        LocalDate to = f.to();
        for (List<String> row : rows) {
            if (row == null) {
                continue;
            }
            if (!matches(mk, cellAt(row, iMachine)) || !matches(pk, cellAt(row, iProcess))) {
                continue;
            }
            LocalDate d = parseDate(cellAt(row, iDate));
            if (d == null || d.isBefore(from) || d.isAfter(to)) {
                continue;
            }
            // 当日以前の配台数量は実績と二重になるため予定に載せない（見込＝実績+翌日以降予定）
            if (today != null && !d.isAfter(today)) {
                continue;
            }
            double v = parseDouble(cellAt(row, iQty));
            if (Math.abs(v) <= EPS) {
                continue;
            }
            String task = iTask >= 0 ? cellAt(row, iTask).strip() : "";
            String process = iProcess >= 0 ? cellAt(row, iProcess).strip() : "";
            out.add(new QuantityLine(d, task, v, process));
        }
        return out;
    }

    /** m トレンドと同趣旨の当日以降キャップ（簡易移植）。 */
    private static void capRemainingPlan(
            List<String> row,
            Map<LocalDate, Double> rowValues,
            List<LocalDate> futureDates,
            int iConv,
            int iDone,
            int iUnprocessed,
            int iCompletion) {
        if (iCompletion >= 0) {
            String flag = cellAt(row, iCompletion);
            if (flag.contains("完了") && !flag.contains("未完")) {
                for (LocalDate d : futureDates) {
                    rowValues.put(d, 0.0);
                }
                return;
            }
        }
        double unprocessed = iUnprocessed >= 0 ? parseDouble(cellAt(row, iUnprocessed)) : Double.NaN;
        if (iConv >= 0 && iDone >= 0) {
            double conv = parseDouble(cellAt(row, iConv));
            double done = parseDouble(cellAt(row, iDone));
            if (conv > EPS && Math.abs(done) <= EPS && Math.abs(unprocessed) <= EPS) {
                unprocessed = conv;
            }
        }
        if (Double.isNaN(unprocessed) || unprocessed < 0) {
            return;
        }
        double futureSum = 0;
        for (LocalDate d : futureDates) {
            futureSum += rowValues.getOrDefault(d, 0.0);
        }
        if (futureSum <= unprocessed + EPS) {
            return;
        }
        double excess = futureSum - unprocessed;
        for (int i = futureDates.size() - 1; i >= 0 && excess > EPS; i--) {
            LocalDate d = futureDates.get(i);
            double v = rowValues.getOrDefault(d, 0.0);
            if (v <= EPS) {
                continue;
            }
            double cut = Math.min(v, excess);
            rowValues.put(d, v - cut);
            excess -= cut;
        }
    }

    private static boolean isAladdinTotalRow(
            List<String> row, int iWarehouse, int iMachine, int iTask) {
        if (iWarehouse >= 0 && cellAt(row, iWarehouse).strip().startsWith(TOTAL_ROW_PREFIX)) {
            return true;
        }
        return iMachine >= 0
                && iTask >= 0
                && cellAt(row, iMachine).isBlank()
                && cellAt(row, iTask).isBlank();
    }

    private static LocalDate rowActualDate(List<String> row, int iStartDt, int iDailyDate, int iKakouDate) {
        if (iStartDt >= 0) {
            LocalDate d = parseDate(cellAt(row, iStartDt));
            if (d != null) {
                return d;
            }
        }
        if (iDailyDate >= 0) {
            LocalDate d = parseDate(cellAt(row, iDailyDate));
            if (d != null) {
                return d;
            }
        }
        return iKakouDate >= 0 ? parseDate(cellAt(row, iKakouDate)) : null;
    }

    /** 加工日付＋終了時間 → 最終工程比較用の日時。解釈できないときは null。 */
    static LocalDateTime composeFinishedAt(LocalDate day, String endTimeRaw) {
        if (day == null) {
            return null;
        }
        LocalTime t = parseEndTime(endTimeRaw);
        if (t == null) {
            LocalDateTime asDt = parseDateTime(endTimeRaw);
            return asDt;
        }
        return LocalDateTime.of(day, t);
    }

    static LocalTime parseEndTime(String raw) {
        if (raw == null) {
            return null;
        }
        String s = raw.strip();
        if (s.isEmpty()) {
            return null;
        }
        // 日付付きなら時刻部だけ取る
        if (s.length() > 10 && (s.charAt(10) == ' ' || s.charAt(10) == 'T')) {
            s = s.substring(11).strip();
        }
        Matcher hm = TIME_HM.matcher(s);
        if (hm.matches()) {
            int h = Integer.parseInt(hm.group(1));
            int m = Integer.parseInt(hm.group(2));
            int sec = hm.group(3) != null ? Integer.parseInt(hm.group(3)) : 0;
            if (h >= 0 && h <= 23 && m >= 0 && m <= 59 && sec >= 0 && sec <= 59) {
                return LocalTime.of(h, m, sec);
            }
            return null;
        }
        Matcher compact = TIME_COMPACT.matcher(s);
        if (compact.matches()) {
            int h = Integer.parseInt(compact.group(1));
            int m = Integer.parseInt(compact.group(2));
            if (h >= 0 && h <= 23 && m >= 0 && m <= 59) {
                return LocalTime.of(h, m);
            }
            return null;
        }
        // 整数時のみ（例: 15）
        try {
            double n = Double.parseDouble(s.replace(",", ""));
            int h = (int) n;
            if (h == n && h >= 0 && h <= 23) {
                return LocalTime.of(h, 0);
            }
        } catch (NumberFormatException ignored) {
            // fall through
        }
        return null;
    }

    private static LocalDateTime parseDateTime(String raw) {
        if (raw == null) {
            return null;
        }
        String s = raw.strip();
        if (s.isEmpty() || s.length() <= 10) {
            return null;
        }
        String[] patterns = {
            "yyyy/M/d H:mm:ss",
            "yyyy/M/d H:mm",
            "yyyy-M-d H:mm:ss",
            "yyyy-M-d H:mm",
            "yyyy-MM-dd'T'HH:mm:ss"
        };
        for (String p : patterns) {
            try {
                return LocalDateTime.parse(s, DateTimeFormatter.ofPattern(p, Locale.JAPAN));
            } catch (DateTimeParseException ignored) {
                // try next
            }
        }
        return null;
    }

    private static int firstCol(List<String> headers, String... names) {
        if (names == null) {
            return -1;
        }
        for (String name : names) {
            int i = colIdx(headers, name);
            if (i >= 0) {
                return i;
            }
        }
        return -1;
    }

    private static int colIdx(List<String> headers, String name) {
        if (headers == null || name == null) {
            return -1;
        }
        String want = normHeader(name);
        for (int i = 0; i < headers.size(); i++) {
            if (want.equals(normHeader(headers.get(i)))) {
                return i;
            }
        }
        return -1;
    }

    private static String normHeader(String h) {
        if (h == null) {
            return "";
        }
        return Normalizer.normalize(h, Normalizer.Form.NFKC).strip();
    }

    private static String cellAt(List<String> row, int i) {
        if (i < 0 || row == null || i >= row.size() || row.get(i) == null) {
            return "";
        }
        return row.get(i);
    }

    private static String normKey(String s) {
        if (s == null || s.isBlank() || "（すべて）".equals(s.strip())) {
            return "";
        }
        return Normalizer.normalize(s, Normalizer.Form.NFKC).strip();
    }

    private static boolean matches(String filterKey, String cell) {
        if (filterKey == null || filterKey.isEmpty()) {
            return true;
        }
        return filterKey.equals(normKey(cell));
    }

    private static double parseDouble(String raw) {
        if (raw == null) {
            return 0;
        }
        String s = WHITESPACE.matcher(raw.strip()).replaceAll("");
        if (s.isEmpty()) {
            return 0;
        }
        s = s.replace(",", "").replace("，", "");
        try {
            return Double.parseDouble(s);
        } catch (NumberFormatException ex) {
            return 0;
        }
    }

    private static LocalDate parseDate(String raw) {
        if (raw == null) {
            return null;
        }
        String s = raw.strip();
        if (s.isEmpty()) {
            return null;
        }
        if (s.length() > 10 && s.charAt(10) == ' ') {
            s = s.substring(0, 10);
        }
        String[] patterns = {
            "yyyy/M/d", "yyyy-M-d", "yyyy/MM/dd", "yyyy-MM-dd", "yyyy/M/d H:mm:ss", "yyyy-MM-dd'T'HH:mm:ss"
        };
        for (String p : patterns) {
            try {
                if (p.contains("H") || p.contains("'T'")) {
                    return LocalDateTime.parse(raw.strip(), DateTimeFormatter.ofPattern(p, Locale.JAPAN))
                            .toLocalDate();
                }
                return LocalDate.parse(s, DateTimeFormatter.ofPattern(p, Locale.JAPAN));
            } catch (DateTimeParseException ignored) {
                // try next
            }
        }
        return null;
    }
}
