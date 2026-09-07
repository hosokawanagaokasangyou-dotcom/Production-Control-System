package jp.co.pm.ai.desktop.io.actuals;

import java.text.Collator;
import java.text.Normalizer;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.dispatch.ResultDispatchInteractiveConsolidator;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchPlanningStageSupport;
import jp.co.pm.ai.desktop.dispatch.ResultDispatchSchema;

/**
 * 加工実績（加工実績明細）と加工予定（アラジン加工計画または配台結果）を日別に集計し、
 * 実績・予定・累計・見込のトレンド系列を組み立てる。
 *
 * <p>単位はすべて m（換算数量ベース・工程延べ）。同一依頼でも工程ごとに行があるため、全工程を合算した値は
 * 「工程延べ m」であり依頼の生産量ではない。金額は Java 側にソースが無いため扱わない。
 *
 * <p>見込（projected）の定義: <b>当日より前は実績、翌日以降は予定、当日は実績と予定の大きい方</b>を採用した
 * 日別値の累計。当日は加工途中で実績が確定しないため予定側に寄せるが、終業後に実績が予定を超えている場合は
 * 実績を採用する（境界は {@code today}）。
 *
 * <p>アラジン加工計画の日付列は「計画者による現時点の残予定の配置」であり、完了した依頼は抽出から消える。
 * そのため <b>前日までの予定は構造的に欠け得る</b>（進捗率は参考値）。また日付列に既加工分が残る行があるため、
 * 当日以降の残予定は行ごとに {@code 未加工}（全数未加工ルール適用後）で上限を掛ける。
 */
public final class ProcessingTrendAggregator {

    /** 実績系列の取得元。 */
    public enum ActualSource {
        DAILY_REPORT("日報（実加工量）"),
        DETAIL("実績明細（実加工数）");

        private final String label;

        ActualSource(String label) {
            this.label = label;
        }

        public String label() {
            return label;
        }

        @Override
        public String toString() {
            return label;
        }
    }

    /** 予定系列の取得元。 */
    public enum PlanSource {
        ALADDIN("アラジン加工計画"),
        DISPATCH("配台結果 (JSON)");

        private final String label;

        PlanSource(String label) {
            this.label = label;
        }

        public String label() {
            return label;
        }

        @Override
        public String toString() {
            return label;
        }
    }

    /**
     * 集計条件。
     *
     * @param from 期間開始（含む）
     * @param to 期間終了（含む）
     * @param actualSource 実績の取得元（デフォルト: 日報）
     * @param planSource 予定の取得元
     * @param machine 機械名（{@code null} / 空 = 全機械）
     * @param process 工程名（{@code null} / 空 = 全工程）
     * @param movingAverageDays 移動平均の窓幅（日）。許可値は 7 / 14 / 30。既定 30
     */
    public record Filter(
            LocalDate from,
            LocalDate to,
            ActualSource actualSource,
            PlanSource planSource,
            String machine,
            String process,
            int movingAverageDays) {

        /** 移動平均窓幅の既定値（日）。 */
        public static final int DEFAULT_MOVING_AVERAGE_DAYS = 30;

        public Filter {
            Objects.requireNonNull(from, "from");
            Objects.requireNonNull(to, "to");
            if (to.isBefore(from)) {
                LocalDate t = from;
                from = to;
                to = t;
            }
            actualSource = actualSource != null ? actualSource : ActualSource.DAILY_REPORT;
            planSource = planSource != null ? planSource : PlanSource.ALADDIN;
            machine = blankToNull(machine);
            process = blankToNull(process);
            movingAverageDays = normalizeMovingAverageDays(movingAverageDays);
        }

        /** 従来互換コンストラクタ（移動平均は既定 30 日）。 */
        public Filter(
                LocalDate from,
                LocalDate to,
                ActualSource actualSource,
                PlanSource planSource,
                String machine,
                String process) {
            this(from, to, actualSource, planSource, machine, process, DEFAULT_MOVING_AVERAGE_DAYS);
        }

        /** 従来互換コンストラクタ（実績ソースは日報、移動平均は既定 30 日）。 */
        public Filter(LocalDate from, LocalDate to, PlanSource planSource, String machine, String process) {
            this(from, to, ActualSource.DAILY_REPORT, planSource, machine, process, DEFAULT_MOVING_AVERAGE_DAYS);
        }

        public boolean hasMachine() {
            return machine != null;
        }

        public boolean hasProcess() {
            return process != null;
        }

        static int normalizeMovingAverageDays(int days) {
            if (days == 7 || days == 14 || days == 30) {
                return days;
            }
            return DEFAULT_MOVING_AVERAGE_DAYS;
        }
    }

    /**
     * 日別 1 点。
     *
     * @param date 日付
     * @param actualM 当日実績 (m)
     * @param planM 当日予定 (m)
     * @param actualCumM 期間開始からの実績累計 (m)
     * @param planCumM 期間開始からの予定累計 (m)
     * @param projectedCumM 見込累計 (m)。当日より前は実績、当日以降は予定を積む
     * @param usesPlanForProjection 見込の当日値が予定側かどうか（{@code date >= today}）
     * @param compareActualM 比較対象実績 (m)（日報選択時は実績明細、明細選択時は日報）
     * @param compareActualCumM 比較対象実績の累計 (m)
     * @param actualMaM 当日を含む直近 N 日間の実績移動平均 (m)。N は {@link Filter#movingAverageDays()}
     */
    public record DayPoint(
            LocalDate date,
            double actualM,
            double planM,
            double actualCumM,
            double planCumM,
            double projectedCumM,
            boolean usesPlanForProjection,
            double compareActualM,
            double compareActualCumM,
            double actualMaM) {

        /** 従来互換コンストラクタ（移動平均は 0.0）。 */
        public DayPoint(
                LocalDate date,
                double actualM,
                double planM,
                double actualCumM,
                double planCumM,
                double projectedCumM,
                boolean usesPlanForProjection,
                double compareActualM,
                double compareActualCumM) {
            this(
                    date,
                    actualM,
                    planM,
                    actualCumM,
                    planCumM,
                    projectedCumM,
                    usesPlanForProjection,
                    compareActualM,
                    compareActualCumM,
                    0.0);
        }

        /** 従来互換コンストラクタ（比較実績および移動平均は 0）。 */
        public DayPoint(
                LocalDate date,
                double actualM,
                double planM,
                double actualCumM,
                double planCumM,
                double projectedCumM,
                boolean usesPlanForProjection) {
            this(date, actualM, planM, actualCumM, planCumM, projectedCumM, usesPlanForProjection, 0.0, 0.0, 0.0);
        }

        /** @deprecated 互換用。{@link #actualMaM()} を使うこと。 */
        @Deprecated
        public double actual7dMaM() {
            return actualMaM;
        }

        /** 当日差異 = 実績 − 予定。 */
        public double diffM() {
            return actualM - planM;
        }

        /** 比較対象実績との差異 = 主実績 − 比較実績（日報選択時は 日報 − 実績明細）。 */
        public double actualCompareDiffM() {
            return actualM - compareActualM;
        }
    }

    /**
     * 集計結果。
     *
     * @param days 期間内の全日（欠損日も 0 で埋める）
     * @param actualTotalM 期間実績合計
     * @param planTotalM 期間予定合計
     * @param actualToDateM 当日より前の実績合計（進捗率の分子）
     * @param planToDateM 当日より前の予定合計（進捗率の分母）
     * @param remainingPlanM 当日以降の見込合計（翌日以降は予定、当日は実績と予定の大きい方）
     * @param projectedTotalM 見込合計 = actualToDateM + remainingPlanM
     * @param today 見込境界に用いた当日
     * @param actualRowsCounted 集計に採用した実績行数（フィルタ後・期間内で非 0 の行）
     * @param planRowsCounted 集計に採用した予定行数（フィルタ後・期間内で非 0 の行）
     * @param actualMinDate 実績ソース（機械・工程フィルタ後）に含まれる最小加工日。無ければ {@code null}
     * @param actualMaxDate 実績ソース（機械・工程フィルタ後）に含まれる最大加工日。無ければ {@code null}
     * @param warnings 集計上の注意（列欠落など）。UI へそのまま表示できる日本語
     * @param compareActualTotalM 比較対象実績の期間合計 (m)
     * @param compareSourceLabel 比較対象実績のラベル（例: "実績明細" または "日報"）
     */
    public record Result(
            List<DayPoint> days,
            double actualTotalM,
            double planTotalM,
            double actualToDateM,
            double planToDateM,
            double remainingPlanM,
            double projectedTotalM,
            LocalDate today,
            int actualRowsCounted,
            int planRowsCounted,
            LocalDate actualMinDate,
            LocalDate actualMaxDate,
            List<String> warnings,
            double compareActualTotalM,
            String compareSourceLabel) {

        public Result {
            days = days != null ? List.copyOf(days) : List.of();
            warnings = warnings != null ? List.copyOf(warnings) : List.of();
            compareSourceLabel = compareSourceLabel != null ? compareSourceLabel : "";
        }

        /** 従来互換コンストラクタ（比較実績なし）。 */
        public Result(
                List<DayPoint> days,
                double actualTotalM,
                double planTotalM,
                double actualToDateM,
                double planToDateM,
                double remainingPlanM,
                double projectedTotalM,
                LocalDate today,
                int actualRowsCounted,
                int planRowsCounted,
                LocalDate actualMinDate,
                LocalDate actualMaxDate,
                List<String> warnings) {
            this(
                    days,
                    actualTotalM,
                    planTotalM,
                    actualToDateM,
                    planToDateM,
                    remainingPlanM,
                    projectedTotalM,
                    today,
                    actualRowsCounted,
                    planRowsCounted,
                    actualMinDate,
                    actualMaxDate,
                    warnings,
                    0.0,
                    "");
        }

        public static Result empty(LocalDate today) {
            return new Result(List.of(), 0, 0, 0, 0, 0, 0, today, 0, 0, null, null, List.of(), 0.0, "");
        }

        /** 比較対象実績との合計差異 = 主実績合計 − 比較実績合計。 */
        public double actualCompareDiffTotalM() {
            return actualTotalM - compareActualTotalM;
        }

        /**
         * 進捗率 (%) = 当日より前の実績 ÷ 当日より前の予定。分母 0 のとき {@code NaN}。
         * 予定が構造的に欠けると生値が 100% を超え得るが、進捗率としては 100% を上限とする。
         */
        public double progressPct() {
            if (planToDateM <= EPS) {
                return Double.NaN;
            }
            return Math.min(100.0, actualToDateM / planToDateM * 100.0);
        }

        /**
         * 進捗率の分母（前日までの予定）が期間予定合計に対して十分か。
         * アラジン予定は完了依頼が抽出から消えるため、月初などで分母が極端に小さいと参考にならない。
         */
        public boolean progressDenominatorSufficient() {
            if (planToDateM <= EPS || planTotalM <= EPS) {
                return false;
            }
            return planToDateM / planTotalM >= PROGRESS_DENOMINATOR_MIN_RATIO;
        }

        /** 見込差異 = 見込合計 − 予定合計。 */
        public double projectedDiffM() {
            return projectedTotalM - planTotalM;
        }

        /** 期間開始が実績ソースの最小加工日より前か（実績が無いのではなくソースに含まれていない）。 */
        public boolean periodStartsBeforeActualSource() {
            return actualMinDate != null && !days.isEmpty() && days.get(0).date().isBefore(actualMinDate);
        }

        public boolean isEmpty() {
            return actualRowsCounted == 0 && planRowsCounted == 0;
        }

        /** 期間末日時点の実績累計。日数0のときは 0。 */
        public double actualCumAtEnd() {
            return days.isEmpty() ? 0.0 : days.get(days.size() - 1).actualCumM();
        }

        /** 期間末日時点の予定累計。日数0のときは 0。 */
        public double planCumAtEnd() {
            return days.isEmpty() ? 0.0 : days.get(days.size() - 1).planCumM();
        }
    }

    /**
     * 月別 1 点。単位は日次と同じ工程延べ m。
     *
     * @param month 年月
     * @param fromInclusive 月内の集計開始日
     * @param toInclusive 月内の集計終了日
     * @param daysInBucket 期間に含まれる日数
     * @param calendarDaysInMonth その月の暦日数
     * @param actualM 月内実績合計
     * @param planM 月内予定合計
     * @param actualCumM 期間開始からの実績累計（月末時点）
     * @param planCumM 期間開始からの予定累計（月末時点）
     * @param projectedCumM 期間開始からの見込累計（月末時点）
     * @param projectedM 月内の見込寄与
     * @param usesPlanForProjection 当月以降かどうか
     * @param incomplete 月の一部のみ期間に含まれるか
     * @param isCurrentMonth 当月かどうか
     * @param compareActualM 比較対象実績の月内合計 (m)
     * @param compareActualCumM 比較対象実績の累計 (m)
     */
    public record MonthPoint(
            YearMonth month,
            LocalDate fromInclusive,
            LocalDate toInclusive,
            int daysInBucket,
            int calendarDaysInMonth,
            double actualM,
            double planM,
            double actualCumM,
            double planCumM,
            double projectedCumM,
            double projectedM,
            boolean usesPlanForProjection,
            boolean incomplete,
            boolean isCurrentMonth,
            double compareActualM,
            double compareActualCumM) {

        /** 従来互換コンストラクタ（比較実績なし）。 */
        public MonthPoint(
                YearMonth month,
                LocalDate fromInclusive,
                LocalDate toInclusive,
                int daysInBucket,
                int calendarDaysInMonth,
                double actualM,
                double planM,
                double actualCumM,
                double planCumM,
                double projectedCumM,
                double projectedM,
                boolean usesPlanForProjection,
                boolean incomplete,
                boolean isCurrentMonth) {
            this(
                    month,
                    fromInclusive,
                    toInclusive,
                    daysInBucket,
                    calendarDaysInMonth,
                    actualM,
                    planM,
                    actualCumM,
                    planCumM,
                    projectedCumM,
                    projectedM,
                    usesPlanForProjection,
                    incomplete,
                    isCurrentMonth,
                    0.0,
                    0.0);
        }

        /** 当月差異 = 実績 − 予定。 */
        public double diffM() {
            return actualM - planM;
        }

        /** 比較対象実績との差異 = 主実績 − 比較実績。 */
        public double actualCompareDiffM() {
            return actualM - compareActualM;
        }

        /** 過去月（完了月）の参考進捗率 (%)。当月・未来・予定0は {@code NaN}。上限 100%。 */
        public double monthProgressPct() {
            if (isCurrentMonth || usesPlanForProjection || planM <= EPS) {
                return Double.NaN;
            }
            return Math.min(100.0, actualM / planM * 100.0);
        }
    }

    /**
     * 月別集計結果。
     */
    public record MonthlyResult(
            List<MonthPoint> months,
            double actualTotalM,
            double planTotalM,
            double actualToDateM,
            double planToDateM,
            double remainingPlanM,
            double projectedTotalM,
            LocalDate today,
            LocalDate periodFrom,
            LocalDate periodTo,
            int actualRowsCounted,
            int planRowsCounted,
            LocalDate actualMinDate,
            LocalDate actualMaxDate,
            List<String> warnings,
            double compareActualTotalM,
            String compareSourceLabel) {

        public MonthlyResult {
            months = months != null ? List.copyOf(months) : List.of();
            warnings = warnings != null ? List.copyOf(warnings) : List.of();
            compareSourceLabel = compareSourceLabel != null ? compareSourceLabel : "";
        }

        /** 従来互換コンストラクタ（比較実績なし）。 */
        public MonthlyResult(
                List<MonthPoint> months,
                double actualTotalM,
                double planTotalM,
                double actualToDateM,
                double planToDateM,
                double remainingPlanM,
                double projectedTotalM,
                LocalDate today,
                LocalDate periodFrom,
                LocalDate periodTo,
                int actualRowsCounted,
                int planRowsCounted,
                LocalDate actualMinDate,
                LocalDate actualMaxDate,
                List<String> warnings) {
            this(
                    months,
                    actualTotalM,
                    planTotalM,
                    actualToDateM,
                    planToDateM,
                    remainingPlanM,
                    projectedTotalM,
                    today,
                    periodFrom,
                    periodTo,
                    actualRowsCounted,
                    planRowsCounted,
                    actualMinDate,
                    actualMaxDate,
                    warnings,
                    0.0,
                    "");
        }

        public static MonthlyResult empty(LocalDate today, LocalDate from, LocalDate to) {
            return new MonthlyResult(
                    List.of(), 0, 0, 0, 0, 0, 0, today, from, to, 0, 0, null, null, List.of(), 0.0, "");
        }

        public double progressPct() {
            if (planToDateM <= EPS) {
                return Double.NaN;
            }
            return Math.min(100.0, actualToDateM / planToDateM * 100.0);
        }

        public boolean progressDenominatorSufficient() {
            if (planToDateM <= EPS || planTotalM <= EPS) {
                return false;
            }
            return planToDateM / planTotalM >= PROGRESS_DENOMINATOR_MIN_RATIO;
        }

        public double projectedDiffM() {
            return projectedTotalM - planTotalM;
        }

        /** 比較対象実績との合計差異 = 主実績合計 − 比較実績合計。 */
        public double actualCompareDiffTotalM() {
            return actualTotalM - compareActualTotalM;
        }

        public boolean isEmpty() {
            return actualRowsCounted == 0 && planRowsCounted == 0;
        }
    }

    private static final double EPS = 1e-9;
    /** 期間は最長でこの日数に丸める（UI の誤操作でメモリを食い潰さない）。 */
    static final int MAX_DAYS = 1000;
    /** 進捗率を表示するために必要な「前日までの予定 ÷ 期間予定合計」の下限。 */
    static final double PROGRESS_DENOMINATOR_MIN_RATIO = 0.10;

    private static final String COL_MACHINE = "機械名";
    private static final String COL_PROCESS = "工程名";
    private static final String COL_ACTUAL_QTY_DETAIL = "実加工数";
    private static final String COL_ACTUAL_QTY = COL_ACTUAL_QTY_DETAIL;
    private static final String COL_ACTUAL_QTY_DAILY = "実加工量";
    private static final String COL_ACTUAL_DATE_DAILY = "加工日付";
    private static final String COL_ACTUAL_DATE = "加工日";
    private static final String COL_ACTUAL_START_DT = "加工開始日時";
    private static final String COL_WAREHOUSE = "倉庫";
    private static final String COL_TASK_ID = "依頼NO";
    private static final String COL_CONVERSION_QTY = "換算数量";
    private static final String COL_UNPROCESSED = "未加工";
    /** 完了判定に使う列。{@code 加工登録区分} は同一行で「完了」と「未完」が共存するため使わない。 */
    private static final String COL_COMPLETION_FLAG = "加工完了区分";
    private static final String TOTAL_ROW_PREFIX = "[合計]";
    private static final Pattern DATE_HEADER = Pattern.compile("\\d{4}/\\d{2}/\\d{2}");
    /** 機械名セルに行の他列が結合された不正値を弾く（yyyy/M/d または yyyy-M-d）。 */
    private static final Pattern EMBEDDED_CALENDAR_DATE =
            Pattern.compile("\\d{4}[/\\-]\\d{1,2}[/\\-]\\d{1,2}");
    /** 短い設備コード（例: {@code W9-1}, {@code EC-2}, {@code Z-9}）。 */
    private static final Pattern EQUIPMENT_CODE =
            Pattern.compile("(?i)^[A-Z]{1,6}\\d{0,3}(-\\d{1,4})?$");
    /** 工程名に混入する数値・寸法コード（例: {@code 000,2}, {@code 750,1}）。 */
    private static final Pattern NUMERIC_PROCESS_JUNK =
            Pattern.compile("^\\d+([.,]\\d+)*$");
    private static final Pattern ZERO_WIDTH = Pattern.compile("[\u200b\u200c\u200d\ufeff]");
    private static final Pattern DASH_LIKE = Pattern.compile("[\u2010-\u2015\u2212\u30fc\uff0d]");
    private static final Pattern WHITESPACE = Pattern.compile("\\s+");
    private static final Collator JA = Collator.getInstance(Locale.JAPAN);
    /**
     * 機械コンボに載せるラベルの上限。実機名（例: {@code スライス機1 湖南}）より十分長く、
     * 加工内容・得意先・完了日がカンマ結合された不正値は超える。
     */
    static final int MAX_PLAUSIBLE_MACHINE_LABEL_LEN = 48;
    /** 工程コンボに載せるラベルの上限（{@code 欠点表示} 等より十分長い）。 */
    static final int MAX_PLAUSIBLE_PROCESS_LABEL_LEN = 24;

    static final String WARN_ACTUAL_QTY_COLUMN_MISSING =
            "実績ソースに「" + COL_ACTUAL_QTY_DETAIL + "」列が無いため、実績は集計していません（累積値からの推定は行いません）。";

    static {
        JA.setStrength(Collator.PRIMARY);
    }

    /**
     * 日別集計結果を月別にロールアップする。
     *
     * <p>期間内の月を連続で列挙し、月内の実績・予定の合算と月末時点の累計を保持する。
     * 期間全体の合計・進捗率・見込指標は日次集計と完全一致する。
     */
    public static MonthlyResult rollUpMonthly(Result daily, Filter filter, LocalDate today) {
        Objects.requireNonNull(daily, "daily");
        Objects.requireNonNull(filter, "filter");
        LocalDate t = today != null ? today : LocalDate.now();
        YearMonth currentYm = YearMonth.from(t);

        LocalDate from = filter.from();
        LocalDate to = filter.to();
        if (to.isBefore(from)) {
            LocalDate tmp = from;
            from = to;
            to = tmp;
        }

        YearMonth startYm = YearMonth.from(from);
        YearMonth endYm = YearMonth.from(to);

        Map<YearMonth, List<DayPoint>> daysByMonth = new LinkedHashMap<>();
        for (DayPoint dp : daily.days()) {
            YearMonth ym = YearMonth.from(dp.date());
            daysByMonth.computeIfAbsent(ym, k -> new ArrayList<>()).add(dp);
        }

        List<MonthPoint> monthPoints = new ArrayList<>();
        double runningActualCum = 0.0;
        double runningPlanCum = 0.0;
        double runningProjectedCum = 0.0;
        double runningCompareActualCum = 0.0;

        for (YearMonth ym = startYm; !ym.isAfter(endYm); ym = ym.plusMonths(1)) {
            List<DayPoint> monthDays = daysByMonth.getOrDefault(ym, List.of());
            int calDays = ym.lengthOfMonth();
            LocalDate ymStart = ym.atDay(1);
            LocalDate ymEnd = ym.atEndOfMonth();

            LocalDate bStart = ymStart.isBefore(from) ? from : ymStart;
            LocalDate bEnd = ymEnd.isAfter(to) ? to : ymEnd;
            int daysInBucket = monthDays.size();
            boolean incomplete = daysInBucket < calDays;
            boolean isCurrent = ym.equals(currentYm);
            boolean usesPlan = !ym.isBefore(currentYm);

            double actSum = 0.0;
            double planSum = 0.0;
            double projContrib = 0.0;
            double compareActSum = 0.0;

            for (DayPoint dp : monthDays) {
                actSum += dp.actualM();
                planSum += dp.planM();
                compareActSum += dp.compareActualM();
                if (dp.date().isBefore(t)) {
                    projContrib += dp.actualM();
                } else if (dp.date().equals(t)) {
                    projContrib += Math.max(dp.actualM(), dp.planM());
                } else {
                    projContrib += dp.planM();
                }
            }

            runningActualCum += actSum;
            runningPlanCum += planSum;
            runningProjectedCum += projContrib;
            runningCompareActualCum += compareActSum;

            double lastActualCum =
                    monthDays.isEmpty() ? runningActualCum : monthDays.get(monthDays.size() - 1).actualCumM();
            double lastPlanCum =
                    monthDays.isEmpty() ? runningPlanCum : monthDays.get(monthDays.size() - 1).planCumM();
            double lastProjCum =
                    monthDays.isEmpty()
                            ? runningProjectedCum
                            : monthDays.get(monthDays.size() - 1).projectedCumM();
            double lastCompareActualCum =
                    monthDays.isEmpty()
                            ? runningCompareActualCum
                            : monthDays.get(monthDays.size() - 1).compareActualCumM();

            monthPoints.add(
                    new MonthPoint(
                            ym,
                            bStart,
                            bEnd,
                            daysInBucket,
                            calDays,
                            actSum,
                            planSum,
                            lastActualCum,
                            lastPlanCum,
                            lastProjCum,
                            projContrib,
                            usesPlan,
                            incomplete,
                            isCurrent,
                            compareActSum,
                            lastCompareActualCum));
        }

        return new MonthlyResult(
                monthPoints,
                daily.actualTotalM(),
                daily.planTotalM(),
                daily.actualToDateM(),
                daily.planToDateM(),
                daily.remainingPlanM(),
                daily.projectedTotalM(),
                t,
                from,
                to,
                daily.actualRowsCounted(),
                daily.planRowsCounted(),
                daily.actualMinDate(),
                daily.actualMaxDate(),
                daily.warnings(),
                daily.compareActualTotalM(),
                daily.compareSourceLabel());
    }

    private ProcessingTrendAggregator() {}

    /**
     * 日報実績と実績明細を同時に受けて集計し、選択された実績と他方との差分を計算する。
     */
    public static Result aggregate(
            ActualsSnapshot dailyReport,
            ActualsSnapshot detailActuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            Filter filter,
            LocalDate today) {
        return aggregate(dailyReport, detailActuals, aladdin, dispatch, filter, today, null, null);
    }

    /**
     * 安定日キャッシュ付き集計。{@code cache}/{@code pathIdentity} が揃うとき、
     * {@code today - 30} 日以前はソース再走査を省略する。
     */
    public static Result aggregate(
            ActualsSnapshot dailyReport,
            ActualsSnapshot detailActuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            Filter filter,
            LocalDate today,
            ProcessingTrendStableDayCache cache,
            String pathIdentity) {
        Objects.requireNonNull(filter, "filter");
        LocalDate t = today != null ? today : LocalDate.now();
        LocalDate from = filter.from();
        LocalDate to = filter.to();
        if (from.plusDays(MAX_DAYS - 1).isBefore(to)) {
            to = from.plusDays(MAX_DAYS - 1);
        }

        TreeMap<LocalDate, double[]> byDay = new TreeMap<>();
        for (LocalDate d = from; !d.isAfter(to); d = d.plusDays(1)) {
            byDay.put(d, new double[3]); // [0]=主実績, [1]=予定, [2]=比較対象実績
        }

        List<String> warnings = new ArrayList<>();
        boolean isDailyPrimary = filter.actualSource() == ActualSource.DAILY_REPORT;
        ActualsSnapshot primaryActuals = isDailyPrimary ? dailyReport : detailActuals;
        ActualsSnapshot compareActuals = isDailyPrimary ? detailActuals : dailyReport;
        String compareLabel = isDailyPrimary ? "実績明細" : "日報";

        // 移動平均のため、期間初日の直前 (窓幅-1) 日間の主実績も集計
        int maDays = filter.movingAverageDays();
        Map<LocalDate, double[]> priorActuals = new HashMap<>();
        for (int i = 1; i < maDays; i++) {
            priorActuals.put(from.minusDays(i), new double[1]);
        }

        ProcessingTrendStableDayCache.SeriesKey seriesKey = null;
        boolean skipStable = false;
        if (cache != null && pathIdentity != null && !pathIdentity.isBlank()) {
            seriesKey =
                    new ProcessingTrendStableDayCache.SeriesKey(
                            pathIdentity,
                            filter.actualSource(),
                            filter.planSource(),
                            normKey(filter.machine()),
                            normKey(filter.process()));
            skipStable = cache.tryFillStableDays(seriesKey, from, to, t, byDay);
            if (skipStable) {
                cache.fillPriorActuals(seriesKey, priorActuals, t);
            }
        }

        LocalDate stableEnd = ProcessingTrendStableDayCache.stableEndInclusive(t);
        ActualsAccumulation act =
                accumulateActuals(
                        primaryActuals,
                        filter,
                        byDay,
                        priorActuals,
                        warnings,
                        0,
                        isDailyPrimary,
                        skipStable ? stableEnd : null);
        accumulateActuals(
                compareActuals,
                filter,
                byDay,
                warnings,
                2,
                !isDailyPrimary,
                skipStable ? stableEnd : null);

        int planRows =
                filter.planSource() == PlanSource.DISPATCH
                        ? accumulateDispatch(dispatch, filter, byDay, skipStable ? stableEnd : null)
                        : accumulateAladdin(aladdin, filter, byDay, t, skipStable ? stableEnd : null);
        if (skipStable) {
            LocalDate observeEnd = to.isBefore(stableEnd) ? to : stableEnd;
            for (LocalDate d = from; !d.isAfter(observeEnd); d = d.plusDays(1)) {
                double[] slot = byDay.get(d);
                if (slot != null && Math.abs(slot[0]) > EPS) {
                    act.observe(d);
                }
            }
        }

        List<DayPoint> days = new ArrayList<>(byDay.size());
        double actCum = 0;
        double planCum = 0;
        double projCum = 0;
        double compareActCum = 0;
        double actualToDate = 0;
        double planToDate = 0;
        double remainingPlan = 0;
        for (Map.Entry<LocalDate, double[]> e : byDay.entrySet()) {
            LocalDate d = e.getKey();
            double actual = e.getValue()[0];
            double plan = e.getValue()[1];
            double compareActual = e.getValue()[2];
            boolean usesPlan = !d.isBefore(t);
            // 当日のみ: 実績が予定を上回っていれば（終業後など）実績を見込に採用する
            double projected = !usesPlan ? actual : d.equals(t) ? Math.max(actual, plan) : plan;
            actCum += actual;
            planCum += plan;
            projCum += projected;
            compareActCum += compareActual;
            if (usesPlan) {
                remainingPlan += projected;
            } else {
                actualToDate += actual;
                planToDate += plan;
            }

            // N 日間移動平均 (当日を含む直近 N 日間の主実績平均)
            double sumMa = 0.0;
            int countMa = 0;
            for (int k = 0; k < maDays; k++) {
                LocalDate past = d.minusDays(k);
                double[] s = byDay.get(past);
                if (s != null) {
                    sumMa += s[0];
                    countMa++;
                } else {
                    double[] ps = priorActuals.get(past);
                    if (ps != null) {
                        sumMa += ps[0];
                        countMa++;
                    }
                }
            }
            double actualMa = countMa > 0 ? (sumMa / countMa) : actual;

            days.add(
                    new DayPoint(
                            d,
                            actual,
                            plan,
                            actCum,
                            planCum,
                            projCum,
                            usesPlan,
                            compareActual,
                            compareActCum,
                            actualMa));
        }
        if (seriesKey != null && cache != null) {
            cache.putStableDays(seriesKey, byDay, t);
        }
        return new Result(
                days,
                actCum,
                planCum,
                actualToDate,
                planToDate,
                remainingPlan,
                actualToDate + remainingPlan,
                t,
                act.rowsCounted,
                planRows,
                act.minDate,
                act.maxDate,
                warnings,
                compareActCum,
                compareLabel);
    }

    /**
     * 従来互換オーバーロード（単一の実績スナップショットを集計）。
     * ヘッダ内容から日報／明細を自動判別して委譲する。
     */
    public static Result aggregate(
            ActualsSnapshot actuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            Filter filter,
            LocalDate today) {
        boolean hasDailyQty = actuals != null && colIdx(actuals.headers(), COL_ACTUAL_QTY_DAILY) >= 0;
        boolean hasDetailQty = actuals != null && colIdx(actuals.headers(), COL_ACTUAL_QTY_DETAIL) >= 0;
        ActualSource effSource = filter.actualSource();
        if (hasDailyQty && !hasDetailQty) {
            effSource = ActualSource.DAILY_REPORT;
        } else if (hasDetailQty && !hasDailyQty) {
            effSource = ActualSource.DETAIL;
        } else if (!hasDailyQty && !hasDetailQty) {
            boolean hasDailyDate = actuals != null && colIdx(actuals.headers(), COL_ACTUAL_DATE_DAILY) >= 0;
            effSource = hasDailyDate ? ActualSource.DAILY_REPORT : ActualSource.DETAIL;
        }
        Filter effFilter = new Filter(
                filter.from(),
                filter.to(),
                effSource,
                filter.planSource(),
                filter.machine(),
                filter.process(),
                filter.movingAverageDays());
        ActualsSnapshot daily = effSource == ActualSource.DAILY_REPORT ? actuals : null;
        ActualsSnapshot detail = effSource == ActualSource.DETAIL ? actuals : null;
        return aggregate(daily, detail, aladdin, dispatch, effFilter, today);
    }

    /** 4 ソースに現れる機械名の和集合（日本語照合順）。 */
    public static List<String> machineNames(
            ActualsSnapshot dailyReport,
            ActualsSnapshot detailActuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch) {
        return distinctColumnValues(dailyReport, detailActuals, aladdin, dispatch, COL_MACHINE);
    }

    /** 3 ソースに現れる機械名の和集合（従来互換）。 */
    public static List<String> machineNames(
            ActualsSnapshot actuals, AladdinSnapshot aladdin, DispatchSnapshot dispatch) {
        return machineNames(actuals, null, aladdin, dispatch);
    }

    /** 4 ソースに現れる工程名の和集合（日本語照合順）。 */
    public static List<String> processNames(
            ActualsSnapshot dailyReport,
            ActualsSnapshot detailActuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch) {
        return distinctColumnValues(dailyReport, detailActuals, aladdin, dispatch, COL_PROCESS);
    }

    /** 3 ソースに現れる工程名の和集合（従来互換）。 */
    public static List<String> processNames(
            ActualsSnapshot actuals, AladdinSnapshot aladdin, DispatchSnapshot dispatch) {
        return processNames(actuals, null, aladdin, dispatch);
    }

    // ---- 実績 ----------------------------------------------------------------------------

    private static final class ActualsAccumulation {
        int rowsCounted;
        LocalDate minDate;
        LocalDate maxDate;

        void observe(LocalDate d) {
            if (minDate == null || d.isBefore(minDate)) {
                minDate = d;
            }
            if (maxDate == null || d.isAfter(maxDate)) {
                maxDate = d;
            }
        }
    }

    /**
     * 実績は日別値（日報の「{@code 実加工量}」または明細の「{@code 実加工数}」）を採用する。
     */
    private static ActualsAccumulation accumulateActuals(
            ActualsSnapshot actuals,
            Filter f,
            TreeMap<LocalDate, double[]> byDay,
            Map<LocalDate, double[]> priorActuals,
            List<String> warnings,
            int slotIndex,
            boolean isDailyReport,
            LocalDate skipOnOrBefore) {
        ActualsAccumulation acc = new ActualsAccumulation();
        if (actuals == null || actuals.headers() == null || actuals.rows() == null) {
            return acc;
        }
        List<String> headers = actuals.headers();
        int iQty = resolveActualQtyCol(headers, isDailyReport);
        if (iQty < 0) {
            if (!actuals.rows().isEmpty() && slotIndex == 0) {
                warnings.add(isDailyReport
                        ? "日報ソースに「" + COL_ACTUAL_QTY_DAILY + "」列が無いため、実績は集計していません。"
                        : WARN_ACTUAL_QTY_COLUMN_MISSING);
            }
            return acc;
        }
        int iMachine = colIdx(headers, COL_MACHINE);
        int iProcess = colIdx(headers, COL_PROCESS);
        int iStartDt = colIdx(headers, COL_ACTUAL_START_DT);
        int iDailyDate = colIdx(headers, COL_ACTUAL_DATE_DAILY);
        int iKakouDate = colIdx(headers, COL_ACTUAL_DATE);
        String mk = normKey(f.machine());
        String pk = normKey(f.process());
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
            if (skipOnOrBefore != null && !d.isAfter(skipOnOrBefore)) {
                continue;
            }
            acc.observe(d);
            double[] slot = byDay.get(d);
            boolean inPeriod = slot != null;
            if (slot == null && priorActuals != null && slotIndex == 0) {
                slot = priorActuals.get(d);
            }
            if (slot == null) {
                continue;
            }
            double qty = parseDouble(cellAt(row, iQty));
            if (Math.abs(qty) <= EPS) {
                continue;
            }
            slot[slotIndex] += qty;
            if (inPeriod) {
                acc.rowsCounted++;
            }
        }
        return acc;
    }

    private static ActualsAccumulation accumulateActuals(
            ActualsSnapshot actuals,
            Filter f,
            TreeMap<LocalDate, double[]> byDay,
            List<String> warnings,
            int slotIndex,
            boolean isDailyReport,
            LocalDate skipOnOrBefore) {
        return accumulateActuals(
                actuals, f, byDay, null, warnings, slotIndex, isDailyReport, skipOnOrBefore);
    }

    private static int resolveActualQtyCol(List<String> headers, boolean preferDailyReport) {
        if (preferDailyReport) {
            int i = colIdx(headers, COL_ACTUAL_QTY_DAILY);
            if (i >= 0) return i;
            return colIdx(headers, COL_ACTUAL_QTY_DETAIL);
        } else {
            int i = colIdx(headers, COL_ACTUAL_QTY_DETAIL);
            if (i >= 0) return i;
            return colIdx(headers, COL_ACTUAL_QTY_DAILY);
        }
    }

    /**
     * 実績行の加工日: {@code 加工開始日時} の日付部、{@code 加工日付}、{@code 加工日} の順で解決。
     */
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

    // ---- 予定: アラジン（日付列グリッド） ------------------------------------------------

    private static int accumulateAladdin(
            AladdinSnapshot aladdin,
            Filter f,
            TreeMap<LocalDate, double[]> byDay,
            LocalDate today,
            LocalDate skipOnOrBefore) {
        if (aladdin == null || aladdin.headers() == null || aladdin.rows() == null) {
            return 0;
        }
        List<String> headers = aladdin.headers();
        int iMachine = colIdx(headers, COL_MACHINE);
        int iProcess = colIdx(headers, COL_PROCESS);
        int iWarehouse = colIdx(headers, COL_WAREHOUSE);
        int iTask = colIdx(headers, COL_TASK_ID);
        int iConv = colIdx(headers, COL_CONVERSION_QTY);
        int iDone = colIdx(headers, COL_ACTUAL_QTY);
        int iUnprocessed = colIdx(headers, COL_UNPROCESSED);
        int iCompletion = colIdx(headers, COL_COMPLETION_FLAG);

        // 全日付列（期間外も含む）。残予定の上限は行全体の当日以降合計に掛けるため期間外も必要
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
        boolean anyInPeriod = false;
        for (LocalDate d : allDateCols.keySet()) {
            if (byDay.containsKey(d)) {
                anyInPeriod = true;
                break;
            }
        }
        if (!anyInPeriod) {
            return 0;
        }
        List<LocalDate> futureDates = new ArrayList<>(allDateCols.tailMap(today, true).keySet());

        String mk = normKey(f.machine());
        String pk = normKey(f.process());
        int counted = 0;
        Map<LocalDate, Double> rowValues = new LinkedHashMap<>();
        for (List<String> row : aladdin.rows()) {
            if (row == null || isAladdinTotalRow(row, iWarehouse, iMachine, iTask)) {
                continue;
            }
            if (!matches(mk, cellAt(row, iMachine)) || !matches(pk, cellAt(row, iProcess))) {
                continue;
            }
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
            boolean any = false;
            for (Map.Entry<LocalDate, Double> e : rowValues.entrySet()) {
                LocalDate day = e.getKey();
                if (skipOnOrBefore != null && day != null && !day.isAfter(skipOnOrBefore)) {
                    continue;
                }
                double[] slot = byDay.get(day);
                if (slot == null || Math.abs(e.getValue()) <= EPS) {
                    continue;
                }
                slot[1] += e.getValue();
                any = true;
            }
            if (any) {
                counted++;
            }
        }
        return counted;
    }

    /** {@code [合計]} 行（倉庫が {@code [合計]} 始まり、または機械名・依頼NO とも空）を除外する。 */
    private static boolean isAladdinTotalRow(List<String> row, int iWarehouse, int iMachine, int iTask) {
        if (iWarehouse >= 0 && cellAt(row, iWarehouse).strip().startsWith(TOTAL_ROW_PREFIX)) {
            return true;
        }
        return iMachine >= 0
                && iTask >= 0
                && cellAt(row, iMachine).isBlank()
                && cellAt(row, iTask).isBlank();
    }

    /**
     * 当日以降の予定を行単位で {@code 未加工}（全数未加工ルール適用後）以下に丸める。超過分は遅い日付から削る。
     * 完了行（{@code 加工完了区分} に「完了」）は当日以降を 0 にする。過去日の値は触らない
     * （完了依頼の過去予定を落とすと前日までの予定がさらに欠けるため）。
     */
    private static void capRemainingPlan(
            List<String> row,
            Map<LocalDate, Double> rowValues,
            List<LocalDate> futureDates,
            int iConv,
            int iDone,
            int iUnprocessed,
            int iCompletion) {
        if (futureDates.isEmpty()) {
            return;
        }
        boolean completed = iCompletion >= 0 && cellAt(row, iCompletion).contains("完了");
        double cap;
        if (completed) {
            cap = 0.0;
        } else if (iUnprocessed >= 0) {
            double unprocessed = parseDouble(cellAt(row, iUnprocessed));
            double conv = iConv >= 0 ? parseDouble(cellAt(row, iConv)) : 0.0;
            double done = iDone >= 0 ? parseDouble(cellAt(row, iDone)) : 0.0;
            // 全数未加工ルール: 換算数量>0・実加工数=0・未加工=0 → 換算数量ぶんが未加工
            if (conv > EPS && Math.abs(done) <= EPS && Math.abs(unprocessed) <= EPS) {
                unprocessed = conv;
            }
            cap = Math.max(0.0, unprocessed);
        } else {
            return;
        }
        double future = 0.0;
        for (LocalDate d : futureDates) {
            future += rowValues.getOrDefault(d, 0.0);
        }
        double excess = future - cap;
        if (excess <= EPS) {
            return;
        }
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

    // ---- 予定: 配台結果（配台日 × 当日配台数量） -----------------------------------------

    private static int accumulateDispatch(
            DispatchSnapshot dispatch,
            Filter f,
            TreeMap<LocalDate, double[]> byDay,
            LocalDate skipOnOrBefore) {
        if (dispatch == null || dispatch.headers() == null || dispatch.rows() == null) {
            return 0;
        }
        List<String> headers = dispatch.headers();
        int iMachine = colIdx(headers, ResultDispatchSchema.COL_MACHINE);
        int iProcess = colIdx(headers, ResultDispatchSchema.COL_PROCESS);
        int iDate = colIdx(headers, ResultDispatchSchema.COL_DISPATCH_DATE);
        int iQty = colIdx(headers, ResultDispatchSchema.COL_DISPATCH_QTY);
        if (iDate < 0 || iQty < 0) {
            return 0;
        }
        List<List<String>> rows = normalizeLegacyDispatchRows(headers, dispatch.rows());
        String mk = normKey(f.machine());
        String pk = normKey(f.process());
        int counted = 0;
        for (List<String> row : rows) {
            if (row == null) {
                continue;
            }
            if (!matches(mk, cellAt(row, iMachine)) || !matches(pk, cellAt(row, iProcess))) {
                continue;
            }
            LocalDate d = parseDate(cellAt(row, iDate));
            if (d == null) {
                continue;
            }
            if (skipOnOrBefore != null && !d.isAfter(skipOnOrBefore)) {
                continue;
            }
            double[] slot = byDay.get(d);
            if (slot == null) {
                continue;
            }
            double v = parseDouble(cellAt(row, iQty));
            if (Math.abs(v) <= EPS) {
                continue;
            }
            slot[1] += v;
            counted++;
        }
        return counted;
    }

    /**
     * 旧 段階3 JSON（{@code 実配台数量} 列あり）は編集目標行とタイムライン実績行が重複するため、
     * 配台結果タブと同じ統合（孤立目標行の除去・実配台&gt;0 を主数量に）を掛けてから合算する。
     * 現行の段階2出力（列なし）はそのまま返す。
     */
    private static List<List<String>> normalizeLegacyDispatchRows(List<String> headers, List<List<String>> rows) {
        if (!ResultDispatchPlanningStageSupport.hasActualDispatchQtyColumn(headers)) {
            return rows;
        }
        List<String> cols = new ArrayList<>(headers);
        List<Map<String, String>> maps = new ArrayList<>(rows.size());
        for (List<String> row : rows) {
            if (row == null) {
                continue;
            }
            Map<String, String> m = new LinkedHashMap<>();
            for (int i = 0; i < cols.size(); i++) {
                m.put(cols.get(i), cellAt(row, i));
            }
            maps.add(m);
        }
        ResultDispatchInteractiveConsolidator.consolidatePlanAndTimelineRowsInPlace(cols, maps);
        ResultDispatchPlanningStageSupport.applyActualQtyDisplayQuantities(cols, maps);
        List<List<String>> out = new ArrayList<>(maps.size());
        for (Map<String, String> m : maps) {
            List<String> row = new ArrayList<>(headers.size());
            for (String h : headers) {
                row.add(m.getOrDefault(h, ""));
            }
            out.add(row);
        }
        return out;
    }

    // ---- 共通 ------------------------------------------------------------------------------

    private static List<String> distinctColumnValues(
            ActualsSnapshot dailyReport,
            ActualsSnapshot detailActuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            String col) {
        Map<String, String> byKey = new LinkedHashMap<>();
        if (dailyReport != null) {
            collectColumn(dailyReport.headers(), dailyReport.rows(), col, byKey);
        }
        if (detailActuals != null) {
            collectColumn(detailActuals.headers(), detailActuals.rows(), col, byKey);
        }
        if (aladdin != null) {
            collectColumn(aladdin.headers(), aladdin.rows(), col, byKey);
        }
        if (dispatch != null) {
            collectColumn(dispatch.headers(), dispatch.rows(), col, byKey);
        }
        TreeSet<String> sorted = new TreeSet<>(JA);
        sorted.addAll(byKey.values());
        return List.copyOf(sorted);
    }

    private static void collectColumn(
            List<String> headers, List<List<String>> rows, String col, Map<String, String> out) {
        int idx = colIdx(headers, col);
        if (idx < 0 || rows == null) {
            return;
        }
        boolean machineCol = COL_MACHINE.equals(col);
        boolean processCol = COL_PROCESS.equals(col);
        for (List<String> row : rows) {
            String raw = cellAt(row, idx);
            if (machineCol && !isPlausibleMachineLabel(raw)) {
                continue;
            }
            if (processCol && !isPlausibleProcessLabel(raw)) {
                continue;
            }
            String key = normKey(raw);
            if (!key.isEmpty()) {
                out.putIfAbsent(key, raw.strip());
            }
        }
    }

    /**
     * 機械コンボ用: ソースの「機械名」列に、加工内容・品種・集計見出しなどが混入することがあるため、
     * 工場の機械名として妥当なものだけ残す。
     */
    static boolean isPlausibleMachineLabel(String raw) {
        if (raw == null) {
            return false;
        }
        String s = raw.strip();
        if (s.isEmpty() || s.length() > MAX_PLAUSIBLE_MACHINE_LABEL_LEN) {
            return false;
        }
        if (s.indexOf(',') >= 0 || s.indexOf(':') >= 0 || s.indexOf('：') >= 0) {
            return false;
        }
        if (s.indexOf('(') >= 0 || s.indexOf('（') >= 0 || s.indexOf(')') >= 0 || s.indexOf('）') >= 0) {
            return false;
        }
        if (s.contains("株式会社") || s.contains("（株）") || s.contains("(株)")) {
            return false;
        }
        if (s.contains("1:完了") || s.contains("0:未完") || s.contains(":完了") || s.contains(":未完")) {
            return false;
        }
        if (s.contains("合計") || s.contains("品種") || s.contains("欠点数") || s.contains("接続点")) {
            return false;
        }
        if (EMBEDDED_CALENDAR_DATE.matcher(s).find()) {
            return false;
        }
        String key = normKey(s);
        if (key.isEmpty()) {
            return false;
        }
        // 実機名: サイト接尾・「機」を含む・短い設備コード（W9-1 / EC-2 等）
        if (key.endsWith("湖南") || key.endsWith("湘南") || key.endsWith("国分")) {
            return true;
        }
        if (key.contains("機")) {
            return true;
        }
        return EQUIPMENT_CODE.matcher(key).matches();
    }

    /**
     * 工程コンボ用: 「工程名」列に寸法・数量コード（{@code 000,2} 等）や他列の結合値が混入することがあるため除外する。
     */
    static boolean isPlausibleProcessLabel(String raw) {
        if (raw == null) {
            return false;
        }
        String s = raw.strip();
        if (s.isEmpty() || s.length() > MAX_PLAUSIBLE_PROCESS_LABEL_LEN) {
            return false;
        }
        if (s.indexOf(',') >= 0 || s.indexOf(':') >= 0 || s.indexOf('：') >= 0) {
            return false;
        }
        if (s.indexOf('(') >= 0 || s.indexOf('（') >= 0 || s.indexOf(')') >= 0 || s.indexOf('）') >= 0) {
            return false;
        }
        if (s.contains("株式会社") || s.contains("（株）") || s.contains("(株)")) {
            return false;
        }
        if (s.contains("1:完了") || s.contains("0:未完") || s.contains(":完了") || s.contains(":未完")) {
            return false;
        }
        if (s.contains("合計") || s.contains("品種") || s.contains("欠点数")) {
            return false;
        }
        if (EMBEDDED_CALENDAR_DATE.matcher(s).find()) {
            return false;
        }
        String key = normKey(s);
        if (key.isEmpty()) {
            return false;
        }
        // 工場サイト付き機械名が工程列に入っているケース
        if (key.endsWith("湖南") || key.endsWith("湘南") || key.endsWith("国分") || key.contains("機")) {
            return false;
        }
        if (NUMERIC_PROCESS_JUNK.matcher(key).matches()) {
            return false;
        }
        return true;
    }

    private static boolean matches(String wantedKey, String cell) {
        return wantedKey.isEmpty() || wantedKey.equals(normKey(cell));
    }

    /**
     * 機械名・工程名の照合キー（NFKC・ダッシュ類統一・空白正規化）。
     * {@code Ｗ９－１} / {@code W9‐1}（U+2010）/ {@code W9−1}（U+2212）/ 長音 {@code ー} を {@code -} に畳む。
     */
    static String normKey(String val) {
        if (val == null || val.isBlank()) {
            return "";
        }
        String t = Normalizer.normalize(val, Normalizer.Form.NFKC);
        t = t.replace('\u00a0', ' ').replace('\u3000', ' ');
        t = ZERO_WIDTH.matcher(t).replaceAll("");
        t = DASH_LIKE.matcher(t).replaceAll("-");
        return WHITESPACE.matcher(t).replaceAll(" ").strip();
    }

    /**
     * {@code yyyy/MM/dd} / {@code yyyy-MM-dd} / {@code yyyy/M/d}（末尾に時刻があっても可）。
     * 年は 4 桁必須。存在しない日付（2/30 など）は丸めずに {@code null}。
     */
    static LocalDate parseDate(String raw) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String s = raw.strip();
        int sp = s.indexOf(' ');
        if (sp > 0) {
            s = s.substring(0, sp);
        }
        int tIdx = s.indexOf('T');
        if (tIdx > 0) {
            s = s.substring(0, tIdx);
        }
        String[] parts = s.split("[/\\-]");
        if (parts.length != 3 || parts[0].strip().length() != 4) {
            return null;
        }
        try {
            int y = Integer.parseInt(parts[0].strip());
            int mo = Integer.parseInt(parts[1].strip());
            int d = Integer.parseInt(parts[2].strip());
            return LocalDate.of(y, mo, d);
        } catch (NumberFormatException | DateTimeException e) {
            return null;
        }
    }

    private static double parseDouble(String s) {
        if (s == null || s.isBlank()) {
            return 0.0;
        }
        try {
            return Double.parseDouble(s.strip().replace(",", ""));
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private static int colIdx(List<String> headers, String title) {
        if (headers == null || title == null) {
            return -1;
        }
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (h != null && title.equals(h.strip())) {
                return i;
            }
        }
        return -1;
    }

    private static String cellAt(List<String> row, int idx) {
        return (idx >= 0 && row != null && idx < row.size() && row.get(idx) != null)
                ? row.get(idx)
                : "";
    }

    private static String blankToNull(String s) {
        return s == null || s.isBlank() ? null : s.strip();
    }
}
