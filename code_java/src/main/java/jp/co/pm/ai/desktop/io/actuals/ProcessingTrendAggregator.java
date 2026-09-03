package jp.co.pm.ai.desktop.io.actuals;

import java.text.Collator;
import java.text.Normalizer;
import java.time.DateTimeException;
import java.time.LocalDate;
import java.util.ArrayList;
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
     * @param planSource 予定の取得元
     * @param machine 機械名（{@code null} / 空 = 全機械）
     * @param process 工程名（{@code null} / 空 = 全工程）
     */
    public record Filter(
            LocalDate from, LocalDate to, PlanSource planSource, String machine, String process) {

        public Filter {
            Objects.requireNonNull(from, "from");
            Objects.requireNonNull(to, "to");
            if (to.isBefore(from)) {
                LocalDate t = from;
                from = to;
                to = t;
            }
            planSource = planSource != null ? planSource : PlanSource.ALADDIN;
            machine = blankToNull(machine);
            process = blankToNull(process);
        }

        public boolean hasMachine() {
            return machine != null;
        }

        public boolean hasProcess() {
            return process != null;
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
     */
    public record DayPoint(
            LocalDate date,
            double actualM,
            double planM,
            double actualCumM,
            double planCumM,
            double projectedCumM,
            boolean usesPlanForProjection) {

        /** 当日差異 = 実績 − 予定。 */
        public double diffM() {
            return actualM - planM;
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
            List<String> warnings) {

        public Result {
            days = days != null ? List.copyOf(days) : List.of();
            warnings = warnings != null ? List.copyOf(warnings) : List.of();
        }

        public static Result empty(LocalDate today) {
            return new Result(List.of(), 0, 0, 0, 0, 0, 0, today, 0, 0, null, null, List.of());
        }

        /** 進捗率 (%) = 当日より前の実績 ÷ 当日より前の予定。分母 0 のとき {@code NaN}。 */
        public double progressPct() {
            if (planToDateM <= EPS) {
                return Double.NaN;
            }
            return actualToDateM / planToDateM * 100.0;
        }

        /**
         * 進捗率の分母（前日までの予定）が期間予定合計に対して十分か。
         * アラジン予定は完了依頼が抽出から消えるため、月初などで分母が極端に小さいと数百 % になる。
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
    }

    private static final double EPS = 1e-9;
    /** 期間は最長でこの日数に丸める（UI の誤操作でメモリを食い潰さない）。 */
    static final int MAX_DAYS = 1000;
    /** 進捗率を表示するために必要な「前日までの予定 ÷ 期間予定合計」の下限。 */
    static final double PROGRESS_DENOMINATOR_MIN_RATIO = 0.10;

    private static final String COL_MACHINE = "機械名";
    private static final String COL_PROCESS = "工程名";
    private static final String COL_ACTUAL_QTY = "実加工数";
    private static final String COL_WAREHOUSE = "倉庫";
    private static final String COL_TASK_ID = "依頼NO";
    private static final String COL_CONVERSION_QTY = "換算数量";
    private static final String COL_UNPROCESSED = "未加工";
    /** 完了判定に使う列。{@code 加工登録区分} は同一行で「完了」と「未完」が共存するため使わない。 */
    private static final String COL_COMPLETION_FLAG = "加工完了区分";
    private static final String TOTAL_ROW_PREFIX = "[合計]";
    private static final Pattern DATE_HEADER = Pattern.compile("\\d{4}/\\d{2}/\\d{2}");
    private static final Pattern ZERO_WIDTH = Pattern.compile("[\u200b\u200c\u200d\ufeff]");
    private static final Pattern DASH_LIKE = Pattern.compile("[\u2010-\u2015\u2212\u30fc\uff0d]");
    private static final Pattern WHITESPACE = Pattern.compile("\\s+");
    private static final Collator JA = Collator.getInstance(Locale.JAPAN);

    static final String WARN_ACTUAL_QTY_COLUMN_MISSING =
            "実績ソースに「" + COL_ACTUAL_QTY + "」列が無いため、実績は集計していません（累積値からの推定は行いません）。";

    static {
        JA.setStrength(Collator.PRIMARY);
    }

    private ProcessingTrendAggregator() {}

    public static Result aggregate(
            ActualsSnapshot actuals,
            AladdinSnapshot aladdin,
            DispatchSnapshot dispatch,
            Filter filter,
            LocalDate today) {
        Objects.requireNonNull(filter, "filter");
        LocalDate t = today != null ? today : LocalDate.now();
        LocalDate from = filter.from();
        LocalDate to = filter.to();
        if (from.plusDays(MAX_DAYS - 1).isBefore(to)) {
            to = from.plusDays(MAX_DAYS - 1);
        }

        TreeMap<LocalDate, double[]> byDay = new TreeMap<>();
        for (LocalDate d = from; !d.isAfter(to); d = d.plusDays(1)) {
            byDay.put(d, new double[2]);
        }

        List<String> warnings = new ArrayList<>();
        ActualsAccumulation act = accumulateActuals(actuals, filter, byDay, warnings);
        int planRows =
                filter.planSource() == PlanSource.DISPATCH
                        ? accumulateDispatch(dispatch, filter, byDay)
                        : accumulateAladdin(aladdin, filter, byDay, t);

        List<DayPoint> days = new ArrayList<>(byDay.size());
        double actCum = 0;
        double planCum = 0;
        double projCum = 0;
        double actualToDate = 0;
        double planToDate = 0;
        double remainingPlan = 0;
        for (Map.Entry<LocalDate, double[]> e : byDay.entrySet()) {
            LocalDate d = e.getKey();
            double actual = e.getValue()[0];
            double plan = e.getValue()[1];
            boolean usesPlan = !d.isBefore(t);
            // 当日のみ: 実績が予定を上回っていれば（終業後など）実績を見込に採用する
            double projected = !usesPlan ? actual : d.equals(t) ? Math.max(actual, plan) : plan;
            actCum += actual;
            planCum += plan;
            projCum += projected;
            if (usesPlan) {
                remainingPlan += projected;
            } else {
                actualToDate += actual;
                planToDate += plan;
            }
            days.add(new DayPoint(d, actual, plan, actCum, planCum, projCum, usesPlan));
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
                warnings);
    }

    /** 3 ソースに現れる機械名の和集合（日本語照合順）。 */
    public static List<String> machineNames(
            ActualsSnapshot actuals, AladdinSnapshot aladdin, DispatchSnapshot dispatch) {
        return distinctColumnValues(actuals, aladdin, dispatch, COL_MACHINE);
    }

    /** 3 ソースに現れる工程名の和集合（日本語照合順）。 */
    public static List<String> processNames(
            ActualsSnapshot actuals, AladdinSnapshot aladdin, DispatchSnapshot dispatch) {
        return distinctColumnValues(actuals, aladdin, dispatch, COL_PROCESS);
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
     * 実績は「{@code 実加工数}」列（日別値）だけを採用する。列が無いときは集計せず警告を返す。
     * {@code 累積実績} や {@code 換算数量×完了率} は累計値なので日別に足すと多重計上になる。
     */
    private static ActualsAccumulation accumulateActuals(
            ActualsSnapshot actuals, Filter f, TreeMap<LocalDate, double[]> byDay, List<String> warnings) {
        ActualsAccumulation acc = new ActualsAccumulation();
        if (actuals == null || actuals.headers() == null || actuals.rows() == null) {
            return acc;
        }
        List<String> headers = actuals.headers();
        int iQty = colIdx(headers, COL_ACTUAL_QTY);
        if (iQty < 0) {
            if (!actuals.rows().isEmpty()) {
                warnings.add(WARN_ACTUAL_QTY_COLUMN_MISSING);
            }
            return acc;
        }
        int iMachine = colIdx(headers, COL_MACHINE);
        int iProcess = colIdx(headers, COL_PROCESS);
        String mk = normKey(f.machine());
        String pk = normKey(f.process());
        for (List<String> row : actuals.rows()) {
            if (row == null) {
                continue;
            }
            if (!matches(mk, cellAt(row, iMachine)) || !matches(pk, cellAt(row, iProcess))) {
                continue;
            }
            LocalDate d = EquipmentStatusDashboardBuilder.rowActualDate(headers, row);
            if (d == null) {
                continue;
            }
            acc.observe(d);
            double[] slot = byDay.get(d);
            if (slot == null) {
                continue;
            }
            double qty = parseDouble(cellAt(row, iQty));
            if (Math.abs(qty) <= EPS) {
                continue;
            }
            slot[0] += qty;
            acc.rowsCounted++;
        }
        return acc;
    }

    // ---- 予定: アラジン（日付列グリッド） ------------------------------------------------

    private static int accumulateAladdin(
            AladdinSnapshot aladdin, Filter f, TreeMap<LocalDate, double[]> byDay, LocalDate today) {
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
                double[] slot = byDay.get(e.getKey());
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
            DispatchSnapshot dispatch, Filter f, TreeMap<LocalDate, double[]> byDay) {
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
            ActualsSnapshot actuals, AladdinSnapshot aladdin, DispatchSnapshot dispatch, String col) {
        Map<String, String> byKey = new LinkedHashMap<>();
        if (actuals != null) {
            collectColumn(actuals.headers(), actuals.rows(), col, byKey);
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
        for (List<String> row : rows) {
            String raw = cellAt(row, idx);
            String key = normKey(raw);
            if (!key.isEmpty()) {
                out.putIfAbsent(key, raw.strip());
            }
        }
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
