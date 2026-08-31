package jp.co.pm.ai.desktop.dispatch;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.io.ExcelCellReadSupport;
import jp.co.pm.ai.desktop.ui.PlanInputDateColumnSupport;

/**
 * アラジン入力用 Excel（機械名ごとのシート×日別2段セル）のモデルを、結果_配台表の日別行から組み立てる。
 *
 * <p>1 行 = 依頼NO×工程名（同一機械シート内）。日付列は当日〜全データの最遅配台日で全シート共通。
 * 各日付セルは上段（現アラ計）・下段（シス計）の2段表示。
 */
public final class DispatchAladdinEntrySheetBuilder {

    private static final String COL_TID = "依頼NO";
    private static final String COL_INPUT_DATE = "原反投入日";
    private static final String COL_KAITO_NOKI = "回答納期";
    private static final String COL_PROCESS_COMPLETE = "加工完了日";
    private static final String COL_CONVERSION_QTY = "換算数量";
    private static final String COL_COMPLETED_QTY = "実加工数";

    /** 数量照合の許容誤差。 */
    public static final double QTY_MATCH_EPS = 0.01;

    /** 完了日チェックの判定・表示色。 */
    public enum CompletionDateCheckStatus {
        BLANK,
        OK_WHITE,
        OK_YELLOW,
        NG_RED
    }

    /** セル上段の接頭辞。 */
    public static final String ALADDIN_LINE_PREFIX = "（現アラ計）";

    /** セル下段の接頭辞。 */
    public static final String SYSTEM_LINE_PREFIX = "（シス計）";

    /** 依頼書目次の納期など、年なし {@code M/d} 表示。 */
    private static final Pattern MONTH_DAY_SLASH =
            Pattern.compile("^\\s*(\\d{1,2})/(\\d{1,2})\\s*$");

    private DispatchAladdinEntrySheetBuilder() {}

    /** 依頼書原本 目次シート由来の補足情報（回答納期 K列・契約No N列）。 */
    public record IndexInfo(String kaitoNoki, String contractNo) {}

    /** 日付セル（上段=現アラ計、下段=シス計）。 */
    public record EntryCell(double aladdinQty, double systemQty) {

        public boolean isEmpty() {
            return Math.abs(aladdinQty) <= QTY_MATCH_EPS && Math.abs(systemQty) <= QTY_MATCH_EPS;
        }

        /** シス計あり かつ 現アラ計と不一致（＝アラジンへ入力が必要な箇所）。 */
        public boolean mismatch() {
            return Math.abs(systemQty) > QTY_MATCH_EPS
                    && Math.abs(aladdinQty - systemQty) > QTY_MATCH_EPS;
        }

        /** 2段セル文字列。両段 0 のときは空文字。 */
        public String cellText() {
            if (isEmpty()) {
                return "";
            }
            return ALADDIN_LINE_PREFIX
                    + ResultDispatchNormalizer.formatQty(aladdinQty)
                    + "\n"
                    + SYSTEM_LINE_PREFIX
                    + ResultDispatchNormalizer.formatQty(systemQty);
        }
    }

    /** 1 タスク行（依頼NO×工程名、機械シート内）。 */
    public record EntryRow(
            String taskId,
            String contractNo,
            String processName,
            String inputDate,
            String kaitoNoki,
            String processCompleteDate,
            double conversionQty,
            double completedQty,
            double dispatchTotal,
            Map<LocalDate, EntryCell> cells,
            LocalDate earliestDispatchDate,
            int referenceYear) {

        /** 配台合計 + 加工完了数量 が 換算数量 と一致（誤差 {@link #QTY_MATCH_EPS} 未満）なら OK。 */
        public boolean quantityOk() {
            return Math.abs(dispatchTotal + completedQty - conversionQty) < QTY_MATCH_EPS;
        }

        /** 数量チェックセル文字列（{@code OK} / {@code NG (差 -200)}）。差 = 配台合計+加工完了−換算数量。 */
        public String quantityCheckText() {
            if (quantityOk()) {
                return "OK";
            }
            double diff = dispatchTotal + completedQty - conversionQty;
            String signed =
                    (diff > 0 ? "+" : "") + ResultDispatchNormalizer.formatQty(diff);
            return "NG (差 " + signed + ")";
        }

        /** 実効完了日（加工完了日と最終シス計日の遅い方）が回答納期の一日前以前か。 */
        public boolean completionDateCheckOk() {
            return switch (completionDateCheckStatus()) {
                case OK_WHITE, OK_YELLOW -> true;
                case BLANK, NG_RED -> false;
            };
        }

        /**
         * 完了日チェック（{@code OK} / {@code NG}）。
         * 回答納期が空欄のときは空文字。実効完了日が取れないときも空文字。
         *
         * <p>実効完了日 = {@code max(加工完了日, 最終シス計日)}。画面のシス計日が加工完了日より遅い場合は
         * そちらを優先する（配台日が納期当日なのに加工完了日だけ早いと誤って OK になるのを防ぐ）。
         */
        public String completionDateCheckText() {
            return switch (completionDateCheckStatus()) {
                case OK_WHITE, OK_YELLOW -> "OK";
                case NG_RED -> "NG";
                case BLANK -> "";
            };
        }

        /** 完了日チェックの判定とExcel背景色の区分。 */
        public CompletionDateCheckStatus completionDateCheckStatus() {
            return DispatchAladdinEntrySheetBuilder.completionDateCheckStatus(
                    processCompleteDate, kaitoNoki, referenceYear, lastSystemDispatchDate());
        }

        /** シス計が 0 超の日付のうち最遅。無ければ null。 */
        LocalDate lastSystemDispatchDate() {
            if (cells == null || cells.isEmpty()) {
                return null;
            }
            LocalDate last = null;
            for (Map.Entry<LocalDate, EntryCell> e : cells.entrySet()) {
                EntryCell c = e.getValue();
                if (c == null || Math.abs(c.systemQty()) <= QTY_MATCH_EPS) {
                    continue;
                }
                if (last == null || e.getKey().isAfter(last)) {
                    last = e.getKey();
                }
            }
            return last;
        }
    }

    /** 機械名 1 つぶんのシート。 */
    public record MachineSheet(String machineName, List<EntryRow> rows) {}

    /** ブック全体（日付列は全シート共通）。 */
    public record EntryWorkbook(List<LocalDate> dates, List<MachineSheet> sheets) {}

    /** 機械名が空のときのシート名。 */
    public static final String MACHINE_NAME_FALLBACK = "(機械名なし)";

    /**
     * 結果_配台表の日別行からモデルを組み立てる。
     *
     * @param columns 結果_配台表の列
     * @param rows 日別行（段階3表示変換済み）
     * @param aladdinLookup {@link AladdinShapedPlanQtyLookup#buildLookup} の結果
     * @param indexByTid 依頼NO 正規化キー → 目次情報（null 可）
     * @param today 日付列の起点（当日）
     */
    public static EntryWorkbook build(
            List<String> columns,
            List<Map<String, String>> rows,
            Map<String, Map<String, Map<String, Map<String, Double>>>> aladdinLookup,
            Map<String, IndexInfo> indexByTid,
            LocalDate today) {
        if (rows == null || rows.isEmpty() || today == null) {
            return new EntryWorkbook(List.of(), List.of());
        }
        Map<String, Map<String, Map<String, Map<String, Double>>>> al =
                aladdinLookup != null ? aladdinLookup : Map.of();
        Map<String, IndexInfo> idx = indexByTid != null ? indexByTid : Map.of();

        List<LocalDate> dates = dateAxis(rows, today);

        // 機械名 → (依頼NO×工程名) → 日別行
        Map<String, Map<String, List<Map<String, String>>>> byMachine = new TreeMap<>();
        for (Map<String, String> row : rows) {
            String machine = nz(row.get(ResultDispatchSchema.COL_MACHINE)).strip();
            String mkey = machine.isEmpty() ? MACHINE_NAME_FALLBACK : machine;
            String tid = nz(row.get(COL_TID)).strip();
            String proc = nz(row.get(ResultDispatchSchema.COL_PROCESS)).strip();
            String gk = tid + "\u0000" + proc;
            byMachine
                    .computeIfAbsent(mkey, k -> new LinkedHashMap<>())
                    .computeIfAbsent(gk, k -> new ArrayList<>())
                    .add(row);
        }

        List<MachineSheet> sheets = new ArrayList<>(byMachine.size());
        for (Map.Entry<String, Map<String, List<Map<String, String>>>> me : byMachine.entrySet()) {
            String machine = me.getKey();
            List<EntryRow> outRows = new ArrayList<>(me.getValue().size());
            for (List<Map<String, String>> group : me.getValue().values()) {
                outRows.add(buildRow(machine, group, dates, al, idx, today));
            }
            outRows.sort(
                    Comparator.comparing(
                                    EntryRow::earliestDispatchDate,
                                    Comparator.nullsLast(Comparator.naturalOrder()))
                            .thenComparing(EntryRow::taskId, String.CASE_INSENSITIVE_ORDER)
                            .thenComparing(EntryRow::processName, String.CASE_INSENSITIVE_ORDER));
            sheets.add(new MachineSheet(machine, List.copyOf(outRows)));
        }
        return new EntryWorkbook(dates, List.copyOf(sheets));
    }

    /** 日付列: 当日〜全データの最遅配台日（配台日がすべて過去なら当日のみ）。 */
    static List<LocalDate> dateAxis(List<Map<String, String>> rows, LocalDate today) {
        LocalDate max = null;
        for (Map<String, String> row : rows) {
            LocalDate d =
                    ResultDispatchPivot.parseIsoDate(
                            row.get(ResultDispatchSchema.COL_DISPATCH_DATE));
            if (d != null && (max == null || d.isAfter(max))) {
                max = d;
            }
        }
        LocalDate last = max != null && max.isAfter(today) ? max : today;
        List<LocalDate> out = new ArrayList<>();
        for (LocalDate d = today; !d.isAfter(last); d = d.plusDays(1)) {
            out.add(d);
        }
        return List.copyOf(out);
    }

    private static EntryRow buildRow(
            String machine,
            List<Map<String, String>> group,
            List<LocalDate> dates,
            Map<String, Map<String, Map<String, Map<String, Double>>>> aladdinLookup,
            Map<String, IndexInfo> indexByTid,
            LocalDate today) {
        Map<String, String> first = group.getFirst();
        String tid = nz(first.get(COL_TID)).strip();
        String proc = nz(first.get(ResultDispatchSchema.COL_PROCESS)).strip();

        // シス計: 配台日 → 当日配台数量の合算（表示範囲外の日も配台合計に含める）
        Map<LocalDate, Double> systemByDate = new TreeMap<>();
        double dispatchTotal = 0;
        LocalDate earliest = null;
        for (Map<String, String> row : group) {
            double qty =
                    ResultDispatchNormalizer.parseDouble(
                            row.get(ResultDispatchSchema.COL_DISPATCH_QTY));
            dispatchTotal += qty;
            LocalDate d =
                    ResultDispatchPivot.parseIsoDate(
                            row.get(ResultDispatchSchema.COL_DISPATCH_DATE));
            if (d != null) {
                if (Math.abs(qty) > 1e-12) {
                    systemByDate.merge(d, qty, Double::sum);
                }
                if (earliest == null || d.isBefore(earliest)) {
                    earliest = d;
                }
            }
        }
        dispatchTotal = ResultDispatchNormalizer.snapDispatchQtyM(dispatchTotal);

        Map<LocalDate, EntryCell> cells = new LinkedHashMap<>();
        String tidLookupKey = AladdinShapedPlanQtyLookup.normalizeTaskIdKey(tid);
        for (LocalDate d : dates) {
            double system = systemByDate.getOrDefault(d, 0d);
            double aladdin =
                    AladdinShapedPlanQtyLookup.lookup(
                            aladdinLookup, machine, tidLookupKey, isoSlash(d), proc);
            cells.put(d, new EntryCell(aladdin, system));
        }

        IndexInfo info = indexByTid.get(AladdinShapedPlanQtyLookup.normalizeTaskIdKey(tid));
        String kaitoNoki =
                info != null && info.kaitoNoki() != null && !info.kaitoNoki().isBlank()
                        ? info.kaitoNoki().strip()
                        : nz(first.get(COL_KAITO_NOKI)).strip();
        String contractNo =
                info != null && info.contractNo() != null ? info.contractNo().strip() : "";

        int referenceYear = resolveReferenceYear(first, today, earliest);

        return new EntryRow(
                tid,
                contractNo,
                proc,
                nz(first.get(COL_INPUT_DATE)).strip(),
                kaitoNoki,
                nz(first.get(COL_PROCESS_COMPLETE)).strip(),
                ResultDispatchNormalizer.parseDouble(first.get(COL_CONVERSION_QTY)),
                ResultDispatchNormalizer.parseDouble(first.get(COL_COMPLETED_QTY)),
                dispatchTotal,
                Map.copyOf(cells),
                earliest,
                referenceYear);
    }

    /**
     * 実効完了日が回答納期の一日前以前か（{@code OK} / {@code NG}）。
     * 回答納期が空欄、または実効完了日が取れないときは空文字。
     *
     * <p>依頼書目次の {@code M/d} や {@code yyyy/M/d} なども解釈する。
     */
    static String completionDateOneDayBeforeAnswerCheck(
            String processCompleteDate, String answerNoki, int referenceYear) {
        return completionDateOneDayBeforeAnswerCheck(
                processCompleteDate, answerNoki, referenceYear, null);
    }

    /**
     * @param lastSystemDispatch シス計のある最遅配台日（null 可）。加工完了日より遅いときはこちらを実効完了日とする。
     */
    static String completionDateOneDayBeforeAnswerCheck(
            String processCompleteDate,
            String answerNoki,
            int referenceYear,
            LocalDate lastSystemDispatch) {
        return switch (completionDateCheckStatus(
                processCompleteDate, answerNoki, referenceYear, lastSystemDispatch)) {
            case OK_WHITE, OK_YELLOW -> "OK";
            case NG_RED -> "NG";
            case BLANK -> "";
        };
    }

    private static CompletionDateCheckStatus completionDateCheckStatus(
            String processCompleteDate,
            String answerNoki,
            int referenceYear,
            LocalDate lastSystemDispatch) {
        if (answerNoki == null || answerNoki.isBlank()) {
            return CompletionDateCheckStatus.BLANK;
        }
        LocalDate answer = parseAladdinEntryDate(answerNoki, referenceYear);
        int completeYear = answer != null ? answer.getYear() : referenceYear;
        LocalDate complete = parseAladdinEntryDate(processCompleteDate, completeYear);
        LocalDate effective = complete;
        if (lastSystemDispatch != null
                && (effective == null || lastSystemDispatch.isAfter(effective))) {
            effective = lastSystemDispatch;
        }
        if (effective == null || answer == null) {
            return CompletionDateCheckStatus.BLANK;
        }
        LocalDate oneDayBeforeAnswer = answer.minusDays(1);
        if (effective.isBefore(oneDayBeforeAnswer)) {
            return CompletionDateCheckStatus.OK_WHITE;
        }
        if (effective.equals(oneDayBeforeAnswer)) {
            return CompletionDateCheckStatus.OK_YELLOW;
        }
        return CompletionDateCheckStatus.NG_RED;
    }

    /**
     * 年なし {@code M/d} 補完用。原反投入日・受注日・加工完了日・最遅配台日の順で年を推定する。
     */
    static int resolveReferenceYear(
            Map<String, String> row, LocalDate today, LocalDate earliestDispatch) {
        int fallback = today != null ? today.getYear() : LocalDate.now().getYear();
        if (row == null) {
            return fallback;
        }
        for (String col : List.of(COL_INPUT_DATE, "受注日", COL_PROCESS_COMPLETE)) {
            LocalDate d = parseAladdinEntryDate(nz(row.get(col)), fallback);
            if (d != null) {
                return d.getYear();
            }
        }
        if (earliestDispatch != null) {
            return earliestDispatch.getYear();
        }
        return fallback;
    }

    /**
     * アラジン入力用 Excel の日付照合向けパース。
     *
     * <p>{@code yyyy-MM-dd} / {@code yyyy/MM/dd} / {@code yyyy/M/d} / 年なし {@code M/d} に対応。
     */
    static LocalDate parseAladdinEntryDate(String raw, int referenceYear) {
        if (raw == null || raw.isBlank()) {
            return null;
        }
        String t = ExcelCellReadSupport.stripMidnightDateTimeSuffix(raw.strip()).strip();
        if (t.isEmpty()) {
            return null;
        }
        int space = t.indexOf(' ');
        if (space > 0) {
            t = t.substring(0, space).strip();
        }
        LocalDate d = ResultDispatchPivot.parseIsoDate(t);
        if (d != null) {
            return d;
        }
        d = PlanInputDateColumnSupport.parseCellValue(t).orElse(null);
        if (d != null) {
            return d;
        }
        String[] parts = t.split("[/.\\-]");
        if (parts.length == 3) {
            try {
                int y = Integer.parseInt(parts[0].strip());
                int mo = Integer.parseInt(parts[1].strip());
                int day = Integer.parseInt(parts[2].strip());
                if (y >= 1900 && mo >= 1 && mo <= 12 && day >= 1 && day <= 31) {
                    return LocalDate.of(y, mo, day);
                }
            } catch (Exception ignored) {
                // fall through
            }
        }
        Matcher m = MONTH_DAY_SLASH.matcher(t);
        if (m.matches()) {
            try {
                return LocalDate.of(
                        referenceYear,
                        Integer.parseInt(m.group(1)),
                        Integer.parseInt(m.group(2)));
            } catch (Exception ignored) {
                return null;
            }
        }
        return null;
    }

    private static String isoSlash(LocalDate d) {
        return String.format("%04d/%02d/%02d", d.getYear(), d.getMonthValue(), d.getDayOfMonth());
    }

    private static String nz(String s) {
        return s != null ? s : "";
    }
}
