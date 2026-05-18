package jp.co.pm.ai.desktop.io.gantt;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.Normalizer;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeSet;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import jp.co.pm.ai.desktop.io.JsonTableIo;
/**
 * {@code *_equipment_gantt_contract.json}（設備ガント描画契約）から、
 * {@link jp.co.pm.ai.desktop.ui.EquipmentGraphicGanttPane} が期待する
 * 「日付・機械名・工程名・タスク概覝・HH:MM 列…」形式の 1 シート分を組み立てる。
 *
 * <p>Excel シートと完全一致は目指さず、タイムラインイベントからスロットを充填する近似。
 * xlsx 不要でグラフィック表示に使える。
 *
 * <p>{@code sorted_dates} は計画ホライズン全体を含むことがあり、イベントがまだ無い暦日が先頭に並ぶ。
 * その日はタイムラインがすべて空になるため、グラフィック用の暦日は {@code timeline_events} に現れる日付に限定する。
 *
 * <p>設備列キーと {@code machine} の対応づけは、Excel 出力側の
 * {@code _eq_grid_events_for_equipment_column} と同様に正規化照合する（厳密一致のみだと表記ゆれで行が空になる）。
 */
public final class EquipmentGanttContractSheetTableBuilder {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String COL_DATE = "日付";
    private static final String COL_MACH = "機械名";
    private static final String COL_PROC = "工程名";
    private static final String COL_TASK = "タスク概覝";

    private static final int SLOT_MINUTES = 10;

    /** 1 暦日あたりの分（壁時計）。スロット列の上限計算に使う。 */
    private static final int MINUTES_PER_DAY = 24 * 60;

    /**
     * 暦日内で最後に許容するスロット開始の「0:00 からの分」。{@link #ceilToSlot} が 24:00 相当まで丸めて
     * {@link LocalTime#of(int, int)} が HourOfDay=24 で例外になるのを防ぐ。
     */
    private static final int LAST_SLOT_START_MINUTE_OF_DAY = MINUTES_PER_DAY - SLOT_MINUTES;

    private static final LocalTime FALLBACK_DAY_START = LocalTime.of(8, 0);
    private static final LocalTime FALLBACK_DAY_END = LocalTime.of(21, 0);

    private EquipmentGanttContractSheetTableBuilder() {}

    /**
     * 契約 JSON から表と担当バッジ列（スロットごと）を返す。
     *
     * @deprecated 呼び出し側は {@link #buildBundleFromContractPath(Path)} を使用すること。
     */
    @Deprecated
    public static JsonTableIo.SheetTable buildFromContractPath(Path contractPath) throws IOException {
        return buildBundleFromContractPath(contractPath).table();
    }

    public static EquipmentGanttSheetBundle buildBundleFromContractPath(Path contractPath)
            throws IOException {
        JsonNode root = JSON.readTree(Files.readString(contractPath, StandardCharsets.UTF_8));
        JsonNode packed = root.get("kwargs_packed");
        if (packed == null || !packed.isObject()) {
            throw new IOException("契約 JSON に kwargs_packed がありません: " + contractPath);
        }
        JsonNode eventsNode = packed.get("timeline_events");
        JsonNode equipNode = packed.get("equipment_list");
        JsonNode datesNode = packed.get("sorted_dates");
        if (eventsNode == null || !eventsNode.isArray()) {
            throw new IOException("契約に timeline_events がありません");
        }
        if (equipNode == null || !equipNode.isArray() || equipNode.isEmpty()) {
            throw new IOException("契約に equipment_list がありません");
        }
        if (datesNode == null || !datesNode.isArray() || datesNode.isEmpty()) {
            throw new IOException("契約に sorted_dates がありません");
        }

        List<TimelineEvent> events = new ArrayList<>();
        for (JsonNode en : eventsNode) {
            TimelineEvent te = TimelineEvent.from(en);
            if (te != null) {
                events.add(te);
            }
        }

        events = applyGapAwareMachiningLabels(events);

        List<LocalDate> sortedDates = new ArrayList<>();
        for (JsonNode dn : datesNode) {
            Object d = GanttContractValueDecoder.decodeValue(dn);
            LocalDate ld = GanttContractValueDecoder.toLocalDate(d);
            if (ld != null) {
                sortedDates.add(ld);
            }
        }

        List<String> equipmentLines = new ArrayList<>();
        for (JsonNode eq : equipNode) {
            if (eq != null && eq.isTextual()) {
                equipmentLines.add(eq.asText());
            }
        }

        Map<String, List<TimelineEvent>> machineToEvents = buildMachineToEventsMap(events);
        List<String> machineHeaderLabels = equipmentScheduleHeaderLabels(equipmentLines);

        List<LocalTime> slotStarts = computeSlotTimes(events);
        List<String> columns = new ArrayList<>();
        columns.add(COL_DATE);
        columns.add(COL_MACH);
        columns.add(COL_PROC);
        columns.add(COL_TASK);
        for (LocalTime t : slotStarts) {
            columns.add(formatSlotColumn(t));
        }

        List<Map<String, String>> rows = new ArrayList<>();
        List<List<String>> badgeSlotRows = new ArrayList<>();

        List<LocalDate> graphicDays = graphicCalendarDates(events, sortedDates);
        for (LocalDate day : graphicDays) {
            Map<String, String> section = new LinkedHashMap<>();
            for (String col : columns) {
                section.put(col, col.equals(COL_DATE) ? formatSectionBanner(day) : "");
            }
            rows.add(section);
            badgeSlotRows.add(emptyBadgeSlots(slotStarts));

            for (int eqIdx = 0; eqIdx < equipmentLines.size(); eqIdx++) {
                String equipLine = equipmentLines.get(eqIdx);
                String[] split = splitEquipmentLine(equipLine);
                String proc = split[0];
                String machDisplay = machineHeaderLabels.get(eqIdx);
                Map<String, String> row = new LinkedHashMap<>();
                row.put(COL_DATE, "");
                row.put(COL_MACH, machDisplay);
                row.put(COL_PROC, proc);
                row.put(COL_TASK, "—");

                List<TimelineEvent> columnEvents =
                        new ArrayList<>(eventsForEquipmentColumn(machineToEvents, equipLine));
                columnEvents.sort(
                        Comparator.comparing(
                                        (TimelineEvent e) ->
                                                e.start != null
                                                        ? e.start
                                                        : LocalDateTime.MIN)
                                .thenComparing(
                                        e -> e.taskId != null ? e.taskId : ""));

                List<String> badgeSlots = new ArrayList<>();
                for (LocalTime slotStart : slotStarts) {
                    String col = formatSlotColumn(slotStart);
                    LocalDateTime winStart = LocalDateTime.of(day, slotStart);
                    LocalDateTime winEnd = winStart.plusMinutes(SLOT_MINUTES);
                    String cell = "";
                    String badgeCell = "";
                    for (TimelineEvent ev : columnEvents) {
                        if (!eventTouchesCalendarDay(ev, day)) {
                            continue;
                        }
                        SlotOverlapRange overlap = slotOverlapRangeForDisplay(ev);
                        if (!rangesOverlap(
                                overlap.start(), overlap.end(), winStart, winEnd)) {
                            continue;
                        }
                        if (ev.isInBreaks(winStart, winEnd)) {
                            continue;
                        }
                        cell = ev.timelineCellLabel();
                        badgeCell = ev.badgeSlotFragment();
                        break;
                    }
                    row.put(col, cell);
                    badgeSlots.add(badgeCell);
                }
                rows.add(row);
                badgeSlotRows.add(badgeSlots);
            }
        }

        return new EquipmentGanttSheetBundle(
                new JsonTableIo.SheetTable(columns, rows), badgeSlotRows);
    }

    /**
     * 同一日・同一機械・同一依頼の加工イベントを時系列で並べ、直前イベント終了より後に開始がある場合は
     * 「休憩で途切れた」別ブロックとして分割する。各ブロックの延長は {@code Σ units_done×unit_m}。
     */
    private static List<TimelineEvent> applyGapAwareMachiningLabels(List<TimelineEvent> raw) {
        Map<Integer, GapSegMeta> metaByIndex = new HashMap<>();
        Map<String, List<Integer>> group = new LinkedHashMap<>();
        for (int i = 0; i < raw.size(); i++) {
            TimelineEvent e = raw.get(i);
            if (!TimelineEvent.isMachiningDispatch(e)) {
                continue;
            }
            String k = TimelineEvent.gapGroupKey(e);
            group.computeIfAbsent(k, kk -> new ArrayList<>()).add(i);
        }
        for (Map.Entry<String, List<Integer>> en : group.entrySet()) {
            List<Integer> ix = en.getValue();
            ix.sort(Comparator.comparing(i -> raw.get(i).start));
            List<List<Integer>> segments = new ArrayList<>();
            List<Integer> cur = new ArrayList<>();
            LocalDateTime prevEnd = null;
            for (int ii : ix) {
                TimelineEvent ev = raw.get(ii);
                if (prevEnd != null && ev.start.isAfter(prevEnd)) {
                    segments.add(cur);
                    cur = new ArrayList<>();
                }
                cur.add(ii);
                prevEnd = ev.end;
            }
            if (!cur.isEmpty()) {
                segments.add(cur);
            }
            int segCount = segments.size();
            for (int si = 0; si < segments.size(); si++) {
                double sum = 0.0;
                for (int ii : segments.get(si)) {
                    TimelineEvent ev = raw.get(ii);
                    double ud = ev.unitsDone != null ? ev.unitsDone : 0.0;
                    double um = ev.unitM != null ? ev.unitM : 0.0;
                    sum += ud * um;
                }
                for (int ii : segments.get(si)) {
                    metaByIndex.put(ii, new GapSegMeta(si, segCount, sum));
                }
            }
        }
        List<TimelineEvent> out = new ArrayList<>(raw.size());
        for (int i = 0; i < raw.size(); i++) {
            TimelineEvent e = raw.get(i);
            GapSegMeta g = metaByIndex.get(i);
            if (g == null) {
                out.add(e);
            } else {
                out.add(e.withGapSegment(g.segmentIndex, g.segmentCount, g.segmentSumM));
            }
        }
        return out;
    }

    private record GapSegMeta(int segmentIndex, int segmentCount, double segmentSumM) {}

    /** 休憩で途切れた複数ブロックの見出し（3 ブロック目以降は「区間N」）。 */
    private static String segmentPhaseLabel(int idx, int total) {
        if (total <= 1) {
            return "";
        }
        if (idx == 0) {
            return "休憩前";
        }
        if (idx == 1) {
            return "休憩後";
        }
        return "区間" + (idx + 1);
    }

    private static List<String> emptyBadgeSlots(List<LocalTime> slotStarts) {
        List<String> z = new ArrayList<>();
        for (int i = 0; i < slotStarts.size(); i++) {
            z.add("");
        }
        return z;
    }

    /**
     * Python {@code _normalize_equipment_match_key} と同等（NFKC・NBSP/全角空白・ゼロ幅・連続空白）。
     */
    static String normalizeEquipmentMatchKey(String val) {
        if (val == null) {
            return "";
        }
        String t = Normalizer.normalize(val, Normalizer.Form.NFKC);
        t = t.replace('\u00a0', ' ').replace('\u3000', ' ');
        t = t.replaceAll("[\u200b\u200c\u200d\ufeff]", "");
        t = t.replaceAll("\\s+", " ").strip();
        return t;
    }

    /** timeline_events を ev.machine 文字列（JSON 上の生キー）で束ねる。挿入順を保つ。 */
    private static Map<String, List<TimelineEvent>> buildMachineToEventsMap(
            List<TimelineEvent> events) {
        Map<String, List<TimelineEvent>> map = new LinkedHashMap<>();
        for (TimelineEvent ev : events) {
            String mk = ev.machine != null ? ev.machine : "";
            map.computeIfAbsent(mk, k -> new ArrayList<>()).add(ev);
        }
        return map;
    }

    /**
     * Python {@code _eq_grid_events_for_equipment_column} と同じ解決順で、設備列に紐づくイベント一覧を返す。
     */
    static List<TimelineEvent> eventsForEquipmentColumn(
            Map<String, List<TimelineEvent>> machineToEvents, String eqCol) {
        if (eqCol == null || eqCol.isEmpty() || machineToEvents.isEmpty()) {
            return List.of();
        }
        List<TimelineEvent> evs = machineToEvents.get(eqCol);
        if (evs != null && !evs.isEmpty()) {
            return evs;
        }
        String nk = normalizeEquipmentMatchKey(eqCol);
        if (nk.isEmpty()) {
            return List.of();
        }
        for (Map.Entry<String, List<TimelineEvent>> e : machineToEvents.entrySet()) {
            if (normalizeEquipmentMatchKey(e.getKey()).equals(nk)) {
                return e.getValue();
            }
        }
        String[] pmEq = splitEquipmentLine(eqCol);
        String peN = normalizeEquipmentMatchKey(pmEq[0]);
        String meN = normalizeEquipmentMatchKey(pmEq[1]);
        if (!peN.isEmpty() && !meN.isEmpty()) {
            for (Map.Entry<String, List<TimelineEvent>> e : machineToEvents.entrySet()) {
                String[] pmMk = splitEquipmentLine(e.getKey());
                String pk = normalizeEquipmentMatchKey(pmMk[0]);
                String mkM = normalizeEquipmentMatchKey(pmMk[1]);
                if (peN.equals(pk) && meN.equals(mkM)) {
                    return e.getValue();
                }
            }
        }
        return List.of();
    }

    /**
     * Python {@code _equipment_schedule_header_labels} と同様。
     * 機械名が複数行で重なるときだけ「機械名（工程名）」で区別して Excel と視認性を揃える。
     */
    static List<String> equipmentScheduleHeaderLabels(List<String> equipmentList) {
        List<String> raw = new ArrayList<>(equipmentList.size());
        for (String eq : equipmentList) {
            String s = eq != null ? eq.strip() : "";
            if (s.contains("+")) {
                String mpart = s.split("\\+", 2)[1].strip();
                raw.add(!mpart.isEmpty() ? mpart : s);
            } else {
                raw.add(s);
            }
        }
        Map<String, Integer> counts = new LinkedHashMap<>();
        for (String r : raw) {
            counts.put(r, counts.getOrDefault(r, 0) + 1);
        }
        List<String> out = new ArrayList<>(equipmentList.size());
        for (int i = 0; i < equipmentList.size(); i++) {
            String eq = equipmentList.get(i);
            String r = raw.get(i);
            if (counts.getOrDefault(r, 0) > 1) {
                String s = eq != null ? eq.strip() : "";
                if (s.contains("+")) {
                    String p = s.split("\\+", 2)[0].strip();
                    out.add((!p.isEmpty()) ? (r + "（" + p + "）") : r);
                } else {
                    out.add(r);
                }
            } else {
                out.add(r);
            }
        }
        return out;
    }

    /**
     * タイムラインにイベントが存在する暦日のみ（イベントの {@code date} および {@code start_dt} の日付）。
     * イベントが 1 件も無いときだけフォールバックで {@code sorted_dates} をそのまま使う。
     */
    private static List<LocalDate> graphicCalendarDates(
            List<TimelineEvent> events, List<LocalDate> sortedDatesFallback) {
        TreeSet<LocalDate> days = new TreeSet<>();
        for (TimelineEvent ev : events) {
            if (ev.date != null) {
                days.add(ev.date);
            }
            if (ev.start != null) {
                days.add(ev.start.toLocalDate());
            }
        }
        if (!days.isEmpty()) {
            return new ArrayList<>(days);
        }
        return new ArrayList<>(sortedDatesFallback);
    }

    /** 契約の {@code date} または開始／終了時刻が、その暦日と関係するイベントを対象にする。 */
    private static boolean eventTouchesCalendarDay(TimelineEvent ev, LocalDate day) {
        if (ev.date != null && ev.date.equals(day)) {
            return true;
        }
        if (ev.start != null && ev.start.toLocalDate().equals(day)) {
            return true;
        }
        if (ev.end != null && ev.end.toLocalDate().equals(day)) {
            return true;
        }
        return false;
    }

    /** "工程+機械…" を最初の '+' で分割（無ければ全体を機械名）。 */
    static String[] splitEquipmentLine(String line) {
        if (line == null || line.isEmpty()) {
            return new String[] {"", ""};
        }
        int p = line.indexOf('+');
        if (p < 0) {
            return new String[] {"", line};
        }
        return new String[] {line.substring(0, p).strip(), line.substring(p + 1).strip()};
    }

    static String formatSectionBanner(LocalDate day) {
        return "【"
                + day.getYear()
                + "/"
                + String.format("%02d", day.getMonthValue())
                + "/"
                + String.format("%02d", day.getDayOfMonth())
                + "】";
    }

    static String formatSlotColumn(LocalTime t) {
        return t.getHour() + ":" + String.format("%02d", t.getMinute());
    }

    static boolean rangesOverlap(
            LocalDateTime a0, LocalDateTime a1, LocalDateTime b0, LocalDateTime b1) {
        return a0.isBefore(b1) && a1.isAfter(b0);
    }

    /**
     * ガントのスロット充填用区間。加工が {@link #SLOT_MINUTES} 未満のときは開始時刻を含む1スロット枠
     * （床時刻〜+{@link #SLOT_MINUTES} 分）に拡張する。配台の実際の {@code end} は変えない。
     */
    record SlotOverlapRange(LocalDateTime start, LocalDateTime end) {}

    static SlotOverlapRange slotOverlapRangeForDisplay(TimelineEvent ev) {
        if (ev == null || ev.start == null || ev.end == null) {
            return new SlotOverlapRange(
                    ev != null ? ev.start : null, ev != null ? ev.end : null);
        }
        if (!TimelineEvent.isMachiningDispatch(ev)) {
            return new SlotOverlapRange(ev.start, ev.end);
        }
        long durationSec = Duration.between(ev.start, ev.end).getSeconds();
        if (durationSec >= SLOT_MINUTES * 60L) {
            return new SlotOverlapRange(ev.start, ev.end);
        }
        LocalTime slotFloor = floorToSlot(ev.start.toLocalTime());
        LocalDateTime dispStart = LocalDateTime.of(ev.start.toLocalDate(), slotFloor);
        LocalDateTime dispEnd = dispStart.plusMinutes(SLOT_MINUTES);
        return new SlotOverlapRange(dispStart, dispEnd);
    }

    /**
     * イベントの発生時刻を包含するよう 10 分刻みスロット開始時刻を列挙。
     * 無イベント時は 8:00〜21:00（21:00 排他的終端）。
     */
    static List<LocalTime> computeSlotTimes(List<TimelineEvent> events) {
        LocalTime start = FALLBACK_DAY_START;
        LocalTime endCap = FALLBACK_DAY_END;
        for (TimelineEvent ev : events) {
            LocalTime fs = floorToSlot(ev.start.toLocalTime());
            LocalTime ce = ceilToSlot(ev.end.toLocalTime());
            if (fs.isBefore(start)) {
                start = fs;
            }
            if (ce.isAfter(endCap)) {
                endCap = ce;
            }
        }
        List<LocalTime> out = new ArrayList<>();
        LocalTime t = start;
        while (t.isBefore(endCap)) {
            out.add(t);
            t = t.plusMinutes(SLOT_MINUTES);
            if (out.size() > 5000) {
                break;
            }
        }
        return out;
    }

    static LocalTime floorToSlot(LocalTime t) {
        int m = t.getHour() * 60 + t.getMinute();
        m = (m / SLOT_MINUTES) * SLOT_MINUTES;
        return LocalTime.of(m / 60, m % 60);
    }

    static LocalTime ceilToSlot(LocalTime t) {
        int m = t.getHour() * 60 + t.getMinute();
        int rem = m % SLOT_MINUTES;
        if (rem != 0) {
            m += SLOT_MINUTES - rem;
        }
        if (m > LAST_SLOT_START_MINUTE_OF_DAY) {
            m = LAST_SLOT_START_MINUTE_OF_DAY;
        }
        return LocalTime.of(m / 60, m % 60);
    }

    static final class TimelineEvent {
        final LocalDate date;
        final String machine;
        final String taskId;
        final String eventKind;
        final LocalDateTime start;
        final LocalDateTime end;
        /** メートル／単位（ロール単位長さなど）。Python {@code unit_m}。 */
        final Double unitM;
        /** 当該イベントの単位数（ロール数など）。Python {@code units_done}。 */
        final Double unitsDone;
        /**
         * 当該タスクの配台単位総数（例: 全ロール数）。Python {@code total_units}。
         * ガントの「依頼単位の延長m」は {@code total_units×unit_m} を優先し、配台表と揃える。
         */
        final Double totalUnits;
        /** 実績明細由来の加工長さ(m)。Python {@code label_len_m}。 */
        final Double labelLenM;
        /** Python {@code label_len_m_is_cumulative}。累積ラベルはスロット按分しない。 */
        final boolean labelLenMIsCumulative;
        /** メイン担当（JSON {@code op}）。 */
        final String op;
        /** サブ担当など（JSON {@code sub}）。 */
        final String sub;
        final List<List<LocalDateTime>> breaks;
        /**
         * 時間ギャップで分割した加工ブロックの 0 始まりインデックス。{@code -1} は未使用（単一ブロックまたは非加工）。
         */
        final int gapSegmentIndex;
        /** {@link #gapSegmentIndex} を付けたときのブロック数（2 以上で休憩前／休憩後表記）。 */
        final int gapSegmentCount;
        /** 同一ギャップブロック内の {@code Σ units_done×unit_m}（m）。 */
        final double gapSegmentSumM;

        TimelineEvent(
                LocalDate date,
                String machine,
                String taskId,
                String eventKind,
                LocalDateTime start,
                LocalDateTime end,
                Double unitM,
                Double unitsDone,
                Double totalUnits,
                Double labelLenM,
                boolean labelLenMIsCumulative,
                String op,
                String sub,
                List<List<LocalDateTime>> breaks,
                int gapSegmentIndex,
                int gapSegmentCount,
                double gapSegmentSumM) {
            this.date = date;
            this.machine = machine;
            this.taskId = taskId;
            this.eventKind = eventKind;
            this.start = start;
            this.end = end;
            this.unitM = unitM;
            this.unitsDone = unitsDone;
            this.totalUnits = totalUnits;
            this.labelLenM = labelLenM;
            this.labelLenMIsCumulative = labelLenMIsCumulative;
            this.op = op != null ? op : "";
            this.sub = sub != null ? sub : "";
            this.breaks = breaks;
            this.gapSegmentIndex = gapSegmentIndex;
            this.gapSegmentCount = gapSegmentCount;
            this.gapSegmentSumM = gapSegmentSumM;
        }

        TimelineEvent withGapSegment(int segmentIndex, int segmentCount, double segmentSumM) {
            return new TimelineEvent(
                    date,
                    machine,
                    taskId,
                    eventKind,
                    start,
                    end,
                    unitM,
                    unitsDone,
                    totalUnits,
                    labelLenM,
                    labelLenMIsCumulative,
                    op,
                    sub,
                    breaks,
                    segmentIndex,
                    segmentCount,
                    segmentSumM);
        }

        static boolean isMachiningDispatch(TimelineEvent e) {
            return e != null
                    && "machining".equals(e.eventKind)
                    && e.taskId != null
                    && !e.taskId.isBlank();
        }

        static String gapGroupKey(TimelineEvent e) {
            return String.valueOf(e.date)
                    + "|"
                    + (e.machine != null ? e.machine : "")
                    + "|"
                    + (e.taskId != null ? e.taskId : "");
        }

        static TimelineEvent from(JsonNode n) {
            if (n == null || !n.isObject()) {
                return null;
            }
            Object d = GanttContractValueDecoder.decodeValue(n.get("date"));
            LocalDate date = GanttContractValueDecoder.toLocalDate(d);
            Object sdt = GanttContractValueDecoder.decodeValue(n.get("start_dt"));
            Object edt = GanttContractValueDecoder.decodeValue(n.get("end_dt"));
            LocalDateTime start = GanttContractValueDecoder.toLocalDateTime(sdt);
            LocalDateTime end = GanttContractValueDecoder.toLocalDateTime(edt);
            if (date == null || start == null || end == null) {
                return null;
            }
            String machine = text(n, "machine");
            String taskId = text(n, "task_id");
            String eventKind = text(n, "event_kind");
            String op = text(n, "op");
            String sub = text(n, "sub");
            Double unitM = numberOrNull(n, "unit_m");
            Double unitsDone = numberOrNull(n, "units_done");
            Double totalUnits = numberOrNull(n, "total_units");
            Double labelLenM = numberOrNull(n, "label_len_m");
            boolean labelCumulative =
                    n.has("label_len_m_is_cumulative") && n.get("label_len_m_is_cumulative").asBoolean();
            List<List<LocalDateTime>> breaks = new ArrayList<>();
            JsonNode bn = n.get("breaks");
            if (bn != null && bn.isArray()) {
                for (JsonNode b : bn) {
                    Object tup = GanttContractValueDecoder.decodeValue(b);
                    if (tup instanceof List<?> list && list.size() >= 2) {
                        LocalDateTime b0 =
                                GanttContractValueDecoder.toLocalDateTime(list.get(0));
                        LocalDateTime b1 =
                                GanttContractValueDecoder.toLocalDateTime(list.get(1));
                        if (b0 != null && b1 != null) {
                            List<LocalDateTime> pair = new ArrayList<>(2);
                            pair.add(b0);
                            pair.add(b1);
                            breaks.add(pair);
                        }
                    }
                }
            }
            return new TimelineEvent(
                    date,
                    machine,
                    taskId,
                    eventKind,
                    start,
                    end,
                    unitM,
                    unitsDone,
                    totalUnits,
                    labelLenM,
                    labelCumulative,
                    op,
                    sub,
                    breaks,
                    -1,
                    0,
                    Double.NaN);
        }

        static String text(JsonNode n, String field) {
            JsonNode x = n.get(field);
            return x != null && x.isTextual() ? x.asText() : "";
        }

        static Double numberOrNull(JsonNode n, String field) {
            JsonNode x = n.get(field);
            if (x != null && x.isNumber()) {
                return x.doubleValue();
            }
            return null;
        }

        boolean isInBreaks(LocalDateTime winStart, LocalDateTime winEnd) {
            for (List<LocalDateTime> br : breaks) {
                if (br.size() < 2) {
                    continue;
                }
                LocalDateTime b0 = br.get(0);
                LocalDateTime b1 = br.get(1);
                if (rangesOverlap(b0, b1, winStart, winEnd)) {
                    return true;
                }
            }
            return false;
        }

        /**
         * Python {@code _gantt_format_length_m} 相当（整数に近いときは整数文字列、それ以外は小数1桁まで）。
         */
        static String formatLengthM(double v) {
            if (Double.isNaN(v) || v <= 1e-12) {
                return "";
            }
            double r = Math.round(v);
            if (Math.abs(v - r) <= 1e-9) {
                return Long.toString((long) r);
            }
            String s = String.format(java.util.Locale.ROOT, "%.1f", v);
            s = s.replaceAll("0+$", "").replaceAll("\\.$", "");
            return s;
        }

        /**
         * イベントに表示する総加工長さ(m)。
         * {@code label_len_m} があれば優先。無ければ契約の {@code total_units×unit_m}（配台表の依頼総量と整合）、
         * それも無ければ {@code units_done×unit_m}（当該タイムライン分割の1片）。
         */
        Double eventTotalLengthMeters() {
            if (labelLenM != null && !labelLenMIsCumulative) {
                return labelLenM;
            }
            if (labelLenMIsCumulative) {
                return null;
            }
            double um = unitM != null ? unitM : 0.0;
            if (totalUnits != null && totalUnits > 1e-12 && um > 1e-12) {
                return totalUnits * um;
            }
            double u = unitsDone != null ? unitsDone : 0.0;
            if (u > 1e-12 && um > 1e-12) {
                return u * um;
            }
            return null;
        }

        /**
         * タイムライン1マスに表示する文字列（依頼NO＋契約に基づく総加工量m）。
         * 加工が時間ギャップで複数ブロックに分かれるときは「休憩前／休憩後」とブロック合計m。
         */
        String timelineCellLabel() {
            if ("machine_daily_startup".equals(eventKind)) {
                return "日次始業準備";
            }
            if ("request_switch_prep".equals(eventKind)) {
                return "依頼切替準備";
            }
            if ("break_resume_prep".equals(eventKind)) {
                return "休憩再開準備";
            }
            if ("machine_daily_inspection".equals(eventKind)
                    || "daily_inspection".equals(eventKind)) {
                return "日次点検";
            }
            String tid = taskId != null ? taskId.strip() : "";
            if (labelLenM != null && !labelLenMIsCumulative) {
                String len = formatLengthM(labelLenM);
                if (!len.isEmpty()) {
                    return tid.isEmpty() ? len + "m" : tid + " " + len + "m";
                }
            }
            // 加工はギャップ集計の Σ(units_done×unit_m) を優先（複数日の続きは当日ブロック分のみとなり total_units× を上回らない）。
            if ("machining".equals(eventKind)
                    && gapSegmentIndex >= 0
                    && !Double.isNaN(gapSegmentSumM)
                    && gapSegmentSumM > 1e-12) {
                String len = formatLengthM(gapSegmentSumM);
                if (!len.isEmpty()) {
                    if (gapSegmentCount > 1) {
                        String phase =
                                EquipmentGanttContractSheetTableBuilder.segmentPhaseLabel(
                                        gapSegmentIndex, gapSegmentCount);
                        if (!phase.isEmpty()) {
                            return tid.isEmpty()
                                    ? phase + " " + len + "m"
                                    : tid + " " + phase + " " + len + "m";
                        }
                    }
                    return tid.isEmpty() ? len + "m" : tid + " " + len + "m";
                }
            }
            Double totalM = eventTotalLengthMeters();
            if (totalM == null
                    && labelLenMIsCumulative
                    && labelLenM != null
                    && labelLenM > 1e-12) {
                totalM = labelLenM;
            }
            if (totalM != null && totalM > 1e-12) {
                String len = formatLengthM(totalM);
                if (!len.isEmpty()) {
                    return tid.isEmpty() ? len + "m" : tid + " " + len + "m";
                }
            }
            if (!tid.isEmpty()) {
                return tid;
            }
            return eventKind != null && !eventKind.isEmpty() ? eventKind : "";
        }

        /**
         * タイムスロット1マスに書き込むバッジセル（複数人は {@link PersonNameBadgeText#UNIT_SEPARATOR} 連結）。
         */
        String badgeSlotFragment() {
            boolean startupSplit =
                    "machine_daily_startup".equals(eventKind)
                            || "request_switch_prep".equals(eventKind)
                            || "break_resume_prep".equals(eventKind)
                            || "machine_daily_inspection".equals(eventKind)
                            || "daily_inspection".equals(eventKind);
            return PersonNameBadgeText.joinBadgeCells(
                    PersonNameBadgeText.badgeListFromOpSub(op, sub, startupSplit));
        }
    }
}
