package jp.co.pm.ai.desktop.reconciliation;

import java.io.BufferedReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Stream;

/**
 * 湖南工場「加工日報発行問合せ」CSV から (依頼NO, 工程名, 機械名) ごとの {@code 完了区分} を引く。
 *
 * <p>原本転記・計画確認タブ専用。配台・タスク入力には使わない。
 */
public final class KonanDailyReportLookup {

    public static final String KEY_DAILY_REPORT_SOURCE_DIR = "PM_AI_DAILY_REPORT_SOURCE_DIR";
    public static final String KEY_DAILY_REPORT_CSV_PATH = "PM_AI_DAILY_REPORT_CSV_PATH";
    public static final String KEY_DAILY_REPORT_LOOKUP = "PM_AI_DAILY_REPORT_LOOKUP";

    private static final int META_SKIP_LINES = 3;
    private static final String FILENAME_PREFIX = "加工日報発行問合せ_";

    private static final String COL_TASK_ID = "依頼NO";
    private static final String COL_PROCESS = "工程名";
    private static final String COL_MACHINE = "機械名";
    private static final String COL_DAY = "加工日付";
    private static final String COL_COMPLETION = "完了区分";

    private static final String DEFAULT_SOURCE_DIR =
            "\\\\192.168.0.101\\"
                    + "共有フォルダ\\"
                    + "湖南工場\\"
                    + "湖南共有\\"
                    + "生産管理システム\\"
                    + "管理システム\\"
                    + "●DATA\\"
                    + "加工日報";

    private final Map<TaskKey, String> completionByKey;
    private final Map<TaskKey2, List<TaskKey>> keysByTaskProcess;
    /** 読込に成功した CSV の絶対パス（正規化済み）。未読込時は空。 */
    private final String sourcePath;

    private KonanDailyReportLookup(
            Map<TaskKey, String> completionByKey,
            Map<TaskKey2, List<TaskKey>> keysByTaskProcess,
            String sourcePath) {
        this.completionByKey = Map.copyOf(completionByKey);
        this.keysByTaskProcess = Map.copyOf(keysByTaskProcess);
        this.sourcePath = sourcePath != null ? sourcePath : "";
    }

    public static KonanDailyReportLookup empty() {
        return new KonanDailyReportLookup(Map.of(), Map.of(), "");
    }

    public boolean isLoaded() {
        return !completionByKey.isEmpty();
    }

    /** 読込元 CSV のフルパス。ファイル未解決・未読込時は空。 */
    public String sourcePath() {
        return sourcePath;
    }

    /** {@link #sourcePath()} のファイル名のみ（ログ等向け）。 */
    public String sourceFileName() {
        if (sourcePath.isEmpty()) {
            return "";
        }
        Path p = Path.of(sourcePath);
        Path name = p.getFileName();
        return name != null ? name.toString() : sourcePath;
    }

    /**
     * 加工日報の完了区分を UI 向け短ラベルで返す。
     *
     * @return {@code 完了} / {@code 未完} / 空（日報に該当なし、または日報未読込）
     */
    public String completionDisplay(String iraiNo, String processName, String machineName) {
        String raw = rawCompletion(iraiNo, processName, machineName);
        return formatCompletionDisplay(raw);
    }

    static String formatCompletionDisplay(String raw) {
        if (raw == null || raw.isBlank()) {
            return "";
        }
        if (raw.contains("未完")) {
            return "未了";
        }
        if (raw.contains("完了")) {
            return "完了";
        }
        return raw.strip();
    }

    /**
     * 依頼NO 単位の加工日報ステータス。
     *
     * <p>工程行のいずれかが未完なら {@code 未了}、すべて完了なら {@code 完了}、日報に行が無ければ {@code ―}。
     */
    public String orderCompletionStatus(String iraiNo) {
        String tid = normalizePart(iraiNo);
        if (tid.isEmpty() || completionByKey.isEmpty()) {
            return "―";
        }
        boolean anyRow = false;
        boolean anyIncomplete = false;
        for (Map.Entry<TaskKey, String> e : completionByKey.entrySet()) {
            if (!tid.equalsIgnoreCase(e.getKey().taskId())) {
                continue;
            }
            anyRow = true;
            String raw = e.getValue() != null ? e.getValue() : "";
            if (raw.contains("未完")) {
                anyIncomplete = true;
                break;
            }
        }
        if (!anyRow) {
            return "―";
        }
        if (anyIncomplete) {
            return "未了";
        }
        return "完了";
    }

    /** 依頼NO に紐づく加工日報行（工程×機械）。 */
    public List<OrderDailyReportEntry> entriesForOrder(String iraiNo) {
        String tid = normalizePart(iraiNo);
        if (tid.isEmpty()) {
            return List.of();
        }
        List<OrderDailyReportEntry> out = new ArrayList<>();
        for (Map.Entry<TaskKey, String> e : completionByKey.entrySet()) {
            if (!tid.equalsIgnoreCase(e.getKey().taskId())) {
                continue;
            }
            TaskKey key = e.getKey();
            out.add(
                    new OrderDailyReportEntry(
                            key.process(),
                            key.machine(),
                            formatCompletionDisplay(e.getValue())));
        }
        out.sort(
                Comparator.comparing(OrderDailyReportEntry::processName)
                        .thenComparing(OrderDailyReportEntry::machineName));
        return List.copyOf(out);
    }

    public record OrderDailyReportEntry(
            String processName, String machineName, String completionStatus) {}

    private String rawCompletion(String iraiNo, String processName, String machineName) {
        if (completionByKey.isEmpty()) {
            return "";
        }
        String tid = normalizePart(iraiNo);
        String proc = normalizePart(processName);
        String mach = normalizePart(machineName);
        if (tid.isEmpty() || proc.isEmpty()) {
            return "";
        }
        TaskKey k3 = new TaskKey(tid, proc, mach);
        String hit = completionByKey.get(k3);
        if (hit != null) {
            return hit;
        }
        List<TaskKey> cands = keysByTaskProcess.get(new TaskKey2(tid, proc));
        if (cands != null && cands.size() == 1) {
            return completionByKey.getOrDefault(cands.getFirst(), "");
        }
        return "";
    }

    /** 加工日報フォルダ／単一 CSV 指定から最新 CSV の絶対パスを返す（読込失敗時は空）。 */
    public static Optional<Path> resolveNewestCsvPath(Map<String, String> ui) {
        try {
            Path path = resolveCsvPath(ui);
            return path != null ? Optional.of(path) : Optional.empty();
        } catch (IOException ex) {
            return Optional.empty();
        }
    }

    public static KonanDailyReportLookup load(Map<String, String> ui, List<String> warnings) {
        if (!lookupEnabled(ui)) {
            return empty();
        }
        try {
            Path path = resolveCsvPath(ui);
            if (path == null) {
                if (warnings != null) {
                    warnings.add(
                            "加工日報 CSV が見つかりません（PM_AI_DAILY_REPORT_CSV_PATH / 加工日報フォルダ）。");
                }
                return empty();
            }
            return loadFromPath(path);
        } catch (Exception ex) {
            if (warnings != null) {
                warnings.add(
                        "加工日報 CSV の読込に失敗: "
                                + (ex.getMessage() != null ? ex.getMessage() : ex.toString()));
            }
            return empty();
        }
    }

    static KonanDailyReportLookup loadFromPath(Path path) throws IOException {
        String pathText = path.toAbsolutePath().normalize().toString();
        List<String> lines = readAllLines(path);
        if (lines.size() <= META_SKIP_LINES) {
            return new KonanDailyReportLookup(Map.of(), Map.of(), pathText);
        }
        List<String> headers = parseCsvLine(lines.get(META_SKIP_LINES));
        stripBomFromFirstHeader(headers);
        Map<String, Integer> colIndex = indexColumns(headers);
        int iTask = requireColumn(colIndex, COL_TASK_ID);
        int iProc = requireColumn(colIndex, COL_PROCESS);
        int iMach = requireColumn(colIndex, COL_MACHINE);
        int iDay = requireColumn(colIndex, COL_DAY);
        int iComp = requireColumn(colIndex, COL_COMPLETION);

        Map<TaskKey, RowAgg> latest = new HashMap<>();
        for (int i = META_SKIP_LINES + 1; i < lines.size(); i++) {
            List<String> cells = parseCsvLine(lines.get(i));
            if (cells.isEmpty()) {
                continue;
            }
            String tid = cellAt(cells, iTask);
            String proc = cellAt(cells, iProc);
            String mach = cellAt(cells, iMach);
            if (normalizePart(tid).isEmpty() || normalizePart(proc).isEmpty()) {
                continue;
            }
            TaskKey key =
                    new TaskKey(normalizePart(tid), normalizePart(proc), normalizePart(mach));
            String day = cellAt(cells, iDay);
            String comp = cellAt(cells, iComp);
            RowAgg prev = latest.get(key);
            if (prev == null || compareDay(day, prev.day()) > 0) {
                latest.put(key, new RowAgg(day, comp));
            }
        }

        Map<TaskKey, String> completion = new HashMap<>();
        Map<TaskKey2, List<TaskKey>> byTaskProc = new HashMap<>();
        for (Map.Entry<TaskKey, RowAgg> e : latest.entrySet()) {
            completion.put(e.getKey(), e.getValue().completion());
            TaskKey2 k2 = new TaskKey2(e.getKey().taskId(), e.getKey().process());
            byTaskProc.computeIfAbsent(k2, k -> new ArrayList<>()).add(e.getKey());
        }
        for (List<TaskKey> list : byTaskProc.values()) {
            list.sort(Comparator.comparing(TaskKey::machine));
        }
        return new KonanDailyReportLookup(completion, byTaskProc, pathText);
    }

    private static boolean lookupEnabled(Map<String, String> ui) {
        String raw = ui != null ? ui.getOrDefault(KEY_DAILY_REPORT_LOOKUP, "").strip() : "";
        if (raw.isEmpty()) {
            return true;
        }
        String lower = raw.toLowerCase(Locale.ROOT);
        return !("0".equals(lower) || "false".equals(lower) || "off".equals(lower) || "no".equals(lower));
    }

    private static Path resolveCsvPath(Map<String, String> ui) throws IOException {
        String explicit =
                ui != null ? ui.getOrDefault(KEY_DAILY_REPORT_CSV_PATH, "").strip() : "";
        if (!explicit.isEmpty()) {
            Path p = Path.of(explicit);
            if (Files.isRegularFile(p)) {
                return p.toAbsolutePath().normalize();
            }
            return null;
        }
        String dir = ui != null ? ui.getOrDefault(KEY_DAILY_REPORT_SOURCE_DIR, "").strip() : "";
        if (dir.isEmpty()) {
            dir = DEFAULT_SOURCE_DIR;
        }
        Path dirPath = Path.of(dir);
        if (!Files.isDirectory(dirPath)) {
            return null;
        }
        Path best = null;
        long bestKey = Long.MIN_VALUE;
        try (Stream<Path> stream = Files.list(dirPath)) {
            for (Path p : stream.toList()) {
                if (!Files.isRegularFile(p)) {
                    continue;
                }
                String name = p.getFileName().toString();
                if (!name.startsWith(FILENAME_PREFIX) || !name.toLowerCase(Locale.ROOT).endsWith(".csv")) {
                    continue;
                }
                long t = Files.getLastModifiedTime(p).toMillis();
                if (t > bestKey) {
                    bestKey = t;
                    best = p;
                }
            }
        }
        return best != null ? best.toAbsolutePath().normalize() : null;
    }

    private static List<String> readAllLines(Path path) throws IOException {
        List<String> lines = new ArrayList<>();
        try (BufferedReader r = Files.newBufferedReader(path, StandardCharsets.UTF_8)) {
            String line;
            while ((line = r.readLine()) != null) {
                lines.add(line);
            }
        }
        return lines;
    }

    private static List<String> parseCsvLine(String line) {
        List<String> cells = new ArrayList<>();
        if (line.isEmpty()) {
            return cells;
        }
        StringBuilder cur = new StringBuilder();
        boolean inQ = false;
        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            if (inQ) {
                if (c == '"') {
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                        cur.append('"');
                        i++;
                    } else {
                        inQ = false;
                    }
                } else {
                    cur.append(c);
                }
            } else {
                if (c == '"') {
                    inQ = true;
                } else if (c == ',') {
                    cells.add(cur.toString());
                    cur.setLength(0);
                } else {
                    cur.append(c);
                }
            }
        }
        cells.add(cur.toString());
        return cells;
    }

    private static void stripBomFromFirstHeader(List<String> headers) {
        if (!headers.isEmpty()) {
            String h0 = headers.getFirst();
            if (h0 != null && !h0.isEmpty() && h0.charAt(0) == '\uFEFF') {
                headers.set(0, h0.substring(1));
            }
        }
    }

    private static Map<String, Integer> indexColumns(List<String> headers) {
        Map<String, Integer> idx = new HashMap<>();
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i) != null ? headers.get(i).strip() : "";
            if (!h.isEmpty()) {
                idx.putIfAbsent(h, i);
            }
        }
        return idx;
    }

    private static int requireColumn(Map<String, Integer> colIndex, String name) throws IOException {
        Integer i = colIndex.get(name);
        if (i == null) {
            throw new IOException("加工日報に列「" + name + "」がありません。");
        }
        return i;
    }

    private static String cellAt(List<String> cells, int index) {
        return index >= 0 && index < cells.size() && cells.get(index) != null
                ? cells.get(index).strip()
                : "";
    }

    private static int compareDay(String a, String b) {
        return normalizePart(a).compareTo(normalizePart(b));
    }

    private static String normalizePart(String val) {
        return JuchuTransferValueNormalizer.normalizeText(val);
    }

    private record TaskKey(String taskId, String process, String machine) {}

    private record TaskKey2(String taskId, String process) {}

    private record RowAgg(String day, String completion) {}
}
