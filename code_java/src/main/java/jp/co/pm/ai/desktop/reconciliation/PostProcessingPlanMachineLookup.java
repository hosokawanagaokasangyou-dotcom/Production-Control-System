package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.OptionalInt;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.TaskInputSourceRawGridIo;

/**
 * 加工計画DATA（{@code PM_AI_TASK_INPUT_SOURCE_DIR} 最新 / {@code PM_AI_PROCESSING_PLAN_PATH}）の
 * 「機械」「機械名」列から、後加工商品マスタ編集フォーム用のコード一覧を構築する。
 */
public final class PostProcessingPlanMachineLookup {

    /** 加工計画DATA の機械コード列（Aladdin 見出し）。 */
    public static final String COL_MACHINE_CODE = "機械";

    /** 旧エクスポート等の互換用。 */
    public static final String COL_MACHINE_CODE_LEGACY = "機械コード";

    public static final String COL_MACHINE_NAME = "機械名";

    private static final Pattern MACHINE_CODE_COLUMN = Pattern.compile("^機械コード(\\d+)$");

    private static volatile Snapshot cached;

    private PostProcessingPlanMachineLookup() {}

    public record Snapshot(
            Path sourcePath,
            long lastModified,
            boolean hasCodeColumn,
            boolean hasNameColumn,
            Map<String, String> machineCodeToName,
            List<String> machineComboLabels) {

        public static Snapshot empty() {
            return new Snapshot(Path.of(""), -1L, false, false, Map.of(), List.of());
        }

        public boolean loaded() {
            return !machineCodeToName.isEmpty();
        }
    }

    public static Snapshot snapshot(Map<String, String> ui) throws IOException {
        NetworkSourceDirResolver.Result net = NetworkSourceDirResolver.resolve(ui);
        Path planPath = net.taskInputPath().orElse(null);
        if (planPath == null || !Files.isRegularFile(planPath)) {
            return Snapshot.empty();
        }
        return snapshotFromFile(planPath);
    }

    public static Snapshot snapshotFromFile(Path planPath) throws IOException {
        if (!Files.isRegularFile(planPath)) {
            return Snapshot.empty();
        }
        Path abs = planPath.toAbsolutePath().normalize();
        long mtime = Files.getLastModifiedTime(abs).toMillis();
        Snapshot hit = cached;
        if (hit != null && Objects.equals(hit.sourcePath(), abs) && hit.lastModified() == mtime) {
            return hit;
        }
        synchronized (PostProcessingPlanMachineLookup.class) {
            hit = cached;
            if (hit != null && Objects.equals(hit.sourcePath(), abs) && hit.lastModified() == mtime) {
                return hit;
            }
            Snapshot loaded = load(abs, mtime);
            cached = loaded;
            return loaded;
        }
    }

    public static void invalidate() {
        cached = null;
    }

    public static boolean isMachineCodeColumn(String columnName) {
        if (columnName == null) {
            return false;
        }
        return MACHINE_CODE_COLUMN.matcher(columnName.trim()).matches();
    }

    public static OptionalInt machineStepIndex(String columnName) {
        if (columnName == null) {
            return OptionalInt.empty();
        }
        Matcher m = MACHINE_CODE_COLUMN.matcher(columnName.trim());
        if (!m.matches()) {
            return OptionalInt.empty();
        }
        try {
            return OptionalInt.of(Integer.parseInt(m.group(1)));
        } catch (NumberFormatException ex) {
            return OptionalInt.empty();
        }
    }

    public static String normalizeMachineCode(String raw) {
        return PostProcessingKouteiNaiyoMasterLookup.normalizeCode(raw, 0);
    }

    public static String resolveMachineName(Snapshot snap, String rawCode) {
        if (snap == null) {
            return "";
        }
        String code = normalizeMachineCode(rawCode);
        if (code.isEmpty()) {
            return "";
        }
        return snap.machineCodeToName().getOrDefault(code, "");
    }

    public static String resolveCodeFromComboInput(Snapshot snap, String text) {
        if (text == null) {
            return "";
        }
        String trimmed = text.trim();
        if (trimmed.isEmpty()) {
            return "";
        }
        String norm = normalizeMachineCode(trimmed);
        if (snap != null && snap.machineCodeToName().containsKey(norm)) {
            return norm;
        }
        if (snap != null) {
            for (String label : snap.machineComboLabels()) {
                if (label.equals(trimmed) || label.startsWith(norm + " ")) {
                    int sp = label.indexOf(' ');
                    return sp > 0 ? label.substring(0, sp) : norm;
                }
            }
        }
        return norm;
    }

    private static Snapshot load(Path abs, long mtime) throws IOException {
        PlanInputTabularIo.TabularSheet sheet = readProcessingPlanSheet(abs);
        List<String> headers = sheet.headers();
        int idxCode = indexOfMachineCodeHeader(headers);
        int idxName = indexOfHeader(headers, COL_MACHINE_NAME);
        boolean hasCode = idxCode >= 0;
        boolean hasName = idxName >= 0;
        if (!hasCode && !hasName) {
            return new Snapshot(abs, mtime, false, false, Map.of(), List.of());
        }

        Map<String, String> codeToName = new LinkedHashMap<>();
        for (List<String> row : sheet.rows()) {
            String code =
                    hasCode ? cellAt(row, idxCode) : "";
            String name =
                    hasName ? cellAt(row, idxName) : "";
            code = normalizeMachineCode(code);
            name = name != null ? name.trim() : "";
            if (code.isEmpty() && name.isEmpty()) {
                continue;
            }
            if (code.isEmpty()) {
                code = name;
            }
            if (name.isEmpty()) {
                name = codeToName.getOrDefault(code, "");
            }
            codeToName.putIfAbsent(code, name);
        }

        List<String> labels = new ArrayList<>();
        for (Map.Entry<String, String> e : codeToName.entrySet()) {
            labels.add(PostProcessingKouteiNaiyoMasterLookup.displayLabel(e.getKey(), e.getValue()));
        }
        return new Snapshot(
                abs,
                mtime,
                hasCode,
                hasName,
                Map.copyOf(codeToName),
                List.copyOf(labels));
    }

    private static PlanInputTabularIo.TabularSheet readProcessingPlanSheet(Path path)
            throws IOException {
        String low = path.getFileName().toString().toLowerCase();
        if (low.endsWith(".xlsx")
                || low.endsWith(".xlsm")
                || low.endsWith(".xltx")
                || low.endsWith(".xltm")) {
            return TaskInputSourceRawGridIo.applyAladdinProcessingPlanDisplaySteps(
                    TaskInputSourceRawGridIo.readRaw(path, 0));
        }
        if (low.endsWith(".csv")) {
            return PlanInputTabularIo.read(path, "");
        }
        throw new IOException("加工計画の形式未対応: " + path);
    }

    private static int indexOfMachineCodeHeader(List<String> headers) {
        int idx = indexOfHeader(headers, COL_MACHINE_CODE);
        if (idx >= 0) {
            return idx;
        }
        return indexOfHeader(headers, COL_MACHINE_CODE_LEGACY);
    }

    private static int indexOfHeader(List<String> headers, String name) {
        if (headers == null) {
            return -1;
        }
        for (int i = 0; i < headers.size(); i++) {
            String h = headers.get(i);
            if (h != null && name.equals(h.trim())) {
                return i;
            }
        }
        return -1;
    }

    private static String cellAt(List<String> row, int index) {
        if (row == null || index < 0 || index >= row.size()) {
            return "";
        }
        String v = row.get(index);
        return v != null ? v.trim() : "";
    }
}
