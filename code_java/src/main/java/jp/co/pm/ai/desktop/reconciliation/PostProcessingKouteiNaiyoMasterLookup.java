package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.OptionalInt;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

/**
 * {@code 後加工工程マスタ.xlsx} / {@code 後加工加工内容マスタ.xlsx} のコード→名称参照。
 * 商品マスタ編集フォームの工程コード・加工内容コード選択に使う。
 */
public final class PostProcessingKouteiNaiyoMasterLookup {

    public static final String KOUTEI_FILE_NAME = "後加工工程マスタ.xlsx";
    public static final String NAIYO_FILE_NAME = "後加工加工内容マスタ.xlsx";

    private static final Pattern STEP_CODE_COLUMN =
            Pattern.compile("^(工程コード|加工内容コード)(\\d+)$");

    private static volatile Snapshot cached;

    private PostProcessingKouteiNaiyoMasterLookup() {}

    public record NaiyoEntry(
            String naiyoCode, String naiyoName, String kouteiCode, String kouteiName) {}

    public record Snapshot(
            Path kouteiPath,
            long kouteiMtime,
            Path naiyoPath,
            long naiyoMtime,
            Map<String, String> kouteiCodeToName,
            List<String> kouteiComboLabels,
            Map<String, NaiyoEntry> naiyoCodeToEntry,
            List<String> naiyoComboLabels) {

        public static Snapshot empty() {
            return new Snapshot(
                    Path.of(""),
                    -1L,
                    Path.of(""),
                    -1L,
                    Map.of(),
                    List.of(),
                    Map.of(),
                    List.of());
        }

        public boolean loaded() {
            return !kouteiCodeToName.isEmpty() || !naiyoCodeToEntry.isEmpty();
        }
    }

    public static Snapshot snapshot(Map<String, String> ui) throws IOException {
        Path masterDir = AppPaths.resolveAladdinMasterDir(ui);
        Path kouteiPath = masterDir.resolve(KOUTEI_FILE_NAME);
        Path naiyoPath = masterDir.resolve(NAIYO_FILE_NAME);
        if (!Files.isRegularFile(kouteiPath) && !Files.isRegularFile(naiyoPath)) {
            return Snapshot.empty();
        }
        long kMtime =
                Files.isRegularFile(kouteiPath)
                        ? Files.getLastModifiedTime(kouteiPath).toMillis()
                        : -1L;
        long nMtime =
                Files.isRegularFile(naiyoPath)
                        ? Files.getLastModifiedTime(naiyoPath).toMillis()
                        : -1L;
        Path kAbs = kouteiPath.toAbsolutePath().normalize();
        Path nAbs = naiyoPath.toAbsolutePath().normalize();
        Snapshot hit = cached;
        if (hit != null
                && Objects.equals(hit.kouteiPath(), kAbs)
                && hit.kouteiMtime() == kMtime
                && Objects.equals(hit.naiyoPath(), nAbs)
                && hit.naiyoMtime() == nMtime) {
            return hit;
        }
        synchronized (PostProcessingKouteiNaiyoMasterLookup.class) {
            hit = cached;
            if (hit != null
                    && Objects.equals(hit.kouteiPath(), kAbs)
                    && hit.kouteiMtime() == kMtime
                    && Objects.equals(hit.naiyoPath(), nAbs)
                    && hit.naiyoMtime() == nMtime) {
                return hit;
            }
            Snapshot loaded = load(kAbs, kMtime, nAbs, nMtime);
            cached = loaded;
            return loaded;
        }
    }

    public static void invalidate() {
        cached = null;
    }

    public static boolean isKouteiCodeColumn(String columnName) {
        if (columnName == null) {
            return false;
        }
        Matcher m = STEP_CODE_COLUMN.matcher(columnName.trim());
        return m.matches() && "工程コード".equals(m.group(1));
    }

    public static boolean isNaiyoCodeColumn(String columnName) {
        if (columnName == null) {
            return false;
        }
        Matcher m = STEP_CODE_COLUMN.matcher(columnName.trim());
        return m.matches() && "加工内容コード".equals(m.group(1));
    }

    public static OptionalInt stepIndex(String columnName) {
        if (columnName == null) {
            return OptionalInt.empty();
        }
        Matcher m = STEP_CODE_COLUMN.matcher(columnName.trim());
        if (!m.matches()) {
            return OptionalInt.empty();
        }
        try {
            return OptionalInt.of(Integer.parseInt(m.group(2)));
        } catch (NumberFormatException ex) {
            return OptionalInt.empty();
        }
    }

    public static String kouteiColumnForStep(int step) {
        return "工程コード" + step;
    }

    public static String normalizeKouteiCode(String raw) {
        return normalizeCode(raw, 4);
    }

    public static String normalizeNaiyoCode(String raw) {
        return normalizeCode(raw, 4);
    }

    public static String normalizeCode(String raw, int padLen) {
        if (raw == null) {
            return "";
        }
        String val = raw.trim();
        if (val.isEmpty()) {
            return "";
        }
        if (val.endsWith(".0")) {
            val = val.substring(0, val.length() - 2).trim();
        }
        try {
            int n = (int) Double.parseDouble(val);
            if (padLen > 0) {
                return String.format("%0" + padLen + "d", n);
            }
            return String.valueOf(n);
        } catch (NumberFormatException ex) {
            return val;
        }
    }

    public static String displayLabel(String code, String name) {
        String c = code != null ? code.trim() : "";
        String n = name != null ? name.trim() : "";
        if (c.isEmpty() && n.isEmpty()) {
            return "";
        }
        if (n.isEmpty()) {
            return c;
        }
        if (c.isEmpty()) {
            return n;
        }
        return c + " " + n;
    }

    public static String resolveKouteiName(Snapshot snap, String rawCode) {
        if (snap == null) {
            return "";
        }
        String code = normalizeKouteiCode(rawCode);
        if (code.isEmpty()) {
            return "";
        }
        return snap.kouteiCodeToName().getOrDefault(code, "");
    }

    public static NaiyoEntry resolveNaiyo(Snapshot snap, String rawCode) {
        if (snap == null) {
            return null;
        }
        String code = normalizeNaiyoCode(rawCode);
        if (code.isEmpty()) {
            return null;
        }
        return snap.naiyoCodeToEntry().get(code);
    }

    public static String resolveCodeFromComboInput(
            Snapshot snap, String text, boolean koutei, int padLen) {
        if (text == null) {
            return "";
        }
        String trimmed = text.trim();
        if (trimmed.isEmpty()) {
            return "";
        }
        String norm = normalizeCode(trimmed, padLen);
        if (koutei) {
            if (snap != null && snap.kouteiCodeToName().containsKey(norm)) {
                return norm;
            }
            if (snap != null) {
                for (String label : snap.kouteiComboLabels()) {
                    if (label.equals(trimmed) || label.startsWith(norm + " ")) {
                        int sp = label.indexOf(' ');
                        return sp > 0 ? label.substring(0, sp) : norm;
                    }
                }
            }
        } else {
            if (snap != null && snap.naiyoCodeToEntry().containsKey(norm)) {
                return norm;
            }
            if (snap != null) {
                for (String label : snap.naiyoComboLabels()) {
                    if (label.equals(trimmed) || label.startsWith(norm + " ")) {
                        int sp = label.indexOf(' ');
                        return sp > 0 ? label.substring(0, sp) : norm;
                    }
                }
            }
        }
        return norm;
    }

    private static Snapshot load(Path kouteiPath, long kMtime, Path naiyoPath, long nMtime)
            throws IOException {
        Map<String, String> kouteiCodeToName = new LinkedHashMap<>();
        if (Files.isRegularFile(kouteiPath)) {
            PlanInputTabularIo.TabularSheet sheet =
                    PlanInputTabularIo.read(kouteiPath, PostProcessingProductMasterIo.DEFAULT_SHEET_NAME);
            for (List<String> row : sheet.rows()) {
                Map<String, String> map =
                        PostProcessingProductMasterIo.rowToMap(sheet.headers(), row);
                String code = normalizeKouteiCode(map.getOrDefault("工程コード", ""));
                if (code.isEmpty()) {
                    continue;
                }
                kouteiCodeToName.putIfAbsent(code, map.getOrDefault("工程名", "").trim());
            }
        }

        Map<String, NaiyoEntry> naiyoCodeToEntry = new LinkedHashMap<>();
        if (Files.isRegularFile(naiyoPath)) {
            PlanInputTabularIo.TabularSheet sheet =
                    PlanInputTabularIo.read(naiyoPath, PostProcessingProductMasterIo.DEFAULT_SHEET_NAME);
            for (List<String> row : sheet.rows()) {
                Map<String, String> map =
                        PostProcessingProductMasterIo.rowToMap(sheet.headers(), row);
                String naiyoCode = normalizeNaiyoCode(map.getOrDefault("加工内容コード", ""));
                if (naiyoCode.isEmpty()) {
                    continue;
                }
                String kouteiCode = normalizeKouteiCode(map.getOrDefault("工程コード", ""));
                String kouteiName = kouteiCodeToName.getOrDefault(kouteiCode, "");
                naiyoCodeToEntry.putIfAbsent(
                        naiyoCode,
                        new NaiyoEntry(
                                naiyoCode,
                                map.getOrDefault("加工内容名", "").trim(),
                                kouteiCode,
                                kouteiName));
            }
        }

        List<String> kouteiLabels = new ArrayList<>();
        for (Map.Entry<String, String> e : kouteiCodeToName.entrySet()) {
            kouteiLabels.add(displayLabel(e.getKey(), e.getValue()));
        }

        List<String> naiyoLabels = new ArrayList<>();
        for (NaiyoEntry e : naiyoCodeToEntry.values()) {
            naiyoLabels.add(displayLabel(e.naiyoCode(), e.naiyoName()));
        }

        return new Snapshot(
                kouteiPath,
                kMtime,
                naiyoPath,
                nMtime,
                Map.copyOf(kouteiCodeToName),
                List.copyOf(kouteiLabels),
                Map.copyOf(naiyoCodeToEntry),
                List.copyOf(naiyoLabels));
    }
}
