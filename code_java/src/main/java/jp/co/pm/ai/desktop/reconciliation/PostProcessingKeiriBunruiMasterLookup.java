package jp.co.pm.ai.desktop.reconciliation;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.PlanInputTabularIo;
import jp.co.pm.ai.desktop.io.PostProcessingProductMasterIo;

/**
 * {@code 後加工経理分類マスタ.xlsx}（列 {@code 経理分類} / {@code 経理分類名}）のコード→名称参照。
 * 商品マスタの {@code 単価分類コード}（Aladdin「経理分類」・現行 xlsx 見出し）選択に使う。
 */
public final class PostProcessingKeiriBunruiMasterLookup {

    public static final String FILE_NAME = "後加工経理分類マスタ.xlsx";

    /** 経理分類マスタでコンボ化する商品マスタ列（154列 xlsx の見出し名）。 */
    public static final String PRODUCT_COLUMN = "単価分類コード";

    private static volatile Snapshot cached;

    private PostProcessingKeiriBunruiMasterLookup() {}

    public record Snapshot(
            Path path,
            long mtime,
            Map<String, String> codeToName,
            List<String> comboLabels) {

        public static Snapshot empty() {
            return new Snapshot(Path.of(""), -1L, Map.of(), List.of());
        }

        public boolean loaded() {
            return !codeToName.isEmpty();
        }
    }

    public static boolean isKeiriBunruiProductColumn(String columnName) {
        return columnName != null && PRODUCT_COLUMN.equals(columnName.trim());
    }

    public static Snapshot snapshot(Map<String, String> ui) throws IOException {
        Path masterDir = AppPaths.resolveAladdinMasterDir(ui);
        Path path = masterDir.resolve(FILE_NAME);
        if (!Files.isRegularFile(path)) {
            return Snapshot.empty();
        }
        long mtime = Files.getLastModifiedTime(path).toMillis();
        Path abs = path.toAbsolutePath().normalize();
        Snapshot hit = cached;
        if (hit != null && Objects.equals(hit.path(), abs) && hit.mtime() == mtime) {
            return hit;
        }
        synchronized (PostProcessingKeiriBunruiMasterLookup.class) {
            hit = cached;
            if (hit != null && Objects.equals(hit.path(), abs) && hit.mtime() == mtime) {
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

    public static String normalizeCode(String raw) {
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
        int colon = val.indexOf(':');
        if (colon >= 0) {
            val = val.substring(0, colon).trim();
        }
        int sp = val.indexOf(' ');
        if (sp > 0) {
            val = val.substring(0, sp).trim();
        }
        try {
            int n = (int) Double.parseDouble(val);
            if (n >= 0 && n <= 99) {
                return String.format("%02d", n);
            }
        } catch (NumberFormatException ignored) {
        }
        return val;
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

    public static String resolveName(Snapshot snap, String rawCode) {
        if (snap == null) {
            return "";
        }
        String code = normalizeCode(rawCode);
        if (code.isEmpty()) {
            return "";
        }
        return snap.codeToName().getOrDefault(code, "");
    }

    public static String resolveCodeFromComboInput(Snapshot snap, String text) {
        if (text == null) {
            return "";
        }
        String trimmed = text.trim();
        if (trimmed.isEmpty()) {
            return "";
        }
        String norm = normalizeCode(trimmed);
        if (snap != null && snap.codeToName().containsKey(norm)) {
            return norm;
        }
        if (snap != null) {
            for (String label : snap.comboLabels()) {
                if (label.equals(trimmed) || label.startsWith(norm + " ")) {
                    int sp = label.indexOf(' ');
                    return sp > 0 ? label.substring(0, sp) : norm;
                }
            }
        }
        return norm;
    }

    private static Snapshot load(Path path, long mtime) throws IOException {
        Map<String, String> codeToName = new LinkedHashMap<>();
        PlanInputTabularIo.TabularSheet sheet =
                PlanInputTabularIo.read(path, PostProcessingProductMasterIo.DEFAULT_SHEET_NAME);
        for (List<String> row : sheet.rows()) {
            Map<String, String> map =
                    PostProcessingProductMasterIo.rowToMap(sheet.headers(), row);
            String code = normalizeCode(map.getOrDefault("経理分類", ""));
            if (code.isEmpty()) {
                continue;
            }
            codeToName.putIfAbsent(code, map.getOrDefault("経理分類名", "").trim());
        }
        List<String> labels = new ArrayList<>();
        for (Map.Entry<String, String> e : codeToName.entrySet()) {
            labels.add(displayLabel(e.getKey(), e.getValue()));
        }
        return new Snapshot(path, mtime, Map.copyOf(codeToName), List.copyOf(labels));
    }
}
