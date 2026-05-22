package jp.co.pm.ai.desktop;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.io.CodeDispatchLookupTableIo;
import jp.co.pm.ai.desktop.io.CodeDispatchLookupTableIo.KeyValTable;
import jp.co.pm.ai.planning.stage2.core.Stage2RollUnitLengthTables;

/**
 * {@code code/} 配下の材料・製品種類ルックアップ表に、キーはあるが値が空欄の行が無いか検証する。
 * 段階2・段階3実行前のゲートに使う。
 */
public final class CodeDispatchLookupTablesValidator {

    private record TableSpec(String relativePath, String defaultHeaderLine, String labelJa) {}

    private static final List<TableSpec> TABLES =
            List.of(
                    new TableSpec(
                            CodeDispatchLookupTablesMerge.FILE_USED_RAW_ROLL,
                            CodeDispatchLookupTablesMerge.FILE_USED_RAW_ROLL.replace(".txt", ""),
                            "使用原反→ロール長(m)"),
                    new TableSpec(
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_ROLL,
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_ROLL.replace(".txt", ""),
                            "製品名→ロール長(m)"),
                    new TableSpec(
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_WIDTH,
                            "製品名,製品幅",
                            "製品名→製品幅(mm)"),
                    new TableSpec(
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK,
                            "製品名,製品厚み",
                            "製品名→厚み(mm)"),
                    new TableSpec(
                            CodeDispatchLookupTablesMerge.FILE_PRODUCT_LENGTH,
                            "製品名,製品長",
                            "製品名→製品長(mm)"),
                    new TableSpec(
                            CodeDispatchLookupTablesMerge.FILE_USED_RAW_WIDTH,
                            "使用原反,原反幅",
                            "使用原反→原反幅(mm)"));

    public record BlankValueIssue(String tableLabelJa, String key) {}

    public record ValidationResult(List<BlankValueIssue> issues) {
        public boolean ok() {
            return issues == null || issues.isEmpty();
        }

        /** ダイアログ向け（先頭数件＋省略）。 */
        public String messageJa(int maxLines) {
            if (ok()) {
                return "";
            }
            int cap = maxLines > 0 ? maxLines : 8;
            StringBuilder sb =
                    new StringBuilder(
                            "材料・製品種類情報（code/）に値が空欄の行があります。"
                                    + "「材料・製品種類情報」タブで入力してから再実行してください。\n\n");
            int n = Math.min(issues.size(), cap);
            for (int i = 0; i < n; i++) {
                BlankValueIssue issue = issues.get(i);
                sb.append("・").append(issue.tableLabelJa()).append(": ").append(issue.key()).append('\n');
            }
            if (issues.size() > cap) {
                sb.append("…他 ").append(issues.size() - cap).append(" 件");
            }
            return sb.toString().stripTrailing();
        }

        public List<String> logLines() {
            if (ok()) {
                return List.of();
            }
            List<String> out = new ArrayList<>(issues.size());
            for (BlankValueIssue issue : issues) {
                out.add(
                        "[材料テーブル] 値が空欄: "
                                + issue.tableLabelJa()
                                + " / キー="
                                + issue.key());
            }
            return out;
        }
    }

    private CodeDispatchLookupTablesValidator() {}

    public static ValidationResult validateNoBlankValues(Map<String, String> ui) throws IOException {
        Map<String, String> u = ui != null ? ui : Map.of();
        Path codeDir = AppPaths.resolveCodeDir(u);
        List<BlankValueIssue> issues = new ArrayList<>();
        for (TableSpec spec : TABLES) {
            Path path = codeDir.resolve(spec.relativePath());
            if (!Files.isRegularFile(path)) {
                continue;
            }
            KeyValTable table = CodeDispatchLookupTableIo.readOrEmpty(path, spec.defaultHeaderLine());
            collectBlankValues(spec.labelJa(), table.rows(), issues);
        }
        return new ValidationResult(List.copyOf(issues));
    }

    private static void collectBlankValues(
            String tableLabelJa, LinkedHashMap<String, String> rows, List<BlankValueIssue> issues) {
        if (rows == null || rows.isEmpty()) {
            return;
        }
        for (Map.Entry<String, String> e : rows.entrySet()) {
            String key = e.getKey() != null ? e.getKey().strip() : "";
            if (key.isEmpty()) {
                continue;
            }
            String nk = Stage2RollUnitLengthTables.normalizeKey(key);
            if (nk.isEmpty()) {
                continue;
            }
            String val = e.getValue() != null ? e.getValue().strip() : "";
            if (val.isEmpty()) {
                issues.add(new BlankValueIssue(tableLabelJa, key));
            }
        }
    }
}
