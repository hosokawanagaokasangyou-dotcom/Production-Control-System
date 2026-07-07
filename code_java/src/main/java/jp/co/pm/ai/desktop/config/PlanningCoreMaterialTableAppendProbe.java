package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * {@code planning_core} の段階1材料テーブル追記仕様（空欄追記・行残し）が新実装かを判定する。
 *
 * <p>{@code _core.py} はファサード化され定数は {@code planning_core/core/columns.py} 等に分離されているため、
 * 単一ファイルの文字列検索だけでは誤って「旧」と判定しない。
 */
public final class PlanningCoreMaterialTableAppendProbe {

    private static final Pattern BUILD_ID =
            Pattern.compile("_STAGE1_MATERIAL_TABLE_APPEND_BUILD\\s*=\\s*\"([^\"]+)\"");

    public enum Spec {
        CURRENT("新（空欄追記・行は残す）"),
        LEGACY("旧（要 code/python 更新）");

        private final String logLabel;

        Spec(String logLabel) {
            this.logLabel = logLabel;
        }

        public String logLabel() {
            return logLabel;
        }
    }

    public record Result(Spec spec, Optional<String> buildId) {}

    private PlanningCoreMaterialTableAppendProbe() {}

    public static Result detect(Path codePythonDir) {
        if (codePythonDir == null) {
            return new Result(Spec.LEGACY, Optional.empty());
        }
        Path planningCore = codePythonDir.resolve("planning_core");
        Path corePy = planningCore.resolve("_core.py");
        if (!Files.isRegularFile(corePy)) {
            return new Result(Spec.LEGACY, Optional.empty());
        }
        try {
            String coreText = Files.readString(corePy, StandardCharsets.UTF_8);
            if (hasNewMaterialTableAppendMarker(coreText)) {
                return new Result(Spec.CURRENT, extractBuildId(coreText));
            }
            if (isPlanningCoreFacade(coreText)) {
                Path columnsPy = planningCore.resolve("core").resolve("columns.py");
                if (Files.isRegularFile(columnsPy)) {
                    String columnsText = Files.readString(columnsPy, StandardCharsets.UTF_8);
                    if (hasNewMaterialTableAppendMarker(columnsText)) {
                        return new Result(Spec.CURRENT, extractBuildId(columnsText));
                    }
                }
            }
            return new Result(Spec.LEGACY, Optional.empty());
        } catch (IOException ex) {
            return new Result(Spec.LEGACY, Optional.empty());
        }
    }

    static boolean isPlanningCoreFacade(String coreText) {
        return coreText.contains("_exec_into_ns") && coreText.contains("_MODULE_ORDER");
    }

    static boolean hasNewMaterialTableAppendMarker(String text) {
        return text.contains("_STAGE1_MATERIAL_TABLE_APPEND_BUILD")
                || text.contains("材料テーブルへ追記し製品厚みは空欄で出力");
    }

    static Optional<String> extractBuildId(String text) {
        Matcher m = BUILD_ID.matcher(text);
        if (m.find()) {
            return Optional.of(m.group(1));
        }
        return Optional.empty();
    }
}
