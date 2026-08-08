package jp.co.pm.ai.desktop.io;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.ui.PlanInputAiSpecialParseColumn;
import javafx.collections.ObservableList;

/**
 * 段階2 が Excel へ書けなかったときの AI 解析結果（{@code code/json/plan_input_ai_special_parse.json}）。
 * 配台計画_タスク入力タブの再読込時に表へマージする。
 */
public final class PlanInputAiSpecialParseSidecar {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private PlanInputAiSpecialParseSidecar() {}

    public static Path resolveSidecarPath(Map<String, String> ui) {
        return AppPaths.resolveCodeDir(ui).resolve("json").resolve("plan_input_ai_special_parse.json");
    }

    /**
     * サイドカーがあれば解析列へマージする。
     *
     * @return マージしたセル数（0 なら未適用）
     */
    public static int applyIfPresent(
            Map<String, String> ui,
            String sheetName,
            List<String> headers,
            List<? extends ObservableList<String>> rows) {
        if (headers == null || rows == null || rows.isEmpty()) {
            return 0;
        }
        int parseCol = PlanInputAiSpecialParseColumn.indexOfParseColumn(headers);
        if (parseCol < 0) {
            return 0;
        }
        Path sidecar = resolveSidecarPath(ui);
        if (!Files.isRegularFile(sidecar)) {
            return 0;
        }
        JsonNode root;
        try {
            root = MAPPER.readTree(sidecar.toFile());
        } catch (IOException ex) {
            return 0;
        }
        String scSheet = root.path("sheet").asText("").trim();
        if (!scSheet.isEmpty() && sheetName != null && !scSheet.equals(sheetName.trim())) {
            return 0;
        }
        JsonNode byRow = root.get("by_excel_row");
        if (byRow == null || !byRow.isObject()) {
            return 0;
        }
        Map<Integer, String> values = new HashMap<>();
        byRow.fields()
                .forEachRemaining(
                        e -> {
                            try {
                                int excelRow = Integer.parseInt(e.getKey());
                                values.put(excelRow, e.getValue().asText(""));
                            } catch (NumberFormatException ignored) {
                                // skip malformed keys
                            }
                        });
        if (values.isEmpty()) {
            return 0;
        }
        int merged = 0;
        for (int dataIndex = 0; dataIndex < rows.size(); dataIndex++) {
            int excelRow = dataIndex + 2;
            if (!values.containsKey(excelRow)) {
                continue;
            }
            ObservableList<String> row = rows.get(dataIndex);
            while (row.size() <= parseCol) {
                row.add("");
            }
            String incoming = values.get(excelRow);
            String current = row.get(parseCol);
            if (incoming == null || incoming.isBlank()) {
                if (current != null && !current.isBlank()) {
                    row.set(parseCol, "");
                    merged++;
                }
            } else if (!incoming.equals(current)) {
                row.set(parseCol, incoming);
                merged++;
            }
        }
        return merged;
    }
}
