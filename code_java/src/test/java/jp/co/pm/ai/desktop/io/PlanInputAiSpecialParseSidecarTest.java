package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class PlanInputAiSpecialParseSidecarTest {

    @Test
    void applyIfPresentMergesByExcelRow(@TempDir Path temp) throws Exception {
        Path jsonDir = temp.resolve("json");
        Files.createDirectories(jsonDir);
        Path sidecar = jsonDir.resolve("plan_input_ai_special_parse.json");
        Files.writeString(
                sidecar,
                """
                {
                  "version": 1,
                  "sheet": "タスク一覧",
                  "by_excel_row": {
                    "2": "{\\"priority\\":1}"
                  }
                }
                """);

        List<String> headers = List.of("依頼NO", "AI納期回答_解析");
        ObservableList<String> row = FXCollections.observableArrayList("Y3-26", "");
        List<ObservableList<String>> rows = new ArrayList<>();
        rows.add(row);

        int merged =
                PlanInputAiSpecialParseSidecar.applyIfPresent(
                        Map.of("PM_AI_CODE_DIR", temp.toString()),
                        "タスク一覧",
                        headers,
                        rows);

        assertEquals(1, merged);
        assertEquals("{\"priority\":1}", row.get(1));
    }
}
