package jp.co.pm.ai.desktop.dispatch;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class ResultDispatchProvenanceTest {

    @TempDir Path tmp;

    @Test
    void readPrefersJsonGeneratedFields() throws Exception {
        Path json = tmp.resolve("結果_配台表.json");
        Files.writeString(
                json,
                """
                {
                  "format_version": 1,
                  "sheet_name": "結果_配台表",
                  "excel_table_name": "_t結果_配台表",
                  "generated_by": "テスト操作者",
                  "generated_at": "2026-09-11T10:30:00+09:00",
                  "columns": ["依頼NO"],
                  "row_count": 0,
                  "rows": []
                }
                """,
                StandardCharsets.UTF_8);
        var info = ResultDispatchProvenance.read(json).orElseThrow();
        assertEquals("テスト操作者", info.generatedBy());
        assertTrue(info.hasGenerator());
        assertFalse(info.generatedAtFromFileMtime());
        assertTrue(info.generatedAt().isPresent());
        assertTrue(info.formatDisplayLine().contains("テスト操作者"));
        assertFalse(info.formatDisplayLine().contains("ファイル更新"));
    }

    @Test
    void readFallsBackToFileMtimeWhenMetaMissing() throws Exception {
        Path json = tmp.resolve("結果_配台表.json");
        Files.writeString(
                json,
                """
                {
                  "format_version": 1,
                  "sheet_name": "結果_配台表",
                  "columns": ["依頼NO"],
                  "row_count": 0,
                  "rows": []
                }
                """,
                StandardCharsets.UTF_8);
        var info = ResultDispatchProvenance.read(json).orElseThrow();
        assertFalse(info.hasGenerator());
        assertTrue(info.generatedAtFromFileMtime());
        assertTrue(info.generatedAt().isPresent());
        assertTrue(info.formatDisplayLine().contains("未記録"));
        assertTrue(info.formatDisplayLine().contains("ファイル更新"));
    }

    @Test
    void writeStampsProvenanceAndRoundTrips() throws Exception {
        Path json = tmp.resolve("out.json");
        ResultDispatchDocument doc =
                new ResultDispatchDocument(
                        new java.util.ArrayList<>(List.of("依頼NO")),
                        new java.util.ArrayList<>(
                                List.of(new java.util.LinkedHashMap<>(Map.of("依頼NO", "R1")))));
        doc.setGeneratedBy("書込者");
        doc.setGeneratedAt(ResultDispatchProvenance.formatInstantIso(Instant.parse("2026-09-11T01:00:00Z")));
        ResultDispatchJsonIo.write(json, doc);
        ResultDispatchDocument read = ResultDispatchJsonIo.read(json);
        assertEquals("書込者", read.generatedBy());
        assertEquals("2026-09-11T01:00:00Z", read.generatedAt());
    }
}
