package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.util.HashMap;
import java.util.Map;

import org.junit.jupiter.api.Test;

class RequestFormOriginalIndexSheetMergerTest {

    private static RequestFormOriginalIndexSheetReader.IndexEntry indexWithInputDate(
            String inputDate) {
        return new RequestFormOriginalIndexSheetReader.IndexEntry(
                "JR260601", "", "", inputDate, "", "", "", "", "");
    }

    @Test
    void preservesSheetInputDate_beforeIndexOverride() {
        Map<String, String> raw = new HashMap<>();
        raw.put("投入日", "2026/7/5");

        RequestFormOriginalIndexSheetMerger.applyIndexOverrides(raw, indexWithInputDate("2026/7/6"));

        assertEquals("2026/7/5", raw.get(RequestFormOriginalIndexSheetMeta.KEY_SHEET_INPUT_DATE));
        assertEquals("2026/7/6", raw.get("投入日"));
        assertEquals("2026/7/6", raw.get(RequestFormOriginalIndexSheetMeta.KEY_INPUT_DATE));
    }

    @Test
    void keepsSheetInputDate_whenIndexInputDateBlank() {
        Map<String, String> raw = new HashMap<>();
        raw.put("投入日", "2026/7/5");

        RequestFormOriginalIndexSheetMerger.applyIndexOverrides(raw, indexWithInputDate(""));

        assertEquals("2026/7/5", raw.get(RequestFormOriginalIndexSheetMeta.KEY_SHEET_INPUT_DATE));
        assertEquals("2026/7/5", raw.get("投入日"));
    }
}
