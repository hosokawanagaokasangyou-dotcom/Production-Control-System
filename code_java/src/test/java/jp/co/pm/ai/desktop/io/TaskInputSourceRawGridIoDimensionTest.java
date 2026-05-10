package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class TaskInputSourceRawGridIoDimensionTest {

    @Test
    void parseDimensionMaxRow0FromSheetXmlPrefix_readsRangeEnd() {
        String xml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                        + "<dimension ref=\"A1:AP504\"/>";
        assertEquals(503, TaskInputSourceRawGridIo.parseDimensionMaxRow0FromSheetXmlPrefix(xml));
    }

    @Test
    void parseDimensionMaxRow0FromSheetXmlPrefix_missingReturnsMinus1() {
        assertEquals(-1, TaskInputSourceRawGridIo.parseDimensionMaxRow0FromSheetXmlPrefix("<worksheet/>"));
    }
}
