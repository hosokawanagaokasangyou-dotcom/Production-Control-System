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

    @Test
    void parseCellReferenceColumn0Based_standardAndMultiLetterReferences() {
        assertEquals(0, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based("A1"));
        assertEquals(1, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based("B5"));
        assertEquals(25, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based("Z100"));
        assertEquals(26, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based("AA1"));
        assertEquals(27, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based("AB2"));
        assertEquals(51, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based("AZ999"));
        assertEquals(52, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based("BA12"));
        assertEquals(32, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based("AG65142"));
        assertEquals(2, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based("$C$10"));
        assertEquals(0, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based(null));
        assertEquals(0, TaskInputSourceRawGridIo.parseCellReferenceColumn0Based(""));
    }
}
