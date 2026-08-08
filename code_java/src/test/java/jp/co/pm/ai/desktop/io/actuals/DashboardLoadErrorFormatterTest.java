package jp.co.pm.ai.desktop.io.actuals;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

class DashboardLoadErrorFormatterTest {

    @Test
    void formatDetail_nullReturnsUnknown() {
        Assertions.assertEquals("原因不明", DashboardLoadErrorFormatter.formatDetail(null));
    }

    @Test
    void formatDetail_usesClassNameOnlyWhenMessageBlank() {
        Assertions.assertEquals(
                "IllegalStateException", DashboardLoadErrorFormatter.formatDetail(new IllegalStateException("  ")));
    }

    @Test
    void formatDetail_joinsCauseChain() {
        Throwable root = new java.io.IOException("ファイルなし");
        Throwable wrapped = new RuntimeException("読込失敗", root);
        Assertions.assertEquals(
                "RuntimeException: 読込失敗\n原因: IOException: ファイルなし",
                DashboardLoadErrorFormatter.formatDetail(wrapped));
    }

    @Test
    void formatDetail_stopsAtMaxDepth() {
        Throwable ex = new RuntimeException("深さ0");
        for (int i = 1; i <= 10; i++) {
            ex = new RuntimeException("深さ" + i, ex);
        }
        String detail = DashboardLoadErrorFormatter.formatDetail(ex);
        Assertions.assertEquals(
                DashboardLoadErrorFormatter.MAX_CAUSE_DEPTH, detail.split("\n").length);
    }

    @Test
    void formatShortDetail_keepsFirstLineOnly() {
        Throwable wrapped = new RuntimeException("外側", new IllegalArgumentException("内側"));
        Assertions.assertEquals(
                "RuntimeException: 外側", DashboardLoadErrorFormatter.formatShortDetail(wrapped));
    }

    @Test
    void formatStackTrace_containsExceptionName() {
        String trace = DashboardLoadErrorFormatter.formatStackTrace(new IllegalStateException("x"));
        Assertions.assertTrue(trace.startsWith("java.lang.IllegalStateException: x"));
        Assertions.assertEquals("", DashboardLoadErrorFormatter.formatStackTrace(null));
    }
}
