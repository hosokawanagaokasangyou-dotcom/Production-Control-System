package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class LogLineKindTest {

    @Test
    void portableSyncExtractPath_withExceptionsModule_isNormal() {
        assertEquals(
                LogLineKind.NORMAL,
                LogLineKind.classify(
                        "[portable-sync] 展開: pm-ai-data/runtime/python-embed/Lib/site-packages/anyio/_core/_exceptions.py"));
    }

    @Test
    void portableSyncExtractPath_withAssemblyException_isNormal() {
        assertEquals(
                LogLineKind.NORMAL,
                LogLineKind.classify(
                        "[portable-sync] 展開: runtime/legal/java.base/ASSEMBLY_EXCEPTION"));
    }

    @Test
    void portableSyncExtractPath_withWarningsInName_isNormal() {
        assertEquals(
                LogLineKind.NORMAL,
                LogLineKind.classify(
                        "[portable-sync] 展開: pm-ai-data/runtime/python-embed/Lib/site-packages/numpy/typing/tests/data/pass/warnings_and_errors.py"));
    }

    @Test
    void realExceptionLine_isError() {
        assertEquals(
                LogLineKind.ERROR,
                LogLineKind.classify("Traceback (most recent call last): ValueError"));
    }

    @Test
    void portableSyncCleanupFailure_canStillBeError() {
        assertEquals(
                LogLineKind.ERROR,
                LogLineKind.classify("[portable-sync] cleanup walk failed: java.io.IOException: denied"));
    }
}
