package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class OperatorUserPathsTest {

    @Test
    void sanitizeOperatorDirName_rejectsDotSegments() {
        assertEquals(OperatorUserPaths.UNKNOWN_OPERATOR_DIR, OperatorUserPaths.sanitizeOperatorDirName(".."));
        assertEquals(OperatorUserPaths.UNKNOWN_OPERATOR_DIR, OperatorUserPaths.sanitizeOperatorDirName("."));
        assertEquals(OperatorUserPaths.UNKNOWN_OPERATOR_DIR, OperatorUserPaths.sanitizeOperatorDirName("../x"));
    }
}
