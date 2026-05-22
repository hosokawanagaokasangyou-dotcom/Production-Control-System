package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class CodeDispatchLookupTablesValidatorTest {

    @TempDir Path tmp;

    @Test
    void okWhenAllValuesPresent() throws IOException {
        Path code = tmp.resolve("code");
        Files.createDirectories(code);
        Path summaryDir = tmp.resolve("shared");
        Files.createDirectories(summaryDir);
        Path summary = summaryDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summary);
        Files.writeString(
                summaryDir.resolve(CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK),
                "製品名,製品厚み\nA-001,2.0\n",
                StandardCharsets.UTF_8);
        var vr =
                CodeDispatchLookupTablesValidator.validateNoBlankValues(
                        Map.of(
                                AppPaths.KEY_PM_AI_REPO_ROOT,
                                tmp.toString(),
                                AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                                summary.toString()));
        assertTrue(vr.ok());
    }

    @Test
    void detectsBlankThicknessValue() throws IOException {
        Path code = tmp.resolve("code");
        Files.createDirectories(code);
        Path summaryDir = tmp.resolve("shared");
        Files.createDirectories(summaryDir);
        Path summary = summaryDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summary);
        Files.writeString(
                summaryDir.resolve(CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK),
                "製品名,製品厚み\nC4300-1056-820x114YA,\nFILLED,3.0\n",
                StandardCharsets.UTF_8);
        var vr =
                CodeDispatchLookupTablesValidator.validateNoBlankValues(
                        Map.of(
                                AppPaths.KEY_PM_AI_REPO_ROOT,
                                tmp.toString(),
                                AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                                summary.toString()));
        assertFalse(vr.ok());
        assertEquals(1, vr.issues().size());
        assertEquals("C4300-1056-820x114YA", vr.issues().getFirst().key());
        assertEquals("製品名→厚み(mm)", vr.issues().getFirst().tableLabelJa());
    }

    @Test
    void skipsMissingTableFiles() throws IOException {
        Path code = tmp.resolve("code");
        Files.createDirectories(code);
        Path summaryDir = tmp.resolve("shared");
        Files.createDirectories(summaryDir);
        Path summary = summaryDir.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createFile(summary);
        var vr =
                CodeDispatchLookupTablesValidator.validateNoBlankValues(
                        Map.of(
                                AppPaths.KEY_PM_AI_REPO_ROOT,
                                tmp.toString(),
                                AppPaths.KEY_PM_AI_CODE_DIR,
                                code.toString(),
                                AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                                summary.toString()));
        assertTrue(vr.ok());
    }
}
