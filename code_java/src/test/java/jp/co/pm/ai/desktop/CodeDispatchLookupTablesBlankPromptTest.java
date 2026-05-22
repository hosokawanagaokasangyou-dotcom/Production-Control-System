package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class CodeDispatchLookupTablesBlankPromptTest {

    @TempDir Path tmp;

    @Test
    void collectGroupsBlankIssuesByProductAndUsedRaw() throws IOException {
        Path code = tmp.resolve("code");
        Files.createDirectories(code);
        Files.writeString(
                code.resolve(CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK),
                "製品名,製品厚み\nNEW-PROD,\n",
                StandardCharsets.UTF_8);
        Files.writeString(
                code.resolve(CodeDispatchLookupTablesMerge.FILE_USED_RAW_WIDTH),
                "使用原反,原反幅\nNEW-RAW,\n",
                StandardCharsets.UTF_8);

        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_CODE_DIR, code.toString());

        var vr = CodeDispatchLookupTablesValidator.validateNoBlankValues(ui);
        assertFalse(vr.ok());

        var bundle = CodeDispatchLookupTablesBlankPrompt.collectPrompt(ui, vr);
        assertEquals(1, bundle.products().size());
        assertEquals("NEW-PROD", bundle.products().getFirst().productName());
        assertTrue(bundle.products().getFirst().needThickness());
        assertEquals(1, bundle.usedRaws().size());
        assertEquals("NEW-RAW", bundle.usedRaws().getFirst().usedRaw());
        assertTrue(bundle.usedRaws().getFirst().needRawWidth());
    }

    @Test
    void applyInputsWritesValuesToTables() throws IOException {
        Path code = tmp.resolve("code");
        Files.createDirectories(code);
        Files.writeString(
                code.resolve(CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK),
                "製品名,製品厚み\nNEW-PROD,\n",
                StandardCharsets.UTF_8);
        Files.writeString(
                code.resolve(CodeDispatchLookupTablesMerge.FILE_USED_RAW_WIDTH),
                "使用原反,原反幅\nNEW-RAW,\n",
                StandardCharsets.UTF_8);

        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_CODE_DIR, code.toString());
        var applied =
                CodeDispatchLookupTablesBlankPrompt.applyInputs(
                        ui,
                        List.of(
                                new CodeDispatchLookupTablesBlankPrompt.ProductInput(
                                        "NEW-PROD", "", "", "3.3", "")),
                        List.of(
                                new CodeDispatchLookupTablesBlankPrompt.UsedRawInput(
                                        "NEW-RAW", "", "1040")));

        assertTrue(applied.updatedFields() >= 2);

        var vr = CodeDispatchLookupTablesValidator.validateNoBlankValues(ui);
        assertTrue(vr.ok());

        String thick =
                Files.readString(
                        code.resolve(CodeDispatchLookupTablesMerge.FILE_PRODUCT_THICK),
                        StandardCharsets.UTF_8);
        assertTrue(thick.contains("NEW-PROD,3.3"));
    }
}
