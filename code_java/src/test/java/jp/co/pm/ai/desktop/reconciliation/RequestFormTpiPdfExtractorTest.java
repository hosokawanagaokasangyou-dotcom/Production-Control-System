package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.GlobalInitSettingTarget;

class RequestFormTpiPdfExtractorTest {

    @Test
    void extractFromPdf_jr260604_fields() throws Exception {
        Path pdf = Path.of("src/test/resources/tpi-request-forms/ECOWD-JR260604.pdf");
        assertTrue(Files.isRegularFile(pdf));
        Map<String, String> raw = RequestFormTpiPdfExtractor.extractEntries(pdf.toFile()).get(0);
        assertEquals("JR260604", raw.get("依頼Ｎｏ"));
        assertEquals("R10W-870-870X95", raw.get("製品"), "raw=" + raw);
        assertEquals("950", raw.get("数量1"));
        assertEquals("1000", raw.get("原反数量"));
        assertTrue(raw.get("特記事項2").contains("入庫お願いします。『P000075425』"));
        assertTrue(raw.get("特記事項2").contains("EL原反は6/10"));
    }

    @Test
    void extractFromText_ecowdJr260604() throws IOException {
        String text =
                Files.readString(
                        Path.of("src/test/resources/tpi-request-forms/ecowd-jr260604-clean.txt"),
                        StandardCharsets.UTF_8);
        Map<String, String> raw =
                RequestFormTpiPdfExtractor.extractFromTextForTest(
                        "ECOWDシート加工注文書（JR260604).pdf", text);

        assertEquals("JR260604", raw.get("依頼Ｎｏ"));
        assertEquals("R10W-870-870X95", raw.get("製品"));
        assertEquals("950", raw.get("数量1"));
        assertEquals("1000", raw.get("原反数量"));
    }

    @Test
    void extractFromPdf_jr260603_fields() throws Exception {
        Path pdf = Path.of("src/test/resources/tpi-request-forms/ECOWD-JR260603.pdf");
        assertTrue(Files.isRegularFile(pdf));
        Map<String, String> raw = RequestFormTpiPdfExtractor.extractEntries(pdf.toFile()).get(0);
        assertEquals("R10W-870-870X95", raw.get("製品"), "raw=" + raw);
        assertEquals("475", raw.get("数量1"));
        assertEquals("ﾗｲﾄｸﾞﾚｰ", raw.get("色1"));
        assertEquals("500", raw.get("原反数量"));
        assertEquals("ﾗｲﾄｸﾞﾚｰ", raw.get("原反色"));
        assertTrue(
                raw.get("特記事項2").contains("入庫お願いします。『P000075424』"),
                raw.get("特記事項2"));
        assertTrue(
                raw.get("特記事項2").contains("EL原反は6/10"),
                raw.get("特記事項2"));
    }

    @Test
    void extractFromPdf_jr260604_1_fields() throws Exception {
        Path pdf = Path.of("src/test/resources/tpi-request-forms/ECOWD-JR260604-1.pdf");
        assertTrue(Files.isRegularFile(pdf));
        Map<String, String> raw = RequestFormTpiPdfExtractor.extractEntries(pdf.toFile()).get(0);
        assertEquals("JR260604-1", raw.get("依頼Ｎｏ"));
        assertEquals("R10W-870-870X95", raw.get("製品"), "raw=" + raw);
        assertEquals("190", raw.get("数量1"));
        assertEquals("200", raw.get("原反数量"));
        assertEquals("ナチュラル", raw.get("色1"));
        assertEquals("ナチュラル", raw.get("原反色"));
        assertTrue(raw.get("特記事項2").contains("入庫お願いします。『P000075425』"));
    }

    @Test
    void extractFromText_ecowdSample() throws IOException {
        String text =
                Files.readString(
                        Path.of("src/test/resources/tpi-request-forms/ecowd-jr260604-1.txt"),
                        StandardCharsets.UTF_8);
        Map<String, String> raw =
                RequestFormTpiPdfExtractor.extractFromTextForTest(
                        "ECOWDシート加工注文書（JR260604-1熱融着).pdf", text);

        assertEquals("JR260604-1", raw.get("依頼Ｎｏ"));
        assertEquals("2026-06-22", raw.get("希望納期"));
        assertEquals("40040", raw.get("品名"));
        assertEquals("R10W-870-870X95", raw.get("製品"));
        assertEquals("190", raw.get("数量1"));
        assertEquals("200", raw.get("原反数量"));
        assertEquals("ナチュラル", raw.get("色1"));
        assertEquals("ナチュラル", raw.get("原反色"));
        assertEquals("X000080855", raw.get("契約Ｎｏ"));
        assertEquals("ECOWD", raw.get(RequestFormTpiPdfFieldLayout.META_TPI_LAYOUT));
        assertEquals(
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF,
                raw.get(RequestFormTpiPdfFieldLayout.META_SOURCE_KIND));

        Map<String, String> db = RequestFormOriginalExtractor.buildTpiDbDefaultsFromRaw(raw);
        assertEquals("TPI", db.get("加工区分"));
        assertEquals("JR（屋根）", db.get("用途"));
        assertFalse(db.get("加工内容").isBlank());
    }

    @Test
    void extractFromText_ecowdJr260603() throws IOException {
        String text =
                Files.readString(
                        Path.of("src/test/resources/tpi-request-forms/ecowd-jr260603-clean.txt"),
                        StandardCharsets.UTF_8);
        Map<String, String> raw =
                RequestFormTpiPdfExtractor.extractFromTextForTest(
                        "ECOWDシート加工注文書（JR260603).pdf", text);

        assertEquals("JR260603", raw.get("依頼Ｎｏ"));
        assertEquals("40040", raw.get("品名"));
        assertEquals("R10W-870-870X95", raw.get("製品"));
        assertEquals("475", raw.get("数量1"));
        assertEquals("ﾗｲﾄｸﾞﾚｰ", raw.get("色1"));
        assertEquals("F-A", raw.get("梱-等1"));
        assertEquals("500", raw.get("原反数量"));
        assertEquals("ﾗｲﾄｸﾞﾚｰ", raw.get("原反色"));
        assertEquals(
                "入庫お願いします。『P000075424』　EL原反は6/10(水)投入します。",
                raw.get("特記事項2"));
    }

    @Test
    void extractFromText_ecowdJr260603_scrambledExport() throws IOException {
        String text =
                Files.readString(
                        Path.of("src/test/resources/tpi-request-forms/ecowd-jr260603.txt"),
                        StandardCharsets.UTF_8);
        Map<String, String> raw =
                RequestFormTpiPdfExtractor.extractFromTextForTest(
                        "ECOWDシート加工注文書（JR260603).pdf", text);

        assertEquals("JR260603", raw.get("依頼Ｎｏ"));
        assertEquals("475", raw.get("数量1"));
        assertEquals("ﾗｲﾄｸﾞﾚｰ", raw.get("色1"));
        assertEquals("F-A", raw.get("梱-等1"));
        assertTrue(raw.get("製品").endsWith("X95"), raw.get("製品"));
        assertEquals("500", raw.get("原反数量"));
        assertEquals("ﾗｲﾄｸﾞﾚｰ", raw.get("原反色"));
        assertTrue(raw.get("特記事項2").contains("入庫お願いします。『P000075424』"));
        assertTrue(raw.get("特記事項2").contains("EL原反は6/10"));
    }

    @Test
    void extractFromText_pnSample() throws IOException {
        String text =
                Files.readString(
                        Path.of("src/test/resources/tpi-request-forms/pn04-03.txt"),
                        StandardCharsets.UTF_8);
        Map<String, String> raw =
                RequestFormTpiPdfExtractor.extractFromTextForTest(
                        "後加工注文書（PN04-03)20260416.pdf", text);

        assertEquals("PN04-03", raw.get("依頼Ｎｏ"));
        assertEquals("2026-04-30", raw.get("希望納期"));
        assertEquals("7C8", raw.get("品名"));
        assertEquals("FEL3002BY05WDLG-EC", raw.get("製品"));
        assertEquals("X000079828", raw.get("契約Ｎｏ"));
        assertEquals("PN", raw.get(RequestFormTpiPdfFieldLayout.META_TPI_LAYOUT));

        Map<String, String> db = RequestFormOriginalExtractor.buildTpiDbDefaultsFromRaw(raw);
        assertEquals("TPI", db.get("加工区分"));
        assertEquals("V（TPI）", db.get("用途"));
    }

    @Test
    void resolveRequestFormTpiPdfDir_konanDefault() {
        GlobalInitSettingTarget.save(FactorySite.KONAN);
        assertEquals(
                AppPaths.DEFAULT_PM_AI_REQUEST_FORM_TPI_PDF_DIR_KONAN,
                AppPaths.resolveRequestFormTpiPdfDir(Map.of()).get().toString());
    }

    @Test
    void resolveRequestFormTpiPdfDir_kokubuEmptyByDefault() {
        GlobalInitSettingTarget.save(FactorySite.KOKUBU);
        assertTrue(AppPaths.resolveRequestFormTpiPdfDir(Map.of()).isEmpty());
    }

    static Stream<Arguments> pdfFixtures() {
        return Stream.of(
                Arguments.of("ECOWD-JR260603.pdf", "JR260603", "ECOWD"),
                Arguments.of("ECOWD-JR260604.pdf", "JR260604", "ECOWD"),
                Arguments.of("ECOWD-JR260604-1.pdf", "JR260604-1", "ECOWD"),
                Arguments.of("ECOWD-JR260605.pdf", "JR260605", "ECOWD"),
                Arguments.of("ECOWD-JR260701.pdf", "JR260701", "ECOWD"),
                Arguments.of("PN04-03.pdf", "PN04-03", "PN"),
                Arguments.of("PN06-01.pdf", "PN06-01", "PN"),
                Arguments.of("PN06-02.pdf", "PN06-02", "PN"));
    }

    @ParameterizedTest
    @MethodSource("pdfFixtures")
    void extractFromPdfFixtures(String fileName, String expectedIraiNo, String expectedLayout)
            throws Exception {
        Path pdf = Path.of("src/test/resources/tpi-request-forms", fileName);
        assertTrue(Files.isRegularFile(pdf), "fixture missing: " + fileName);
        List<Map<String, String>> entries = RequestFormTpiPdfExtractor.extractEntries(pdf.toFile());
        assertEquals(1, entries.size(), fileName);
        Map<String, String> raw = entries.get(0);
        assertEquals(expectedIraiNo, raw.get("依頼Ｎｏ"), fileName);
        assertEquals(expectedLayout, raw.get(RequestFormTpiPdfFieldLayout.META_TPI_LAYOUT), fileName);
        assertEquals(
                RequestFormTpiPdfFieldLayout.META_SOURCE_KIND_TPI_PDF,
                raw.get(RequestFormTpiPdfFieldLayout.META_SOURCE_KIND),
                fileName);
        assertNotNull(raw.get("契約Ｎｏ"), fileName);
        assertFalse(raw.get("契約Ｎｏ").isBlank(), fileName);

        Map<String, String> db = RequestFormOriginalExtractor.buildTpiDbDefaultsFromRaw(raw);
        assertEquals("TPI", db.get("加工区分"), fileName);
        if ("ECOWD".equals(expectedLayout)) {
            assertEquals("JR（屋根）", db.get("用途"), fileName);
        } else {
            assertEquals("V（TPI）", db.get("用途"), fileName);
        }
    }

    @Test
    void bulkScanTpiPdfFolder_whenUncAccessible() throws Exception {
        Path dir = resolveTpiPdfScanDirForVerification();
        if (dir == null) {
            return;
        }
        File[] pdfs =
                dir.toFile()
                        .listFiles(
                                (d, name) ->
                                        name != null && name.toLowerCase().endsWith(".pdf"));
        if (pdfs == null || pdfs.length == 0) {
            return;
        }
        int ok = 0;
        for (File pdf : pdfs) {
            List<Map<String, String>> entries = RequestFormTpiPdfExtractor.extractEntries(pdf);
            assertEquals(1, entries.size(), pdf.getName());
            assertFalse(entries.get(0).get("依頼Ｎｏ").isBlank(), pdf.getName());
            ok++;
        }
        assertTrue(ok >= 8, "expected at least 8 PDFs in TPI folder, got " + ok);
    }

    /** UNC 到達時は実フォルダ、CI/WSL ではテスト用 8 件フィクスチャで一括検証する。 */
    private static Path resolveTpiPdfScanDirForVerification() {
        String uncDir = System.getenv("PM_AI_REQUEST_FORM_TPI_PDF_DIR");
        if (uncDir != null && !uncDir.isBlank()) {
            Path fromEnv = Path.of(uncDir);
            if (Files.isDirectory(fromEnv)) {
                return fromEnv;
            }
        }
        Path factoryDefault = Path.of(AppPaths.DEFAULT_PM_AI_REQUEST_FORM_TPI_PDF_DIR_KONAN);
        if (Files.isDirectory(factoryDefault)) {
            return factoryDefault;
        }
        Path fixtures = Path.of("src/test/resources/tpi-request-forms");
        if (Files.isDirectory(fixtures)) {
            return fixtures;
        }
        return null;
    }
}
