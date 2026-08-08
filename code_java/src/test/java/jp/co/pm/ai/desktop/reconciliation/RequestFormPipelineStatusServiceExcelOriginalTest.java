package jp.co.pm.ai.desktop.reconciliation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferCoverageCheck.CoverageResult;
import jp.co.pm.ai.desktop.reconciliation.RawInputDateCrossSourceCheck.CrossSourceResult;
import jp.co.pm.ai.desktop.reconciliation.RawInputDateCrossSourceCheck.SourceValues;

class RequestFormPipelineStatusServiceExcelOriginalTest {

    @Test
    void loadExcelParseEntriesFromCacheFile_readsExcelSchema(@TempDir Path temp) throws Exception {
        Path cacheJson = temp.resolve("sample.json");
        String payload =
                """
                {
                  "schemaVersion": "%s",
                  "cachedAtMillis": %d,
                  "entries": [
                    {
                      "依頼Ｎｏ": "Y8-99",
                      "原本ファイル名": "Y-8月（2026年）加工依頼書（国分Y）.xlsm"
                    }
                  ]
                }
                """
                        .formatted(
                                RequestFormSourceCache.PARSE_SCHEMA_VERSION,
                                System.currentTimeMillis());
        Files.writeString(cacheJson, payload);

        Optional<List<Map<String, String>>> entries =
                RequestFormSourceCache.loadExcelParseEntriesFromCacheFile(cacheJson.toFile());

        assertTrue(entries.isPresent());
        assertEquals("Y8-99", entries.get().getFirst().get("依頼Ｎｏ"));
    }

    @Test
    void resolveLinkedExcelOriginalRaw_findsEntryInParseCache(@TempDir Path temp) throws Exception {
        File parseRoot = temp.resolve("preview_cache").toFile();
        File parseDir = RequestFormSourceCache.parseDir(parseRoot);
        String payload =
                """
                {
                  "schemaVersion": "%s",
                  "cachedAtMillis": %d,
                  "entries": [
                    {
                      "依頼Ｎｏ": "Y8-99",
                      "ユーザー": "フカサワ",
                      "原本ファイル名": "Y-8月（2026年）加工依頼書（国分Y）.xlsm"
                    }
                  ]
                }
                """
                        .formatted(
                                RequestFormSourceCache.PARSE_SCHEMA_VERSION,
                                System.currentTimeMillis());
        Files.writeString(
                new File(parseDir, "Y-8月（2026年）加工依頼書（国分Y）.json").toPath(), payload);

        Optional<Map<String, String>> linked =
                RequestFormPipelineStatusService.resolveLinkedExcelOriginalRaw(
                        "Y8-99", Map.of(), parseRoot, new ArrayList<>());

        assertTrue(linked.isPresent());
        assertEquals("Y8-99", linked.get().get("依頼Ｎｏ"));
        assertEquals(
                "Y-8月（2026年）加工依頼書（国分Y）.xlsm",
                linked.get().get("_sourceFileName"));
    }

    @Test
    void appendExcelParseCacheFallback_addsMissingKeys(@TempDir Path temp) throws Exception {
        File parseRoot = temp.resolve("preview_cache").toFile();
        File parseDir = RequestFormSourceCache.parseDir(parseRoot);
        String payload =
                """
                {
                  "schemaVersion": "%s",
                  "cachedAtMillis": %d,
                  "entries": [
                    {
                      "依頼Ｎｏ": "Y8-99",
                      "原本ファイル名": "Y-8月（2026年）加工依頼書（国分Y）.xlsm"
                    }
                  ]
                }
                """
                        .formatted(
                                RequestFormSourceCache.PARSE_SCHEMA_VERSION,
                                System.currentTimeMillis());
        Files.writeString(
                new File(parseDir, "Y-8月（2026年）加工依頼書（国分Y）.json").toPath(), payload);

        List<Map<String, String>> rawRequests = new ArrayList<>();
        Set<String> keys = new HashSet<>();

        RequestFormPipelineStatusService.appendExcelParseCacheFallback(
                rawRequests, keys, parseRoot);

        assertEquals(1, rawRequests.size());
        assertEquals("Y8-99", rawRequests.getFirst().get("依頼Ｎｏ"));
        assertTrue(keys.contains(JuchuTransferValueNormalizer.normalizeKey("Y8-99")));
    }

    @Test
    void issueCheck_noOriginal_falseWhenOriginalPresent() {
        CoverageResult coverage = new CoverageResult(true, 4, 4, 100.0, List.of());
        CrossSourceResult cross =
                new CrossSourceResult(
                        RawInputDateCrossSourceCheck.STATUS_MATCH,
                        new SourceValues("2026/8/25", "2026/8/25", "", ""),
                        "");
        RequestFormPipelineStatusService.PipelineStatusRow row =
                new RequestFormPipelineStatusService.PipelineStatusRow(
                        "Y8-99",
                        "Y-8月（2026年）加工依頼書（国分Y）.xlsm",
                        true,
                        "フカサワ",
                        true,
                        coverage.rateDisplay(),
                        coverage.ratePercent(),
                        coverage.mismatchCount(),
                        "190-999T",
                        "",
                        true,
                        List.of(),
                        coverage,
                        List.of(),
                        LocalDate.of(2026, 8, 21),
                        "2026/8/21",
                        "担当",
                        LocalDate.of(2026, 8, 28),
                        "2026/8/28",
                        "2026/8/25",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        cross);

        assertFalse(
                RequestFormPipelineIssueCheck.detect(row, true)
                        .contains(RequestFormPipelineIssueCheck.IssueKind.NO_ORIGINAL));
    }

    @Test
    void resolveLinkedTpiPdfRaw_skippedForKokubuFactory(@TempDir Path temp) throws Exception {
        File parseRoot = temp.resolve("preview_cache").toFile();
        File parseDir = RequestFormSourceCache.parseDir(parseRoot);
        Files.createDirectories(parseDir.toPath());
        Map<String, String> kokubuEnv =
                Map.of(AppPaths.KEY_PM_AI_FACTORY_SITE, FactorySite.KOKUBU.name());

        assertTrue(
                RequestFormPipelineStatusService.resolveLinkedTpiPdfRaw(
                                "GB60604", kokubuEnv, parseRoot, new ArrayList<>())
                        .isEmpty());
    }
}
