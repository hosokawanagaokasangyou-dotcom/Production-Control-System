package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.time.Instant;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class SourceFileExtensionPolicyTest {

    @TempDir
    Path temp;

    @Test
    void processingPlan_okWhenNewestIsXlsx() throws Exception {
        Path olderCsv = temp.resolve("plan.csv");
        Path newerXlsx = temp.resolve("plan.xlsx");
        Files.writeString(olderCsv, "csv");
        Files.writeString(newerXlsx, "xlsx");
        bumpNewer(newerXlsx, olderCsv);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkProcessingPlanDirectory(temp);

        assertTrue(r.ok());
        assertEquals(newerXlsx.toAbsolutePath().normalize(), r.loadablePath().orElseThrow());
        assertTrue(r.errorMessage().isEmpty());
    }

    @Test
    void processingPlan_errorWhenNewestIsCsvEvenIfOlderXlsxExists() throws Exception {
        Path olderXlsx = temp.resolve("plan.xlsx");
        Path newerCsv = temp.resolve("受注明細表.csv");
        Files.writeString(olderXlsx, "xlsx");
        Files.writeString(newerCsv, "csv");
        bumpNewer(newerCsv, olderXlsx);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkProcessingPlanDirectory(temp);

        assertFalse(r.ok());
        assertTrue(r.loadablePath().isEmpty());
        assertEquals(newerCsv.toAbsolutePath().normalize(), r.newestCandidatePath().orElseThrow());
        assertTrue(r.errorMessage().contains(".xlsx"));
        assertTrue(r.errorMessage().contains(".csv") || r.errorMessage().contains("csv"));
    }

    @Test
    void processingPlan_ignoresXlsmForLoadAndErrorsIfNewest() throws Exception {
        Path xlsm = temp.resolve("plan.xlsm");
        Files.writeString(xlsm, "xlsm");

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkProcessingPlanDirectory(temp);

        assertFalse(r.ok());
        assertTrue(r.loadablePath().isEmpty());
    }

    @Test
    void dailyReport_okWhenNewestMatchingPrefixIsCsv() throws Exception {
        Path older = temp.resolve("加工日報発行問合せ_20260101.csv");
        Path newer = temp.resolve("加工日報発行問合せ_20260201.csv");
        Files.writeString(older, "a");
        Files.writeString(newer, "b");
        bumpNewer(newer, older);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkDailyReportDirectory(temp);

        assertTrue(r.ok());
        assertEquals(newer.toAbsolutePath().normalize(), r.loadablePath().orElseThrow());
    }

    @Test
    void dailyReport_errorWhenNewestMatchingPrefixIsXlsx() throws Exception {
        Path olderCsv = temp.resolve("加工日報発行問合せ_20260101.csv");
        Path newerXlsx = temp.resolve("加工日報発行問合せ_20260201.xlsx");
        Files.writeString(olderCsv, "csv");
        Files.writeString(newerXlsx, "xlsx");
        bumpNewer(newerXlsx, olderCsv);

        SourceFileExtensionPolicy.Result r =
                SourceFileExtensionPolicy.checkDailyReportDirectory(temp);

        assertFalse(r.ok());
        assertTrue(r.loadablePath().isEmpty());
        assertEquals(newerXlsx.toAbsolutePath().normalize(), r.newestCandidatePath().orElseThrow());
        assertTrue(r.errorMessage().contains(".csv"));
    }

    private static void bumpNewer(Path newer, Path older) throws Exception {
        long base = Instant.now().toEpochMilli();
        Files.setLastModifiedTime(older, FileTime.fromMillis(base));
        Files.setLastModifiedTime(newer, FileTime.fromMillis(base + 60_000L));
    }
}
