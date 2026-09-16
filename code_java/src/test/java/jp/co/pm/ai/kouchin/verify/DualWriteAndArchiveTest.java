package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class DualWriteAndArchiveTest {

    @TempDir Path tmp;

    @Test
    void dualWriteOneFailureStillSucceedsOther() throws Exception {
        Path okDir = tmp.resolve("ok");
        Path badParent = tmp.resolve("file-not-dir");
        Files.writeString(badParent, "x");
        Path bad = badParent.resolve("out.xlsx");
        Path ok = okDir.resolve("out.xlsx");
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("s");
            DualWriteFiles.WriteOutcome o = DualWriteFiles.writeWorkbook(wb, List.of(ok, bad), Map.of());
            assertEquals(1, o.succeeded().size());
            assertEquals(ok, o.succeeded().get(0));
            assertTrue(o.anyFailed());
            assertFalse(o.allFailed());
        }
    }

    @Test
    void bothFailure() throws Exception {
        Path bad1 = tmp.resolve("a").resolve("x.xlsx");
        Path bad2 = tmp.resolve("b").resolve("x.xlsx");
        Files.writeString(tmp.resolve("a"), "x");
        Files.writeString(tmp.resolve("b"), "y");
        DualWriteFiles.WriteOutcome o = DualWriteFiles.writeBytes(new byte[] {1}, List.of(bad1, bad2));
        assertTrue(o.allFailed());
    }

    @Test
    void archiveKeepIsPerDirectory() throws Exception {
        Path d1 = tmp.resolve("k1");
        Path d2 = tmp.resolve("k2");
        Files.createDirectories(d1);
        Files.createDirectories(d2);
        Path keep1 = d1.resolve("検証結果_国分工場_1.xlsx");
        Path old1 = d1.resolve("検証結果_国分工場_0.xlsx");
        Path keep2 = d2.resolve("検証結果_国分工場_1.xlsx");
        Path old2 = d2.resolve("検証結果_国分工場_0.xlsx");
        Files.write(keep1, new byte[] {1});
        Files.write(old1, new byte[] {2});
        Files.write(keep2, new byte[] {3});
        Files.write(old2, new byte[] {4});
        ResultArchive.archiveOldVerifyResults(d1, List.of(keep1));
        ResultArchive.archiveOldVerifyResults(d2, List.of(keep2));
        assertTrue(Files.exists(keep1));
        assertTrue(Files.exists(keep2));
        assertFalse(Files.exists(old1));
        assertTrue(Files.exists(d1.resolve(ResultArchive.PAST_VERIFY).resolve("検証結果_国分工場_0.xlsx")));
        assertTrue(Files.exists(old2) || Files.exists(d2.resolve(ResultArchive.PAST_VERIFY).resolve("検証結果_国分工場_0.xlsx")));
        assertTrue(Files.exists(d2.resolve(ResultArchive.PAST_VERIFY).resolve("検証結果_国分工場_0.xlsx")));
        assertFalse(Files.exists(d1.resolve(ResultArchive.PAST_VERIFY).resolve("検証結果_国分工場_1.xlsx")));
    }
}
