package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.AppPaths;

class KouchinOutputDirsTest {

    @Test
    void resolveAllAlwaysIncludesKokubuAndKonan() {
        List<Path> dirs = KouchinOutputDirs.resolveAll(Map.of());
        assertEquals(2, dirs.size());
        assertEquals(
                Path.of(AppPaths.DEFAULT_KOUCHIN_BASE_DIR).toAbsolutePath().normalize(),
                dirs.get(0));
        assertEquals(
                Path.of(AppPaths.DEFAULT_KOUCHIN_KONAN_OUTPUT_DIR).toAbsolutePath().normalize(),
                dirs.get(1));
    }

    @Test
    void extraOutputDirIsThirdCopy() {
        Path extra = Path.of("C:\\tmp\\kouchin-extra").toAbsolutePath().normalize();
        List<Path> dirs =
                KouchinOutputDirs.resolveAll(
                        Map.of(AppPaths.KEY_PM_AI_KOUCHIN_OUTPUT_DIR, extra.toString()));
        assertEquals(3, dirs.size());
        assertEquals(extra, dirs.get(2));
    }
}

class KouchinDualWriteSmokeTest {

    @Test
    void dualWriteContinuesWhenOneTargetFails(@TempDir Path tmp) throws IOException {
        Path okDir = tmp.resolve("ok");
        Path failFile = tmp.resolve("missing-parent-as-file");
        Files.writeString(failFile, "not-a-dir");
        Path ok = okDir.resolve("検証結果_国分工場_1.xlsx");
        Path fail = failFile.resolve("検証結果_国分工場_1.xlsx");
        DualWriteFiles.WriteOutcome out =
                DualWriteFiles.writeBytes("abc".getBytes(), List.of(ok, fail));
        assertEquals(List.of(ok), out.succeeded());
        assertFalse(out.failures().isEmpty());
        assertTrue(Files.isRegularFile(ok));
    }

    @Test
    void archiveDoesNotMoveKeepInSameDir(@TempDir Path dir) throws IOException {
        Path keep = dir.resolve("検証結果_国分工場_new.xlsx");
        Path old = dir.resolve("検証結果_国分工場_old.xlsx");
        Files.writeString(keep, "new");
        Files.writeString(old, "old");
        ResultArchive.archiveOldVerifyResults(dir, List.of(keep));
        assertTrue(Files.isRegularFile(keep));
        assertFalse(Files.isRegularFile(old));
        assertTrue(Files.isRegularFile(dir.resolve("過去検証結果").resolve("検証結果_国分工場_old.xlsx")));
    }

    @Test
    void archiveIgnoresKeepFromOtherDirectory(@TempDir Path tmp) throws IOException {
        Path a = tmp.resolve("a");
        Path b = tmp.resolve("b");
        Files.createDirectories(a);
        Files.createDirectories(b);
        Path justWritten = a.resolve("検証結果_国分工場_x.xlsx");
        Files.writeString(justWritten, "x");
        Path keepFromB = b.resolve("検証結果_国分工場_x.xlsx");
        Files.writeString(keepFromB, "other");
        ResultArchive.archiveOldVerifyResults(a, List.of(keepFromB));
        assertFalse(Files.isRegularFile(justWritten));
        assertTrue(Files.isRegularFile(a.resolve("過去検証結果").resolve("検証結果_国分工場_x.xlsx")));
    }
}
