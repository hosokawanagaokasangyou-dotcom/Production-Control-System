package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class KouchinTorayCsvDropSupportTest {

    @TempDir Path tmp;

    @Test
    void copiesCsvAndOverwritesSameName() throws Exception {
        Path dest = tmp.resolve("toray");
        Path src = tmp.resolve("RVSHEET202609.csv");
        Files.writeString(src, "a", StandardCharsets.UTF_8);
        var o = KouchinTorayCsvDropSupport.copyCsvFiles(List.of(src), dest, null);
        assertTrue(o.copiedAny());
        assertEquals(1, o.copied().size());
        Files.writeString(src, "b", StandardCharsets.UTF_8);
        var o2 = KouchinTorayCsvDropSupport.copyCsvFiles(List.of(src), dest, null);
        assertEquals("b", Files.readString(dest.resolve("RVSHEET202609.csv")));
        assertTrue(o2.warnings().isEmpty() || o2.warnings().stream().noneMatch(s -> s.contains("拒否")));
    }

    @Test
    void rejectsNonCsv() throws Exception {
        Path dest = tmp.resolve("toray");
        Path xlsx = tmp.resolve("a.xlsx");
        Files.write(xlsx, new byte[] {1});
        var o = KouchinTorayCsvDropSupport.copyCsvFiles(List.of(xlsx), dest, null);
        assertFalse(o.copiedAny());
        assertTrue(o.warnings().stream().anyMatch(s -> s.contains("csv以外")));
    }

    @Test
    void folderDropsOnlyImmediateCsvNotRecursive() throws Exception {
        Path dest = tmp.resolve("toray");
        Path dir = tmp.resolve("drop");
        Files.createDirectories(dir.resolve("sub"));
        Files.writeString(dir.resolve("RVSHEET202609.csv"), "a", StandardCharsets.UTF_8);
        Files.writeString(dir.resolve("sub").resolve("inner.csv"), "x", StandardCharsets.UTF_8);
        var o = KouchinTorayCsvDropSupport.copyCsvFiles(List.of(dir), dest, null);
        assertEquals(1, o.copied().size());
        assertFalse(Files.exists(dest.resolve("inner.csv")));
    }

    @Test
    void sourceDeletedDuringCopyIsErrorNotDetectionUpdate() throws Exception {
        Path dest = tmp.resolve("toray");
        Path missing = tmp.resolve("gone.csv");
        var o = KouchinTorayCsvDropSupport.copyCsvFiles(List.of(missing), dest, null);
        assertFalse(o.copiedAny());
    }

    @Test
    void nonRvsheetNameWarnsButCopies() throws Exception {
        Path dest = tmp.resolve("toray");
        Path src = tmp.resolve("outlook-attach.csv");
        Files.writeString(src, "x", StandardCharsets.UTF_8);
        var o = KouchinTorayCsvDropSupport.copyCsvFiles(List.of(src), dest, null);
        assertTrue(o.copiedAny());
        assertTrue(o.warnings().stream().anyMatch(s -> s.contains("RVSHEET")));
    }
}
