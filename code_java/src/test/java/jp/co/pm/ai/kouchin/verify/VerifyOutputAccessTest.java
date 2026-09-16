package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class VerifyOutputAccessTest {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("出力フォルダへ書けなければ検証不可")
    void allUnwritableWhenEveryDirFails() throws Exception {
        Path ok = tmp.resolve("ok");
        Files.createDirectories(ok);
        Path notDir = tmp.resolve("file");
        Files.writeString(notDir, "x");
        assertTrue(VerifyOutputAccess.canWriteDir(ok));
        assertFalse(VerifyOutputAccess.canWriteDir(notDir));
        assertFalse(VerifyOutputAccess.canWriteDir(null));
        assertTrue(VerifyOutputAccess.anyOutputWritable(List.of(notDir, ok)));
        assertFalse(VerifyOutputAccess.anyOutputWritable(List.of(notDir)));
        assertFalse(VerifyOutputAccess.anyOutputWritable(List.of()));
        assertNull(VerifyOutputAccess.writeBlockReason(List.of(ok)));
        assertEquals("結果Excelを書き込めません", VerifyOutputAccess.writeBlockReason(List.of(notDir)));
    }
}
