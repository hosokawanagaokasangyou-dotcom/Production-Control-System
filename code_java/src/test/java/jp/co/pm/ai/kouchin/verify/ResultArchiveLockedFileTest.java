package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class ResultArchiveLockedFileTest {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("過去フォルダを作れなくても新規結果の退避処理は例外にしない")
    void archiveDoesNotThrowWhenPastFolderBlocked() throws Exception {
        Path keep = tmp.resolve("検証結果_湖南工場_new.xlsx");
        Path old = tmp.resolve("検証結果_湖南工場_old.xlsx");
        Files.write(keep, new byte[] {1});
        Files.write(old, new byte[] {2});
        Files.writeString(tmp.resolve(ResultArchive.PAST_VERIFY), "not-a-dir");
        assertDoesNotThrow(() -> ResultArchive.archiveOldVerifyResults(tmp, List.of(keep)));
        assertTrue(Files.isRegularFile(keep));
        assertTrue(Files.isRegularFile(old));
    }
}
