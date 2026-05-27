package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class ProcessOwnedLockFilesTest {

    @TempDir
    Path temp;

    @Test
    void releaseIfOwnedByCurrentProcessDeletesOnlyMatchingPid() throws Exception {
        Path ownLock = temp.resolve("own.export.lock");
        Path otherLock = temp.resolve("other.export.lock");
        long pid = ProcessHandle.current().pid();
        Files.writeString(
                ownLock,
                "version=1\npid=" + pid + "\n",
                StandardCharsets.UTF_8);
        Files.writeString(otherLock, "version=1\npid=999999\n", StandardCharsets.UTF_8);

        ProcessOwnedLockFiles.register(ownLock);
        ProcessOwnedLockFiles.register(otherLock);

        ProcessOwnedLockFiles.releaseAllOwnedQuietly();

        assertFalse(Files.isRegularFile(ownLock));
        assertTrue(Files.isRegularFile(otherLock));
    }

    @Test
    void releaseIfOwnedByCurrentProcessSkipsForeignPid() throws Exception {
        Path lock = temp.resolve("foreign.save.lock");
        Files.writeString(lock, "version=1\npid=424242\n", StandardCharsets.UTF_8);

        assertFalse(ProcessOwnedLockFiles.releaseIfOwnedByCurrentProcess(lock));
        assertTrue(Files.isRegularFile(lock));
    }
}
