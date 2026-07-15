package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;

class MainShellStage2GuardConcurrencyContractTest {

    @Test
    void guardBusyBlocksRelatedStartsAndRevalidatesSnapshotBeforeHandoff() throws Exception {
        String source =
                Files.readString(
                        Path.of(
                                "src/main/java/jp/co/pm/ai/desktop/MainShellController.java"));

        assertTrue(source.contains("Stage2SourceGuardCoordinator"));
        assertTrue(source.contains("blockIfStage2SourceGuardBusy(\"段階1\")"));
        assertTrue(source.contains("stage2SourceGuardCoordinator.isRunning()"));
        assertTrue(source.contains("captureStage2SourceGuardSnapshot"));
        assertTrue(source.contains("startedSnapshot.matches(currentSnapshot)"));
        assertTrue(source.contains("ガード中に実行条件が変更されました"));
        assertTrue(source.contains("stage2SourceGuardRunHandoff = true"));
        assertTrue(source.contains("stage2SourceGuardRunHandoff = false"));
        assertTrue(source.contains("applyRunTabGating()"));
        assertTrue(source.contains("Task<List<Stage1SourcePairMatcher.MatchedPair>>"));
        assertTrue(source.contains("\"today-dispatch-source-scan\""));
        assertTrue(source.contains("startStage1AfterStrictBundleInvalidation"));
        assertTrue(source.contains("Stage1SourceBundleCompletionGate.invalidateBeforeStage1"));
        assertTrue(source.contains("Stage1SourceBundleCompletionGate.persist"));
        assertTrue(source.contains("if (!bundleResult.completionAllowed())"));
    }
}
