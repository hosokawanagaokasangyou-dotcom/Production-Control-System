package jp.co.pm.ai.planning.stage2.source;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.util.concurrent.atomic.AtomicBoolean;

import org.junit.jupiter.api.Test;

class Stage1SourceBundleCompletionGateTest {

    @Test
    void invalidateBeforeStage1_blocksStartWhenStrictDeleteFails() {
        var result =
                Stage1SourceBundleCompletionGate.invalidateBeforeStage1(
                        () -> {
                            throw new IOException("delete failed");
                        });

        assertFalse(result.completionAllowed());
    }

    @Test
    void persist_blocksCompletionWhenRequiredBundleSaveFails() {
        var result =
                Stage1SourceBundleCompletionGate.persist(
                        true,
                        true,
                        () -> {},
                        () -> {
                            throw new IOException("save failed");
                        });

        assertFalse(result.completionAllowed());
    }

    @Test
    void persist_todayDispatchOffInvalidatesOldBundleAndAllowsCompletion() {
        AtomicBoolean invalidated = new AtomicBoolean();

        var result =
                Stage1SourceBundleCompletionGate.persist(
                        false,
                        false,
                        () -> invalidated.set(true),
                        () -> {});

        assertTrue(result.completionAllowed());
        assertTrue(invalidated.get());
    }
}
