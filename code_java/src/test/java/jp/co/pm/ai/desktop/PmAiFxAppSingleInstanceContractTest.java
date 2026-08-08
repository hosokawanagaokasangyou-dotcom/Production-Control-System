package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;

class PmAiFxAppSingleInstanceContractTest {

    @Test
    void mainSourceWiresSingleInstanceGuardEarly() throws Exception {
        Path src =
                Path.of("src/main/java/jp/co/pm/ai/desktop/PmAiFxApp.java");
        String text = Files.readString(src, StandardCharsets.UTF_8);
        assertTrue(text.contains("SingleInstanceGuard"));
        assertTrue(text.contains("tryAcquire"));
        assertTrue(text.contains("Role.SECONDARY"));
        assertTrue(text.contains("setOnActivateRequest"));
        assertTrue(text.contains("bringPrimaryStageToFront"));
    }
}
