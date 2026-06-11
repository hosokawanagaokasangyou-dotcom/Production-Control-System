package jp.co.pm.ai.desktop.io;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class RdpLaunchSharedPathValidatorTest {

    @Test
    void isUncPath_detectsUncPrefix() {
        assertTrue(RdpLaunchSharedPathValidator.isUncPath("\\\\192.168.0.101\\share\\a.exe"));
        assertTrue(RdpLaunchSharedPathValidator.isUncPath("//192.168.0.101/share/a.exe"));
        assertFalse(RdpLaunchSharedPathValidator.isUncPath("C:\\Windows\\notepad.exe"));
        assertFalse(RdpLaunchSharedPathValidator.isUncPath("Z:\\portable\\Aladdin_RPA_Studio.exe"));
    }

    @Test
    void validateProgramPath_emptyProgram() {
        RdpLaunchSharedPathValidator.Result result =
                RdpLaunchSharedPathValidator.validateProgramPath("  ");
        assertFalse(result.ok());
        assertTrue(result.issues().get(0).detail().contains("空"));
    }

    @Test
    void validateScenarioArguments_noScenario_isOk() {
        RdpLaunchSharedPathValidator.Result result =
                RdpLaunchSharedPathValidator.validateScenarioArguments("");
        assertTrue(result.ok());
        assertTrue(result.checkedPaths().isEmpty());
    }

    @Test
    void formatExistenceOkMessage_listsCheckedPaths() {
        RdpLaunchSharedPathValidator.Result result =
                new RdpLaunchSharedPathValidator.Result(
                        true,
                        List.of(),
                        List.of("\\\\server\\share\\a.ardrpa"));
        assertTrue(result.formatExistenceOkMessage().contains("a.ardrpa"));
    }

    @Test
    void validateBeforeConnect_skipsLocalPaths(@TempDir Path tmp) throws Exception {
        Path localExe = tmp.resolve("Aladdin_RPA_Studio.exe");
        Files.writeString(localExe, "stub");
        Path localScenario = tmp.resolve("工程別.ardrpa");
        Files.writeString(localScenario, "stub");

        RdpLaunchSharedPathValidator.Result result =
                RdpLaunchSharedPathValidator.validateBeforeConnect(
                        localExe.toString(),
                        "--scenario " + localScenario,
                        null);

        assertTrue(result.ok(), "issues=" + result.issues());
    }

    @Test
    void validateBeforeConnect_reportsMissingUncScenario() {
        String program = "C:\\portable\\Aladdin_RPA_Studio.exe";
        String scenario =
                "\\\\127.0.0.99\\share\\002  加工G\\scenarios\\工程別生産計画問い合わせ.ardrpa";
        RdpLaunchSharedPathValidator.Result result =
                RdpLaunchSharedPathValidator.validateBeforeConnect(
                        program, "--scenario " + scenario, null);

        assertFalse(result.ok());
        assertTrue(result.issues().stream().anyMatch(i -> i.label().contains("シナリオ")));
        assertTrue(result.formatBlockingMessage().contains("工程別生産計画問い合わせ.ardrpa"));
    }

    @Test
    void validateBeforeConnect_repairsCollapsed002BeforeCheck() {
        String wrongScenario =
                "\\\\127.0.0.99\\share\\002 加工G\\scenarios\\missing.ardrpa";
        RdpLaunchSharedPathValidator.Result result =
                RdpLaunchSharedPathValidator.validateBeforeConnect(
                        "C:\\app.exe", "--scenario " + wrongScenario, null);

        assertFalse(result.ok());
        assertTrue(
                result.issues().stream()
                        .anyMatch(i -> i.path().contains("002  加工G")),
                "issues=" + result.issues());
    }
}
