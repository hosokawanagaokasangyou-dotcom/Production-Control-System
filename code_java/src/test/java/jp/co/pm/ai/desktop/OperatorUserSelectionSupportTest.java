package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javafx.scene.control.Dialog;
import javafx.stage.Stage;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;

class OperatorUserSelectionSupportTest {

    private String priorDesktopHomeDirName;
    private String priorTestStore;

    @BeforeEach
    void activateRdpAppHome() {
        priorDesktopHomeDirName = AppPaths.desktopAppHomeDirName();
        priorTestStore = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        RemoteDesktopStandaloneBootstrap.activate();
    }

    @AfterEach
    void restoreDesktopHomeDirName() throws Exception {
        AppPaths.setDesktopAppHomeDirName(priorDesktopHomeDirName);
        FactoryOperatorUserStore.resetStoreForTests();
        if (priorTestStore != null) {
            System.setProperty("pm.ai.test.factoryOperatorUserStore", priorTestStore);
        } else {
            System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        }
    }

    @Test
    void rdpScopeLabels_omitFactorySiteName() {
        assertEquals("", OperatorUserSelectionSupport.operatorSelectionScopeLabel(FactorySite.RDP_LAUNCHER));
        assertEquals(
                "部署 製造1課",
                OperatorUserSelectionSupport.operatorSelectionScopeSuffix(
                        FactorySite.RDP_LAUNCHER, "製造1課"));
        assertEquals(
                "",
                OperatorUserSelectionSupport.operatorSelectionScopeSuffix(FactorySite.RDP_LAUNCHER, ""));
        assertEquals(
                "",
                OperatorUserSelectionSupport.operatorSelectionScopeLabel(FactorySite.KONAN));
    }

    @Test
    void remoteDesktopShellFxml_hasNoDepartmentToolbarButton() throws Exception {
        Path fxml =
                Path.of("src/main/resources/jp/co/pm/ai/desktop/fxml/RemoteDesktopShell.fxml");
        String text = Files.readString(fxml, StandardCharsets.UTF_8);
        assertTrue(text.contains("changeSessionOperatorToolbarButton"));
        assertFalse(text.contains("changeSessionDepartmentToolbarButton"));
        assertFalse(text.contains("部署を変更"));
    }

    @Test
    void rdpOperatorChange_cancelDepartmentRestoresSession(@TempDir Path tmp) throws Exception {
        configureRdpStore(tmp, "製造1課", "製造2課");
        FactoryOperatorUserStore.selectSessionRdpDepartment("製造1課");
        FactoryOperatorUserStore.addName(FactorySite.RDP_LAUNCHER, "山田");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.RDP_LAUNCHER, "山田");

        OperatorUserSelectionSupport.performRdpOperatorChange(
                new StubHost(),
                FactorySite.RDP_LAUNCHER,
                "山田",
                "製造1課",
                (host, departments) -> Optional.empty(),
                (host, factory) -> Optional.of("別ユーザー"));

        assertEquals("山田", FactoryOperatorUserStore.sessionOperatorName());
        assertEquals("製造1課", FactoryOperatorUserStore.sessionRdpDepartmentKey());
    }

    @Test
    void rdpOperatorChange_cancelOperatorRestoresSession(@TempDir Path tmp) throws Exception {
        configureRdpStore(tmp, "製造1課", "製造2課");
        FactoryOperatorUserStore.selectSessionRdpDepartment("製造1課");
        FactoryOperatorUserStore.addName(FactorySite.RDP_LAUNCHER, "山田");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.RDP_LAUNCHER, "山田");

        OperatorUserSelectionSupport.performRdpOperatorChange(
                new StubHost(),
                FactorySite.RDP_LAUNCHER,
                "山田",
                "製造1課",
                (host, departments) -> Optional.of("製造2課"),
                (host, factory) -> Optional.empty());

        assertEquals("山田", FactoryOperatorUserStore.sessionOperatorName());
        assertEquals("製造1課", FactoryOperatorUserStore.sessionRdpDepartmentKey());
    }

    @Test
    void rdpOperatorChange_departmentBeforeOperatorAndScopesUserList(@TempDir Path tmp)
            throws Exception {
        configureRdpStore(tmp, "製造1課", "製造2課");
        FactoryOperatorUserStore.selectSessionRdpDepartment("製造1課");
        FactoryOperatorUserStore.setAdminRdpDepartmentContext("製造1課");
        FactoryOperatorUserStore.addName(FactorySite.RDP_LAUNCHER, "製造1専用");
        FactoryOperatorUserStore.setAdminRdpDepartmentContext("製造2課");
        FactoryOperatorUserStore.selectSessionRdpDepartment("製造2課");
        FactoryOperatorUserStore.addName(FactorySite.RDP_LAUNCHER, "製造2専用");
        FactoryOperatorUserStore.setAdminRdpDepartmentContext("");
        FactoryOperatorUserStore.selectSessionRdpDepartment("製造1課");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.RDP_LAUNCHER, "製造1専用");

        List<String> promptOrder = new ArrayList<>();

        OperatorUserSelectionSupport.performRdpOperatorChange(
                new StubHost(),
                FactorySite.RDP_LAUNCHER,
                "製造1専用",
                "製造1課",
                (host, departments) -> {
                    promptOrder.add("department");
                    return Optional.of("製造2課");
                },
                (host, factory) -> {
                    promptOrder.add("operator");
                    assertEquals("製造2課", FactoryOperatorUserStore.sessionRdpDepartmentKey());
                    return Optional.of(FactoryOperatorUserStore.GUEST_OPERATOR_NAME);
                });

        assertEquals(List.of("department", "operator"), promptOrder);
        assertEquals(FactoryOperatorUserStore.GUEST_OPERATOR_NAME, FactoryOperatorUserStore.sessionOperatorName());
        assertEquals("製造2課", FactoryOperatorUserStore.sessionRdpDepartmentKey());
    }

    @Test
    void mainRunTabFxml_hasChangeSessionOperatorButton() throws Exception {
        Path fxml = Path.of("src/main/resources/jp/co/pm/ai/desktop/fxml/MainRunTab.fxml");
        String text = Files.readString(fxml, StandardCharsets.UTF_8);
        assertTrue(text.contains("changeSessionOperatorButton"));
        assertTrue(text.contains("操作者変更"));
    }

    @Test
    void productionOperatorChange_cancelRestoresSession(@TempDir Path tmp) throws Exception {
        Path customDir = tmp.resolve("factory-user-store");
        System.setProperty(
                "pm.ai.test.factoryOperatorUserStore",
                customDir.resolve("operators.bin").toString());
        System.setProperty(
                "pm.ai.test.factoryOperatorLastSelectedDir", tmp.resolve("last-selected").toString());
        FactoryOperatorUserStore.resetStoreForTests();
        FactoryOperatorUserStore.configureForCurrentApp(Map.of(), FactorySite.KONAN);
        FactoryOperatorUserStore.addName(FactorySite.KONAN, "operatorA");
        FactoryOperatorUserStore.addName(FactorySite.KONAN, "operatorB");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "operatorA");

        OperatorUserSelectionSupport.performProductionOperatorChange(
                new StubHost(),
                FactorySite.KONAN,
                "operatorA",
                (host, factory) -> Optional.empty());

        assertEquals("operatorA", FactoryOperatorUserStore.sessionOperatorName());
    }

    private static void configureRdpStore(Path tmp, String... departments) throws Exception {
        Path customDir = tmp.resolve("rdp-user-store");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR, customDir.toString());
        System.setProperty(
                "pm.ai.test.factoryOperatorUserStore",
                customDir.resolve("operators.bin").toString());
        System.setProperty(
                "pm.ai.test.factoryOperatorLastSelectedDir", tmp.resolve("last-selected").toString());
        FactoryOperatorUserStore.resetStoreForTests();
        FactoryOperatorUserStore.configureForStandaloneLocalOnly(ui, FactorySite.RDP_LAUNCHER);
        for (String dept : departments) {
            FactoryOperatorUserStore.addRdpDepartment(dept);
        }
    }

    private static final class StubHost implements DesktopShellHost {
        @Override
        public Map<String, String> snapshotUiEnv() {
            return Map.of();
        }

        @Override
        public void updateEnvTabValue(String envKey, String value) {}

        @Override
        public void appendLog(String line) {}

        @Override
        public Stage primaryStageForDialogs() {
            return null;
        }

        @Override
        public void requireOperatorSelectionForFactory(FactorySite site, boolean startup) {}

        @Override
        public void changeSessionOperator(FactorySite site) {}

        @Override
        public void refreshOperatorUserPresentation() {}

        @Override
        public void prepareDialogForMainTheme(Dialog<?> dialog) {}

        @Override
        public void showWarningDialog(String title, String message) {}

        @Override
        public void showInformationDialog(String title, String message) {}
    }
}
