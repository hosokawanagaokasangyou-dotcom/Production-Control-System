package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.Map;

import javafx.application.Platform;
import javafx.scene.control.Button;
import javafx.scene.control.Label;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.FactoryOperatorUserStore;
import jp.co.pm.ai.desktop.config.FactorySite;

class ShellFactoryOperatorToolbarTest {

    @TempDir Path tmp;

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    @BeforeEach
    void isolateStore() throws Exception {
        System.setProperty(
                "pm.ai.test.factoryOperatorUserStore", tmp.resolve("operators.bin").toString());
        System.setProperty(
                "pm.ai.test.factoryOperatorLastSelectedDir", tmp.resolve("last-selected").toString());
        FactoryOperatorUserStore.resetStoreForTests();
        FactoryOperatorUserStore.configureForCurrentApp(Map.of(), FactorySite.KONAN);
    }

    @AfterEach
    void clearProperty() throws Exception {
        FactoryOperatorUserStore.resetStoreForTests();
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        System.clearProperty("pm.ai.test.factoryOperatorLastSelectedDir");
    }

    @Test
    void guestSession_keepsOperatorChangeEnabledAndDisablesPinChange() throws Exception {
        FactoryOperatorUserStore.selectSessionOperator(
                FactorySite.KONAN, FactoryOperatorUserStore.GUEST_OPERATOR_NAME);
        Button changeOperator = new Button();
        Button changePin = new Button();
        ShellFactoryOperatorToolbar toolbar =
                new ShellFactoryOperatorToolbar(
                        null, null, null, null, new Label(), changeOperator, changePin);

        toolbar.refreshOperatorUserLabel();

        assertFalse(
                changeOperator.isDisable(),
                "ゲストでも登録操作者へ切り替えられること");
        assertTrue(changePin.isDisable(), "ゲストは PIN 変更不可");
    }

    @Test
    void registeredOperator_enablesOperatorChangeAndPinChange() throws Exception {
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        Button changeOperator = new Button();
        Button changePin = new Button();
        ShellFactoryOperatorToolbar toolbar =
                new ShellFactoryOperatorToolbar(
                        null, null, null, null, new Label(), changeOperator, changePin);

        toolbar.refreshOperatorUserLabel();

        assertFalse(changeOperator.isDisable());
        assertFalse(changePin.isDisable());
    }
}
