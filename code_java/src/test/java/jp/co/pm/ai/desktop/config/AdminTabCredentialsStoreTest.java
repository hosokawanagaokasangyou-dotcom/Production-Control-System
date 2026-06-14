package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class AdminTabCredentialsStoreTest {

    @TempDir
    Path tmp;

    private String priorStoreProperty;

    @BeforeEach
    void isolateStore() {
        priorStoreProperty = System.getProperty("pm.ai.test.adminTabCredentialsStore");
        System.setProperty("pm.ai.test.adminTabCredentialsStore", tmp.resolve("admin-tab-unlock.json").toString());
        AdminTabCredentialsStore.clearSavedUnlock();
    }

    @AfterEach
    void restoreProperty() {
        AdminTabCredentialsStore.clearSavedUnlock();
        if (priorStoreProperty != null) {
            System.setProperty("pm.ai.test.adminTabCredentialsStore", priorStoreProperty);
        } else {
            System.clearProperty("pm.ai.test.adminTabCredentialsStore");
        }
    }

    @Test
    void hasValidSavedUnlock_falseWhenMissing() {
        assertFalse(AdminTabCredentialsStore.hasValidSavedUnlock());
    }

    @Test
    void saveAndReloadUnlockToken() throws Exception {
        AdminTabCredentialsStore.saveAfterSuccessfulUnlock();
        assertTrue(Files.isRegularFile(AdminTabCredentialsStore.resolveStorePath()));
        assertTrue(AdminTabCredentialsStore.hasValidSavedUnlock());
    }

    @Test
    void clearSavedUnlock_removesFile() throws Exception {
        AdminTabCredentialsStore.saveAfterSuccessfulUnlock();
        assertTrue(AdminTabCredentialsStore.hasValidSavedUnlock());
        AdminTabCredentialsStore.clearSavedUnlock();
        assertFalse(Files.exists(AdminTabCredentialsStore.resolveStorePath()));
        assertFalse(AdminTabCredentialsStore.hasValidSavedUnlock());
    }

    @Test
    void hasValidSavedUnlock_falseForTamperedToken() throws Exception {
        AdminTabCredentialsStore.saveAfterSuccessfulUnlock();
        Path path = AdminTabCredentialsStore.resolveStorePath();
        String json = Files.readString(path);
        Files.writeString(path, json.replace("unlockTokenB64", "unlockTokenB64X"));
        assertFalse(AdminTabCredentialsStore.hasValidSavedUnlock());
    }

    @Test
    void hasValidSavedUnlock_falseForWrongUsernameInFile() throws Exception {
        AdminTabCredentialsStore.saveAfterSuccessfulUnlock();
        Path path = AdminTabCredentialsStore.resolveStorePath();
        String json = Files.readString(path).replace(FactoryOperatorUserStore.ADMIN_TAB_USERNAME, "admin");
        Files.writeString(path, json);
        assertFalse(AdminTabCredentialsStore.hasValidSavedUnlock());
    }

    @Test
    void usesSeparateStorePerDesktopAppHome() throws Exception {
        AdminTabCredentialsStore.saveAfterSuccessfulUnlock();
        assertTrue(AdminTabCredentialsStore.hasValidSavedUnlock());

        String priorHome = AppPaths.desktopAppHomeDirName();
        try {
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            System.setProperty(
                    "pm.ai.test.adminTabCredentialsStore",
                    tmp.resolve("rdp-admin-tab-unlock.json").toString());
            assertFalse(AdminTabCredentialsStore.hasValidSavedUnlock());
            AdminTabCredentialsStore.saveAfterSuccessfulUnlock();
            assertTrue(AdminTabCredentialsStore.hasValidSavedUnlock());
        } finally {
            AppPaths.setDesktopAppHomeDirName(priorHome);
            System.setProperty("pm.ai.test.adminTabCredentialsStore", tmp.resolve("admin-tab-unlock.json").toString());
        }
    }
}
