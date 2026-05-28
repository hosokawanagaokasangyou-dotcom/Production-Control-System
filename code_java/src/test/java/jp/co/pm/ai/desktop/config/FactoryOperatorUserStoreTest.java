package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class FactoryOperatorUserStoreTest {

    @TempDir
    Path tmp;

    @BeforeEach
    void isolateStore() throws Exception {
        System.setProperty("pm.ai.test.factoryOperatorUserStore", tmp.resolve("operators.bin").toString());
        FactoryOperatorUserStore.resetStoreForTests();
    }

    @AfterEach
    void clearProperty() {
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        System.clearProperty("pm.ai.test.factoryOperatorUserLegacyStore");
    }

    @Test
    void defaultNamesForBothFactories() throws Exception {
        assertEquals(FactoryOperatorUserStore.DEFAULT_NAMES, FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN));
        assertEquals(FactoryOperatorUserStore.DEFAULT_NAMES, FactoryOperatorUserStore.namesForFactory(FactorySite.KOKUBU));
    }

    @Test
    void addRemoveAndSelectAreFactoryScoped() throws Exception {
        FactoryOperatorUserStore.addName(FactorySite.KONAN, "山田");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "山田");
        assertEquals("山田", FactoryOperatorUserStore.sessionOperatorName());
        assertTrue(FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN).contains("山田"));
        assertTrue(!FactoryOperatorUserStore.namesForFactory(FactorySite.KOKUBU).contains("山田"));

        FactoryOperatorUserStore.removeName(FactorySite.KONAN, "山田");
        assertTrue(!FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN).contains("山田"));
    }

    @Test
    void cannotRemoveLastName() throws Exception {
        for (String name : FactoryOperatorUserStore.DEFAULT_NAMES) {
            if (!name.equals("砂田")) {
                FactoryOperatorUserStore.removeName(FactorySite.KONAN, name);
            }
        }
        assertThrows(IllegalStateException.class, () -> FactoryOperatorUserStore.removeName(FactorySite.KONAN, "砂田"));
    }

    @Test
    void resetNamesToDefaults() throws Exception {
        FactoryOperatorUserStore.addName(FactorySite.KOKUBU, "テスト");
        FactoryOperatorUserStore.resetNamesToDefaults(FactorySite.KOKUBU);
        assertEquals(FactoryOperatorUserStore.DEFAULT_NAMES, FactoryOperatorUserStore.namesForFactory(FactorySite.KOKUBU));
    }

    @Test
    void persistsLastSelected() throws Exception {
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "古家");
        FactoryOperatorUserStore.clearSessionOperatorName();
        assertEquals("古家", FactoryOperatorUserStore.lastSelectedForFactory(FactorySite.KONAN));
        assertTrue(Files.isRegularFile(FactoryOperatorUserStore.storePath()));
    }

    @Test
    void verifyAdminTabAccess_requiresAdministratorUsernameAndPassword() {
        assertTrue(
                FactoryOperatorUserStore.verifyAdminTabAccess(
                        FactoryOperatorUserStore.ADMIN_TAB_USERNAME,
                        FactoryOperatorUserStore.ADMIN_TAB_PASSWORD));
        assertTrue(!FactoryOperatorUserStore.verifyAdminTabAccess("admin", FactoryOperatorUserStore.ADMIN_TAB_PASSWORD));
        assertTrue(!FactoryOperatorUserStore.verifyAdminTabAccess(FactoryOperatorUserStore.ADMIN_TAB_USERNAME, "wrong"));
        assertTrue(!FactoryOperatorUserStore.verifyAdminTabAccess("", FactoryOperatorUserStore.ADMIN_TAB_PASSWORD));
    }

    @Test
    void adminViewablePin_storedOnIssueAndUpdatedOnChange() throws Exception {
        String pin = FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "砂田");
        assertEquals(pin, FactoryOperatorUserStore.adminViewablePin(FactorySite.KONAN, "砂田").orElse(""));
        assertEquals(pin, FactoryOperatorUserStore.adminPinDisplayLabel(FactorySite.KONAN, "砂田"));
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.changePinByUser(FactorySite.KONAN, "砂田", pin, "2468");
        assertEquals("2468", FactoryOperatorUserStore.adminViewablePin(FactorySite.KONAN, "砂田").orElse(""));
        assertEquals("2468", FactoryOperatorUserStore.adminPinDisplayLabel(FactorySite.KONAN, "砂田"));
    }

    @Test
    void adminPinDisplayLabel_showsPlaceholderWhenPlaintextMissing() throws Exception {
        FactoryOperatorUserStore.writeRawJsonForTests(
                """
                {
                  "schemaVersion": 5,
                  "factories": {
                    "KONAN": {
                      "names": ["砂田"],
                      "lastSelected": "",
                      "pinHashes": { "砂田": "deadbeef" }
                    }
                  }
                }
                """);
        assertEquals("（再発行で確認）", FactoryOperatorUserStore.adminPinDisplayLabel(FactorySite.KONAN, "砂田"));
    }

    @Test
    void issuePinVerifyAndPersist() throws Exception {
        String pin = FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "砂田");
        assertEquals(4, pin.length());
        assertTrue(FactoryOperatorUserStore.hasPin(FactorySite.KONAN, "砂田"));
        assertTrue(FactoryOperatorUserStore.mustChangePin(FactorySite.KONAN, "砂田"));
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", pin));
        assertTrue(!FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", "0000"));
        assertEquals("初回変更待", FactoryOperatorUserStore.pinStatusLabel(FactorySite.KONAN, "砂田"));
    }

    @Test
    void addNameIssuesRandomPinRequiringFirstLoginChange() throws Exception {
        String pin = FactoryOperatorUserStore.addName(FactorySite.KONAN, "山田");
        assertEquals(4, pin.length());
        assertTrue(FactoryOperatorUserStore.hasPin(FactorySite.KONAN, "山田"));
        assertTrue(FactoryOperatorUserStore.mustChangePin(FactorySite.KONAN, "山田"));
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "山田", pin));
        assertEquals("初回変更待", FactoryOperatorUserStore.pinStatusLabel(FactorySite.KONAN, "山田"));
    }

    @Test
    void changePinOnFirstLoginClearsMustChangeFlag() throws Exception {
        String pin = FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.changePinOnFirstLogin(FactorySite.KONAN, "砂田", pin, "8765");
        assertTrue(!FactoryOperatorUserStore.mustChangePin(FactorySite.KONAN, "砂田"));
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", "8765"));
        assertEquals("設定済", FactoryOperatorUserStore.pinStatusLabel(FactorySite.KONAN, "砂田"));
    }

    @Test
    void normalizePinAcceptsFourToTenDigits() {
        assertEquals("1234", FactoryOperatorUserStore.normalizePin("1234"));
        assertEquals("1234567890", FactoryOperatorUserStore.normalizePin("1234567890"));
        assertTrue(FactoryOperatorUserStore.normalizePin("123") == null);
        assertTrue(FactoryOperatorUserStore.normalizePin("12345678901") == null);
        assertTrue(FactoryOperatorUserStore.normalizePin("12a4") == null);
    }

    @Test
    void changePinByUser_setsAndChangesOwnPin() throws Exception {
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.changePinByUser(FactorySite.KONAN, "砂田", "", "567890");
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", "567890"));
        FactoryOperatorUserStore.changePinByUser(FactorySite.KONAN, "砂田", "567890", "4321");
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", "4321"));
    }

    @Test
    void changePinByUser_clearsMustChangeFlag() throws Exception {
        String pin = FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.changePinByUser(FactorySite.KONAN, "砂田", pin, "9999");
        assertTrue(!FactoryOperatorUserStore.mustChangePin(FactorySite.KONAN, "砂田"));
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", "9999"));
    }

    @Test
    void changePinByUser_rejectsOtherUserOrWrongCurrent() throws Exception {
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.changePinByUser(FactorySite.KONAN, "砂田", "", "1111");
        assertThrows(
                IllegalStateException.class,
                () -> FactoryOperatorUserStore.changePinByUser(FactorySite.KONAN, "古家", "1111", "2222"));
        assertThrows(
                IllegalArgumentException.class,
                () -> FactoryOperatorUserStore.changePinByUser(FactorySite.KONAN, "砂田", "9999", "2222"));
    }

    @Test
    void verifyPinAllowsWhenUnset() throws Exception {
        assertTrue(!FactoryOperatorUserStore.hasPin(FactorySite.KONAN, "砂田"));
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", "1234"));
    }

    @Test
    void removeNameClearsPin() throws Exception {
        FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "古家");
        FactoryOperatorUserStore.removeName(FactorySite.KONAN, "古家");
        assertTrue(!FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN).contains("古家"));
        String pin = FactoryOperatorUserStore.addName(FactorySite.KONAN, "古家");
        assertTrue(FactoryOperatorUserStore.hasPin(FactorySite.KONAN, "古家"));
        assertTrue(FactoryOperatorUserStore.mustChangePin(FactorySite.KONAN, "古家"));
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "古家", pin));
    }

    @Test
    void migratesSchemaV1Json() throws Exception {
        FactoryOperatorUserStore.writeRawJsonForTests(
                """
                {
                  "schemaVersion": 1,
                  "factories": {
                    "KONAN": {
                      "names": ["砂田", "古家"],
                      "lastSelected": "砂田"
                    }
                  }
                }
                """);
        assertEquals(List.of("砂田", "古家"), FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN));
        assertTrue(!FactoryOperatorUserStore.hasPin(FactorySite.KONAN, "砂田"));
        String pin = FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "砂田");
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", pin));
        byte[] saved = Files.readAllBytes(FactoryOperatorUserStore.storePath());
        assertEquals('P', saved[0]);
        assertEquals('M', saved[1]);
        assertEquals('O', saved[2]);
        assertEquals('U', saved[3]);
    }

    @Test
    void configureFromUiUsesSummaryWorkbookParent() throws Exception {
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        Path summary = tmp.resolve("shared").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createDirectories(summary.getParent());
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summary.toString());
        FactoryOperatorUserStore.configureFromUi(ui);
        Path expected =
                summary.getParent()
                        .resolve(AppPaths.FACTORY_OPERATOR_USERS_BIN)
                        .toAbsolutePath()
                        .normalize();
        assertEquals(expected, FactoryOperatorUserStore.storePath());
        FactoryOperatorUserStore.addName(FactorySite.KONAN, "テスト");
        assertTrue(Files.isRegularFile(expected));
        byte[] saved = Files.readAllBytes(expected);
        assertEquals('P', saved[0]);
    }

    @Test
    void migratesLegacyHomeJsonToBinary() throws Exception {
        Path legacy = tmp.resolve("legacy").resolve("factory-operator-users.json");
        Files.createDirectories(legacy.getParent());
        Files.writeString(
                legacy,
                """
                {
                  "schemaVersion": 3,
                  "factories": {
                    "KONAN": {
                      "names": ["砂田", "古家"],
                      "lastSelected": "古家"
                    }
                  }
                }
                """,
                StandardCharsets.UTF_8);
        Path target = tmp.resolve("shared").resolve(AppPaths.FACTORY_OPERATOR_USERS_BIN);
        Files.createDirectories(target.getParent());
        System.setProperty("pm.ai.test.factoryOperatorUserStore", target.toString());
        System.setProperty("pm.ai.test.factoryOperatorUserLegacyStore", legacy.toString());
        FactoryOperatorUserStore.resetStoreForTests();
        assertEquals(List.of("砂田", "古家"), FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN));
        assertEquals("古家", FactoryOperatorUserStore.lastSelectedForFactory(FactorySite.KONAN));
        assertTrue(Files.isRegularFile(target));
        byte[] saved = Files.readAllBytes(target);
        assertEquals('P', saved[0]);
    }

    @Test
    void locksAfterTwentyConsecutiveWrongPins() throws Exception {
        String pin = FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "砂田");
        for (int i = 0; i < FactoryOperatorUserStore.MAX_CONSECUTIVE_PIN_FAILURES - 1; i++) {
            assertEquals(
                    FactoryOperatorUserStore.PinVerificationResult.WRONG_PIN,
                    FactoryOperatorUserStore.verifyPinAttempt(FactorySite.KONAN, "砂田", "0000"));
            assertTrue(!FactoryOperatorUserStore.isPinLocked(FactorySite.KONAN, "砂田"));
        }
        assertEquals(
                FactoryOperatorUserStore.PinVerificationResult.LOCKED,
                FactoryOperatorUserStore.verifyPinAttempt(FactorySite.KONAN, "砂田", "0000"));
        assertTrue(FactoryOperatorUserStore.isPinLocked(FactorySite.KONAN, "砂田"));
        assertEquals("ロック", FactoryOperatorUserStore.pinStatusLabel(FactorySite.KONAN, "砂田"));
        assertEquals(
                FactoryOperatorUserStore.PinVerificationResult.LOCKED,
                FactoryOperatorUserStore.verifyPinAttempt(FactorySite.KONAN, "砂田", pin));
        FactoryOperatorUserStore.unlockPin(FactorySite.KONAN, "砂田");
        assertTrue(!FactoryOperatorUserStore.isPinLocked(FactorySite.KONAN, "砂田"));
        assertEquals(
                FactoryOperatorUserStore.PinVerificationResult.SUCCESS,
                FactoryOperatorUserStore.verifyPinAttempt(FactorySite.KONAN, "砂田", pin));
    }

    @Test
    void issuePinClearsLock() throws Exception {
        FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "古家");
        for (int i = 0; i < FactoryOperatorUserStore.MAX_CONSECUTIVE_PIN_FAILURES; i++) {
            FactoryOperatorUserStore.verifyPinAttempt(FactorySite.KONAN, "古家", "9999");
        }
        assertTrue(FactoryOperatorUserStore.isPinLocked(FactorySite.KONAN, "古家"));
        FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "古家");
        assertTrue(!FactoryOperatorUserStore.isPinLocked(FactorySite.KONAN, "古家"));
    }
}
