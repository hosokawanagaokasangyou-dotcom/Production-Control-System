package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
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

import jp.co.pm.ai.desktop.io.OperatorAladdinCredentialsLauncherJson;

class FactoryOperatorUserStoreTest {

    @TempDir
    Path tmp;

    @BeforeEach
    void isolateStore() throws Exception {
        System.setProperty("pm.ai.test.factoryOperatorUserStore", tmp.resolve("operators.bin").toString());
        System.setProperty("pm.ai.test.factoryOperatorLastSelectedDir", tmp.resolve("last-selected").toString());
        FactoryOperatorUserStore.resetStoreForTests();
    }

    @AfterEach
    void clearProperty() {
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        System.clearProperty("pm.ai.test.factoryOperatorUserLegacyStore");
        System.clearProperty("pm.ai.test.factoryOperatorLastSelectedDir");
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
    void persistsLastSelectedLocallyNotInSharedBin() throws Exception {
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "古家");
        FactoryOperatorUserStore.clearSessionOperatorName();
        assertEquals("古家", FactoryOperatorUserStore.lastSelectedForFactory(FactorySite.KONAN));
        Path redirected = tmp.resolve("last-selected").resolve("last-factory-operator-konan.txt");
        assertTrue(Files.isRegularFile(redirected));
        assertEquals("古家", Files.readString(redirected, StandardCharsets.UTF_8).strip());
    }

    @Test
    void tryRestoreSessionFromLocalLastSelected_restoresWithoutPin() throws Exception {
        FactoryOperatorUserStore.assignPinByAdmin(FactorySite.KONAN, "砂田", "1234");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.clearSessionOperatorName();
        assertTrue(FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(FactorySite.KONAN));
        assertEquals("砂田", FactoryOperatorUserStore.sessionOperatorName());
    }

    @Test
    void tryRestoreSessionFromLocalLastSelected_failsWhenPinLocked() throws Exception {
        String pin = FactoryOperatorUserStore.assignPinByAdmin(FactorySite.KONAN, "砂田", "1234");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.clearSessionOperatorName();
        for (int i = 0; i < FactoryOperatorUserStore.MAX_CONSECUTIVE_PIN_FAILURES; i++) {
            FactoryOperatorUserStore.verifyPinAttempt(FactorySite.KONAN, "砂田", "0000");
        }
        assertTrue(FactoryOperatorUserStore.isPinLocked(FactorySite.KONAN, "砂田"));
        assertTrue(!FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(FactorySite.KONAN));
        assertEquals("", FactoryOperatorUserStore.sessionOperatorName());
        FactoryOperatorUserStore.unlockPin(FactorySite.KONAN, "砂田");
        assertTrue(FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(FactorySite.KONAN));
        assertEquals("砂田", FactoryOperatorUserStore.sessionOperatorName());
        FactoryOperatorUserStore.changePinByUser(FactorySite.KONAN, "砂田", pin, "2468");
    }

    @Test
    void tryRestoreSessionFromLocalLastSelected_failsWhenMustChangePin() throws Exception {
        FactoryOperatorUserStore.addName(FactorySite.KONAN, "新規");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "新規");
        FactoryOperatorUserStore.clearSessionOperatorName();
        assertTrue(FactoryOperatorUserStore.mustChangePin(FactorySite.KONAN, "新規"));
        assertTrue(!FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(FactorySite.KONAN));
        assertEquals("", FactoryOperatorUserStore.sessionOperatorName());
    }

    @Test
    void tryRestoreSessionFromLocalLastSelected_restoresGuest() throws Exception {
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, FactoryOperatorUserStore.GUEST_OPERATOR_NAME);
        FactoryOperatorUserStore.clearSessionOperatorName();
        assertTrue(FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(FactorySite.KONAN));
        assertEquals(FactoryOperatorUserStore.GUEST_OPERATOR_NAME, FactoryOperatorUserStore.sessionOperatorName());
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
        FactoryOperatorUserStore.configureFromUi(ui, FactorySite.KONAN);
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
    void configureFromUiUsesLocalFallbackWhenNetworkStoreNotWritable() throws Exception {
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        // 親が通常ファイルのとき共有 DATA 相当のパスへは書けない（UNC 権限不足の近似）
        Path blockedParent = tmp.resolve("blocked-not-dir");
        Files.writeString(blockedParent, "x");
        Path summary = blockedParent.resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Path networkBin = blockedParent.resolve(AppPaths.FACTORY_OPERATOR_USERS_BIN);
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK, summary.toString());
        Path local = AppPaths.localFactoryOperatorUsersStorePath(FactorySite.KONAN);
        Files.deleteIfExists(local);
        FactoryOperatorUserStore.resetStoreForTests();
        FactoryOperatorUserStore.configureFromUi(ui, FactorySite.KONAN);
        assertEquals(local, FactoryOperatorUserStore.storePath());
        assertTrue(FactoryOperatorUserStore.usingLocalStoreFallback());
        FactoryOperatorUserStore.addName(FactorySite.KONAN, "ローカル");
        assertTrue(Files.isRegularFile(local));
        assertTrue(!Files.isRegularFile(networkBin));
        Files.deleteIfExists(local);
    }

    @Test
    void configureFromUiUsesEffectiveFactoryWhenSummaryPointsToOtherFactory() throws Exception {
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
        Path konanSummary =
                tmp.resolve("湖南工場").resolve("共有DATA").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createDirectories(konanSummary.getParent());
        Map<String, String> ui =
                Map.of(
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        konanSummary.toString(),
                        AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR,
                        AppPaths.DEFAULT_PM_AI_ALADDIN_MASTER_DIR_KOKUBU);
        FactoryOperatorUserStore.configureFromUi(ui, FactorySite.KOKUBU);
        Path store = FactoryOperatorUserStore.storePath();
        Path konanBin =
                konanSummary.getParent()
                        .resolve(AppPaths.FACTORY_OPERATOR_USERS_BIN)
                        .toAbsolutePath()
                        .normalize();
        assertNotEquals(konanBin, store, "湖南サマリ配下ではなく国分側 bin を参照");
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
        System.setProperty("pm.ai.test.factoryOperatorLastSelectedDir", tmp.resolve("legacy-last").toString());
        FactoryOperatorUserStore.resetStoreForTests();
        assertEquals(List.of("砂田", "古家"), FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN));
        assertEquals("古家", FactoryOperatorUserStore.lastSelectedForFactory(FactorySite.KONAN));
        assertTrue(
                Files.isRegularFile(
                        tmp.resolve("legacy-last").resolve("last-factory-operator-konan.txt")));
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

    @Test
    void loginChoicesIncludeGuestWithoutPersistingToNames() throws Exception {
        List<String> choices = FactoryOperatorUserStore.loginChoicesForFactory(FactorySite.KONAN);
        assertTrue(choices.contains(FactoryOperatorUserStore.GUEST_OPERATOR_NAME));
        assertTrue(!FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN).contains(FactoryOperatorUserStore.GUEST_OPERATOR_NAME));
    }

    @Test
    void guestSessionRequiresNoPinAndBlocksSummaryExcel() throws Exception {
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, FactoryOperatorUserStore.GUEST_OPERATOR_NAME);
        assertEquals(FactoryOperatorUserStore.GUEST_OPERATOR_NAME, FactoryOperatorUserStore.sessionOperatorName());
        assertTrue(FactoryOperatorUserStore.isGuestSession());
        assertTrue(!FactoryOperatorUserStore.sessionMayGenerateSummaryExcel());
        assertTrue(!FactoryOperatorUserStore.sessionMayMutateRequestFormInput());
        assertTrue(!FactoryOperatorUserStore.hasPin(FactorySite.KONAN, FactoryOperatorUserStore.GUEST_OPERATOR_NAME));
        assertEquals(
                FactoryOperatorUserStore.GUEST_OPERATOR_NAME,
                FactoryOperatorUserStore.lastSelectedForFactory(FactorySite.KONAN));
    }

    @Test
    void addNameRejectsGuestOperatorName() {
        assertThrows(
                IllegalArgumentException.class,
                () -> FactoryOperatorUserStore.addName(FactorySite.KONAN, FactoryOperatorUserStore.GUEST_OPERATOR_NAME));
    }

    @Test
    void assignPinByAdmin_setsSpecifiedPinWithoutMustChange() throws Exception {
        FactoryOperatorUserStore.assignPinByAdmin(FactorySite.KONAN, "砂田", "1357");
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", "1357"));
        assertTrue(!FactoryOperatorUserStore.mustChangePin(FactorySite.KONAN, "砂田"));
        assertEquals("1357", FactoryOperatorUserStore.adminViewablePin(FactorySite.KONAN, "砂田").orElse(""));
        assertEquals("設定済", FactoryOperatorUserStore.pinStatusLabel(FactorySite.KONAN, "砂田"));
    }

    @Test
    void assignPinByAdmin_rejectsGuestOperator() {
        assertThrows(
                IllegalArgumentException.class,
                () ->
                        FactoryOperatorUserStore.assignPinByAdmin(
                                FactorySite.KONAN, FactoryOperatorUserStore.GUEST_OPERATOR_NAME, "1234"));
    }

    @Test
    void aladdinCredentials_roundTripAndPinChangePreserves() throws Exception {
        FactoryOperatorUserStore.setAladdinCredentials(FactorySite.KONAN, "砂田", "000585", "000585585");
        assertTrue(FactoryOperatorUserStore.hasAladdinCredentials(FactorySite.KONAN, "砂田"));
        var creds = FactoryOperatorUserStore.aladdinCredentialsFor(FactorySite.KONAN, "砂田");
        assertTrue(creds.isPresent());
        assertEquals("000585", creds.get().loginId());
        assertEquals("000585585", creds.get().password());

        FactoryOperatorUserStore.assignPinByAdmin(FactorySite.KONAN, "砂田", "2468");
        assertTrue(FactoryOperatorUserStore.hasAladdinCredentials(FactorySite.KONAN, "砂田"));
        assertEquals("000585", FactoryOperatorUserStore.aladdinLoginIdFor(FactorySite.KONAN, "砂田"));
    }

    @Test
    void aladdinCredentials_schema5MigrationPreservesPins() throws Exception {
        FactoryOperatorUserStore.writeRawJsonForTests(
                """
                {
                  "schemaVersion": 5,
                  "factories": {
                    "KONAN": {
                      "names": ["砂田", "古家", "図司", "細川"],
                      "pinHashes": {},
                      "pinMustChange": []
                    }
                  }
                }
                """);
        FactoryOperatorUserStore.setAladdinCredentials(FactorySite.KONAN, "細川", "111111", "secret-pass");
        assertEquals("111111", FactoryOperatorUserStore.aladdinLoginIdFor(FactorySite.KONAN, "細川"));
        assertTrue(FactoryOperatorUserStore.namesForFactory(FactorySite.KONAN).contains("細川"));

        byte[] saved = Files.readAllBytes(FactoryOperatorUserStore.storePath());
        FactoryOperatorUserStore.resetStoreForTests();
        Path store = Path.of(System.getProperty("pm.ai.test.factoryOperatorUserStore"));
        Files.write(store, saved);
        assertTrue(FactoryOperatorUserStore.hasAladdinCredentials(FactorySite.KONAN, "細川"));
    }

    @Test
    void removeName_dropsAladdinCredentialsOnlyForRemoved() throws Exception {
        FactoryOperatorUserStore.setAladdinCredentials(FactorySite.KONAN, "砂田", "a", "pass1");
        FactoryOperatorUserStore.setAladdinCredentials(FactorySite.KONAN, "古家", "b", "pass2");
        FactoryOperatorUserStore.removeName(FactorySite.KONAN, "砂田");
        assertTrue(!FactoryOperatorUserStore.hasAladdinCredentials(FactorySite.KONAN, "砂田"));
        assertTrue(FactoryOperatorUserStore.hasAladdinCredentials(FactorySite.KONAN, "古家"));
    }

    @Test
    void resolveRdpLauncherOperatorName_prefersSessionOverUiEnv() throws Exception {
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        try {
            assertEquals(
                    "砂田",
                    FactoryOperatorUserStore.resolveRdpLauncherOperatorName(
                            Map.of(AppPaths.KEY_PM_AI_OPERATOR_USER, "古家")));
        } finally {
            FactoryOperatorUserStore.clearSessionOperatorName();
        }
    }

    @Test
    void resolveRdpLauncherOperatorName_fallsBackToUiEnvWhenSessionEmpty() {
        FactoryOperatorUserStore.clearSessionOperatorName();
        assertEquals(
                "古家",
                FactoryOperatorUserStore.resolveRdpLauncherOperatorName(
                        Map.of(AppPaths.KEY_PM_AI_OPERATOR_USER, "古家")));
    }

    @Test
    void syncLauncherCredentialsJsonToDeployDir_writesBesideRapIni(@TempDir Path fakeRepo)
            throws Exception {
        Path summary = fakeRepo.resolve("code").resolve(AppPaths.SUMMARY_AI_DISPATCH_XLSX);
        Files.createDirectories(summary.getParent());
        Files.writeString(summary, "x");
        Map<String, String> ui = Map.of(AppPaths.KEY_PM_AI_REPO_ROOT, fakeRepo.toString());

        FactoryOperatorUserStore.setAladdinCredentials(FactorySite.KONAN, "砂田", "000585", "000585585");
        FactoryOperatorUserStore.selectSessionOperator(FactorySite.KONAN, "砂田");
        FactoryOperatorUserStore.syncLauncherCredentialsJsonToDeployDir(ui);

        Path jsonPath =
                AppPaths.resolveRdpLauncherIni(ui, "砂田")
                        .getParent()
                        .resolve(OperatorAladdinCredentialsLauncherJson.FILE_NAME);
        assertTrue(Files.isRegularFile(jsonPath));
        String text = Files.readString(jsonPath, StandardCharsets.UTF_8);
        assertTrue(text.contains("000585"));
        assertTrue(text.contains("砂田"));
    }

    @Test
    void configureForStandaloneLocalOnly_usesRdpLauncherStorePath() {
        String priorHome = AppPaths.desktopAppHomeDirName();
        String priorTestStore = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        try {
            if (priorTestStore != null) {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            FactoryOperatorUserStore.configureForStandaloneLocalOnly(Map.of(), FactorySite.RDP_LAUNCHER);
            assertEquals(
                    AppPaths.resolveRdpLauncherOperatorUsersStorePath(Map.of()),
                    FactoryOperatorUserStore.storePath());
            assertEquals(
                    AppPaths.resolveRdpLauncherOperatorUsersStorePath(Map.of()),
                    FactoryOperatorUserStore.networkStorePath());
        } finally {
            AppPaths.setDesktopAppHomeDirName(priorHome);
            if (priorTestStore != null) {
                System.setProperty("pm.ai.test.factoryOperatorUserStore", priorTestStore);
            } else {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
        }
    }

    @Test
    void tryRestoreSessionFromLocalLastSelected_rdpLauncherApp(@TempDir Path tmp) throws Exception {
        String priorHome = AppPaths.desktopAppHomeDirName();
        String priorTestStore = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        Path customDir = tmp.resolve("rdp-user-store");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR, customDir.toString());
        try {
            if (priorTestStore != null) {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            FactoryOperatorUserStore.configureForStandaloneLocalOnly(ui, FactorySite.RDP_LAUNCHER);
            FactoryOperatorUserStore.selectSessionRdpDepartment(
                    FactoryOperatorUserStore.DEFAULT_RDP_DEPARTMENT_LABEL);
            FactoryOperatorUserStore.addName(FactorySite.RDP_LAUNCHER, "山田");
            FactoryOperatorUserStore.assignPinByAdmin(FactorySite.RDP_LAUNCHER, "山田", "1234");
            FactoryOperatorUserStore.selectSessionOperator(FactorySite.RDP_LAUNCHER, "山田");
            FactoryOperatorUserStore.clearSessionOperatorName();
            assertTrue(
                    FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(
                            FactorySite.RDP_LAUNCHER));
            assertEquals("山田", FactoryOperatorUserStore.sessionOperatorName());
        } finally {
            AppPaths.setDesktopAppHomeDirName(priorHome);
            if (priorTestStore != null) {
                System.setProperty("pm.ai.test.factoryOperatorUserStore", priorTestStore);
            } else {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
        }
    }

    @Test
    void rdpLauncherDepartments_areScopedPerDepartment(@TempDir Path tmp) throws Exception {
        String priorHome = AppPaths.desktopAppHomeDirName();
        String priorTestStore = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        Path customDir = tmp.resolve("rdp-user-store");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR, customDir.toString());
        try {
            if (priorTestStore != null) {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            FactoryOperatorUserStore.configureForStandaloneLocalOnly(ui, FactorySite.RDP_LAUNCHER);
            FactoryOperatorUserStore.addRdpDepartment("製造1課");
            FactoryOperatorUserStore.selectSessionRdpDepartment("製造1課");
            FactoryOperatorUserStore.setAdminRdpDepartmentContext("製造1課");
            FactoryOperatorUserStore.addName(FactorySite.RDP_LAUNCHER, "山田");
            FactoryOperatorUserStore.setAdminRdpDepartmentContext(
                    FactoryOperatorUserStore.DEFAULT_RDP_DEPARTMENT_LABEL);
            assertTrue(
                    FactoryOperatorUserStore.namesForFactory(FactorySite.RDP_LAUNCHER).contains("山田"),
                    "操作者変更ダイアログはセッション部署（製造1課）のユーザーを使う");
            assertTrue(
                    !FactoryOperatorUserStore.namesForAdminTable(FactorySite.RDP_LAUNCHER).contains("山田"),
                    "ユーザー管理者タブは編集中部署（既定）のユーザーを使う");
            FactoryOperatorUserStore.setAdminRdpDepartmentContext("製造1課");
            assertTrue(
                    FactoryOperatorUserStore.namesForAdminTable(FactorySite.RDP_LAUNCHER).contains("山田"));
            assertEquals(List.of("既定", "製造1課"), FactoryOperatorUserStore.listRdpDepartmentKeys());
        } finally {
            AppPaths.setDesktopAppHomeDirName(priorHome);
            if (priorTestStore != null) {
                System.setProperty("pm.ai.test.factoryOperatorUserStore", priorTestStore);
            } else {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
        }
    }

    @Test
    void pinStatusLabel_usesAdminDepartmentNotSession(@TempDir Path tmp) throws Exception {
        String priorHome = AppPaths.desktopAppHomeDirName();
        String priorTestStore = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        Path customDir = tmp.resolve("rdp-user-store");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR, customDir.toString());
        try {
            if (priorTestStore != null) {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            FactoryOperatorUserStore.configureForStandaloneLocalOnly(ui, FactorySite.RDP_LAUNCHER);
            FactoryOperatorUserStore.addRdpDepartment("国分工場");
            FactoryOperatorUserStore.addRdpDepartment("湖南工場");
            FactoryOperatorUserStore.selectSessionRdpDepartment("湖南工場");
            FactoryOperatorUserStore.setAdminRdpDepartmentContext("国分工場");
            FactoryOperatorUserStore.addName(FactorySite.RDP_LAUNCHER, "河合");
            FactoryOperatorUserStore.issuePin(FactorySite.RDP_LAUNCHER, "河合");
            assertFalse(
                    FactoryOperatorUserStore.hasPin(FactorySite.RDP_LAUNCHER, "河合"),
                    "ログイン検証はセッション部署を参照する");
            assertEquals(
                    "初回変更待",
                    FactoryOperatorUserStore.pinStatusLabel(FactorySite.RDP_LAUNCHER, "河合"),
                    "管理者一覧は編集中部署の PIN を表示する");
            assertNotEquals(
                    "—",
                    FactoryOperatorUserStore.adminPinDisplayLabel(FactorySite.RDP_LAUNCHER, "河合"));
            FactoryOperatorUserStore.setAdminRdpDepartmentContext("湖南工場");
            assertEquals(
                    "未設定",
                    FactoryOperatorUserStore.pinStatusLabel(FactorySite.RDP_LAUNCHER, "河合"),
                    "別部署を編集中は当該ユーザーの PIN を表示しない");
        } finally {
            AppPaths.setDesktopAppHomeDirName(priorHome);
            if (priorTestStore != null) {
                System.setProperty("pm.ai.test.factoryOperatorUserStore", priorTestStore);
            } else {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
        }
    }

    @Test
    void loginChoicesForFactory_rdpFollowsSessionDepartmentWhenAdminContextDiffers(@TempDir Path tmp)
            throws Exception {
        String priorHome = AppPaths.desktopAppHomeDirName();
        String priorTestStore = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        Path customDir = tmp.resolve("rdp-user-store");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR, customDir.toString());
        try {
            if (priorTestStore != null) {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            FactoryOperatorUserStore.configureForStandaloneLocalOnly(ui, FactorySite.RDP_LAUNCHER);
            FactoryOperatorUserStore.addRdpDepartment("湖南工場");
            FactoryOperatorUserStore.selectSessionRdpDepartment("湖南工場");
            FactoryOperatorUserStore.setAdminRdpDepartmentContext("湖南工場");
            FactoryOperatorUserStore.addName(FactorySite.RDP_LAUNCHER, "細川");
            FactoryOperatorUserStore.setAdminRdpDepartmentContext(
                    FactoryOperatorUserStore.DEFAULT_RDP_DEPARTMENT_LABEL);
            List<String> choices =
                    FactoryOperatorUserStore.loginChoicesForFactory(FactorySite.RDP_LAUNCHER);
            assertTrue(choices.contains("細川"));
            assertTrue(choices.contains(FactoryOperatorUserStore.GUEST_OPERATOR_NAME));
        } finally {
            AppPaths.setDesktopAppHomeDirName(priorHome);
            if (priorTestStore != null) {
                System.setProperty("pm.ai.test.factoryOperatorUserStore", priorTestStore);
            } else {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
        }
    }

    @Test
    void tryRestoreSession_rdpLauncherRestoresDepartmentAndOperator(@TempDir Path tmp) throws Exception {
        String priorHome = AppPaths.desktopAppHomeDirName();
        String priorTestStore = System.getProperty("pm.ai.test.factoryOperatorUserStore");
        Path customDir = tmp.resolve("rdp-user-store");
        Map<String, String> ui =
                Map.of(AppPaths.KEY_PM_AI_RDP_OPERATOR_USERS_STORE_DIR, customDir.toString());
        try {
            if (priorTestStore != null) {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
            AppPaths.setDesktopAppHomeDirName(AppPaths.REMOTE_DESKTOP_APP_HOME_DIR_NAME);
            FactoryOperatorUserStore.configureForStandaloneLocalOnly(ui, FactorySite.RDP_LAUNCHER);
            FactoryOperatorUserStore.addRdpDepartment("製造1課");
            FactoryOperatorUserStore.selectSessionRdpDepartment("製造1課");
            FactoryOperatorUserStore.setAdminRdpDepartmentContext("製造1課");
            FactoryOperatorUserStore.addName(FactorySite.RDP_LAUNCHER, "山田");
            FactoryOperatorUserStore.assignPinByAdmin(FactorySite.RDP_LAUNCHER, "山田", "1234");
            FactoryOperatorUserStore.selectSessionOperator(FactorySite.RDP_LAUNCHER, "山田");
            FactoryOperatorUserStore.clearSessionOperatorName();
            FactoryOperatorUserStore.clearSessionRdpDepartmentKey();
            assertTrue(FactoryOperatorUserStore.tryRestoreSessionRdpDepartmentFromLocal());
            assertEquals("製造1課", FactoryOperatorUserStore.sessionRdpDepartmentKey());
            assertTrue(
                    FactoryOperatorUserStore.tryRestoreSessionFromLocalLastSelected(
                            FactorySite.RDP_LAUNCHER));
            assertEquals("山田", FactoryOperatorUserStore.sessionOperatorName());
        } finally {
            AppPaths.setDesktopAppHomeDirName(priorHome);
            if (priorTestStore != null) {
                System.setProperty("pm.ai.test.factoryOperatorUserStore", priorTestStore);
            } else {
                System.clearProperty("pm.ai.test.factoryOperatorUserStore");
            }
        }
    }
}
