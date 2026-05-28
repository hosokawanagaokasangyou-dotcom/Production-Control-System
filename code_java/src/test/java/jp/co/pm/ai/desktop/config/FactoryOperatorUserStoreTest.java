package jp.co.pm.ai.desktop.config;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class FactoryOperatorUserStoreTest {

    @TempDir
    Path tmp;

    @BeforeEach
    void isolateStore() throws Exception {
        System.setProperty("pm.ai.test.factoryOperatorUserStore", tmp.resolve("operators.json").toString());
        FactoryOperatorUserStore.resetStoreForTests();
    }

    @AfterEach
    void clearProperty() {
        System.clearProperty("pm.ai.test.factoryOperatorUserStore");
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
    void issuePinVerifyAndPersist() throws Exception {
        String pin = FactoryOperatorUserStore.issuePin(FactorySite.KONAN, "砂田");
        assertEquals(4, pin.length());
        assertTrue(FactoryOperatorUserStore.hasPin(FactorySite.KONAN, "砂田"));
        assertTrue(FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", pin));
        assertTrue(!FactoryOperatorUserStore.verifyPin(FactorySite.KONAN, "砂田", "0000"));
        assertEquals("設定済", FactoryOperatorUserStore.pinStatusLabel(FactorySite.KONAN, "砂田"));
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
        FactoryOperatorUserStore.addName(FactorySite.KONAN, "古家");
        assertTrue(!FactoryOperatorUserStore.hasPin(FactorySite.KONAN, "古家"));
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
        String saved = Files.readString(FactoryOperatorUserStore.storePath());
        assertTrue(saved.contains("\"schemaVersion\" : 2"));
        assertTrue(saved.contains("pinHashes"));
    }
}
