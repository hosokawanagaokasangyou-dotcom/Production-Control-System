package jp.co.pm.ai.desktop.crypto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

import com.fasterxml.jackson.databind.node.ObjectNode;

class AladdinOperatorCredentialsCryptoTest {

    @Test
    void roundTrip_encryptAndDecrypt() throws Exception {
        ObjectNode payload = AladdinOperatorCredentialsCrypto.encryptToPayload("000585585");
        String decrypted = AladdinOperatorCredentialsCrypto.decryptFromPayload(payload);
        assertEquals("000585585", decrypted);
    }

    @Test
    void decrypt_rejectsEmptyPasswordOnEncrypt() {
        assertThrows(IllegalArgumentException.class, () -> AladdinOperatorCredentialsCrypto.encryptToPayload(""));
    }
}
