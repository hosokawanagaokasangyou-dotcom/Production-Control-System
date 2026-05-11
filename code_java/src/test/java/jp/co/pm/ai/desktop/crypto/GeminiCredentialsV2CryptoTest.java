package jp.co.pm.ai.desktop.crypto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

class GeminiCredentialsV2CryptoTest {

    @Test
    void roundTripMatchesPassphrase(@TempDir Path tmp) throws Exception {
        Path out = tmp.resolve("gemini_credentials.encrypted.json");
        String plain = "test-api-key-roundtrip";
        GeminiCredentialsV2Crypto.writeEncryptedCredentials(out, plain);
        String json = Files.readString(out);
        String back =
                GeminiCredentialsV2Crypto.decryptGeminiApiKeyFromJsonString(
                        json, GeminiCredentialsV2Crypto.DEFAULT_PASSPHRASE);
        assertEquals(plain, back);
    }

    @Test
    void canDecryptRepositoryBundledEncryptedJsonIfPresent() throws Exception {
        Path cwd = Path.of("").toAbsolutePath();
        Path bundled =
                cwd.getFileName() != null && "code_java".equals(cwd.getFileName().toString())
                        ? cwd.getParent().resolve("code/gemini_credentials.encrypted.json")
                        : cwd.resolve("code/gemini_credentials.encrypted.json");
        assumeTrue(Files.isRegularFile(bundled), "repo bundled credentials: " + bundled);
        String json = Files.readString(bundled);
        String key =
                GeminiCredentialsV2Crypto.decryptGeminiApiKeyFromJsonString(
                        json, GeminiCredentialsV2Crypto.DEFAULT_PASSPHRASE);
        assumeTrue(key != null && !key.isBlank());
    }
}
