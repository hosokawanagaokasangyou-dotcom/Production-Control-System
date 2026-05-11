package jp.co.pm.ai.desktop.crypto;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.security.spec.KeySpec;
import java.util.Arrays;
import java.util.Base64;

import javax.crypto.Cipher;
import javax.crypto.Mac;
import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.PBEKeySpec;
import javax.crypto.spec.SecretKeySpec;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * Gemini 証明書 JSON（{@code format_version: 2}）の暗号化。Python {@code cryptography.fernet} +
 * PBKDF2-HMAC-SHA256 と同一形式で、{@code planning_core._core} の復号と互換である。
 */
public final class GeminiCredentialsV2Crypto {

    /** {@code planning_core} の {@code _GEMINI_CREDENTIALS_PASSPHRASE_FIXED} と同一。 */
    public static final String DEFAULT_PASSPHRASE = "nagaoka1234";

    public static final int DEFAULT_ITERATIONS = 480_000;

    private static final byte FERNET_VERSION = (byte) 0x80;

    private static final ObjectMapper MAPPER =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    private GeminiCredentialsV2Crypto() {}

    /**
     * 平文 API キーを {@code format_version 2} の JSON にし、{@code passphrase} で暗号化してファイルへ書く。
     *
     * @param outputPath 通常は環境変数 {@code GEMINI_CREDENTIALS_JSON} が指すパス
     */
    public static void writeEncryptedCredentials(
            Path outputPath, String geminiApiKeyPlain, String passphrase, int iterations)
            throws IOException, GeneralSecurityException {
        String trimmed = geminiApiKeyPlain != null ? geminiApiKeyPlain.strip() : "";
        if (trimmed.isEmpty()) {
            throw new IllegalArgumentException("gemini_api_key が空です。");
        }
        String phrase = passphrase != null ? passphrase.strip() : "";
        if (phrase.isEmpty()) {
            throw new IllegalArgumentException("パスフレーズが空です。");
        }
        if (iterations < 1) {
            throw new IllegalArgumentException("iterations が不正です。");
        }

        ObjectNode inner = JsonNodeFactory.instance.objectNode();
        inner.put("gemini_api_key", trimmed);
        byte[] innerUtf8 = MAPPER.writeValueAsBytes(inner);

        SecureRandom rng = new SecureRandom();
        byte[] salt = new byte[16];
        rng.nextBytes(salt);

        String fernetTokenAscii =
                fernetEncrypt(innerUtf8, phrase, salt, iterations);

        ObjectNode root = JsonNodeFactory.instance.objectNode();
        root.put("format_version", 2);
        root.put("kdf", "pbkdf2_sha256");
        root.put("iterations", iterations);
        root.put("salt_b64", Base64.getEncoder().encodeToString(salt));
        root.put("fernet_ciphertext", fernetTokenAscii);
        root.put(
                "description",
                "encrypt_gemini_credentials.py / GeminiCredentialsV2Crypto で生成。"
                        + "復号は planning_core の定数のみ（パスフレーズは社内手順で管理）。");

        Path parent = outputPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        MAPPER.writeValue(outputPath.toFile(), root);
    }

    /** {@link #DEFAULT_PASSPHRASE} と {@link #DEFAULT_ITERATIONS} で書き込む。 */
    public static void writeEncryptedCredentials(Path outputPath, String geminiApiKeyPlain)
            throws IOException, GeneralSecurityException {
        writeEncryptedCredentials(outputPath, geminiApiKeyPlain, DEFAULT_PASSPHRASE, DEFAULT_ITERATIONS);
    }

    /**
     * 単体テスト・検証用: 暗号化 JSON 文字列から API キーを復号する（runtime での利用は想定しない）。
     */
    public static String decryptGeminiApiKeyFromJsonString(String jsonUtf8, String passphrase)
            throws IOException, GeneralSecurityException {
        JsonNode root = MAPPER.readTree(jsonUtf8);
        String ct = root.path("fernet_ciphertext").asText(null);
        String saltB64 = root.path("salt_b64").asText(null);
        if (ct == null || ct.isBlank() || saltB64 == null || saltB64.isBlank()) {
            throw new IllegalArgumentException("暗号化フィールドが不足しています。");
        }
        int iterations = root.path("iterations").asInt(DEFAULT_ITERATIONS);
        byte[] salt = Base64.getDecoder().decode(saltB64);
        byte[] plain = fernetDecrypt(ct, passphrase, salt, iterations);
        JsonNode inner = MAPPER.readTree(plain);
        String k = textOrNull(inner.get("gemini_api_key"));
        if (k == null || k.isBlank()) {
            k = textOrNull(inner.get("GEMINI_API_KEY"));
        }
        return k != null ? k.strip() : "";
    }

    private static String textOrNull(JsonNode n) {
        return n != null && !n.isNull() ? n.asText() : null;
    }

    static String fernetEncrypt(byte[] plaintextUtf8, String passphrase, byte[] salt, int iterations)
            throws GeneralSecurityException {
        byte[] key32 = pbkdf2(passphrase, salt, iterations);
        byte[] signingKey = Arrays.copyOfRange(key32, 0, 16);
        byte[] encryptionKey = Arrays.copyOfRange(key32, 16, 32);

        long epochSeconds = System.currentTimeMillis() / 1000L;
        byte[] iv = new byte[16];
        new SecureRandom().nextBytes(iv);

        Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");
        cipher.init(Cipher.ENCRYPT_MODE, new SecretKeySpec(encryptionKey, "AES"), new IvParameterSpec(iv));
        byte[] ciphertext = cipher.doFinal(plaintextUtf8);

        byte[] basic =
                new byte[1 + 8 + 16 + ciphertext.length];
        basic[0] = FERNET_VERSION;
        putUint64Be(epochSeconds, basic, 1);
        System.arraycopy(iv, 0, basic, 9, 16);
        System.arraycopy(ciphertext, 0, basic, 25, ciphertext.length);

        Mac mac = Mac.getInstance("HmacSHA256");
        mac.init(new SecretKeySpec(signingKey, "HmacSHA256"));
        byte[] hmac = mac.doFinal(basic);

        byte[] token = Arrays.copyOf(basic, basic.length + hmac.length);
        System.arraycopy(hmac, 0, token, basic.length, hmac.length);

        return Base64.getUrlEncoder().withoutPadding().encodeToString(token);
    }

    static byte[] fernetDecrypt(String tokenUrlSafeB64, String passphrase, byte[] salt, int iterations)
            throws GeneralSecurityException, IOException {
        byte[] decoded;
        try {
            decoded = Base64.getUrlDecoder().decode(tokenUrlSafeB64.strip());
        } catch (IllegalArgumentException ex) {
            throw new IllegalArgumentException("fernet_ciphertext の Base64 が不正です。", ex);
        }
        if (decoded.length < 1 + 8 + 16 + 32) {
            throw new IllegalArgumentException("トークンが短すぎます。");
        }
        byte[] hmac = Arrays.copyOfRange(decoded, decoded.length - 32, decoded.length);
        byte[] basic = Arrays.copyOfRange(decoded, 0, decoded.length - 32);
        byte[] key32 = pbkdf2(passphrase, salt, iterations);
        byte[] signingKey = Arrays.copyOfRange(key32, 0, 16);
        byte[] encryptionKey = Arrays.copyOfRange(key32, 16, 32);

        Mac mac = Mac.getInstance("HmacSHA256");
        mac.init(new SecretKeySpec(signingKey, "HmacSHA256"));
        byte[] expected = mac.doFinal(basic);
        if (!MessageDigest.isEqual(hmac, expected)) {
            throw new SecurityException("HMAC が一致しません。");
        }
        if (basic.length < 25 || basic[0] != FERNET_VERSION) {
            throw new IllegalArgumentException("未対応の Fernet バージョンです。");
        }
        byte[] iv = Arrays.copyOfRange(basic, 9, 25);
        byte[] ct = Arrays.copyOfRange(basic, 25, basic.length);
        Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");
        cipher.init(Cipher.DECRYPT_MODE, new SecretKeySpec(encryptionKey, "AES"), new IvParameterSpec(iv));
        return cipher.doFinal(ct);
    }

    private static byte[] pbkdf2(String passphrase, byte[] salt, int iterations)
            throws GeneralSecurityException {
        KeySpec spec = new PBEKeySpec(passphrase.toCharArray(), salt, iterations, 256);
        SecretKeyFactory skf = SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256");
        return skf.generateSecret(spec).getEncoded();
    }

    private static void putUint64Be(long v, byte[] buf, int offset) {
        for (int i = 7; i >= 0; i--) {
            buf[offset + i] = (byte) (v & 0xff);
            v >>>= 8;
        }
    }

}
