package jp.co.pm.ai.desktop.crypto;

import java.nio.charset.StandardCharsets;
import java.security.GeneralSecurityException;
import java.security.SecureRandom;
import java.security.spec.KeySpec;
import java.util.Base64;

import javax.crypto.Cipher;
import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.PBEKeySpec;
import javax.crypto.spec.SecretKeySpec;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * 操作者別アラジン RPA ログインパスワードの復号可能暗号化。
 *
 * <p>C# {@code AladdinOperatorCredentialsCrypto} と同一形式（AES-256-CBC + PBKDF2-HMAC-SHA256）。
 */
public final class AladdinOperatorCredentialsCrypto {

    public static final int FORMAT_VERSION = 1;
    public static final int DEFAULT_ITERATIONS = 480_000;
    /** 工場内共有ストア向け固定パスフレーズ（Gemini 証明書と別系統）。 */
    public static final String DEFAULT_PASSPHRASE = "pm-ai-aladdin-operator";

    private static final ObjectMapper JSON = new ObjectMapper();
    private static final SecureRandom SECURE_RANDOM = new SecureRandom();

    private AladdinOperatorCredentialsCrypto() {}

    public static ObjectNode encryptToPayload(String plaintext) throws GeneralSecurityException {
        return encryptToPayload(plaintext, DEFAULT_PASSPHRASE, DEFAULT_ITERATIONS);
    }

    public static ObjectNode encryptToPayload(String plaintext, String passphrase, int iterations)
            throws GeneralSecurityException {
        String trimmed = plaintext != null ? plaintext : "";
        if (trimmed.isEmpty()) {
            throw new IllegalArgumentException("パスワードが空です。");
        }
        String phrase = passphrase != null ? passphrase.strip() : "";
        if (phrase.isEmpty()) {
            throw new IllegalArgumentException("パスフレーズが空です。");
        }
        if (iterations < 1) {
            throw new IllegalArgumentException("iterations が不正です。");
        }

        byte[] salt = new byte[16];
        byte[] iv = new byte[16];
        SECURE_RANDOM.nextBytes(salt);
        SECURE_RANDOM.nextBytes(iv);

        byte[] key = deriveKey(phrase, salt, iterations);
        Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");
        cipher.init(Cipher.ENCRYPT_MODE, new SecretKeySpec(key, "AES"), new IvParameterSpec(iv));
        byte[] ciphertext = cipher.doFinal(trimmed.getBytes(StandardCharsets.UTF_8));

        ObjectNode node = JSON.createObjectNode();
        node.put("v", FORMAT_VERSION);
        node.put("kdf", "pbkdf2_sha256");
        node.put("iterations", iterations);
        node.put("salt_b64", Base64.getEncoder().encodeToString(salt));
        node.put("iv_b64", Base64.getEncoder().encodeToString(iv));
        node.put("ciphertext_b64", Base64.getEncoder().encodeToString(ciphertext));
        return node;
    }

    public static String decryptFromPayload(JsonNode payload) throws GeneralSecurityException {
        return decryptFromPayload(payload, DEFAULT_PASSPHRASE);
    }

    public static String decryptFromPayload(JsonNode payload, String passphrase)
            throws GeneralSecurityException {
        if (payload == null || !payload.isObject()) {
            throw new IllegalArgumentException("暗号化ペイロードが不正です。");
        }
        int version = payload.path("v").asInt(0);
        if (version != FORMAT_VERSION) {
            throw new IllegalArgumentException("未対応の暗号化形式です: v=" + version);
        }
        String phrase = passphrase != null ? passphrase.strip() : "";
        if (phrase.isEmpty()) {
            throw new IllegalArgumentException("パスフレーズが空です。");
        }
        int iterations = payload.path("iterations").asInt(DEFAULT_ITERATIONS);
        byte[] salt = Base64.getDecoder().decode(payload.path("salt_b64").asText(""));
        byte[] iv = Base64.getDecoder().decode(payload.path("iv_b64").asText(""));
        byte[] ciphertext = Base64.getDecoder().decode(payload.path("ciphertext_b64").asText(""));

        byte[] key = deriveKey(phrase, salt, iterations);
        Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");
        cipher.init(Cipher.DECRYPT_MODE, new SecretKeySpec(key, "AES"), new IvParameterSpec(iv));
        byte[] plain = cipher.doFinal(ciphertext);
        return new String(plain, StandardCharsets.UTF_8);
    }

    private static byte[] deriveKey(String passphrase, byte[] salt, int iterations)
            throws GeneralSecurityException {
        SecretKeyFactory factory = SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256");
        KeySpec spec = new PBEKeySpec(passphrase.toCharArray(), salt, iterations, 256);
        return factory.generateSecret(spec).getEncoded();
    }
}
