package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.GeneralSecurityException;
import java.security.MessageDigest;
import java.util.Base64;

import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * ユーザー管理者タブ解錠の PC ローカル永続化。
 *
 * <p>平文パスワードは保存せず、{@code user.home} 由来の鍵で HMAC-SHA256 した解錠トークンのみを
 * {@link AppPaths#resolveAdminTabCredentialsStorePath()} に書き込む。Git 追跡対象外。
 */
public final class AdminTabCredentialsStore {

    public static final int SCHEMA_VERSION = 1;

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final String TEST_STORE_PROPERTY = "pm.ai.test.adminTabCredentialsStore";

    private AdminTabCredentialsStore() {}

    public static Path resolveStorePath() {
        String test = System.getProperty(TEST_STORE_PROPERTY);
        if (test != null && !test.isBlank()) {
            return Path.of(test.strip()).toAbsolutePath().normalize();
        }
        return AppPaths.resolveAdminTabCredentialsStorePath();
    }

    /** 保存済み解錠トークンが現在の管理者資格情報と一致するとき true。 */
    public static boolean hasValidSavedUnlock() {
        try {
            Path path = resolveStorePath();
            if (!Files.isRegularFile(path)) {
                return false;
            }
            JsonNode root = JSON.readTree(path.toFile());
            if (root == null || !root.isObject()) {
                return false;
            }
            if (root.path("schemaVersion").asInt(0) != SCHEMA_VERSION) {
                return false;
            }
            String username = root.path("username").asText("").strip();
            if (!FactoryOperatorUserStore.ADMIN_TAB_USERNAME.equals(username)) {
                return false;
            }
            byte[] stored = decodeToken(root.path("unlockTokenB64").asText(""));
            if (stored.length == 0) {
                return false;
            }
            byte[] expected =
                    computeUnlockToken(username, FactoryOperatorUserStore.ADMIN_TAB_PASSWORD);
            return MessageDigest.isEqual(stored, expected);
        } catch (GeneralSecurityException e) {
            return false;
        } catch (Exception e) {
            return false;
        }
    }

    /** ダイアログで解錠成功後に呼ぶ。 */
    public static void saveAfterSuccessfulUnlock() throws IOException {
        String username = FactoryOperatorUserStore.ADMIN_TAB_USERNAME;
        byte[] token;
        try {
            token = computeUnlockToken(username, FactoryOperatorUserStore.ADMIN_TAB_PASSWORD);
        } catch (GeneralSecurityException e) {
            throw new IOException("解錠トークンの生成に失敗しました。", e);
        }
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", SCHEMA_VERSION);
        root.put("username", username);
        root.put("unlockTokenB64", Base64.getEncoder().encodeToString(token));
        Path path = resolveStorePath();
        Files.createDirectories(path.getParent());
        JSON.writerWithDefaultPrettyPrinter().writeValue(path.toFile(), root);
    }

    /** 保存済み解錠を削除する（再プロンプト用）。 */
    public static void clearSavedUnlock() {
        try {
            Files.deleteIfExists(resolveStorePath());
        } catch (IOException ignored) {
            // best effort
        }
    }

    static byte[] computeUnlockToken(String username, String password) throws GeneralSecurityException {
        Mac mac = Mac.getInstance("HmacSHA256");
        mac.init(new SecretKeySpec(deriveMachineKey(), "HmacSHA256"));
        String user = username != null ? username.strip() : "";
        String pass = password != null ? password : "";
        mac.update((user + "\0" + pass).getBytes(StandardCharsets.UTF_8));
        return mac.doFinal();
    }

    private static byte[] deriveMachineKey() throws GeneralSecurityException {
        String home = System.getProperty("user.home", "").strip();
        MessageDigest digest = MessageDigest.getInstance("SHA-256");
        digest.update("pm-ai-admin-tab-unlock:".getBytes(StandardCharsets.UTF_8));
        digest.update(home.getBytes(StandardCharsets.UTF_8));
        return digest.digest();
    }

    private static byte[] decodeToken(String encoded) {
        if (encoded == null || encoded.isBlank()) {
            return new byte[0];
        }
        try {
            return Base64.getDecoder().decode(encoded.strip());
        } catch (IllegalArgumentException e) {
            return new byte[0];
        }
    }
}
