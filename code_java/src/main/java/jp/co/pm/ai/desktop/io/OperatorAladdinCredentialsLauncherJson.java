package jp.co.pm.ai.desktop.io;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.FactorySite;

/**
 * C# {@code PmAiRdpRemoteLauncher} 向けの操作者別アラジン資格情報キャッシュ JSON。
 *
 * <p>正本は {@link jp.co.pm.ai.desktop.config.FactoryOperatorUserStore} の bin。保存のたびに再生成する。
 * ログイン情報は操作者ごとに一意（工場キーでは分けない。配備先フォルダが工場を区別する）。
 */
public final class OperatorAladdinCredentialsLauncherJson {

    public static final String FILE_NAME = "operator-aladdin-credentials.launcher.json";
    public static final int SCHEMA_VERSION = 2;

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    public record OperatorEntry(String loginId, ObjectNode passwordPayload) {}

    private OperatorAladdinCredentialsLauncherJson() {}

    /** 操作者名をキーに一意の資格情報を書き出す。 */
    public static void writeOperators(Path path, Map<String, OperatorEntry> operators)
            throws IOException {
        Objects.requireNonNull(path, "path");
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", SCHEMA_VERSION);
        ObjectNode ops = root.putObject("operators");
        if (operators != null) {
            for (Map.Entry<String, OperatorEntry> e : operators.entrySet()) {
                putOperatorRow(ops, e.getKey(), e.getValue());
            }
        }
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        JSON.writeValue(path.toFile(), root);
    }

    /**
     * 工場別マップを操作者名で一意に畳んで書き出す。
     *
     * @param preferredFactory 同名が複数工場にあるときの優先（配備先の工場）
     */
    public static void writeAllFactories(
            Path path,
            Map<FactorySite, Map<String, OperatorEntry>> operatorsByFactory,
            FactorySite preferredFactory)
            throws IOException {
        writeOperators(path, flattenUniqueOperators(operatorsByFactory, preferredFactory));
    }

    public static void writeAllFactories(
            Path path, Map<FactorySite, Map<String, OperatorEntry>> operatorsByFactory)
            throws IOException {
        writeAllFactories(path, operatorsByFactory, FactorySite.KONAN);
    }

    static Map<String, OperatorEntry> flattenUniqueOperators(
            Map<FactorySite, Map<String, OperatorEntry>> operatorsByFactory,
            FactorySite preferredFactory) {
        Map<String, OperatorEntry> unique = new LinkedHashMap<>();
        if (operatorsByFactory == null || operatorsByFactory.isEmpty()) {
            return unique;
        }
        FactorySite preferred = preferredFactory != null ? preferredFactory : FactorySite.KONAN;
        Map<String, OperatorEntry> preferredOps = operatorsByFactory.get(preferred);
        if (preferredOps != null) {
            for (Map.Entry<String, OperatorEntry> e : preferredOps.entrySet()) {
                putIfValid(unique, e.getKey(), e.getValue());
            }
        }
        for (Map.Entry<FactorySite, Map<String, OperatorEntry>> factoryEntry :
                operatorsByFactory.entrySet()) {
            if (factoryEntry.getKey() == preferred || factoryEntry.getValue() == null) {
                continue;
            }
            for (Map.Entry<String, OperatorEntry> e : factoryEntry.getValue().entrySet()) {
                putIfValid(unique, e.getKey(), e.getValue());
            }
        }
        return unique;
    }

    private static void putIfValid(Map<String, OperatorEntry> unique, String name, OperatorEntry entry) {
        if (name == null || name.isBlank() || entry == null || unique.containsKey(name.strip())) {
            return;
        }
        String loginId = entry.loginId() != null ? entry.loginId().strip() : "";
        if (loginId.isEmpty() || entry.passwordPayload() == null) {
            return;
        }
        unique.put(name.strip(), entry);
    }

    private static void putOperatorRow(ObjectNode ops, String name, OperatorEntry entry) {
        if (name == null || name.isBlank() || entry == null) {
            return;
        }
        String loginId = entry.loginId() != null ? entry.loginId().strip() : "";
        if (loginId.isEmpty() || entry.passwordPayload() == null) {
            return;
        }
        ObjectNode row = ops.putObject(name.strip());
        row.put("loginId", loginId);
        row.set("password", entry.passwordPayload());
    }

    /** 既存ファイルを読み、操作者名でエントリを返す（無ければ空）。工場引数は旧 factories 形式の読取用。 */
    public static Map<String, OperatorEntry> readOperators(Path path, FactorySite factory)
            throws IOException {
        Map<String, OperatorEntry> out = new LinkedHashMap<>();
        if (!Files.isRegularFile(path)) {
            return out;
        }
        JsonNode root = JSON.readTree(path.toFile());
        if (root == null || !root.isObject()) {
            return out;
        }
        JsonNode operators = root.get("operators");
        if (operators != null && operators.isObject()) {
            parseOperatorObject(operators, out);
            if (!out.isEmpty()) {
                return out;
            }
        }
        FactorySite effective = factory != null ? factory : FactorySite.KONAN;
        JsonNode factories = root.get("factories");
        if (factories != null && factories.isObject()) {
            JsonNode ops = factories.get(effective.name());
            if (ops != null && ops.isObject()) {
                parseOperatorObject(ops, out);
            }
            if (out.isEmpty()) {
                factories
                        .fields()
                        .forEachRemaining(
                                e -> {
                                    if (out.isEmpty() && e.getValue() != null && e.getValue().isObject()) {
                                        parseOperatorObject(e.getValue(), out);
                                    }
                                });
            }
            return out;
        }
        String factoryName = root.path("factory").asText("");
        if (!factoryName.isBlank() && !factoryName.equals(effective.name())) {
            return out;
        }
        JsonNode legacyOps = root.get("operators");
        if (legacyOps != null && legacyOps.isObject()) {
            parseOperatorObject(legacyOps, out);
        }
        return out;
    }

    private static void parseOperatorObject(JsonNode ops, Map<String, OperatorEntry> out) {
        ops.fields()
                .forEachRemaining(
                        e -> {
                            String name = e.getKey();
                            JsonNode node = e.getValue();
                            if (name == null || name.isBlank() || node == null || !node.isObject()) {
                                return;
                            }
                            String loginId = node.path("loginId").asText("").strip();
                            JsonNode password = node.get("password");
                            if (loginId.isEmpty() || password == null || !password.isObject()) {
                                return;
                            }
                            out.put(name.strip(), new OperatorEntry(loginId, (ObjectNode) password));
                        });
    }

    public static String readRawUtf8(Path path) throws IOException {
        if (!Files.isRegularFile(path)) {
            return "";
        }
        return Files.readString(path, StandardCharsets.UTF_8);
    }
}
