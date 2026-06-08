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
 */
public final class OperatorAladdinCredentialsLauncherJson {

    public static final String FILE_NAME = "operator-aladdin-credentials.launcher.json";
    public static final int SCHEMA_VERSION = 1;

    private static final ObjectMapper JSON =
            new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    public record OperatorEntry(String loginId, ObjectNode passwordPayload) {}

    private OperatorAladdinCredentialsLauncherJson() {}

    public static void writeAllFactories(
            Path path, Map<FactorySite, Map<String, OperatorEntry>> operatorsByFactory)
            throws IOException {
        Objects.requireNonNull(path, "path");
        ObjectNode root = JSON.createObjectNode();
        root.put("schemaVersion", SCHEMA_VERSION);
        ObjectNode factories = root.putObject("factories");
        if (operatorsByFactory != null) {
            for (Map.Entry<FactorySite, Map<String, OperatorEntry>> factoryEntry :
                    operatorsByFactory.entrySet()) {
                FactorySite site = factoryEntry.getKey();
                if (site == null) {
                    continue;
                }
                ObjectNode ops = factories.putObject(site.name());
                Map<String, OperatorEntry> operators = factoryEntry.getValue();
                if (operators == null) {
                    continue;
                }
                for (Map.Entry<String, OperatorEntry> e : operators.entrySet()) {
                    String name = e.getKey();
                    OperatorEntry entry = e.getValue();
                    if (name == null || name.isBlank() || entry == null) {
                        continue;
                    }
                    String loginId = entry.loginId() != null ? entry.loginId().strip() : "";
                    if (loginId.isEmpty() || entry.passwordPayload() == null) {
                        continue;
                    }
                    ObjectNode row = ops.putObject(name.strip());
                    row.put("loginId", loginId);
                    row.set("password", entry.passwordPayload());
                }
            }
        }
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        JSON.writeValue(path.toFile(), root);
    }

    /** 既存ファイルを読み、工場・操作者名でエントリを返す（無ければ空）。 */
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
        FactorySite effective = factory != null ? factory : FactorySite.KONAN;
        JsonNode factories = root.get("factories");
        if (factories != null && factories.isObject()) {
            JsonNode ops = factories.get(effective.name());
            if (ops != null && ops.isObject()) {
                parseOperatorObject(ops, out);
            }
            return out;
        }
        // 旧単一 factory 形式
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
