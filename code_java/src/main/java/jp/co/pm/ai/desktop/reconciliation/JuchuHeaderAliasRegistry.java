package jp.co.pm.ai.desktop.reconciliation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.desktop.config.InitSettingPaths;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.Set;

/**
 * 受注ファイルごとの見出し別名・期待見出し上書き（列位置は {@link JuchuSheetColumnLayout.Col} 固定）。
 *
 * <p>永続化は工場別（{@link FactorySite}）のユーザーホームファイル。工場出荷既定はリポジトリ {@code init_setting} の
 * {@link InitSettingPaths#juchuHeaderAliasesFileForFactory(FactorySite)} に書き出す。
 */
public final class JuchuHeaderAliasRegistry {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static final Path LEGACY_STORE_PATH =
            Path.of(
                    System.getProperty("user.home"),
                    ".pm-ai-desktop",
                    "request-form-juchu-header-aliases.properties");

    private static final String KIND_ALIAS = "alias";
    private static final String KIND_EXPECTED = "expected";
    private static final String KIND_EXPECTED_PICK = "expectedPick";
    private static final String KIND_IGNORED = "ignored";
    private static final String KIND_EXCLUDE = "exclude";
    private static final String KIND_HEADER_ROW = "headerRow";
    private static final String FACTORY_DEFAULT_PATH_KEY = "@default";

    /** 受注ﾌｧｲﾙ見出し行の既定（1-based）。 */
    public static final int DEFAULT_HEADER_ROW_ONE_BASED = 3;

    private final Path storePath;

    private final Map<String, Map<String, Set<String>>> aliasesByFilePath = new LinkedHashMap<>();
    private final Map<String, Map<String, String>> expectedOverrideByFilePath = new LinkedHashMap<>();
    /** REDEFINE 時にユーザーが選んだ {@code XX列: 見出し} 表示ラベル。 */
    private final Map<String, Map<String, String>> expectedPickLabelByFilePath = new LinkedHashMap<>();
    private final Map<String, Set<String>> ignoredUnknownColumnLettersByFilePath = new LinkedHashMap<>();
    private final Map<String, Set<String>> excludedKnownColumnsByFilePath = new LinkedHashMap<>();
    /** 受注ファイル絶対パス → 見出し行（1-based）。 */
    private final Map<String, Integer> headerRowOneBasedByFilePath = new LinkedHashMap<>();
    private int factoryDefaultHeaderRowOneBased = DEFAULT_HEADER_ROW_ONE_BASED;

    public JuchuHeaderAliasRegistry(Path storePath) {
        this.storePath = storePath != null ? storePath : storePathForFactory(FactorySite.KONAN);
    }

    /** テスト用（湖南・工場別ストアパス）。 */
    JuchuHeaderAliasRegistry() {
        this(storePathForFactory(FactorySite.KONAN));
    }

    public static JuchuHeaderAliasRegistry loadDefault() {
        return loadForFactory(FactorySite.KONAN, Map.of());
    }

    public static JuchuHeaderAliasRegistry loadForFactory(FactorySite site, Map<String, String> ui) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        Path path = storePathForFactory(effective);
        migrateLegacyStoreIfNeeded(path, effective);
        JuchuHeaderAliasRegistry registry = new JuchuHeaderAliasRegistry(path);
        registry.reloadFromDisk();
        if (registry.isEmpty() && ui != null && !ui.isEmpty()) {
            Path initSetting =
                    InitSettingPaths.resolveRepoInitSettingDir(ui)
                            .resolve(InitSettingPaths.juchuHeaderAliasesFileForFactory(effective));
            if (Files.isRegularFile(initSetting)) {
                try {
                    registry.importFromJsonFile(initSetting);
                    registry.saveToDisk();
                } catch (IOException ignored) {
                }
            }
        }
        return registry;
    }

    public Path storePath() {
        return storePath;
    }

    public synchronized boolean isEmpty() {
        return aliasesByFilePath.isEmpty()
                && expectedOverrideByFilePath.isEmpty()
                && expectedPickLabelByFilePath.isEmpty()
                && ignoredUnknownColumnLettersByFilePath.isEmpty()
                && excludedKnownColumnsByFilePath.isEmpty()
                && headerRowOneBasedByFilePath.isEmpty()
                && factoryDefaultHeaderRowOneBased == DEFAULT_HEADER_ROW_ONE_BASED;
    }

    public synchronized void reloadFromDisk() {
        aliasesByFilePath.clear();
        expectedOverrideByFilePath.clear();
        expectedPickLabelByFilePath.clear();
        ignoredUnknownColumnLettersByFilePath.clear();
        excludedKnownColumnsByFilePath.clear();
        headerRowOneBasedByFilePath.clear();
        factoryDefaultHeaderRowOneBased = DEFAULT_HEADER_ROW_ONE_BASED;
        if (!Files.isRegularFile(storePath)) {
            return;
        }
        PropertiesLike props = PropertiesLike.read(storePath);
        applyPropertiesEntries(props.entries());
    }

    private synchronized void applyPropertiesEntries(Map<String, String> entries) {
        for (Map.Entry<String, String> entry : entries.entrySet()) {
            String fileKey = entry.getKey();
            String value = entry.getValue();
            if (fileKey == null || fileKey.isBlank() || value == null || value.isBlank()) {
                continue;
            }
            int firstSep = fileKey.indexOf('|');
            if (firstSep <= 0 || firstSep >= fileKey.length() - 1) {
                continue;
            }
            String path = fileKey.substring(0, firstSep).trim();
            String remainder = fileKey.substring(firstSep + 1).trim();
            if (KIND_HEADER_ROW.equals(remainder)) {
                int rowOneBased = parseHeaderRowOneBased(value);
                if (FACTORY_DEFAULT_PATH_KEY.equals(path)) {
                    factoryDefaultHeaderRowOneBased = rowOneBased;
                } else {
                    headerRowOneBasedByFilePath.put(normalizePathKey(path), rowOneBased);
                }
                continue;
            }
            int secondSep = remainder.indexOf('|');
            if (secondSep < 0) {
                putAliases(path, remainder, value);
                continue;
            }
            String kind = remainder.substring(0, secondSep).trim();
            String colName = remainder.substring(secondSep + 1).trim();
            if (colName.isBlank()) {
                continue;
            }
            if (KIND_EXPECTED.equals(kind)) {
                expectedOverrideByFilePath
                        .computeIfAbsent(normalizePathKey(path), k -> new LinkedHashMap<>())
                        .put(colName, value.strip());
            } else if (KIND_EXPECTED_PICK.equals(kind)) {
                expectedPickLabelByFilePath
                        .computeIfAbsent(normalizePathKey(path), k -> new LinkedHashMap<>())
                        .put(colName, value.strip());
            } else if (KIND_ALIAS.equals(kind)) {
                putAliases(path, colName, value);
            } else if (KIND_IGNORED.equals(kind)) {
                ignoredUnknownColumnLettersByFilePath
                        .computeIfAbsent(normalizePathKey(path), k -> new LinkedHashSet<>())
                        .add(colName);
            } else if (KIND_EXCLUDE.equals(kind)) {
                excludedKnownColumnsByFilePath
                        .computeIfAbsent(normalizePathKey(path), k -> new LinkedHashSet<>())
                        .add(colName);
            }
        }
    }

    public synchronized void importFromJsonFile(Path jsonFile) throws IOException {
        JsonNode root = JSON.readTree(jsonFile.toFile());
        importFromJsonNode(root);
    }

    public synchronized void importFromJsonNode(JsonNode root) {
        if (root == null || !root.isObject()) {
            return;
        }
        JsonNode defaultHeaderRow = root.get("defaultHeaderRowOneBased");
        if (defaultHeaderRow != null && defaultHeaderRow.isInt()) {
            factoryDefaultHeaderRowOneBased = parseHeaderRowOneBased(defaultHeaderRow.asInt());
        }
        JsonNode entries = root.get("entries");
        if (entries != null && entries.isObject()) {
            PropertiesLike props = new PropertiesLike();
            Iterator<String> fn = entries.fieldNames();
            while (fn.hasNext()) {
                String key = fn.next();
                JsonNode valueNode = entries.get(key);
                if (valueNode != null && valueNode.isTextual()) {
                    props.put(key, valueNode.asText(""));
                }
            }
            applyPropertiesEntries(props.entries());
            return;
        }
        Iterator<String> fn = root.fieldNames();
        PropertiesLike props = new PropertiesLike();
        while (fn.hasNext()) {
            String key = fn.next();
            JsonNode valueNode = root.get(key);
            if (valueNode != null && valueNode.isTextual()) {
                props.put(key, valueNode.asText(""));
            }
        }
        applyPropertiesEntries(props.entries());
    }

    public synchronized void exportToJsonFile(Path jsonFile) throws IOException {
        Files.createDirectories(jsonFile.getParent());
        JSON.writerWithDefaultPrettyPrinter().writeValue(jsonFile.toFile(), toJsonObject());
    }

    public synchronized ObjectNode toJsonObject() {
        ObjectNode root = JSON.createObjectNode();
        if (factoryDefaultHeaderRowOneBased != DEFAULT_HEADER_ROW_ONE_BASED) {
            root.put("defaultHeaderRowOneBased", factoryDefaultHeaderRowOneBased);
        }
        ObjectNode entries = root.putObject("entries");
        PropertiesLike props = buildPropertiesSnapshot();
        for (Map.Entry<String, String> entry : props.entries().entrySet()) {
            entries.put(entry.getKey(), entry.getValue());
        }
        return root;
    }

    private synchronized PropertiesLike buildPropertiesSnapshot() {
        PropertiesLike props = new PropertiesLike();
        for (Map.Entry<String, Map<String, Set<String>>> fileEntry : aliasesByFilePath.entrySet()) {
            String pathKey = fileEntry.getKey();
            for (Map.Entry<String, Set<String>> colEntry : fileEntry.getValue().entrySet()) {
                String key = pathKey + "|" + KIND_ALIAS + "|" + colEntry.getKey();
                props.put(key, String.join("\u0001", colEntry.getValue()));
            }
        }
        for (Map.Entry<String, Map<String, String>> fileEntry : expectedOverrideByFilePath.entrySet()) {
            String pathKey = fileEntry.getKey();
            for (Map.Entry<String, String> colEntry : fileEntry.getValue().entrySet()) {
                String key = pathKey + "|" + KIND_EXPECTED + "|" + colEntry.getKey();
                props.put(key, colEntry.getValue());
            }
        }
        for (Map.Entry<String, Map<String, String>> fileEntry : expectedPickLabelByFilePath.entrySet()) {
            String pathKey = fileEntry.getKey();
            for (Map.Entry<String, String> colEntry : fileEntry.getValue().entrySet()) {
                String key = pathKey + "|" + KIND_EXPECTED_PICK + "|" + colEntry.getKey();
                props.put(key, colEntry.getValue());
            }
        }
        for (Map.Entry<String, Set<String>> fileEntry : ignoredUnknownColumnLettersByFilePath.entrySet()) {
            String pathKey = fileEntry.getKey();
            for (String letter : fileEntry.getValue()) {
                String key = pathKey + "|" + KIND_IGNORED + "|" + letter;
                props.put(key, "1");
            }
        }
        for (Map.Entry<String, Set<String>> fileEntry : excludedKnownColumnsByFilePath.entrySet()) {
            String pathKey = fileEntry.getKey();
            for (String colName : fileEntry.getValue()) {
                String key = pathKey + "|" + KIND_EXCLUDE + "|" + colName;
                props.put(key, "1");
            }
        }
        if (factoryDefaultHeaderRowOneBased != DEFAULT_HEADER_ROW_ONE_BASED) {
            props.put(
                    FACTORY_DEFAULT_PATH_KEY + "|" + KIND_HEADER_ROW,
                    String.valueOf(factoryDefaultHeaderRowOneBased));
        }
        for (Map.Entry<String, Integer> fileEntry : headerRowOneBasedByFilePath.entrySet()) {
            props.put(
                    fileEntry.getKey() + "|" + KIND_HEADER_ROW,
                    String.valueOf(fileEntry.getValue()));
        }
        return props;
    }

    private void putAliases(String path, String colName, String value) {
        Map<String, Set<String>> perFile =
                aliasesByFilePath.computeIfAbsent(normalizePathKey(path), k -> new LinkedHashMap<>());
        Set<String> aliases = perFile.computeIfAbsent(colName, k -> new LinkedHashSet<>());
        for (String part : value.split("\u0001")) {
            String alias = part != null ? part.strip() : "";
            if (!alias.isEmpty()) {
                aliases.add(alias);
            }
        }
    }

    public synchronized Optional<String> aliasMatch(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column, String actualHeader) {
        if (juchuFileAbsolutePath == null || column == null || actualHeader == null) {
            return Optional.empty();
        }
        Map<String, Set<String>> perFile = aliasesByFilePath.get(normalizePathKey(juchuFileAbsolutePath));
        if (perFile == null) {
            return Optional.empty();
        }
        Set<String> aliases = perFile.get(column.name());
        if (aliases == null || aliases.isEmpty()) {
            return Optional.empty();
        }
        String trimmed = actualHeader.strip();
        return aliases.contains(trimmed) ? Optional.of(trimmed) : Optional.empty();
    }

    public synchronized List<String> aliasesFor(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column) {
        if (juchuFileAbsolutePath == null || column == null) {
            return List.of();
        }
        Map<String, Set<String>> perFile = aliasesByFilePath.get(normalizePathKey(juchuFileAbsolutePath));
        if (perFile == null) {
            return List.of();
        }
        Set<String> aliases = perFile.get(column.name());
        return aliases == null || aliases.isEmpty() ? List.of() : List.copyOf(aliases);
    }

    public synchronized List<String> extraAliasesFor(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column) {
        return aliasesFor(juchuFileAbsolutePath, column);
    }

    public synchronized Optional<String> expectedOverrideFor(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column) {
        if (juchuFileAbsolutePath == null || column == null) {
            return Optional.empty();
        }
        Map<String, String> perFile = expectedOverrideByFilePath.get(normalizePathKey(juchuFileAbsolutePath));
        if (perFile == null || !perFile.containsKey(column.name())) {
            return Optional.empty();
        }
        return Optional.ofNullable(perFile.get(column.name()));
    }

    public synchronized String expectedHeaderFor(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column) {
        if (juchuFileAbsolutePath == null || column == null) {
            return column != null ? column.primaryHeader() : "";
        }
        Map<String, String> perFile = expectedOverrideByFilePath.get(normalizePathKey(juchuFileAbsolutePath));
        if (perFile == null) {
            return column.primaryHeader();
        }
        return perFile.getOrDefault(column.name(), column.primaryHeader());
    }

    public synchronized void addAlias(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column, String aliasHeader) {
        if (juchuFileAbsolutePath == null
                || juchuFileAbsolutePath.isBlank()
                || column == null
                || aliasHeader == null
                || aliasHeader.isBlank()) {
            return;
        }
        String pathKey = normalizePathKey(juchuFileAbsolutePath);
        Map<String, Set<String>> perFile =
                aliasesByFilePath.computeIfAbsent(pathKey, k -> new LinkedHashMap<>());
        perFile.computeIfAbsent(column.name(), k -> new LinkedHashSet<>()).add(aliasHeader.strip());
    }

    public synchronized void setExpectedOverride(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column, String expectedHeader) {
        if (juchuFileAbsolutePath == null || juchuFileAbsolutePath.isBlank() || column == null) {
            return;
        }
        expectedOverrideByFilePath
                .computeIfAbsent(normalizePathKey(juchuFileAbsolutePath), k -> new LinkedHashMap<>())
                .put(column.name(), expectedHeader != null ? expectedHeader.strip() : "");
    }

    public synchronized Optional<String> expectedPickLabelFor(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column) {
        if (juchuFileAbsolutePath == null || column == null) {
            return Optional.empty();
        }
        Map<String, String> perFile =
                expectedPickLabelByFilePath.get(normalizePathKey(juchuFileAbsolutePath));
        if (perFile == null || !perFile.containsKey(column.name())) {
            return Optional.empty();
        }
        return Optional.ofNullable(perFile.get(column.name()));
    }

    public synchronized void setExpectedPickLabel(
            String juchuFileAbsolutePath,
            JuchuSheetColumnLayout.Col column,
            String pickDisplayLabel) {
        if (juchuFileAbsolutePath == null || juchuFileAbsolutePath.isBlank() || column == null) {
            return;
        }
        String pathKey = normalizePathKey(juchuFileAbsolutePath);
        if (pickDisplayLabel == null || pickDisplayLabel.isBlank()) {
            Map<String, String> perFile = expectedPickLabelByFilePath.get(pathKey);
            if (perFile != null) {
                perFile.remove(column.name());
            }
            return;
        }
        expectedPickLabelByFilePath
                .computeIfAbsent(pathKey, k -> new LinkedHashMap<>())
                .put(column.name(), pickDisplayLabel.strip());
    }

    public synchronized boolean isUnknownColumnIgnored(
            String juchuFileAbsolutePath, String columnLetter) {
        if (juchuFileAbsolutePath == null || columnLetter == null || columnLetter.isBlank()) {
            return false;
        }
        Set<String> ignored =
                ignoredUnknownColumnLettersByFilePath.get(normalizePathKey(juchuFileAbsolutePath));
        return ignored != null && ignored.contains(columnLetter.strip().toUpperCase(Locale.ROOT));
    }

    public synchronized void setUnknownColumnIgnored(
            String juchuFileAbsolutePath, String columnLetter) {
        if (juchuFileAbsolutePath == null || juchuFileAbsolutePath.isBlank() || columnLetter == null) {
            return;
        }
        ignoredUnknownColumnLettersByFilePath
                .computeIfAbsent(normalizePathKey(juchuFileAbsolutePath), k -> new LinkedHashSet<>())
                .add(columnLetter.strip().toUpperCase(Locale.ROOT));
    }

    public synchronized void clearUnknownColumnIgnored(
            String juchuFileAbsolutePath, String columnLetter) {
        if (juchuFileAbsolutePath == null || columnLetter == null || columnLetter.isBlank()) {
            return;
        }
        Set<String> ignored =
                ignoredUnknownColumnLettersByFilePath.get(normalizePathKey(juchuFileAbsolutePath));
        if (ignored != null) {
            ignored.remove(columnLetter.strip().toUpperCase(Locale.ROOT));
        }
    }

    public synchronized boolean isExcludedFromTransfer(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column) {
        if (juchuFileAbsolutePath == null || column == null) {
            return false;
        }
        Set<String> excluded =
                excludedKnownColumnsByFilePath.get(normalizePathKey(juchuFileAbsolutePath));
        return excluded != null && excluded.contains(column.name());
    }

    public synchronized void setExcludedFromTransfer(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column) {
        if (juchuFileAbsolutePath == null || juchuFileAbsolutePath.isBlank() || column == null) {
            return;
        }
        excludedKnownColumnsByFilePath
                .computeIfAbsent(normalizePathKey(juchuFileAbsolutePath), k -> new LinkedHashSet<>())
                .add(column.name());
    }

    public synchronized void clearExcludedFromTransfer(
            String juchuFileAbsolutePath, JuchuSheetColumnLayout.Col column) {
        if (juchuFileAbsolutePath == null || column == null) {
            return;
        }
        Set<String> excluded =
                excludedKnownColumnsByFilePath.get(normalizePathKey(juchuFileAbsolutePath));
        if (excluded != null) {
            excluded.remove(column.name());
        }
    }

    /** 見出し行（1-based）。ファイル別設定が無ければ工場既定（既定 3 行目）。 */
    public synchronized int headerRowOneBasedFor(String juchuFileAbsolutePath) {
        if (juchuFileAbsolutePath == null || juchuFileAbsolutePath.isBlank()) {
            return factoryDefaultHeaderRowOneBased;
        }
        return headerRowOneBasedByFilePath.getOrDefault(
                normalizePathKey(juchuFileAbsolutePath), factoryDefaultHeaderRowOneBased);
    }

    /** 見出し行（0-based、POI {@link Row} index）。 */
    public synchronized int headerRowIndexFor(String juchuFileAbsolutePath) {
        return Math.max(0, headerRowOneBasedFor(juchuFileAbsolutePath) - 1);
    }

    public synchronized int factoryDefaultHeaderRowOneBased() {
        return factoryDefaultHeaderRowOneBased;
    }

    public synchronized void setFactoryDefaultHeaderRowOneBased(int rowOneBased) {
        factoryDefaultHeaderRowOneBased = parseHeaderRowOneBased(rowOneBased);
    }

    public synchronized void setHeaderRowOneBasedFor(
            String juchuFileAbsolutePath, int rowOneBased) {
        if (juchuFileAbsolutePath == null || juchuFileAbsolutePath.isBlank()) {
            return;
        }
        int clamped = parseHeaderRowOneBased(rowOneBased);
        String pathKey = normalizePathKey(juchuFileAbsolutePath);
        if (clamped == factoryDefaultHeaderRowOneBased) {
            headerRowOneBasedByFilePath.remove(pathKey);
        } else {
            headerRowOneBasedByFilePath.put(pathKey, clamped);
        }
    }

    private static int parseHeaderRowOneBased(String value) {
        if (value == null || value.isBlank()) {
            return DEFAULT_HEADER_ROW_ONE_BASED;
        }
        try {
            return parseHeaderRowOneBased(Integer.parseInt(value.strip()));
        } catch (NumberFormatException ex) {
            return DEFAULT_HEADER_ROW_ONE_BASED;
        }
    }

    private static int parseHeaderRowOneBased(int rowOneBased) {
        return Math.max(1, Math.min(200, rowOneBased));
    }

    public synchronized void replaceFromInitSetting(Path initSettingJson) throws IOException {
        aliasesByFilePath.clear();
        expectedOverrideByFilePath.clear();
        expectedPickLabelByFilePath.clear();
        ignoredUnknownColumnLettersByFilePath.clear();
        excludedKnownColumnsByFilePath.clear();
        headerRowOneBasedByFilePath.clear();
        factoryDefaultHeaderRowOneBased = DEFAULT_HEADER_ROW_ONE_BASED;
        if (initSettingJson != null && Files.isRegularFile(initSettingJson)) {
            importFromJsonFile(initSettingJson);
        }
        saveToDisk();
    }

    public synchronized void saveToDisk() throws IOException {
        Files.createDirectories(storePath.getParent());
        buildPropertiesSnapshot().write(storePath);
    }

    public static Path storePathForFactory(FactorySite site) {
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        String suffix = effective.name().toLowerCase(Locale.ROOT);
        return Path.of(
                System.getProperty("user.home"),
                ".pm-ai-desktop",
                "request-form-juchu-header-aliases_" + suffix + ".properties");
    }

    private static void migrateLegacyStoreIfNeeded(Path targetPath, FactorySite site) {
        if (site != FactorySite.KONAN || Files.isRegularFile(targetPath)) {
            return;
        }
        if (!Files.isRegularFile(LEGACY_STORE_PATH)) {
            return;
        }
        try {
            Files.createDirectories(targetPath.getParent());
            Files.copy(LEGACY_STORE_PATH, targetPath, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException ignored) {
        }
    }

    private static String normalizePathKey(String absolutePath) {
        try {
            return Path.of(absolutePath).toAbsolutePath().normalize().toString();
        } catch (Exception ex) {
            return absolutePath.strip();
        }
    }

    /** UTF-8 properties の最小実装（エスケープは改行・\\のみ）。 */
    private static final class PropertiesLike {
        private final Map<String, String> map = new LinkedHashMap<>();

        void put(String key, String value) {
            map.put(key, value);
        }

        Map<String, String> entries() {
            return map;
        }

        static PropertiesLike read(Path path) {
            PropertiesLike props = new PropertiesLike();
            try (InputStream in = Files.newInputStream(path)) {
                String text = new String(in.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
                for (String line : text.split("\n")) {
                    String trimmed = line.strip();
                    if (trimmed.isEmpty() || trimmed.startsWith("#")) {
                        continue;
                    }
                    int eq = trimmed.indexOf('=');
                    if (eq <= 0) {
                        continue;
                    }
                    String key = unescape(trimmed.substring(0, eq).strip());
                    String value = unescape(trimmed.substring(eq + 1).strip());
                    props.put(key, value);
                }
            } catch (IOException ex) {
                return props;
            }
            return props;
        }

        void write(Path path) throws IOException {
            StringBuilder sb = new StringBuilder();
            sb.append("# request-form juchu header aliases / expected overrides (UTF-8)\n");
            for (Map.Entry<String, String> entry : map.entrySet()) {
                sb.append(escape(entry.getKey()))
                        .append('=')
                        .append(escape(entry.getValue()))
                        .append('\n');
            }
            try (OutputStream out = Files.newOutputStream(path)) {
                out.write(sb.toString().getBytes(java.nio.charset.StandardCharsets.UTF_8));
            }
        }

        private static String escape(String text) {
            if (text == null) {
                return "";
            }
            return text.replace("\\", "\\\\").replace("\n", "\\n");
        }

        private static String unescape(String text) {
            if (text == null) {
                return "";
            }
            return text.replace("\\n", "\n").replace("\\\\", "\\");
        }
    }
}
