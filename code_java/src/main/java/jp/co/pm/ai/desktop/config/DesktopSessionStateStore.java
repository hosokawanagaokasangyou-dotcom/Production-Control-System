package jp.co.pm.ai.desktop.config;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;

import jp.co.pm.ai.desktop.reconciliation.RequestFormComboChoices;

/**
 * Persists last-used paths under {@code ~/.pm-ai-desktop/session-state.json} so tabs reload the same files on
 * the next launch.
 *
 * <p>初回などセッションファイルが無いときは、{@code pm-ai-data/config/bundled_session_ui_defaults.json}（メインシェル
 * タブ順・見出しグロー、設備ガント表示、バッジ等）またはクラスパス上の同名リソースを読み込む。旧ファイル名
 * {@code bundled_session_badge_defaults.json} が残っているインストーラでも読み込める。
 *
 * <p>ポータブル自動バージョンアップ後は {@link #applyPortableUpgradeBundledPolicyToSessionStore()} で同ファイルのキーを
 * {@code ~/.pm-ai-desktop/session-state.json} に上書きし、配台不要 JSON パスも正本に合わせる。
 */
public final class DesktopSessionStateStore {

    private static final ObjectMapper JSON = new ObjectMapper();

    private static Path storePath() {
        return AppPaths.resolveSessionStateStorePath();
    }

    private static final String BUNDLED_SESSION_UI_DEFAULTS_RESOURCE =
            "/jp/co/pm/ai/desktop/config/bundled_session_ui_defaults.json";

    private static final String LEGACY_BUNDLED_SESSION_BADGE_DEFAULTS_RESOURCE =
            "/jp/co/pm/ai/desktop/config/bundled_session_badge_defaults.json";

    private DesktopSessionStateStore() {}

    public static DesktopSessionState load() {
        try {
            JsonNode root;
            if (Files.isRegularFile(storePath())) {
                root = JSON.readTree(storePath().toFile());
                if (root == null || !root.isObject()) {
                    return DesktopSessionState.empty();
                }
            } else {
                root = sessionRootWhenStoreFileMissing();
            }
            return parseDesktopSessionState(root);
        } catch (IOException e) {
            return DesktopSessionState.empty();
        }
    }

    /**
     * {@code session-state.json} が無いときの JSON ルート。クラスパス・{@code pm-ai-data/config}・{@code init_setting}
     * をマージした UI フラグメント。
     */
    private static JsonNode sessionRootWhenStoreFileMissing() {
        JsonNode merged = readMergedSessionUiDefaultsNode();
        if (merged != null && merged.isObject() && merged.size() > 0) {
            return merged;
        }
        return JSON.createObjectNode();
    }

    /**
     * 優先順（後勝ち）: クラスパス {@code bundled_session_ui_defaults.json} → 旧バッジ既定 → {@code
     * pm-ai-data/config/bundled_session_ui_defaults.json} → 同 {@code bundled_session_badge_defaults.json} →
     * {@code init_setting/session_defaults.json}（カレントディレクトリ基準）→ {@code
     * pm-ai-data/init_setting/session_defaults.json}。リポジトリ側は {@link #readMergedSessionUiDefaultsNode(Map)} を参照。
     */
    static JsonNode readMergedSessionUiDefaultsNode() {
        return readMergedSessionUiDefaultsNode(null);
    }

    /**
     * {@link #readMergedSessionUiDefaultsNode()} と同じマージに加え、{@code ui} が非 null のとき
     * {@link InitSettingPaths#resolveRepoInitSettingDir(Map)}/{@link InitSettingPaths#SESSION_DEFAULTS_FILE} を重ね、続けて
     * {@link GlobalInitSettingTarget} に応じた {@link InitSettingPaths#sessionDefaultsFileForFactory(FactorySite)} を
     * <strong>最後</strong>に重ねる（工場別の上書き。無い場合は無視）。グローバル設定の書き出し先と工場リセットの参照先を揃える。
     */
    static JsonNode readMergedSessionUiDefaultsNode(Map<String, String> ui) {
        ObjectNode acc = JSON.createObjectNode();
        mergeSessionUiFromClasspath(acc, BUNDLED_SESSION_UI_DEFAULTS_RESOURCE);
        mergeSessionUiFromClasspath(acc, LEGACY_BUNDLED_SESSION_BADGE_DEFAULTS_RESOURCE);
        Path cfgDir =
                Path.of(System.getProperty("user.dir", "."))
                        .toAbsolutePath()
                        .normalize()
                        .resolve("pm-ai-data")
                        .resolve("config");
        mergeSessionUiFromPath(acc, cfgDir.resolve("bundled_session_ui_defaults.json"));
        mergeSessionUiFromPath(acc, cfgDir.resolve("bundled_session_badge_defaults.json"));
        mergeSessionUiFromPath(
                acc, InitSettingPaths.cwdInitSettingDir().resolve(InitSettingPaths.SESSION_DEFAULTS_FILE));
        mergeSessionUiFromPath(
                acc,
                InitSettingPaths.portableBundleInitSettingDir()
                        .resolve(InitSettingPaths.SESSION_DEFAULTS_FILE));
        if (ui != null) {
            Path repoInit = InitSettingPaths.resolveRepoInitSettingDir(ui);
            mergeSessionUiFromPath(acc, repoInit.resolve(InitSettingPaths.SESSION_DEFAULTS_FILE));
            FactorySite g = GlobalInitSettingTarget.load();
            mergeSessionUiFromPath(acc, repoInit.resolve(InitSettingPaths.sessionDefaultsFileForFactory(g)));
        }
        return acc.size() > 0 ? acc : null;
    }

    /** 工場別 {@code init_setting} から依頼書フォーム ComboBox 候補を読む。 */
    public static RequestFormComboChoices loadFactoryRequestFormComboChoices(Map<String, String> ui) {
        return loadFactoryRequestFormComboChoices(ui, GlobalInitSettingTarget.load());
    }

    public static RequestFormComboChoices loadFactoryRequestFormComboChoices(
            Map<String, String> ui, FactorySite site) {
        JsonNode merged = readMergedSessionUiDefaultsNode(ui);
        if (merged == null || !merged.isObject()) {
            return RequestFormComboChoices.empty();
        }
        FactorySite effective = site != null ? site : FactorySite.KONAN;
        Path factoryFile =
                ui != null
                        ? InitSettingPaths.resolveRepoInitSettingDir(ui)
                                .resolve(InitSettingPaths.sessionDefaultsFileForFactory(effective))
                        : null;
        if (factoryFile != null && Files.isRegularFile(factoryFile)) {
            try {
                JsonNode factoryRoot = JSON.readTree(factoryFile.toFile());
                RequestFormComboChoices fromFactory = RequestFormComboChoices.fromJson(factoryRoot);
                if (!fromFactory.isEmpty()) {
                    return fromFactory;
                }
            } catch (IOException ignored) {
            }
        }
        return RequestFormComboChoices.fromJson(merged);
    }

    /**
     * 依頼書入力【設定】タブの「工場出荷状態に戻す」用。{@code init_setting} の工場別既定を優先し、
     * 無ければマージ済み bundled、それも空なら {@link RequestFormComboChoices#bundledDefaults()}。
     */
    public static RequestFormComboChoices factoryShipmentRequestFormComboChoices(
            Map<String, String> ui, FactorySite site) {
        RequestFormComboChoices fromInit = loadFactoryRequestFormComboChoices(ui, site);
        if (!fromInit.isEmpty()) {
            return fromInit.mergedWithDefaults();
        }
        return RequestFormComboChoices.bundledDefaults();
    }

    public static Path factoryJuchuHeaderAliasesPath(Map<String, String> ui, FactorySite site) {
        FactorySite effective = site != null ? site : GlobalInitSettingTarget.load();
        if (ui == null || ui.isEmpty()) {
            return null;
        }
        return InitSettingPaths.resolveRepoInitSettingDir(ui)
                .resolve(InitSettingPaths.juchuHeaderAliasesFileForFactory(effective));
    }

    private static void mergeSessionUiFromClasspath(ObjectNode acc, String resourcePath) {
        JsonNode n = readBundledJsonFromClasspath(resourcePath);
        if (n != null && n.isObject()) {
            deepMergeInto(acc, (ObjectNode) n);
        }
    }

    private static void mergeSessionUiFromPath(ObjectNode acc, Path file) {
        try {
            if (!Files.isRegularFile(file)) {
                return;
            }
            JsonNode n = JSON.readTree(file.toFile());
            if (n != null && n.isObject()) {
                deepMergeInto(acc, (ObjectNode) n);
            }
        } catch (IOException ignored) {
        }
    }

    private static void deepMergeInto(ObjectNode acc, ObjectNode overlay) {
        Iterator<String> fn = overlay.fieldNames();
        while (fn.hasNext()) {
            String k = fn.next();
            JsonNode v = overlay.get(k);
            if (v != null) {
                acc.set(k, v.deepCopy());
            }
        }
    }

    /**
     * 工場出荷 UI リセット用: マージ済み既定に対し、環境・実行パスだけ {@code bootstrap} を適用する。
     *
     * @param ui 非 null のとき {@code resolveRepoInitSettingDir(ui)} 配下の {@code session_defaults.json} と工場別
     *     {@link InitSettingPaths#sessionDefaultsFileForFactory(FactorySite)} をマージ最終層に含め、グローバル設定で書き出した UI
     *     既定に近づける
     */
    public static DesktopSessionState buildFactoryResetSession(
            DesktopSessionState bootstrap, Map<String, String> ui) {
        JsonNode merged = readMergedSessionUiDefaultsNode(ui);
        ObjectNode root =
                merged != null && merged.isObject()
                        ? (ObjectNode) merged
                        : JSON.createObjectNode();
        DesktopSessionState factory = parseDesktopSessionState(root);
        return factory.withBootstrapFieldsFrom(bootstrap);
    }

    /** {@link #save(DesktopSessionState)} と同一形状の JSON（{@code init_setting} 書き出し用）。 */
    public static ObjectNode toJsonObject(DesktopSessionState state) {
        ObjectNode root = JSON.createObjectNode();
        put(root, "planInputPath", state.planInputPath());
        put(root, "planInputSheet", state.planInputSheet());
        put(root, "stage1PreviewPath", state.stage1PreviewPath());
        put(root, "stage1PreviewSheet", state.stage1PreviewSheet());
        put(root, "excludeRulesPath", state.excludeRulesPath());
        put(root, "mainRunWorkbook", state.mainRunWorkbook());
        put(root, "mainRunScriptDir", state.mainRunScriptDir());
        put(root, "uiTheme", state.uiTheme());
        put(root, "logFontFamily", state.logFontFamily());
        putLogFontSize(root, state.logFontSize());
        put(root, "mainRunLogFilter", state.mainRunLogFilter());
        putMainRunLogLines(root, state.mainRunLogLines());
        putMainRunLogScroll(root, state.mainRunLogScroll());
        put(root, "mainRunStage2ProductionPlan", state.mainRunStage2ProductionPlan());
        put(root, "mainRunStage2MemberSchedule", state.mainRunStage2MemberSchedule());
        root.put("mainRunStage2SkipTodayDispatch", state.mainRunStage2SkipTodayDispatch());
        root.put(
                "planInputStage2InProgressNextDayPrompt",
                state.planInputStage2InProgressNextDayPrompt());
        put(root, "mainRunStage2ResultBookFont", state.mainRunStage2ResultBookFont());
        root.put("mainRunSkipGeminiApi", state.mainRunSkipGeminiApi());
        root.put("mainRunStage1MarkAllExcludeAfterRun", state.mainRunStage1MarkAllExcludeAfterRun());
        root.put("mainRunApplyLearnedSpeedFromActuals", state.mainRunApplyLearnedSpeedFromActuals());
        putUiEnvRows(root, state.uiEnvRows());
        putMainShellTabOrder(root, state.mainShellTabOrder());
        putMainShellTabLayout(root, state.mainShellTabLayout());
        putStringStringMap(root, "mainShellTabTitleAliases", state.mainShellTabTitleAliases());
        putStringIntMap(root, "innerTabSelectedIndexByShellTabKey", state.innerTabSelectedIndexByShellTabKey());
        putStringStringMap(root, "innerTabHeaderColorByKey", state.innerTabHeaderColorByKey());
        putEquipmentGanttGraphicPrefs(root, state);
        putStage1NetworkCacheBadgePrefs(root, state);
        putRequestFormPreviewUpdateBadgePrefs(root, state);
        root.put("mainShellTabOrganizerHeaderGlow", state.mainShellTabOrganizerHeaderGlow());
        root.put(
                "mainShellTabOrganizerHeaderGlowStrength",
                state.mainShellTabOrganizerHeaderGlowStrength());
        putPushButtonDesignPrefs(root, state);
        putMemorySettingsPrefs(root, state);
        putEquipmentStatusDashboardPrefs(root, state);
        putWindowGeometry(root, state);
        // 依頼書 ComboBox 候補の正本はサマリ Excel 同フォルダの request_form_input_settings.json
        return root;
    }

    /**
     * ポータブル自動バージョンアップ（正本→ローカル {@code pm-ai-data} 同期）のあと、同梱された
     * {@code bundled_session_ui_defaults.json} の各キーを {@link #storePath()} に上書きし、
     * {@code pm-ai-data/code/exclude_rules.json} があれば {@code excludeRulesPath} と環境タブ相当の
     * {@link AppPaths#KEY_PM_AI_EXCLUDE_RULES_JSON} 行も最新の絶対パスへ更新する。
     *
     * <p>{@link #applyPortableUpgradeBundledPolicyToSessionStore(Map)} に {@code null} を渡すのと同じ。
     */
    public static void applyPortableUpgradeBundledPolicyToSessionStore() throws IOException {
        applyPortableUpgradeBundledPolicyToSessionStore(null);
    }

    /**
     * {@link #applyPortableUpgradeBundledPolicyToSessionStore()} と同様だが、{@code ui} が非 null のときは
     * {@link #readMergedSessionUiDefaultsNode(Map)} を用い、リポジトリ {@code init_setting/session_defaults.json} と
     * 工場別 {@link InitSettingPaths#sessionDefaultsFileForFactory}（アップデートで配布側が上書きした内容）をマージの最終層に含める。
     */
    public static void applyPortableUpgradeBundledPolicyToSessionStore(Map<String, String> ui) throws IOException {
        JsonNode bundled = readMergedSessionUiDefaultsNode(ui);
        ObjectNode root;
        if (Files.isRegularFile(storePath())) {
            JsonNode cur = JSON.readTree(storePath().toFile());
            root = cur != null && cur.isObject() ? (ObjectNode) cur : JSON.createObjectNode();
        } else {
            root = JSON.createObjectNode();
        }
        if (bundled != null && bundled.isObject()) {
            ObjectNode bo = (ObjectNode) bundled;
            Iterator<String> fn = bo.fieldNames();
            while (fn.hasNext()) {
                String k = fn.next();
                JsonNode v = bo.get(k);
                if (v != null) {
                    root.set(k, v.deepCopy());
                }
            }
            if (bo.has("mainShellTabOrder")) {
                root.remove("mainShellTabLayout");
            }
        }
        Path cwd = Path.of(System.getProperty("user.dir", ".")).toAbsolutePath().normalize();
        Path exclude = cwd.resolve("pm-ai-data").resolve("code").resolve("exclude_rules.json");
        if (Files.isRegularFile(exclude)) {
            String abs = exclude.toAbsolutePath().normalize().toString();
            root.put("excludeRulesPath", abs);
            putOrUpdateUiEnvRowValueInSessionRoot(root, AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON, abs);
        }
        // init_setting の開発機 BOX パス等を持ち込まない（未設定時は実行時フォールバック）
        putOrUpdateUiEnvRowValueInSessionRoot(root, AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR, "");
        Files.createDirectories(storePath().getParent());
        JSON.writerWithDefaultPrettyPrinter().writeValue(storePath().toFile(), root);
    }

    private static void putOrUpdateUiEnvRowValueInSessionRoot(ObjectNode root, String key, String value) {
        JsonNode arrNode = root.get("uiEnvRows");
        ArrayNode arr;
        if (arrNode != null && arrNode.isArray()) {
            arr = (ArrayNode) arrNode;
        } else {
            arr = JSON.createArrayNode();
            root.set("uiEnvRows", arr);
        }
        for (int i = 0; i < arr.size(); i++) {
            JsonNode el = arr.get(i);
            if (el != null && el.isObject() && key.equals(text(el, "name"))) {
                ((ObjectNode) el).put("value", value != null ? value : "");
                return;
            }
        }
        ObjectNode o = arr.addObject();
        o.put("name", key);
        o.put("value", value != null ? value : "");
        o.put("description", "");
    }

    private static JsonNode readBundledJsonFromClasspath(String resourcePath) {
        try (InputStream in = DesktopSessionStateStore.class.getResourceAsStream(resourcePath)) {
            if (in == null) {
                return null;
            }
            return JSON.readTree(in);
        } catch (IOException e) {
            return null;
        }
    }

    /**
     * ユーザープロファイル等、{@link #toJsonObject(DesktopSessionState)} 相当の JSON オブジェクトから復元する。
     */
    public static DesktopSessionState desktopSessionFromStoredJson(JsonNode root) {
        if (root == null || !root.isObject()) {
            return DesktopSessionState.empty();
        }
        return parseDesktopSessionState(root);
    }

    private static DesktopSessionState parseDesktopSessionState(JsonNode root) {
        return new DesktopSessionState(
                text(root, "planInputPath"),
                text(root, "planInputSheet"),
                text(root, "stage1PreviewPath"),
                text(root, "stage1PreviewSheet"),
                text(root, "excludeRulesPath"),
                text(root, "mainRunWorkbook"),
                text(root, "mainRunScriptDir"),
                optionalDouble(root, "windowWidth", 0d),
                optionalDouble(root, "windowHeight", 0d),
                optionalDouble(root, "windowX", Double.NaN),
                optionalDouble(root, "windowY", Double.NaN),
                text(root, "uiTheme"),
                text(root, "logFontFamily"),
                optionalDouble(root, "logFontSize", 0d),
                text(root, "mainRunLogFilter"),
                loadStringList(root, "mainRunLogLines"),
                optionalDouble(root, "mainRunLogScroll", Double.NaN),
                text(root, "mainRunStage2ProductionPlan"),
                text(root, "mainRunStage2MemberSchedule"),
                optionalBoolean(root, "mainRunStage2SkipTodayDispatch", false),
                optionalBoolean(root, "planInputStage2InProgressNextDayPrompt", true),
                text(root, "mainRunStage2ResultBookFont"),
                optionalBoolean(root, "mainRunSkipGeminiApi", false),
                optionalBoolean(root, "mainRunStage1MarkAllExcludeAfterRun", false),
                optionalBoolean(root, "mainRunApplyLearnedSpeedFromActuals", false),
                loadUiEnvRows(root),
                loadStringList(root, "mainShellTabOrder"),
                loadMainShellTabLayout(root),
                loadStringStringMap(root, "mainShellTabTitleAliases"),
                loadStringIntMap(root, "innerTabSelectedIndexByShellTabKey"),
                loadStringStringMap(root, "innerTabHeaderColorByKey"),
                optionalDouble(root, "equipmentGanttGraphicZoomPercent", 0d),
                optionalDouble(root, "equipmentGanttDateColWidth", 0d),
                optionalDouble(root, "equipmentGanttMachineColWidth", 0d),
                optionalDouble(root, "equipmentGanttProcessColWidth", 0d),
                text(root, "equipmentGanttBarFontFamily"),
                optionalDouble(root, "equipmentGanttBarFontPercent", 0d),
                optionalDouble(root, "equipmentGanttRowHeightPercent", 0d),
                optionalDouble(root, "equipmentGanttHeaderHeightPercent", 0d),
                optionalDouble(root, "equipmentGanttSlotWidthPercent", 0d),
                optionalDouble(root, "equipmentGanttShiftWheelHScrollPercent", 0d),
                optionalBoolean(
                        root,
                        "equipmentGanttPrepTimeLabelsEnabled",
                        DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PREP_TIME_LABELS_ENABLED),
                loadEquipmentGanttPersonBadgeGapPx(root),
                loadEquipmentGanttPersonBadgeBandVerticalOffsetPx(root),
                text(root, "equipmentGanttGraphicDataFingerprint"),
                loadEquipmentGanttBadgeDragDeltas(root),
                optionalBoolean(root, "equipmentGanttPersonBadgeDragAdjustEnabled", false),
                optionalBoolean(root, "equipmentGanttPersonBadgeEnabled", true),
                optionalBoolean(
                        root,
                        "equipmentGanttPersonBadgeWireEnabled",
                        DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_ENABLED),
                text(root, "equipmentGanttPersonBadgeWireStrokeHex"),
                optionalDouble(
                        root,
                        "equipmentGanttPersonBadgeWireWidthPx",
                        DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_WIDTH_PX),
                text(root, "equipmentGanttPersonBadgeWireDashStyleKey"),
                loadEquipmentGanttPersonBadgeWireMaxLengthPx(root),
                text(root, "equipmentGanttPersonBadgeFontFamily"),
                optionalDouble(root, "equipmentGanttPersonBadgeFontPercent", 0d),
                text(root, "equipmentGanttPersonBadgeFillHex"),
                text(root, "equipmentGanttPersonBadgeTextHex"),
                text(root, "equipmentGanttPersonBadgeStrokeHex"),
                optionalDouble(root, "equipmentGanttPersonBadgeStrokeWidth", -1d),
                optionalDouble(root, "equipmentGanttPersonBadgeCornerRadius", -1d),
                optionalBoolean(root, "equipmentGanttPersonBadgePill", false),
                text(root, "equipmentGanttPersonBadgeGlowColorHex"),
                optionalDouble(root, "equipmentGanttPersonBadgeGlowRadius", -1d),
                optionalDouble(root, "equipmentGanttPersonBadgeGlowSpread", -1d),
                optionalDouble(root, "equipmentGanttPersonBadgeOpacity", -1d),
                loadPersonBadgeStyleMap(root, "equipmentGanttPersonBadgeStylesByLabel"),
                loadPersonBadgeStyleMap(root, "equipmentGanttPersonBadgeStylesByMemberKey"),
                text(root, "equipmentGanttPlanJsonPath"),
                text(root, "stage1NetworkCacheBadgeLabel"),
                loadStage1NetworkCacheBadgeStyle(root),
                text(root, "requestFormPreviewUpdateBadgeLabel"),
                loadRequestFormPreviewUpdateBadgeStyle(root),
                optionalBoolean(root, "mainShellTabOrganizerHeaderGlow", true),
                clamp01(optionalDouble(root, "mainShellTabOrganizerHeaderGlowStrength", 1d)),
                loadPushButtonDesignPrefs(root),
                optionalBoolean(root, "memoryMonitorEnabled", false),
                optionalLongClamped(root, "memoryMonitorIntervalSec", 5L, 1L, 3600L),
                loadNextLaunchHeapFixed(root),
                loadNextLaunchHeapMinMiB(root),
                loadNextLaunchHeapMaxMiB(root),
                text(root, "equipmentStatusDashboardActualDate"),
                text(root, "equipmentStatusDashboardPlanDate"),
                optionalIntClamped(root, "equipmentStatusDashboardActualDayOffset", 0, -366, 366),
                optionalIntClamped(root, "equipmentStatusDashboardPlanDayOffset", 0, -366, 366),
                optionalBoolean(root, "equipmentStatusDashboardAutoRefreshEnabled", true),
                optionalIntClamped(
                        root,
                        "equipmentStatusDashboardAutoRefreshIntervalSec",
                        DesktopSessionState.DEFAULT_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                        DesktopSessionState.MIN_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                        DesktopSessionState.MAX_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC),
                optionalBoolean(root, "equipmentStatusDashboardShowAladdinPlans", true),
                optionalBoolean(root, "equipmentStatusDashboardShowDispatchPlans", true),
                loadEquipmentStatusDashboardAppearancePrefs(root),
                jp.co.pm.ai.desktop.reconciliation.RequestFormComboChoices.fromJson(root));
    }

    public static void save(DesktopSessionState state) {
        try {
            Files.createDirectories(storePath().getParent());
            JSON.writerWithDefaultPrettyPrinter().writeValue(storePath().toFile(), toJsonObject(state));
        } catch (IOException ignored) {
        }
    }

    /**
     * リモートデスクトップ配布用シェル等、環境変数タブ無し UI から env 行とテーマだけ session JSON を部分更新する。
     */
    public static void patchUiEnvMapAndTheme(Map<String, String> envMap, String uiThemeId) {
        if (envMap == null && (uiThemeId == null || uiThemeId.isBlank())) {
            return;
        }
        try {
            ObjectNode root;
            if (Files.isRegularFile(storePath())) {
                JsonNode existing = JSON.readTree(storePath().toFile());
                root =
                        existing != null && existing.isObject()
                                ? ((ObjectNode) existing).deepCopy()
                                : JSON.createObjectNode();
            } else {
                JsonNode defaults = readMergedSessionUiDefaultsNode();
                root =
                        defaults != null && defaults.isObject()
                                ? ((ObjectNode) defaults).deepCopy()
                                : JSON.createObjectNode();
            }
            if (uiThemeId != null && !uiThemeId.isBlank()) {
                root.put("uiTheme", uiThemeId.strip());
            }
            if (envMap != null && !envMap.isEmpty()) {
                ArrayNode rows =
                        root.has("uiEnvRows") && root.get("uiEnvRows").isArray()
                                ? (ArrayNode) root.get("uiEnvRows").deepCopy()
                                : JSON.createArrayNode();
                LinkedHashMap<String, ObjectNode> byKey = new LinkedHashMap<>();
                for (JsonNode row : rows) {
                    if (row != null && row.isObject() && row.has("name")) {
                        byKey.put(row.get("name").asText(""), (ObjectNode) row.deepCopy());
                    }
                }
                for (Map.Entry<String, String> e : envMap.entrySet()) {
                    String key = e.getKey() != null ? e.getKey().trim() : "";
                    if (key.isEmpty()) {
                        continue;
                    }
                    ObjectNode row = byKey.computeIfAbsent(key, k -> JSON.createObjectNode());
                    row.put("name", key);
                    row.put("value", e.getValue() != null ? e.getValue() : "");
                    if (!row.has("description")) {
                        row.put("description", "");
                    }
                }
                ArrayNode merged = JSON.createArrayNode();
                byKey.values().forEach(merged::add);
                root.set("uiEnvRows", merged);
            }
            Files.createDirectories(storePath().getParent());
            JSON.writerWithDefaultPrettyPrinter().writeValue(storePath().toFile(), root);
        } catch (IOException ignored) {
        }
    }

    /**
     * 環境変数タブの行スナップショットとテーマだけ session JSON を部分更新する（説明列を保持）。
     */
    public static void patchUiEnvRowsAndTheme(List<UiEnvRowSnapshot> rows, String uiThemeId) {
        if ((rows == null || rows.isEmpty()) && (uiThemeId == null || uiThemeId.isBlank())) {
            return;
        }
        try {
            ObjectNode root;
            if (Files.isRegularFile(storePath())) {
                JsonNode existing = JSON.readTree(storePath().toFile());
                root =
                        existing != null && existing.isObject()
                                ? ((ObjectNode) existing).deepCopy()
                                : JSON.createObjectNode();
            } else {
                JsonNode defaults = readMergedSessionUiDefaultsNode();
                root =
                        defaults != null && defaults.isObject()
                                ? ((ObjectNode) defaults).deepCopy()
                                : JSON.createObjectNode();
            }
            if (uiThemeId != null && !uiThemeId.isBlank()) {
                root.put("uiTheme", uiThemeId.strip());
            }
            if (rows != null && !rows.isEmpty()) {
                ArrayNode array = JSON.createArrayNode();
                for (UiEnvRowSnapshot snap : rows) {
                    if (snap == null) {
                        continue;
                    }
                    String name = snap.name() != null ? snap.name().trim() : "";
                    if (name.isEmpty()) {
                        continue;
                    }
                    ObjectNode row = JSON.createObjectNode();
                    row.put("name", name);
                    row.put("value", snap.value() != null ? snap.value() : "");
                    row.put(
                            "description",
                            snap.description() != null ? snap.description() : "");
                    array.add(row);
                }
                root.set("uiEnvRows", array);
            }
            Files.createDirectories(storePath().getParent());
            JSON.writerWithDefaultPrettyPrinter().writeValue(storePath().toFile(), root);
        } catch (IOException ignored) {
        }
    }

    private static String text(JsonNode root, String key) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull() || !n.isTextual()) {
            return "";
        }
        return n.asText("");
    }

    private static double optionalDouble(JsonNode root, String key, double defaultValue) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull() || !n.isNumber()) {
            return defaultValue;
        }
        return n.asDouble();
    }

    private static double clamp01(double v) {
        if (!Double.isFinite(v)) {
            return 1d;
        }
        return Math.max(0d, Math.min(1d, v));
    }

    private static long optionalLongClamped(
            JsonNode root, String key, long defaultValue, long min, long max) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull() || !n.isNumber()) {
            return defaultValue;
        }
        long v = n.asLong();
        return Math.max(min, Math.min(max, v));
    }

    private static int optionalIntClamped(
            JsonNode root, String key, int defaultValue, int min, int max) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull() || !n.isNumber()) {
            return defaultValue;
        }
        int v = n.asInt();
        return Math.max(min, Math.min(max, v));
    }

    private static long optionalNonNegativeLong(JsonNode root, String key, long defaultValue) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull() || !n.isNumber()) {
            return defaultValue;
        }
        long v = n.asLong();
        if (v < 0) {
            return defaultValue;
        }
        return Math.min(v, 65536L);
    }

    private static void putMemorySettingsPrefs(ObjectNode root, DesktopSessionState state) {
        root.put("memoryMonitorEnabled", state.memoryMonitorEnabled());
        long iv = state.memoryMonitorIntervalSec();
        root.put("memoryMonitorIntervalSec", Math.max(1L, Math.min(3600L, iv <= 0 ? 5L : iv)));
        root.put("nextLaunchHeapFixed", state.nextLaunchHeapFixed());
        long nMin = state.nextLaunchHeapMinMiB();
        if (nMin > 0) {
            root.put("nextLaunchHeapMinMiB", nMin);
        }
        long nx = state.nextLaunchHeapMaxMiB();
        if (nx > 0) {
            root.put("nextLaunchHeapMaxMiB", nx);
        }
    }

    /**
     * 旧セッション（{@code nextLaunchHeapFixed} 無し）: 上限が保存されていれば固定（-Xms=-Xmx）、未設定なら変動既定。
     */
    private static boolean loadNextLaunchHeapFixed(JsonNode root) {
        JsonNode n = root.get("nextLaunchHeapFixed");
        if (n != null && !n.isNull() && n.isBoolean()) {
            return n.asBoolean();
        }
        long savedMax = optionalNonNegativeLong(root, "nextLaunchHeapMaxMiB", 0L);
        return savedMax > 0;
    }

    private static long loadNextLaunchHeapMinMiB(JsonNode root) {
        JsonNode n = root.get("nextLaunchHeapMinMiB");
        if (n != null && !n.isNull() && n.isNumber()) {
            long v = n.asLong();
            if (v > 0) {
                return Math.min(v, 65536L);
            }
        }
        long savedMax = optionalNonNegativeLong(root, "nextLaunchHeapMaxMiB", 0L);
        if (savedMax > 0) {
            return savedMax;
        }
        return DesktopSessionState.DEFAULT_NEXT_LAUNCH_HEAP_MIN_MIB;
    }

    private static long loadNextLaunchHeapMaxMiB(JsonNode root) {
        long saved = optionalNonNegativeLong(root, "nextLaunchHeapMaxMiB", 0L);
        if (saved > 0) {
            return saved;
        }
        return DesktopSessionState.DEFAULT_NEXT_LAUNCH_HEAP_MAX_MIB;
    }

    private static void putEquipmentStatusDashboardPrefs(ObjectNode root, DesktopSessionState state) {
        put(root, "equipmentStatusDashboardActualDate", state.equipmentStatusDashboardActualDate());
        put(root, "equipmentStatusDashboardPlanDate", state.equipmentStatusDashboardPlanDate());
        root.put("equipmentStatusDashboardAutoRefreshEnabled", state.equipmentStatusDashboardAutoRefreshEnabled());
        root.put(
                "equipmentStatusDashboardAutoRefreshIntervalSec",
                Math.max(
                        DesktopSessionState.MIN_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                        Math.min(
                                DesktopSessionState.MAX_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC,
                                state.equipmentStatusDashboardAutoRefreshIntervalSec() <= 0
                                        ? DesktopSessionState
                                                .DEFAULT_EQUIPMENT_STATUS_DASHBOARD_AUTO_REFRESH_INTERVAL_SEC
                                        : state.equipmentStatusDashboardAutoRefreshIntervalSec())));
        root.put("equipmentStatusDashboardShowAladdinPlans", state.equipmentStatusDashboardShowAladdinPlans());
        root.put("equipmentStatusDashboardShowDispatchPlans", state.equipmentStatusDashboardShowDispatchPlans());
        putEquipmentStatusDashboardAppearancePrefs(root, state.equipmentStatusDashboardAppearance());
    }

    private static EquipmentStatusDashboardAppearancePrefs loadEquipmentStatusDashboardAppearancePrefs(
            JsonNode root) {
        EquipmentStatusDashboardAppearancePrefs d = EquipmentStatusDashboardAppearancePrefs.defaults();
        JsonNode n = root.get("equipmentStatusDashboardAppearance");
        if (n == null || !n.isObject()) {
            return d;
        }
        return new EquipmentStatusDashboardAppearancePrefs(
                optionalIntClamped(n, "columnCount", d.columnCount(), 0, 12),
                optionalDouble(n, "cardWidth", d.cardWidth()),
                optionalDouble(n, "fullscreenCardWidthPercent", d.fullscreenCardWidthPercent()),
                optionalDouble(n, "cardPadding", d.cardPadding()),
                optionalDouble(n, "cardGapH", d.cardGapH()),
                optionalDouble(n, "cardGapV", d.cardGapV()),
                optionalDouble(n, "cardBorderRadius", d.cardBorderRadius()),
                nzFallback(text(n, "cardShadowStyle"), d.cardShadowStyle()),
                text(n, "fontFamily"),
                optionalDouble(n, "machineFontPx", d.machineFontPx()),
                optionalDouble(n, "metaFontPx", d.metaFontPx()),
                optionalDouble(n, "planFontPx", d.planFontPx()),
                optionalDouble(n, "pctFontPx", d.pctFontPx()),
                optionalDouble(n, "chartSizePx", d.chartSizePx()),
                nzFallback(text(n, "chartDoneColorHex"), d.chartDoneColorHex()),
                nzFallback(text(n, "chartRemainColorHex"), d.chartRemainColorHex()),
                nzFallback(text(n, "chartStyle"), d.chartStyle()),
                optionalBoolean(n, "chartShadowEnabled", d.chartShadowEnabled()),
                nzFallback(text(n, "fullscreenTheme"), d.fullscreenTheme()));
    }

    private static void putEquipmentStatusDashboardAppearancePrefs(
            ObjectNode root, EquipmentStatusDashboardAppearancePrefs prefs) {
        EquipmentStatusDashboardAppearancePrefs p =
                prefs != null ? prefs : EquipmentStatusDashboardAppearancePrefs.defaults();
        if (p.equals(EquipmentStatusDashboardAppearancePrefs.defaults())) {
            return;
        }
        ObjectNode o = root.putObject("equipmentStatusDashboardAppearance");
        o.put("columnCount", p.columnCount());
        o.put("cardWidth", p.cardWidth());
        o.put("fullscreenCardWidthPercent", p.fullscreenCardWidthPercent());
        o.put("cardPadding", p.cardPadding());
        o.put("cardGapH", p.cardGapH());
        o.put("cardGapV", p.cardGapV());
        o.put("cardBorderRadius", p.cardBorderRadius());
        o.put("cardShadowStyle", p.cardShadowStyle());
        put(o, "fontFamily", p.fontFamily());
        o.put("machineFontPx", p.machineFontPx());
        o.put("metaFontPx", p.metaFontPx());
        o.put("planFontPx", p.planFontPx());
        o.put("pctFontPx", p.pctFontPx());
        o.put("chartSizePx", p.chartSizePx());
        o.put("chartDoneColorHex", p.chartDoneColorHex());
        o.put("chartRemainColorHex", p.chartRemainColorHex());
        o.put("chartStyle", p.chartStyle());
        o.put("chartShadowEnabled", p.chartShadowEnabled());
        o.put("fullscreenTheme", p.fullscreenTheme());
    }

    private static boolean optionalBoolean(JsonNode root, String key, boolean defaultValue) {
        JsonNode n = root.get(key);
        if (n == null || n.isNull()) {
            return defaultValue;
        }
        if (n.isBoolean()) {
            return n.booleanValue();
        }
        if (n.isTextual()) {
            String t = n.asText("").trim().toLowerCase(Locale.ROOT);
            return !List.of("0", "false", "no", "off", "none").contains(t);
        }
        return defaultValue;
    }

    private static void put(ObjectNode root, String key, String value) {
        if (value != null && !value.isBlank()) {
            root.put(key, value.trim());
        }
    }

    private static void putLogFontSize(ObjectNode root, double sizePoints) {
        if (Double.isFinite(sizePoints) && sizePoints > 0) {
            root.put("logFontSize", sizePoints);
        }
    }

    private static List<String> loadStringList(JsonNode root, String key) {
        JsonNode arr = root.get(key);
        if (arr == null || !arr.isArray()) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (JsonNode el : arr) {
            if (el != null && el.isTextual()) {
                out.add(el.asText(""));
            } else if (el != null && el.isValueNode()) {
                out.add(el.asText(""));
            }
        }
        return List.copyOf(out);
    }

    private static void putMainRunLogLines(ObjectNode root, List<String> lines) {
        if (lines == null || lines.isEmpty()) {
            return;
        }
        ArrayNode arr = JSON.createArrayNode();
        for (String s : lines) {
            arr.add(s != null ? s : "");
        }
        root.set("mainRunLogLines", arr);
    }

    private static void putMainRunLogScroll(ObjectNode root, double scroll) {
        if (Double.isFinite(scroll)) {
            root.put("mainRunLogScroll", scroll);
        }
    }

    private static List<UiEnvRowSnapshot> loadUiEnvRows(JsonNode root) {
        JsonNode arr = root.get("uiEnvRows");
        if (arr == null || !arr.isArray()) {
            return List.of();
        }
        List<UiEnvRowSnapshot> out = new ArrayList<>();
        for (JsonNode el : arr) {
            if (el != null && el.isObject()) {
                out.add(
                        new UiEnvRowSnapshot(
                                text(el, "name"), text(el, "value"), text(el, "description")));
            }
        }
        return List.copyOf(out);
    }

    private static void putUiEnvRows(ObjectNode root, List<UiEnvRowSnapshot> rows) {
        ArrayNode arr = JSON.createArrayNode();
        if (rows != null) {
            for (UiEnvRowSnapshot r : rows) {
                ObjectNode o = arr.addObject();
                o.put("name", r.name() != null ? r.name() : "");
                o.put("value", r.value() != null ? r.value() : "");
                o.put("description", r.description() != null ? r.description() : "");
            }
        }
        root.set("uiEnvRows", arr);
    }

    private static void putMainShellTabOrder(ObjectNode root, List<String> order) {
        if (order == null || order.isEmpty()) {
            return;
        }
        ArrayNode arr = JSON.createArrayNode();
        for (String s : order) {
            if (s != null && !s.isBlank()) {
                arr.add(s.trim());
            }
        }
        if (!arr.isEmpty()) {
            root.set("mainShellTabOrder", arr);
        }
    }

    private static List<MainShellTabLayoutNode> loadMainShellTabLayout(JsonNode root) {
        JsonNode arr = root.get("mainShellTabLayout");
        if (arr == null || !arr.isArray()) {
            return List.of();
        }
        List<MainShellTabLayoutNode> out = new ArrayList<>();
        for (JsonNode el : arr) {
            MainShellTabLayoutNode n = MainShellTabLayoutNode.fromJson(el);
            if (n != null) {
                out.add(n);
            }
        }
        return List.copyOf(out);
    }

    private static void putMainShellTabLayout(ObjectNode root, List<MainShellTabLayoutNode> layout) {
        if (layout == null || layout.isEmpty()) {
            return;
        }
        ArrayNode arr = JsonNodeFactory.instance.arrayNode();
        for (MainShellTabLayoutNode n : layout) {
            arr.add(layoutNodeToJson(n));
        }
        root.set("mainShellTabLayout", arr);
    }

    private static Map<String, String> loadStringStringMap(JsonNode root, String key) {
        JsonNode o = root.get(key);
        if (o == null || !o.isObject()) {
            return Map.of();
        }
        Map<String, String> out = new LinkedHashMap<>();
        Iterator<String> names = o.fieldNames();
        while (names.hasNext()) {
            String k = names.next();
            if (k == null || k.isBlank()) {
                continue;
            }
            JsonNode vn = o.get(k);
            if (vn != null && vn.isTextual()) {
                String v = vn.asText("");
                if (!v.isBlank()) {
                    out.put(k.trim(), v.strip());
                }
            }
        }
        return Map.copyOf(out);
    }

    private static void putStringStringMap(ObjectNode root, String key, Map<String, String> map) {
        if (map == null || map.isEmpty()) {
            return;
        }
        ObjectNode o = JSON.createObjectNode();
        for (Map.Entry<String, String> e : map.entrySet()) {
            if (e.getKey() != null
                    && !e.getKey().isBlank()
                    && e.getValue() != null
                    && !e.getValue().isBlank()) {
                o.put(e.getKey().trim(), e.getValue().strip());
            }
        }
        if (o.size() > 0) {
            root.set(key, o);
        }
    }

    private static Map<String, Integer> loadStringIntMap(JsonNode root, String key) {
        JsonNode o = root.get(key);
        if (o == null || !o.isObject()) {
            return Map.of();
        }
        Map<String, Integer> out = new LinkedHashMap<>();
        Iterator<String> names = o.fieldNames();
        while (names.hasNext()) {
            String k = names.next();
            if (k == null || k.isBlank()) {
                continue;
            }
            JsonNode vn = o.get(k);
            if (vn != null && vn.isNumber()) {
                int idx = vn.asInt(-1);
                if (idx >= 0) {
                    out.put(k.trim(), idx);
                }
            }
        }
        return Map.copyOf(out);
    }

    private static void putStringIntMap(ObjectNode root, String key, Map<String, Integer> map) {
        if (map == null || map.isEmpty()) {
            return;
        }
        ObjectNode o = JSON.createObjectNode();
        for (Map.Entry<String, Integer> e : map.entrySet()) {
            if (e.getKey() != null
                    && !e.getKey().isBlank()
                    && e.getValue() != null
                    && e.getValue() >= 0) {
                o.put(e.getKey().trim(), e.getValue());
            }
        }
        if (o.size() > 0) {
            root.set(key, o);
        }
    }

    private static JsonNode layoutNodeToJson(MainShellTabLayoutNode n) {
        ObjectNode o = JsonNodeFactory.instance.objectNode();
        o.put("kind", n.kind());
        if (!n.colorHex().isBlank()) {
            o.put("color", n.colorHex());
        }
        if (n.isTab()) {
            o.put("id", n.id());
            return o;
        }
        o.put("title", n.title().isBlank() ? "グループ" : n.title());
        ArrayNode ch = JsonNodeFactory.instance.arrayNode();
        for (MainShellTabLayoutNode c : n.children()) {
            ch.add(layoutNodeToJson(c));
        }
        o.set("children", ch);
        return o;
    }

    private static double loadEquipmentGanttPersonBadgeGapPx(JsonNode root) {
        double gap = optionalDouble(root, "equipmentGanttPersonBadgeGapPx", -1d);
        if (Double.isFinite(gap) && gap >= 0) {
            return Math.clamp(
                    gap, 0.0, DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX);
        }
        double legacy = optionalDouble(root, "equipmentGanttPersonBadgeOverlapPercent", -1d);
        if (Double.isFinite(legacy) && legacy >= 0) {
            return legacyPersonBadgeOverlapPercentToGapPx(legacy);
        }
        return DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX;
    }

    private static double loadEquipmentGanttPersonBadgeBandVerticalOffsetPx(JsonNode root) {
        double v =
                optionalDouble(
                        root,
                        "equipmentGanttPersonBadgeBandVerticalOffsetPx",
                        DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX);
        if (!Double.isFinite(v)) {
            return DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX;
        }
        return Math.clamp(
                v,
                DesktopSessionState.MIN_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX,
                DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX);
    }

    private static double loadEquipmentGanttPersonBadgeWireMaxLengthPx(JsonNode root) {
        double v =
                optionalDouble(
                        root,
                        "equipmentGanttPersonBadgeWireMaxLengthPx",
                        DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX);
        if (!Double.isFinite(v) || v < 0) {
            return DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX;
        }
        return Math.min(v, DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX);
    }

    /**
     * 旧「重なり量 0〜80%」を、おおよそ同程度の疎さになるよう固定間隔（px）に変換する移行用。
     */
    private static double legacyPersonBadgeOverlapPercentToGapPx(double overlapPercent0To80) {
        double op = Math.clamp(overlapPercent0To80, 0.0, 80.0);
        return Math.clamp(
                (80.0 - op) / 80.0 * 12.0,
                0.0,
                DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX);
    }

    private static Map<String, EquipmentGanttBadgeDragDelta> loadEquipmentGanttBadgeDragDeltas(
            JsonNode root) {
        JsonNode obj = root.get("equipmentGanttBadgeDragDeltas");
        if (obj == null || !obj.isObject()) {
            return Map.of();
        }
        Map<String, EquipmentGanttBadgeDragDelta> out = new LinkedHashMap<>();
        for (Iterator<String> it = obj.fieldNames(); it.hasNext(); ) {
            String field = it.next();
            JsonNode el = obj.get(field);
            if (el != null && el.isObject()) {
                double dx = el.path("dx").asDouble(0.0);
                double dy = el.path("dy").asDouble(0.0);
                out.put(field, new EquipmentGanttBadgeDragDelta(dx, dy));
            }
        }
        return Map.copyOf(out);
    }

    private static void putEquipmentGanttBadgeDragDeltas(
            ObjectNode root, Map<String, EquipmentGanttBadgeDragDelta> deltas) {
        if (deltas == null || deltas.isEmpty()) {
            return;
        }
        ObjectNode o = JSON.createObjectNode();
        for (Map.Entry<String, EquipmentGanttBadgeDragDelta> e : deltas.entrySet()) {
            if (e.getKey() == null || e.getKey().isBlank() || e.getValue() == null) {
                continue;
            }
            ObjectNode one = JSON.createObjectNode();
            one.put("dx", e.getValue().dx());
            one.put("dy", e.getValue().dy());
            o.set(e.getKey(), one);
        }
        if (!o.isEmpty()) {
            root.set("equipmentGanttBadgeDragDeltas", o);
        }
    }

    private static void putEquipmentGanttGraphicPrefs(ObjectNode root, DesktopSessionState state) {
        double z = state.equipmentGanttGraphicZoomPercent();
        if (Double.isFinite(z) && z >= 50 && z <= 200) {
            root.put("equipmentGanttGraphicZoomPercent", z);
        }
        double dw = state.equipmentGanttDateColWidth();
        if (Double.isFinite(dw) && dw > 0) {
            root.put("equipmentGanttDateColWidth", dw);
        }
        double mw = state.equipmentGanttMachineColWidth();
        if (Double.isFinite(mw) && mw > 0) {
            root.put("equipmentGanttMachineColWidth", mw);
        }
        double pw = state.equipmentGanttProcessColWidth();
        if (Double.isFinite(pw) && pw > 0) {
            root.put("equipmentGanttProcessColWidth", pw);
        }
        String bf = state.equipmentGanttBarFontFamily();
        if (bf != null && !bf.isBlank()) {
            root.put("equipmentGanttBarFontFamily", bf.strip());
        }
        double bfp = state.equipmentGanttBarFontPercent();
        if (Double.isFinite(bfp) && bfp >= 50 && bfp <= 200) {
            root.put("equipmentGanttBarFontPercent", bfp);
        }
        double rh = state.equipmentGanttRowHeightPercent();
        if (Double.isFinite(rh) && rh >= 50 && rh <= 200) {
            root.put("equipmentGanttRowHeightPercent", rh);
        }
        double hh = state.equipmentGanttHeaderHeightPercent();
        if (Double.isFinite(hh) && hh >= 50 && hh <= 200) {
            root.put("equipmentGanttHeaderHeightPercent", hh);
        }
        double sw = state.equipmentGanttSlotWidthPercent();
        if (Double.isFinite(sw) && sw >= 50 && sw <= 500) {
            root.put("equipmentGanttSlotWidthPercent", sw);
        }
        double sh = state.equipmentGanttShiftWheelHScrollPercent();
        if (Double.isFinite(sh) && sh >= 50 && sh <= 1000) {
            root.put("equipmentGanttShiftWheelHScrollPercent", sh);
        }
        root.put(
                "equipmentGanttPrepTimeLabelsEnabled",
                state.equipmentGanttPrepTimeLabelsEnabled());
        double bgap = state.equipmentGanttPersonBadgeGapPx();
        if (Double.isFinite(bgap) && bgap >= 0) {
            root.put("equipmentGanttPersonBadgeGapPx", bgap);
        }
        double bvOff = state.equipmentGanttPersonBadgeBandVerticalOffsetPx();
        if (Double.isFinite(bvOff)) {
            root.put(
                    "equipmentGanttPersonBadgeBandVerticalOffsetPx",
                    Math.clamp(
                            bvOff,
                            DesktopSessionState.MIN_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX,
                            DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX));
        }
        put(root, "equipmentGanttGraphicDataFingerprint", state.equipmentGanttGraphicDataFingerprint());
        putEquipmentGanttBadgeDragDeltas(root, state.equipmentGanttBadgeDragDeltas());
        root.put(
                "equipmentGanttPersonBadgeDragAdjustEnabled",
                state.equipmentGanttPersonBadgeDragAdjustEnabled());
        root.put("equipmentGanttPersonBadgeEnabled", state.equipmentGanttPersonBadgeEnabled());
        root.put(
                "equipmentGanttPersonBadgeWireEnabled",
                state.equipmentGanttPersonBadgeWireEnabled());
        put(root, "equipmentGanttPersonBadgeWireStrokeHex", state.equipmentGanttPersonBadgeWireStrokeHex());
        double wWire = state.equipmentGanttPersonBadgeWireWidthPx();
        if (Double.isFinite(wWire)) {
            root.put(
                    "equipmentGanttPersonBadgeWireWidthPx",
                    Math.min(
                            Math.max(0d, wWire),
                            DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_WIDTH_PX));
        }
        String dashKey = state.equipmentGanttPersonBadgeWireDashStyleKey();
        root.put(
                "equipmentGanttPersonBadgeWireDashStyleKey",
                dashKey != null && !dashKey.isBlank()
                        ? dashKey.strip()
                        : DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_DASH_STYLE_KEY);
        double wMaxLen = state.equipmentGanttPersonBadgeWireMaxLengthPx();
        if (Double.isFinite(wMaxLen) && wMaxLen >= 0) {
            root.put(
                    "equipmentGanttPersonBadgeWireMaxLengthPx",
                    Math.min(
                            wMaxLen,
                            DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_MAX_LENGTH_PX));
        }
        put(root, "equipmentGanttPersonBadgeFontFamily", state.equipmentGanttPersonBadgeFontFamily());
        double bpf = state.equipmentGanttPersonBadgeFontPercent();
        if (Double.isFinite(bpf) && bpf > 0 && bpf <= 300) {
            root.put("equipmentGanttPersonBadgeFontPercent", bpf);
        }
        put(root, "equipmentGanttPersonBadgeFillHex", state.equipmentGanttPersonBadgeFillHex());
        put(root, "equipmentGanttPersonBadgeTextHex", state.equipmentGanttPersonBadgeTextHex());
        put(root, "equipmentGanttPersonBadgeStrokeHex", state.equipmentGanttPersonBadgeStrokeHex());
        double stw = state.equipmentGanttPersonBadgeStrokeWidth();
        if (Double.isFinite(stw) && stw >= 0) {
            root.put("equipmentGanttPersonBadgeStrokeWidth", stw);
        }
        double cr = state.equipmentGanttPersonBadgeCornerRadius();
        if (Double.isFinite(cr) && cr >= 0) {
            root.put("equipmentGanttPersonBadgeCornerRadius", cr);
        }
        root.put("equipmentGanttPersonBadgePill", state.equipmentGanttPersonBadgePill());
        put(root, "equipmentGanttPersonBadgeGlowColorHex", state.equipmentGanttPersonBadgeGlowColorHex());
        double gr = state.equipmentGanttPersonBadgeGlowRadius();
        if (Double.isFinite(gr) && gr >= 0) {
            root.put("equipmentGanttPersonBadgeGlowRadius", gr);
        }
        double gs = state.equipmentGanttPersonBadgeGlowSpread();
        if (Double.isFinite(gs) && gs >= 0 && gs <= 1) {
            root.put("equipmentGanttPersonBadgeGlowSpread", gs);
        }
        double bop = state.equipmentGanttPersonBadgeOpacity();
        if (Double.isFinite(bop) && bop >= 0 && bop <= 1) {
            root.put("equipmentGanttPersonBadgeOpacity", bop);
        }
        putPersonBadgeStyleMap(root, state.equipmentGanttPersonBadgeStylesByLabel(), "equipmentGanttPersonBadgeStylesByLabel");
        putPersonBadgeStyleMap(root, state.equipmentGanttPersonBadgeStylesByMemberKey(), "equipmentGanttPersonBadgeStylesByMemberKey");
        put(root, "equipmentGanttPlanJsonPath", state.equipmentGanttPlanJsonPath());
    }

    private static Map<String, PersonBadgeStyle> loadPersonBadgeStyleMap(JsonNode root, String jsonKey) {
        JsonNode obj = root.get(jsonKey);
        if (obj == null || !obj.isObject()) {
            return Map.of();
        }
        Map<String, PersonBadgeStyle> out = new LinkedHashMap<>();
        for (Iterator<String> it = obj.fieldNames(); it.hasNext(); ) {
            String field = it.next();
            JsonNode el = obj.get(field);
            if (el != null && el.isObject()) {
                PersonBadgeStyle st = loadPersonBadgeStyleObject(el);
                String k = PersonBadgeStyle.normalizeLabelKey(field);
                if (st != null && !k.isEmpty()) {
                    out.put(k, st);
                }
            }
        }
        return Map.copyOf(out);
    }

    private static PersonBadgeStyle loadStage1NetworkCacheBadgeStyle(JsonNode root) {
        JsonNode n = root.get("stage1NetworkCacheBadgeStyle");
        if (n == null || !n.isObject()) {
            return PersonBadgeStyle.networkSourceCacheBadgeDefault();
        }
        PersonBadgeStyle st = loadPersonBadgeStyleObject(n);
        return st != null ? st : PersonBadgeStyle.networkSourceCacheBadgeDefault();
    }

    private static PersonBadgeStyle loadRequestFormPreviewUpdateBadgeStyle(JsonNode root) {
        JsonNode n = root.get("requestFormPreviewUpdateBadgeStyle");
        if (n == null || !n.isObject()) {
            return PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();
        }
        PersonBadgeStyle st = loadPersonBadgeStyleObject(n);
        return st != null ? st : PersonBadgeStyle.requestFormPreviewUpdateBadgeDefault();
    }

    private static void putStage1NetworkCacheBadgePrefs(ObjectNode root, DesktopSessionState state) {
        put(root, "stage1NetworkCacheBadgeLabel", state.stage1NetworkCacheBadgeLabel());
        PersonBadgeStyle st = state.stage1NetworkCacheBadgeStyle();
        if (st == null) {
            return;
        }
        ObjectNode o = root.putObject("stage1NetworkCacheBadgeStyle");
        o.put("fontFamily", st.fontFamily() != null ? st.fontFamily() : "");
        o.put("fontPercent", st.fontPercent());
        o.put("fillHex", st.fillHex());
        o.put("textHex", st.textHex());
        o.put("strokeHex", st.strokeHex());
        o.put("strokeWidth", st.strokeWidth());
        o.put("cornerRadius", st.cornerRadius());
        o.put("pill", st.pill());
        o.put("glowColorHex", st.glowColorHex());
        o.put("glowRadius", st.glowRadius());
        o.put("glowSpread", st.glowSpread());
        o.put("opacity", st.opacity());
    }

    private static void putRequestFormPreviewUpdateBadgePrefs(ObjectNode root, DesktopSessionState state) {
        put(root, "requestFormPreviewUpdateBadgeLabel", state.requestFormPreviewUpdateBadgeLabel());
        PersonBadgeStyle st = state.requestFormPreviewUpdateBadgeStyle();
        if (st == null) {
            return;
        }
        ObjectNode o = root.putObject("requestFormPreviewUpdateBadgeStyle");
        o.put("fontFamily", st.fontFamily() != null ? st.fontFamily() : "");
        o.put("fontPercent", st.fontPercent());
        o.put("fillHex", st.fillHex());
        o.put("textHex", st.textHex());
        o.put("strokeHex", st.strokeHex());
        o.put("strokeWidth", st.strokeWidth());
        o.put("cornerRadius", st.cornerRadius());
        o.put("pill", st.pill());
        o.put("glowColorHex", st.glowColorHex());
        o.put("glowRadius", st.glowRadius());
        o.put("glowSpread", st.glowSpread());
        o.put("opacity", st.opacity());
    }

    private static PersonBadgeStyle loadPersonBadgeStyleObject(JsonNode o) {
        PersonBadgeStyle d = PersonBadgeStyle.defaultStyle();
        String ff = text(o, "fontFamily");
        double fp = optionalDouble(o, "fontPercent", 0d);
        String fill = text(o, "fillHex");
        String tx = text(o, "textHex");
        String stk = text(o, "strokeHex");
        double stw = optionalDouble(o, "strokeWidth", -1d);
        double cr = optionalDouble(o, "cornerRadius", -1d);
        boolean pill = optionalBoolean(o, "pill", false);
        String glow = text(o, "glowColorHex");
        double gr = optionalDouble(o, "glowRadius", -1d);
        double gs = optionalDouble(o, "glowSpread", -1d);
        double op = optionalDouble(o, "opacity", -1d);
        return new PersonBadgeStyle(
                ff,
                fp > 0 && fp <= 300 ? fp : d.fontPercent(),
                nzStored(fill, d.fillHex()),
                nzStored(tx, d.textHex()),
                nzStored(stk, d.strokeHex()),
                stw >= 0 ? stw : d.strokeWidth(),
                cr >= 0 ? cr : d.cornerRadius(),
                pill,
                nzStored(glow, d.glowColorHex()),
                gr >= 0 ? gr : d.glowRadius(),
                gs >= 0 && gs <= 1 ? gs : d.glowSpread(),
                op >= 0 && op <= 1 ? op : d.opacity());
    }

    private static String nzStored(String s, String def) {
        return s != null && !s.isBlank() ? s.strip() : def;
    }

    private static void putPersonBadgeStyleMap(
            ObjectNode root, Map<String, PersonBadgeStyle> styleMap, String jsonKey) {
        Map<String, PersonBadgeStyle> m = styleMap;
        if (m == null || m.isEmpty()) {
            return;
        }
        ObjectNode bag = JSON.createObjectNode();
        for (Map.Entry<String, PersonBadgeStyle> e : m.entrySet()) {
            if (e.getKey() == null || e.getKey().isBlank() || e.getValue() == null) {
                continue;
            }
            String canon = PersonBadgeStyle.normalizeLabelKey(e.getKey());
            if (canon.isEmpty()) {
                continue;
            }
            PersonBadgeStyle st = e.getValue();
            ObjectNode o = bag.putObject(canon);
            o.put("fontFamily", st.fontFamily() != null ? st.fontFamily() : "");
            o.put("fontPercent", st.fontPercent());
            o.put("fillHex", st.fillHex());
            o.put("textHex", st.textHex());
            o.put("strokeHex", st.strokeHex());
            o.put("strokeWidth", st.strokeWidth());
            o.put("cornerRadius", st.cornerRadius());
            o.put("pill", st.pill());
            o.put("glowColorHex", st.glowColorHex());
            o.put("glowRadius", st.glowRadius());
            o.put("glowSpread", st.glowSpread());
            o.put("opacity", st.opacity());
        }
        if (!bag.isEmpty()) {
            root.set(jsonKey, bag);
        }
    }

    private static PushButtonDesignPrefs loadPushButtonDesignPrefs(JsonNode root) {
        JsonNode n = root.get("pushButtonDesignPrefs");
        PushButtonDesignPrefs d = PushButtonDesignPrefs.inactiveDefaults();
        if (n == null || !n.isObject()) {
            return d;
        }
        boolean cg = optionalBoolean(n, "customizeGeneralRunTab", false);
        boolean cs = optionalBoolean(n, "customizeStageRunButtons", false);
        double gr = optionalDouble(n, "generalBorderRadius", d.generalBorderRadius());
        double gpv = optionalDouble(n, "generalPaddingV", d.generalPaddingV());
        double gph = optionalDouble(n, "generalPaddingH", d.generalPaddingH());
        double gf = optionalDouble(n, "generalFontPx", d.generalFontPx());
        String gbg = nzFallback(text(n, "generalBgHex"), d.generalBgHex());
        String gb = nzFallback(text(n, "generalBorderHex"), d.generalBorderHex());
        String gt = nzFallback(text(n, "generalTextHex"), d.generalTextHex());
        String gh = nzFallback(text(n, "generalHoverBgHex"), d.generalHoverBgHex());
        String gp = nzFallback(text(n, "generalPressedBgHex"), d.generalPressedBgHex());
        double sf = optionalDouble(n, "stageFontPx", d.stageFontPx());
        double smw = optionalDouble(n, "stageMinWidth", d.stageMinWidth());
        double smh = optionalDouble(n, "stageMinHeight", d.stageMinHeight());
        double spv = optionalDouble(n, "stagePaddingV", d.stagePaddingV());
        double sph = optionalDouble(n, "stagePaddingH", d.stagePaddingH());
        double sr = optionalDouble(n, "stageBorderRadius", d.stageBorderRadius());
        String s1b = nzFallback(text(n, "stage1BgHex"), d.stage1BgHex());
        String s1bo = nzFallback(text(n, "stage1BorderHex"), d.stage1BorderHex());
        String s1h = nzFallback(text(n, "stage1HoverBgHex"), d.stage1HoverBgHex());
        String s1p = nzFallback(text(n, "stage1PressedBgHex"), d.stage1PressedBgHex());
        String s2b = nzFallback(text(n, "stage2BgHex"), d.stage2BgHex());
        String s2bo = nzFallback(text(n, "stage2BorderHex"), d.stage2BorderHex());
        String s2h = nzFallback(text(n, "stage2HoverBgHex"), d.stage2HoverBgHex());
        String s2p = nzFallback(text(n, "stage2PressedBgHex"), d.stage2PressedBgHex());
        String s3b = nzFallback(text(n, "stage3BgHex"), d.stage3BgHex());
        String s3bo = nzFallback(text(n, "stage3BorderHex"), d.stage3BorderHex());
        String s3h = nzFallback(text(n, "stage3HoverBgHex"), d.stage3HoverBgHex());
        String s3p = nzFallback(text(n, "stage3PressedBgHex"), d.stage3PressedBgHex());
        boolean ca = optionalBoolean(n, "customizeAlignAladdinPlanButton", false);
        double aar = optionalDouble(n, "alignAladdinBorderRadius", d.alignAladdinBorderRadius());
        double aapv = optionalDouble(n, "alignAladdinPaddingV", d.alignAladdinPaddingV());
        double aaph = optionalDouble(n, "alignAladdinPaddingH", d.alignAladdinPaddingH());
        double aaf = optionalDouble(n, "alignAladdinFontPx", d.alignAladdinFontPx());
        String aabg = nzFallback(text(n, "alignAladdinBgHex"), d.alignAladdinBgHex());
        String aabo = nzFallback(text(n, "alignAladdinBorderHex"), d.alignAladdinBorderHex());
        String aat = nzFallback(text(n, "alignAladdinTextHex"), d.alignAladdinTextHex());
        String aah = nzFallback(text(n, "alignAladdinHoverBgHex"), d.alignAladdinHoverBgHex());
        String aap = nzFallback(text(n, "alignAladdinPressedBgHex"), d.alignAladdinPressedBgHex());
        boolean cd = optionalBoolean(n, "customizeDialogButtons", false);
        double dpr = optionalDouble(n, "dialogPrimaryBorderRadius", d.dialogPrimaryBorderRadius());
        double dpv = optionalDouble(n, "dialogPrimaryPaddingV", d.dialogPrimaryPaddingV());
        double dph = optionalDouble(n, "dialogPrimaryPaddingH", d.dialogPrimaryPaddingH());
        double df = optionalDouble(n, "dialogPrimaryFontPx", d.dialogPrimaryFontPx());
        String dpbg = nzFallback(text(n, "dialogPrimaryBgHex"), d.dialogPrimaryBgHex());
        String dpbo = nzFallback(text(n, "dialogPrimaryBorderHex"), d.dialogPrimaryBorderHex());
        String dpt = nzFallback(text(n, "dialogPrimaryTextHex"), d.dialogPrimaryTextHex());
        String dphov = nzFallback(text(n, "dialogPrimaryHoverBgHex"), d.dialogPrimaryHoverBgHex());
        String dppr = nzFallback(text(n, "dialogPrimaryPressedBgHex"), d.dialogPrimaryPressedBgHex());
        double dsr = optionalDouble(n, "dialogSecondaryBorderRadius", d.dialogSecondaryBorderRadius());
        double dsv = optionalDouble(n, "dialogSecondaryPaddingV", d.dialogSecondaryPaddingV());
        double dsh = optionalDouble(n, "dialogSecondaryPaddingH", d.dialogSecondaryPaddingH());
        double dsf = optionalDouble(n, "dialogSecondaryFontPx", d.dialogSecondaryFontPx());
        String dsbg = nzFallback(text(n, "dialogSecondaryBgHex"), d.dialogSecondaryBgHex());
        String dsbo = nzFallback(text(n, "dialogSecondaryBorderHex"), d.dialogSecondaryBorderHex());
        String dst = nzFallback(text(n, "dialogSecondaryTextHex"), d.dialogSecondaryTextHex());
        String dshov = nzFallback(text(n, "dialogSecondaryHoverBgHex"), d.dialogSecondaryHoverBgHex());
        String dspr = nzFallback(text(n, "dialogSecondaryPressedBgHex"), d.dialogSecondaryPressedBgHex());
        return new PushButtonDesignPrefs(
                cg,
                gr,
                gpv,
                gph,
                gf,
                gbg,
                gb,
                gt,
                gh,
                gp,
                cs,
                sf,
                smw,
                smh,
                spv,
                sph,
                sr,
                s1b,
                s1bo,
                s1h,
                s1p,
                s2b,
                s2bo,
                s2h,
                s2p,
                s3b,
                s3bo,
                s3h,
                s3p,
                ca,
                aar,
                aapv,
                aaph,
                aaf,
                aabg,
                aabo,
                aat,
                aah,
                aap,
                cd,
                dpr,
                dpv,
                dph,
                df,
                dpbg,
                dpbo,
                dpt,
                dphov,
                dppr,
                dsr,
                dsv,
                dsh,
                dsf,
                dsbg,
                dsbo,
                dst,
                dshov,
                dspr);
    }

    private static String nzFallback(String s, String def) {
        return s != null && !s.isBlank() ? s.strip() : def;
    }

    private static void putPushButtonDesignPrefs(ObjectNode root, DesktopSessionState state) {
        PushButtonDesignPrefs p = state.pushButtonDesignPrefs();
        if (p == null || p.equals(PushButtonDesignPrefs.inactiveDefaults())) {
            return;
        }
        ObjectNode o = root.putObject("pushButtonDesignPrefs");
        o.put("customizeGeneralRunTab", p.customizeGeneralRunTab());
        o.put("customizeStageRunButtons", p.customizeStageRunButtons());
        o.put("generalBorderRadius", p.generalBorderRadius());
        o.put("generalPaddingV", p.generalPaddingV());
        o.put("generalPaddingH", p.generalPaddingH());
        o.put("generalFontPx", p.generalFontPx());
        o.put("generalBgHex", p.generalBgHex());
        o.put("generalBorderHex", p.generalBorderHex());
        o.put("generalTextHex", p.generalTextHex());
        o.put("generalHoverBgHex", p.generalHoverBgHex());
        o.put("generalPressedBgHex", p.generalPressedBgHex());
        o.put("stageFontPx", p.stageFontPx());
        o.put("stageMinWidth", p.stageMinWidth());
        o.put("stageMinHeight", p.stageMinHeight());
        o.put("stagePaddingV", p.stagePaddingV());
        o.put("stagePaddingH", p.stagePaddingH());
        o.put("stageBorderRadius", p.stageBorderRadius());
        o.put("stage1BgHex", p.stage1BgHex());
        o.put("stage1BorderHex", p.stage1BorderHex());
        o.put("stage1HoverBgHex", p.stage1HoverBgHex());
        o.put("stage1PressedBgHex", p.stage1PressedBgHex());
        o.put("stage2BgHex", p.stage2BgHex());
        o.put("stage2BorderHex", p.stage2BorderHex());
        o.put("stage2HoverBgHex", p.stage2HoverBgHex());
        o.put("stage2PressedBgHex", p.stage2PressedBgHex());
        o.put("stage3BgHex", p.stage3BgHex());
        o.put("stage3BorderHex", p.stage3BorderHex());
        o.put("stage3HoverBgHex", p.stage3HoverBgHex());
        o.put("stage3PressedBgHex", p.stage3PressedBgHex());
        o.put("customizeAlignAladdinPlanButton", p.customizeAlignAladdinPlanButton());
        o.put("alignAladdinBorderRadius", p.alignAladdinBorderRadius());
        o.put("alignAladdinPaddingV", p.alignAladdinPaddingV());
        o.put("alignAladdinPaddingH", p.alignAladdinPaddingH());
        o.put("alignAladdinFontPx", p.alignAladdinFontPx());
        o.put("alignAladdinBgHex", p.alignAladdinBgHex());
        o.put("alignAladdinBorderHex", p.alignAladdinBorderHex());
        o.put("alignAladdinTextHex", p.alignAladdinTextHex());
        o.put("alignAladdinHoverBgHex", p.alignAladdinHoverBgHex());
        o.put("alignAladdinPressedBgHex", p.alignAladdinPressedBgHex());
        o.put("customizeDialogButtons", p.customizeDialogButtons());
        o.put("dialogPrimaryBorderRadius", p.dialogPrimaryBorderRadius());
        o.put("dialogPrimaryPaddingV", p.dialogPrimaryPaddingV());
        o.put("dialogPrimaryPaddingH", p.dialogPrimaryPaddingH());
        o.put("dialogPrimaryFontPx", p.dialogPrimaryFontPx());
        o.put("dialogPrimaryBgHex", p.dialogPrimaryBgHex());
        o.put("dialogPrimaryBorderHex", p.dialogPrimaryBorderHex());
        o.put("dialogPrimaryTextHex", p.dialogPrimaryTextHex());
        o.put("dialogPrimaryHoverBgHex", p.dialogPrimaryHoverBgHex());
        o.put("dialogPrimaryPressedBgHex", p.dialogPrimaryPressedBgHex());
        o.put("dialogSecondaryBorderRadius", p.dialogSecondaryBorderRadius());
        o.put("dialogSecondaryPaddingV", p.dialogSecondaryPaddingV());
        o.put("dialogSecondaryPaddingH", p.dialogSecondaryPaddingH());
        o.put("dialogSecondaryFontPx", p.dialogSecondaryFontPx());
        o.put("dialogSecondaryBgHex", p.dialogSecondaryBgHex());
        o.put("dialogSecondaryBorderHex", p.dialogSecondaryBorderHex());
        o.put("dialogSecondaryTextHex", p.dialogSecondaryTextHex());
        o.put("dialogSecondaryHoverBgHex", p.dialogSecondaryHoverBgHex());
        o.put("dialogSecondaryPressedBgHex", p.dialogSecondaryPressedBgHex());
    }

    private static void putWindowGeometry(ObjectNode root, DesktopSessionState state) {
        double w = state.windowWidth();
        double h = state.windowHeight();
        if (Double.isFinite(w) && w > 0) {
            root.put("windowWidth", w);
        }
        if (Double.isFinite(h) && h > 0) {
            root.put("windowHeight", h);
        }
        double x = state.windowX();
        double y = state.windowY();
        if (Double.isFinite(x)) {
            root.put("windowX", x);
        }
        if (Double.isFinite(y)) {
            root.put("windowY", y);
        }
    }
}
