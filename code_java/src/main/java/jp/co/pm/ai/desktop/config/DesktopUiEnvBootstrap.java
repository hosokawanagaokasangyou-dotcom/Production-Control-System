package jp.co.pm.ai.desktop.config;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

import jp.co.pm.ai.desktop.bridge.StagePythonExecutable;

/**
 * 環境変数タブのブートストラップ既定（{@link jp.co.pm.ai.desktop.MainShellController} の
 * {@code bootstrapDefaultValueForKey} と同一ソース）。
 */
public final class DesktopUiEnvBootstrap {

    /** {@link jp.co.pm.ai.desktop.MainShellController} の {@code BOOTSTRAP_ORDER} と同期。 */
    public static final List<String> BOOTSTRAP_ORDER =
            List.of(
                    AppPaths.KEY_PM_AI_PYTHON,
                    AppPaths.KEY_PM_AI_REPO_ROOT,
                    AppPaths.KEY_PM_AI_OUTPUT_DIR,
                    AppPaths.KEY_PM_AI_CODE_PYTHON_DIR,
                    AppPaths.KEY_PM_AI_WORKSPACE,
                    AppPaths.KEY_GEMINI_CREDENTIALS_JSON,
                    AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON,
                    AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                    AppPaths.KEY_PM_AI_COLUMN_CONFIG_WORKBOOK,
                    AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK,
                    AppPaths.KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV,
                    AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET,
                    AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                    AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH,
                    AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
                    AppPaths.KEY_PM_AI_DAILY_REPORT_SOURCE_DIR,
                    AppPaths.KEY_PM_AI_ORDER_DETAIL_SOURCE_DIR,
                    AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR,
                    AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR,
                    AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE,
                    AppPaths.KEY_PM_AI_MACHINE_DELIVERY_MANAGEMENT_XLSM,
                    AppPaths.KEY_PM_AI_REQUEST_FORM_TPI_PDF_DIR,
                    AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR,
                    AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON,
                    AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH,
                    AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR);

    private DesktopUiEnvBootstrap() {}

    /** マップ内の空欄キーへブートストラップ既定を適用する（破壊的）。 */
    public static void fillEmptyBootstrapDefaults(Map<String, String> map) {
        if (map == null) {
            return;
        }
        for (String k : BOOTSTRAP_ORDER) {
            if (map.getOrDefault(k, "").isBlank()) {
                String v = defaultValueForKey(k, map);
                if (!v.isBlank()) {
                    map.put(k, v);
                }
            }
        }
    }

    public static String defaultValueForKey(String k, Map<String, String> ui) {
        Map<String, String> u = ui != null ? ui : Map.of();
        if (k == null || k.isBlank()) {
            return "";
        }
        switch (k) {
            case AppPaths.KEY_PM_AI_PYTHON -> {
                return StagePythonExecutable.defaultPmAiPythonForBootstrap();
            }
            case AppPaths.KEY_PM_AI_REPO_ROOT -> {
                return AppPaths.resolveRepoRoot(u).toString();
            }
            case AppPaths.KEY_PM_AI_CODE_PYTHON_DIR -> {
                return AppPaths.resolvePythonScriptDir(u).toString();
            }
            case AppPaths.KEY_PM_AI_WORKSPACE -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR -> {
                return AppPaths.DEFAULT_PM_AI_TASK_INPUT_SOURCE_DIR;
            }
            case AppPaths.KEY_PM_AI_PROCESSING_PLAN_PATH -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR -> {
                return AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_SOURCE_DIR;
            }
            case AppPaths.KEY_PM_AI_DAILY_REPORT_SOURCE_DIR -> {
                return AppPaths.defaultDailyReportSourceDirForFactory(GlobalInitSettingTarget.load());
            }
            case AppPaths.KEY_PM_AI_ORDER_DETAIL_SOURCE_DIR -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR -> {
                return AppPaths.resolveAladdinMasterDir(u).toString();
            }
            case AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE -> {
                return AppPaths.defaultRequestFormJuchuFileForFactory(GlobalInitSettingTarget.load());
            }
            case AppPaths.KEY_PM_AI_MACHINE_DELIVERY_MANAGEMENT_XLSM -> {
                return AppPaths.defaultMachineDeliveryManagementXlsmForFactory(
                        GlobalInitSettingTarget.load());
            }
            case AppPaths.KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_REQUEST_FORM_TPI_PDF_DIR -> {
                return AppPaths.defaultRequestFormTpiPdfDirForFactory(GlobalInitSettingTarget.load());
            }
            case AppPaths.KEY_PM_AI_RESULT_DISPATCH_TABLE_DIR -> {
                return AppPaths.resolveResultDispatchTableDir(u).toString();
            }
            case AppPaths.KEY_PM_AI_OUTPUT_DIR -> {
                return AppPaths.resolveDefaultOutputDir(u).toString();
            }
            case AppPaths.KEY_GEMINI_CREDENTIALS_JSON -> {
                Path root = AppPaths.resolveRepoRoot(u);
                Path underCode =
                        root.resolve("code").resolve("gemini_credentials.encrypted.json");
                if (Files.isRegularFile(underCode)) {
                    return underCode.toAbsolutePath().normalize().toString();
                }
                Path atRoot = root.resolve("gemini_credentials.encrypted.json");
                return Files.isRegularFile(atRoot)
                        ? atRoot.toAbsolutePath().normalize().toString()
                        : "";
            }
            case AppPaths.KEY_PM_AI_EXCLUDE_RULES_JSON -> {
                return AppPaths.resolveDefaultExcludeRulesJsonPath(u).map(Path::toString).orElse("");
            }
            case AppPaths.KEY_PM_AI_MASTER_WORKBOOK -> {
                return AppPaths.resolveMasterWorkbookCandidate(u).map(Path::toString).orElse("");
            }
            case AppPaths.KEY_PM_AI_COLUMN_CONFIG_WORKBOOK,
                    AppPaths.KEY_PM_AI_DATA_EXTRACTION_SOURCE_WORKBOOK,
                    AppPaths.KEY_PM_AI_RESULT_TASK_COLUMN_CONFIG_CSV -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_SKIP_WORKBOOK_ENV_SHEET -> {
                return "1";
            }
            case AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON,
                    AppPaths.KEY_PM_AI_PLAN_RESULT_TASK_JSON_PATH -> {
                return "";
            }
            case AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR -> {
                return AppPaths.DEFAULT_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR;
            }
            default -> {
                return "";
            }
        }
    }
}
