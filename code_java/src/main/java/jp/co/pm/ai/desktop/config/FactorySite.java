package jp.co.pm.ai.desktop.config;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
import java.util.Optional;

/**
 * 工場別の環境タブ既定（ネットワークソース・バージョンアップ正本 ZIP・マスタ名・サマリ用ブック）。
 *
 * <p>ポータル自動バージョンアップ完了時、および「環境変数を初期値に戻す」で工場を選んだときに適用する。
 * 湖南（{@link #KONAN}）は従来のコード既定、国分（{@link #KOKUBU}）は国分共有パスと {@code code/国分master.xlsm} 絶対パス等を使う。
 */
public enum FactorySite {

    /** 湖南工場（従来のコード既定・工場共有 UNC）。 */
    KONAN(
            "湖南工場",
            AppPaths.DEFAULT_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR,
            AppPaths.DEFAULT_PM_AI_TASK_INPUT_SOURCE_DIR,
            AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
            "",
            AppPaths.SUMMARY_AI_DISPATCH_XLSX),

    /** 国分工場（国分共有・DATA 配下・マスタは {@code 国分master.xlsm}）。 */
    KOKUBU(
            "国分工場",
            AppPaths.DEFAULT_PM_AI_PORTABLE_BUNDLE_RELEASE_DIR_KOKUBU,
            "\\\\192.168.0.101\\共有フォルダ\\国分工場\\国分共有\\●配台AIシステム\\DATA\\計画",
            "\\\\192.168.0.101\\共有フォルダ\\国分工場\\国分共有\\●配台AIシステム\\DATA\\実績",
            "国分master.xlsm",
            AppPaths.KOKUBU_SUMMARY_AI_DISPATCH_WORKBOOK_XLSX);

    private final String displayLabelJa;
    private final String portableBundleSourceDir;
    private final String taskInputSourceDir;
    private final String actualDetailSourceDir;
    /** {@link AppPaths#KEY_MASTER_WORKBOOK_FILE}。空のとき planning_core 側は {@code master.xlsm} 相当。 */
    private final String masterWorkbookFileBasename;
    /** {@code code/} 直下の {@link AppPaths#KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK} 用ファイル名。 */
    private final String summaryAiDispatchWorkbookCodeFilename;

    FactorySite(
            String displayLabelJa,
            String portableBundleSourceDir,
            String taskInputSourceDir,
            String actualDetailSourceDir,
            String masterWorkbookFileBasename,
            String summaryAiDispatchWorkbookCodeFilename) {
        this.displayLabelJa = displayLabelJa;
        this.portableBundleSourceDir = portableBundleSourceDir;
        this.taskInputSourceDir = taskInputSourceDir;
        this.actualDetailSourceDir = actualDetailSourceDir;
        this.masterWorkbookFileBasename = masterWorkbookFileBasename;
        this.summaryAiDispatchWorkbookCodeFilename = summaryAiDispatchWorkbookCodeFilename;
    }

    /** UI 表示用（ダイアログの選択肢文言）。 */
    public String displayLabelJa() {
        return displayLabelJa;
    }

    /** {@link javafx.scene.control.ChoiceDialog} のコンボ表示用。 */
    @Override
    public String toString() {
        return displayLabelJa;
    }

    /** {@link AppPaths#KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR}（ZIP または正本フォルダ）。 */
    public String portableBundleSourceDir() {
        return portableBundleSourceDir;
    }

    /** {@link AppPaths#KEY_PM_AI_TASK_INPUT_SOURCE_DIR}。 */
    public String taskInputSourceDir() {
        return taskInputSourceDir;
    }

    /** {@link AppPaths#KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR}。 */
    public String actualDetailSourceDir() {
        return actualDetailSourceDir;
    }

    /** {@link AppPaths#KEY_MASTER_WORKBOOK_FILE}（basename）。 */
    public String masterWorkbookFileBasename() {
        return masterWorkbookFileBasename;
    }

    /**
     * {@link AppPaths#KEY_PM_AI_MASTER_WORKBOOK} 環境タブへ書く既定。
     *
     * <p>湖南は空（{@code MASTER_WORKBOOK_FILE} の basename のみで解決）。国分は
     * {@code (リポジトリルート)/code/(}{@link #masterWorkbookFileBasename()}{@code )} の絶対パス文字列。
     *
     * @param ui {@link AppPaths#resolveRepoRoot(Map)} に必要なキーを含む（未設定時は cwd 系の既定）
     */
    public String pmAiMasterWorkbookEnvValue(Map<String, String> ui) {
        if (this != KOKUBU) {
            return "";
        }
        Path p =
                AppPaths.resolveRepoRoot(ui != null ? ui : Map.of())
                        .resolve("code")
                        .resolve(masterWorkbookFileBasename)
                        .normalize()
                        .toAbsolutePath();
        return p.toString();
    }

    /**
     * {@link AppPaths#KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK} 環境タブへ書く既定（絶対パス）。
     *
     * <p>{@code (リポジトリルート)/code/}{@link #summaryAiDispatchWorkbookCodeFilename}。
     *
     * @param ui {@link AppPaths#resolveRepoRoot(Map)} に必要なキーを含む
     */
    public String pmAiSummaryAiDispatchWorkbookEnvValue(Map<String, String> ui) {
        Path p =
                AppPaths.resolveRepoRoot(ui != null ? ui : Map.of())
                        .resolve("code")
                        .resolve(summaryAiDispatchWorkbookCodeFilename)
                        .normalize()
                        .toAbsolutePath();
        return p.toString();
    }

    /**
     * ポータブル同梱の {@code pm-ai-data/init_setting/session_defaults.json} から工場を推定する。
     *
     * <p>初回起動マーカー処理で湖南固定にしないため。{@link AppPaths#KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR}
     * の UNC に「国分」「湖南」が含まれるかで判定する。
     */
    public static Optional<FactorySite> inferFromPortableBundleInitSetting(Path portableExeDir) {
        if (portableExeDir == null) {
            return Optional.empty();
        }
        Path defaults =
                portableExeDir
                        .toAbsolutePath()
                        .normalize()
                        .resolve("pm-ai-data")
                        .resolve("init_setting")
                        .resolve(InitSettingPaths.SESSION_DEFAULTS_FILE);
        if (!Files.isRegularFile(defaults)) {
            return Optional.empty();
        }
        try {
            JsonNode root = new ObjectMapper().readTree(defaults.toFile());
            if (!root.isArray()) {
                return Optional.empty();
            }
            for (JsonNode row : root) {
                String name = textOrEmpty(row, "name");
                if (!AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR.equals(name)) {
                    continue;
                }
                return inferFromPortableBundleSourceValue(textOrEmpty(row, "value"));
            }
        } catch (Exception ignored) {
            return Optional.empty();
        }
        return Optional.empty();
    }

    static Optional<FactorySite> inferFromPortableBundleSourceValue(String raw) {
        if (raw == null || raw.isBlank()) {
            return Optional.empty();
        }
        if (raw.contains("国分")) {
            return Optional.of(KOKUBU);
        }
        if (raw.contains("湖南")) {
            return Optional.of(KONAN);
        }
        return Optional.empty();
    }

    private static String textOrEmpty(JsonNode row, String field) {
        JsonNode n = row.get(field);
        if (n == null || n.isNull()) {
            return "";
        }
        return n.asText("").trim();
    }
}
