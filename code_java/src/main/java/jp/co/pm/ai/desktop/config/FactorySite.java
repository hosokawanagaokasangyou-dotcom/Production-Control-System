package jp.co.pm.ai.desktop.config;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * 工場別の環境タブ既定（ネットワークソース・バージョンアップ正本 ZIP・マスタ名・サマリ用ブック）。
 *
 * <p>ポータル自動バージョンアップ完了時、および「環境変数を初期値に戻す」で工場を選んだときに適用する。
 * 湖南（{@link #KONAN}）は {@link AppPaths#DEFAULT_KONAN_SHARED_DATA_DIR}、国分（{@link #KOKUBU}）は
 * {@link AppPaths#DEFAULT_KOKUBU_DATA_DIR} 配下のマスタ／サマリ UNC を既定とする。
 */
public enum FactorySite {

    /** 湖南工場（工場共有 UNC・{@link AppPaths#DEFAULT_KONAN_SHARED_DATA_DIR} のマスタ／サマリ）。 */
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
            AppPaths.SUMMARY_AI_DISPATCH_XLSX);

    private final String displayLabelJa;
    private final String portableBundleSourceDir;
    private final String taskInputSourceDir;
    private final String actualDetailSourceDir;
    /** {@link AppPaths#KEY_PM_AI_MASTER_WORKBOOK} 未設定時の basename ヒント（工場プリセット識別用）。 */
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

    /** 工場別マスタ basename（{@link #pmAiMasterWorkbookEnvValue} の UNC ファイル名と対応）。 */
    public String masterWorkbookFileBasename() {
        return masterWorkbookFileBasename;
    }

    /**
     * {@link AppPaths#KEY_PM_AI_MASTER_WORKBOOK} 環境タブへ書く既定。
     *
     * <p>湖南は {@link AppPaths#DEFAULT_PM_AI_MASTER_WORKBOOK_KONAN}。国分は
     * {@link AppPaths#DEFAULT_PM_AI_MASTER_WORKBOOK_KOKUBU}。
     */
    public String pmAiMasterWorkbookEnvValue(Map<String, String> ui) {
        if (this == KONAN) {
            return AppPaths.DEFAULT_PM_AI_MASTER_WORKBOOK_KONAN;
        }
        if (this == KOKUBU) {
            return AppPaths.DEFAULT_PM_AI_MASTER_WORKBOOK_KOKUBU;
        }
        return "";
    }

    /**
     * {@link AppPaths#KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK} 環境タブへ書く既定（絶対パス）。
     *
     * <p>湖南は {@link AppPaths#DEFAULT_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK_KONAN}。国分は
     * {@link AppPaths#DEFAULT_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK_KOKUBU}。
     */
    public String pmAiSummaryAiDispatchWorkbookEnvValue(Map<String, String> ui) {
        if (this == KONAN) {
            return AppPaths.DEFAULT_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK_KONAN;
        }
        if (this == KOKUBU) {
            return AppPaths.DEFAULT_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK_KOKUBU;
        }
        return "";
    }

    /** {@link AppPaths#KEY_PM_AI_ALADDIN_MASTER_DIR} 環境タブへ書く既定（UNC）。 */
    public String aladdinMasterDir() {
        return AppPaths.defaultAladdinMasterDirForFactory(this);
    }

    /** {@link AppPaths#KEY_PM_AI_REQUEST_FORM_JUCHU_FILE} 環境タブへ書く既定（UNC）。 */
    public String requestFormJuchuFile() {
        return AppPaths.defaultRequestFormJuchuFileForFactory(this);
    }

    /** {@link AppPaths#KEY_PM_AI_REQUEST_FORM_ORIGINAL_DIR} 環境タブへ書く既定（UNC）。 */
    public String requestFormOriginalDir() {
        return AppPaths.defaultRequestFormOriginalDirForFactory(this);
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

    /**
     * 環境変数タブの工場別 UNC 等から利用工場を推定する。
     *
     * <p>複数キーを集計し、国分／湖南の票数が多い方を返す。同点・判定不能のときは empty（呼び出し側は {@link
     * GlobalInitSettingTarget#load()} を参照）。
     */
    public static Optional<FactorySite> inferFromUiEnv(Map<String, String> ui) {
        if (ui == null || ui.isEmpty()) {
            return Optional.empty();
        }
        int[] scores = new int[2];
        List<String> keys =
                List.of(
                        AppPaths.KEY_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR,
                        AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR,
                        AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
                        AppPaths.KEY_PM_AI_MASTER_WORKBOOK,
                        AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                        AppPaths.KEY_PM_AI_ALADDIN_MASTER_DIR,
                        AppPaths.KEY_PM_AI_REQUEST_FORM_JUCHU_FILE);
        for (String key : keys) {
            int weight =
                    AppPaths.KEY_PM_AI_TASK_INPUT_SOURCE_DIR.equals(key)
                                    || AppPaths.KEY_PM_AI_ACTUAL_DETAIL_SOURCE_DIR.equals(key)
                            ? 2
                            : 1;
            scoreEnvValue(ui.getOrDefault(key, ""), weight, scores);
        }
        if (scores[1] > scores[0] && scores[1] > 0) {
            return Optional.of(KOKUBU);
        }
        if (scores[0] > scores[1] && scores[0] > 0) {
            return Optional.of(KONAN);
        }
        return Optional.empty();
    }

    private static void scoreEnvValue(String raw, int weight, int[] scores) {
        Optional<FactorySite> site = inferFromPortableBundleSourceValue(raw);
        if (site.isEmpty() || weight <= 0) {
            return;
        }
        if (site.get() == KOKUBU) {
            scores[1] += weight;
        } else {
            scores[0] += weight;
        }
    }

    private static String textOrEmpty(JsonNode row, String field) {
        JsonNode n = row.get(field);
        if (n == null || n.isNull()) {
            return "";
        }
        return n.asText("").trim();
    }
}
