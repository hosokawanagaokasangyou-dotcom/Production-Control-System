package jp.co.pm.ai.desktop.config;

import java.nio.file.Path;
import java.util.Map;

/**
 * 工場別の環境タブ既定（ネットワークソース・バージョンアップ正本 ZIP・マスタ名）。
 *
 * <p>ポータブル自動バージョンアップ完了時にユーザーが選択する。通常の「環境変数を初期化」は {@link #KONAN} と同一の既定を使う。
 */
public enum FactorySite {

    /** 湖南工場（従来のコード既定・工場共有 UNC）。 */
    KONAN(
            "湖南工場",
            AppPaths.DEFAULT_PM_AI_PORTABLE_BUNDLE_SOURCE_DIR,
            AppPaths.DEFAULT_PM_AI_TASK_INPUT_SOURCE_DIR,
            AppPaths.DEFAULT_PM_AI_ACTUAL_DETAIL_SOURCE_DIR,
            ""),

    /** 国分工場（国分共有・DATA 配下・マスタは {@code 国分master.xlsm}）。 */
    KOKUBU(
            "国分工場",
            "\\\\192.168.0.101\\共有フォルダ\\国分工場\\国分共有\\●配台AIシステム\\pm-ai-package-release\\"
                    + "PMD_version_upgrade.zip",
            "\\\\192.168.0.101\\共有フォルダ\\国分工場\\国分共有\\●配台AIシステム\\DATA\\計画",
            "\\\\192.168.0.101\\共有フォルダ\\国分工場\\国分共有\\●配台AIシステム\\DATA\\実績",
            "国分master.xlsm");

    private final String displayLabelJa;
    private final String portableBundleSourceDir;
    private final String taskInputSourceDir;
    private final String actualDetailSourceDir;
    /** {@link AppPaths#KEY_MASTER_WORKBOOK_FILE}。空のとき planning_core 側は {@code master.xlsm} 相当。 */
    private final String masterWorkbookFileBasename;

    FactorySite(
            String displayLabelJa,
            String portableBundleSourceDir,
            String taskInputSourceDir,
            String actualDetailSourceDir,
            String masterWorkbookFileBasename) {
        this.displayLabelJa = displayLabelJa;
        this.portableBundleSourceDir = portableBundleSourceDir;
        this.taskInputSourceDir = taskInputSourceDir;
        this.actualDetailSourceDir = actualDetailSourceDir;
        this.masterWorkbookFileBasename = masterWorkbookFileBasename;
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
}
