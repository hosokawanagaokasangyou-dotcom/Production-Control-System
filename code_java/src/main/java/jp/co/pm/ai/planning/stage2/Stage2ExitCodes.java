package jp.co.pm.ai.planning.stage2;

/** Python {@code plan_simulation_stage2.py} の終了コードに揃えた定数。 */
public final class Stage2ExitCodes {

    public static final int OK = 0;
    /** 入力ファイル欠如など（Python の FileNotFoundError exit 2）。 */
    public static final int FILE_NOT_FOUND = 2;
    /** 検証エラー（Python の PlanningValidationError exit 3）。 */
    public static final int VALIDATION = 3;
    /** その他の失敗。 */
    public static final int GENERAL_FAILURE = 1;

    private Stage2ExitCodes() {}
}
