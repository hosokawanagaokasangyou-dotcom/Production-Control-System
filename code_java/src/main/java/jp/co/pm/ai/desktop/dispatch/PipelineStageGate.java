package jp.co.pm.ai.desktop.dispatch;

/**
 * 段階2.0 / 段階2.1 のボタン活性ゲート。
 *
 * <p>段階の前提充足を「成果物の存在」から判定する純ロジック（永続フラグは持たない）。
 */
public final class PipelineStageGate {

    /**
     * ゲート判定の入力。
     *
     * @param planInputExists 配台計画_タスク入力（入力1表）が保存済みか
     * @param stage2ResultExists 段階2.0/2.1 の {@code 結果_配台表.json} が存在するか
     */
    public record State(boolean planInputExists, boolean stage2ResultExists) {}

    private PipelineStageGate() {}

    /** 段階2.0: 入力1表が保存済みなら実行可能。 */
    public static boolean canRunStage20(State s) {
        return s != null && s.planInputExists();
    }

    /** 段階2.1(時間外): 段階2.0 完了（結果_配台表.json 存在）が前提。 */
    public static boolean canRunStage21(State s) {
        return s != null && s.stage2ResultExists();
    }

    /** ボタン無効時のツールチップ文言（実行可能なら空文字）。 */
    public static String stage21DisabledReason(State s) {
        return canRunStage21(s)
                ? ""
                : "段階2.1(時間外)の前に段階2.0を実行し、結果_配台表.json を生成してください。";
    }
}
