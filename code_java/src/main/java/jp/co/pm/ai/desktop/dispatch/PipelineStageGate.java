package jp.co.pm.ai.desktop.dispatch;

/**
 * 段階2.0 / 段階2.1 のボタン活性ゲート。
 *
 * <p>永続フラグは持たない。段階2.0/2.1 とも入力1表の存在で実行可否を判定する。
 */
public final class PipelineStageGate {

    /**
     * ゲート判定の入力。
     *
     * @param planInputExists 配台計画_タスク入力（入力1表）が保存済みか
     * @param stage2ResultExists 段階2.0/2.1 の {@code 結果_配台表.json} が存在するか（比較用。実行可否には使わない）
     */
    public record State(boolean planInputExists, boolean stage2ResultExists) {}

    private PipelineStageGate() {}

    /** 段階2.0: 入力1表が保存済みなら実行可能。 */
    public static boolean canRunStage20(State s) {
        return s != null && s.planInputExists();
    }

    /**
     * 段階2.1(時間外): 残業を決めたうえで段階2（配台A）を実行するシミュレーション。
     * 段階2.0 の事前実行は不要（入力1表があれば可）。
     */
    public static boolean canRunStage21(State s) {
        return canRunStage20(s);
    }

    /** ボタン無効時のツールチップ文言（実行可能なら空文字）。 */
    public static String stage21DisabledReason(State s) {
        return canRunStage21(s)
                ? ""
                : "段階2.1(時間外)の前に配台計画_タスク入力を保存してください。";
    }
}
