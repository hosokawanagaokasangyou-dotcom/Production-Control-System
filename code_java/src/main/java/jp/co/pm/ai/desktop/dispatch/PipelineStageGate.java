package jp.co.pm.ai.desktop.dispatch;

/**
 * 段階2.0〜3.2 のボタン活性ゲート。
 *
 * <p>段階の前提充足を「成果物の存在」から判定する純ロジック（永続フラグは持たない）。状態は実体（結果_配台表.json /
 * 入力3表シート / 段階3成果物）を反映するため、外部更新・復元後も整合する。可視化用の段階マトリックスには使わない（ボタン活性専用）。
 */
public final class PipelineStageGate {

    /**
     * ゲート判定の入力。
     *
     * @param planInputExists 配台計画_タスク入力（入力1表）が保存済みか
     * @param stage2ResultExists 段階2.0/2.1 の {@code 結果_配台表.json} が存在するか
     * @param stage3InputExists 入力3表（{@code 配台計画_タスク入力3.0} シート）が存在するか
     * @param stage3ResultExists 段階3.0 以降の成果物が存在するか
     */
    public record State(
            boolean planInputExists,
            boolean stage2ResultExists,
            boolean stage3InputExists,
            boolean stage3ResultExists) {}

    private PipelineStageGate() {}

    /** 段階2.0: 入力1表が保存済みなら実行可能。 */
    public static boolean canRunStage20(State s) {
        return s != null && s.planInputExists();
    }

    /** 段階2.1(時間外): 段階2.0 完了（結果_配台表.json 存在）が前提。 */
    public static boolean canRunStage21(State s) {
        return s != null && s.stage2ResultExists();
    }

    /** 入力3表を生成: 段階2.0 完了（結果_配台表.json 存在）が前提。 */
    public static boolean canBuildStage3Input(State s) {
        return s != null && s.stage2ResultExists();
    }

    /** 段階3.0: 入力3表が存在すれば実行可能。 */
    public static boolean canRunStage30(State s) {
        return s != null && s.stage3InputExists();
    }

    /** 段階3.1(時間外): 入力3表が存在し、段階3.0 完了が前提。 */
    public static boolean canRunStage31(State s) {
        return s != null && s.stage3InputExists() && s.stage3ResultExists();
    }

    /** 段階3.2(数量厳守): 入力3表が存在すれば実行可能。 */
    public static boolean canRunStage32(State s) {
        return s != null && s.stage3InputExists();
    }

    /** ボタン無効時のツールチップ文言（実行可能なら空文字）。 */
    public static String stage21DisabledReason(State s) {
        return canRunStage21(s)
                ? ""
                : "段階2.1(時間外)の前に段階2.0を実行し、結果_配台表.json を生成してください。";
    }

    public static String buildStage3InputDisabledReason(State s) {
        return canBuildStage3Input(s)
                ? ""
                : "入力3表の生成には段階2.0完了（結果_配台表.json）と手動修正の保存が必要です。";
    }

    public static String stage30DisabledReason(State s) {
        return canRunStage30(s) ? "" : "段階3.0の前に「入力3表を生成」してください。";
    }

    public static String stage31DisabledReason(State s) {
        if (canRunStage31(s)) {
            return "";
        }
        if (s == null || !s.stage3InputExists()) {
            return "段階3.1(時間外)の前に「入力3表を生成」してください。";
        }
        return "段階3.1(時間外)の前に段階3.0を実行してください。";
    }

    public static String stage32DisabledReason(State s) {
        return canRunStage32(s) ? "" : "段階3.2(数量厳守)の前に「入力3表を生成」してください。";
    }
}
