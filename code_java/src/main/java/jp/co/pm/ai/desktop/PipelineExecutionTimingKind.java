package jp.co.pm.ai.desktop;

/** 実行・ログタブに表示するパイプライン処理の計測種別。 */
public enum PipelineExecutionTimingKind {
    STAGE1("段階1"),
    /** 段階2.0（配台可能日時ベースの配台A）。旧履歴の {@code STAGE2} は読込時に本値へ正規化する。 */
    STAGE2_0("段階2.0"),
    STAGE2_1("段階2.1"),
    /** 段階3.0（入力3表・枝番分解後の配台A）。 */
    STAGE3_0("段階3.0"),
    /** 段階3.1（入力3表・時間外 hybrid）。 */
    STAGE3_1("段階3.1(時間外)"),
    /** 段階3.2（数量厳守: 同日完走必須・人ブロック無視）。 */
    STAGE3_2("段階3.2(数量厳守)"),
    /** 配台計画手動修正タブの単体試行（{@code dispatch_interactive_trial.py}）。 */
    STAGE3("配台試行"),
    SUMMARY_EXCEL("サマリ Excel"),
    DELIVERY_CALENDAR_VIEW("納期管理ビュー");

    private final String label;

    PipelineExecutionTimingKind(String label) {
        this.label = label;
    }

    public String label() {
        return label;
    }
}
