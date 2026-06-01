package jp.co.pm.ai.desktop;

/** 実行・ログタブに表示するパイプライン処理の計測種別。 */
public enum PipelineExecutionTimingKind {
    STAGE1("段階1"),
    STAGE2("段階2"),
    /** 段階2.0（配台可能日時ベースの配台A）。計測表示は段階2にまとめる。 */
    STAGE2_0("段階2.0"),
    STAGE2_5("段階2.5(AI)"),
    STAGE3("段階3"),
    STAGE2_1("段階2.1"),
    /** 段階3.0（入力3表・枝番分解後の配台A）。 */
    STAGE3_0("段階3.0"),
    /** 段階3.1（入力3表・時間外 hybrid）。 */
    STAGE3_1("段階3.1(時間外)"),
    /** 段階3.2（数量厳守: 同日完走必須・人ブロック無視）。 */
    STAGE3_2("段階3.2(数量厳守)"),
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
