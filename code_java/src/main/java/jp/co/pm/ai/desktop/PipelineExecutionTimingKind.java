package jp.co.pm.ai.desktop;

/** 実行・ログタブに表示するパイプライン処理の計測種別。 */
public enum PipelineExecutionTimingKind {
    STAGE1("段階1"),
    STAGE2("段階2"),
    STAGE2_5("段階2.5(AI)"),
    STAGE3("段階3"),
    STAGE3_5("段階3.5"),
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
