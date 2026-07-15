package jp.co.pm.ai.desktop;

import javafx.scene.control.Label;
import javafx.scene.control.TitledPane;

/** 段階2開始からExcel自動生成完了までの進捗表示。 */
final class MainRunStage2Progress {

    private static final String RUNNING_STYLE = "pm-stage2-progress-running";
    private static final String SUCCESS_STYLE = "pm-stage2-progress-success";
    private static final String FAILURE_STYLE = "pm-stage2-progress-failure";

    enum State {
        RUNNING("段階2を実行中…", RUNNING_STYLE),
        DISPATCH_RELOADING("配台結果を再読込中…", RUNNING_STYLE),
        DELIVERY_RELOADING("納期管理を更新中…", RUNNING_STYLE),
        EXCEL_GENERATING("アラジン入力用Excelを生成中…", RUNNING_STYLE),
        COMPLETED("アラジン入力用Excelを生成しました", SUCCESS_STYLE),
        STAGE2_FAILED("段階2の実行に失敗しました", FAILURE_STYLE),
        FAILED("アラジン入力用Excelの生成に失敗しました", FAILURE_STYLE);

        private final String message;
        private final String styleClass;

        State(String message, String styleClass) {
            this.message = message;
            this.styleClass = styleClass;
        }

        String message() {
            return message;
        }

        String styleClass() {
            return styleClass;
        }
    }

    private MainRunStage2Progress() {}

    static void apply(TitledPane pane, Label label, State state, String detail) {
        if (pane != null) {
            pane.setExpanded(true);
        }
        if (label == null || state == null) {
            return;
        }
        String suffix = detail != null && !detail.isBlank() ? " — " + detail.strip() : "";
        label.setText(state.message + suffix);
        label.getStyleClass().removeAll(RUNNING_STYLE, SUCCESS_STYLE, FAILURE_STYLE);
        label.getStyleClass().add(state.styleClass);
    }
}
